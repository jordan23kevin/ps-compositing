"""黑T贴图处理 v2.4 — 反相完成后调用，PS贴图+BW合成全覆盖

变更 v2.3：
  - 复用 wb_sticker_ps.StickerSession 进行黑T贴图，与通用贴图共享 JSX 路径模板，
    修复因 {{TORSO_DOC_NAME}} / {{DESIGN_DOC_NAME}} 占位符与 JSX 不匹配导致的执行失败。
  - BW 合成仍在本脚本内完成，直接操作 StickerSession 的 Photoshop COM 会话。
  - 单 DX 内只开启一次 Photoshop，全部黑T贴图与 BW 合成完成后才关闭。
"""
import os, re, sys, tempfile, time
from pathlib import Path

sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from config import SOURCE_BASE, parse_side_suffix
from wb_sticker_ps import StickerSession, _role_from_name, cleanup_stale_uploads
import wb_naming  # 命名规则唯一出处（sys.path 由 wb_sticker_ps 注入）

BASE = Path(SOURCE_BASE)
TORSO = Path(r"D:\Semems\1胚衣")

# ---------------------------------------------------------------------------
# 元数据辅助
# ---------------------------------------------------------------------------
_MIGRATED_DX = set()


def _infer_meta(path):
    """sidecar 完全缺失时的文件名兜底推断"""
    name = os.path.basename(path)
    dx = name.split("_")[0] if "_" in name else "DX"
    role = _role_from_name(name)
    uid = f"UID_{dx}_{re.sub(r'[^A-Za-z0-9]', '_', role)}"
    group_id = f"G_{dx}_{role}"
    return {"uid": uid, "group_id": group_id, "role": role, "stage": "unknown"}


def _get_meta(path):
    """读取 sidecar；缺失时 migrate_dx，最后兜底推断"""
    if wb_meta is None:
        return None
    dx_dir = str(Path(path).parent.parent)
    meta = wb_meta.read_meta(path)
    if meta:
        return meta
    if dx_dir not in _MIGRATED_DX:
        try:
            wb_meta.migrate_dx(dx_dir)
        except Exception as e:
            print(f"  ⚠️ migrate_dx 失败 {dx_dir}: {e}")
        _MIGRATED_DX.add(dx_dir)
    meta = wb_meta.read_meta(path)
    if meta:
        return meta
    return _infer_meta(path)


BACK_NEW = {
    "torso_white": "白背2.jpg", "torso_black": "黑背2.jpg",
    "scale_percent": 30, "rotation": -1,
    "target_center_x": 680, "target_top_y": 570,
}
FRONT_NEW = {
    "torso_white": "白正2.jpg", "torso_black": "黑正2.jpg",
    "scale_percent": 13.33, "rotation": -1,
    "target_center_x": 888, "target_top_y": 612,
}


# ---------------------------------------------------------------------------
# Photoshop 会话：单 DX 内复用 StickerSession
# ---------------------------------------------------------------------------
class BlackStickerSession:
    def __init__(self):
        # StickerSession 内部已经做了 CoInitialize / PS 启动 / 隐藏窗口
        self.sticker = StickerSession()

    # _open_file 已移除：纯软件实现不再需要打开 Photoshop 文档

    def place_one(self, side, cfg, inv_path, dx, upload, cut_meta=None):
        """贴一张黑T：直接复用 StickerSession 的路径模板 JSX。"""
        output_name = wb_naming.flat_name(dx, side, "黑")
        output_path = str(upload / output_name)
        torso_path = str(TORSO / cfg["torso_black"])
        self.sticker.place_design(inv_path, torso_path, output_path, cfg, cut_meta=cut_meta)
        return output_name

    def bw_synth(self, dx, upload):
        """BW合成（纯软件）：复用 ps_batch.process_dx 合成该款全部 BW。"""
        sys.path.insert(0, os.path.dirname(__file__))
        try:
            import ps_batch
            results = ps_batch.process_dx(None, dx)
            ok = any(ok_flag for _, ok_flag in results)
            return "✅ 黑BW 合成完成（纯软件）" if ok else "⏭️ 黑BW 合成跳过（缺平铺图）"
        except Exception as e:
            return f"⚠️ 黑BW 合成失败: {e}"

    def close(self):
        self.sticker.close()


def main():
    dx = sys.argv[1] if len(sys.argv) > 1 else ""
    if not dx:
        print("用法: python process_black.py DX0124")
        return
    print(f"=== 黑T贴图处理: {dx} ===")
    t0 = time.time()
    upload = BASE / dx / "03_UPLOAD"
    rembg = BASE / dx / "02_REM_BG"

    # 单面款先清理 03_UPLOAD 旧互补面/平铺残留（如只有 _黑W 时清掉旧 B_* 与 *BW）
    cleanup_stale_uploads(str(BASE / dx))

    # 找黑版文件（支持版本号后缀：黑B2 / 黑W1 / 黑BW3 等）
    tasks = []
    for inv_path in sorted(rembg.glob(f"{dx}_黑*_cut.png")):
        letter = Path(inv_path).stem.replace(f"{dx}_黑", "").replace("_cut", "")
        side, version = parse_side_suffix(letter)
        meta = _get_meta(str(inv_path))
        if side == "B":
            tasks.append(("B", BACK_NEW, inv_path, meta))
        elif side == "W":
            tasks.append(("W", FRONT_NEW, inv_path, meta))
        elif side == "WB" or side == "BW":
            tasks.append(("B", BACK_NEW, inv_path, meta))
            tasks.append(("W", FRONT_NEW, inv_path, meta))
        else:
            print(f"  ⚠️ 跳过无法识别的黑版文件: {inv_path.name}")

    if not tasks:
        print("❌ 未找到黑版_cut文件")
        return

    session = BlackStickerSession()
    try:
        # 贴图：只做黑色胚衣版本
        print(f"\n--- 贴图 ({len(tasks)}张) ---")
        for side, cfg, inv_path, meta in tasks:
            session.place_one(side, cfg, inv_path, dx, upload, cut_meta=meta)

        # BW合成
        has_b = any(s == "B" for s, _, _, _ in tasks)
        has_w = any(s == "W" for s, _, _, _ in tasks)
        if has_b and has_w:
            print("\n--- BW合成 ---")
            print(f"  {session.bw_synth(dx, upload)}")
        else:
            print("⏭️ 只有单面，跳过BW合成")
    finally:
        session.close()

    dt = time.time() - t0
    print(f"\n⏱️ 总耗时: {dt:.1f}秒")
    print("=" * 40)


if __name__ == "__main__":
    main()
