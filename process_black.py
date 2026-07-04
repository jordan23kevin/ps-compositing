"""黑T贴图处理 v2.3 — 反相完成后调用，PS贴图+BW合成全覆盖

变更 v2.3：
  - 复用 wb_sticker_ps.StickerSession 进行黑T贴图，与通用贴图共享 JSX 路径模板，
    修复因 {{TORSO_DOC_NAME}} / {{DESIGN_DOC_NAME}} 占位符与 JSX 不匹配导致的执行失败。
  - BW 合成仍在本脚本内完成，直接操作 StickerSession 的 Photoshop COM 会话。
  - 单 DX 内只开启一次 Photoshop，全部黑T贴图与 BW 合成完成后才关闭。
"""
import os, re, sys, tempfile, time
import win32com.client
from pathlib import Path

sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from config import SOURCE_BASE, parse_side_suffix
from wb_sticker_ps import StickerSession, _role_from_name

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

    def _open_file(self, file_path):
        """通过 JSX app.open 打开文件，并返回当前活动文档对象。"""
        doc_name = os.path.basename(file_path)
        jsx = (
            'var f = new File("' + file_path.replace("\\", "\\\\") + '");\n'
            'app.open(f);\n'
        )
        temp_jsx = os.path.join(tempfile.gettempdir(), f"_open_{doc_name}.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx)
        self.sticker.ps_app.DoJavaScriptFile(temp_jsx)
        return self.sticker.ps_app.ActiveDocument

    def place_one(self, side, cfg, inv_path, dx, upload, cut_meta=None):
        """贴一张黑T：直接复用 StickerSession 的路径模板 JSX。"""
        output_name = f"{dx}_{side}_黑T.jpg"
        output_path = str(upload / output_name)
        torso_path = str(TORSO / cfg["torso_black"])
        self.sticker.place_design(inv_path, torso_path, output_path, cfg, cut_meta=cut_meta)
        return output_name

    def bw_synth(self, dx, upload):
        """BW合成：用PS动作合并B和W（黑T版）。"""
        b_img = str(upload / f"{dx}_B_黑T.jpg")
        w_img = str(upload / f"{dx}_W_黑T.jpg")
        out_path = str(upload / f"{dx}_黑BW.jpg")

        if not os.path.exists(b_img) or not os.path.exists(w_img):
            return "⏭️ 缺少 B/W 黑T文件，跳过黑BW合成"

        b_doc = self._open_file(b_img)
        w_doc = self._open_file(w_img)

        # 主动轮询等待两个文档都打开
        t0 = time.time()
        while self.sticker.ps_app.Documents.Count < 2 and time.time() - t0 < 10:
            time.sleep(0.05)

        # 激活背面文档并执行动作
        self.sticker.ps_app.ActiveDocument = b_doc
        self.sticker.ps_app.DoAction('黑', '正反图')

        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except Exception as e:
                print(f"  ⚠️ 删除旧黑BW失败: {e}")

        jpgOpt = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        jpgOpt.Quality = 12
        self.sticker.ps_app.ActiveDocument.SaveAs(out_path, jpgOpt, True, 2)

        # 关闭本次打开的 B/W 文档
        for doc in (b_doc, w_doc):
            try:
                doc.Close(2)
            except:
                pass

        if wb_meta is not None:
            try:
                meta_b = _get_meta(b_img)
                meta_w = _get_meta(w_img)
                uid = meta_b.get("uid") or meta_w.get("uid")
                group_id = meta_b.get("group_id") or meta_w.get("group_id")
                bw_role = "黑BW"
                bw_uid = f"{uid}_{bw_role}" if uid else None
                wb_meta.register_bw(
                    out_path,
                    uid=bw_uid,
                    group_id=group_id,
                    role=bw_role,
                    source_uids=[meta_b.get("uid"), meta_w.get("uid")],
                    source_files=[os.path.basename(b_img), os.path.basename(w_img)],
                )
                print(f"  BW元数据已注册: {bw_role}")
            except Exception as e:
                print(f"  ⚠️ 黑BW元数据注册失败: {e}")

        return "✅ 黑BW 合成完成"

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
