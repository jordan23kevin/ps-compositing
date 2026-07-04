"""黑T贴图处理 v2.2 — 反相完成后调用，PS贴图+BW合成全覆盖

变更 v2.2：
  - 单 DX 内复用 Photoshop COM 会话，缓存黑胚衣文档
  - 用主动轮询替代硬编码 sleep
  - 与 check_rem.py /invert-rem 联动：反相生成黑版专用图后自动调用本脚本
"""
import os, re, sys, tempfile, time
import win32com.client
import pythoncom
from pathlib import Path
from PIL import Image
import numpy as np

sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from config import hide_ps_window, parse_side_suffix

ALPHA_THRESHOLD = 20
BASE = Path(r"D:\Semems WB\02_PROJECTS")
TORSO = Path(r"D:\Semems\1胚衣")

# ---------------------------------------------------------------------------
# 元数据辅助
# ---------------------------------------------------------------------------
_MIGRATED_DX = set()


def _role_from_name(name):
    """从文件名推断 role（支持 _cut.png / 上传图 / BW 合成图）"""
    stem = os.path.splitext(name)[0]
    if stem.endswith("_cut"):
        stem = stem[:-4]
    parts = stem.split("_")
    if len(parts) >= 3 and parts[-1] in ("白T", "黑T"):
        side = parts[-2]
        torso = parts[-1]
        if torso == "黑T":
            return f"黑{side}"
        return side
    if len(parts) >= 2:
        return parts[-1]
    return "?"


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
# Photoshop 会话：单 DX 内复用 COM、缓存黑胚衣文档
# ---------------------------------------------------------------------------
class BlackStickerSession:
    def __init__(self):
        pythoncom.CoInitialize()
        self.ps_app = win32com.client.Dispatch("Photoshop.Application")
        self.ps_app.DisplayDialogs = 3  # DialogModes.NO
        hide_ps_window(self.ps_app)
        self.torso_docs = {}      # torso_path -> doc
        self.design_doc = None
        self.design_doc_name = None
        self.temp_dir = tempfile.gettempdir()

    def _open_file(self, file_path):
        ps_file = win32com.client.Dispatch("Photoshop.File")
        ps_file.Path = file_path
        return self.ps_app.Open(ps_file)

    def _get_torso_doc(self, torso_path):
        if torso_path not in self.torso_docs:
            self.torso_docs[torso_path] = self._open_file(torso_path)
            hide_ps_window(self.ps_app)
        return self.torso_docs[torso_path]

    def _prepare_scaled_design(self, inv_path, cfg):
        img = Image.open(inv_path).convert("RGBA")
        a = np.array(img)
        alpha = a[:, :, 3]
        mask = alpha >= ALPHA_THRESHOLD
        ys, xs = np.where(mask)
        if len(ys) == 0:
            raise ValueError(f"无有效像素: {inv_path}")
        x0, x1 = xs.min(), xs.max() + 1
        y0, y1 = ys.min(), ys.max() + 1
        trimmed = img.crop((x0, y0, x1, y1))
        scale = cfg["scale_percent"] / 100
        new_w = int((x1 - x0) * scale)
        new_h = int((y1 - y0) * scale)
        scaled = trimmed.resize((new_w, new_h), Image.LANCZOS)

        torso_name = os.path.splitext(os.path.basename(cfg["torso_black"]))[0]
        temp_path = os.path.join(self.temp_dir, f"temp_inv_{torso_name}_scaled.png")
        scaled.save(temp_path, "PNG")

        move_x = cfg["target_center_x"] - new_w / 2
        move_y = cfg["target_top_y"]
        return temp_path, move_x, move_y

    def _open_design_doc(self, temp_design_path):
        self._close_design_doc()
        self.design_doc = self._open_file(temp_design_path)
        self.design_doc_name = os.path.basename(temp_design_path)
        hide_ps_window(self.ps_app)

    def _close_design_doc(self):
        if self.design_doc is not None:
            try:
                self.design_doc.Close(2)
            except Exception:
                pass
            self.design_doc = None
            self.design_doc_name = None

    def place_one(self, side, cfg, inv_path, dx, upload, cut_meta=None):
        """贴一张黑T。"""
        temp_design, move_x, move_y = self._prepare_scaled_design(inv_path, cfg)
        self._open_design_doc(temp_design)

        torso_path = str(TORSO / cfg["torso_black"])
        torso_doc = self._get_torso_doc(torso_path)

        color_suffix = "黑T"
        output_name = f"{dx}_{side}_{color_suffix}.jpg"
        output_path = str(upload / output_name)

        # 若旧文件存在则先删除，确保 saveAs 直接覆盖不弹窗
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"  ⚠️ 删除旧{output_name}失败: {e}")

        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_design.jsx")
        with open(jsx_path, "r", encoding="utf-8") as jf:
            jsx = jf.read()
        jsx = jsx.replace("{{DESIGN_DOC_NAME}}", self.design_doc_name)
        jsx = jsx.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx = jsx.replace("{{ROTATION}}", str(cfg["rotation"]))
        jsx = jsx.replace("{{MOVE_X}}", str(move_x))
        jsx = jsx.replace("{{MOVE_Y}}", str(move_y))

        temp_jsx = os.path.join(self.temp_dir, "temp_place.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as jf:
            jf.write(jsx)

        t0 = time.time()
        self.ps_app.ActiveDocument = torso_doc
        self.ps_app.DoJavaScriptFile(temp_jsx)
        hide_ps_window(self.ps_app)
        dt = time.time() - t0
        print(f"  ✅ {output_name} ({dt:.1f}秒)")

        if cut_meta is not None and wb_meta is not None:
            try:
                role = _role_from_name(output_name)
                wb_meta.register_sticker(
                    output_path,
                    uid=cut_meta["uid"],
                    group_id=cut_meta["group_id"],
                    role=role,
                    parent_uid=cut_meta["uid"],
                    cut_file=os.path.basename(inv_path),
                )
            except Exception as e:
                print(f"  ⚠️ 元数据注册失败({side}): {e}")

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
        while self.ps_app.Documents.Count < 2 and time.time() - t0 < 10:
            time.sleep(0.05)
        hide_ps_window(self.ps_app)

        # 激活背面文档并执行动作
        self.ps_app.ActiveDocument = b_doc
        self.ps_app.DoAction('黑', '正反图')
        hide_ps_window(self.ps_app)

        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except Exception as e:
                print(f"  ⚠️ 删除旧黑BW失败: {e}")

        jpgOpt = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        jpgOpt.Quality = 12
        self.ps_app.ActiveDocument.SaveAs(out_path, jpgOpt, True, 2)

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
        self._close_design_doc()
        for doc in list(self.torso_docs.values()):
            try:
                doc.Close(2)
            except:
                pass
        self.torso_docs.clear()
        pythoncom.CoUninitialize()


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
