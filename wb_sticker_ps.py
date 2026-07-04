# ===== WB 贴图主控 v2.2 =====
# 变更 v2.2：
#   - 单 DX 内复用 Photoshop COM 会话，避免每贴一张图都重新连接 PS
#   - 黑T优先使用 02_REM_BG 中的 _黑B/_黑W/_黑BW 专用文件
#   - 检测到黑版专用文件时，通用图不再输出黑T成品，仅输出白T
import os
import re
import sys
import io
import time
import win32com.client
import pythoncom
import tempfile
from pathlib import Path
from PIL import Image
import numpy as np
import config

# 确保 Windows 控制台能输出中文/emoji
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None

ALPHA_THRESHOLD = 20

# ---------------------------------------------------------------------------
# 元数据辅助（读取 _cut.png sidecar，为上传图注册）
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


def _get_cut_meta(cut_path):
    """读取去背图 sidecar；缺失时尝试 migrate_dx，最后兜底推断"""
    if wb_meta is None:
        return None
    dx_dir = str(Path(cut_path).parent.parent)
    meta = wb_meta.read_meta(cut_path)
    if meta:
        return meta
    if dx_dir not in _MIGRATED_DX:
        try:
            wb_meta.migrate_dx(dx_dir)
        except Exception as e:
            print(f"  ⚠️ migrate_dx 失败 {dx_dir}: {e}")
        _MIGRATED_DX.add(dx_dir)
    meta = wb_meta.read_meta(cut_path)
    if meta:
        return meta
    return _infer_meta(cut_path)


def calculate_sticker_position(png_path):
    """
    计算贴图的关键位置（复刻美图秀秀 calc_offset_back 逻辑）：
    1. 最高有效像素的 y 坐标：从上往下找，找到一行里有效像素≥50个的那行
       （防止顶部噪点干扰）
    2. 中心点 x 坐标：透明度加权质心 (x * alpha).sum() / alpha.sum()
    """
    img = Image.open(png_path).convert("RGBA")
    a = np.array(img)
    h, w = a.shape[0], a.shape[1]
    alpha = a[:, :, 3].astype(np.float64)

    mask = alpha >= ALPHA_THRESHOLD
    ys, xs = np.where(mask)

    if len(ys) == 0:
        top_y = 0
        center_x = w / 2
    else:
        # 最高有效像素y：从最小y开始往下找，找到一行里有效像素≥50个的那行
        y_min = int(ys.min())
        for y in range(y_min, min(y_min + 200, h)):
            if (alpha[y] >= ALPHA_THRESHOLD).sum() >= 50:
                y_min = y
                break
        top_y = y_min

        # 中心点x：透明度加权质心
        weights = np.where(mask, alpha, 0.0)
        center_x = float((np.indices(alpha.shape)[1] * weights).sum() / weights.sum())

    return {
        "top_y": float(top_y),
        "center_x": float(center_x),
        "original_width": w,
        "original_height": h,
    }


# ---------------------------------------------------------------------------
# Photoshop 会话：单 DX 内复用 COM
# ---------------------------------------------------------------------------
class StickerSession:
    def __init__(self):
        pythoncom.CoInitialize()
        self.ps_app = win32com.client.Dispatch("Photoshop.Application")
        self.ps_app.DisplayDialogs = 3  # DialogModes.NO：不弹任何对话框
        config.hide_ps_window(self.ps_app)
        self.temp_dir = tempfile.gettempdir()

    def _prepare_scaled_design(self, design_path, placement_cfg):
        """预先 trim + 缩放贴图，返回 (临时文件路径, move_x, move_y)。"""
        pos = calculate_sticker_position(design_path)
        scale = placement_cfg["scale_percent"] / 100

        img = Image.open(design_path).convert("RGBA")
        a = np.array(img)
        alpha = a[:, :, 3]
        mask = alpha >= ALPHA_THRESHOLD
        ys, xs = np.where(mask)
        if len(ys) == 0:
            raise ValueError(f"无有效像素: {design_path}")

        x0, x1 = xs.min(), xs.max() + 1
        y0, y1 = ys.min(), ys.max() + 1
        trimmed = img.crop((x0, y0, x1, y1))
        new_w = int((x1 - x0) * scale)
        new_h = int((y1 - y0) * scale)
        scaled = trimmed.resize((new_w, new_h), Image.LANCZOS)

        # 用 torso 文件名区分临时文件，避免同进程内冲突
        torso_name = os.path.splitext(os.path.basename(placement_cfg["torso_white"]))[0]
        temp_path = os.path.join(self.temp_dir, f"temp_design_{torso_name}_scaled.png")
        scaled.save(temp_path, "PNG")

        move_x = placement_cfg["target_center_x"] - new_w / 2
        move_y = placement_cfg["target_top_y"]
        return temp_path, move_x, move_y

    def place_design(self, design_path, torso_path, output_path, placement_cfg, cut_meta=None):
        """通用贴图函数：JSX 内打开胚衣和设计图，贴图后关闭。"""
        temp_design, move_x, move_y = self._prepare_scaled_design(design_path, placement_cfg)

        # 若旧文件存在则先删除，确保 saveAs 直接覆盖不弹窗
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception as e:
                print(f"  ⚠️ 删除旧{os.path.basename(output_path)}失败: {e}")

        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_design.jsx")
        with open(jsx_path, "r", encoding="utf-8") as f:
            jsx_content = f.read()

        jsx_content = jsx_content.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{DESIGN_PATH}}", temp_design.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{ROTATION}}", str(placement_cfg["rotation"]))
        jsx_content = jsx_content.replace("{{MOVE_X}}", str(move_x))
        jsx_content = jsx_content.replace("{{MOVE_Y}}", str(move_y))

        temp_jsx = os.path.join(self.temp_dir, "temp_place_design.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx_content)

        t0 = time.time()
        self.ps_app.DoJavaScriptFile(temp_jsx)
        config.hide_ps_window(self.ps_app)
        dt = time.time() - t0
        print(f"✅ 生成: {output_path}  ({dt:.1f}秒)")

        if cut_meta is not None and wb_meta is not None:
            try:
                out_name = os.path.basename(output_path)
                role = _role_from_name(out_name)
                wb_meta.register_sticker(
                    output_path,
                    uid=cut_meta["uid"],
                    group_id=cut_meta["group_id"],
                    role=role,
                    parent_uid=cut_meta["uid"],
                    cut_file=os.path.basename(design_path),
                )
            except Exception as e:
                print(f"  ⚠️ 元数据注册失败: {e}")

    def close(self):
        pythoncom.CoUninitialize()


def classify_design(filename):
    """根据文件名分类设计图
    BW/WB → 正面+背面都要做
    B → 只做背面
    W → 只做正面
    """
    base = os.path.splitext(filename)[0]

    if "_BW_" in base or "_BW" in base or "_WB_" in base or "_WB" in base:
        return "BW"
    elif "_B_" in base or "_B" in base:
        return "B"
    elif "_W_" in base or "_W" in base:
        return "W"
    elif base.endswith("_1"):
        return "B"
    elif base.endswith("_2"):
        return "W"
    else:
        return "W"


def black_counterpart(file, dx):
    """返回通用_cut文件对应的黑版专用文件名（如 DX_B_cut.png → DX_黑B_cut.png）。
    支持版本号后缀：DX_B2_cut.png → DX_黑B2_cut.png。
    不存在对应关系时返回 None。"""
    if "_黑" in file or not file.lower().endswith("_cut.png"):
        return None
    stem = file[:-len("_cut.png")]  # e.g. DXxxxx_B / DXxxxx_BW / DXxxxx_B2
    suffix = stem[len(dx)+1:] if stem.startswith(dx + "_") else ""
    side, version = config.parse_side_suffix(suffix)
    if side in ("B", "W", "BW", "WB"):
        return f"{dx}_黑{side}{version}_cut.png"
    return None


def process_dx_folder(dx_folder, session=None):
    """处理单个 DX 文件夹，返回耗时（秒）。session 为 None 时内部创建。"""
    dx_name = os.path.basename(dx_folder)
    rem_bg_folder = os.path.join(dx_folder, "02_REM_BG")
    upload_folder = os.path.join(dx_folder, "03_UPLOAD")

    if not os.path.exists(rem_bg_folder):
        return 0.0

    os.makedirs(upload_folder, exist_ok=True)
    own_session = session is None
    if own_session:
        session = StickerSession()

    t_dx = time.time()
    try:
        for file in os.listdir(rem_bg_folder):
            # 只处理 _cut.png 文件，跳过其他（如 DXxxxx_B.png）
            if not file.lower().endswith("_cut.png"):
                continue
            # 跳过黑版（反相生成的文件，含_黑），黑T贴图由 process_black.py 处理
            if "_黑" in file:
                continue

            print(f"\n处理: {file}")
            design_path = os.path.join(rem_bg_folder, file)
            design_type = classify_design(file)
            cut_meta = _get_cut_meta(design_path)

            # 如果存在对应的黑版专用文件，则通用图不再用于黑T
            black_file = black_counterpart(file, dx_name)
            has_black = black_file and os.path.exists(os.path.join(rem_bg_folder, black_file))
            if has_black:
                print(f"  发现黑版专用 {black_file}，通用图仅用于白T")

            if design_type == "BW":
                # ===== BW 类型：生成 W 和 B 两套文件，供 ps_batch.py 合成 =====
                print("  → 生成 W 正面文件（新方案）...")
                session.place_design(
                    design_path,
                    os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_white"]),
                    os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"),
                    config.FRONT_NEW,
                    cut_meta=cut_meta,
                )
                if not has_black:
                    session.place_design(
                        design_path,
                        os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_black"]),
                        os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"),
                        config.FRONT_NEW,
                        cut_meta=cut_meta,
                    )

                print("  → 生成 B 背面文件（新方案）...")
                session.place_design(
                    design_path,
                    os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_white"]),
                    os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"),
                    config.BACK_NEW,
                    cut_meta=cut_meta,
                )
                if not has_black:
                    session.place_design(
                        design_path,
                        os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_black"]),
                        os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"),
                        config.BACK_NEW,
                        cut_meta=cut_meta,
                    )
                print("  ✅ BW 准备完成，可运行 ps_batch.py 合成最终 BW 图！")

            elif design_type == "W":
                # ===== W 类型：正图新方案 =====
                session.place_design(
                    design_path,
                    os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_white"]),
                    os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"),
                    config.FRONT_NEW,
                    cut_meta=cut_meta,
                )
                if not has_black:
                    session.place_design(
                        design_path,
                        os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_black"]),
                        os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"),
                        config.FRONT_NEW,
                        cut_meta=cut_meta,
                    )

            elif design_type == "B":
                # ===== B 类型：背图新方案 =====
                session.place_design(
                    design_path,
                    os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_white"]),
                    os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"),
                    config.BACK_NEW,
                    cut_meta=cut_meta,
                )
                if not has_black:
                    session.place_design(
                        design_path,
                        os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_black"]),
                        os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"),
                        config.BACK_NEW,
                        cut_meta=cut_meta,
                    )

        dt_dx = time.time() - t_dx
        print(f"⏱️  {dx_name} 完成，耗时 {dt_dx:.1f}秒")
        return dt_dx
    finally:
        if own_session:
            session.close()


def main(start_dx=None):
    """批量处理所有 DX 文件夹
    start_dx: 指定起始DX（如'DX0124'），从该文件夹开始处理（含）；None表示处理全部
    """
    dx_root = config.SOURCE_BASE
    folders = sorted([d for d in os.listdir(dx_root)
                       if d.startswith("DX") and os.path.isdir(os.path.join(dx_root, d))])
    if start_dx:
        idx = next((i for i, f in enumerate(folders) if f == start_dx), 0)
        folders = folders[idx:]
        print(f"从 {start_dx} 开始，共 {len(folders)} 个文件夹")

    t_total = time.time()
    dx_times = []
    session = StickerSession()
    try:
        for folder in folders:
            dx_folder = os.path.join(dx_root, folder)
            print(f"\n===== 处理: {dx_folder} =====")
            dt = process_dx_folder(dx_folder, session=session)
            dx_times.append((folder, dt))
    finally:
        session.close()

    dt_total = time.time() - t_total
    # 汇总
    print(f"\n{'='*60}")
    print(f"📊 全部完成！共 {len(dx_times)} 款")
    print(f"{'='*60}")
    print(f"{'DX':<12} {'耗时(秒)':>10}")
    print(f"{'-'*24}")
    for name, dt in dx_times:
        print(f"{name:<12} {dt:>10.1f}")
    print(f"{'-'*24}")
    print(f"{'总计':<12} {dt_total:>10.1f}  ({dt_total/60:.1f}分钟)")
    if dx_times:
        print(f"{'平均':<12} {dt_total/len(dx_times):>10.1f}")
    print(f"{'='*60}")


if __name__ == "__main__":
    import sys
    start = sys.argv[1] if len(sys.argv) > 1 else None
    main(start)
