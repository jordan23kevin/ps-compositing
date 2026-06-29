import os
import time
import win32com.client
import pythoncom
import tempfile
from PIL import Image
import numpy as np
import config

ALPHA_THRESHOLD = 20


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


def place_design(design_path, torso_path, output_path, placement_cfg):
    """
    通用贴图函数（正图/背图共用）：
    1. Python预先trim透明边距 + 缩放贴图，保存为临时PNG
    2. JSX只需duplicate + 移动 + 旋转
    trim后图层左上角=有效像素左上角，移动moveY=targetTopY即可对齐
    """
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        t0 = time.time()

        # 计算贴图位置（用于trim和缩放）
        pos = calculate_sticker_position(design_path)
        scale = placement_cfg["scale_percent"] / 100

        # Python预先trim透明边距 + 缩放，保存为临时PNG
        img = Image.open(design_path).convert("RGBA")
        a = np.array(img)
        alpha = a[:, :, 3]
        mask = alpha >= ALPHA_THRESHOLD
        ys, xs = np.where(mask)
        # trim到有效像素区域
        x0, x1 = xs.min(), xs.max() + 1
        y0, y1 = ys.min(), ys.max() + 1
        trimmed = img.crop((x0, y0, x1, y1))
        # 缩放
        new_w = int((x1 - x0) * scale)
        new_h = int((y1 - y0) * scale)
        scaled = trimmed.resize((new_w, new_h), Image.LANCZOS)

        # 保存临时PNG
        temp_dir = tempfile.gettempdir()
        temp_design = os.path.join(temp_dir, "temp_design_scaled.png")
        scaled.save(temp_design, "PNG")

        # trim后图层左上角=有效像素左上角
        # 中心点在图层内的x偏移 = new_w/2
        # 要让中心点在target_center_x，图层左上角x = target_center_x - new_w/2
        # 要让顶部在target_top_y，图层左上角y = target_top_y
        move_x = placement_cfg["target_center_x"] - new_w / 2
        move_y = placement_cfg["target_top_y"]

        # 读取并替换JSX
        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_design.jsx")
        with open(jsx_path, "r", encoding="utf-8") as f:
            jsx_content = f.read()

        jsx_content = jsx_content.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{DESIGN_PATH}}", temp_design.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{ROTATION}}", str(placement_cfg["rotation"]))
        jsx_content = jsx_content.replace("{{MOVE_X}}", str(move_x))
        jsx_content = jsx_content.replace("{{MOVE_Y}}", str(move_y))

        # 写入临时JSX
        temp_dir = tempfile.gettempdir()
        temp_jsx = os.path.join(temp_dir, "temp_place_design.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx_content)

        psApp.DoJavaScriptFile(temp_jsx)
        dt = time.time() - t0
        print(f"✅ 生成: {output_path}  ({dt:.1f}秒)")

    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
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


def process_dx_folder(dx_folder):
    """处理单个 DX 文件夹，返回耗时（秒）"""
    dx_name = os.path.basename(dx_folder)
    rem_bg_folder = os.path.join(dx_folder, "02_REM_BG")
    upload_folder = os.path.join(dx_folder, "03_UPLOAD")

    if not os.path.exists(rem_bg_folder):
        return 0.0

    os.makedirs(upload_folder, exist_ok=True)
    t_dx = time.time()

    for file in os.listdir(rem_bg_folder):
        # 只处理 _cut.png 文件，跳过其他（如 DXxxxx_B.png）
        if not file.lower().endswith("_cut.png"):
            continue
        # 跳过黑版（反相生成的文件，含_黑）
        if "_黑" in file:
            continue

        print(f"\n处理: {file}")
        design_path = os.path.join(rem_bg_folder, file)
        design_type = classify_design(file)

        if design_type == "BW":
            # ===== BW 类型：生成 W 和 B 两套文件，供 ps_batch.py 合成 =====
            print("  → 生成 W 正面文件（新方案）...")
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_white"]),
                os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"),
                config.FRONT_NEW,
            )
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_black"]),
                os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"),
                config.FRONT_NEW,
            )

            print("  → 生成 B 背面文件（新方案）...")
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_white"]),
                os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"),
                config.BACK_NEW,
            )
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_black"]),
                os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"),
                config.BACK_NEW,
            )
            print("  ✅ BW 准备完成，可运行 ps_batch.py 合成最终 BW 图！")

        elif design_type == "W":
            # ===== W 类型：正图新方案 =====
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_white"]),
                os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"),
                config.FRONT_NEW,
            )
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_black"]),
                os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"),
                config.FRONT_NEW,
            )

        elif design_type == "B":
            # ===== B 类型：背图新方案 =====
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_white"]),
                os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"),
                config.BACK_NEW,
            )
            place_design(
                design_path,
                os.path.join(config.BASE_TORSO, config.BACK_NEW["torso_black"]),
                os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"),
                config.BACK_NEW,
            )

    dt_dx = time.time() - t_dx
    print(f"⏱️  {dx_name} 完成，耗时 {dt_dx:.1f}秒")
    return dt_dx


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
    for folder in folders:
        dx_folder = os.path.join(dx_root, folder)
        print(f"\n===== 处理: {dx_folder} =====")
        dt = process_dx_folder(dx_folder)
        dx_times.append((folder, dt))

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
    print(f"{'平均':<12} {dt_total/len(dx_times):>10.1f}")
    print(f"{'='*60}")


if __name__ == "__main__":
    import sys
    start = sys.argv[1] if len(sys.argv) > 1 else None
    main(start)
