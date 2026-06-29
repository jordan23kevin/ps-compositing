import os
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
    1. 用 duplicate 复制图层（保留原始尺寸和透明度）
    2. 按配置缩放（百分比，相对于原始尺寸）
    3. 旋转
    4. 贴图最高有效像素 → PS Y = target_top_y
    5. 贴图中心点 → PS X = target_center_x
    """
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")

        # 计算贴图位置
        pos = calculate_sticker_position(design_path)
        scale = placement_cfg["scale_percent"] / 100
        scaled_top_y = pos["top_y"] * scale
        scaled_center_x = pos["center_x"] * scale

        # 读取并替换JSX
        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_design.jsx")
        with open(jsx_path, "r", encoding="utf-8") as f:
            jsx_content = f.read()

        jsx_content = jsx_content.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{DESIGN_PATH}}", design_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{TARGET_CENTER_X}}", str(placement_cfg["target_center_x"]))
        jsx_content = jsx_content.replace("{{TARGET_TOP_Y}}", str(placement_cfg["target_top_y"]))
        jsx_content = jsx_content.replace("{{SCALE_PERCENT}}", str(placement_cfg["scale_percent"]))
        jsx_content = jsx_content.replace("{{ROTATION}}", str(placement_cfg["rotation"]))
        jsx_content = jsx_content.replace("{{SCALED_TOP_Y}}", str(scaled_top_y))
        jsx_content = jsx_content.replace("{{SCALED_CENTER_X}}", str(scaled_center_x))

        # 写入临时JSX
        temp_dir = tempfile.gettempdir()
        temp_jsx = os.path.join(temp_dir, "temp_place_design.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx_content)

        psApp.DoJavaScriptFile(temp_jsx)
        print(f"✅ 生成: {output_path}")

    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()


def classify_design(filename):
    """根据文件名分类设计图"""
    base = os.path.splitext(filename)[0]

    if "_BW_" in base or "_BW" in base:
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
    """处理单个 DX 文件夹"""
    dx_name = os.path.basename(dx_folder)
    rem_bg_folder = os.path.join(dx_folder, "02_REM_BG")
    upload_folder = os.path.join(dx_folder, "03_UPLOAD")

    if not os.path.exists(rem_bg_folder):
        return

    os.makedirs(upload_folder, exist_ok=True)

    for file in os.listdir(rem_bg_folder):
        if not file.lower().endswith(".png"):
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


def main():
    dx_root = config.SOURCE_BASE
    for folder in os.listdir(dx_root):
        if folder.startswith("DX"):
            dx_folder = os.path.join(dx_root, folder)
            if os.path.isdir(dx_folder):
                print(f"\n===== 处理: {dx_folder} =====")
                process_dx_folder(dx_folder)


if __name__ == "__main__":
    main()
