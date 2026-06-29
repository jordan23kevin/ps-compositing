import os
import win32com.client
import pythoncom
import config
from PIL import Image
import tempfile


def make_black_transparent(input_path, output_path):
    """将 PNG 的黑色背景转为透明"""
    img = Image.open(input_path).convert("RGBA")
    datas = img.getdata()

    new_data = []
    for item in datas:
        # 判断黑色（RGB 都接近0）
        if item[0] < 20 and item[1] < 20 and item[2] < 20:
            # 设为透明
            new_data.append((item[0], item[1], item[2], 0))
        else:
            new_data.append(item)

    img.putdata(new_data)
    img.save(output_path, "PNG")


def place_design(design_path, torso_path, output_path, placement, is_front=True):
    """使用 JSX 脚本将设计贴到T恤上"""
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")

        # 先预处理黑色背景为透明
        temp_dir = tempfile.gettempdir()
        temp_design = os.path.join(temp_dir, "temp_design_transparent.png")
        make_black_transparent(design_path, temp_design)

        # 选择 JSX 脚本
        if is_front:
            jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_front.jsx")
        else:
            jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_back.jsx")

        # 读取 JSX 内容，替换参数
        with open(jsx_path, "r", encoding="utf-8") as f:
            jsx_content = f.read()

        # 替换占位符
        jsx_content = jsx_content.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{DESIGN_PATH}}", temp_design.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{TARGET_HEIGHT}}", str(placement["target_height"]))
        jsx_content = jsx_content.replace("{{CENTER_X}}", str(placement["center_x"]))
        jsx_content = jsx_content.replace("{{CENTER_Y}}", str(placement["center_y"]))
        jsx_content = jsx_content.replace("{{ROTATION}}", str(placement["rotation"]))

        # 写入临时 JSX
        temp_jsx = os.path.join(temp_dir, "temp_place.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx_content)

        # 运行 JSX
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
        return "W"
    elif base.endswith("_2"):
        return "B"
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
        
        # 提取基础 DX 名称（去掉 _cut.png 等后缀）
        base_name = dx_name  # 直接用 DXxxxx

        if design_type == "BW":
            # ===== BW 类型：生成 W 和 B 两套文件，供 ps_batch.py 合成 =====
            # W 套（正面）
            print("  → 生成 W 正面文件...")
            place_design(design_path, os.path.join(config.BASE_TORSO, config.FRONT["torso_white"]),
                         os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"), config.FRONT, is_front=True)
            place_design(design_path, os.path.join(config.BASE_TORSO, config.FRONT["torso_black"]),
                         os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"), config.FRONT, is_front=True)
            
            # B 套（背面）
            print("  → 生成 B 背面文件...")
            place_design(design_path, os.path.join(config.BASE_TORSO, config.BACK["torso_white"]),
                         os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"), config.BACK, is_front=False)
            place_design(design_path, os.path.join(config.BASE_TORSO, config.BACK["torso_black"]),
                         os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"), config.BACK, is_front=False)
            print("  ✅ BW 准备完成，可运行 ps_batch.py 合成最终 BW 图！")
        
        elif design_type == "W":
            # ===== W 类型：只生成正面 =====
            place_design(design_path, os.path.join(config.BASE_TORSO, config.FRONT["torso_white"]),
                         os.path.join(upload_folder, f"{base_name}_W_白T.jpg"), config.FRONT, is_front=True)
            place_design(design_path, os.path.join(config.BASE_TORSO, config.FRONT["torso_black"]),
                         os.path.join(upload_folder, f"{base_name}_W_黑T.jpg"), config.FRONT, is_front=True)
        
        elif design_type == "B":
            # ===== B 类型：只生成背面 =====
            place_design(design_path, os.path.join(config.BASE_TORSO, config.BACK["torso_white"]),
                         os.path.join(upload_folder, f"{base_name}_B_白T.jpg"), config.BACK, is_front=False)
            place_design(design_path, os.path.join(config.BASE_TORSO, config.BACK["torso_black"]),
                         os.path.join(upload_folder, f"{base_name}_B_黑T.jpg"), config.BACK, is_front=False)


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
