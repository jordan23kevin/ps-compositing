import os
import win32com.client
import pythoncom
import tempfile
from calc_position import calculate_sticker_position
import config

def place_back(design_path, torso_path, output_path, scaled_top_y, scaled_center_x):
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        
        # 读取并替换JSX
        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_back_new.jsx")
        with open(jsx_path, "r", encoding="utf-8") as f:
            jsx_content = f.read()
        
        jsx_content = jsx_content.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{DESIGN_PATH}}", design_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx_content = jsx_content.replace("{{SCALED_TOP_Y}}", str(scaled_top_y))
        jsx_content = jsx_content.replace("{{SCALED_CENTER_X}}", str(scaled_center_x))
        
        # 写入临时JSX
        temp_dir = tempfile.gettempdir()
        temp_jsx = os.path.join(temp_dir, "temp_place_back_new.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as f:
            f.write(jsx_content)
        
        # 运行JSX
        psApp.DoJavaScriptFile(temp_jsx)
        
        print(f"✅ 生成: {output_path}")
        
    except Exception as e:
        print(f"❌ 错误: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()


def test_dx0086_back():
    dx_name = "DX0086"
    design_path = os.path.join(config.SOURCE_BASE, dx_name, "02_REM_BG", f"{dx_name}_BW_cut.png")
    upload_folder = os.path.join(config.SOURCE_BASE, dx_name, "03_UPLOAD")
    os.makedirs(upload_folder, exist_ok=True)
    
    # 计算位置
    pos = calculate_sticker_position(design_path)
    scale = 0.30
    scaled_top_y = pos["top_y"] * scale
    scaled_center_x = pos["center_x"] * scale
    print(f"计算结果：缩放后最高有效y={scaled_top_y}, 中心点x={scaled_center_x}")
    
    # 白t背图
    print("\n--- 白T背图 ---")
    place_back(
        design_path,
        os.path.join(config.BASE_TORSO, config.BACK["torso_white"]),
        os.path.join(upload_folder, f"{dx_name}_B_白T.jpg"),
        scaled_top_y, scaled_center_x
    )
    
    # 黑t背图
    print("\n--- 黑T背图 ---")
    place_back(
        design_path,
        os.path.join(config.BASE_TORSO, config.BACK["torso_black"]),
        os.path.join(upload_folder, f"{dx_name}_B_黑T.jpg"),
        scaled_top_y, scaled_center_x
    )


if __name__ == "__main__":
    test_dx0086_back()
