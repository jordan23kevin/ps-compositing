import os
import sys
sys.path.insert(0, r"E:\Claude code\ps")
import config
from wb_sticker_ps import place_design

# 测试 DX0086 正图（白T + 黑T）
dx_name = "DX0086"
design_path = os.path.join(config.SOURCE_BASE, dx_name, "02_REM_BG", f"{dx_name}_BW_cut.png")
upload_folder = os.path.join(config.SOURCE_BASE, dx_name, "03_UPLOAD")
os.makedirs(upload_folder, exist_ok=True)

print("=== 测试正图（W）新方案 ===")
print(f"缩放: {config.FRONT_NEW['scale_percent']}%")
print(f"旋转: {config.FRONT_NEW['rotation']}°")
print(f"目标中心点X: {config.FRONT_NEW['target_center_x']}")
print(f"目标最高有效像素Y: {config.FRONT_NEW['target_top_y']}")

# 白T正图
print("\n--- 白T正图 ---")
place_design(
    design_path,
    os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_white"]),
    os.path.join(upload_folder, f"{dx_name}_W_白T.jpg"),
    config.FRONT_NEW,
)

# 黑T正图
print("\n--- 黑T正图 ---")
place_design(
    design_path,
    os.path.join(config.BASE_TORSO, config.FRONT_NEW["torso_black"]),
    os.path.join(upload_folder, f"{dx_name}_W_黑T.jpg"),
    config.FRONT_NEW,
)
