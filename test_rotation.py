"""验证 PS 贴图旋转方向和角度
生成 5 张对比图：0°、+1°、-1°、+3°、-3°，方便肉眼确认方向。
正数在 Photoshop DOM 里表示顺时针（向右旋转）。
"""
import os
import sys
sys.path.insert(0, r"E:\Claude code\ps")
import config
from wb_sticker_ps import place_design

# 选一个 BW 文件做测试
dx_name = "DX0086"
design_path = os.path.join(config.SOURCE_BASE, dx_name, "02_REM_BG", f"{dx_name}_BW_cut.png")
upload_folder = os.path.join(config.SOURCE_BASE, dx_name, "03_UPLOAD")
os.makedirs(upload_folder, exist_ok=True)

if not os.path.exists(design_path):
    print(f"测试文件不存在: {design_path}")
    sys.exit(1)

print("=== PS 贴图旋转方向验证 ===")
print(f"测试文件: {design_path}")
print("正数 = Photoshop 顺时针（向右旋转）")
print("负数 = Photoshop 逆时针（向左旋转）\n")

base_cfg = dict(config.FRONT_NEW)
torso_white = os.path.join(config.BASE_TORSO, base_cfg["torso_white"])

for angle in (0, 1, -1, 3, -3):
    cfg = dict(base_cfg)
    cfg["rotation"] = angle
    out_path = os.path.join(upload_folder, f"{dx_name}_W_白T_旋转{angle}度.jpg")
    print(f"--- 生成旋转 {angle}° -> {out_path} ---")
    place_design(design_path, torso_white, out_path, cfg)

print("\n✅ 5 张对比图已生成，请打开查看旋转方向。")
print("如果实际方向与标记相反，说明 Photoshop DOM 的 rotate() 方向与预期不同，需要把配置里的 rotation 取反。")
