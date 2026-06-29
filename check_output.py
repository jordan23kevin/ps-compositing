from PIL import Image
import numpy as np

# 检查生成的输出图
output_path = r"D:\Semems WB\02_PROJECTS\DX0086\03_UPLOAD\DX0086_B_白T.jpg"

# 检查胚衣原图
torso_path = r"D:\Semems\1胚衣\白背2.jpg"

out_img = Image.open(output_path)
torso_img = Image.open(torso_path)

print(f"输出图: {out_img.size}, 模式: {out_img.mode}")
print(f"胚衣图: {torso_img.size}, 模式: {torso_img.mode}")

out_arr = np.array(out_img.convert("RGB"))
torso_arr = np.array(torso_img.convert("RGB"))

# 计算差异
diff = np.abs(out_arr.astype(int) - torso_arr.astype(int))
print(f"\n输出图与胚衣的差异:")
print(f"  最大差异: {diff.max()}")
print(f"  平均差异: {diff.mean():.2f}")
print(f"  有差异的像素数 (>10): {(diff.sum(axis=2) > 10).sum()}")

if diff.max() < 10:
    print("\n❌ 输出图和胚衣完全一样！印花没有被贴上去！")
else:
    print("\n✅ 输出图和胚衣有差异，印花应该贴上去了")
