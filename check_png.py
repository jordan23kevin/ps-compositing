from PIL import Image
import numpy as np

path = r"D:\Semems WB\02_PROJECTS\DX0086\02_REM_BG\DX0086_BW_cut.png"
img = Image.open(path)
print(f"模式: {img.mode}")
print(f"尺寸: {img.size}")

a = np.array(img.convert("RGBA"))
print(f"shape: {a.shape}")

# 检查alpha通道
alpha = a[:, :, 3]
print(f"\nalpha通道统计:")
print(f"  最小值: {alpha.min()}")
print(f"  最大值: {alpha.max()}")
print(f"  平均值: {alpha.mean():.2f}")
print(f"  alpha=0 的像素数: {(alpha == 0).sum()}")
print(f"  alpha>0 的像素数: {(alpha > 0).sum()}")
print(f"  alpha>=20 的像素数: {(alpha >= 20).sum()}")

# 检查RGB通道
r = a[:, :, 0]
g = a[:, :, 1]
b = a[:, :, 2]
print(f"\nRGB通道统计:")
print(f"  R: min={r.min()}, max={r.max()}, mean={r.mean():.2f}")
print(f"  G: min={g.min()}, max={g.max()}, mean={g.mean():.2f}")
print(f"  B: min={b.min()}, max={b.max()}, mean={b.mean():.2f}")

# 看看四个角的颜色
print(f"\n四个角的颜色 (R,G,B,A):")
print(f"  左上: {a[0, 0]}")
print(f"  右上: {a[0, -1]}")
print(f"  左下: {a[-1, 0]}")
print(f"  右下: {a[-1, -1]}")
print(f"  中心: {a[1024, 1024]}")
