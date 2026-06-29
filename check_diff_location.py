from PIL import Image
import numpy as np

output_path = r"D:\Semems WB\02_PROJECTS\DX0086\03_UPLOAD\DX0086_B_白T.jpg"
torso_path = r"D:\Semems\1胚衣\白背2.jpg"

out_img = Image.open(output_path).convert("RGB")
torso_img = Image.open(torso_path).convert("RGB")

out_arr = np.array(out_img)
torso_arr = np.array(torso_img)

diff = np.abs(out_arr.astype(int) - torso_arr.astype(int))
diff_sum = diff.sum(axis=2)

# 找出有差异的像素的位置
ys, xs = np.where(diff_sum > 10)
if len(ys) > 0:
    print(f"有差异的像素数: {len(ys)}")
    print(f"差异区域 X 范围: {xs.min()} ~ {xs.max()}")
    print(f"差异区域 Y 范围: {ys.min()} ~ {ys.max()}")
    print(f"差异区域中心: ({xs.mean():.0f}, {ys.mean():.0f})")
else:
    print("没有差异")
