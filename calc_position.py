from PIL import Image
import numpy as np

ALPHA_THRESHOLD = 20

def calculate_sticker_position(png_path):
    """
    计算贴图的关键位置：
    1. 最高有效像素的 y 坐标
    2. 中心点的 x 坐标
    3. 贴图原始尺寸
    """
    img = Image.open(png_path).convert("RGBA")
    a = np.array(img)
    h, w = a.shape[0], a.shape[1]
    alpha = a[:, :, 3]
    
    # 找出所有 alpha >= 20 的有效像素
    mask = alpha >= ALPHA_THRESHOLD
    
    # 计算最高有效像素的 y 坐标
    top_y = None
    for y in range(h):
        if np.any(mask[y, :]):
            top_y = y
            break
    
    if top_y is None:
        top_y = 0
    
    # 计算有效像素的中心点 x 坐标
    ys, xs = np.where(mask)
    if len(xs) == 0:
        center_x = w / 2
    else:
        center_x = np.mean(xs)
    
    return {
        "top_y": top_y,
        "center_x": center_x,
        "original_width": w,
        "original_height": h
    }


if __name__ == "__main__":
    # 测试 DX0086
    test_path = r"D:\Semems WB\02_PROJECTS\DX0086\02_REM_BG\DX0086_BW_cut.png"
    pos = calculate_sticker_position(test_path)
    print(f"贴图原始尺寸: {pos['original_width']} x {pos['original_height']}")
    print(f"最高有效像素 y: {pos['top_y']}")
    print(f"有效像素中心点 x: {pos['center_x']}")
