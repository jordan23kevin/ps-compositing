
"""WB贴图 PS 配置 — 路径、坐标、缩放、旋转"""

# 胚衣路径
BASE_TORSO = r"D:\Semems\1胚衣"
SOURCE_BASE = r"D:\Semems WB\02_PROJECTS"

# 图像参数
ALPHA_THRESHOLD = 20
MIN_FILE_SIZE = 50 * 1024  # 50KB

# PS 画布尺寸（固定：1340×1785）
CANVAS_WIDTH = 1340
CANVAS_HEIGHT = 1785

# 背图参数
BACK = {
    "torso_white": "白背2.jpg",
    "torso_black": "黑背2.jpg",
    "target_height": 498,   # 设计图缩放目标高度 (px)
    "center_x": 718,        # 图层中心 X
    "center_y": 808,        # 图层中心 Y
    "rotation": 0,          # 旋转角度 (度)
}

# 正图参数
FRONT = {
    "torso_white": "白正2.jpg",
    "torso_black": "黑正2.jpg",
    "target_height": 400,   # 设计图缩放目标高度 (px)
    "center_x": 850,        # 图层中心 X
    "center_y": 600,        # 图层中心 Y
    "rotation": 0,          # 旋转角度 (度)
}

# 分类逻辑
def get_type(filename):
    """从文件名获取类型（复刻自原 utils/sticker_typer.py）"""
    import os
    stem = os.path.splitext(os.path.basename(filename))[0]
    s = stem.upper()
    if "BW" in s or "WB" in s:
        return "both"
    elif "B" in s:
        return "back"
    elif "W" in s:
        return "front"
    last = stem[-1] if stem else ""
    if last == "1":
        return "back"
    if last == "2":
        return "front"
    return None

