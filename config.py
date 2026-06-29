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

# ===== 背图新方案参数 =====
# 复刻美图秀秀的逻辑：
# 1. 置入贴图，缩放30%，旋转1°
# 2. 贴图最高有效像素 → PS Y=570
# 3. 贴图中心点 → PS X=680
BACK_NEW = {
    "torso_white": "白背2.jpg",
    "torso_black": "黑背2.jpg",
    "scale_percent": 30,      # 缩放30%
    "rotation": 1,            # 向右旋转1°
    "target_center_x": 680,   # 贴图中心点目标X
    "target_top_y": 570,      # 贴图最高有效像素目标Y
}

# ===== 正图新方案参数 =====
# 和背图逻辑一样，仅参数不同：
# 1. 置入贴图，缩放10%，旋转1°
# 2. 贴图最高有效像素不高于Y=612
# 3. 贴图中心点 → PS X=888
# 注：实测手动"置入后设10%"等效于相对原始尺寸13.33%
FRONT_NEW = {
    "torso_white": "白正2.jpg",
    "torso_black": "黑正2.jpg",
    "scale_percent": 13.33,    # 实测等效于手动置入后设10%
    "rotation": 1,            # 向右旋转1°
    "target_center_x": 888,   # 贴图中心点目标X
    "target_top_y": 612,      # 贴图最高有效像素目标Y
}

# ===== 旧版背图参数（保留，供BW两阶段流程的ps_batch.py使用） =====
BACK = {
    "torso_white": "白背2.jpg",
    "torso_black": "黑背2.jpg",
    "target_height": 498,
    "center_x": 718,
    "center_y": 808,
    "rotation": 0,
}

# 正图参数（待调整）
FRONT = {
    "torso_white": "白正2.jpg",
    "torso_black": "黑正2.jpg",
    "target_height": 400,
    "center_x": 850,
    "center_y": 600,
    "rotation": 0,
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
