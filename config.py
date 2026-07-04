"""WB贴图 PS 配置 — 路径、坐标、缩放、旋转"""
import re

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
    "rotation": -1,           # 向左旋转1°（逆时针）
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
    "rotation": -1,           # 向左旋转1°（逆时针）
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

def _find_ps_hwnd():
    """查找 Photoshop 主窗口句柄。"""
    try:
        import win32gui
        hwnd = win32gui.FindWindow("Photoshop", None)
        if hwnd:
            return hwnd
        # fallback：枚举可见窗口，标题含 Photoshop
        handles = []
        def _enum(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if "Photoshop" in title:
                    extra.append(hwnd)
        win32gui.EnumWindows(_enum, handles)
        return handles[0] if handles else 0
    except Exception:
        return 0


def show_ps_minimized(ps_app=None):
    """让 Photoshop 窗口可见并最小化到任务栏，全程保持最小化。
    如果窗口已经存在且已最小化，则不再重复设置 Visible（避免把窗口恢复出来）。
    ps_app: 已获取的 Photoshop COM Application 对象（可选）"""
    try:
        import win32gui
        import win32con
        hwnd = _find_ps_hwnd()
        if hwnd and win32gui.IsIconic(hwnd):
            # 已经最小化，不要再次 Visible=True，否则会把窗口 restore 出来
            return
        if hwnd and not win32gui.IsIconic(hwnd):
            # 窗口已存在但没最小化，直接最小化
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            return
        # 没有窗口：让 COM 创建窗口并最小化
        if ps_app is not None:
            ps_app.Visible = True
        import time
        time.sleep(0.3)
        hwnd = _find_ps_hwnd()
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    except Exception:
        pass


def hide_ps_window(ps_app=None):
    """让 Photoshop 主窗口完全隐藏（后台运行，不显示在屏幕/任务栏）。
    适合配合 Web 页面状态摘要使用。
    ps_app: 已获取的 Photoshop COM Application 对象（可选）"""
    try:
        import win32gui
        import win32con
        hwnd = _find_ps_hwnd()
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            return
        # 没有窗口：让 COM 创建窗口后再隐藏
        if ps_app is not None:
            ps_app.Visible = True
        import time
        time.sleep(0.3)
        hwnd = _find_ps_hwnd()
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    except Exception:
        pass


# 文件名后缀解析：支持版本号，如 B2 / W1 / BW3 / WB2
_SIDE_RE = re.compile(r'^(BW|WB|B|W)(\d*)$', re.IGNORECASE)


def parse_side_suffix(suffix):
    """解析去背图后缀中的正反面与版本号。
    例如：'B2' -> ('B', '2')；'BW' -> ('BW', '')；'WB1' -> ('WB', '1')。
    无法解析时返回 (None, None)。"""
    m = _SIDE_RE.match(str(suffix))
    if not m:
        return None, None
    return m.group(1).upper(), m.group(2)


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
