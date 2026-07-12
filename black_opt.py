# ===== 黑衫平铺图贴花优化 v1.0.0（纯软件，自 check_rem.py 抽取） =====
# 作用：把「通用 / 白版」设计图贴在黑衫上时，自动做
#   1) 自适应白墨打底（add_white_underbase）：越暗白墨越厚，使暗部在黑布上显色
#   2) 暗部智能提亮 + 饱和补偿（enhance_dark_print_for_black_shirt）
# 解决：通用图直接贴黑衫时半透明边缘/纹理与黑色混合导致「变暗 / 发脏」。
# 仅在贴在黑胚衣时使用（wb_sticker_ps.place_design(black_optimize=True)）。
#
# 依赖 cv2（生产环境来自 E:\python_packages）。若当前环境缺 cv2，自动跳过优化
# 并返回原图，保证贴图流水线绝不崩溃。
import sys

try:
    import cv2
except Exception:
    try:
        sys.path.insert(0, r"E:\python_packages")
        import cv2
    except Exception:
        cv2 = None
CV2_OK = cv2 is not None

import numpy as np
from PIL import Image


def add_white_underbase(design, max_white_opacity=0.9, min_white_opacity=0.05,
                        transition_threshold=130, edge_feather=5, boost_sat=0.35):
    """自适应浓度白墨打底（黑衫显色）：越暗白墨越厚，亮/饱和色保留原色。
    模拟真实 DTG 黑衫印花『先喷白墨、再喷彩色』，使极暗区域在黑布上可见，
    同时避免全铺白底导致颜色漂白。design: PIL RGBA 图。返回 PIL RGBA 图。
    """
    if not CV2_OK:
        return design
    if getattr(design, "mode", None) != "RGBA":
        design = design.convert("RGBA")
    arr = np.array(design)
    rgb = arr[..., :3].astype(np.float32)
    alpha = arr[..., 3].astype(np.float32) / 255.0
    valid = alpha > 0.01
    if not valid.any():
        return design
    rgb_u = rgb.astype(np.uint8)
    gray = cv2.cvtColor(rgb_u, cv2.COLOR_RGB2GRAY).astype(np.float32) / 255.0
    hsv = cv2.cvtColor(rgb_u, cv2.COLOR_RGB2HSV).astype(np.float32)
    sat = hsv[..., 1] / 255.0
    th = transition_threshold / 255.0
    white_mask = np.clip((th - gray) / th, 0.0, 1.0)            # 越暗越 1
    white_alpha = min_white_opacity + white_mask * (max_white_opacity - min_white_opacity)
    white_alpha = white_alpha * alpha
    if edge_feather > 0:
        k = int(edge_feather) * 2 + 1
        white_alpha = cv2.GaussianBlur(white_alpha, (k, k), 0)
    white_alpha = np.clip(white_alpha, 0.0, 1.0)
    white = np.ones_like(rgb) * 255.0
    keep = np.clip(gray + boost_sat * sat, 0.0, 1.0)            # 暗但饱和的颜色多保留原色
    color_on_white = white * (1 - keep)[..., None] + rgb * keep[..., None]
    final = rgb * (1 - white_alpha)[..., None] + color_on_white * white_alpha[..., None]
    final = np.clip(final, 0, 255).astype(np.uint8)
    out = np.dstack([final, arr[..., 3]])
    return Image.fromarray(out.astype(np.uint8), "RGBA")


def enhance_dark_print_for_black_shirt(design, shirt="black",
        dark_boost=0.55, protect_threshold=140, min_brightness=20,
        sat_compensation=0.3, smooth_radius=9):
    """黑衫/白衫智能显色：仅处理暗部(黑衫)或亮部(白衫)，完整保留颜色。
    替代旧版『非透明像素变纯白/纯黑』剪影。shirt='black' 提亮暗部；shirt='white' 压暗亮部。
    design: PIL RGBA 图。返回 PIL RGBA 图。
    """
    if not CV2_OK:
        return design
    if getattr(design, "mode", None) != "RGBA":
        design = design.convert("RGBA")
    arr = np.array(design)
    rgb = arr[..., :3].astype(np.float32)
    alpha = arr[..., 3].astype(np.float32) / 255.0
    valid = alpha > 0.01
    if not valid.any():
        return design
    lab = cv2.cvtColor(rgb.astype(np.uint8), cv2.COLOR_RGB2LAB).astype(np.float32)
    L, A, B = lab[..., 0], lab[..., 1], lab[..., 2]
    L_norm = L / 255.0
    thresh_norm = protect_threshold / 255.0
    if shirt == "black":
        weight = np.clip((thresh_norm - L_norm) / thresh_norm, 0.0, 1.0)
        target_gamma = max(0.15, 1.0 - dark_boost * 0.7)
        L_target = 255.0 * np.power(L_norm, target_gamma)
        min_mask = (L < min_brightness) & valid
        L_target[min_mask] = np.maximum(L_target[min_mask], min_brightness)
    else:
        weight = np.clip((L_norm - thresh_norm) / (1.0 - thresh_norm), 0.0, 1.0)
        target_gamma = max(0.15, 1.0 + dark_boost * 0.7)
        L_target = 255.0 * np.power(L_norm, target_gamma)
        max_mask = (L > (255 - min_brightness)) & valid
        L_target[max_mask] = np.minimum(L_target[max_mask], 255 - min_brightness)
    weight = weight * alpha
    if smooth_radius > 0:
        sr = int(smooth_radius) | 1
        weight = cv2.GaussianBlur(weight, (sr, sr), 0)
    weight = np.clip(weight, 0.0, 1.0)
    L_final = L * (1 - weight) + L_target * weight
    sat_gain = 1.0 + sat_compensation * weight
    A_final = np.clip((A - 128.0) * sat_gain + 128.0, 0.0, 255.0)
    B_final = np.clip((B - 128.0) * sat_gain + 128.0, 0.0, 255.0)
    lab_final = np.stack([L_final, A_final, B_final], axis=-1).astype(np.uint8)
    bgr = cv2.cvtColor(lab_final, cv2.COLOR_LAB2BGR)
    rgb_final = cv2.cvtColor(bgr, cv2.COLOR_BGR2RGB)
    out = np.dstack([rgb_final, arr[..., 3]])
    return Image.fromarray(out.astype(np.uint8), "RGBA")


def black_shirt_print_optimize(design):
    """黑衫终极显色流水线：自适应白墨打底 + 轻度暗部提亮 + 饱和补偿。

    与 check_rem.py 的「反黑」按钮使用完全相同的参数，保证平铺图与模特图
    黑衫成品观感一致。
    """
    if not CV2_OK:
        print("  ⚠️ cv2 不可用，跳过黑衫优化（设计将按原样贴到黑衫，可能出现变暗）")
        return design
    step1 = add_white_underbase(design)
    step2 = enhance_dark_print_for_black_shirt(
        step1, shirt="black", dark_boost=0.25, protect_threshold=160,
        min_brightness=20, sat_compensation=0.2
    )
    return step2
