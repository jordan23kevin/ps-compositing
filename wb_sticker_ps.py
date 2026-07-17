# ===== WB 贴图主控 v2.5.2（纯软件，不再依赖 Photoshop） =====
# 变更 v2.5.4（2026-07-15）：
#   - place_design 五参缩放从「硬拉宽度到 width」改为「等比适应(contain)」：
#     width/height 视为一个框，设计图最长边不超过框，另一边同比例缩放。
#     即 scale = min(width/tw, height/th)。
#   - 效果：又高又大的设计（如 DX0694BW 背面 CHOSEN 1099×1894 缩到 317×546）
#     不再把后背占满；宽图仍按宽度缩，高图按高度缩。
#   - 当 meta 中 height=0 时仍退回到旧行为 width-only，保证单面款/历史五参兼容。
# 变更 v2.5.3（2026-07-15）：
#   - 修复 classify_design 对「拆开命名」双面款(BW)的误判：实际 cut 文件为
#     DXxxxxBW_W_cut.png / DXxxxxBW_B_cut.png（款号本身带 BW，BW 前无下划线），
#     原判定 "_BW" in base 永远不匹配，导致被误判成单面款(W/B)，正面黑W11/白W11
#     始终走顶层①(546) 而非 bw ②(245) —— 即「双面款尺寸不对」根因。
#   - 改为：款号首段 endswith BW/WB 即视为双面款，其 W 面 cut → "BW"（正面用 bw ②）、
#     B 面 cut → "B"（背面用顶层①）；保留旧 "_BW"/"_WB" 单 cut 命名兼容。
#   - process_dx_folder 的 BW 分支增加 is_split_w 判定：拆开的 W 面 cut 只生成正面
#     （bw ②），背面由 B 面 cut 负责，避免背面被正面图覆盖。
#   - 效果：BW 款贴图时，黑W11.jpg 与 白W11.jpg 各自读取自己的 "bw" ② 五参。
# 变更 v2.5.2（2026-07-13）：
#   - 正面胚衣（W白/白W11、W黑/黑W11）支持「双组五参」：素材库 .meta.json 顶层五参
#     用于① 单面款（只有W贴图），新增 "bw" 子块五参用于② 双面款（有W+B贴图）。
#   - process_dx_folder 按产品类型自动选组：单面款(W)用顶层、双面款(BW)正面用 bw 块；
#     双面款缺 bw 第二组直接报错（不偷偷用第一组兜底，强制两组都填）。
#   - 背面胚衣（B白/B黑）只一组，行为不变。
# 变更 v2.5.1（2026-07-12）：
#   - 黑胚衣落点 black_optimize 改回 False（用户确认黑衫就要「不加白墨打底」的
#     效果，与其批准的 DX0648_W黑T_五参黑W11 / _v_现在代码(topcenter,-meta) 一致）。
#     白墨打底（v2.4.1）保留在 black_opt.py，仅当用户后续要求「防黑衫发暗」时按需开启。
# 变更 v2.5.0（2026-07-12）：
#   - 平铺图定位改为「素材库五参驱动」：place_design 优先读取胚衣同名
#     .meta.json（width/height/rotation/highest_y/center_x），按素材库 黑W11/白W11/
#     白B12/黑B7 的真实尺寸贴图，彻底替代 config.FRONT_NEW/BACK_NEW 写死的
#     scale/center（旧 13.33% 把图缩成胸口小标、位置偏右）。
#   - config 新增 MATERIAL_BASE / FLAT_TORSO / flat_torso() / load_meta()；
#     process_dx_folder 经 flat_torso 取底图 + 五参，黑胚衣落点按本版默认 black_optimize=False。
#   - 经像素比对，素材库胚衣与旧 1胚衣 为同一件衣服（JPEG 重编码差 <1/255），
#     切换底图不改变成品外观，仅修正定位。
# 变更 v2.4.1（2026-07-12）：
#   - StickerSession.place_design 新增 black_optimize 开关：贴在黑胚衣时自动调用
#     black_opt.black_shirt_print_optimize（白墨打底 + 暗部提亮），解决通用/白版
#     设计图直接贴黑衫时半透明边缘/纹理与黑色混合导致的「变暗 / 发脏」问题。
#   - process_dx_folder 中 4 个黑胚衣落点（W黑T×2、B黑T×2）均传入 black_optimize=True；
#     白胚衣落点不受影响。已有 _黑*_cut.png 专用图的款仍走 process_black.py（已优化）。
#   - 新增 black_opt.py（自 check_rem.py 抽取黑衫优化，带 cv2 兜底，缺 cv2 时自动跳过）。
# 变更 v2.4.0（2026-07-12）：
#   - StickerSession.place_design 改为纯 PIL 实现（trim→缩放→平移→绕中心旋转→normal 合成），
#     复刻原 place_design.jsx 定位；移除 win32com/pythoncom，不再连接 Photoshop。
#   - 黑T优先使用 02_REM_BG 中的 _黑B/_黑W/_黑BW 专用文件（逻辑不变）。
#   - 检测到黑版专用文件时，通用图不再输出黑T成品，仅输出白T（逻辑不变）。
import os
import re
import sys
import io
import time
import tempfile
from pathlib import Path
from PIL import Image
import numpy as np
import config

# 确保 Windows 控制台能输出中文/emoji，同时防止桥接进程的 stdout 句柄已关闭导致 print 崩溃
class _SafeStream:
    def __init__(self, stream):
        self._stream = stream

    def write(self, s):
        try:
            self._stream.write(s)
        except Exception:
            pass

    def flush(self):
        try:
            self._stream.flush()
        except Exception:
            pass

    def __getattr__(self, name):
        return getattr(self._stream, name)


sys.stdout = _SafeStream(io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8'))

sys.path.insert(0, r"E:\Claude code\ps")
sys.path.insert(0, r"D:\Semems WB\04_OS\engine")
import wb_naming  # 命名规则唯一出处

try:
    import wb_meta
except Exception:
    wb_meta = None

ALPHA_THRESHOLD = 20

VERSION = "2.5.5"

# ---------------------------------------------------------------------------
# 元数据辅助（读取 _cut.png sidecar，为上传图注册）
# ---------------------------------------------------------------------------
_MIGRATED_DX = set()


def _role_from_name(name):
    """从文件名推断 role（规则见 wb_naming.role_from_name）"""
    return wb_naming.role_from_name(name)


def _infer_meta(path):
    """sidecar 完全缺失时的文件名兜底推断"""
    name = os.path.basename(path)
    dx = name.split("_")[0] if "_" in name else "DX"
    role = _role_from_name(name)
    uid = f"UID_{dx}_{re.sub(r'[^A-Za-z0-9]', '_', role)}"
    group_id = f"G_{dx}_{role}"
    return {"uid": uid, "group_id": group_id, "role": role, "stage": "unknown"}


def _get_cut_meta(cut_path):
    """读取去背图 sidecar；缺失时尝试 migrate_dx，最后兜底推断"""
    if wb_meta is None:
        return None
    dx_dir = str(Path(cut_path).parent.parent)
    meta = wb_meta.read_meta(cut_path)
    if meta:
        return meta
    if dx_dir not in _MIGRATED_DX:
        try:
            wb_meta.migrate_dx(dx_dir)
        except Exception as e:
            print(f"  ⚠️ migrate_dx 失败 {dx_dir}: {e}")
        _MIGRATED_DX.add(dx_dir)
    meta = wb_meta.read_meta(cut_path)
    if meta:
        return meta
    return _infer_meta(cut_path)


def calculate_sticker_position(png_path):
    """
    计算贴图的关键位置（复刻美图秀秀 calc_offset_back 逻辑）：
    1. 最高有效像素的 y 坐标：从上往下找，找到一行里有效像素≥50个的那行
       （防止顶部噪点干扰）
    2. 中心点 x 坐标：透明度加权质心 (x * alpha).sum() / alpha.sum()
    """
    img = Image.open(png_path).convert("RGBA")
    a = np.array(img)
    h, w = a.shape[0], a.shape[1]
    alpha = a[:, :, 3].astype(np.float64)

    mask = alpha >= ALPHA_THRESHOLD
    ys, xs = np.where(mask)

    if len(ys) == 0:
        top_y = 0
        center_x = w / 2
    else:
        # 最高有效像素y：从最小y开始往下找，找到一行里有效像素≥50个的那行
        y_min = int(ys.min())
        for y in range(y_min, min(y_min + 200, h)):
            if (alpha[y] >= ALPHA_THRESHOLD).sum() >= 50:
                y_min = y
                break
        top_y = y_min

        # 中心点x：透明度加权质心
        weights = np.where(mask, alpha, 0.0)
        center_x = float((np.indices(alpha.shape)[1] * weights).sum() / weights.sum())

    return {
        "top_y": float(top_y),
        "center_x": float(center_x),
        "original_width": w,
        "original_height": h,
    }


# 保存图片，确保文件大小不超过 max_size_mb（默认 2MB）
# 如果初始质量下超过限制，逐步降低 quality 直到满足，分辨率始终不变
_SAVE_QUALITIES = [95, 90, 85, 80, 75, 70]


def save_with_size_limit(img, output_path, max_size_mb=2.0):
    for q in _SAVE_QUALITIES:
        img.save(output_path, "JPEG", quality=q, optimize=True)
        size_mb = os.path.getsize(output_path) / (1024 * 1024)
        if size_mb <= max_size_mb:
            if q < _SAVE_QUALITIES[0]:
                print(f"  ⚠️ 质量降至 {q} 以满足 < {max_size_mb}MB: {os.path.basename(output_path)}")
            return
    print(f"  ⚠️ 警告: 即使 quality={_SAVE_QUALITIES[-1]} 仍超过 {max_size_mb}MB: {os.path.basename(output_path)}")


# ---------------------------------------------------------------------------
# 平铺图贴花会话：纯软件实现（不再依赖 Photoshop）
# 复刻原 place_design.jsx 的 affine 变换：trim → 缩放 → 平移 → 绕中心旋转 → normal 合成
# ---------------------------------------------------------------------------
class StickerSession:
    def __init__(self):
        # 纯软件实现，无需连接 Photoshop
        self.temp_dir = tempfile.gettempdir()

    def place_design(self, design_path, torso_path, output_path, placement_cfg=None, meta=None, cut_meta=None, black_optimize=False):
        """纯软件贴花，接口与原 Photoshop 版完全一致。

        定位参数：优先用 meta（素材库五参 width/height/rotation/highest_y/center_x）；
        未传 meta 时回退 placement_cfg（config.FRONT_NEW/BACK_NEW，供 process_black/white 旧线）。

        变换顺序与 JSX(place_design.jsx) 对齐：
          1. 按 ALPHA_THRESHOLD 裁剪设计图透明边距
          2. 按 scale 缩放（meta: width/设计宽；cfg: scale_percent）
          3. 平移 move_x = target_center_x - new_w/2, move_y = target_top_y（图层左上角对齐）
          4. 绕图层中心旋转 rotation 度（PS rotate 默认锚点=图层中心；
             PS rotate(正)=顺时针，PIL rotate(正)=逆时针 → 取负匹配）
          5. normal 混合（alpha 合成）到胚衣，与 JSX duplicate 图层默认一致
        """
        alpha_thr = ALPHA_THRESHOLD

        torso = Image.open(torso_path).convert("RGBA")
        design = Image.open(design_path).convert("RGBA")
        if black_optimize:
            try:
                import black_opt
                design = black_opt.black_shirt_print_optimize(design)
                print("  🖤 已对黑衫应用白墨打底 + 暗部提亮优化")
            except Exception as e:
                print(f"  ⚠️ 黑衫优化跳过（不影响贴图）: {e}")
        a = np.array(design)
        alpha = a[:, :, 3]
        mask = alpha >= alpha_thr
        ys, xs = np.where(mask)
        if len(xs) == 0:
            raise ValueError(f"无有效像素: {design_path}")
        x0, y0, x1, y1 = xs.min(), ys.min(), xs.max() + 1, ys.max() + 1

        trimmed = design.crop((x0, y0, x1, y1))
        tw, th = x1 - x0, y1 - y0

        # 定位参数：meta 优先（素材库五参），否则回退 config
        if meta is not None:
            target_w = float(meta["width"])
            target_h = float(meta.get("height", 0) or 0)
            if target_h > 0 and th:
                # 等比适应(contain)：设计图最长边不超过框(width×height)，另一边同比例缩放。
                # 例：框 546×546、设计又高又大(如 1297×1675) → 缩到 423×546，不再占满后背。
                scale = min(target_w / tw, target_h / th)
            else:
                # 旧行为：只约束宽度（顶层五参 height=0 的单面款保持原样）
                scale = (target_w / tw) if tw else 1.0
            rot = -meta["rotation"]
            target_cx = meta["center_x"]
            target_top_y = meta["highest_y"]
        else:
            scale = placement_cfg["scale_percent"] / 100.0
            rot = -placement_cfg["rotation"]
            target_cx = placement_cfg["target_center_x"]
            target_top_y = placement_cfg["target_top_y"]

        new_w = max(1, int(round(tw * scale)))
        new_h = max(1, int(round(th * scale)))
        scaled = trimmed.resize((new_w, new_h), Image.BICUBIC)

        # 参考点 = 缩放后设计图「顶边中点」(new_w/2, 0)，应落在胚衣坐标
        # (target_cx, target_top_y)。先求该点绕 scaled 中心旋转后在 R 中的位置，
        # 再平移使参考点对齐目标。与生成批准图的 place_meta 逻辑一致。
        import math
        theta = math.radians(rot)  # rot 已取负（PIL 正角=逆时针），单位度
        vx = (new_h / 2) * math.sin(theta)
        vy = -(new_h / 2) * math.cos(theta)

        # 绕 scaled 中心旋转，旋转中心 = scaled 中心
        rotated = scaled.rotate(rot, expand=True, resample=Image.BICUBIC)
        rw, rh = rotated.size
        anchor_x = rw / 2 + vx
        anchor_y = rh / 2 + vy
        px = int(round(target_cx - anchor_x))
        py = int(round(target_top_y - anchor_y))

        canvas = Image.new("RGBA", torso.size)
        canvas.paste(rotated, (px, py), rotated)
        out = Image.alpha_composite(torso, canvas)

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
            except Exception:
                pass
        save_with_size_limit(out.convert("RGB"), output_path, max_size_mb=2.0)
        print(f"✅ 生成: {output_path}")

        # 元数据注册（与原实现一致）
        if cut_meta is not None and wb_meta is not None:
            try:
                out_name = os.path.basename(output_path)
                role = _role_from_name(out_name)
                wb_meta.register_sticker(
                    output_path,
                    uid=cut_meta["uid"],
                    group_id=cut_meta["group_id"],
                    role=role,
                    parent_uid=cut_meta["uid"],
                    cut_file=os.path.basename(design_path),
                )
            except Exception as e:
                print(f"  ⚠️ 元数据注册失败: {e}")

    def close(self):
        # 纯软件实现，无需关闭 Photoshop
        pass


def classify_design(filename):
    """根据文件名分类设计图
    BW/WB → 正面+背面都要做
    B → 只做背面
    W → 只做正面
    """
    base = os.path.splitext(filename)[0]
    # 款号本身是否为双面款（DXxxxxBW / DXxxxxWB）：其 W/B 两个 cut 都属于双面款，
    # 正面(W) cut 需走「bw 第二套五参」，背面(B) cut 走顶层第一组。
    dx = base.split("_")[0]
    is_bw_dx = dx.endswith("BW") or dx.endswith("WB")

    if "_BW_" in base or "_BW" in base or "_WB_" in base or "_WB" in base:
        return "BW"
    if is_bw_dx:
        # 双面款拆成 W/B 两个 cut（如 DXxxxxBW_W_cut.png / DXxxxxBW_B_cut.png）：
        #   W 面 cut → 正面（bw 第二套）；B 面 cut → 背面（顶层）。
        if "_W" in base:
            return "BW"
        if "_B" in base:
            return "B"
        return "BW"  # 无 W/B 后缀的单 BW cut（两面都做）
    elif "_B_" in base or "_B" in base:
        return "B"
    elif "_W_" in base or "_W" in base:
        return "W"
    elif base.endswith("_1"):
        return "B"
    elif base.endswith("_2"):
        return "W"
    else:
        return "W"


def black_counterpart(file, dx):
    """返回通用_cut文件对应的黑版专用文件名（如 DX_B_cut.png → DX_黑B_cut.png）。
    支持版本号后缀：DX_B2_cut.png → DX_黑B2_cut.png。
    不存在对应关系时返回 None。"""
    if "_黑" in file or not file.lower().endswith("_cut.png"):
        return None
    stem = file[:-len("_cut.png")]  # e.g. DXxxxx_B / DXxxxx_BW / DXxxxx_B2
    suffix = stem[len(dx)+1:] if stem.startswith(dx + "_") else ""
    side, version = config.parse_side_suffix(suffix)
    if side in ("B", "W", "BW", "WB"):
        return f"{dx}_黑{side}{version}_cut.png"
    return None


def white_counterpart(file, dx):
    """返回通用_cut文件对应的白版专用文件名（如 DX_B_cut.png → DX_白B_cut.png）。
    支持版本号后缀：DX_B2_cut.png → DX_白B2_cut.png。
    不存在对应关系时返回 None。"""
    if "_白" in file or not file.lower().endswith("_cut.png"):
        return None
    stem = file[:-len("_cut.png")]
    suffix = stem[len(dx)+1:] if stem.startswith(dx + "_") else ""
    side, version = config.parse_side_suffix(suffix)
    if side in ("B", "W", "BW", "WB"):
        return f"{dx}_白{side}{version}_cut.png"
    return None


_SIDE_RE = re.compile(r'^(BW|WB|B|W)(\d*)$', re.IGNORECASE)


def real_sides(dx_folder):
    """解析 02_REM_BG 全部 _cut.png（含 _黑/_白 专用），返回该 DX 真实拥有的面集合。

    - 去掉 黑/白 前缀后解析 side；BW/WB 展开为 {B, W}。
    - 例如只有 DX_W_cut.png + DX_黑W_cut.png → {'W'}（单面 W）。
    """
    sides = set()
    rem_bg = os.path.join(dx_folder, "02_REM_BG")
    if not os.path.isdir(rem_bg):
        return sides
    dx = os.path.basename(dx_folder)
    prefix = dx + "_"
    for fn in os.listdir(rem_bg):
        if not fn.lower().endswith("_cut.png"):
            continue
        stem = os.path.splitext(fn)[0][:-4]  # 去掉 _cut
        if not stem.startswith(prefix):
            continue
        suffix = stem[len(prefix):]
        if suffix.startswith("黑") or suffix.startswith("白"):
            suffix = suffix[1:]
        m = _SIDE_RE.match(suffix)
        side = m.group(1).upper() if m else None
        if side in ("BW", "WB"):
            sides.update(("B", "W"))
        elif side in ("B", "W"):
            sides.add(side)
    return sides


def cleanup_stale_uploads(dx_folder):
    """单面款（只有 B 或只有 W）：清理 03_UPLOAD 中已不属于本款真实面的旧产物。

    旧款从双面改成单面后，旧互补面胚衣图（_W_白T/_W_黑T 或 _B_*）和旧 BW 平铺图
    （_白BW/_黑BW）会残留在 03_UPLOAD，导致 ps_batch 误把"新单面 + 旧互补面"当成
    完整 B+W 合成平铺图。本函数在每个贴图入口开头调用，按 02_REM_BG 真实面数清理。

    仅当真实面严格为 {'B'} 或 {'W'} 时清理；双面（含 BW/WB 或 B+W）/空/未知不动。
    返回实际删除的文件名列表。
    """
    sides = real_sides(dx_folder)
    if sides not in ({"B"}, {"W"}):
        return []
    dx = os.path.basename(dx_folder)
    upload = os.path.join(dx_folder, "03_UPLOAD")
    if not os.path.isdir(upload):
        return []
    missing_side = "W" if sides == {"B"} else "B"
    candidates = [
        wb_naming.bw_name(dx, "白"), wb_naming.bw_name(dx, "黑"),
        wb_naming.flat_name(dx, missing_side, "白"), wb_naming.flat_name(dx, missing_side, "黑"),
    ]
    removed = []
    for name in candidates:
        fp = os.path.join(upload, name)
        if os.path.exists(fp):
            try:
                os.remove(fp)
                removed.append(name)
            except Exception as e:
                print(f"  ⚠️ 清理残留失败 {name}: {e}")
    if removed:
        print(f"  🧹 单面款({''.join(sorted(sides))})清理旧互补面/平铺残留: {removed}")
    return removed


# 平铺图五参：单面款(只有W)用顶层五参，双面款(W+B)正面用 "bw" 第二套。
_FLAT_KEYS = ["width", "height", "rotation", "highest_y", "center_x"]


def _select_meta(full_meta, meta_set="w"):
    """从素材库 .meta.json 选出本次贴图用的五参。

    meta_set=="w"  -> 顶层五参（① 单面款 / 只有W贴图）
    meta_set=="bw" -> "bw" 子块（② 双面款 / 有W+B贴图）
    双面款要求第二组必须填：缺失/非数字/为0 直接报错，不偷偷用第一组兜底。
    """
    if not isinstance(full_meta, dict):
        full_meta = {}
    if meta_set == "bw":
        bw = full_meta.get("bw")
        if not isinstance(bw, dict):
            raise ValueError("缺少双面款(W+B)第二组参数「bw」，请到素材库填写")
        sel = {}
        for k in _FLAT_KEYS:
            v = bw.get(k)
            try:
                v = float(v)
            except (TypeError, ValueError):
                raise ValueError(f"双面款(W+B)第二组参数「{k}」无效，请到素材库填写")
            if v == 0:
                raise ValueError(f"双面款(W+B)第二组参数「{k}」为0，请到素材库填写")
            sel[k] = v
        return sel
    # 顶层 = 单面款（缺省回退 0，与旧行为一致）
    sel = {}
    for k in _FLAT_KEYS:
        v = full_meta.get(k)
        try:
            v = float(v)
        except (TypeError, ValueError):
            v = 0.0
        sel[k] = v
    return sel


def process_dx_folder(dx_folder, session=None):
    """处理单个 DX 文件夹，返回耗时（秒）。session 为 None 时内部创建。"""
    dx_name = os.path.basename(dx_folder)
    rem_bg_folder = os.path.join(dx_folder, "02_REM_BG")
    upload_folder = os.path.join(dx_folder, "03_UPLOAD")

    if not os.path.exists(rem_bg_folder):
        return 0.0

    os.makedirs(upload_folder, exist_ok=True)
    # 单面款贴图前先清理 03_UPLOAD 中已不存在的互补面/平铺旧残留
    cleanup_stale_uploads(dx_folder)
    own_session = session is None
    if own_session:
        session = StickerSession()

    # 五参定位：place_design 接收素材库 .meta.json，按胚衣真实尺寸贴图
    # （不再用 config.FRONT_NEW/BACK_NEW 写死的 scale/center）
    # meta_set: "w"=单面款用顶层第一组；"bw"=双面款正面用 bw 第二套（缺则报错）
    def _run(side, color, black_opt, meta_set="w"):
        torso_p, meta_p = config.flat_torso(side, color)
        full = config.load_meta(meta_p) or {}
        try:
            meta = _select_meta(full, meta_set)
        except ValueError as e:
            raise ValueError(f"{side}{color} 胚衣: {e}")
        session.place_design(
            design_path, torso_p,
            os.path.join(upload_folder, wb_naming.flat_name(dx_name, side, color)),
            meta=meta,
            black_optimize=black_opt, cut_meta=cut_meta,
        )

    t_dx = time.time()
    try:
        for file in os.listdir(rem_bg_folder):
            # 只处理 _cut.png 文件，跳过其他（如 DXxxxx_B.png）
            if not file.lower().endswith("_cut.png"):
                continue
            # 跳过黑版/白版专用文件，它们分别由 process_black.py / process_white.py 处理
            if "_黑" in file or "_白" in file:
                continue

            print(f"\n处理: {file}")
            design_path = os.path.join(rem_bg_folder, file)
            design_type = classify_design(file)
            base = os.path.splitext(file)[0]
            cut_meta = _get_cut_meta(design_path)

            # 如果存在对应的黑版专用文件，则通用图不再用于黑T
            black_file = black_counterpart(file, dx_name)
            has_black = black_file and os.path.exists(os.path.join(rem_bg_folder, black_file))
            if has_black:
                print(f"  发现黑版专用 {black_file}，通用图不输出黑T")

            # 如果存在对应的白版专用文件，则通用图不再用于白T
            white_file = white_counterpart(file, dx_name)
            has_white = white_file and os.path.exists(os.path.join(rem_bg_folder, white_file))
            if has_white:
                print(f"  发现白版专用 {white_file}，通用图不输出白T")

            if design_type == "BW":
                # ===== BW 类型：正面(W)用 bw 第二套五参，背面(B)用顶层第一组 =====
                # 双面款拆成 W/B 两个 cut（DXxxxxBW_W_cut.png / DXxxxxBW_B_cut.png）时，
                # W 面 cut 只生成正面、B 面 cut 只生成背面，避免背面被正面图覆盖。
                dx_is_bw = dx_name.endswith("BW") or dx_name.endswith("WB")
                is_split_w = dx_is_bw and "_W" in base
                if not has_white:
                    print("  → 生成 W 正面文件（双面款W+B第二组五参）...")
                    _run("W", "白", False, "bw")
                if not has_black:
                    _run("W", "黑", False, "bw")
                if not is_split_w:
                    if not has_white:
                        print("  → 生成 B 背面文件（五参定位）...")
                        _run("B", "白", False, "w")
                    if not has_black:
                        _run("B", "黑", False, "w")
                print("  ✅ BW 准备完成，可运行 ps_batch.py 合成最终 BW 图！")

            elif design_type == "W":
                # ===== W 类型：正图五参定位（单面款，用顶层第一组）=====
                if not has_white:
                    _run("W", "白", False, "w")
                if not has_black:
                    _run("W", "黑", False, "w")

            elif design_type == "B":
                # ===== B 类型：背图五参定位（背面只一组）=====
                if not has_white:
                    _run("B", "白", False, "w")
                if not has_black:
                    _run("B", "黑", False, "w")

        dt_dx = time.time() - t_dx
        print(f"⏱️  {dx_name} 完成，耗时 {dt_dx:.1f}秒")
        return dt_dx
    finally:
        if own_session:
            session.close()


def main(start_dx=None):
    """批量处理所有 DX 文件夹
    start_dx: 指定起始DX（如'DX0124'），从该文件夹开始处理（含）；None表示处理全部
    """
    dx_root = config.SOURCE_BASE
    folders = sorted([d for d in os.listdir(dx_root)
                       if d.startswith("DX") and os.path.isdir(os.path.join(dx_root, d))])
    if start_dx:
        idx = next((i for i, f in enumerate(folders) if f == start_dx), 0)
        folders = folders[idx:]
        print(f"从 {start_dx} 开始，共 {len(folders)} 个文件夹")

    t_total = time.time()
    dx_times = []
    session = StickerSession()
    try:
        for folder in folders:
            dx_folder = os.path.join(dx_root, folder)
            print(f"\n===== 处理: {dx_folder} =====")
            dt = process_dx_folder(dx_folder, session=session)
            dx_times.append((folder, dt))
    finally:
        session.close()

    dt_total = time.time() - t_total
    # 汇总
    print(f"\n{'='*60}")
    print(f"📊 全部完成！共 {len(dx_times)} 款")
    print(f"{'='*60}")
    print(f"{'DX':<12} {'耗时(秒)':>10}")
    print(f"{'-'*24}")
    for name, dt in dx_times:
        print(f"{name:<12} {dt:>10.1f}")
    print(f"{'-'*24}")
    print(f"{'总计':<12} {dt_total:>10.1f}  ({dt_total/60:.1f}分钟)")
    if dx_times:
        print(f"{'平均':<12} {dt_total/len(dx_times):>10.1f}")
    print(f"{'='*60}")


if __name__ == "__main__":
    import sys
    start = sys.argv[1] if len(sys.argv) > 1 else None
    main(start)
