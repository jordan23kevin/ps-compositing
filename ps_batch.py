# ===== WB 正反图批处理 v2.0.0（纯软件，不再依赖 Photoshop） =====
# 变更 v2.0.0（2026-07-12）：
#   - 彻底移除 Photoshop 依赖：删除 get_ps/open_doc/wait_docs/close_docs/
#     export_bw 等 COM 函数与 win32com/pythoncom 导入；main() 不再 CoInitialize。
#   - BW 合成全程纯 PIL，整条「贴图 + BW」流水线从此零 PS 依赖、零 PS 动作集依赖。
# 变更 v1.8：按 DX0481 参考图 + 用户规格精确调整 BW 合成（见 compose_bw_pil）。
# 变更 v1.7：圆形插图保留正面图完整背景（木地板/报纸/鞋子/衣架），不再填衣服色。
# 变更 v1.6：BW 合成由 PS 动作集「正反图」改为 PIL 直拼。
# 变更 v1.4.0：单 DX 复用同一 PS COM 会话（v2.0.0 起已废弃）。
# 直读 03_UPLOAD 贴图结果，直写 03_UPLOAD BW 合成图。
import io, os, time, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

BASE = r"D:\Semems WB\02_PROJECTS"

sys.path.insert(0, r"E:\Claude code\ps")
sys.path.insert(0, r"D:\Semems WB\04_OS\engine")
import wb_naming  # 命名规则唯一出处

try:
    import wb_meta
except Exception:
    wb_meta = None
try:
    from wb_sticker_ps import real_sides
except Exception:
    real_sides = None

VERSION = "2.0.1"

_MIGRATED_DX = set()


def log(msg):
    print(f"[BW] {msg}", flush=True)


def _role_from_name(name):
    """从文件名推断 role（规则见 wb_naming.role_from_name）"""
    return wb_naming.role_from_name(name)


def _infer_meta(path):
    """sidecar 完全缺失时的文件名兜底推断"""
    name = os.path.basename(path)
    dx = name.split("_")[0] if "_" in name else "DX"
    role = _role_from_name(name)
    uid = f"UID_{dx}_{role}"
    group_id = f"G_{dx}_{role}"
    return {"uid": uid, "group_id": group_id, "role": role, "stage": "unknown"}


def _get_meta(path):
    """读取上传图 sidecar；缺失时 migrate_dx，最后兜底推断"""
    if wb_meta is None:
        return None
    dx_dir = str(Path(path).parent.parent)
    meta = wb_meta.read_meta(path)
    if meta:
        return meta
    if dx_dir not in _MIGRATED_DX:
        try:
            wb_meta.migrate_dx(dx_dir)
        except Exception as e:
            log(f"migrate_dx 失败 {dx_dir}: {e}")
        _MIGRATED_DX.add(dx_dir)
    meta = wb_meta.read_meta(path)
    if meta:
        return meta
    return _infer_meta(path)


def register_bw_meta(back_img, front_img, out_path):
    """注册 BW 合成图元数据。"""
    if wb_meta is None:
        return
    try:
        meta_b = _get_meta(back_img)
        meta_w = _get_meta(front_img)
        uid = meta_b.get("uid") or meta_w.get("uid")
        group_id = meta_b.get("group_id") or meta_w.get("group_id")
        bw_role = _role_from_name(os.path.basename(out_path))
        bw_uid = f"{uid}_{bw_role}" if uid else None
        wb_meta.register_bw(
            out_path,
            uid=bw_uid,
            group_id=group_id,
            role=bw_role,
            source_uids=[meta_b.get("uid"), meta_w.get("uid")],
            source_files=[os.path.basename(back_img), os.path.basename(front_img)],
        )
        log(f"  BW元数据已注册: {bw_role}")
    except Exception as e:
        log(f"  BW元数据注册失败: {e}")


def compose_bw_pil(front_path, back_path, out_path, shirt_color="white",
                   diameter=595, center_x_ratio=0.7567, center_y_ratio=0.8118,
                   front_scale=None, border_width=5,
                   shadow_offset=(0, 0), shadow_blur=0, shadow_opacity=0.0):
    """用 PIL 像素级复刻「正反图」BW 合成（依据 DX0481 参考图实测参数）。

    规格（用户给定 + 参考图实测）：
      - 底图 1340×1785，圆形插图直径 595（半径 297）。
      - 正面图缩放到「宽度 = 圆圈直径」（实测比例≈0.444，即 595/1340）。
        此时高度（≈793）大于圆圈，垂直居中后上下各裁掉约 99px，
        只保留圆圈内的部分 —— 与参考图一致。
      - 圆心位于 (1014, 1449)，即宽 75.67% / 高 81.18%（参考图实测）。
      - 白边宽度 5px（在圆圈外侧加一圈白）。
      - 不要阴影。

    front_scale：正面图缩放系数（相对原图像素）。默认 None = 宽度贴合圆圈。
      若用户要求「缩放到 50%」，传 front_scale=0.5 即可。
    """
    if shirt_color in ("白", "white"):
        shirt_color = "white"
    elif shirt_color in ("黑", "black"):
        shirt_color = "black"

    back = Image.open(back_path).convert("RGB")
    front = Image.open(front_path).convert("RGB")

    w, h = back.size
    cx = int(round(w * center_x_ratio))
    cy = int(round(h * center_y_ratio))
    radius = diameter // 2

    # 正面图缩放
    fw, fh = front.size
    if front_scale is None:
        front_scale = diameter / fw          # 宽度贴合圆圈
    scaled = front.resize((int(fw * front_scale), int(fh * front_scale)), Image.LANCZOS)
    sw, sh = scaled.size

    # 直径×直径透明画布，正面图居中（宽/高超出则自然裁切）
    canvas = Image.new("RGBA", (diameter, diameter), (0, 0, 0, 0))
    ox = (diameter - sw) // 2
    oy = (diameter - sh) // 2
    canvas.paste(scaled, (ox, oy))

    # 圆形 mask（0.5px 羽化去锯齿）
    mask = Image.new("L", (diameter, diameter), 0)
    ImageDraw.Draw(mask).ellipse((0, 0, diameter - 1, diameter - 1), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(radius=0.5))
    canvas.putalpha(mask)

    # 完整对象：白边(外圈5px) + 圆形正面图
    full_size = diameter + 2 * border_width
    full = Image.new("RGBA", (full_size, full_size), (0, 0, 0, 0))
    white_bg = Image.new("RGBA", (full_size, full_size), (255, 255, 255, 255))
    border_mask = Image.new("L", (full_size, full_size), 0)
    ImageDraw.Draw(border_mask).ellipse((0, 0, full_size - 1, full_size - 1), fill=255)
    border_mask = border_mask.filter(ImageFilter.GaussianBlur(radius=0.5))
    white_bg.putalpha(border_mask)
    full.paste(white_bg, (0, 0), white_bg)
    full.paste(canvas, (border_width, border_width), mask)

    base = back.convert("RGBA")
    # 阴影：默认关闭（shadow_opacity=0）
    if shadow_opacity > 0 and shadow_blur > 0:
        shadow_size = full_size + 2 * shadow_blur
        sd = Image.new("L", (shadow_size, shadow_size), 0)
        ImageDraw.Draw(sd).ellipse(
            (shadow_blur, shadow_blur, shadow_blur + full_size - 1, shadow_blur + full_size - 1), fill=255)
        sd = sd.filter(ImageFilter.GaussianBlur(radius=shadow_blur))
        shadow = Image.new("RGBA", (shadow_size, shadow_size), (0, 0, 0, int(255 * shadow_opacity)))
        shadow.putalpha(sd)
        sx = cx - radius - border_width + shadow_offset[0] - shadow_blur
        sy = cy - radius - border_width + shadow_offset[1] - shadow_blur
        base.paste(shadow, (sx, sy), shadow)

    base.paste(full, (cx - radius - border_width, cy - radius - border_width), full)
    base.convert("RGB").save(out_path, quality=95, optimize=True)
    return out_path


def process_dx(ps, dx_folder):
    """处理单个 DX 的 BW 合成（白T + 黑T）。ps 参数为兼容旧调用保留，纯软件下未使用。"""
    upload = os.path.join(BASE, dx_folder, "03_UPLOAD")

    # 单面款（02_REM_BG 只有 B 或只有 W）不该合成 BW 平铺图：
    # 拦截并清掉旧 _白BW/_黑BW 残留，避免“新单面图 + 旧互补面图”被误拼成平铺。
    single_side = None
    if real_sides is not None:
        try:
            sides = real_sides(os.path.join(BASE, dx_folder))
            if sides in ({"B"}, {"W"}):
                single_side = next(iter(sides))
        except Exception as e:
            log(f"  ⚠️ real_sides 解析失败: {e}")

    colors = [
        ("白T", "白", wb_naming.bw_name(dx_folder, "白")),
        ("黑T", "黑", wb_naming.bw_name(dx_folder, "黑")),
    ]

    results = []
    for color, action_name, output_name in colors:
        back_img = os.path.join(upload, wb_naming.flat_name(dx_folder, "B", action_name))
        front_img = os.path.join(upload, wb_naming.flat_name(dx_folder, "W", action_name))
        out_path = os.path.join(upload, output_name)

        if single_side is not None:
            if os.path.exists(out_path):
                try:
                    os.remove(out_path)
                    log(f"  🧹 单面款({single_side})删除残留平铺: {output_name}")
                except Exception as e:
                    log(f"  ⚠️ 删除{output_name}失败: {e}")
            else:
                log(f"  跳过{color}：单面款({single_side})不合成BW")
            results.append((color, False))
            continue

        if not os.path.exists(back_img) or not os.path.exists(front_img):
            log(f"  跳过{color}：缺少 {os.path.basename(back_img)} / {os.path.basename(front_img)}")
            results.append((color, False))
            continue

        if os.path.exists(out_path):
            log(f"  覆盖{color}：{output_name}已存在，将重新生成")

        try:
            compose_bw_pil(front_img, back_img, out_path, shirt_color=action_name)
            size = os.path.getsize(out_path)
            log(f"  OK: {output_name} ({size/1024:.0f}KB)")
            register_bw_meta(back_img, front_img, out_path)
            results.append((color, True))
        except Exception as e:
            log(f"  BW合成失败: {e}")
            results.append((color, False))

    return results


def main(start_dx=None):
    """纯软件批量合成所有 DX 的 BW 图（不连接 Photoshop）。"""
    log(f"扫描: {BASE}")
    folders = sorted([d for d in os.listdir(BASE)
                      if os.path.isdir(os.path.join(BASE, d)) and d.startswith('DX')])
    if start_dx:
        idx = next((i for i, f in enumerate(folders) if f == start_dx), 0)
        folders = folders[idx:]
    log(f"从 {start_dx or '开头'} 开始，发现 {len(folders)} 个 DX 文件夹")

    results = {"白T": 0, "黑T": 0, "跳过白": 0, "跳过黑": 0, "失败": []}
    dx_times = []
    t_total = time.time()

    for folder_name in folders:
        log(f"\n{'='*50}")
        log(">> " + folder_name)
        log(f"{'='*50}")
        t_dx = time.time()

        try:
            color_results = process_dx(None, folder_name)
            for color, ok in color_results:
                key = "白T" if color == "白T" else "黑T"
                skip_key = "跳过白" if color == "白T" else "跳过黑"
                if ok:
                    results[key] += 1
                else:
                    results[skip_key] += 1
        except Exception as e:
            log(f"  异常: {e}")
            results["失败"].append(folder_name)

        dt_dx = time.time() - t_dx
        dx_times.append((folder_name, dt_dx))
        log(f"⏱️  {folder_name} 完成，耗时 {dt_dx:.1f}秒")

    dt_total = time.time() - t_total
    log(f"\n{'='*60}")
    log(f"📊 全部完成！共 {len(dx_times)} 款")
    log(f"  白T生成: {results['白T']}  跳过: {results['跳过白']}")
    log(f"  黑T生成: {results['黑T']}  跳过: {results['跳过黑']}")
    if results["失败"]:
        log(f"  异常需重试: {', '.join(results['失败'])}")
    log(f"{'='*60}")
    log(f"{'DX':<12} {'耗时(秒)':>10}")
    log(f"{'-'*24}")
    for name, dt in dx_times:
        log(f"{name:<12} {dt:>10.1f}")
    log(f"{'-'*24}")
    log(f"{'总计':<12} {dt_total:>10.1f}  ({dt_total/60:.1f}分钟)")
    if dx_times:
        log(f"{'平均':<12} {dt_total/len(dx_times):>10.1f}")
    log(f"{'='*60}")


if __name__ == "__main__":
    import sys
    start = sys.argv[1] if len(sys.argv) > 1 else None
    main(start)
