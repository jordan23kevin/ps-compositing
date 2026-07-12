# ===== PS 正反图批处理 v1.5 =====
# 变更 v1.5：
#   - 将 BW 合成从 Photoshop 动作集「正反图」改为 PIL 直拼，
#     消除对私有 PS 动作集（白T/黑T）的依赖。新 PS 安装/迁移后动作集
#     丢失也不会导致 BW 合成失败。
# 变更 v1.4.0：
#   - 整个批次复用同一个 Photoshop COM 会话
#   - 每个 DX 一次打开 B/W 正背图，连续执行白/黑动作后统一关闭
#   - 用主动轮询替代硬编码 sleep，减少等待
# 直读03_UPLOAD贴图结果，直写03_UPLOAD BW合成图
import io, win32com.client, pythoncom, os, time, sys, tempfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter
try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

import config

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

# ---------------------------------------------------------------------------
# 元数据辅助
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

def log(msg):
    print(f"[PS] {msg}", flush=True)


def get_ps(timeout=10):
    """获取 Photoshop COM 对象，带超时。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            return win32com.client.GetObject(Class='Photoshop.Application')
        except:
            time.sleep(0.1)
    raise Exception("PS 连接超时")


def open_doc(ps, img_path):
    """通过 JSX app.open 打开一个图片文件，并返回当前活动文档。"""
    doc_name = os.path.basename(img_path)
    jsx = (
        'var f = new File("' + img_path.replace("\\", "\\\\") + '");\n'
        'app.open(f);\n'
    )
    temp_jsx = os.path.join(tempfile.gettempdir(), f"_open_{doc_name}.jsx")
    with open(temp_jsx, "w", encoding="utf-8") as f:
        f.write(jsx)
    ps.DoJavaScriptFile(temp_jsx)
    config.hide_ps_window(ps)
    return ps.ActiveDocument


def wait_docs(ps, target_count, timeout=10):
    """主动轮询，等待文档数量达标。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if ps.Documents.Count >= target_count:
                return
        except:
            pass
        time.sleep(0.05)


def close_docs(docs):
    """关闭指定文档列表。"""
    for doc in docs:
        try:
            doc.Close(2)
        except:
            pass


def export_bw(ps, out_path):
    """用 ExportOptionsSaveForWeb 导出当前活动文档为 JPG。"""
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
        except Exception as e:
            log(f"  删除旧{os.path.basename(out_path)}失败: {e}")

    export_opts = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    export_opts.Format = 6
    export_opts.Quality = 100
    export_opts.Optimized = True

    ps.ActiveDocument.Export(ExportIn=out_path, ExportAs=2, Options=export_opts)
    config.hide_ps_window(ps)
    size = os.path.getsize(out_path)
    log(f"  OK: {os.path.basename(out_path)} ({size/1024:.0f}KB)")


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


def compose_bw_pil(front_path, back_path, out_path,
                   inset_ratio=0.36, center_x_ratio=0.80, center_y_ratio=0.84,
                   shadow_offset=(8, 8), shadow_blur=12, shadow_opacity=0.35):
    """用 PIL 将正面图以圆形插图形式合成到背面图右下角，不依赖 PS 动作集。

    复刻原「正反图」动作集效果：
      - 背面图做为主画布
      - 正面图缩放后裁剪成圆形，贴在右下角，带阴影
    参数为相对比例，自适应不同尺寸图片。
    """
    back = Image.open(back_path).convert("RGB")
    front = Image.open(front_path).convert("RGB")
    w, h = back.size

    diameter = int(w * inset_ratio)
    if diameter <= 2:
        raise ValueError("图片宽度太小，无法生成圆形插图")
    radius = diameter // 2
    cx = int(w * center_x_ratio)
    cy = int(h * center_y_ratio)

    # 正面图缩放到正方形
    front_sq = front.resize((diameter, diameter), Image.LANCZOS)

    # 圆形 mask（抗锯齿羽化 1px）
    mask = Image.new("L", (diameter, diameter), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0, diameter - 1, diameter - 1), fill=255)
    mask = mask.filter(ImageFilter.GaussianBlur(radius=1))

    # 阴影层：圆形高斯模糊 + 半透明黑色
    base = back.convert("RGBA")
    sd = Image.new("L", (diameter, diameter), 0)
    sdraw = ImageDraw.Draw(sd)
    sdraw.ellipse((0, 0, diameter - 1, diameter - 1), fill=255)
    sd = sd.filter(ImageFilter.GaussianBlur(radius=shadow_blur))
    shadow = Image.new("RGBA", (diameter, diameter), (0, 0, 0, int(255 * shadow_opacity)))
    shadow.putalpha(sd)
    sx = cx - radius + shadow_offset[0]
    sy = cy - radius + shadow_offset[1]
    base.paste(shadow, (sx, sy), shadow)

    # 圆形正面图贴到主画布
    circle = front_sq.convert("RGBA")
    circle.putalpha(mask)
    base.paste(circle, (cx - radius, cy - radius), mask)

    base.convert("RGB").save(out_path, quality=100, optimize=True)
    return out_path


def process_dx(ps, dx_folder):
    """处理单个 DX 的 BW 合成（白T + 黑T）。"""
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
            compose_bw_pil(front_img, back_img, out_path)
            size = os.path.getsize(out_path)
            log(f"  OK: {output_name} ({size/1024:.0f}KB)")
            register_bw_meta(back_img, front_img, out_path)
            results.append((color, True))
        except Exception as e:
            log(f"  BW合成失败: {e}")
            results.append((color, False))

    return results


def main(start_dx=None):
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

    pythoncom.CoInitialize()
    ps = None
    try:
        ps = get_ps(timeout=15)
        ps.DisplayDialogs = 3
        config.hide_ps_window(ps)

        for folder_name in folders:
            log(f"\n{'='*50}")
            log(">> " + folder_name)
            log(f"{'='*50}")
            t_dx = time.time()

            try:
                color_results = process_dx(ps, folder_name)
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
                try:
                    # 尽量清理残留文档
                    for i in range(ps.Documents.Count, 0, -1):
                        try:
                            ps.Documents(i).Close(2)
                        except:
                            pass
                except:
                    pass

            dt_dx = time.time() - t_dx
            dx_times.append((folder_name, dt_dx))
            log(f"⏱️  {folder_name} 完成，耗时 {dt_dx:.1f}秒")
    finally:
        if ps is not None:
            try:
                for i in range(ps.Documents.Count, 0, -1):
                    try:
                        ps.Documents(i).Close(2)
                    except:
                        pass
            except:
                pass
        pythoncom.CoUninitialize()

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
