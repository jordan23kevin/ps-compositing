# ===== PS 正反图批处理 v1.4.1 =====
# 变更 v1.4.0：
#   - 整个批次复用同一个 Photoshop COM 会话
#   - 每个 DX 一次打开 B/W 正背图，连续执行白/黑动作后统一关闭
#   - 用主动轮询替代硬编码 sleep，减少等待
# 直读03_UPLOAD贴图结果，直写03_UPLOAD BW合成图
import io, win32com.client, pythoncom, os, time, sys, tempfile
from pathlib import Path
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import config

BASE = r"D:\Semems WB\02_PROJECTS"

sys.path.insert(0, r"E:\Claude code\ps")
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
    """从文件名推断 role（支持上传图 / BW 合成图）"""
    stem = os.path.splitext(name)[0]
    if stem.endswith("_cut"):
        stem = stem[:-4]
    parts = stem.split("_")
    if len(parts) >= 3 and parts[-1] in ("白T", "黑T"):
        side = parts[-2]
        torso = parts[-1]
        if torso == "黑T":
            return f"黑{side}"
        return side
    if len(parts) >= 2:
        return parts[-1]
    return "?"


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
        ("白T", "白", f"{dx_folder}_白BW.jpg"),
        ("黑T", "黑", f"{dx_folder}_黑BW.jpg"),
    ]

    results = []
    for color, action_name, output_name in colors:
        back_img = os.path.join(upload, f"{dx_folder}_B_{color}.jpg")
        front_img = os.path.join(upload, f"{dx_folder}_W_{color}.jpg")
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
            log(f"  跳过{color}：缺少 {dx_folder}_B/W_{color}.jpg")
            results.append((color, False))
            continue

        if os.path.exists(out_path):
            log(f"  覆盖{color}：{output_name}已存在，将重新生成")

        log(f"  打开{color}正背两张...")
        back_doc = open_doc(ps, back_img)
        front_doc = open_doc(ps, front_img)
        wait_docs(ps, 2)
        ps.DisplayDialogs = 3  # DialogModes.NO
        config.hide_ps_window(ps)
        log(f"  已打开 {ps.Documents.Count} 个文档")

        # 激活背面文档并执行动作
        ps.ActiveDocument = back_doc
        log(f"  激活: {back_doc.Name}")
        ps.DoAction(action_name, '正反图')
        log(f"  动作: 正反图 > {action_name}")
        config.hide_ps_window(ps)

        export_bw(ps, out_path)
        register_bw_meta(back_img, front_img, out_path)

        close_docs([back_doc, front_doc])
        results.append((color, True))

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
