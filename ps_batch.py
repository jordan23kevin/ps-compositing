# ===== PS 正反图批处理 v1.3.0 =====
# 变更 v1.3.0：Photoshop 窗口全程最小化/隐藏，不抢焦点
# 直读03_UPLOAD贴图结果，直写03_UPLOAD BW合成图
import io, win32com.client, os, time, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"D:\Semems WB\02_PROJECTS"
PS_EXE = r"D:\Program Files\Adobe Photoshop 2025 v26.0\Adobe Photoshop 2025\Photoshop.exe"

def log(msg):
    print(f"[PS] {msg}", flush=True)

def ps_open_both(back_img, front_img):
    shell = win32com.client.Dispatch('WScript.Shell')
    # 7 = SW_SHOWMINNOACTIVE：最小化打开，不抢焦点
    shell.Run(f'"{PS_EXE}" "{front_img}" "{back_img}"', 7, False)

def wait_ps_docs(target_count=2, timeout=20):
    for _ in range(timeout):
        try:
            ps = win32com.client.GetObject(Class='Photoshop.Application')
            if ps.Documents.Count >= target_count:
                return ps
        except:
            pass
        time.sleep(0.5)
    return win32com.client.GetObject(Class='Photoshop.Application')

def get_ps():
    for _ in range(15):
        try:
            return win32com.client.GetObject(Class='Photoshop.Application')
        except:
            time.sleep(0.5)
    raise Exception("PS 连接超时")

def close_all_docs(ps):
    for i in range(ps.Documents.Count, 0, -1):
        try:
            ps.Documents(i).Close(2)
        except:
            pass

def process_color(dx_folder, color, action_name, output_name):
    """处理一个颜色，从03_UPLOAD读取，输出到03_UPLOAD"""
    upload = os.path.join(BASE, dx_folder, "03_UPLOAD")
    back_img = os.path.join(upload, f"{dx_folder}_B_{color}.jpg")
    front_img = os.path.join(upload, f"{dx_folder}_W_{color}.jpg")
    out_path = os.path.join(upload, output_name)

    if not os.path.exists(back_img):
        log(f"  跳过{color}：找不到{dx_folder}_B_{color}.jpg")
        return False

    if not os.path.exists(front_img):
        log(f"  跳过{color}：找不到{dx_folder}_W_{color}.jpg")
        return False

    if os.path.exists(out_path):
        log(f"  跳过{color}：{output_name}已存在")
        return False

    log(f"  打开正背两张...")
    try:
        ps_clean = get_ps()
        close_all_docs(ps_clean)
    except:
        pass
    ps_open_both(back_img, front_img)
    ps = wait_ps_docs(2)
    try:
        ps.Visible = False
    except Exception:
        pass
    log(f"  已打开 {ps.Documents.Count} 个文档")

    for i in range(1, ps.Documents.Count + 1):
        doc = ps.Documents(i)
        if f"_B_{color}" in doc.Name:
            ps.ActiveDocument = doc
            log(f"  激活: {doc.Name}")
            break

    ps.DoAction(action_name, '正反图')
    log(f"  动作: 正反图 > {action_name}")

    if os.path.exists(out_path):
        os.remove(out_path)

    export_opts = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
    export_opts.Format = 6
    export_opts.Quality = 100
    export_opts.Optimized = True

    ps.ActiveDocument.Export(ExportIn=out_path, ExportAs=2, Options=export_opts)
    size = os.path.getsize(out_path)
    log(f"  OK: {output_name} ({size/1024:.0f}KB)")

    close_all_docs(ps)
    return True

def main(start_dx=None):
    log(f"扫描: {BASE}")
    folders = sorted([d for d in os.listdir(BASE)
                      if os.path.isdir(os.path.join(BASE, d)) and d.startswith('DX')])
    if start_dx:
        idx = next((i for i, f in enumerate(folders) if f == start_dx), 0)
        folders = folders[idx:]
    log(f"从 {start_dx or '开头'} 开始，发现 {len(folders)} 个 DX 文件夹")

    results = {"白T": 0, "黑T": 0, "跳过白": 0, "跳过黑": 0, "失败黑": []}
    import time
    dx_times = []
    t_total = time.time()

    for folder_name in folders:
        log(f"\n{'='*50}")
        log(">> " + folder_name)
        log(f"{'='*50}")
        t_dx = time.time()

        if process_color(folder_name, "白T", "白", f"{folder_name}_白BW.jpg"):
            results["白T"] += 1
        else:
            results["跳过白"] += 1

        try:
            ps = get_ps()
            if process_color(folder_name, "黑T", "黑", f"{folder_name}_黑BW.jpg"):
                results["黑T"] += 1
            else:
                results["跳过黑"] += 1
        except Exception as e:
            log(f"  黑T异常: {e}")
            results["失败黑"].append(folder_name)
            try:
                ps = get_ps()
                close_all_docs(ps)
            except:
                pass
            dt_dx = time.time() - t_dx
            dx_times.append((folder_name, dt_dx))
            log(f"⏱️  {folder_name} 完成，耗时 {dt_dx:.1f}秒")
            continue

        dt_dx = time.time() - t_dx
        dx_times.append((folder_name, dt_dx))
        log(f"⏱️  {folder_name} 完成，耗时 {dt_dx:.1f}秒")

    dt_total = time.time() - t_total
    log(f"\n{'='*60}")
    log(f"📊 全部完成！共 {len(dx_times)} 款")
    log(f"  白T生成: {results['白T']}  跳过: {results['跳过白']}")
    log(f"  黑T生成: {results['黑T']}  跳过: {results['跳过黑']}")
    if results["失败黑"]:
        log(f"  黑T异常需重试: {', '.join(results['失败黑'])}")
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
