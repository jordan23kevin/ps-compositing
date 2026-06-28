# ===== PS 正反图批处理 v1.2.0 =====
# 直读03_UPLOAD贴图结果，直写03_UPLOAD BW合成图
import io, win32com.client, os, time, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = r"D:\Semems WB\02_PROJECTS"
PS_EXE = r"D:\Program Files\Adobe Photoshop 2025 v26.0\Adobe Photoshop 2025\Photoshop.exe"

def log(msg):
    print(f"[PS] {msg}", flush=True)

def ps_open_both(back_img, front_img):
    shell = win32com.client.Dispatch('WScript.Shell')
    shell.Run(f'"{PS_EXE}" "{front_img}" "{back_img}"', 1, False)

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

def main():
    log(f"扫描: {BASE}")
    folders = sorted([d for d in os.listdir(BASE)
                      if os.path.isdir(os.path.join(BASE, d)) and d.startswith('DX')])
    log(f"发现 {len(folders)} 个 DX 文件夹")

    results = {"白T": 0, "黑T": 0, "跳过白": 0, "跳过黑": 0, "失败黑": []}

    for folder_name in folders:
        log(f"\n{'='*50}")
        log(">> " + folder_name)
        log(f"{'='*50}")

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
            continue

    log(f"\n{'='*50}")
    log(f"全部完成！")
    log(f"  白T生成: {results['白T']}  跳过: {results['跳过白']}")
    log(f"  黑T生成: {results['黑T']}  跳过: {results['跳过黑']}")
    if results["失败黑"]:
        log(f"  黑T异常需重试: {', '.join(results['失败黑'])}")
    log(f"{'='*50}")

if __name__ == "__main__":
    main()
