# ===== PS 正反图批处理 v1.1.0 =====
# 更新日期: 2026-06-21
# 核心功能: 遍历 DX 文件夹，正反图合成白BW/黑BW
# 关键优化: 一次打开两张图 + Export保存不弹窗 + 开图前清空标签
# =========================================
import io, win32com.client, os, time, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

SRC = r"D:\Semems\1AI"
PS_EXE = r"D:\Program Files\Adobe Photoshop 2025 v26.0\Adobe Photoshop 2025\Photoshop.exe"

def log(msg):
    print(f"[PS] {msg}", flush=True)

def ps_open_both(back_img, front_img):
    """一次性打开正背两张图（PS命令行支持多文件）"""
    shell = win32com.client.Dispatch('WScript.Shell')
    # PS 最后打开的文件会成为活动文档，所以正先打开、背后打开 → 背自动激活
    shell.Run(f'"{PS_EXE}" "{front_img}" "{back_img}"', 1, False)

def wait_ps_docs(target_count=2, timeout=20):
    """轮询等待 PS 打开指定数量的文档"""
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
    """获取 PS 实例"""
    for _ in range(15):
        try:
            return win32com.client.GetObject(Class='Photoshop.Application')
        except:
            time.sleep(0.5)
    raise Exception("PS 连接超时")

def close_all_docs(ps):
    """关闭所有文档，不保存"""
    for i in range(ps.Documents.Count, 0, -1):
        try:
            ps.Documents(i).Close(2)
        except:
            pass

def process_color(ps, folder, color, action_name, output_name):
    """处理一个颜色：白→白正/白背，黑→黑正/黑背"""
    back_img = os.path.join(folder, f"{color}背2_副本.jpg")
    front_img = os.path.join(folder, f"{color}正2_副本.jpg")
    out_path = os.path.join(folder, output_name)

    if not os.path.exists(back_img):
        log(f"  跳过{color}T：找不到{color}背2_副本.jpg")
        return False

    if os.path.exists(out_path):
        log(f"  跳过{color}T：{output_name}已存在")
        return False

    log(f"  清理旧标签 + 打开正背两张...")
    try:
        ps_clean = get_ps()
        close_all_docs(ps_clean)
    except:
        pass  # PS 可能还没启动
    ps_open_both(back_img, front_img)
    ps = wait_ps_docs(2)
    log(f"  已打开 {ps.Documents.Count} 个文档")

    # 激活背面的文档
    for i in range(1, ps.Documents.Count + 1):
        doc = ps.Documents(i)
        if f"{color}背" in doc.Name:
            ps.ActiveDocument = doc
            log(f"  激活: {doc.Name}")
            break

    # 运行动作
    ps.DoAction(action_name, '正反图')
    log(f"  动作: 正反图 > {action_name}")

    # 保存
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
    import sys as _sys
    start_from = None
    for a in _sys.argv[1:]:
        if a.startswith('DX'):
            start_from = a
    if start_from:
        log(f"起始文件夹: {start_from}")

    log(f"扫描: {SRC}")
    folders = sorted([d for d in os.listdir(SRC)
                      if os.path.isdir(os.path.join(SRC, d)) and d.startswith('DX')])
    if start_from and start_from in folders:
        folders = folders[folders.index(start_from):]
    log(f"发现 {len(folders)} 个 DX 文件夹" + (f"（从{start_from}开始）" if start_from else ""))

    results = {"白T": 0, "黑T": 0, "跳过白": 0, "跳过黑": 0, "失败黑": []}

    for folder_name in folders:
        folder_path = os.path.join(SRC, folder_name)
        log(f"\n{'='*50}")
        log(">> " + folder_name)
        log(f"{'='*50}")

        # 白T
        if process_color(None, folder_path, "白", "白", "白BW.jpg"):
            results["白T"] += 1
        else:
            results["跳过白"] += 1

        # 黑T（加异常保护，失败则记录跳过下一张）
        try:
            ps = get_ps()
            if process_color(ps, folder_path, "黑", "黑", "黑BW.jpg"):
                results["黑T"] += 1
            else:
                results["跳过黑"] += 1
        except Exception as e:
            log(f"  ⚠️ 黑T异常: {e}")
            results["失败黑"].append(folder_name)
            # 尝试关闭所有文档恢复PS状态
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
