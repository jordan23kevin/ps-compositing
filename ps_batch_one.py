"""PS BW合成 单款入口 v2.1 — 被 check_rem.py 调用处理单个 DX 的 BW 合成

变更 v2.1：
  - 作为 check_rem.py 贴图流水线最后一步，由 /ps-sticker 自动调用
  - Photoshop 窗口最小化/隐藏由 ps_batch.py 统一处理
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from ps_batch import process_color, get_ps, close_all_docs
from config import SOURCE_BASE

if __name__ == "__main__":
    dx = sys.argv[1]
    dx_folder = dx  # process_color 内部用 BASE/dx_folder/03_UPLOAD
    if not os.path.isdir(os.path.join(SOURCE_BASE, dx)):
        print(f"❌ {dx} 不存在")
        sys.exit(1)
    t0 = time.time()
    print(f"\n=== BW合成: {dx} ===")

    # 先关掉PS里的文档
    try:
        ps = get_ps()
        close_all_docs(ps)
    except:
        pass

    # 白T BW合成
    print(f"  白T BW合成...")
    ok_white = process_color(dx_folder, "白T", "白", f"{dx}_白BW.jpg")
    # 黑T BW合成
    print(f"  黑T BW合成...")
    try:
        ok_black = process_color(dx_folder, "黑T", "黑", f"{dx}_黑BW.jpg")
    except Exception as e:
        print(f"  黑T异常: {e}")
        ok_black = False

    dt = time.time() - t0
    print(f"\n✓ {dx} BW合成完成 白T={'✅' if ok_white else '⏭'} 黑T={'✅' if ok_black else '⏭'} 耗时 {dt:.1f}秒")
