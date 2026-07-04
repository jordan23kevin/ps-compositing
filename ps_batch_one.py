"""PS BW合成 单款入口 v2.2 — 被 check_rem.py 调用处理单个 DX 的 BW 合成

变更 v2.2：
  - 修复从 ps_batch 导入不存在的 process_color / close_all_docs 导致的 ImportError
  - 改为直接调用 ps_batch.process_dx，由 process_dx 内部完成白T/黑T BW 合成
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from ps_batch import process_dx, get_ps
from config import SOURCE_BASE

if __name__ == "__main__":
    dx = sys.argv[1]
    if not os.path.isdir(os.path.join(SOURCE_BASE, dx)):
        print(f"❌ {dx} 不存在")
        sys.exit(1)
    t0 = time.time()
    print(f"\n=== BW合成: {dx} ===")

    ps = get_ps(timeout=15)
    ps.DisplayDialogs = 3  # DialogModes.NO
    process_dx(ps, dx)

    dt = time.time() - t0
    print(f"\n✓ {dx} BW合成完成 耗时 {dt:.1f}秒")
