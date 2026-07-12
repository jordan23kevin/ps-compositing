"""PS BW合成 单款入口 v2.3 — 被 check_rem.py 调用处理单个 DX 的 BW 合成

变更 v2.3：
  - BW 合成已改为纯 PIL 实现（不再依赖 Photoshop 私有动作集），
    入口不再强制连接 PS，避免独立「BW合成」按钮因 PS 未启动而失败。

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
from ps_batch import process_dx
from config import SOURCE_BASE

if __name__ == "__main__":
    dx = sys.argv[1]
    if not os.path.isdir(os.path.join(SOURCE_BASE, dx)):
        print(f"❌ {dx} 不存在")
        sys.exit(1)
    t0 = time.time()
    print(f"\n=== BW合成: {dx} ===")

    # BW 合成已改为纯 PIL，无需连接 Photoshop
    process_dx(None, dx)

    dt = time.time() - t0
    print(f"\n✓ {dx} BW合成完成 耗时 {dt:.1f}秒")
