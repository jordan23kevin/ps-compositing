"""PS贴图 单款入口 v2.1 — 被 check_rem.py 调用处理单个 DX

变更 v2.1：
  - 作为 check_rem.py 贴图流水线第二步，由 /ps-sticker 自动调用
  - 黑T优先逻辑在 wb_sticker_ps.py 中实现：存在 _黑B/_黑W/_黑BW 时通用图不再输出黑T
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
from wb_sticker_ps import process_dx_folder
from config import SOURCE_BASE

if __name__ == "__main__":
    dx = sys.argv[1]
    dx_folder = os.path.join(SOURCE_BASE, dx)
    if not os.path.isdir(dx_folder):
        print(f"❌ {dx_folder} 不存在")
        sys.exit(1)
    t0 = time.time()
    print(f"\n=== PS贴图: {dx} ===")
    process_dx_folder(dx_folder)
    dt = time.time() - t0
    print(f"✓ {dx} 完成，耗时 {dt:.1f}秒")
