"""PS贴图 单款入口 v2.1 — 被 check_rem.py 调用处理单个 DX

变更 v2.1：
  - 作为 check_rem.py 贴图流水线第二步，由 /ps-sticker 自动调用
  - 黑T优先逻辑在 wb_sticker_ps.py 中实现：存在 _黑B/_黑W/_黑BW 时通用图不再输出黑T
"""
import sys, os, time
sys.path.insert(0, os.path.dirname(__file__))
sys.path.insert(0, r"E:\Claude code\ps")
try:
    import wb_meta
except Exception:
    wb_meta = None
from wb_sticker_ps import process_dx_folder
from config import SOURCE_BASE

if __name__ == "__main__":
    dx = sys.argv[1]
    only_color = None
    if "--only-color" in sys.argv:
        i = sys.argv.index("--only-color")
        if i + 1 < len(sys.argv):
            only_color = sys.argv[i + 1]
    dx_folder = os.path.join(SOURCE_BASE, dx)
    if not os.path.isdir(dx_folder):
        print(f"❌ {dx_folder} 不存在")
        sys.exit(1)
    t0 = time.time()
    print(f"\n=== PS贴图: {dx} {'(仅'+only_color+')' if only_color else ''} ===")
    process_dx_folder(dx_folder, only_color=only_color)
    dt = time.time() - t0
    print(f"✓ {dx} 完成，耗时 {dt:.1f}秒")
