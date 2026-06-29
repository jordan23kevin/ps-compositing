"""黑T贴图处理 — 反相完成后调用，PS贴图+BW合成全覆盖"""
import os, sys, tempfile, time
from pathlib import Path
from PIL import Image
import numpy as np

ALPHA_THRESHOLD = 20
BASE = Path(r"D:\Semems WB\02_PROJECTS")
TORSO = Path(r"D:\Semems\1胚衣")

BACK_NEW = {
    "torso_white": "白背2.jpg", "torso_black": "黑背2.jpg",
    "scale_percent": 30, "rotation": 1,
    "target_center_x": 680, "target_top_y": 570,
}
FRONT_NEW = {
    "torso_white": "白正2.jpg", "torso_black": "黑正2.jpg",
    "scale_percent": 13.33, "rotation": 1,
    "target_center_x": 888, "target_top_y": 612,
}


def place_one(side, cfg, inv_path, dx, upload, torso_color="black"):
    """PS贴图：trim+缩放+贴到胚衣+保存"""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.DisplayDialogs = 3

        torso_file = cfg[f"torso_{torso_color}"]
        torso_path = str(TORSO / torso_file)
        scale = cfg["scale_percent"] / 100

        # Trim + 缩放
        img = Image.open(inv_path).convert("RGBA")
        a = np.array(img)
        alpha = a[:, :, 3]
        mask = alpha >= ALPHA_THRESHOLD
        ys, xs = np.where(mask)
        if len(ys) == 0:
            return None
        x0, x1 = xs.min(), xs.max() + 1
        y0, y1 = ys.min(), ys.max() + 1
        trimmed = img.crop((x0, y0, x1, y1))
        new_w = int((x1 - x0) * scale)
        new_h = int((y1 - y0) * scale)
        scaled = trimmed.resize((new_w, new_h), Image.LANCZOS)

        temp_dir = tempfile.gettempdir()
        temp_design = os.path.join(temp_dir, f"temp_{dx}_inv.png")
        scaled.save(temp_design, "PNG")

        move_x = cfg["target_center_x"] - new_w / 2
        move_y = cfg["target_top_y"]

        color_suffix = "白T" if torso_color == "white" else "黑T"
        output_name = f"{dx}_{side}_{color_suffix}.jpg"
        output_path = str(upload / output_name)

        jsx_path = os.path.join(os.path.dirname(__file__), "jsx", "place_design.jsx")
        with open(jsx_path, "r", encoding="utf-8") as jf:
            jsx = jf.read()
        jsx = jsx.replace("{{TORSO_PATH}}", torso_path.replace("\\", "\\\\"))
        jsx = jsx.replace("{{DESIGN_PATH}}", temp_design.replace("\\", "\\\\"))
        jsx = jsx.replace("{{OUTPUT_PATH}}", output_path.replace("\\", "\\\\"))
        jsx = jsx.replace("{{ROTATION}}", str(cfg["rotation"]))
        jsx = jsx.replace("{{MOVE_X}}", str(move_x))
        jsx = jsx.replace("{{MOVE_Y}}", str(move_y))

        temp_jsx = os.path.join(temp_dir, "temp_place.jsx")
        with open(temp_jsx, "w", encoding="utf-8") as jf:
            jf.write(jsx)
        psApp.DoJavaScriptFile(temp_jsx)
        return output_name
    except Exception as e:
        print(f"  ❌ PS错误({side},{torso_color}): {e}")
        return None
    finally:
        pythoncom.CoUninitialize()


def bw_synth(dx, upload):
    """BW合成：用PS动作合并B和W"""
    import win32com.client, pythoncom
    pythoncom.CoInitialize()
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.DisplayDialogs = 3
        b_img = str(upload / f"{dx}_B_黑T.jpg")
        w_img = str(upload / f"{dx}_W_黑T.jpg")
        out_path = str(upload / f"{dx}_黑BW.jpg")

        # 用PS动作合成（和ps_batch.py一致）
        shell = win32com.client.Dispatch("WScript.Shell")
        ps_exe = r"D:\Program Files\Adobe Photoshop 2025 v26.0\Adobe Photoshop 2025\Photoshop.exe"
        shell.Run(f'"{ps_exe}" "{w_img}" "{b_img}"', 1, False)
        import time; time.sleep(2)

        for _ in range(20):
            if psApp.Documents.Count >= 2: break
            time.sleep(0.5)

        for i in range(1, psApp.Documents.Count + 1):
            doc = psApp.Documents(i)
            if f"_B_黑T" in doc.Name:
                psApp.ActiveDocument = doc; break
        psApp.DoAction('黑', '正反图')

        jpgOpt = win32com.client.Dispatch("Photoshop.JPEGSaveOptions")
        jpgOpt.Quality = 12
        psApp.ActiveDocument.SaveAs(out_path, jpgOpt, True, 2)
        psApp.ActiveDocument.Close(2)
        for _ in range(psApp.Documents.Count, 0, -1):
            try: psApp.Documents(_).Close(2)
            except: pass
        return "✅ 黑BW 合成完成"
    except Exception as e:
        return f"❌ BW合成错误: {e}"
    finally:
        pythoncom.CoUninitialize()


def main():
    dx = sys.argv[1] if len(sys.argv) > 1 else ""
    if not dx:
        print("用法: python process_black.py DX0124")
        return
    print(f"=== 黑T贴图处理: {dx} ===")
    t0 = time.time()
    upload = BASE / dx / "03_UPLOAD"
    rembg = BASE / dx / "02_REM_BG"

    # 找黑版文件
    tasks = []
    for inv_path in sorted(rembg.glob(f"{dx}_黑*_cut.png")):
        letter = Path(inv_path).stem.replace(f"{dx}_黑", "").replace("_cut", "")
        if letter == "B":
            tasks.append(("B", BACK_NEW, inv_path))
        elif letter == "W":
            tasks.append(("W", FRONT_NEW, inv_path))
        elif letter == "WB":
            tasks.append(("B", BACK_NEW, inv_path))
            tasks.append(("W", FRONT_NEW, inv_path))

    if not tasks:
        print("❌ 未找到黑版_cut文件")
        return

    # 贴图：只做黑色胚衣版本
    print(f"\n--- 贴图 ({len(tasks)}张) ---")
    for side, cfg, inv_path in tasks:
        r = place_one(side, cfg, inv_path, dx, upload, "black")
        if r: print(f"  ✅ {r}")

    # BW合成
    has_b = any(s == "B" for s, _, _ in tasks)
    has_w = any(s == "W" for s, _, _ in tasks)
    if has_b and has_w:
        print("\n--- BW合成 ---")
        print(f"  {bw_synth(dx, upload)}")
    else:
        print("⏭️ 只有单面，跳过BW合成")

    dt = time.time() - t0
    print(f"\n⏱️ 总耗时: {dt:.1f}秒")
    print("=" * 40)


if __name__ == "__main__":
    main()
