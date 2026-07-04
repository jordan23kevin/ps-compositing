"""安全退出 Photoshop。
先尝试 COM Quit()；若 PS 无响应或失败，则强制结束 Photoshop.exe 进程。
注意：不会主动启动 Photoshop；若 PS 未运行则直接返回。
"""
import sys
import os
import time
import subprocess

# 强制 stdout/stderr 使用 UTF-8，避免 Windows GBK 控制台打印特殊字符时崩溃
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout = open(sys.stdout.fileno(), 'w', encoding='utf-8', closefd=False)
        sys.stderr = open(sys.stderr.fileno(), 'w', encoding='utf-8', closefd=False)
    except Exception:
        pass

sys.path.insert(0, os.path.dirname(__file__))


def _ps_is_running():
    """检查是否有 Photoshop.exe 进程在运行。"""
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Photoshop.exe"],
            capture_output=True, text=True, timeout=5
        )
        return "Photoshop.exe" in result.stdout
    except Exception:
        return False


def kill_photoshop():
    """强制结束 Photoshop 相关进程。"""
    for name in ("Photoshop.exe", "AdobeIPCBroker.exe", "Adobe Crash Processor.exe"):
        try:
            subprocess.run(["taskkill", "/IM", name, "/F"], capture_output=True, timeout=10)
        except Exception:
            pass


def quit_photoshop():
    """先尝试优雅退出，失败则强制结束。"""
    if not _ps_is_running():
        print("Photoshop 未运行，无需退出", flush=True)
        return

    try:
        import win32com.client
        try:
            # 只连接已运行的 PS，不要启动新实例
            ps = win32com.client.GetObject(Class="Photoshop.Application")
            ps.Quit()
        except Exception:
            pass

        # 给 PS 最多 5 秒自行退出
        for _ in range(10):
            time.sleep(0.5)
            if not _ps_is_running():
                print("Photoshop 已退出", flush=True)
                return
        print("Photoshop 未在 5 秒内退出，强制结束进程", flush=True)
    except Exception as e:
        print(f"连接 Photoshop 失败: {e}，尝试强制结束", flush=True)

    kill_photoshop()
    if not _ps_is_running():
        print("Photoshop 已强制结束", flush=True)
    else:
        print("强制结束 Photoshop 失败，请手动到任务管理器结束", flush=True)


if __name__ == "__main__":
    quit_photoshop()
