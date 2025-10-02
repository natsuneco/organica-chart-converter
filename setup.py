# setup.py
import sys
from cx_Freeze import setup, Executable

# --- 基本設定 ---
# Windows GUIアプリケーションとして設定
base = None
if sys.platform == "win32":
    base = "Win32GUI"

# main.pyをビルド対象にし、baseを設定する
executables = [Executable("main.py", base=base, icon="chart_convert_icon.ico")]

# --- セットアップ情報の記述 ---
setup(
    name="Organica 譜面変換ツール",
    version="0.0.1",
    executables=executables
)
