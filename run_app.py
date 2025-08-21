import subprocess
import sys
import os

def install_requirements():
    """必要なライブラリをインストール"""
    print("必要なライブラリをインストール中...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
    print("ライブラリのインストールが完了しました。")

def run_streamlit_app():
    """Streamlitアプリを起動"""
    print("Streamlitアプリを起動中...")
    subprocess.run([sys.executable, "-m", "streamlit", "run", "streamlit_app.py"])

if __name__ == "__main__":
    try:
        # 必要なライブラリをインストール
        install_requirements()
        
        # Streamlitアプリを起動
        run_streamlit_app()
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        input("Enterキーを押して終了してください...")