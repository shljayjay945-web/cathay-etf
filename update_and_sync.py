# -*- coding: utf-8 -*-
"""
國泰投信量化分析系統 — 一鍵每日更新與 GitHub 雲端自動同步腳本
"""
import os
import sys
import subprocess
import webbrowser
import datetime

# 設定終端機 UTF-8 編碼
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def print_banner():
    print("=" * 70)
    print(" 🚀 國泰投信 人氣 ETF 與熱門基金 量化分析系統 — 一鍵更新與同步")
    print(f" 時間: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)

def main():
    print_banner()
    current_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(current_dir)

    # 1. 抓取最新官方數據
    print("\n[1/3] 正在向國泰投信官網與市場伺服器抓取最新淨值、配息與投資組合...")
    res1 = subprocess.run([sys.executable, "-u", "cathay_fund_analysis.py"])
    if res1.returncode != 0:
        print("\n❌ [1/3] 數據抓取過程發生異常，請檢查網路連線後重試。")
        return 1

    # 2. 重新編譯 HTML 互動儀表板
    print("\n[2/3] 正在重新編譯 HTML5 互動式儀表板 (含手機端卡片流適配)...")
    res2 = subprocess.run([sys.executable, "-u", "build_html_dashboard.py"])
    if res2.returncode != 0:
        print("\n❌ [2/3] HTML 儀表板生成失敗。")
        return 1

    # 3. 檢查變動並推送到 GitHub 雲端
    print("\n[3/3] 正在檢查是否有數據變動，並自動同步推送至 GitHub 雲端...")
    try:
        subprocess.run(["git", "add", "."], check=True)
        diff_check = subprocess.run(["git", "diff-index", "--quiet", "HEAD", "--"])
        if diff_check.returncode != 0:
            now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            subprocess.run(["git", "commit", "-m", f"chore: 自動更新最新基金與ETF數據 ({now_str})"], check=True)
            push_res = subprocess.run(["git", "push", "origin", "main"])
            if push_res.returncode == 0:
                print("\n🌐 GitHub 雲端儀表板同步更新成功！")
                print("📱 您的手機網址 (30秒後生效): https://shljayjay945-web.github.io/cathay-etf/")
            else:
                print("\n⚠️ 推送到 GitHub 發生異常，但本地數據與 Excel 已更新成功。")
        else:
            print("\nℹ️ 數據無變更，GitHub 雲端已是最新狀態。")
    except Exception as e:
        print(f"\n⚠️ Git 同步提示: {e} (本地資料已成功更新)")

    # 4. 在瀏覽器開啟本地檔案
    html_file = os.path.join(current_dir, "國泰投信_人氣ETF與熱門基金投資分析儀表板.html")
    if os.path.exists(html_file):
        print("\n" + "=" * 70)
        print(" ✅ 全部流程執行完畢！正在為您在瀏覽器開啟最新儀表板...")
        print("=" * 70)
        webbrowser.open(f"file:///{html_file.replace(os.sep, '/')}")

    return 0

if __name__ == "__main__":
    sys.exit(main())
