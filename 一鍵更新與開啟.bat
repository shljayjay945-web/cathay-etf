@echo off
@chcp 65001 > nul
title 國泰投信量化分析系統 - 一鍵每日更新與GitHub雲端同步

echo ======================================================================
echo  🚀 國泰投信 人氣 ETF 與熱門基金 量化分析系統 — 一鍵更新與同步
echo ======================================================================
echo.
echo [1/3] 正在向國泰投信官網與市場伺服器抓取最新淨值、配息與投資組合...
python cathay_fund_analysis.py
if errorlevel 1 (
    echo.
    echo ❌ 數據抓取過程發生異常，請檢查網路連線。
    pause
    exit /b 1
)

echo.
echo [2/3] 正在重新編譯 HTML5 互動式儀表板 (含手機端適配)...
python build_html_dashboard.py
if errorlevel 1 (
    echo.
    echo ❌ HTML 生成過程發生異常。
    pause
    exit /b 1
)

echo.
echo [3/3] 正在檢查是否有新數據，並自動同步推送至 GitHub 雲端...
git add .
git diff-index --quiet HEAD --
if errorlevel 1 (
    git commit -m "chore: 自動更新最新基金與ETF數據 (%date% %time%)"
    git push origin main
    if errorlevel 1 (
        echo.
        echo ⚠️ 推送到 GitHub 發生異常，但本地數據與 Excel 已更新成功。
    ) else (
        echo.
        echo 🌐 GitHub 雲端儀表板同步更新成功！
        echo 📱 手機稍後重新整理網頁即可查看最新數據。
    )
) else (
    echo.
    echo ℹ️ 數據無變動，GitHub 雲端已是最新狀態。
)

echo.
echo ======================================================================
echo  ✅ 全部流程執行完畢！正在為您開啟最新儀表板...
echo ======================================================================
start "" "國泰投信_人氣ETF與熱門基金投資分析儀表板.html"
timeout /t 3 > nul
exit
