"""
國泰投信投資分析儀表板 - 手機區域網路快速分享伺服器 (附 QR Code)
"""
import os
import sys
import socket
import http.server
import socketserver
import urllib.parse

def get_local_ip():
    """取得電腦在區域網路 (Wi-Fi) 中的真實 IP 位址"""
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        # 連接外部公用 DNS（不發送封包，僅用於偵測出網介面 IP）
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return '127.0.0.1'

def start_server():
    port = 8080
    local_ip = get_local_ip()
    url = f"http://{local_ip}:{port}/"

    print("=" * 60)
    print(" 📱 國泰投信投資分析儀表板 — 手機即時查看服務已啟動！")
    print("=" * 60)
    print(f" 電腦區網 IP: {local_ip}")
    print(f" 手機專屬訪問網址:")
    print(f" 👉 {url}")
    print("-" * 60)

    try:
        import qrcode
        qr = qrcode.QRCode(border=1)
        qr.add_data(url)
        qr.make(fit=True)
        print(" 請用手機相機（或 LINE 掃描器）對準下方 QR Code 直接掃描：\n")
        qr.print_ascii(invert=True)
    except Exception as e:
        print(f" (QR Code 生成跳過: {e})")

    print("\n" + "=" * 60)
    print(" 💡 貼心提醒：請確保您的【手機】與【這台電腦】連線至同一個 Wi-Fi 網路！")
    print(" 若要關閉服務，直接關閉此視窗，或在視窗內按下鍵盤 Ctrl + C 即可。")
    print("=" * 60 + "\n")

    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    class CustomHandler(http.server.SimpleHTTPRequestHandler):
        def end_headers(self):
            # 防止手機快取，確保看到最新資料
            self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
            self.send_header('Pragma', 'no-cache')
            self.send_header('Expires', '0')
            super().end_headers()

    # 允許地址重用，避免端口佔用報錯
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(("", port), CustomHandler) as httpd:
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n服務已停止。")

if __name__ == '__main__':
    start_server()
