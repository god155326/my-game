#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
不朽之旅 - 免快取本機伺服器
============================================================
跟 start.bat 預設用的 `python -m http.server` 幾乎一模一樣，
唯一差別：每個檔案的回應都會強制加上「不要快取」的 HTTP 標頭
(Cache-Control / Pragma / Expires)，讓瀏覽器每次重新整理頁面時
都會直接跟這個伺服器要最新的檔案內容，不會沿用舊版的
data/不朽之旅.xlsx 或 assets/images/ 底下的圖片。

不需要單獨執行這個 .py 檔——雙擊 start-nocache.bat 就會自動呼叫它。

跟 index.html 內建的 ASSET_VERSION 快取破壞機制是互補關係：
ASSET_VERSION 那個修正在「GitHub Pages 正式站」也有效(因為是網址帶版本號，
不管伺服器端有沒有設定快取標頭都有用)；這支免快取伺服器則是進一步在
「本機測試」時，從伺服器端直接把快取整個關掉，雙重保險。
============================================================
"""
import functools
import http.server
import os
import socketserver
import sys

# Windows 的命令提示字元預設不是用 UTF-8 顯示文字(常是簡體/繁體的舊版編碼，
# 或甚至是純英文的舊編碼)，直接 print() 中文字串在那種情況下會噴
# UnicodeEncodeError 讓整支程式當掉、伺服器根本沒真的啟動起來
# (瀏覽器端就會看到「拒絕連線」)。這裡強制把 stdout/stderr 換成 UTF-8，
# 遇到真的無法顯示的字元就用替代符號代替，不會再讓程式崩潰。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

PORT = 8787

# 一定要用「這支 .py 檔自己所在的資料夾」，不要依賴啟動當下的「目前工作目錄」。
# 目前工作目錄會受到「怎麼被啟動」影響(捷徑的起始位置、cmd 視窗當時所在的
# 資料夾等等)，就算 start-nocache.bat 裡有 cd /d "%~dp0"，只要中間有任何
# 環節走樣，Python 實際的工作目錄就可能不是遊戲資料夾，這時 index.html
# 就會抓到別的地方去、變成「404 找不到」。改用這支程式自己的路徑當基準，
# 不管從哪裡被啟動，永遠都會抓到正確的遊戲資料夾。
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


class NoCacheHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def log_message(self, format, *args):
        # 精簡輸出：只顯示異常狀態碼(4xx/5xx)的請求，正常的200/304不洗版面
        status = str(args[1]) if len(args) >= 2 else ""
        if not status.startswith(("2", "3")):
            super().log_message(format, *args)


def main():
    # 注意：這裡故意「不」開啟 allow_reuse_address。
    # 如果之前已經有一個伺服器(不管是這支程式還是 start.bat)還在背景執行、
    # 占用著 8787 port，Windows 有時候會允許新的伺服器也綁定同一個 port，
    # 結果變成兩個伺服器同時存在，瀏覽器的請求會隨機被其中一個接到——
    # 如果剛好接到的是舊的那個(可能少了某些檔案或設定不同)，就會出現
    # 「index.html 404 找不到」這種讓人一頭霧水的錯誤。
    # 不開啟這個選項，遇到 port 被占用時就會直接跳出明確的錯誤訊息，
    # 提示你先執行 stop-server.bat，而不是悄悄地同時跑兩個伺服器。
    handler = functools.partial(NoCacheHTTPRequestHandler, directory=SCRIPT_DIR)
    try:
        with socketserver.TCPServer(("", PORT), handler) as httpd:
            print(f"[不朽之旅] 免快取伺服器已啟動：http://localhost:{PORT}/index.html")
            print(f"遊戲資料夾：{SCRIPT_DIR}")
            print("已強制關閉瀏覽器快取，重新整理頁面一定會抓到最新的 xlsx / 圖片。")
            print("關閉這個視窗即可停止伺服器。")
            httpd.serve_forever()
    except OSError as e:
        print(f"[不朽之旅] 啟動失敗：{e}")
        print(f"可能是 port {PORT} 已經被其他程式占用(例如 start.bat 還在跑)，")
        print("請先執行 stop-server.bat 關閉舊的伺服器，再重新執行這個檔案。")
        sys.exit(1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
