"""
Contest Log Reader

コンテストログの音声読み上げアプリ
対応フォーマット：
 JARLコンテストログ R1.0

【実行環境の準備】
このプログラムを実行するには、以下のライブラリをインストールしてください。
コマンドプロンプトやターミナルで以下のコマンドを実行します：

    pip install pywin32

※ このプログラムは Windows 専用です（音声合成に SAPI5 を使用するため）。
"""

import queue
import threading
import tkinter as tk
from tkinter import filedialog, ttk
import re
import time
import win32com.client
import pythoncom

# -----------------------------------------------------------------------------
# 定数定義
# -----------------------------------------------------------------------------

# フォネティックコードの定義
PHONETIC = {
    'A': 'アルファ', 'B': 'ブラボー', 'C': 'チャーリー', 'D': 'デルタ', 'E': 'エコー',
    'F': 'フォックストロット', 'G': 'ゴルフ', 'H': 'ホテル', 'I': 'インディア', 'J': 'ジュリエット',
    'K': 'キロ', 'L': 'リマ', 'M': 'マイク', 'N': 'ノーベンバー', 'O': 'オスカー',
    'P': 'パパ', 'Q': 'ケベック', 'R': 'ロメオ', 'S': 'シエラ', 'T': 'タンゴ',
    'U': 'ユニフォーム', 'V': 'ビクター', 'W': 'ウィスキー', 'X': 'エクスレイ', 'Y': 'ヤンキー',
    'Z': 'ズール', '0': 'ゼロ', '1': 'ワン', '2': 'ツー', '3': 'スリー', '4': 'フォー',
    '5': 'ファイブ', '6': 'シックス', '7': 'セブン', '8': 'エイト', '9': 'ナイン'
}

# -----------------------------------------------------------------------------
# ユーティリティ関数
# -----------------------------------------------------------------------------

def make_ssml(text):
    """
    読み上げ用のSSML（Speech Synthesis Markup Language）を生成する。
    """
    return (
        '<speak version="1.0" '
        'xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="ja-JP">'
        'っ、 <break time="300ms"/>'
        + text +
        '</speak>'
    )

# -----------------------------------------------------------------------------
# メインアプリケーションクラス
# -----------------------------------------------------------------------------

class LogReader:
    """
    ログを読み込み、音声で読み上げるGUIアプリケーション。
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Log Reader")
        
        # 状態管理変数
        self.log_data = []
        self.contest_name = ""
        self.is_running = False
        self.sapi = None
        self.sapi_ready = False
        self._cmd_queue = queue.Queue()
        self.keep_alive = True 

        # UIの構築
        self.setup_ui()
        
        # ウィンドウ終了時のクリーンアップ処理を登録
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # バックグラウンドスレッドの開始
        threading.Thread(target=self._sapi_thread, daemon=True).start()
        threading.Thread(target=self._keep_audio_active, daemon=True).start()

    def _keep_audio_active(self):
        """
        オーディオデバイスをアクティブに保つためのスレッド処理。
        """
        pythoncom.CoInitialize()
        try:
            keep_sapi = win32com.client.Dispatch("SAPI.SpVoice")
            keep_sapi.Volume = 0 
            while self.keep_alive:
                if not self.is_running:
                    # 無音に近い状態で発話させ、デバイスのサスペンドを防ぐ
                    keep_sapi.Speak("、", 1) 
                time.sleep(1) 
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()

    def _sapi_thread(self):
        """
        SAPI音声エンジンの制御を行うスレッド処理。
        """
        pythoncom.CoInitialize()
        try:
            sapi = win32com.client.Dispatch("SAPI.SpVoice")
            # 日本語音声エンジンの選択
            for v in sapi.GetVoices():
                desc = v.GetDescription()
                if "Japanese" in desc or "Haruka" in desc:
                    sapi.Voice = v
                    break
            
            sapi.Speak("っ、。")
            self.sapi = sapi
            self.sapi_ready = True
            self.root.after(0, lambda: self.lbl_sapi.config(text="音声エンジン: 準備完了"))
        except Exception as e:
            self.root.after(0, lambda: self.lbl_sapi.config(text=f"エラー ({e})"))
            return

        while True:
            cmd = self._cmd_queue.get()
            if cmd is None: 
                break
            try:
                # 読み上げフラグのクリア
                sapi.Speak("っ、っ、", 8) 
                cmd(sapi)
            except Exception:
                pass
        pythoncom.CoUninitialize()

    def _speak(self, sapi, text):
        """
        SSMLまたはプレーンテキストで発話を実行する。
        """
        SVSFIsXML = 8
        try:
            sapi.Speak(make_ssml(text), SVSFIsXML)
        except Exception:
            try:
                sapi.Speak("っ、 " + re.sub(r"<[^>]+>", "", text))
            except Exception: 
                pass

    def _make_read_cmd(self, rows, total, speed):
        """
        連続読み上げ用のコマンド関数を生成する。
        """
        def cmd(sapi):
            sapi.Rate = speed  # SAPIは生成したスレッド内で操作する
            self._speak(sapi, self.contest_name)
            for i, row in enumerate(rows):
                if not self.is_running:
                    self.root.after(0, lambda: self.lbl_status.config(text="停止"))
                    return
                
                self.root.after(0, lambda n=i+1: self.lbl_status.config(text=f"読み上げ中: {n}/{total}"))
                say_text = self.create_speech_text(row, rows[i-1] if i > 0 else None)
                self._speak(sapi, say_text)

            self._speak(sapi, "終わり")
            self.is_running = False
            self.root.after(0, lambda: self.lbl_status.config(text="終了"))
        return cmd

    def setup_ui(self):
        """
        GUI画面のレイアウトを構築する。
        """
        # ファイル選択エリア
        top_frame = ttk.Frame(self.root)
        top_frame.pack(pady=10, padx=10, fill="x")
        ttk.Button(top_frame, text="ログファイルを開く", command=self.load_file).pack(side="left")

        # 情報表示エリア
        info_frame = ttk.Frame(self.root)
        info_frame.pack(fill="x", padx=10)
        
        self.lbl_contest = ttk.Label(info_frame, text="コンテスト名: 未読込", font=("", 10, "bold"))
        self.lbl_contest.pack(anchor="w")
        
        self.lbl_status = ttk.Label(info_frame, text="ファイルを選択してください")
        self.lbl_status.pack(anchor="w")
        
        self.lbl_sapi = ttk.Label(info_frame, text="音声エンジン: 初期化中…", foreground="gray")
        self.lbl_sapi.pack(anchor="w")

        # 設定エリア
        config_frame = ttk.LabelFrame(self.root, text="読み上げ項目設定")
        config_frame.pack(pady=10, padx=10, fill="both")

        # 日付設定
        date_f = ttk.Frame(config_frame)
        date_f.pack(anchor="w", pady=2)
        self.var_date_mode = tk.StringVar(value="月日")
        ttk.Label(date_f, text="日付：", width=8).pack(side="left")
        for m in ["年月日", "月日", "読まない"]:
            ttk.Radiobutton(date_f, text=m, value=m, variable=self.var_date_mode).pack(side="left", padx=5)

        # BAND/MODE/SENT設定
        options = [("BAND：", "var_read_band"), ("MODE：", "var_read_mode"), ("SENT：", "var_read_sent")]
        for label, var_name in options:
            f = ttk.Frame(config_frame)
            f.pack(anchor="w", pady=2)
            setattr(self, var_name, tk.StringVar(value="読む"))
            ttk.Label(f, text=label, width=8).pack(side="left")
            for m in ["読む", "読まない"]:
                ttk.Radiobutton(f, text=m, value=m, variable=getattr(self, var_name)).pack(side="left", padx=5)

        # 前行と同じBAND・MODEの省略設定
        self.var_same = tk.BooleanVar(value=False)
        ttk.Checkbutton(config_frame, text="前行と同じBAND・MODEは「同じ」と読む",
                        variable=self.var_same).pack(anchor="w", pady=2)

        # 速度設定
        spd_f = ttk.LabelFrame(self.root, text="速度")
        spd_f.pack(fill="x", padx=10, pady=5)
        self.scale_speed = tk.Scale(spd_f, from_=-5, to=+5, orient="horizontal")
        self.scale_speed.set(0)
        self.scale_speed.pack(fill="x")

        # 操作ボタン
        btn_f = ttk.Frame(self.root)
        btn_f.pack(pady=10)
        ttk.Button(btn_f, text="読み上げ開始", command=self.start_reading).pack(side="left", padx=5)
        ttk.Button(btn_f, text="停止", command=self.stop_reading).pack(side="left", padx=5)

    def load_file(self):
        """
        ログファイルを選択・解析し、データをメモリに保持する。
        """
        path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt;*.log")])
        if not path: 
            return
            
        try:
            with open(path, "r", encoding="shift-jis") as f:
                content = f.read()
            
            # コンテスト名の抽出
            name_match = re.search(r"<CONTESTNAME>(.*?)</CONTESTNAME>", content)
            self.contest_name = name_match.group(1) if name_match else "不明"
            self.lbl_contest.config(text=f"コンテスト名: {self.contest_name}")
            
            # ログデータの抽出
            self.log_data = []
            log_section = re.search(r"<LOGSHEET.*?>(.*?)</LOGSHEET>", content, re.DOTALL)
            if log_section:
                lines = log_section.group(1).strip().split("\n")
                for line in lines:
                    if "DATE" in line or "----" in line or not line.strip(): 
                        continue
                    parts = line.split()
                    if len(parts) >= 9:
                        self.log_data.append(parts)
            
            self.lbl_status.config(text=f"読込完了: {len(self.log_data)}交信")
        except Exception as e:
            self.lbl_status.config(text=f"読込エラー: {e}")

    def create_speech_text(self, row, prev=None):
        """
        1行のログデータから読み上げ用の文字列を生成する。
        """
        parts = ["っ"] 
        
        # 日付処理
        d_m = self.var_date_mode.get()
        if d_m != "読まない":
            d_p = row[0].split("-")
            if len(d_p) == 3:
                y, m, d = int(d_p[0]), int(d_p[1]), int(d_p[2])
                if d_m == "年月日": 
                    parts.append(f"{y}年{m}月{d}日")
                elif d_m == "月日": 
                    parts.append(f"{m}月{d}日")

        # 時間（1文字ずつ分解）
        t_num = re.sub(r'\D', '', row[1])
        parts.append(" ".join(list(t_num)))

        # バンド・モード
        items = []  # (今回の値, 前行の値) の組。読む項目だけ入れる
        if self.var_read_band.get() == "読む":
            items.append((row[2], prev[2] if prev else None))
        if self.var_read_mode.get() == "読む":
            items.append((row[3], prev[3] if prev else None))
        if items:
            # 読む項目がすべて前行と同じ場合だけ「同じ」を1回読む
            if self.var_same.get() and prev and all(cur == old for cur, old in items):
                parts.append("同じ")
            else:
                parts.extend(cur for cur, _ in items)

        # コールサイン処理
        call = row[4].upper()
        # プリフィックス（数字まで）とサフィックスを分離
        match = re.match(r'^([A-Z0-9]{1,2}[0-9])([A-Z0-9/]+)$', call)
        if match:
            prefix, suffix = match.groups()
            p_prefix = [PHONETIC[c] if c in PHONETIC else c for c in prefix]
            p_suffix = [PHONETIC[c] if c in PHONETIC else ("ポータブル" if c == '/' else c) for c in suffix]
            # プリフィックスとサフィックスの間に微小なポーズを挿入
            call_speech = " ".join(p_prefix) + '<break time="80ms"/>' + " ".join(p_suffix)
            parts.append(call_speech)
        else:
            p_call = [PHONETIC[c] if c in PHONETIC else ("ポータブル" if c == '/' else c) for c in call]
            parts.append(" ".join(p_call))

        # 送信・受信コンテストナンバー
        if self.var_read_sent.get() == "読む":
            parts.append(" ".join(list(row[5])) + " " + " ".join(list(row[6])))

        parts.append(" ".join(list(row[7])) + " " + " ".join(list(row[8])))

        return " 、 ".join(parts)

    def start_reading(self):
        """
        読み上げプロセスを開始する。
        """
        if not self.sapi_ready: 
            return
        if not self.log_data or self.is_running: 
            return
            
        speed = int(self.scale_speed.get())  # SAPIのRate範囲(-10〜10)内
        
        self.is_running = True
        rows = list(self.log_data)
        self._cmd_queue.put(self._make_read_cmd(rows, len(rows), speed))

    def stop_reading(self):
        """
        読み上げプロセスを停止する。
        """
        self.is_running = False

    def on_closing(self):
        """
        ウィンドウを閉じる際のクリーンアップ処理。
        """
        self.keep_alive = False  # オーディオ維持ループを停止
        self.is_running = False  # 読み上げループを停止
        
        # SAPIスレッドのループを抜けるためのシグナルを送る
        self._cmd_queue.put(None)
        
        # 少し待機してからウィンドウを破棄（スレッドのクリーンアップ時間を確保）
        self.root.after(100, self.root.destroy)
        
# -----------------------------------------------------------------------------
# エントリーポイント
# -----------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = LogReader(root)
    root.mainloop()