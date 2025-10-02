# -*- coding: utf-8 -*-
import mido
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess

# --- 設定 ---
FORMAT_VERSION = 1
# MIDIノート F3 から B3 をレーン 0 から 6 にマッピングします
NOTE_RANGE_START = 53  # F3のMIDIノート番号
NOTE_RANGE_END = 59    # B3のMIDIノート番号
CRITICAL_VELOCITY_THRESHOLD = 120  # このベロシティ以上のノートはcriticalになります
DEFAULT_OFFSET = 0   # JSONに出力されるオフセット値


def midi_to_json_score(midi_path, output_path):
    """
    MIDIファイルを7レーンのリズムゲーム用JSON譜面データに変換します。

    Args:
        midi_path (str): 入力するMIDIファイルのパス
        output_path (str): 出力するJSONファイルのパス

    Returns:
        dict or None: 成功した場合は変換後のJSONデータを、失敗した場合はNoneを返します。
    """
    try:
        mid = mido.MidiFile(midi_path)
    except Exception as e:
        # GUIでエラーを表示するために、呼び出し元で処理
        raise IOError(f"MIDIファイルの読み込み中にエラーが発生しました: {e}")

    tpb = mid.ticks_per_beat
    notes_data = []

    # --- 初期のBPMと曲名を特定 ---
    initial_bpm = 120  # デフォルトBPM
    title = os.path.splitext(os.path.basename(midi_path))[0]  # デフォルトの曲名

    found_first_tempo = False
    # 通常、メタ情報はトラック0に含まれる
    for msg in mid.tracks[0]:
        if msg.is_meta and msg.type == 'set_tempo' and not found_first_tempo:
            initial_bpm = round(mido.tempo2bpm(msg.tempo), 3)
            found_first_tempo = True
        if msg.is_meta and msg.type == 'track_name':
            title = msg.name

    # --- ノートとイベントの処理 ---
    active_notes = {}  # 処理中のノートオンイベントを格納: {note: {'tick': start_tick, 'velocity': velocity}}

    current_tick = 0
    # 全トラックをマージして時系列順にメッセージを処理
    for msg in mido.merge_tracks(mid.tracks):
        current_tick += msg.time  # デルタタイムを累積して絶対Tickを計算

        # BPM変更イベントの処理 (tick 0 の初期BPMは除く)
        if msg.is_meta and msg.type == 'set_tempo':
            # 既に同じtickにBPM変更イベントがないか確認
            if not any(e.get('type') == '_bpm' and e.get('tick') == current_tick for e in notes_data):
                bpm_event = {
                    "type": "_bpm",
                    "tick": current_tick,
                    "bpm": round(mido.tempo2bpm(msg.tempo), 3)
                }
                # 初期BPMとして設定済みの場合はイベントとして追加しない
                if current_tick > 0:
                    notes_data.append(bpm_event)

        # ノートオンイベントの処理
        elif msg.type == 'note_on' and msg.velocity > 0:
            if NOTE_RANGE_START <= msg.note <= NOTE_RANGE_END:
                # 絶対tickとベロシティを記録
                active_notes[msg.note] = {
                    'tick': current_tick,
                    'velocity': msg.velocity
                }

        # ノートオフイベントの処理 (ベロシティ0のノートオンもノートオフとして扱う)
        elif msg.type == 'note_off' or (msg.type == 'note_on' and msg.velocity == 0):
            if msg.note in active_notes:
                start_note = active_notes.pop(msg.note)
                start_tick = start_note['tick']
                velocity = start_note['velocity']
                duration = current_tick - start_tick

                # 長さが0のノートは無視
                if duration == 0:
                    continue

                # ノートのタイプを決定
                note_type = "critical" if velocity >= CRITICAL_VELOCITY_THRESHOLD else "normal"

                note_event = {
                    "lane": msg.note - NOTE_RANGE_START,
                    "tick": start_tick
                }

                # ロングノートか判定 (長さが1拍より長い)
                if duration > tpb:
                    note_event["type"] = "long"
                    note_event["duration"] = duration
                else:
                    note_event["type"] = note_type

                notes_data.append(note_event)

    # 全てのノートとイベントをtick順にソート
    notes_data.sort(key=lambda x: (x['tick'], 0 if x['type'] == '_bpm' else 1))

    # --- 最終的なJSONオブジェクトの構築 ---
    final_json = {
        "version": FORMAT_VERSION,
        "title": title,
        "bpm": initial_bpm,
        "tpb": tpb,
        "offset": DEFAULT_OFFSET,
        "notes": notes_data
    }

    # --- ファイルへの書き込み ---
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(final_json, f, indent=2, ensure_ascii=False)
    except Exception as e:
        raise IOError(f"JSONファイルへの書き込み中にエラーが発生しました: {e}")

    return final_json


class Application(tk.Frame):
    """Tkinter GUIアプリケーションのメインクラス"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title('Organica 譜面変換ツール')
        self.master.geometry('480x550')
        self.master.iconbitmap(default="chart_converter_icon.ico")
        self.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # --- ファイルパス変数 ---
        self.midi_file_path = ""
        self.output_json_path = ""
        self.music_file_path_convert = tk.StringVar()
        self.json_file_path_preview = tk.StringVar()
        self.music_file_path_preview = tk.StringVar()

        # --- 表示用変数 ---
        self.filepath_var = tk.StringVar(value="MIDIファイルが選択されていません")
        self.title_var = tk.StringVar(value="-")
        self.bpm_var = tk.StringVar(value="-")
        self.tpb_var = tk.StringVar(value="-")
        self.normal_notes_var = tk.StringVar(value="-")
        self.critical_notes_var = tk.StringVar(value="-")
        self.long_notes_var = tk.StringVar(value="-")
        self.bpm_changes_var = tk.StringVar(value="-")
        self.total_notes_var = tk.StringVar(value="-")

        self.create_widgets()

    def create_widgets(self):
        """GUIのウィジェットを作成・配置する"""
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill='both')

        convert_tab = ttk.Frame(notebook, padding=10)
        preview_tab = ttk.Frame(notebook, padding=10)

        notebook.add(convert_tab, text='変換')
        notebook.add(preview_tab, text='プレビュー')

        self.create_convert_tab(convert_tab)
        self.create_preview_tab(preview_tab)

    def create_convert_tab(self, parent):
        """「変換」タブのウィジェットを作成"""
        # --- ファイル選択フレーム ---
        file_frame = ttk.LabelFrame(parent, text="ファイル選択")
        file_frame.pack(fill=tk.X, padx=5, pady=5)

        ttk.Label(file_frame, text="MIDI:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Button(file_frame, text="ファイルを開く", command=self.select_midi_file).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.filepath_var, wraplength=300).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(file_frame, text="楽曲:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Button(file_frame, text="ファイルを開く", command=lambda: self.select_music_file(self.music_file_path_convert)).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.music_file_path_convert, wraplength=300).grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)

        file_frame.columnconfigure(2, weight=1)

        # --- メタデータ表示フレーム ---
        meta_frame = ttk.LabelFrame(parent, text="メタデータ")
        meta_frame.pack(fill=tk.X, padx=5, pady=5, expand=True)
        # ... (以下、既存のメタデータとノーツ数の表示ウィジェット)
        meta_frame.columnconfigure(1, weight=1)
        ttk.Label(meta_frame, text="Title:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(meta_frame, textvariable=self.title_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(meta_frame, text="BPM:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(meta_frame, textvariable=self.bpm_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(meta_frame, text="TPB:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(meta_frame, textvariable=self.tpb_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)

        # --- ノーツ数表示フレーム ---
        notes_frame = ttk.LabelFrame(parent, text="ノーツ数")
        notes_frame.pack(fill=tk.X, padx=5, pady=5)
        notes_frame.columnconfigure(1, weight=1)
        notes_frame.columnconfigure(3, weight=1)
        ttk.Label(notes_frame, text="Normal:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, textvariable=self.normal_notes_var).grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, text="Critical:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, textvariable=self.critical_notes_var).grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, text="Long:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, textvariable=self.long_notes_var).grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, text="BPM Changes:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, textvariable=self.bpm_changes_var).grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        ttk.Separator(notes_frame, orient='horizontal').grid(row=3, column=0, columnspan=4, sticky='ew', pady=5)
        ttk.Label(notes_frame, text="Total Notes:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        ttk.Label(notes_frame, textvariable=self.total_notes_var, font=('Helvetica', 10, 'bold')).grid(row=4, column=1, sticky=tk.W, padx=5, pady=2)

        # --- アクションボタンフレーム ---
        actions_frame = ttk.Frame(parent)
        actions_frame.pack(pady=20)
        ttk.Button(actions_frame, text="変換実行", command=self.convert_file).pack(side=tk.LEFT, padx=5)
        self.preview_button_convert = ttk.Button(actions_frame, text="プレビュー", command=self.preview_chart_from_convert, state=tk.DISABLED)
        self.preview_button_convert.pack(side=tk.LEFT, padx=5)

    def create_preview_tab(self, parent):
        """「プレビュー」タブのウィジェットを作成"""
        # --- ファイル選択フレーム ---
        file_frame = ttk.LabelFrame(parent, text="ファイル選択")
        file_frame.pack(fill=tk.X, padx=5, pady=5, expand=True)

        ttk.Label(file_frame, text="譜面 (JSON):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Button(file_frame, text="ファイルを開く", command=lambda: self.select_json_file(self.json_file_path_preview)).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.json_file_path_preview, wraplength=300).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

        ttk.Label(file_frame, text="楽曲:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        ttk.Button(file_frame, text="ファイルを開く", command=lambda: self.select_music_file(self.music_file_path_preview)).grid(row=1, column=1, padx=5, pady=5)
        ttk.Label(file_frame, textvariable=self.music_file_path_preview, wraplength=300).grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)

        file_frame.columnconfigure(2, weight=1)

        # --- プレビューボタン ---
        actions_frame = ttk.Frame(parent)
        actions_frame.pack(pady=20)
        ttk.Button(actions_frame, text="プレビュー実行", command=self.preview_chart_from_tab).pack()

    def select_midi_file(self):
        """ファイルダイアログを開き、MIDIファイルを選択する"""
        path = filedialog.askopenfilename(title="MIDIファイルを選択", filetypes=[("MIDI files", "*.mid"), ("All files", "*.*")])
        if path:
            self.midi_file_path = path
            self.filepath_var.set(os.path.basename(path))
            self.reset_info()
            # 同名の音楽ファイルを自動で探す
            base, _ = os.path.splitext(path)
            for ext in ['.mp3', '.wav', '.ogg']:
                music_path = base + ext
                if os.path.exists(music_path):
                    self.music_file_path_convert.set(music_path)
                    break

    def select_music_file(self, string_var):
        """音楽ファイルを選択し、指定されたStringVarにパスをセットする"""
        path = filedialog.askopenfilename(title="楽曲ファイルを選択", filetypes=[("Audio Files", "*.mp3 *.wav *.ogg"), ("All files", "*.*")])
        if path:
            string_var.set(path)

    def select_json_file(self, string_var):
        """JSONファイルを選択し、指定されたStringVarにパスをセットする"""
        path = filedialog.askopenfilename(title="譜面ファイルを選択", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if path:
            string_var.set(path)

    def convert_file(self):
        """ファイルを変換し、結果をGUIに表示する"""
        if not self.midi_file_path:
            messagebox.showwarning("エラー", "最初にMIDIファイルを選択してください。")
            return

        self.output_json_path = os.path.splitext(self.midi_file_path)[0] + ".json"

        try:
            result_json = midi_to_json_score(self.midi_file_path, self.output_json_path)
            # ... (メタデータとノーツ数の更新処理)
            self.title_var.set(result_json['title'])
            self.bpm_var.set(result_json['bpm'])
            self.tpb_var.set(result_json['tpb'])
            counts = {'normal': 0, 'critical': 0, 'long': 0, '_bpm': 0}
            for note in result_json['notes']:
                if note['type'] in counts:
                    counts[note['type']] += 1
            total_notes = counts['normal'] + counts['critical'] + counts['long']
            self.normal_notes_var.set(counts['normal'])
            self.critical_notes_var.set(counts['critical'])
            self.long_notes_var.set(counts['long'])
            self.bpm_changes_var.set(counts['_bpm'])
            self.total_notes_var.set(total_notes)

            self.preview_button_convert.config(state=tk.NORMAL)
            messagebox.showinfo("成功", f"変換が完了しました。\n'{self.output_json_path}' に保存されました。")
        except Exception as e:
            messagebox.showerror("変換エラー", str(e))

    def launch_chart_player(self, json_path, music_path):
        """Organica Chart Player.exeを起動する共通関数"""
        chart_player_exe = 'Organica Chart Player.exe'

        if not json_path or not os.path.exists(json_path):
            messagebox.showwarning("プレビューエラー", f"譜面ファイルが見つかりません:\n{json_path}")
            return
        if not music_path or not os.path.exists(music_path):
            messagebox.showwarning("プレビューエラー", f"楽曲ファイルが見つかりません:\n{music_path}")
            return
        if not os.path.exists(chart_player_exe):
            messagebox.showerror("プレビューエラー", f"'{chart_player_exe}' が見つかりません。\nスクリプトと同じディレクトリに配置してください。")
            return

        try:
            subprocess.Popen([chart_player_exe, json_path, music_path])
        except Exception as e:
            messagebox.showerror("プレビューエラー", f"プレビューの起動に失敗しました:\n{e}")

    def preview_chart_from_convert(self):
        """「変換」タブからプレビューを実行"""
        self.launch_chart_player(self.output_json_path, self.music_file_path_convert.get())

    def preview_chart_from_tab(self):
        """「プレビュー」タブからプレビューを実行"""
        self.launch_chart_player(self.json_file_path_preview.get(), self.music_file_path_preview.get())

    def reset_info(self):
        """表示されているメタデータとノーツ数をリセットする"""
        self.title_var.set("-")
        self.bpm_var.set("-")
        self.tpb_var.set("-")
        self.normal_notes_var.set("-")
        self.critical_notes_var.set("-")
        self.long_notes_var.set("-")
        self.bpm_changes_var.set("-")
        self.total_notes_var.set("-")
        self.output_json_path = ""
        self.music_file_path_convert.set("")
        self.preview_button_convert.config(state=tk.DISABLED)


def main_gui():
    """GUIアプリケーションを起動する"""
    root = tk.Tk()
    app = Application(master=root)
    app.mainloop()


if __name__ == "__main__":
    main_gui()
