import tkinter as tk
from tkinter import simpledialog
import threading
import time
import pyautogui
import pyperclip
from sound_utils import play_end_sound, play_stop_sound

is_running = True
is_paused = False
is_skip = False
status_text = ""

def safe_input_text(x, y, text):
    pyautogui.moveTo(x, y, duration=0.5)
    pyautogui.click()
    time.sleep(0.1)
    pyperclip.copy(str(text))
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.1)

def automation_loop(table1, table2, id_text2, repeat_count, root, status_label):
    global is_running, is_paused, is_skip, status_text

    for loop_index in range(repeat_count):
        for row in table1:
            if not is_running:
                status_text = "処理を中断しました"
                return

            while is_paused and is_running:
                time.sleep(0.2)

            x = row.get("マウスX")
            y = row.get("マウスY")
            click_flag = row.get("左クリック")
            dialog = row.get("ダイアログボックス")
            input_word = row.get("入力ワード")
            enter_flag = row.get("エンター")
            wait_time = row.get("待ち時間(秒)")

            # クリック処理
            if x is not None and y is not None and click_flag == "on":
                status_text = f"[{loop_index+1}回目] クリック: X={x}, Y={y}"
                status_label.config(text=status_text)
                root.update()
                pyautogui.moveTo(x, y, duration=0.5)
                pyautogui.click()

            # ダイアログ処理
            if dialog and "Ctrl + A" in str(dialog):
                status_text = "全選択して削除"
                status_label.config(text=status_text)
                root.update()
                pyautogui.hotkey("ctrl", "a")
                pyautogui.press("delete")

            # 入力ワード処理
            if input_word and str(input_word).strip() not in ["", "無し"]:
                input_word_str = str(input_word).strip()

                if "参照" in input_word_str:
                    ref_name = input_word_str.replace("参照", "").strip()
                    if ref_name == id_text2.replace("ID:", ""):
                        if table2 and loop_index < len(table2):
                            ref_text = table2[loop_index].get("ワード")
                            if ref_text:
                                status_text = f"[{loop_index+1}回目] 参照入力: {ref_text}"
                                status_label.config(text=status_text)
                                root.update()
                                safe_input_text(x, y, ref_text)
                else:
                    status_text = f"[{loop_index+1}回目] 入力: {input_word_str}"
                    status_label.config(text=status_text)
                    root.update()
                    safe_input_text(x, y, input_word_str)

            # エンター処理
            if enter_flag == "on":
                status_text = "エンターキーを押す"
                status_label.config(text=status_text)
                root.update()
                pyautogui.press("enter")

            # 待ち時間処理（カウントダウン表示）
            if isinstance(wait_time, (int, float)) and wait_time > 0:
                for i in range(int(wait_time), 0, -1):
                    if is_skip:  # スキップフラグが立っていたら中断
                        status_text = "待機スキップしました"
                        status_label.config(text=status_text)
                        root.update()
                        break

                    status_text = f"待機中... 残り {i} 秒"
                    status_label.config(text=status_text)
                    root.update()
                    time.sleep(1)

                # スキップボタンを押した後はフラグをリセット
                is_skip = False

    # 自然終了
    status_label.config(text="コード実行終了しました ✅")
    play_end_sound()
    root.after(1500, root.destroy)

def gui_loop(table1, table2, id_text2):
    global is_running, is_paused, is_skip, status_text
    
    root = tk.Tk()
    root.title("自動化ステータス")
    root.geometry("400x200")

    # 繰り返し回数を入力するダイアログ
    repeat_count = simpledialog.askinteger("繰り返し回数", "繰り返し処理の回数を入力してください", parent=root)
    if repeat_count is None:
        status_label = tk.Label(root, text="キャンセルされました。終了します。", font=("Arial", 12))
        status_label.pack(pady=20)
        root.after(2000, root.destroy)
        root.mainloop()
        return

    status_label = tk.Label(root, text="待機中", font=("Arial", 12))
    status_label.pack(pady=20)

    def stop():
        global is_running
        is_running = False
        status_label.config(text="処理を中断しました")
        play_stop_sound()
        root.destroy()   # 停止ボタンは即終了

    def pause():
        global is_paused
        is_paused = True
        status_label.config(text="一時停止中")

    def resume():
        global is_paused
        is_paused = False
        status_label.config(text="再開しました")

    def skip():
        global is_skip
        is_skip = True
        status_label.config(text="スキップ要求あり")

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=10)

    tk.Button(btn_frame, text="停止", command=stop, bg="red").grid(row=0, column=0, padx=5)
    tk.Button(btn_frame, text="一時停止", command=pause, bg="yellow").grid(row=0, column=1, padx=5)
    tk.Button(btn_frame, text="再開", command=resume, bg="green").grid(row=0, column=2, padx=5)
    tk.Button(btn_frame, text="スキップ", command=skip, bg="blue", fg="white").grid(row=0, column=3, padx=5)

    threading.Thread(
        target=automation_loop, args=(table1, table2, id_text2, repeat_count, root, status_label), daemon=True
    ).start()

    root.mainloop()
