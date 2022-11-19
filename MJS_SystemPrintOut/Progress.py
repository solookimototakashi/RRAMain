import threading
import time
from tkinter import ttk
import tkinter as tk


# 5秒タイマー
def timer():
    p.start(5)  # プログレスバー開始
    for i in range(6):
        time.sleep(1)  # 1秒待機
        b["text"] = i  # 秒数表示
    p.stop()  # プログレスバー停止


# ボタンクリック時に実行する関数
def button_clicked():
    t = threading.Thread(target=timer)  # スレッド立ち上げ
    t.start()  # スレッド開始


# root設定##################################################
root = tk.Tk()
w = 600
h = 80
x = root.winfo_screenwidth() / 2
y = root.winfo_screenheight() / 2
# 画面中央に表示。
root.geometry("%dx%d+%d+%d" % (w, h, x, y))
root.protocol("WM_DELETE_WINDOW", (lambda: "pass")())
root.resizable(0, 0)
# タイトルを非表示
# root.overrideredirect(1)
root.title("PDF加工処理")
# ##########################################################
label = ttk.Label(root, text="PDF加工中")
label.pack(padx=10, pady=20, side=tk.LEFT, anchor=tk.N)
# プログレスバー
p = ttk.Progressbar(
    root,
    length=600,
    mode="indeterminate",  # 非確定的
)
p.pack()

# ボタン
b = tk.Button(root, width=15, text="start", command=button_clicked)
b.pack()


root.mainloop()
