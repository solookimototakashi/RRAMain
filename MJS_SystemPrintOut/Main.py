import tkinter as tk
import pandas as pd
import PDFReadForList as PRF
import RPA_Function as RPA

# import os
# import MyTable
import datetime as dt

from tkinter import filedialog, messagebox, ttk


class GUI(tk.Frame):
    def __init__(self, root):
        self.widget_list = []
        self.dir_name = idir
        self.dir = RPA.My_Dir("PDFReadForList")
        self.Img_dir = dir + r"\\img"
        self.open_dir = dir + r"\\OMSOpen"
        self.Img_dir_D = Img_dir + r"\\PDFReadForList"
        self.w = int(root.winfo_screenwidth() / 2)
        self.h = int(root.winfo_screenheight() / 2)
        self.x = int(root.winfo_screenwidth() / 4)
        self.y = int(root.winfo_screenheight() / 4)
        super().__init__(root)
        # フレーム
        self.fra = tk.Frame(root, bd=5)
        self.fra.pack(fill=tk.BOTH, expand=True)

        # インナーフレーム
        self.inner_upfra = tk.Frame(
            self.fra, width=self.w, height=(self.h / 2), bd=5, relief=tk.GROOVE
        )
        self.inner_upfra.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # インナーサブフレーム
        self.inner_up1 = tk.Frame(
            self.inner_upfra, width=self.w, height=(self.h / 2), bd=5, relief=tk.GROOVE
        )
        self.inner_up1.pack(side=tk.TOP, padx=5, fill=tk.X, expand=True)
        # インナーサブフレーム要素
        # 作成先選択ボタン
        self.dir_btn = tk.Button(
            self.inner_up1, width=int(self.w / 17.5), text="作成先選択", command=self.dir_set
        )
        self.dir_btn.grid(row=0, column=0, columnspan=2, padx=5, pady=10, sticky=tk.W)
        self.widget_list.append(self.dir_btn)

        # 作成先
        self.i_lb3 = tk.Label(self.inner_up1, text="作成先")
        self.i_lb3.grid(row=1, column=0, padx=5, pady=10, sticky=tk.W)

        self.dir_txt = tk.StringVar(self.inner_up1, str(idir))
        self.dir_ent = tk.Entry(
            self.inner_up1,
            textvariable=self.dir_txt,
            width=int(self.w / 20),
        )
        self.dir_ent.grid(row=1, column=1, padx=5, pady=10, sticky=tk.W)
        self.widget_list.append(self.dir_ent)

        # 開始ボタン
        self.start_btn = tk.Button(
            self.inner_up1,
            width=int(self.w / 30),
            text="PDF抽出開始",
            command=self.start,
        )
        self.start_btn.grid(
            row=2, column=0, columnspan=2, padx=5, pady=10, sticky=tk.NSEW
        )
        self.widget_list.append(self.start_btn)

        # ウィジェット配置とキーバインド
        for w in self.widget_list:
            w.bind("<Key-Return>", self.on_key)

    # -------------------------------------------------------------------------------------------------------------------------------
    # キー入力コールバック
    def on_key(self, event):
        idx = self.widget_list.index(event.widget)
        self.widget_list[(idx + 1) % len(self.widget_list)].focus_set()

    def start(self):
        if tk.messagebox.askyesno("確認", "PDF抽出を開始しますか？") is True:
            self.master.withdraw()
            tk.messagebox.showinfo(
                "注意",
                "PDF抽出を開始します。\n長時間かかるのでターミナルを確認してください。\nGILがめんどいのでプログレス実装してないです。",
            )
            PRF_Flag = PRF.Main(self)
            if PRF_Flag is True:
                tk.messagebox.showinfo("完了", "PDF抽出を完了しました。")
                self.master.deiconify()
            else:
                tk.messagebox.showinfo("失敗", "PDF抽出が失敗しました。")
                self.master.deiconify()
        else:
            tk.messagebox.showinfo(
                "中断",
                "処理を中断します。",
            )

    def dir_set(self):
        self.dir_name = filedialog.askdirectory(
            title="PDF抽出先を指定",
            initialdir=idir,
        )
        self.dir_txt.set(self.dir_name)


if __name__ == "__main__":
    global ID, Pass, dir, idir, Img_dir, Img_dir_D, save_dir, Lyear, Lmon

    ID = "561"
    Pass = "051210561111111"
    # 前月指定
    Lyear = dt.datetime.today().year
    Lmon = dt.datetime.today().month - 1

    idir = r"\\NAS-SV\B_監査etc\B2_電子ﾌｧｲﾙ\ﾒｯｾｰｼﾞﾎﾞｯｸｽ"
    root = tk.Tk()
    root.withdraw()
    # RPA用画像フォルダの作成---------------------------------------------------------
    dir = RPA.My_Dir("PDFReadForList")
    Img_dir = dir + r"\\img"
    open_dir = dir + r"\\OMSOpen"
    Img_dir_D = Img_dir + r"\\PDFReadForList"
    # ------------------------------------------------------------------------------
    # ルート設定#################################
    root = tk.Tk()
    w = int(root.winfo_screenwidth() / 1.5)
    h = int(root.winfo_screenheight() / 4)
    x = int(root.winfo_screenwidth() / 6)
    y = int(root.winfo_screenheight() / 4)
    # 画面中央に表示。
    root.geometry("%dx%d+%d+%d" % (w, h, x, y))
    root.title("TKC 公会計時間集計")
    ############################################
    GuiFrame = GUI(root)
    root.mainloop()
