import tkinter as tk
import pandas as pd
import TKC_SeikyuuNyuuryoku as TKC_Sei
import RPA_Function as RPA
import os
import MyTable
import datetime as dt
import calendar
from tkinter import filedialog, messagebox, ttk
import WarekiHenkan as WH


class GUI(tk.Frame):
    def __init__(self, root):
        self.widget_list = []
        self.sheet_names = []
        self.dir = RPA.My_Dir("MJS_PreFileGet")
        self.Img_dir = dir + r"\\img"
        self.open_dir = dir + r"\\OMSOpen"
        self.Img_dir_D = Img_dir + r"\\TKC_SeikyuuNyuuryoku"
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

        self.lb = tk.Label(self.inner_up1, text="ID")
        self.lb.grid(row=0, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.id_txt = tk.StringVar(self.inner_up1, ID)
        self.id_ent = tk.Entry(
            self.inner_up1, textvariable=self.id_txt, width=int(self.w / 30)
        )
        self.id_ent.grid(row=0, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.id_ent)

        self.lb2 = tk.Label(self.inner_up1, text="Pass")
        self.lb2.grid(row=1, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.pass_txt = tk.StringVar(self.inner_up1, Pass)
        self.pass_ent = tk.Entry(
            self.inner_up1, textvariable=self.pass_txt, width=int(self.w / 30)
        )
        self.pass_ent.grid(row=1, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.pass_ent)

        # インナーサブフレーム
        self.inner_up2 = tk.Frame(
            self.inner_upfra, width=self.w, height=(self.h / 2), bd=5, relief=tk.GROOVE
        )
        self.inner_up2.pack(side=tk.TOP, padx=5, fill=tk.X, expand=True)

        # インナーサブフレーム要素
        self.bt = tk.Button(
            self.inner_up2,
            width=int(self.w / 30),
            text="給与計算請求額表選択",
            command=self.xlsx_get,
        )
        self.bt.pack(side=tk.LEFT, padx=5, pady=10, fill=tk.X, expand=True)
        self.widget_list.append(self.bt)
        # self.bt.grid(row=0,column=0, padx=5, sticky=tk.W + tk.E)

        self.sheet_list_var = tk.StringVar(self.inner_up2, "")
        self.sheet_list = ttk.Combobox(
            self.inner_up2,
            values=self.sheet_names,
            textvariable=self.sheet_list_var,
            state="readonly",
        )
        self.sheet_list.pack(side=tk.LEFT, padx=5, pady=10, fill=tk.X, expand=True)
        self.widget_list.append(self.sheet_list)
        self.sheet_list.bind("<<ComboboxSelected>>", self.change_sheet)

        # インナーフレーム2
        self.inner_lowerfra = tk.Frame(self.fra, width=self.w, bd=5, relief=tk.GROOVE)
        self.inner_lowerfra.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

        # テーブル
        self.inner_lo_left_fra = tk.Frame(
            self.inner_lowerfra, width=int(self.w / 1.5), bd=5, relief=tk.GROOVE
        )
        self.inner_lo_left_fra.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.table = MyTable.MyTable(self.inner_lo_left_fra, width=int(self.w / 1.5))
        self.table_side()

    def table_side(self):
        # テーブル
        self.inner_lo_right_fra = tk.Frame(
            self.inner_lowerfra, width=int(self.w / 4), bd=5, relief=tk.GROOVE
        )
        self.inner_lo_right_fra.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.i_lb = tk.Label(self.inner_lo_right_fra, text="請求日")
        self.i_lb.grid(
            row=0, column=0, columnspan=6, padx=5, pady=10, sticky=tk.W + tk.E
        )
        # 年
        self.i_lb = tk.Label(self.inner_lo_right_fra, text="年")
        self.i_lb.grid(row=1, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt = tk.StringVar(self.inner_lo_right_fra, str(Lyear.year))
        self.i_id_ent = tk.Entry(
            self.inner_lo_right_fra, textvariable=self.i_id_txt, width=int(self.w / 100)
        )
        self.i_id_ent.grid(row=1, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_id_ent)
        # 月
        self.i_lb2 = tk.Label(self.inner_lo_right_fra, text="月")
        self.i_lb2.grid(row=1, column=2, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt2 = tk.StringVar(self.inner_lo_right_fra, str(Lmon))
        self.i_id_ent2 = tk.Entry(
            self.inner_lo_right_fra,
            textvariable=self.i_id_txt2,
            width=int(self.w / 100),
        )
        self.i_id_ent2.grid(row=1, column=3, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_id_ent2)
        # 日
        self.i_lb3 = tk.Label(self.inner_lo_right_fra, text="日")
        self.i_lb3.grid(row=1, column=4, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt3 = tk.StringVar(self.inner_lo_right_fra, str(Lday))
        self.i_id_ent3 = tk.Entry(
            self.inner_lo_right_fra,
            textvariable=self.i_id_txt3,
            width=int(self.w / 100),
        )
        self.i_id_ent3.grid(row=1, column=5, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_id_ent3)
        # 関与先コード列番号
        self.i_lb4 = tk.Label(self.inner_lo_right_fra, text="関与先コード列番号")
        self.i_lb4.grid(row=2, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt4 = tk.StringVar(self.inner_lo_right_fra, "")
        self.i_id_ent4 = tk.Entry(
            self.inner_lo_right_fra,
            textvariable=self.i_id_txt4,
            width=int(self.w / 100),
        )
        self.i_id_ent4.grid(row=2, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_id_ent4)

        self.i_id_bt4 = tk.Button(
            self.inner_lo_right_fra,
            width=int(self.w / 40),
            text="選択列転記",
            command=self.i_id_bt4func,
        )
        self.i_id_bt4.grid(
            row=2, column=2, columnspan=5, padx=5, pady=10, sticky=tk.W + tk.E
        )
        self.widget_list.append(self.i_id_bt4)
        # 関与先名列番号
        self.i_name = tk.Label(self.inner_lo_right_fra, text="関与先名列番号")
        self.i_name.grid(row=3, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_name_txt = tk.StringVar(self.inner_lo_right_fra, "")
        self.i_name_ent = tk.Entry(
            self.inner_lo_right_fra,
            textvariable=self.i_name_txt,
            width=int(self.w / 100),
        )
        self.i_name_ent.grid(row=3, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_name_ent)

        self.i_name_bt = tk.Button(
            self.inner_lo_right_fra,
            width=int(self.w / 40),
            text="選択列転記",
            command=self.i_name_btfunc,
        )
        self.i_name_bt.grid(
            row=3, column=2, columnspan=5, padx=5, pady=10, sticky=tk.W + tk.E
        )
        self.widget_list.append(self.i_name_bt)
        # 入力金額列番号
        self.i_lb5 = tk.Label(self.inner_lo_right_fra, text="入力金額列番号")
        self.i_lb5.grid(row=4, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt5 = tk.StringVar(self.inner_lo_right_fra, "")
        self.i_id_ent5 = tk.Entry(
            self.inner_lo_right_fra, textvariable=self.i_id_txt5, width=int(self.w / 40)
        )
        self.i_id_ent5.grid(row=4, column=1, padx=5, pady=10, sticky=tk.W + tk.E)
        self.widget_list.append(self.i_id_ent5)

        self.i_id_bt5 = tk.Button(
            self.inner_lo_right_fra,
            width=int(self.w / 40),
            text="選択列転記",
            command=self.i_id_bt5func,
        )
        self.i_id_bt5.grid(
            row=4, column=2, columnspan=5, padx=5, pady=10, sticky=tk.W + tk.E
        )
        self.widget_list.append(self.i_id_bt5)
        # 科目
        self.code = tk.Label(self.inner_lo_right_fra, text="科目コード")
        self.code.grid(row=5, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.code_txt = tk.StringVar(self.inner_lo_right_fra, "")
        self.code_ent = tk.Entry(
            self.inner_lo_right_fra, textvariable=self.code_txt, width=int(self.w / 40)
        )
        self.code_ent.grid(
            row=5, column=1, columnspan=5, padx=5, pady=10, sticky=tk.W + tk.E
        )
        self.widget_list.append(self.code_ent)

        # 摘要
        self.i_lb6 = tk.Label(self.inner_lo_right_fra, text="摘要")
        self.i_lb6.grid(row=6, column=0, padx=5, pady=10, sticky=tk.W + tk.E)

        self.i_id_txt6 = tk.StringVar(self.inner_lo_right_fra, "")
        self.i_id_ent6 = tk.Entry(
            self.inner_lo_right_fra, textvariable=self.i_id_txt6, width=int(self.w / 40)
        )
        self.i_id_ent6.grid(
            row=6, column=1, columnspan=5, padx=5, pady=10, sticky=tk.W + tk.E
        )
        self.widget_list.append(self.i_id_ent6)

        # 開始
        self.start_btn = tk.Button(
            self.inner_lo_right_fra,
            width=int(self.w / 40),
            text="請求入力開始(RPA)",
            command=self.start,
        )
        self.start_btn.grid(
            row=7, column=0, columnspan=6, padx=5, pady=10, sticky=tk.W + tk.E
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
        # 年
        if self.i_id_txt.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "請求年が未入力です。",
            )
            return
        # 月
        if self.i_id_txt2.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "請求月が未入力です。",
            )
            return
        # 日
        if self.i_id_txt3.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "請求日が未入力です。",
            )
            return
        # 関与先コード列番号
        if self.i_id_txt4.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "関与先コード列番号が未入力です。",
            )
            return
        # 入力金額列番号
        if self.i_id_txt5.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "入力金額列番号が未入力です。",
            )
            return
        # 科目
        if self.code_txt.get() == "":
            tk.messagebox.showinfo(
                "確認",
                "科目コードが未入力です。",
            )
            return
        if tk.messagebox.askyesno("確認", "TKC-FMSを起動し、請求入力を開始しますか？") is True:
            self.master.withdraw()
            tk.messagebox.showinfo(
                "注意",
                "これよりTKC-FMSを起動し、請求入力を開始します。\n処理が完了するまで[必ず]PC操作を中断して下さい。",
            )
            TKC_Sei_Flag = TKC_Sei.Main(self)
            if TKC_Sei_Flag is True:
                tk.messagebox.showinfo("完了", "請求入力を完了しました。")
                self.master.deiconify()
            else:
                tk.messagebox.showinfo("失敗", "請求入力が失敗しました。")
                self.master.deiconify()
        else:
            tk.messagebox.showinfo(
                "中断",
                "処理を中断します。",
            )

    def i_id_bt4func(self):
        self.i_id_txt4.set(self.table.startcol)

    def i_id_bt5func(self):
        self.i_id_txt5.set(self.table.startcol)

    def i_name_btfunc(self):
        self.i_name_txt.set(self.table.startcol)

    def xlsx_get(self):
        typ = [("Excel files", ".xlsx .xls .xlsm")]
        self.xlsx_name = filedialog.askopenfilename(
            filetypes=typ,
            title="Excelファイルを開く",
            initialdir=Img_dir_D,
        )
        self.excel_data = pd.ExcelFile(self.xlsx_name)
        self.sheet_names = self.excel_data.sheet_names
        self.sheet_list.config(values=self.sheet_names)

    def change_sheet(self, event):
        sheet_df = self.excel_data.parse(self.sheet_list.get())
        self.table.model.df = sheet_df
        self.table.show()


if __name__ == "__main__":
    global ID, Pass, dir, Img_dir, Img_dir_D, save_dir, Lyear, Lmon, Lday
    Lyear = WH.Wareki.from_ad(dt.datetime.today().year)
    ID = "561"
    Pass = "051210561111111"
    # 月中以降は当月末日
    if dt.datetime.today().day > 15:
        Lmon = dt.datetime.today().month
        Lday = calendar.monthrange(dt.datetime.today().year, dt.datetime.today().month)[
            1
        ]
    else:
        Lmon = dt.datetime.today().month - 1
        Lday = calendar.monthrange(
            dt.datetime.today().year, dt.datetime.today().month - 1
        )[1]
    idir = r"\\nas-sv\K_管理\A1_総務\01_総務"
    root = tk.Tk()
    root.withdraw()
    # RPA用画像フォルダの作成---------------------------------------------------------
    dir = RPA.My_Dir("TKC_SeikyuuNyuuryoku")
    Img_dir = dir + r"\\img"
    open_dir = dir + r"\\OMSOpen"
    Img_dir_D = Img_dir + r"\\TKC_SeikyuuNyuuryoku"
    save_dir = r"\\nas-sv\B_監査etc\B2_電子ﾌｧｲﾙ\ﾒｯｾｰｼﾞﾎﾞｯｸｽ"
    # ------------------------------------------------------------------------------
    # ルート設定#################################
    root = tk.Tk()
    w = int(root.winfo_screenwidth() / 1.5)
    h = int(root.winfo_screenheight() / 1.5)
    x = int(root.winfo_screenwidth() / 6)
    y = int(root.winfo_screenheight() / 6)
    # 画面中央に表示。
    root.geometry("%dx%d+%d+%d" % (w, h, x, y))
    root.title("TKC FMS 請求入力")
    ############################################
    GuiFrame = GUI(root)
    root.mainloop()
