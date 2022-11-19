import tkinter as tk
import MJS_PreXMLPrint as MJS_XML
import MJS_PreSinkokuPrint as MJS_PRT
import RPA_Function as RPA
import os
import datetime as dt
from tkinter import filedialog, messagebox

class GUI(tk.Frame):
    def __init__(self,root):
        self.widget_list = []
        self.csv_load = False
        self.table2_load = False
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
        self.inner_upfra.pack(side=tk.TOP, fill=tk.X, expand=True)

        # インナーサブフレーム
        self.inner_up1 = tk.Frame(
            self.inner_upfra, width=self.w, height=(self.h / 2), bd=5, relief=tk.GROOVE
        )
        self.inner_up1.pack(side=tk.TOP, padx=5, fill=tk.X, expand=True)

        self.lb = tk.Label(self.inner_up1, text="ID")
        self.lb.grid(row=0, column=0, padx=5, sticky=tk.W + tk.E)

        self.id_txt = tk.StringVar(self.inner_up1, ID)
        self.id_ent = tk.Entry(
            self.inner_up1, textvariable=self.id_txt, width=int(self.w / 30)
        )
        self.id_ent.grid(row=0, column=1, padx=5, sticky=tk.W + tk.E)
        self.widget_list.append(self.id_ent)

        self.lb2 = tk.Label(self.inner_up1, text="Pass")
        self.lb2.grid(row=1, column=0, padx=5, sticky=tk.W + tk.E)

        self.pass_txt = tk.StringVar(self.inner_up1, Pass)
        self.pass_ent = tk.Entry(
            self.inner_up1, textvariable=self.pass_txt, width=int(self.w / 30)
        )
        self.pass_ent.grid(row=1, column=1, padx=5, sticky=tk.W + tk.E)
        self.widget_list.append(self.pass_ent)

        # インナーサブフレーム
        self.inner_up2 = tk.Frame(
            self.inner_upfra, width=self.w, height=(self.h / 2), bd=5, relief=tk.GROOVE
        )
        self.inner_up2.pack(side=tk.TOP, padx=5, fill=tk.X, expand=True)

        # インナーフレーム要素
        self.bt = tk.Button(self.inner_up2,width=int(self.w/30), text="保存場所選択", command=self.Set_Dir)
        self.bt.pack(side=tk.TOP,padx=5, expand=True)
        
        self.year_lb = tk.Label(self.inner_up2,text="年")
        self.year_lb.pack()
        self.year_txt = tk.StringVar(self.inner_up2,Year_txt)
        self.year = tk.Entry(self.inner_up2,width=int(self.w/5),textvariable=self.year_txt)
        self.year.pack(side=tk.TOP, fill=tk.X, expand=True)

        self.month_lb = tk.Label(self.inner_up2,text="月")
        self.month_lb.pack()
        self.month_txt = tk.StringVar(self.inner_up2,Month_txt)        
        self.month = tk.Entry(self.inner_up2,width=int(self.w/5),textvariable=self.month_txt)
        self.month.pack(side=tk.TOP, fill=tk.X, expand=True)

        # インナーフレーム2
        self.inner_lowerfra = tk.Frame(self.fra, width=self.w, bd=5, relief=tk.GROOVE)
        self.inner_lowerfra.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)
        self.save_dir = tk.StringVar(self.inner_lowerfra,save_dir)
        self.txt = tk.Entry(self.inner_lowerfra,width=int(self.w/5),textvariable=self.save_dir)
        self.txt.pack(side=tk.TOP,padx=5,fill=tk.X, expand=True)

        self.bt2 = tk.Button(self.inner_lowerfra,width=int(self.w/30), text="XML出力(RPA)", command=self.MJS_PRT)
        self.bt2.pack(side=tk.TOP,padx=5,fill=tk.X, expand=True)

        self.bt3 = tk.Button(self.inner_lowerfra,width=int(self.w/30), text="XML→PDF変換(RPA)", command=self.MJS_XML)
        self.bt3.pack(side=tk.TOP,padx=5,fill=tk.X, expand=True)

        # # テーブル
        # self.inner_lo_left_fra = tk.Frame(self.inner_lowerfra, width=int(self.w/1.5), bd=5, relief=tk.GROOVE)
        # self.inner_lo_left_fra.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)   
        # self.table = MyTable.MyTable(self.inner_lo_left_fra, width=int(self.w/1.5))

        # # テーブル
        # self.inner_lo_right_fra = tk.Frame(self.inner_lowerfra, width=int(self.w/4), bd=5, relief=tk.GROOVE)
        # self.inner_lo_right_fra.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        # self.table2 = MyTable.MyTable(self.inner_lo_right_fra, width=int(self.w/4))

        # ウィジェット配置とキーバインド
        for w in self.widget_list:
            w.bind("<Key-Return>", self.on_key)

    # -------------------------------------------------------------------------------------------------------------------------------
    # キー入力コールバック
    def on_key(self, event):
        idx = self.widget_list.index(event.widget)
        self.widget_list[(idx + 1) % len(self.widget_list)].focus_set()

    def Set_Dir(self):
        self.save_dir.set(filedialog.askdirectory(
            title="保存場所を選択",
            initialdir=save_dir,
        ))
        for x in range(1000):
            y_txt = int(self.year_txt.get()) + x
            if str(y_txt) in self.save_dir.get():
                str_split = self.save_dir.get().split(f"/{str(y_txt)}")
                self.save_dir.set(str_split[0])
                return

            y_txt = int(self.year_txt.get()) - x                
            if str(y_txt) in self.save_dir.get():
                str_split = self.save_dir.get().split(f"/{str(y_txt)}")
                self.save_dir.set(str_split[0])                
                return

    def MJS_PRT(self):
        if tk.messagebox.askyesno("確認","NX-PROを起動し、XMLデータダウンロードを開始しますか？") is True:
            self.master.withdraw()
            tk.messagebox.showinfo(
                "注意",
                "これよりNX-PROを起動し、XMLデータダウンロードを開始します。\n処理が完了するまで[必ず]PC操作を中断して下さい。",
            )        
            MJS_PRT_Flag = MJS_PRT.Main(self.save_dir.get(),Img_dir_D, self.year_txt.get(),self.month_txt.get())
            if MJS_PRT_Flag is True:
                tk.messagebox.showinfo(
                    "完了",
                    "XMLデータダウンロードを完了しました。"
                )
                self.master.deiconify()  
            else:
                tk.messagebox.showinfo(
                    "失敗",
                    "XMLデータダウンロードを失敗しました。"
                )
                self.master.deiconify()
        else:
            tk.messagebox.showinfo(
                "中断",
                "処理を中断します。",
            )                                  

    def MJS_XML(self):
        if tk.messagebox.askyesno("確認","NX-PROを起動し、XMLデータをPDFに変換しますか？") is True:
            self.master.withdraw()
            tk.messagebox.showinfo(
                "注意",
                "これよりNX-PROを起動し、XMLデータをPDFに変換します。\n処理が完了するまで[必ず]PC操作を中断して下さい。",
            )        
            MJS_XML_Flag = MJS_XML.Main(self.save_dir.get(),Img_dir_D, self.year_txt.get(),self.month_txt.get())
            if MJS_XML_Flag is True:
                tk.messagebox.showinfo(
                    "完了",
                    "XMLデータダウンロードを完了しました。"
                )
                self.master.deiconify()  
            else:
                tk.messagebox.showinfo(
                    "失敗",
                    "XMLデータダウンロードを失敗しました。"
                )
                self.master.deiconify()
        else:
            tk.messagebox.showinfo(
                "中断",
                "処理を中断します。",
            )


        

if __name__ == "__main__":
    global ID,Pass,dir, Img_dir, Img_dir_D, save_dir,Year_txt,Month_txt  
    ID = "561"
    Pass = "051210561111111"    
    Year_txt,Month_txt = str(dt.datetime.today().year),str((dt.datetime.today().month-1))
    root = tk.Tk()
    root.withdraw()
    # RPA用画像フォルダの作成---------------------------------------------------------
    dir = RPA.My_Dir("MJS_PreFileGet")
    Img_dir = dir + r"\\img"
    Img_dir_D = Img_dir + r"/MJS_DensiSinkoku"
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
    root.title("MJS NX-PRO プレ申告処理")
    ############################################
    GuiFrame = GUI(root)
    root.mainloop()