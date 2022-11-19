import openpyxl

# import ExcelFileAction as EFA
import datetime
import pandas as pd


class Sheet:
    """
    エクセルブック(class)
    """

    def __init__(self, XLSURL, **kw):

        self.mybook_url = XLSURL

        # Ex_file = EFA.XlsmRead(XLSURL)
        # if Ex_file[0] is True:
        # エクセルブック
        # self.book = Ex_file[1]
        self.book = openpyxl.load_workbook(self.mybook_url, keep_vba=True)
        # 全シート
        # self.input_sheet_name = self.book.sheet_names
        self.input_sheet_name = self.book.sheetnames
        # 全シート数
        self.num_sheet = len(self.input_sheet_name)

        self.sheet_header = []
        ExSheet = ""
        NameSheet = ""
        # ExSheet = self.book.parse(sheet_name, skiprows=0)
        # NameSheet = self.book.parse("関与先一覧")
        ExSheet = self.book["法人更新申請"]
        ExSheetdata = ExSheet.values
        print(ExSheetdata)
        ExSheetcolumns = next(ExSheetdata)[0:]

        NameSheet = self.book["関与先一覧"]
        NameSheetdata = NameSheet.values
        NameSheetcolumns = next(NameSheetdata)[0:]

        print(ExSheet)
        # 初回読込時の保存--------------------------
        dt_s = datetime.datetime.now()
        dt_s = dt_s.strftime("%Y-%m-%d %H-%M-%S")
        # self.sheet_df = pd.DataFrame(ExSheet)
        ExSheet = pd.DataFrame(ExSheetdata, columns=ExSheetcolumns)
        # self.name_df = pd.DataFrame(NameSheet)
        self.name_df = pd.DataFrame(NameSheetdata, columns=NameSheetcolumns)

        # 列名整理--------------------------------
        self.sheet_column_count = ExSheet.shape[1]  # 列数
        for Ex in range(self.sheet_column_count):
            ExRow = ExSheet.iloc[0]  # 列名
            ExSecondRow = ExSheet.iloc[1]  # 列名2
            if ExRow[Ex] is None:  # None判定
                # Noneの場合
                Txt = ExRow[Ex - 1]
                if ExSecondRow[Ex] is None:  # None判定
                    # Noneの場合
                    self.sheet_header.append(ExRow[Ex])
                else:
                    # Noneでない場合
                    self.sheet_header.append(Txt + "_" + ExSecondRow[Ex])
            else:
                Txt = ExRow[Ex]
                # Noneでない場合
                if ExSecondRow[Ex] is None:  # None判定
                    # Noneの場合
                    self.sheet_header.append(Txt)
                else:
                    # Noneでない場合
                    self.sheet_header.append(Txt + "_" + ExSecondRow[Ex])
        # データ整理--------------------------------
        # Df作成
        self.sheet_header = [e for e in self.sheet_header if e is not None]

        ExDf = pd.DataFrame(
            ExSheet.values[3:, : len(self.sheet_header)], columns=self.sheet_header
        )
        # Dfnan処理
        ExDf.dropna(how="all", inplace=True)
        print(ExDf)
        self.sheet_df = ExDf
        self.sheet_column_count = self.sheet_df.shape[1]  # 列数
        self.sheet_row_count = self.sheet_df.shape[0]  # 行数


if __name__ == "__main__":
    XLSURL = r"\\NAS-SV\B_監査etc\B2_電子ﾌｧｲﾙ\RPA_ミロクシステム次年更新\一括更新申請\一括更新申請ミロク(三上)TEST.xlsm"
    Ex_File = Sheet(XLSURL)
