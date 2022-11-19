import pandas as pd
import numpy as np
from chardet.universaldetector import UniversalDetector

# -------------------------------------------------------------------------------------------------------------------------------
def getFileEncoding(file_path):
    """
    引数ファイルのエンコードを調べる
    """
    detector = UniversalDetector()
    with open(file_path, mode="rb") as f:
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result["encoding"]


# ------------------------------------------------------------------------------------------------------------------------------
def CsvPlus(URL, LogList, LogMSG):
    """
    引数指定のCSV最終行に行データ追加
    """
    try:
        SerchEnc = format(getFileEncoding(URL))
        df_shape = LogList.shape
        # 最終行に追加
        LogList.loc[df_shape[0]] = LogMSG
        pd.DataFrame(LogList).to_csv(URL, encoding=SerchEnc, index=False)
        return True
    except:
        return False


# -------------------------------------------------------------------------------------------------------------------------------
def XlsmRead(URL):
    """
    引数指定のCSVを読みとる
    """
    try:
        URL = URL.replace("\\", "/")
        C_xlsm = pd.ExcelFile(URL, engine="openpyxl")
        return True, C_xlsm
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def to_Df(Exf, SName):
    """
    エクセルシートをデータフレームへ
    """
    try:
        # DataFrameとしてsheet1枚のデータ(2019)を読込み
        C_xlsm = Exf.parse(sheet_name=SName, index_col=None)
        return True, C_xlsm
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def to_DfList(URL):
    """
    エクセルシートを順々にデータフレーム化しリスト格納
    """
    try:
        df_list = []
        file = pd.ExcelFile(URL)  # bookを読む
        for sheet in file.sheet_names:
            df_list.append(file.parse(sheet))  # シートを順々にデータフレーム化
        return True, df_list
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSortArray(URL, KeyCol, Key, Arg):
    """
    引数指定の条件一致行を対象から抽出
    """
    try:
        Sort_url = URL
        SerchEnc = format(getFileEncoding(Sort_url))
        C_Child = pd.read_csv(Sort_url, encoding=SerchEnc)
        C_CforCount = 0
        C_dfRow = np.array(C_Child).shape[0]  # 配列行数取得
        for x in range(C_dfRow):
            C_ChildDataRow = C_Child.iloc[x, :]
            C_Val = C_ChildDataRow[KeyCol]
            if Arg == "str":
                C_Val = str(C_Val)
            if Arg == "int":
                C_Val = int(C_Val)
            if Key == C_Val:
                return True, C_ChildDataRow
            else:
                C_CforCount = C_CforCount + 1
        return False, ""
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
