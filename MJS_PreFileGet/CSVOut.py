# ----------------------------------------------------------------------------------------------------------------------
# モジュールインポート
# pandasインポート
import pandas as pd

# 配列計算関数numpyインポート
import numpy as np
from chardet.universaldetector import UniversalDetector

# -------------------------------------------------------------------------------------------------------------------------------
def getFileEncoding(file_path):  # .format( getFileEncoding( "sjis.csv" ) )
    detector = UniversalDetector()
    with open(file_path, mode="rb") as f:
        for binary in f:
            detector.feed(binary)
            if detector.done:
                break
    detector.close()
    return detector.result["encoding"]


# ------------------------------------------------------------------------------------------------------------------------------
def CsvPlus(URL, LogList, LogMSG):  # 引数指定のCSV最終行に行データ追加
    try:
        df_shape = LogList.shape
        # 最終行に追加
        LogList.loc[df_shape[0]] = LogMSG
        SerchEnc = format(getFileEncoding(URL))
        print(LogList)
        LogList.to_csv(URL, encoding=SerchEnc, index=False)
        return True
    except:
        return False


# -------------------------------------------------------------------------------------------------------------------------------
def CsvRead(URL):  # 引数指定のCSVを読みとる
    try:
        SerchEnc = format(getFileEncoding(URL))
        if SerchEnc == "None":
            C_csv = pd.read_csv(URL, encoding="cp932")
        else:
            C_csv = pd.read_csv(URL, encoding=SerchEnc)
        return True, C_csv
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvReadDtypeDict(URL, dict):  # 引数指定のCSVを読みとる
    try:
        SerchEnc = format(getFileEncoding(URL))
        if SerchEnc == "None":
            C_csv = pd.read_csv(URL, encoding="cp932", dtype=dict)
        else:
            C_csv = pd.read_csv(URL, encoding=SerchEnc, dtype=dict)
        return True, C_csv
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvReadHeaderless(URL):  # 引数指定のCSVを読みとる
    try:
        SerchEnc = format(getFileEncoding(URL))
        C_csv = pd.read_csv(URL, encoding=SerchEnc, header=None)
        return True, C_csv
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSave(URL, data):
    # DataFrame作成
    df = pd.DataFrame(data)
    df.to_csv(URL, index=False)
    return df


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSaveNoHeader(URL, data, enc):
    # DataFrame作成
    df = pd.DataFrame(data)
    df.to_csv(URL, index=False, header=False, encoding=enc)
    return df


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSaveEnc(URL, data, enc, HLIST):
    # DataFrame作成
    df = pd.DataFrame(data)
    df.columns = HLIST
    with open(URL, mode="w", encoding=enc, errors="ignore", newline="") as f:
        # pandasでファイルオブジェクトに書き込む
        df.to_csv(f, index=False)
    return df


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSortDatetime(URL, ColName, asc):  # asc=False降順
    try:
        SerchEnc = format(getFileEncoding(URL))
        data = pd.read_csv(URL, encoding=SerchEnc)
        data[ColName] = pd.to_datetime(data[ColName], infer_datetime_format=True)
        data.sort_values(by=ColName, ascending=False, inplace=True)
        return True, data
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSortArray(
    URL, KeyCol, Key, Arg
):  # dtype={"TKCKokuzeiUserCode": str,"TKCKokuzeiUserCode": str}.........　#引数指定の条件一致行を対象から抽出
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
def CsvSortRow(
    URL, KeyCol, Key, Arg
):  # dtype={"TKCKokuzeiUserCode": str,"TKCKokuzeiUserCode": str}.........　#引数指定の条件一致行を対象から抽出
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
                return True, C_CforCount
            else:
                C_CforCount = C_CforCount + 1
        return False, ""
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def CsvSortRowDouble(
    URL, KeyCol1, KeyCol2, Key1, Key2
):  # dtype={"TKCKokuzeiUserCode": str,"TKCKokuzeiUserCode": str}.........　#引数指定の条件一致行を対象から抽出
    try:
        Sort_url = URL
        SerchEnc = format(getFileEncoding(Sort_url))
        C_Child = pd.read_csv(Sort_url, encoding=SerchEnc)
        C_CforCount = 0
        C_dfRow = np.array(C_Child).shape[0]  # 配列行数取得
        for x in range(C_dfRow):
            C_ChildDataRow = C_Child.iloc[x, :]
            if C_ChildDataRow[KeyCol2] == "  ":  # TKCFMSで枝番なしは半角空白2つなので
                C_Val = str(C_ChildDataRow[KeyCol1])
            else:
                C_Val = str(C_ChildDataRow[KeyCol1]) + str(C_ChildDataRow[KeyCol2])
            if not Key2:
                Key = str(Key1)
            else:
                Key = str(Key1) + str(Key2)

            if Key == C_Val:
                return True, C_CforCount
            else:
                C_CforCount = C_CforCount + 1
        return False, ""
    except:
        return False, ""


# -------------------------------------------------------------------------------------------------------------------------------
def chkprint(*args):  # 変数名をそのままprint関数内で表示させる関数
    for obj in args:
        for k, v in globals().items():
            if id(v) == id(obj):
                target = k
                break
    return target


# -------------------------------------------------------------------------------------------------------------------------------
def typeInfo(targetData):  # データがどのデータ型か、列数、行数を表示する関数
    if type(targetData) is pd.core.frame.DataFrame:
        print("{} は DataFrame型".format(chkprint(targetData)))
        print(
            "{} の行数, 列数・・・{}\n".format(chkprint(targetData), targetData.shape)
        )  # shapeの表示内容は、(行数, 列数)となる
        return "DataFrame"
    if type(targetData) is list:
        print("{} は list型".format(chkprint(targetData)))
        print(
            "{} の行数, 列数・・・{}\n".format(
                chkprint(targetData), pd.DataFrame(targetData).shape
            )
        )  # shapeの表示内容は、(行数, 列数)となる
        return "list"
    if type(targetData) is np.ndarray:
        print("{} は ndarray型".format(chkprint(targetData)))
        print(
            "{} の行数, 列数・・・{}\n".format(chkprint(targetData), targetData.shape)
        )  # shapeの表示内容は、(行数, 列数)となる
        return "ndarray"
    if type(targetData) is pd.core.series.Series:
        print("{} は Series型".format(chkprint(targetData)))
        print(
            "{} の行数, 列数・・・{}\n".format(chkprint(targetData), targetData.shape)
        )  # shapeの表示内容は、(行数, 列数)となる
        return "Series"
        # 使用例↓
        # data_list = [[1.0, 2.0, 3.0], [11.0, 12.0, 13.0]]
        # data_ndarray = np.array(data_list)      #listをndarrayに変換する
        # data_Series = pd.Series(data_list)      #listをSeriesに変換する
        # data_df = pd.DataFrame(data_list)       #listをDataFrameに変換する
        # typeInfo(data_list)
        # typeInfo(data_ndarray)
        # typeInfo(data_Series)
        # typeInfo(data_df)


# -------------------------------------------------------------------------------------------------------------------------------
