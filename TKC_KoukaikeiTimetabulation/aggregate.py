import pandas as pd
import numpy as np
import os
from datetime import datetime as dt
import codecs
from chardet.universaldetector import UniversalDetector
import tkinter as tk

# ----------------------------------------------------------------------------------------------------------------------
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


# ----------------------------------------------------------------------------------------------------------------------
def Flow(Tan_path, year):
    """
    引数フォルダの担当者別フォルダ内を集計し、引数フォルダ内BACKUPフォルダに集計結果CSVを保存
    """
    try:
        for current_dir, sub_dirs, files_list in os.walk(Tan_path):
            fn = 0
            for fileobj in files_list:
                # PathとFileNameを結合
                fileurl = Tan_path + "/" + fileobj.replace("\u3000", "　")
                # PathとFileNameに当年が含まれていたら
                if year in fileurl:
                    # 列名変更してDataFrame化
                    Enc = getFileEncoding(fileurl)
                    if Enc is None:
                        with codecs.open(
                            fileurl, mode="r", encoding="shift-jis", errors="ignore"
                        ) as file:
                            H_df = pd.read_table(file, delimiter=",")
                    else:
                        with codecs.open(
                            fileurl, mode="r", encoding=Enc, errors="ignore"
                        ) as file:
                            H_df = pd.read_table(file, delimiter=",")
                    H_df = H_df.rename(
                        columns={"実\u3000績(A)": "当  月(A)", "前年実績(B)": "前年同月(B)"}
                    )
                    H_dfRow = np.array(H_df).shape[0]  # DataFrame行数取得
                    H_MdfRow = np.array(H_Mdf).shape[0]  # DataFrame行数取得
                    # CSV行ループ
                    for x in range(H_dfRow):
                        try:
                            if x >= 0 and not x == H_dfRow - 1:

                                if (
                                    not fileobj == "移動時間.CSV"
                                    and not fileobj == "移動時間A8.CSV"
                                    and not fileobj == "移動時間A9.CSV"
                                ):
                                    H_dfDataRow = H_df.loc[x]
                                    H_Tan = fileobj.split("_")
                                    H_Kanyo = H_Tan[0]
                                    H_TTan = H_Tan[1]
                                    H_TTan = (
                                        H_TTan.replace(".CSV", "")
                                        .replace("\u3000", "　")
                                        .replace("\u200b", "　")
                                        .replace(" ", "　")
                                        .replace("B2", "")
                                        .replace("A8", "")
                                        .replace("A9", "")
                                        .replace("A10", "")
                                        .replace("A11", "")
                                    )  # 空白\u3000を置換
                                    for y in range(H_MdfRow):
                                        H_MdfDataRow = H_Mdf.loc[y]
                                        T = (
                                            H_MdfDataRow["氏名"]
                                            .replace("\u3000", "　")
                                            .replace("\u200b", "　")
                                            .replace(" ", "　")
                                            .replace("B2", "")
                                            .replace("A8", "")
                                            .replace("A9", "")
                                            .replace("A10", "")
                                            .replace("A11", "")
                                        )
                                        if T == H_TTan:
                                            H_Katudou = H_dfDataRow["活動"]
                                            H_Tougetu = str(H_dfDataRow["当  月(A)"])
                                            if H_Tougetu == "nan":
                                                break

                                            CSVWriteRow = (
                                                "" + H_TTan + "",
                                                "" + H_Kanyo + "",
                                                "" + H_Katudou + "",
                                                "" + H_Tougetu + "",
                                            )
                                            H_Marges.append(CSVWriteRow)
                                            break
                                else:
                                    H_dfDataRow = H_df.loc[x]
                                    H_Kanyo = "移動時間"
                                    H_TTan = (
                                        H_dfDataRow["担当者"]
                                        .replace("\u3000", "　")
                                        .replace("\u200b", "　")
                                        .replace(" ", "　")
                                        .replace("B2", "")
                                        .replace("A8", "")
                                        .replace("A9", "")
                                        .replace("A10", "")
                                        .replace("A11", "")
                                    )  # 空白\u3000を置換
                                    for y in range(H_MdfRow):
                                        H_MdfDataRow = H_Mdf.loc[y]
                                        T = (
                                            H_MdfDataRow["氏名"]
                                            .replace("\u3000", "　")
                                            .replace("\u200b", "　")
                                            .replace(" ", "　")
                                            .replace("B2", "")
                                            .replace("A8", "")
                                            .replace("A9", "")
                                            .replace("A10", "")
                                            .replace("A11", "")
                                        )
                                        if T == H_TTan:
                                            H_Katudou = "移動時間"
                                            H_Tougetu = H_dfDataRow["当  月(A)"]
                                            if H_Tougetu == "nan":
                                                break

                                            CSVWriteRow = (
                                                "" + H_TTan + "",
                                                "" + H_Kanyo + "",
                                                "" + H_Katudou + "",
                                                "" + H_Tougetu + "",
                                            )
                                            H_Marges.append(CSVWriteRow)
                                            break
                        except:
                            pass
                fn += 1
            try:
                H_MargesRow = np.array(H_Marges).shape[0]  # 配列行数取得
                if H_MargesRow > 0:
                    df = pd.DataFrame(H_Marges)
                    df.columns = ["氏名", "関与先", "活動", "当  期(A)"]
                    pd.DataFrame(df).to_csv(
                        dir_path + "/BACKUP/集計表.csv", index=False, encoding="shift-jis"
                    )
            except:
                pass
        return
    except:
        return


# ----------------------------------------------------------------------------------------------------------------------
# if __name__ == "__main__":
def Main(year, month):
    global H_url, Enc, H_df, H_Murl, H_Mdf, tdy, H_dfRow, OKLog, NGLog, H_Marges, dir_path
    try:
        # 関与先DBをDataFrameに変換
        H_url = "//nas-sv/A_共通/A8_ｼｽﾃﾑ資料/RPA/ALLDataBase/Heidi関与先DB.csv"
        Enc = getFileEncoding(H_url)
        H_df = pd.read_csv(H_url, encoding=Enc)

        # 公会計名簿をDataFrameに変換
        H_Murl = "//nas-sv/A_共通/A8_ｼｽﾃﾑ資料/RPA/公会計時間分析/公会計名簿.csv"
        Enc = getFileEncoding(H_Murl)
        H_Mdf = pd.read_csv(H_Murl, encoding=Enc)
        # 出力用list
        H_Marges = []
        # 月指定
        month = format(int(month), "02")
        dir_path = "//nas-sv/A_共通/A8_ｼｽﾃﾑ資料/RPA/公会計時間分析/" + year + "-" + month
        Tan_path = dir_path + "/担当者別"
        # 実行
        Flow(Tan_path, year + "-")
        return True
    except:
        return False
