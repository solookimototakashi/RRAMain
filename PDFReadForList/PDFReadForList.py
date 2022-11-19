from datetime import datetime
import os
import csv
import numpy as np
from pdfminer.high_level import extract_text
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from PIL import Image
from collections import OrderedDict
import ContextTimeOut as CTO
import CSVOut as FCSV
import FolderCreate as FC
import MiniStrChange as MSC
import toml
import CSVSetting as CSVSet  # CSVの設定ファイルの読込
import PDFCellsImport as FPDF
import tkinter as tk
from tkinter import ttk
import threading

# logger設定------------------------------------------------------------------------------
import logging.config

logging.config.fileConfig(r"LogConf\logging_debug.conf")
logger = logging.getLogger(__name__)
# ----------------------------------------------------------------------------------------
def SerchdirFolders(URL):
    """
    概要: 指定URL配下のサブフォルダを取得
    @param URL: フォルダURL(str)
    @return サブフォルダURLリスト(list)
    """
    List = []
    for fd_path, sb_folder, sb_file in os.walk(URL):
        for fol in sb_folder:
            List.append([fd_path, fol])
    return List


# ----------------------------------------------------------------------------------------
def SerchdirFiles(URL):
    """
    概要: 指定URL配下のファイルを取得
    @param URL: フォルダURL(str)
    @return ファイルURLリスト(list)
    """
    List = []
    for fd_path, sb_folder, sb_file in os.walk(URL):
        for fil in sb_file:
            List.append([fd_path, fil])
    return List


# -------------------------------------------------------------------------------------------------------
def DiffListPlus(ColList, ScrList, Ers):
    """
    概要: 取得した情報を各リストに格納
    @param ColList: PDFから取得したヘッダー情報(list)
    @param ScrList: PDFから取得した情報(list)
    @param Ers: サブテーブルから取得した場合はSubと指定(str)
    @return bool,代入したリスト名(文字列)
    """
    try:
        LNList = Settingtoml["MASTER"]["ListNameList"]  # tomlから各種設定リスト名を抽出
        NewColList = []  # tomlリストの代入変数を初期化
        for LNListItem in LNList:  # tomlリスト数分ループ
            NewColList = Settingtoml["CsvSaveEnc"][LNListItem]  # tomlリストの代入
            SColA = set(ColList) - set(NewColList)  # tomlリストと抽出ヘッダーの比較A
            SColB = set(NewColList) - set(ColList)  # tomlリストと抽出ヘッダーの比較B
            if len(SColA) == 0 and len(SColB) == 0:  # tomlリストと抽出ヘッダーが完全一致したら
                CDict[LNListItem].append(ScrList)  # tomlリストに格納

                # ログ追記------------------------------------------------------------------------------
                with open(
                    MeUrl + r"/NoErrLog.CSV",
                    "a",
                    encoding="utf-8",
                ) as Colup:
                    writer = csv.writer(Colup)
                    writer.writerow(ColList)
                    writer.writerow(ScrList)
                # ----------------------------------------------------------------------------------------

                return True, LNListItem
        print(
            "#########################################################################################"
        )
        print(ColList)
        print(
            "========================================================================================"
        )
        print(ScrList)
        # 全てのtomlリストと不一致の場合-----------------------------------------------------------------
        if Ers == "Sub":  # サブテーブル判定引数が指定されていたら
            print("サブテーブル取得エラー")
            if len(ScrList) < 20:
                LC = 20 - len(ScrList)
                for LCC in range(LC):
                    ScrList.append("")
            CDict["SubErrList"].append(ScrList)

            # Errログ追記------------------------------------------------------------------------------
            with open(
                MeUrl + r"/ErrLog.CSV",
                "a",
                encoding="utf-8",
            ) as Colup:
                writer = csv.writer(Colup)
                writer.writerow(ColList)
                writer.writerow(ScrList)
            # ----------------------------------------------------------------------------------------

            return False, "SubErrList"
        else:
            # サブテーブル判定引数が指定されていなければtomlリスト未設定のPDF形式
            print("指定列名での設定項目がありませんでした。")

            # 列名追記------------------------------------------------------------------------------
            with open(
                MeUrl + r"/ColumnList.CSV",
                "a",
                encoding="utf-8",
            ) as Colup:
                writer = csv.writer(Colup)
                writer.writerow(ColList)
                writer.writerow(ScrList)
            # ----------------------------------------------------------------------------------------

            if len(ScrList) < 20:
                LC = 20 - len(ScrList)
                for LCC in range(LC):
                    ScrList.append("")
            CDict["ErrList"].append(ScrList)
            return False, "ErrList"
    except:
        # エラー処理-----------------------------------------------------------------------------------
        print(
            "#########################################################################################"
        )
        print(ColList)
        print(
            "========================================================================================"
        )
        print(ScrList)
        if Ers == "Sub":
            print("サブテーブル取得エラー")
            if len(ScrList) < 20:
                LC = 20 - len(ScrList)
                for LCC in range(LC):
                    ScrList.append("")
            CDict["SubErrList"].append(ScrList)

            # Errログ追記------------------------------------------------------------------------------
            with open(
                MeUrl + r"/RPAPhoto/PDFReadForList/ErrLog.CSV",
                "a",
                encoding="utf-8",
            ) as Colup:
                writer = csv.writer(Colup)
                writer.writerow(ColList)
                writer.writerow(ScrList)
            # ----------------------------------------------------------------------------------------

            return False, "SubErrList"
        else:
            print("指定列名での設定項目がありませんでした。")

            # 列名追記------------------------------------------------------------------------------
            with open(
                MeUrl + r"/RPAPhoto/PDFReadForList/ColumnList.CSV",
                "a",
                encoding="utf-8",
            ) as Colup:
                writer = csv.writer(Colup)
                writer.writerow(ColList)
                writer.writerow(ScrList)
            # ----------------------------------------------------------------------------------------

            if len(ScrList) < 20:
                LC = 20 - len(ScrList)
                for LCC in range(LC):
                    ScrList.append("")
            CDict["ErrList"].append(ScrList)
            return False, "ErrList"


# -------------------------------------------------------------------------------------------------------
def DiffListCSVOUT(ListURL, ColN):
    """
    概要: PDFデータを格納したtomlリストのCSV書出し
    (外部ファンクションを引用)
    @param ListURL: CSV格納フォルダ(str)
    @param ColN: tomlリスト名=CSVファイル名(str)
    @return なし
    """
    ListURL = ListURL.replace("\\", "/")
    FCSV.CsvSaveEnc(
        ListURL + "/" + ColN + ".csv",
        CDict[ColN],
        "cp932",
        Settingtoml["CsvSaveEnc"][ColN],
    )


# ----------------------------------------------------------------------------------------------
def Camelotsp(Sbtext, path_pdf, PageVol):
    """
    概要: 自作ファンクションCamelotSerchで取得したTKCの完了報告書PDF情報から、
        次ページに取得対象があるかを判定
    @param Sbtext: PDFテキスト(str)
    @param path_pdf: PDFパス(str)
    @param PageVol: PDFページ(str)
    @return bool,判定文,次ページカウント(int),次ページリスト(list)
    """
    try:
        FA = 0  # 次ページカウントを初期化
        FAList = []  # 次ページリストを初期化
        NextFlag = True  # 次ページ判定を初期化
        # 次ページ判定がTrueで次ページカウントが5ページ以下ならループ-----------------------
        while NextFlag is True and FA <= 5:
            CSbtext = Sbtext.split("\n")  # PDFテキストをリスト化
            SBR = len(CSbtext) - 1  # PDFテキストリストのデータ数
            # PDFテキストリストを逆方向でループ------------------------------------------
            for CSbtextItem in reversed(CSbtext):
                if CSbtextItem == "":  # PDFテキスト行にテキストが無ければ
                    CSbtext.pop(SBR)  # PDFテキストリストから行削除
                # PDFテキスト行にテキストが無ければ指定文字列が含まれている場合------------
                elif "国税受付システムからの「受信通知」の内容" in CSbtextItem:
                    # 指定文字で分割しリスト化後、先頭行を削除
                    Relist = CSbtextItem.split("国税受付システムからの「受信通知」の内容")
                    CSbtext[SBR] = CSbtext[SBR].replace(Relist[0], "")
                SBR -= 1
            # 指定文字が処理後のPDFテキストに含まれていなければ--------------------------
            if "通知」の内容" not in Sbtext:
                if "受 付 通 知" in Sbtext:
                    return True, Sbtext, FA, FAList
                else:
                    print("次ページなし")
                    return False, "通知内容無し", FA, FAList
            else:
                SBR = len(CSbtext) - 1  # 処理後PDFテキストリストのデータ数-1
                # 処理後PDFテキストリストの下から10行分ループ---------------------------
                for SBRP in range(SBR - 10, SBR):
                    # 下から10行目以内に指定文字列がある場合----------------------------
                    if "ファイル名称" in CSbtext[SBRP]:
                        print("1ページ")
                        NextFlag = False  # 次ページ取得不要
                        break  # ループを抜ける
                    else:
                        print("続き有")
                        NextFlag = True  # 次ページ取得必要
                # 次ページ取得が必要な場合---------------------------------------------
                if NextFlag is True:
                    yy = int(PageVol) + FA  # PDFページ引数に次ページカウントを加算
                    mp = []  # PDFminerのページ引数がリスト型なのでリストを初期化
                    mp.append(yy)  # PDFminerのページ引数作成
                    # PDFminerで指定ページのテキスト取得-------------------------------
                    NSbtext = extract_text(
                        path_pdf, page_numbers=mp, maxpages=1, codec="utf-8"
                    )
                    # --------------------------------------------------------------
                    NSbtext = NSbtext.replace("\n\n", ":")  # 二重改行文字列があれば置換
                    # 重複読込判定---------------------------------------------------
                    if NSbtext in Sbtext or "通知」の内容" not in NSbtext:
                        print("重複読込")
                        break
                    else:
                        # 重複読込でなければテキスト結合して次ページカウントに加算
                        Sbtext = Sbtext + "\n" + NSbtext
                        print(Sbtext)
                        FAList.append(FA)
                        FA += 1
        return True, Sbtext, FA, FAList
    except:
        return False, "Camelotspエラー", FA, FAList


# ----------------------------------------------------------------------------------------------
def CamelotSerch(CDict, path_pdf, PageVol, Settingtoml, SCode, y, engine, DLCList):
    """
    概要: 外部ファンクションのタイムアウト設定がされたCamelotPDFテキスト取得から、内容に応じて処理を分岐
    @param CDict : 外部CSVSetting.pyから取得したリスト(py)
    @param path_pdf : PDFパス(str)
    @param PageVol : PDFパス(str)
    @param Settingtoml : tomlファイル(toml)
    @param SCode : 関与先コード(int)
    @param y :  ループ変数(int)
    @param engine :  camelotのエンジン(str)
    @param DLCList : 格納完了したリスト名一覧(list)
    @return : bool,TimeOut判定変数(bool),TableList,セル情報(Camelotspの戻り値),内容のタイプ(str),次ページフラグ(bool),ループ変数(int)
    """
    try:
        NextFlag = False
        TO = True  # TimeOut判定変数
        # 第三引数に'stream'を渡すと表外の値を抽出できる
        if engine == "stream":
            tables = CTO.camelotTimeOut(path_pdf, PageVol, engine)
        else:
            tables = CTO.camelotTimeOut(path_pdf, PageVol, "")
        mp = []  # PDFminerのページ引数がリスト型なのでリストを初期化
        mp.append(y)  # PDFminerのページ引数作成
        # PDFminerで指定ページのテキスト取得-------------------------------
        Sbtext = extract_text(
            path_pdf, page_numbers=mp, maxpages=1, codec="utf-8"
        )  # テキストのみ取得できる
        Sbtext = Sbtext.replace("\n\n", ":")
        # 取得テキストに指定文字列が含まれるかで処理分け-----------------------------------------------------------------
        if "申告のお知らせ" in Sbtext:
            if "課税期間分の中間申告について" in Sbtext:
                TaxType = "etaxsyouhicyuukan"
            elif "事業年度等分中間（予定）申告について" in Sbtext:
                TaxType = "etaxjigyounendo"
            else:
                TaxType = "etaxosirase"
        elif "税⽬１" in Sbtext or "税⽬２" in Sbtext or "優先税⽬" in Sbtext or "第２税⽬" in Sbtext:
            TaxType = "etax3retu"
        elif "「消費税の納税義務者でなくなった旨の届出書」を提出していない場合には" in Sbtext:
            TaxType = "etaxsyouhi"
        elif "消費税簡易課税制度選択不適⽤届出" in Sbtext:
            TaxType = "etaxsyouhitodoke"
        elif "源泉所得税及復興特別所得税" in Sbtext:
            if "徴収⾼計算書の送付の要否" in Sbtext:
                TaxType = "etaxsyotoku2"
            else:
                TaxType = "etaxsyotoku"
        elif "前事業年度等" in Sbtext:
            TaxType = "etaxjigyounendo"
        elif "Copyright(C) TKC" in Sbtext:
            if "税務届出書類等作成支援システム(e-DMS)による電子申請・届出が完了しましたので、ご報告いたします。" in Sbtext:
                if "地 方 税 の 電 子 申 請・届 出 完 了 報 告 書" in Sbtext:
                    TaxType = "TKC2"
                else:
                    TaxType = "TKC"
            elif "ＴＫＣ電子申告システム" in Sbtext:
                TaxType = "TKC3"
                tables = CTO.camelotTimeOut(path_pdf, PageVol, "stream")
            elif "地 方 税 ポ ー タ ル シ ス テ ム ( e L T A X ) の" in Sbtext:
                TaxType = "TKC10"
                tables = CTO.camelotTimeOut(path_pdf, PageVol, "stream")
        else:
            if "送信された申告データを受付けました。" in Sbtext:
                if "GWE" in Sbtext:
                    TaxType = "eltaxList"
                else:
                    MOUTList = Settingtoml["OUTLIST"]["MJSOutList"]
                    Sbt = Sbtext.replace("\u3000", "").replace("\n", "")
                    for MOUTListItem in MOUTList:
                        if MOUTListItem in Sbt:
                            TaxType = "MJSOutList"
                            break
                        else:
                            TaxType = "MJS"
            elif "（e-Tax）" in Sbtext:
                TaxType = "etaxList"
            elif "GWE" in Sbtext:
                TaxType = "eltaxList"
            else:
                if "備考:" in Sbtext or "手続名:" in Sbtext:
                    MOUTList = Settingtoml["OUTLIST"]["MJSOutList"]
                    Sbt = Sbtext.replace("\u3000", "").replace("\n", "")
                    for MOUTListItem in MOUTList:
                        if MOUTListItem in Sbt:
                            TaxType = "MJSOutList"
                            break
                        else:
                            TaxType = "MJS"
                else:
                    TaxType = "None"
        # -----------------------------------------------------------------------------------------------------------
        # 取得したTKCの完了報告書PDF情報から、次ページに取得対象があるかを判定
        if "TKC" in TaxType:
            CLsp = Camelotsp(Sbtext, path_pdf, PageVol)
            if CLsp[0] is True:
                PageList = CLsp[3]
                Sbtext = CLsp[1]
                tCells = FPDF.CellsImport(
                    CDict,
                    Settingtoml,
                    SCode,
                    path_pdf,
                    tables,
                    y,
                    TaxType,
                    Sbtext,
                    DLCList,
                    NextFlag,
                    PageList,
                )
                if tCells[0] is False:
                    TO = False
                return True, TO, tables, tCells, TaxType, NextFlag, CLsp[2]
            else:
                if CLsp[1] == "Camelotspエラー":
                    y = y + 1
                return False, "TO", "tables", "tCells", "TaxType", "NextFlag", y
        elif TaxType == "None":
            logger.debug(path_pdf + "_" + str(y + 1) + "ページ目取得対象外")
            # logger.debug(path_pdf + "_" + str(y + 1) + "ページ目取得対象外")
            OutputList = [
                path_pdf.replace("/", "\\"),
                str(y + 1) + "ページ目取得対象外",
                SCode,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
            DLP = DiffListPlus(
                Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
            )  # 抽出リストに格納
            DLCList.append(DLP[1])  # できあがった抽出リストを保管
            return False, "None", "tables", "tCells", "TaxType", "NextFlag", y
        else:
            PageList = []
            tCells = FPDF.CellsImport(
                CDict,
                Settingtoml,
                SCode,
                path_pdf,
                tables,
                y,
                TaxType,
                Sbtext,
                DLCList,
                NextFlag,
                PageList,
            )
            # yを返して現在処理したページを更新------------------------------
            if tCells[0] is False:
                TO = False
                y = 1
            else:
                y = 1
            # ------------------------------------------------------------
            return True, TO, tables, tCells, TaxType, NextFlag, y
    except:  # TimeOut処理を記述
        TO = False  # TimeOut判定変数
        try:
            if len(tables) == 0:
                tables = "tables"
        except:
            tables = "tables"
        try:
            if len(tCells) == 0:
                tCells = "tCells"
        except:
            tCells = "tCells"
        try:
            if len(TaxType) == 0:
                TaxType = "TaxType"
        except:
            TaxType = "TaxType"
        try:
            if len(NextFlag) == 0:
                NextFlag = "NextFlag"
        except:
            NextFlag = "NextFlag"
        if not y == 0:
            y = 1
        return False, TO, tables, tCells, TaxType, NextFlag, y


# ----------------------------------------------------------------------------------------------


def CSVIndexSort(SCode, path_pdf, DLCList):
    """
    概要: PDFファイルに対するメイン処理
    @param SCode : 関与先コード(int)
    @param path_pdf : PDFパス(str)
    @param DLCList : 格納完了したリスト名一覧(list)
    @return : なし
    """
    # # ------------------------------------------------------------------------------------
    fp = open(path_pdf, "rb")  # PDFファイルを読み込み
    parser = PDFParser(fp)  # PDFperserを作成。
    document = PDFDocument(parser)  # PDFperserを格納。
    num_pages = 0  # ページ数格納変数を初期化
    num_pagesList = []
    for page in PDFPage.create_pages(document):  # ページオブジェ分ループ
        num_pages += 1  # ページ数カウント
        num_pagesList.append(num_pages - 1)
    print(num_pages)  # ページ数確認
    # ------------------------------------------------------------------------------------
    xy = 0  # 次ページ取得対象のPDFファイルだった場合のページカウント加算変数の初期化
    try:
        # PDFのページ数分ループ---------------------------------------------------------------------------
        for y in range(num_pages):
            # ページカウント加算変数が0以下もしくはループ変数と一致した場合
            if xy == y or xy <= 0:
                if y == 0:
                    NF = False  # 次ページフラグをFalse
                # 次ページフラグがTrueの場合
                if NF is True:
                    NF = False  # 次ページフラグをFalse
                    continue  # 次のループへ
                else:
                    PageVol = str(y + 1)
                    # TimeOutを加味したPDFRead処理TimeOut設定時間はContextTimeOut.pyにコンテキストで設定する
                    # 一回目の処理-----------------------------------------------
                    # # PDFテキスト内容で税目処理分け
                    # TO = CamelotSerch(
                    #     CDict, path_pdf, PageVol, Settingtoml, SCode, y, "", DLCList
                    # )  # return True, TO,tables, tCells
                    try:
                        # PDFテキスト内容で税目処理分け
                        TO = CamelotSerch(
                            CDict, path_pdf, PageVol, Settingtoml, SCode, y, "", DLCList
                        )  # return True, TO,tables, tCells
                        if type(TO[5]) == str:
                            NF = False
                        else:
                            NF = TO[5]
                    except:
                        NF = False
                    xy = y + TO[6]
                    # ---------------------------------------------------------
                    if "TKC" not in TO[4]:  # 一回目処理結果がTKCじゃなければ
                        if NF is False:
                            if type(TO[1]) == str and TO[1] == "TO":
                                logger.debug(
                                    path_pdf + "_" + str(TO[6] + 1) + "ページ目取得対象外"
                                )
                                # logger.debug(path_pdf + "_" + str(y + 1) + "ページ目取得対象外")
                                OutputList = [
                                    path_pdf.replace("/", "\\"),
                                    str(TO[6] + 1) + "ページ目取得対象外",
                                    SCode,
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                ]
                                DLP = DiffListPlus(
                                    Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
                                )  # 抽出リストに格納
                                DLCList.append(DLP[1])  # できあがった抽出リストを保管
                            elif TO[1] == "None":
                                print("None格納")
                            elif TO[1] is not False:  # 一回目処理がTimeOutしなかったら
                                tables = TO[2]
                                tCells = TO[3]
                                t_count = len(tables)  # PDFのテーブル数を格納
                                # PDFのテーブル数をが二つ以上なら----------------------------------
                                if t_count >= 2:
                                    # 既に取得済みの初めのページを格納
                                    DLP = DiffListPlus(
                                        tCells[1], tCells[2], ""
                                    )  # 抽出リストに格納
                                    DLCList.append(DLP[1])  # できあがった抽出リストを保管
                                    # 二回目の処理第六引数に'stream'を渡すと表外の値を抽出できる------
                                    # PDFテキスト内容で税目処理分け
                                    SB = CamelotSerch(
                                        CDict,
                                        path_pdf,
                                        PageVol,
                                        Settingtoml,
                                        SCode,
                                        y,
                                        "stream",
                                        DLCList,
                                    )  # return True, TO,tables, tCells
                                    # ---------------------------------------------------------
                                    Subtables = SB[2]
                                    SubtCells = SB[3]
                                    SubtCells[2][1] = str(SubtCells[2][1]) + "サブテーブル失敗"
                                    DLP = DiffListPlus(
                                        SubtCells[1], SubtCells[2], "Sub"
                                    )  # 抽出リストに格納
                                    DLCList.append(DLP[1])  # できあがった抽出リストを保管
                                    # TKCPDFの表外テキスト抽出処理------------------------------
                                    # PDFテキスト内容で税目処理分け
                                    SB = CamelotSerch(
                                        CDict,
                                        path_pdf,
                                        PageVol,
                                        Settingtoml,
                                        SCode,
                                        y,
                                        "stream",
                                        DLCList,
                                    )  # return True, TO,tables, tCells
                                    # ---------------------------------------------------------
                                else:
                                    # 既に取得済みの初めのページを格納
                                    # 二回目の処理第六引数に'stream'を渡すと表外の値を抽出できる------
                                    # PDFテキスト内容で税目処理分け
                                    if "TKC" in TO[4]:
                                        SB = CamelotSerch(
                                            CDict,
                                            path_pdf,
                                            PageVol,
                                            Settingtoml,
                                            SCode,
                                            y,
                                            "stream",
                                            DLCList,
                                        )  # return True, TO,tables, tCells
                                    # ---------------------------------------------------------
                                    DLP = DiffListPlus(
                                        tCells[1], tCells[2], ""
                                    )  # 抽出リストに格納
                                    DLCList.append(DLP[1])  # できあがった抽出リストを保管
                                # ------------------------------------------------------------
                            else:
                                # 二回目の処理第六引数に'stream'を渡すと表外の値を抽出できる------
                                # PDFテキスト内容で税目処理分け
                                SB = CamelotSerch(
                                    CDict,
                                    path_pdf,
                                    PageVol,
                                    Settingtoml,
                                    SCode,
                                    y,
                                    "stream",
                                    DLCList,
                                )  # return True, TO,tables, tCells
                                # ---------------------------------------------------------
                                if SB[1] is not False:  # 一回目処理がTimeOutしなかったら
                                    Subtables = SB[2]
                                    SubtCells = SB[3]
                                    DLP = DiffListPlus(
                                        SubtCells[1], SubtCells[2], ""
                                    )  # 抽出リストに格納
                                    DLCList.append(DLP[1])  # できあがった抽出リストを保管
                                else:
                                    logger.debug(
                                        path_pdf + "_" + str(y + 1) + "ページ目タイムアウト"
                                    )
                                    OutputList = [
                                        path_pdf.replace("/", "\\"),
                                        str(y + 1) + "ページ目タイムアウト",
                                        SCode,
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                        "",
                                    ]
                                    DLP = DiffListPlus(
                                        Settingtoml["CsvSaveEnc"]["ErrList"],
                                        OutputList,
                                        "",
                                    )  # 抽出リストに格納
                                    DLCList.append(DLP[1])  # できあがった抽出リストを保管
                        else:
                            tables = TO[2]
                            tCells = TO[3]
                            # 既に取得済みの初めのページを格納
                            if len(tCells[2]) == 3:
                                OutputList = [
                                    path_pdf.replace("/", "\\"),
                                    str(y + 1) + "ページ目NFエラー取得失敗",
                                    SCode,
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                    "",
                                ]
                                DLP = DiffListPlus(
                                    Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
                                )  # 抽出リストに格納
                                DLCList.append(DLP[1])  # できあがった抽出リストを保管
                            else:
                                DLP = DiffListPlus(tCells[1], tCells[2], "")  # 抽出リストに格納
                                DLCList.append(DLP[1])  # できあがった抽出リストを保管
                    else:  # 一回目処理結果がTKCだったら
                        if not TO[4] == "TKC3":
                            tables = TO[2]
                            tCells = TO[3]
                            # 既に取得済みの初めのページを格納
                            DLP = DiffListPlus(tCells[1], tCells[2], "")  # 抽出リストに格納
                            DLCList.append(DLP[1])  # できあがった抽出リストを保管
                        else:
                            print("TKCテキスト完了")
                            NF = TO[5]
            else:
                print("SKIP")
    except Exception as e:
        if TO[0] is not False:
            logger.debug(path_pdf + "_" + str(y + 1) + "ページ目取得失敗")
            OutputList = [
                path_pdf.replace("/", "\\"),
                str(y + 1) + "ページ目取得失敗",
                SCode,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
            DLP = DiffListPlus(
                Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
            )  # 抽出リストに格納
            if DLP[0] is True:
                DLCList.append(DLP[1])  # できあがった抽出リストを保管
            print(e)
        elif e.args[0] == "local variable 'tables' referenced before assignment":
            logger.debug(path_pdf + "_" + str(y + 1) + "ページ目取得失敗")
            OutputList = [
                path_pdf.replace("/", "\\"),
                str(y + 1) + "ページ目取得失敗",
                SCode,
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
                "",
            ]
            DLP = DiffListPlus(
                Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
            )  # 抽出リストに格納
            if DLP[0] is True:
                DLCList.append(DLP[1])  # できあがった抽出リストを保管
            print(e)


# ----------------------------------------------------------------------------------------
def PDFRead(URL, Settingtoml):
    # ------------------------------------------------------------------------------------
    # FCSV.CsvSaveEnc(URL.replace("\\","/") + "/PDFDataSuccess.csv",CSVList,"shiftjis")
    DLCList = []
    NoList = []
    dir_List = SerchdirFolders(URL)  # 指定URL配下のサブフォルダを取得
    print(dir_List)
    if len(dir_List) == 0:
        dir_Files = SerchdirFiles(URL)  # サブフォルダ配下のサブフォルダを取得
        print(dir_Files)
        for dir_FilesItem in dir_Files:
            if URL == dir_FilesItem[0]:
                dif = dir_FilesItem[0] + "\\" + dir_FilesItem[1]  # ファイル名
                dirsplit = dir_FilesItem[1].split("_")
                dirsplit = dirsplit[0].split(".")
                SCode = dirsplit[0]
                try:
                    if dif.endswith(".pdf") is True:  # pdfファイルのみ
                        path_pdf = dif.replace("\\", "/")  # PDFパスを代入
                        CSVIndexSort(SCode, path_pdf, DLCList)
                    else:
                        print("xdw")
                        path_pdf = dif.replace("\\", "/")  # PDFパスを代入
                        logger.debug(path_pdf + "_PDFファイル以外の為取得不可")
                        OutputList = [
                            path_pdf.replace("/", "\\"),
                            "PDFファイル以外の為取得不可",
                            SCode,
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                            "",
                        ]
                        DLP = DiffListPlus(
                            Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
                        )  # 抽出リストに格納
                        if DLP[0] is True:
                            DLCList.append(DLP[1])  # できあがった抽出リストを保管

                except Exception as e:
                    print(e)
                    if e.args[0] == 13:
                        print("PDFファイルアクセス拒否")
                        NoList.append(SCode)
                        NoList.append(path_pdf.replace("/", "\\"))
                        DLP = DiffListPlus(
                            Settingtoml["CsvSaveEnc"]["NotAccessList"], NoList, ""
                        )  # 抽出リストに格納
                        if DLP[0] is True:
                            DLCList.append(DLP[1])  # できあがった抽出リストを保管
                    else:
                        print(e)
    # ------------------------------------------------------------------------------------
    else:
        for dir_ListItem in dir_List:
            Serchd = dir_ListItem[0] + "\\" + dir_ListItem[1]  # サブフォルダ名
            dir_Files = SerchdirFiles(Serchd)  # サブフォルダ配下のサブフォルダを取得
            print(dir_Files)
            for dir_FilesItem in dir_Files:
                if Serchd == dir_FilesItem[0]:
                    dif = dir_FilesItem[0] + "\\" + dir_FilesItem[1]  # ファイル名
                    dirsplit = dir_FilesItem[1].split("_")
                    dirsplit = dirsplit[0].split(".")
                    SCode = dirsplit[0]
                    try:
                        if dif.endswith(".pdf") is True:  # pdfファイルのみ
                            path_pdf = dif.replace("\\", "/")  # PDFパスを代入
                            CSVIndexSort(SCode, path_pdf, DLCList)
                        else:
                            print("xdw")
                            path_pdf = dif.replace("\\", "/")  # PDFパスを代入
                            logger.debug(path_pdf + "_PDFファイル以外の為取得不可")
                            OutputList = [
                                path_pdf.replace("/", "\\"),
                                "PDFファイル以外の為取得不可",
                                SCode,
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                                "",
                            ]
                            DLP = DiffListPlus(
                                Settingtoml["CsvSaveEnc"]["ErrList"], OutputList, ""
                            )  # 抽出リストに格納
                            if DLP[0] is True:
                                DLCList.append(DLP[1])  # できあがった抽出リストを保管

                    except Exception as e:
                        print(e)
                        if e.args[0] == 13:
                            print("PDFファイルアクセス拒否")
                            NoList.append(SCode)
                            NoList.append(path_pdf.replace("/", "\\"))
                            DLP = DiffListPlus(
                                Settingtoml["CsvSaveEnc"]["NotAccessList"], NoList, ""
                            )  # 抽出リストに格納
                            if DLP[0] is True:
                                DLCList.append(DLP[1])  # できあがった抽出リストを保管
                        else:
                            print(e)
        # ------------------------------------------------------------------------------------
    ListURL = FC.CreFol(save_dir, "受信通知CSV")
    for f in os.listdir(ListURL):
        if os.path.isfile(os.path.join(ListURL, f)):
            if ".csv" in f:
                try:
                    # os.remove(os.path.join(ListURL, f))
                    osurl = ListURL + "/" + f
                    os.remove(osurl)
                except:
                    print("OSREMOVEError")
    DLCDiplicated = list(OrderedDict.fromkeys(DLCList))  # 抽出リストの重複削除
    for DItem in DLCDiplicated:
        try:
            DiffListCSVOUT(ListURL, DItem)  # 抽出リストをCSV書出し
        except:
            continue


# ------------------------------------------------------------------------------------
def CSVLog(URL, LogURL):
    URL = URL + "\\受信通知CSV"
    now = datetime.now()
    DY = "{0:%Y-%m-%d %H:%M:%S}".format(now)
    DY = DY.replace(":", "-")
    List = []
    ALLList = []
    for fd_path, sb_folder, sb_file in os.walk(URL):
        for fil in sb_file:
            if fil.endswith(".csv") is True:
                List.append([fd_path, fil])
    for ListItem in List:
        CURL = ListItem[0] + "\\" + ListItem[1]
        DFCSV = FCSV.CsvReadDtypeDict(CURL, str)
        DFCSVRow = np.array(DFCSV[1]).shape[0]
        for DI in range(DFCSVRow):
            AR = DFCSV[1].iloc[DI]
            ALLList.append(AR)
        # print(ALLList)
    LogURL = LogURL.replace("\\", "/") + "/" + DY + "_Log.csv"
    FCSV.CsvSaveNoHeader(LogURL, ALLList, "cp932")


def Main(obj):
    global MeUrl, CDict, LogURL, save_dir, Settingtoml
    # ------------------------------------------------------------------------------------
    MeUrl = obj.Img_dir_D
    save_dir = obj.dir_name
    # toml読込------------------------------------------------------------------------------
    with open(obj.Img_dir + r"/PDFReadForList/Setting.toml", encoding="utf-8") as f:
        Settingtoml = toml.load(f)
        print(Settingtoml)
    # ----------------------------------------------------------------------------------------
    CDict = CSVSet.CSVIndexSortFuncArray  # 外部よりdict変数取得
    URL = obj.dir_name
    # URL = "\\\\nas-sv\\B_監査etc\\B2_電子ﾌｧｲﾙ\\ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\TEST"
    LogURL = "\\\\nas-sv\\B_監査etc\\B2_電子ﾌｧｲﾙ\\ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\PDFREADLog"
    try:
        logger.debug(URL + "内のPDF抽出開始")

        PDFRead(URL, Settingtoml)

        # CSVLog(URL, LogURL)
        logger.debug(URL + "内のPDF抽出完了")
    except Exception as e:
        logger.debug("エラー終了" + e)
