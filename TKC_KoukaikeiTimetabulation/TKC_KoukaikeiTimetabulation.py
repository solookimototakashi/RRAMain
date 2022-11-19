# インポート
import pyautogui as pg
import time
import pandas as pd
import numpy as np
import os
from datetime import datetime as dt
import traceback
import pyautogui
import pyperclip
import aggregate
from tkinter import messagebox
from chardet.universaldetector import UniversalDetector

# 以下は同ディレクトリにpyファイルを保管して下さい。
import WarekiHenkan  # 西暦変換
import OMSOpen  # OMS起動処理

"""
作成者:沖本卓士
作成日:
最終更新日:2022/11/14
稼働設定:解像度 1920*1080 表示スケール125%
####################################################
処理の流れ
####################################################
1:保存フォルダのチェック
↓
2:OMSにログインする
↓
3:時間管理TMSを立ち上げる
↓
4:時間管理TMSで比較期間を前月で指定する
↓
5:間接業務CSVを出力し、業務別データを特定する(順序特定)
↓
6:移動時間の集計を行う
↓
7:B2・A8・A9・A10・A11の作業ごとに以下の集計を行う
    ・関与先CSVを出力(順序特定)
    ・関与先メニュー内へ移動
    ・担当者CSVを出力(順序特定)
    ・担当者毎にファイル名を付け保存フォルダに保存
↓
8:全作業集計完了後エクセルマクロ読込用のCSVファイル作成
####################################################
ディレクトリ
・__pycache__
    実行時コンパイルモジュールのキャッシュ
・img
    画像フォルダ
・aggregate.py
    マクロ集計用CSV作成自作モジュール
・OMSOpen.py
    TKCOMS起動自作モジュール
・WarekiHenkan.py
    西暦変換自作モジュール
"""
# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitXPATH(UIPATH, driver):
    """
    AppiumDriverからXPATH要素を取得するまで待機
    """
    for x in range(1000):
        try:
            driver.find_element_by_xpath(UIPATH)
            Flag = 1
            return True
        except:
            Flag = 0
    if Flag == 0:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitAutomationId(UIPATH, driver):
    """
    AppiumDriverからID要素を取得するまで待機
    """
    for x in range(1000):
        try:
            driver.find_element_by_accessibility_id(UIPATH)
            Flag = 1
            return True
        except:
            Flag = 0
    if Flag == 0:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitName(UIPATH, driver):
    """
    AppiumDriverからName要素を取得するまで待機
    """
    for x in range(1000):
        try:
            driver.find_element_by_Name(UIPATH)
            Flag = 1
            return True
        except:
            Flag = 0
    if Flag == 0:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitclassname(UIPATH, driver):
    """
    AppiumDriverからClassName要素を取得するまで待機
    """
    for x in range(10000):
        try:
            driver.find_element_by_class_name(UIPATH)
            Flag = 1
            return True
        except:
            Flag = 0
    if Flag == 0:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverFindClass(UIPATH, driver):  # XPATH要素を取得するまで待機
    """
    AppiumDriverからClassName要素を複数取得する
    """
    for x in range(10000):
        try:
            elList = driver.find_elements_by_class_name(UIPATH)
            Flag = 1
            return True, elList
        except:
            Flag = 0
    if Flag == 0:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverCheck(Hub, ObjName, driver):
    """
    AppiumDriver要素を取得する
    """
    for x in range(1000):
        if Hub == "AutomationID":
            if (
                DriverUIWaitAutomationId(ObjName, driver) is True
            ):  # OMSメニューの年調起動ボタンを判定して初期処理分け
                # 正常待機後処理
                driver.find_element_by_accessibility_id(ObjName)  # 一括電子申告送信ボタン
                return True
            else:
                # 異常待機後処理
                print("要素取得に失敗しました。")
        elif Hub == "XPATH":
            if DriverUIWaitXPATH(ObjName, driver) is True:  # OMSメニューの年調起動ボタンを判定して初期処理分け
                # 正常待機後処理
                driver.find_element_by_xpath(ObjName)  # 一括電子申告送信ボタン
                return True
            else:
                # 異常待機後処理
                print("要素取得に失敗しました。")
        elif Hub == "Name":
            if DriverUIWaitName(ObjName, driver) is True:  # OMSメニューの年調起動ボタンを判定して初期処理分け
                # 正常待機後処理
                driver.find_element_by_Name(ObjName)  # 一括電子申告送信ボタン
                return True
            else:
                # 異常待機後処理
                print("要素取得に失敗しました。")


# ----------------------------------------------------------------------------------------------------------------------
def DriverClick(Hub, ObjName, driver):
    """
    AppiumDriver要素をクリックする
    """
    if Hub == "AutomationID":
        if (
            DriverUIWaitAutomationId(ObjName, driver) is True
        ):  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            OMSObj = driver.find_element_by_accessibility_id(ObjName)  # 一括電子申告送信ボタン
            OMSObj.click()
            return OMSObj
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")
    elif Hub == "XPATH":
        if DriverUIWaitXPATH(ObjName, driver) is True:  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            OMSObj = driver.find_element_by_xpath(ObjName)  # 一括電子申告送信ボタン
            OMSObj.click()
            return OMSObj
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")
    elif Hub == "Name":
        if DriverUIWaitName(ObjName, driver) is True:  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            OMSObj = driver.find_element_by_Name(ObjName)  # 一括電子申告送信ボタン
            OMSObj.click()
            return OMSObj
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")
    elif Hub == "class_name":
        if DriverUIWaitclassname(ObjName, driver) is True:  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            OMSObj = driver.find_element_by_class_name(ObjName)  # 一括電子申告送信ボタン
            OMSObj.click()
            return OMSObj
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")


# ----------------------------------------------------------------------------------------------------------------------
def ImgCheck(URL, FileName, conf, LoopVal):
    """
    画像の有無を判定
    """
    ImgURL = URL + "/" + FileName
    for x in range(LoopVal):
        try:
            p = pyautogui.locateOnScreen(ImgURL, confidence=conf)
            x, y = pyautogui.center(p)
            return True, x, y
        except:
            Flag = 0
    if Flag == 0:
        return False, "", ""


# ----------------------------------------------------------------------------------------------------------------------
def ImgNothingCheck(URL, FileName, conf, LoopVal):
    """
    画像がなければTrueを返す
    """
    ImgURL = URL + "/" + FileName
    for x in range(LoopVal):
        try:
            p = pyautogui.locateOnScreen(ImgURL, confidence=conf)
            x, y = pyautogui.center(p)
            return False
        except:
            Flag = 0
    if Flag == 0:
        return True


# ----------------------------------------------------------------------------------------------------------------------
def ImgCheckForList(URL, List, conf):
    """
    複数の画像の有無を判定
    """
    for x in range(10):
        for ListItem in List:
            ImgURL = URL + "/" + ListItem
            try:
                p = pyautogui.locateOnScreen(ImgURL, confidence=conf)
                x, y = pyautogui.center(p)
                return True, ListItem
                break
            except:
                Flag = 0
    if Flag == 0:
        return False, ""


# ----------------------------------------------------------------------------------------------------------------------
def ImgClick(URL, FileName, conf, LoopVal):
    """
    画像があればクリックしてx,y軸を返す
    """
    ImgURL = URL + "/" + FileName
    for x in range(10):
        if (
            ImgCheck(URL, FileName, conf, LoopVal)[0] is True
        ):  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            for y in range(10):
                try:
                    p = pyautogui.locateOnScreen(ImgURL, confidence=conf)
                    x, y = pyautogui.center(p)
                    pyautogui.click(x, y)
                    time.sleep(1)
                    return x, y
                except:
                    print("失敗")
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")


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
def CSVGet(FileUrl):
    """
    ファイルのエンコードを調べ、Dataframeにする
    """
    try:
        SerchEnc = format(getFileEncoding(FileUrl))
        MasterCSV = pd.read_csv(FileUrl, encoding=SerchEnc)
        return True, MasterCSV
    except:
        return False, ""


# ----------------------------------------------------------------------------------------------------------------------
def ShiftJisCSVGet(FileUrl):
    """
    ファイルのエンコードを直接SHIFT_JIS指定し、Dataframeにする
    """
    try:
        MasterCSV = pd.read_csv(FileUrl, encoding="SHIFT_JIS")
        return True, MasterCSV
    except:
        return False, ""


# ----------------------------------------------------------------------------------------------------------------------
def CSVCheck(Key, CSVArr, ColName):
    """
    指定CSVDataframeから引数KeyのIndexを調べる
    """
    CSVArrRow = np.array(CSVArr).shape[0]  # 配列行数取得
    try:
        for x in range(CSVArrRow):
            CSVRowData = CSVArr.iloc[x, :]
            CSVTarget = CSVRowData[ColName].replace("\u3000", "")
            if Key == CSVTarget:
                return True, x
    except:
        return False, ""


# ----------------------------------------------------------------------------------------------------------------------
def FirstOpen(obj, URL, Tax):
    """
    初回起動(抽出期間の設定)
    """
    try:
        time.sleep(1)
        ImgClick(URL, "TMSOpen.png", 0.9, 2)
        while pg.locateOnScreen(URL + "/GyoumuBunsekiBtn.png", confidence=0.9) is None:
            time.sleep(1)
        time.sleep(1)
        ImgClick(URL, "GyoumuBunsekiBtn.png", 0.9, 2)
        # while pg.locateOnScreen(URL + "/GyoumuBunsekiWin.png" , confidence=0.9) is None:
        #     time.sleep(1)
        while pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9) is not None:
            time.sleep(1)
        time.sleep(1)
        pg.press("f7")
        time.sleep(1)
        while pg.locateOnScreen(URL + "/SiteiKikanStr.png", confidence=0.9) is None:
            time.sleep(1)
        ImgClick(URL, "SiteiKikanStr.png", 0.9, 2)
        time.sleep(1)
        ImgClick(URL, "SiteiKikanOKBtn.png", 0.9, 2)
        time.sleep(1)
        ImgClick(URL, "TimeBox.png", 0.9, 2)
        time.sleep(1)
        pg.press("return")
        time.sleep(1)
        # 現在の年度(str)
        yPar = str(obj.i_id_txt.get())
        # 先月(str)
        mPar = str(obj.i_id_txt2.get())
        yPar = str(WarekiHenkan.Wareki.from_ad(int(yPar)).year)

        pg.write(yPar, interval=0.01)
        pg.press("return")
        time.sleep(1)
        pg.write(mPar, interval=0.01)
        pg.press("return")
        time.sleep(1)
        pg.press("return")
        time.sleep(1)
        pg.write(yPar, interval=0.01)
        pg.press("return")
        time.sleep(1)
        pg.write(mPar, interval=0.01)
        pg.press("return")
        time.sleep(1)
        pg.press("return")
        time.sleep(3)
        ImgClick(URL, "GyoumubetuTab.png", 0.9, 2)
        time.sleep(1)
        while pg.locateOnScreen(URL + "/TyokusetuOpenFlag.png", confidence=0.9) is None:
            time.sleep(1)
        time.sleep(1)
        ImgClick(URL, "KansetuTab.png", 0.9, 2)
        Kansetu = TKCCSVOut(URL, "KANSETU")
        if Kansetu is True:
            KansetuList = CSVGet(URL + "/KANSETU.CSV")
            if KansetuList[0] is True:
                KansetuListRow = CSVCheck("E4移動時間", KansetuList[1], "業務")
                if KansetuListRow[0] is True:
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 1)
                    for x in range(KansetuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/TantoubetuFlag.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    time.sleep(1)
                    TKCTimeCSVOut(URL, "移動時間", SaveDir, Tax)
                    time.sleep(1)
                    pg.press("f10")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/TyokuKanTab.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    time.sleep(1)
                    ImgClick(URL, "TyokusetuTab.png", 0.9, 1)
        return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def TKCCSVOut(URL, Title):
    """
    TKCメニュー設定CSV出力フロー
    """
    try:
        ImgClick(URL, "ExcelTab.png", 0.9, 2)
        while pg.locateOnScreen(URL + "/KiridasiFlag.png", confidence=0.9) is None:
            time.sleep(1)
        ImgList = ["KiridasiType.png", "KiridasiType2.png"]
        ImgListAns = ImgCheckForList(URL, ImgList, 0.9)
        ImgClick(URL, ImgListAns[1], 0.9, 2)
        time.sleep(1)
        pg.press(["down", "down", "down"])
        pg.press(["return"])
        time.sleep(1)
        pg.press(["tab", "tab", "tab"])
        pg.press("delete")
        URL = URL.replace("/", "\\")
        pyperclip.copy(URL)
        pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        pg.press(["return"])
        time.sleep(1)
        pg.press("delete")
        pg.write(Title, interval=0.01)
        time.sleep(1)
        pg.press(["return"])
        time.sleep(1)
        pg.press(["return"])
        time.sleep(1)
        ImgClick(URL, "KiridasiSave.png", 0.9, 2)
        time.sleep(1)
        if ImgCheck(URL, "OverFileList.png", 0.9, 1)[0] is True:
            time.sleep(1)
            ImgClick(URL, "OverFileListYes.png", 0.9, 2)
        return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def TKCB2CSVOut(URL, CSVURL, Title):
    """
    TKCメニューB2公会計営業CSV出力フロー
    """
    try:
        ImgClick(URL, "ExcelTab.png", 0.9, 2)
        while pg.locateOnScreen(URL + "/KiridasiFlag.png", confidence=0.9) is None:
            time.sleep(1)
        ImgList = ["KiridasiType.png", "KiridasiType2.png"]
        ImgListAns = ImgCheckForList(URL, ImgList, 0.9)
        ImgClick(URL, ImgListAns[1], 0.9, 2)
        time.sleep(1)
        pg.press(["down", "down", "down"])
        pg.press(["return"])
        time.sleep(1)
        pg.press(["tab", "tab", "tab"])
        pg.press("delete")
        CSVURL = CSVURL.replace("/", "\\")
        pyperclip.copy(CSVURL)
        pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        pg.press(["return"])
        time.sleep(1)
        pg.press("delete")
        pyperclip.copy(Title)
        pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        time.sleep(1)
        pg.press(["return"])
        time.sleep(1)
        pg.press(["return"])
        time.sleep(1)
        ImgClick(URL, "KiridasiSave.png", 0.9, 2)
        time.sleep(1)
        if ImgCheck(URL, "OverFileList.png", 0.9, 1)[0] is True:
            time.sleep(1)
            ImgClick(URL, "OverFileListYes.png", 0.9, 2)
        return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def TKCTimeCSVOut(URL, Title, FileURL, Tax):
    """
    TKCメニューCSV出力フロー
    """
    try:
        ImgClick(URL, "ExcelTab.png", 0.9, 2)
        while pg.locateOnScreen(URL + "/KiridasiFlag.png", confidence=0.9) is None:
            time.sleep(1)
        ImgList = ["KiridasiType.png", "KiridasiType2.png"]
        ImgListAns = ImgCheckForList(URL, ImgList, 0.9)
        ImgClick(URL, ImgListAns[1], 0.9, 2)
        time.sleep(1)
        pg.press(["down", "down", "down"])
        pg.press(["return"])
        time.sleep(1)
        pg.press(["tab", "tab", "tab"])
        pg.press("delete")
        FileURL = FileURL.replace("/", "\\")
        pyperclip.copy(FileURL)
        pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        pg.press(["return"])
        time.sleep(1)
        pg.press("delete")
        if Tax != "":
            Title = Title + "_" + Tax
        pyperclip.copy(Title)
        pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        pg.press(["return"])
        time.sleep(1)
        pg.press(["return"])
        time.sleep(1)
        ImgClick(URL, "KiridasiSave.png", 0.9, 2)
        time.sleep(1)
        if ImgCheck(URL, "OverFileList.png", 0.9, 1)[0] is True:
            time.sleep(1)
            ImgClick(URL, "OverFileListYes.png", 0.9, 2)
        return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def KanyoScroll(URL, Tax):  # 関与先毎の時間集計操作
    """
    関与先毎の集計処理
    """
    KCSVO = TKCCSVOut(URL, "KANYOSAKIBETU")
    if KCSVO is True:
        KanyosakibetuList = ShiftJisCSVGet(URL + "/KANYOSAKIBETU.CSV")
        if KanyosakibetuList[0] is True:
            KArr = KanyosakibetuList[1]
            KArrRow = np.array(KArr).shape[0]  # 配列行数取得
            for y in range(KArrRow):
                if y <= 11:
                    KArrRowData = KArr.iloc[y, :]
                    KArrName = KArrRowData["関与先"]
                    KArrListRow = CSVCheck(KArrName, KArr, "関与先")
                    if KArrListRow[0] is True:
                        time.sleep(1)
                        ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                        for z in range(KArrListRow[1]):
                            pg.press("down")
                        time.sleep(1)
                        pg.press("return")
                        time.sleep(1)
                        while (
                            pg.locateOnScreen(
                                URL + "/TantoubetuFlag.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        TantouCSVO = TKCCSVOut(URL, "TANTOUBETU")
                        if TantouCSVO is True:
                            TantouList = CSVGet(URL + "/TANTOUBETU.CSV")
                            if TantouList[0] is True:
                                TanArr = TantouList[1]
                                TanArrRow = np.array(TanArr).shape[0]  # 配列行数取得
                                for z in range(TanArrRow):
                                    TanArrData = TanArr.iloc[z, :]
                                    TanName = TanArrData["担当者"]
                                    TanArrName = TanArrData["担当者"].replace("\u3000", "")
                                    TanArrTime = TanArrData["実\u3000績(A)"]
                                    if not TanArrName == "【合計】":
                                        if TanArrTime is not False:
                                            TanArrRowListRow = CSVCheck(
                                                TanArrName, TanArr, "担当者"
                                            )
                                            if TanArrRowListRow[0] is True:
                                                time.sleep(1)
                                                ImgClick(
                                                    URL,
                                                    "1gyouUnderArrow.png",
                                                    0.9,
                                                    5,
                                                )
                                                for z in range(TanArrRowListRow[1]):
                                                    pg.press("down")
                                                time.sleep(1)
                                                pg.press("return")
                                                time.sleep(1)
                                                while (
                                                    pg.locateOnScreen(
                                                        URL + "/KatudouTab.png",
                                                        confidence=0.9,
                                                    )
                                                    is None
                                                ):
                                                    time.sleep(1)
                                                FN = KArrName + "_" + TanName
                                                TKCTimeCSVOut(
                                                    URL,
                                                    FN,
                                                    SaveDir,
                                                    Tax,
                                                )
                                                time.sleep(1)
                                                pg.press("f10")
                                                time.sleep(1)
                                                while (
                                                    pg.locateOnScreen(
                                                        URL + "/TantoubetuFlag.png",
                                                        confidence=0.9,
                                                    )
                                                    is None
                                                ):
                                                    time.sleep(1)
                                                time.sleep(1)
                                                ImgClick(
                                                    URL,
                                                    "1gyouUnderArrow.png",
                                                    0.9,
                                                    5,
                                                )
                                                for z in range(TanArrRowListRow[1]):
                                                    pg.press("up")
                                time.sleep(1)
                                while (
                                    pg.locateOnScreen(
                                        URL + "/TantoubetuFlag.png", confidence=0.9
                                    )
                                    is None
                                ):
                                    time.sleep(1)
                                time.sleep(1)
                                ImgClick(
                                    URL,
                                    "TantoubetuFlag.png",
                                    0.9,
                                    5,
                                )
                                pg.press("f10")
                                time.sleep(1)
                                KouList = [
                                    "KousagyouOpen.png",
                                    "KousagyouOpen2.png",
                                    "KousagyouOpen3.png",
                                    "KousagyouOpen4.png",
                                    "KousagyouOpen5.png",
                                ]
                                while ImgCheckForList(URL, KouList, 0.9)[0] is False:
                                    time.sleep(1)
                                time.sleep(1)
                                ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                                time.sleep(1)
                                for z in range(KArrListRow[1]):
                                    pg.press("up")
                                time.sleep(1)
            pg.press("f10")
            print("")


# ----------------------------------------------------------------------------------------------------------------------
def EigyouScroll(URL):  # 関与先毎の時間集計操作
    """
    B2公会計営業処理
    """
    KCSVO = TKCCSVOut(URL, "KANYOSAKIBETU")
    if KCSVO is True:
        KanyosakibetuList = ShiftJisCSVGet(URL + "/KANYOSAKIBETU.CSV")
        if KanyosakibetuList[0] is True:
            KArr = KanyosakibetuList[1]
            KArrRow = np.array(KArr).shape[0]  # 配列行数取得
            for y in range(KArrRow):
                if y < KArrRow - 1:
                    KArrRowData = KArr.iloc[y, :]
                    KArrName = KArrRowData["関与先"]
                    KArrTime = KArrRowData["実　績(A)"]
                    if KArrTime != KArrTime is False:
                        KArrListRow = CSVCheck(KArrName, KArr, "関与先")
                        if KArrListRow[0] is True:
                            time.sleep(1)
                            ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                            for z in range(KArrListRow[1]):
                                pg.press("down")
                            time.sleep(1)
                            pg.press("return")

                            time.sleep(1)
                            while (
                                pg.locateOnScreen(
                                    URL + "/TantoubetuFlag.png", confidence=0.9
                                )
                                is None
                            ):
                                time.sleep(1)
                            TantouB2CSVO = TKCB2CSVOut(
                                URL,
                                SaveDir,
                                KArrName + "_B2",
                            )
                            if TantouB2CSVO is True:
                                time.sleep(1)
                                pg.press("f10")
                                time.sleep(1)
                                ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                                for z in reversed(range(KArrListRow[1])):
                                    pg.press("up")
                                time.sleep(1)
                            else:
                                print("B2CSV出力失敗")


# ----------------------------------------------------------------------------------------------------------------------
def MainFlow(obj):
    """
    メイン(作業内容別処理)
    """
    # 4724ポート指定でappiumサーバー起動バッチを開く
    BatUrl = os.getcwd().replace("\\", "/") + r"\bat\AWADriverOpen.bat"
    # OMSを起動しログイン後インスタンス化
    driver = OMSOpen.MainFlow(
        BatUrl, obj.open_dir, "RPAPhoto", obj.id_txt.get(), obj.pass_txt.get()
    )
    # RPA操作用の画像ファイル保管場所
    URL = obj.Img_dir_D
    Tax = ""
    # 初回起動
    FMO = FirstOpen(obj, URL, Tax)
    if FMO is True:
        # 業務別CSV出力
        TCSVO = TKCCSVOut(URL, "TGYOUMULIST")
        if TCSVO is True:
            # 業務別CSV格納
            TgyoumuList = CSVGet(URL + "/TGYOUMULIST.CSV")
            # 業務別処理#########################################################################
            if TgyoumuList[0] is True:
                # ------------------------------------------------------------------------------
                # 業務別CSVから引数1のIndexを取得
                TgyoumuListRow = CSVCheck("B2公会計営業", TgyoumuList[1], "業務")
                if TgyoumuListRow[0] is True:
                    Tax = "B2"
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    for x in range(TgyoumuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/KanyoTab.png", confidence=0.9) is None
                    ):
                        time.sleep(1)
                    ImgClick(URL, "KanyoTab.png", 0.9, 5)
                    time.sleep(1)
                    # 関与先毎の時間集計操作
                    KanyoScroll(URL, Tax)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    for x in range(TgyoumuListRow[1]):
                        pg.press("up")
                # ------------------------------------------------------------------------------
                # 業務別CSVから引数1のIndexを取得
                TgyoumuListRow = CSVCheck("A8公会計作業（固定資産台帳）", TgyoumuList[1], "業務")
                if TgyoumuListRow[0] is True:
                    Tax = "A8"
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/KanyoTab.png", confidence=0.9) is None
                    ):
                        time.sleep(1)
                    ImgClick(URL, "KanyoTab.png", 0.9, 5)
                    time.sleep(1)
                    # 関与先毎の時間集計操作
                    KanyoScroll(URL, Tax)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("up")
                # ------------------------------------------------------------------------------
                # 業務別CSVから引数1のIndexを取得
                TgyoumuListRow = CSVCheck("A9公会計作業（財務書類）", TgyoumuList[1], "業務")
                if TgyoumuListRow[0] is True:
                    Tax = "A9"
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/KanyoTab.png", confidence=0.9) is None
                    ):
                        time.sleep(1)
                    ImgClick(URL, "KanyoTab.png", 0.9, 5)
                    time.sleep(1)
                    # 関与先毎の時間集計操作
                    KanyoScroll(URL, Tax)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("up")
                # ------------------------------------------------------------------------------
                # 業務別CSVから引数1のIndexを取得
                TgyoumuListRow = CSVCheck("A10公会計作業（その他）", TgyoumuList[1], "業務")
                if TgyoumuListRow[0] is True:
                    Tax = "A10"
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/KanyoTab.png", confidence=0.9) is None
                    ):
                        time.sleep(1)
                    ImgClick(URL, "KanyoTab.png", 0.9, 5)
                    time.sleep(1)
                    # 関与先毎の時間集計操作
                    KanyoScroll(URL, Tax)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("up")
                # ------------------------------------------------------------------------------
                # 業務別CSVから引数1のIndexを取得
                TgyoumuListRow = CSVCheck("A11公営作業", TgyoumuList[1], "業務")
                if TgyoumuListRow[0] is True:
                    Tax = "A11"
                    time.sleep(1)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("down")
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/Syuukeityuu.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    while (
                        pg.locateOnScreen(URL + "/KanyoTab.png", confidence=0.9) is None
                    ):
                        time.sleep(1)
                    ImgClick(URL, "KanyoTab.png", 0.9, 5)
                    time.sleep(1)
                    # 関与先毎の時間集計操作
                    KanyoScroll(URL, Tax)
                    ImgClick(URL, "1gyouUnderArrow.png", 0.9, 5)
                    # 目的のIndexまでキー入力
                    for x in range(TgyoumuListRow[1]):
                        pg.press("up")
                # ------------------------------------------------------------------------------
            # ###################################################################################
        return True
    else:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def DirSearch(obj):
    """
    フォルダチェック・作成
    """
    URL = obj.dir_name
    # 現在の年度(str)
    yPar = str(obj.i_id_txt.get())
    # 先月(str)
    mPar = str(obj.i_id_txt2.get())

    for x in range(100):
        for y in range(12):
            yp, mp = int(yPar) + x, int(mPar) + y
            mp = format(mp, "02")
            if f"{str(yp)}-{str(mp)}" in URL:
                Usp = URL.split(f"/{str(yp)}-{str(mp)}")
                URL = Usp[0]
                break

            yp, mp = int(yPar) - x, int(mPar) - y
            mp = format(mp, "02")
            if f"{str(yp)}-{str(mp)}" in URL:
                Usp = URL.split(f"/{str(yp)}-{str(mp)}")
                URL = Usp[0]
                break

    URL = URL + r"\\" + yPar + "-" + mPar

    if os.path.isdir(URL) is True:
        if os.path.isdir(URL + r"\\BACKUP") is False:
            os.mkdir(URL + r"\\BACKUP")
        URL = URL + r"\\担当者別"
        if os.path.isdir(URL) is False:
            os.mkdir(URL)
        return URL
    else:
        os.mkdir(URL)
        if os.path.isdir(URL + r"\\BACKUP") is False:
            os.mkdir(URL + r"\\BACKUP")
        URL = URL + r"\\担当者別"
        if os.path.isdir(URL) is False:
            os.mkdir(URL)
        return URL


# ----------------------------------------------------------------------------------------------------------------------
# if __name__ == "__main__":
def Main(obj):
    global SaveDir

    SaveDir = DirSearch(obj)
    try:
        MF = MainFlow(obj)
        if MF is True:
            # 集計結果をBACKUPフォルダにCSV出力
            AM = aggregate.Main(obj.i_id_txt.get(), obj.i_id_txt2.get())
            if AM is True:
                return True
            else:
                messagebox.showinfo(
                    "注意",
                    "個別データを集計時に失敗しました。",
                )
                return False
        else:
            return False
    except:
        traceback.print_exc()
        return False
