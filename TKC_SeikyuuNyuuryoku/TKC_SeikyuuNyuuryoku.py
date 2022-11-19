###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# モジュールインポート
import pyautogui as pg
import time
import OMSOpen

# pandasインポート
import pandas as pd

# osインポート
import os

# datetimeインポート
from datetime import datetime as dt

# 例外処理判定の為のtracebackインポート
import traceback

# pandas(pd)で関与先データCSVを取得
import pyautogui
import pyperclip
import tkinter
import tkinter.filedialog

# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitXPATH(UIPATH, driver):  # XPATH要素を取得するまで待機
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
def DriverUIWaitAutomationId(UIPATH, driver):  # XPATH要素を取得するまで待機
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
def DriverUIWaitName(UIPATH, driver):  # XPATH要素を取得するまで待機
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
def DriverUIWaitclassname(UIPATH, driver):  # XPATH要素を取得するまで待機
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
# ----------------------------------------------------------------------------------------------------------------------
def DriverFindClass(UIPATH, driver):  # XPATH要素を取得するまで待機
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
def DriverCheck(Hub, ObjName, driver):  # XPATH要素を取得するまで待機
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
def ImgCheck(FolURL2, FileName, conf, LoopVal):  # 画像があればTrueを返す関数
    ImgURL = FolURL2 + "/" + FileName
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
def ImgNothingCheck(FolURL2, FileName, conf, LoopVal):  # 画像がなければTrueを返す
    ImgURL = FolURL2 + "/" + FileName
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
def ImgCheckForList(FolURL2, List, conf):  # リスト内の画像があればTrueと画像名を返す
    for x in range(10):
        for ListItem in List:
            ImgURL = FolURL2 + "/" + ListItem
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
def ImgClick(FolURL2, FileName, conf, LoopVal):  # 画像があればクリックしてx,y軸を返す
    ImgURL = FolURL2 + "/" + FileName
    for x in range(10):
        if (
            ImgCheck(FolURL2, FileName, conf, LoopVal)[0] is True
        ):  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            for y in range(10):
                try:
                    p = pyautogui.locateOnScreen(ImgURL, confidence=conf)
                    x, y = pyautogui.center(p)
                    pyautogui.click(x, y)
                    time.sleep(2)
                    return x, y
                except:
                    print("失敗")
        else:
            # 異常待機後処理
            print("要素取得に失敗しました。")


# ----------------------------------------------------------------------------------------------------------------------
def EraceIMGWait(FolURL2, FileName):
    try:
        while all(pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9)) is True:
            time.sleep(2)
    except:
        print("待機終了")


# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
def SortPDF(PDFName):
    Fol = str(dt.today().year) + "-" + str(dt.today().month)
    pt = "\\\\nas-sv\\B_監査etc\\B2_電子ﾌｧｲﾙ\\ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\" + Fol + "\\送信分受信通知"
    # path = path.replace('\\','/')#先
    PDFFileList = os.listdir(pt)
    Cou = 1
    for PDFItem in PDFFileList:
        PDFName = PDFName.replace("\u3000", "").replace("PDF", "").replace("pdf", "")
        PDFItem = PDFItem.replace("\u3000", "").replace("PDF", "").replace("pdf", "")
        if PDFName in PDFItem:
            Cou = Cou + 1
    return str(Cou), pt


# ----------------------------------------------------------------------------------------------------------------------
def FMSOpen(driver, FolURL2, xls_data, KamokuCD, Lyear, Lmonth, Lday):
    # 要素クリック----------------------------------------------------------------------------------------------------------
    Hub = "AutomationID"
    ObjName = "um10PictureButton"
    DriverClick(Hub, ObjName, driver)
    # ----------------------------------------------------------------------------------------------------------------------
    FileName = "TodayTitle.png"
    while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
        time.sleep(2)
    pg.write(str(Lyear), interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    time.sleep(1)
    pg.write(str(Lmonth), interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    time.sleep(1)
    pg.write(str(Lday), interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    pg.press("return")
    time.sleep(3)
    conf = 0.9
    # -------------------------------------------
    FileName = "F10END.png"
    F10E = ImgCheck(FolURL2, FileName, 0.9, 5)
    pyautogui.click(F10E[1] + 10, F10E[2] - 30)
    # -------------------------------------------
    pg.write("202", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    FileName = "KanyoItiWin.png"
    while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
        time.sleep(2)
    conf = 0.9
    LoopVal = 10
    FileName = "KanyoSyatyouTAB.png"
    ImgClick(FolURL2, FileName, conf, LoopVal)
    FileName = "SyatyouOpen.png"
    while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
        time.sleep(2)


# ----------------------------------------------------------------------------------------------------------------------
def FindMenu(driver, FolURL2, xls_cd):
    pg.press("f8")
    FileName = "FindTitle.png"
    while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
        time.sleep(2)
    pg.write(xls_cd, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    pg.write(xls_cd, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
    pg.press("return")
    pg.press("f4")
    time.sleep(4)
    conf = 0.99999
    LoopVal = 11
    FileName = "NoData.png"

    if ImgCheck(FolURL2, FileName, conf, LoopVal)[0] is True:
        return False
    else:
        return True


# ----------------------------------------------------------------------------------------------------------------------
def FirstAction(driver, FolURL2, xls_cd, xls_name, xls_mn, xls_tx, UpList, KamokuCD):
    FM = FindMenu(driver, FolURL2, xls_cd)
    if FM is True:
        time.sleep(2)
        conf = 0.9
        LoopVal = 10
        # -------------------------------------------
        FileName = "F10Close.png"
        F10C = ImgCheck(FolURL2, FileName, 0.9, 5)
        pyautogui.click(F10C[1] + int(F10C[1] / 100), F10C[2] - int(F10C[2] / 20))
        # -------------------------------------------
        pg.write("1", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        time.sleep(2)
        pg.press("return")
        time.sleep(2)
        pg.press("f4")
        time.sleep(2)
        FileName = "HousyuSaimoku.png"
        NoD = ImgCheck(FolURL2, "NoDataBar.png", 0.9, 5)
        while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
            time.sleep(2)
            NoD = ImgCheck(FolURL2, "NoDataBar.png", 0.9, 5)
            if NoD[0] is True:
                break
        if NoD[0] is True:
            ImgClick(FolURL2, "FullMenu.png", 0.9, 5)
            time.sleep(3)
            conf = 0.9
            # -------------------------------------------
            FileName = "F10END.png"
            F10E = ImgCheck(FolURL2, FileName, 0.9, 5)
            pyautogui.click(F10E[1] + 10, F10E[2] - 30)
            # -------------------------------------------
            pg.write("202", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
            pg.press("return")
            FileName = "KanyoItiWin.png"
            while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
                time.sleep(2)
            conf = 0.9
            LoopVal = 10
            FileName = "KanyoSyatyouTAB.png"
            ImgClick(FolURL2, FileName, conf, LoopVal)
            FileName = "SyatyouOpen.png"
            while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
                time.sleep(2)
            UpList.append([KamokuCD, xls_cd, xls_name, "データ無"])
            with open(
                FolURL2 + "/Log/請求入力フロー結果.csv",
                mode="w",
                encoding="shift-jis",
                errors="ignore",
            ) as f:
                pd.DataFrame(UpList).to_csv(f)
        else:
            conf = 0.9
            LoopVal = 10
            # -------------------------------------------
            FileName = "F8CAN.png"
            F10C = ImgCheck(FolURL2, FileName, 0.9, 5)
            pyautogui.click(F10C[1] + int(F10C[1] / 100), F10C[2] - int(F10C[2] / 20))
            # -------------------------------------------
            pg.write(KamokuCD, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
            pg.press("return")
            pg.press("return")
            time.sleep(4)
            if KamokuCD == "222":
                pg.keyDown("alt")
                pg.press("down")
                pg.keyUp("alt")
                pg.press("2")
                pg.press("return")
            else:
                pg.keyDown("alt")
                pg.press("down")
                pg.keyUp("alt")
                pg.press("4")
                pg.press("return")
            time.sleep(2)
            pg.press(
                [
                    "return",
                    "return",
                    "return",
                    "return",
                    "return",
                    "return",
                    "return",
                    "return",
                ]
            )  # 一巡目
            time.sleep(2)
            pg.press("return")
            # if xls_tx == xls_tx:
            #     print("nan")
            # else:
            pyperclip.copy(xls_tx)
            if xls_tx != "nan":
                pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
            pg.press("return")
            pg.write(xls_mn, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
            pg.press("return")
            time.sleep(2)
            pg.press("f4")
            time.sleep(2)
            pg.press("f4")
            time.sleep(2)
            FileName = "InputOK.png"
            while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
                time.sleep(2)
            pg.press("return")
            FileName = "KanyoItiWin.png"
            while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
                time.sleep(2)
            UpList.append([KamokuCD, xls_cd, xls_name, xls_mn])
            with open(
                FolURL2 + "/Log/請求入力フロー結果.csv",
                mode="w",
                encoding="shift-jis",
                errors="ignore",
            ) as f:
                pd.DataFrame(UpList).to_csv(f)
            # pd.DataFrame(UpList).to_csv(FolURL2 + "/Log/請求入力フロー結果.csv", encoding = "shift-jis")
    else:
        UpList.append([KamokuCD, xls_cd, xls_name, "失敗"])
        with open(
            FolURL2 + "/Log/請求入力フロー結果.csv",
            mode="w",
            encoding="shift-jis",
            errors="ignore",
        ) as f:
            pd.DataFrame(UpList).to_csv(f)


# ----------------------------------------------------------------------------------------------------------------------
def OuterAction(driver, FolURL2, xls_cd, xls_name, xls_mn, xls_tx, UpList, KamokuCD):
    FM = FindMenu(driver, FolURL2, xls_cd)
    if FM is True:
        time.sleep(2)
        # -------------------------------------------
        FileName = "F10Close.png"
        F10C = ImgCheck(FolURL2, FileName, 0.9, 5)
        pyautogui.click(F10C[1] + int(F10C[1] / 100), F10C[2] - int(F10C[2] / 20))
        # -------------------------------------------
        pg.write("1", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        time.sleep(2)
        pg.press("return")
        time.sleep(2)
        pg.press("f4")
        time.sleep(2)
        FileName = "HousyuSaimoku.png"
        while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
            time.sleep(2)
        # -------------------------------------------
        F10C = ImgCheck(FolURL2, FileName, 0.9, 5)
        FileName = "F8CAN.png"
        pyautogui.click(F10C[1] + int(F10C[1] / 100), F10C[2] - int(F10C[2] / 20))
        # -------------------------------------------
        pg.write(KamokuCD, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        pg.press("return")
        pg.press("return")
        time.sleep(4)
        if KamokuCD == "222":
            pg.keyDown("alt")
            pg.press("down")
            pg.keyUp("alt")
            pg.press("2")
            pg.press("return")
        else:
            pg.keyDown("alt")
            pg.press("down")
            pg.keyUp("alt")
            pg.press("4")
            pg.press("return")
        time.sleep(2)
        pg.press(
            [
                "return",
                "return",
                "return",
                "return",
                "return",
                "return",
                "return",
                "return",
            ]
        )  # 一巡目
        time.sleep(2)
        pg.press("return")
        # if xls_tx == xls_tx:
        #     print("nan")
        # else:
        pyperclip.copy(xls_tx)
        if xls_tx != "nan":
            pg.hotkey("ctrl", "v")  # pg日本語不可なのでコピペ
        pg.press("return")
        pg.write(xls_mn, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        pg.press("return")
        time.sleep(2)
        pg.press("f4")
        time.sleep(2)
        pg.press("f4")
        time.sleep(2)
        FileName = "InputOK.png"
        while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
            time.sleep(2)
        pg.press("return")
        FileName = "KanyoItiWin.png"
        while pg.locateOnScreen(FolURL2 + "/" + FileName, confidence=0.9) is None:
            time.sleep(2)
        UpList.append([KamokuCD, xls_cd, xls_name, xls_mn])
        with open(
            FolURL2 + "/Log/請求入力フロー結果.csv",
            mode="w",
            encoding="shift-jis",
            errors="ignore",
        ) as f:
            pd.DataFrame(UpList).to_csv(f)
        # pd.DataFrame(UpList).to_csv(FolURL2 + "/Log/請求入力フロー結果.csv", encoding = "shift-jis")
    else:
        UpList.append([KamokuCD, xls_cd, xls_name, "失敗"])
        with open(
            FolURL2 + "/Log/請求入力フロー結果.csv",
            mode="w",
            encoding="shift-jis",
            errors="ignore",
        ) as f:
            pd.DataFrame(UpList).to_csv(f)


# ----------------------------------------------------------------------------------------------------------------------
def MainFlow(FolURL2, xls_data, KamokuCD, Lyear, Lmonth, Lday, i, p, obj):
    BatUrl = os.getcwd() + r"/bat/AWADriverOpen.bat"  # 4724ポート指定でappiumサーバー起動バッチを開く
    driver = OMSOpen.MainFlow(BatUrl, openurl, "RPAPhoto", i, p)  # OMSを起動しログイン後インスタンス化

    if driver != "TimeOut":
        FMSOpen(driver, FolURL2, xls_data, KamokuCD, Lyear, Lmonth, Lday)
        UpList = []
        No = False
        first = False
        for index, xls_Item in xls_data.iterrows():
            xls_cd = str(xls_Item.values[int(obj.i_id_txt4.get())])
            xls_name = xls_Item.values[int(obj.i_name_txt.get())].replace("\u3000", "")
            xls_mn = str(xls_Item.values[int(obj.i_id_txt5.get())])
            xls_tx = str(obj.i_id_txt6.get())
            try:
                int(xls_cd)
                int(xls_mn)
                No = True
            except:
                No = False
            if No is True:
                if first is False:
                    FirstAction(
                        driver,
                        FolURL2,
                        xls_cd,
                        xls_name,
                        xls_mn,
                        xls_tx,
                        UpList,
                        KamokuCD,
                    )
                    first = True
                else:
                    OuterAction(
                        driver,
                        FolURL2,
                        xls_cd,
                        xls_name,
                        xls_mn,
                        xls_tx,
                        UpList,
                        KamokuCD,
                    )
        while pg.locateOnScreen(FolURL2 + r"\wait_menu.png", confidence=0.9) is None:
            time.sleep(2)
        pg.press("f10")
        while pg.locateOnScreen(FolURL2 + r"\CloseBtn.png", confidence=0.9) is None:
            time.sleep(2)
        ImgClick(FolURL2, r"\End_Btn.png", 0.9, 10)
        while pg.locateOnScreen(FolURL2 + r"\CloseBtn.png", confidence=0.9) is not None:
            time.sleep(2)
        print("処理終了")
        return True
    else:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def Main(obj):
    global openurl
    openurl, FolURL2, xls_data, KamokuCD, Lyear, Lmonth, Lday, i, p = (
        obj.open_dir,
        obj.Img_dir_D,
        obj.table.model.df,
        obj.code_txt.get(),
        obj.i_id_txt.get(),
        obj.i_id_txt2.get(),
        obj.i_id_txt3.get(),
        obj.id_txt.get(),
        obj.pass_txt.get(),
    )
    try:
        MF = MainFlow(FolURL2, xls_data, KamokuCD, Lyear, Lmonth, Lday, i, p, obj)
        if MF is True:
            return True
        else:
            return False
    except:
        traceback.print_exc()
        return False
