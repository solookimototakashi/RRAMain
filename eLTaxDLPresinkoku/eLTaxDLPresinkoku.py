###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# モジュールインポート
import subprocess
import pyautogui as pg
import time
import elTaxDLOpen

# pandasインポート
import pandas as pd

# 配列計算関数numpyインポート
import numpy as np

# osインポート
import os

# datetimeインポート
from datetime import datetime as dt

# 例外処理判定の為のtracebackインポート
import traceback
import calendar
import pyperclip
from collections import OrderedDict
import jaconv

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
            p = pg.locateOnScreen(ImgURL, confidence=conf)
            x, y = pg.center(p)
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
            p = pg.locateOnScreen(ImgURL, confidence=conf)
            x, y = pg.center(p)
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
                p = pg.locateOnScreen(ImgURL, confidence=conf)
                x, y = pg.center(p)
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
                    p = pg.locateOnScreen(ImgURL, confidence=conf)
                    x, y = pg.center(p)
                    pg.click(x, y)
                    time.sleep(1)
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
            time.sleep(1)
    except:
        print("待機終了")


# ----------------------------------------------------------------------------------------------------------------------
def SortPDF(PDFName):
    Fol = str(dt.today().year) + "-" + str(dt.today().month)
    pt = "\\\\nas-sv\\B_監査etc\\B2_電子ﾌｧｲﾙ\\ﾒｯｾｰｼﾞﾎﾞｯｸｽ" + Fol + "\\送信分受信通知"
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
def DeleteData(FolURL2, Cou):  # 初めの画面で引数Cou分削除操作繰り返し
    FileName = "DataSelectCombo.png"
    conf = 0.9
    LoopVal = 10
    ImgClick(FolURL2, FileName, conf, LoopVal)
    for x in range(Cou):
        pg.press("down")
        pg.press("return")
        pg.press(["tab", "tab", "tab", "tab"])
        pg.press("return")
        pg.press("y")
        pg.press("tab")


# ----------------------------------------------------------------------------------------------------------------------
def ReturnPar(FolURL2, Loop_Code, Loop_Name, MasterCSV):
    MRow = np.array(MasterCSV).shape[0]  # 配列行数取得
    for x in range(MRow):
        MDataRow = MasterCSV.iloc[x, :]
        MCode = str(MDataRow["SyanaiCode"])
        if str(Loop_Code) == MCode:
            return (
                str(Loop_Code),
                Loop_Name,
                MDataRow["TKCKokuzeiUserCode"],
                MDataRow["TKCTihouzeiUserID"],
                MDataRow["MirokuKokuzeiUserCode"],
                MDataRow["MirokuTihouzeiUserID"],
                MDataRow["etaxPass"],
                MDataRow["eltaxPass"],
            )


# ----------------------------------------------------------------------------------------------------------------------
def NewEnt(FolURL2, MasterPar):
    try:
        conf = 0.9
        LoopVal = 100
        List = ["NewEntBtn.png", "NewEntBtn2.png"]
        ILA = ImgCheckForList(FolURL2, List, conf)
        SCode = MasterPar[0]
        TName = MasterPar[1]
        Teltax = MasterPar[3]
        Meltax = MasterPar[5]
        if ILA[0] is True:
            ImgClick(FolURL2, ILA[1], conf, LoopVal)
            while pg.locateOnScreen(FolURL2 + "NewEntWin.png", confidence=0.9) is None:
                time.sleep(1)
            try:
                pg.write(Teltax, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
                pg.press("tab")
            except:
                pg.write(Meltax, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
                pg.press("tab")
            try:
                EntName = jaconv.h2z(
                    str(SCode + "_" + TName), digit=True, ascii=True, kana=True
                )
                pyperclip.copy(EntName)
                pg.hotkey("ctrl", "v")
                pg.press("tab")
            except:
                pg.write("ななし", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
            ImgClick(FolURL2, "NewEntEndBtn.png", conf, LoopVal)
            while (
                pg.locateOnScreen(FolURL2 + "NewEntEndCheck.png", confidence=0.9)
                is None
            ):
                time.sleep(1)
            pg.press("return")
            time.sleep(1)
        conf = 0.9
        LoopVal = 100
        List = ["DataSelectCombo.png", "DataSelectCombo2.png"]
        DSC = ImgCheckForList(FolURL2, List, conf)
        if DSC[0] is True:
            ImgClick(FolURL2, DSC[1], conf, LoopVal)
        pg.press("down")
        pg.press("return")
        pg.press("tab")
        pg.press("return")
        while pg.locateOnScreen(FolURL2 + "MainWin.png", confidence=0.9) is None:
            time.sleep(1)
            return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def MsgOpenAction(FolURL2, MasterPar):
    try:
        etaxPass = MasterPar[6]
        eltaxPass = MasterPar[7]
        conf = 0.9
        LoopVal = 10
        FileName = "MsgCheck.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        FileName = "MsgOpenBtn.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        while pg.locateOnScreen(FolURL2 + "PotalLogMenu.png", confidence=0.9) is None:
            time.sleep(1)
        FileName = "PassTxtBox.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        try:
            pg.write(eltaxPass, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        except:
            pg.write(etaxPass, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        FileName = "MsgLoginBtn.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        while pg.locateOnScreen(FolURL2 + "MsgFirstWin.png", confidence=0.9) is None:
            time.sleep(1)
        return True
    except:
        return False


# ----------------------------------------------------------------------------------------------------------------------
def MsgAction(FolURL2, TaisyouNen, TaisyouTuki, MasterPar):
    try:
        DateF = str(TaisyouNen) + str("{0:02d}".format(int(TaisyouTuki))) + "01"
        Lday = calendar.monthrange(int(TaisyouNen), int(TaisyouTuki))
        DateL = str(TaisyouNen) + str("{0:02d}".format(int(TaisyouTuki))) + str(Lday[1])
        conf = 0.9
        LoopVal = 10
        FileName = "DateTimeBox.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        pg.press("delete")
        pg.write(DateF, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        pg.press("return")
        pg.press(["tab", "tab", "tab"])
        pg.press("delete")
        pg.write(DateL, interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        pg.press("return")
        FileName = "MsgFindBtn.png"
        ImgClick(FolURL2, FileName, conf, LoopVal)
        while pg.locateOnScreen(FolURL2 + "MsgWaitBar.png", confidence=0.9) is None:
            time.sleep(1)
        return True
    except:
        return False


def MainFlow(FolURL2, PreList, MasterCSV, NoList):
    BatUrl = FolURL2 + "/bat/AWADriverOpen.bat"  # 4724ポート指定でappiumサーバー起動バッチを開く
    elTaxDLOpen.MainFlow(BatUrl, FolURL2, "RPAPhoto")  # OMSを起動しログイン後インスタンス化
    FolURL2 = FolURL2 + "/RPAPhoto/eLTaxDLPresinkoku/"
    for NoListItem in NoList:
        LoopList = []
        for PreListItem in PreList:  # PreListItem[0]=URL,PreListItem=[1]=関与先コード
            if NoListItem == PreListItem[1]:
                LoopList.append(PreListItem)
    for LoopListItem in LoopList:
        Loop_Code = LoopListItem[1]
        Loop_Row = int(int(LoopListItem[2]) - 1)
        Loop_Name = LoopListItem[4]
        MasterPar = ReturnPar(FolURL2, Loop_Code, Loop_Name, MasterCSV)
        NEF = NewEnt(FolURL2, MasterPar)
        if NEF is True:
            MOA = MsgOpenAction(FolURL2, MasterPar)
            if MOA is True:
                MA = MsgAction(FolURL2, TaisyouNen, TaisyouTuki, MasterPar)
                if MA is True:
                    conf = 0.9
                    LoopVal = 10
                    FileName = "MsgWaitBar.png"
                    RowsTarget = ImgCheck(FolURL2, FileName, conf, LoopVal)
                    xpos = RowsTarget[1]
                    ypos = RowsTarget[2]
                    xpos = xpos
                    ypos = ypos + 30
                    RowC = 25 * Loop_Row
                    ypos = ypos + RowC
                    pg.click(xpos, ypos)
                    time.sleep(1)


# RPA用画像フォルダの作成---------------------------------------------------------
FolURL2 = os.getcwd().replace("\\", "/")
# 既定のプリンターをMSPDFへ---------------------------------------------------------
PT = (
    os.getcwd().replace("\\", "/")
    + "/PowerShellMyScripts/DefaultPrinterChangeMSPDF.ps1"
)
proc = subprocess.call("powershell.exe -File " + PT)
# --------------------------------------------------------------------------------
TaisyouNen = input("対象[年]を西暦で入力してください。\n")
TaisyouTuki = input("対象[月]を西暦で入力してください。\n")
TaisyouFol = str(TaisyouNen) + "-" + str(TaisyouTuki)
# プレ申告のお知らせ保管フォルダチェック---------------------------------------------------------
Fol = TaisyouFol
pt = "\\\\nas-sv\\B_監査etc\\B2_電子ﾌｧｲﾙ\\ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\" + Fol + "\\eLTAX"
# path = path.replace('\\','/')#先
PDFFileList = os.walk(pt)
Cou = 1
PreList = []
NgLog = pd.read_csv(
    FolURL2 + "/RPAPhoto/eLTaxDLPresinkoku/Log/Log.csv", encoding="utf-8"
)
NgRow = np.array(NgLog).shape[0]  # 配列行数取得
NgCol = np.array(NgLog).shape[1]  # 配列列数取得

for current_dir, sub_dirs, files_list in PDFFileList:
    Count_dir = 0
    for file_name in files_list:
        if "プレ申告のお知らせ" in file_name or "プレ申告データに関するお知らせ" in file_name:
            Count_dir = Count_dir + 1
    for file_name in files_list:
        if "プレ申告のお知らせ" in file_name or "プレ申告データに関するお知らせ" in file_name:
            Nos = file_name.split("_")
            FolName = current_dir.split("_")
            FolName = FolName[1]
            NewTitle = os.path.join(current_dir, file_name)
            NewTitle = NewTitle.split("プレ申告データ")
            NewTitle = NewTitle[0] + "プレ申告データ印刷結果.pdf"
            # NGList = ["100","105","106","107","108","121","12","148","183","200","201","204","207","209","221",\
            #    "223","240","249","251","268","282","285","305","306","309","317"]
            NoF = True
            for x in range(NgRow):
                NgDataRow = NgLog.iloc[x, :]
                NgCodeCode = str(NgDataRow[1])
                if not Nos[0] == NgCodeCode:
                    NoF = True
                else:
                    NoF = False
                    break
            if NoF is True:
                PreList.append(
                    [
                        os.path.join(current_dir, file_name),
                        int(Nos[0]),
                        Count_dir,
                        NewTitle,
                        FolName,
                    ]
                )
print(NgLog)
print(PreList)

myList = []
for PreListItem in PreList:
    myList.append(PreListItem[1])
NoList = list(OrderedDict.fromkeys(myList))
print(NoList)
MasterCSV = pd.read_csv(
    FolURL2 + "/RPAPhoto/TKC_PreSinkokuDown/" + "MasterDB.csv",
    dtype={
        "TKCKokuzeiUserCode": str,
        "TKCTihouzeiUserID": str,
        "MirokuKokuzeiUserCode": str,
        "MirokuTihouzeiUserID": str,
        "etaxPass": str,
        "eltaxPass": str,
    },
)
print(MasterCSV)
# --------------------------------------------------------------------------------
try:
    MainFlow(FolURL2, PreList, MasterCSV, NoList)
except:
    traceback.print_exc()
