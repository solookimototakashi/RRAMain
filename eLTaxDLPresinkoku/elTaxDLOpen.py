###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# モジュールインポート
from appium import webdriver
import subprocess
import pyautogui as pg
import time
import os

# ----------------------------------------------------------------------------------------------------------------------
def ExeOpen(AppURL):  # URL指定でアプリ起動関数
    subprocess.Popen(AppURL)


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitXPATH(UIPATH, driver):  # XPATH要素を取得するまで待機
    for x in range(1000000):
        try:
            driver.find_element_by_xpath(UIPATH)
            return True
        except Exception as e:
            print(e)
            return False


# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitAutomationId(UIPATH, driver):  # XPATH要素を取得するまで待機
    for x in range(1000000):
        try:
            driver.find_element_by_accessibility_id(UIPATH)
            return True
        except Exception as e:
            print(e)
            return False


# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
def ImgClick(FolURL2, FileName, conf, LoopVal):  # 画像があればクリックしてx,y軸を返す
    ImgURL = FolURL2 + "/" + FileName
    for x in range(10):
        for y in range(10):
            try:
                p = pg.locateOnScreen(ImgURL, confidence=conf)
                x, y = pg.center(p)
                pg.click(x, y)
                time.sleep(1)
                return x, y
            except:
                print("失敗")


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
def MainFlow(BatUrl, FolURL2, ImgFolName):
    # WebDriver起動バッチを管理者権限で起動---------------------------------------------------------------------------------
    ExeOpen(BatUrl)
    desired_caps = {}
    desired_caps["app"] = "Root"  # Rootを指定してDriverTargetをデスクトップに
    driver = webdriver.Remote(
        "http://127.0.0.1:4724", desired_caps, direct_connection=True
    )  # ポート指定してDriverインスタンス化

    # ----------------------------------------------------------------------------------------------------------------------
    # elTaxを起動-------------------------------------------------------------------------------------------------------------
    elTaxURL = r"C:\Program Files (x86)\LT\LTN\BIN\LtnMain.exe"
    ExeOpen(elTaxURL)
    # time.sleep(10)
    FolURL2 = os.getcwd().replace("\\", "/") + "/RPAPhoto/elTaxDLOpen/"
    FileName = "VCheck.png"
    while pg.locateOnScreen(FolURL2 + FileName, confidence=0.9) is None:
        time.sleep(1)
        conf = 0.9
        LoopVal = 100
        OF = ImgCheck(FolURL2, "OpenWindow.png", conf, LoopVal)
        if OF[0] is True:
            break
    if OF[0] is True:
        FileName = "OpenWindow.png"
        while pg.locateOnScreen(FolURL2 + FileName, confidence=0.9) is None:
            time.sleep(1)
        print("起動しました。")
        return driver
    else:
        FileName = "VCheckNext.png"
        conf = 0.9
        LoopVal = 100
        ImgClick(FolURL2, FileName, conf, LoopVal)
        # ----------------------------------------------------------------------------------------------------------------------
        FileName = "OpenWindow.png"
        while pg.locateOnScreen(FolURL2 + FileName, confidence=0.9) is None:
            time.sleep(1)
        print("起動しました。")
        return driver
    # ----------------------------------------------------------------------------------------------------------------------
