###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# モジュールインポート
# from appium import webdriver
import subprocess
import pyautogui as pg
import time
import wrapt_timeout_decorator

# loggerインポート
from logging import getLogger

logger = getLogger()

TIMEOUT = 100
# ----------------------------------------------------------------------------------------------------------------------
def ExeOpen(AppURL):  # URL指定でアプリ起動関数
    P = subprocess.Popen(AppURL)
    return P


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitXPATH(UIPATH, driver):  # XPATH要素を取得するまで待機
    for x in range(1000000):
        try:
            driver.find_element_by_xpath(UIPATH)
            return True
        except:
            return False


# ----------------------------------------------------------------------------------------------------------------------
def DriverUIWaitAutomationId(UIPATH, driver):  # XPATH要素を取得するまで待機
    for x in range(1000000):
        try:
            driver.find_element_by_accessibility_id(UIPATH)
            return True
        except:
            return False


# ----------------------------------------------------------------------------------------------------------------------
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
    for x in range(100):
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
        return False


# ----------------------------------------------------------------------------------------------------------------------
def ImgClick(FolURL2, FileName, conf, LoopVal):  # 画像があればクリックしてx,y軸を返す
    ImgURL = FolURL2 + "/" + FileName
    for x in range(10000):
        if (
            ImgCheck(FolURL2, FileName, conf, LoopVal)[0] is True
        ):  # OMSメニューの年調起動ボタンを判定して初期処理分け
            # 正常待機後処理
            for y in range(10000):
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
@wrapt_timeout_decorator.timeout(dec_timeout=TIMEOUT)
def Flow(BatUrl, FolURL2, ImgFolName):
    # WebDriver起動バッチを管理者権限で起動---------------------------------------------------------------------------------
    logger.debug("Bat起動: debug level log")
    MSPDFURL = FolURL2 + "/bat/MSPDFSet.bat"  # 規定プリンターをMSPDFに
    ExeOpen(MSPDFURL)
    # ExeOpen(BatUrl)
    # desired_caps = {}
    # desired_caps["app"] = "Root"  # Rootを指定してDriverTargetをデスクトップに
    # logger.debug("Appiumサーバー起動: debug level log")
    # driver = webdriver.Remote(
    #     "http://127.0.0.1:4724", desired_caps, direct_connection=True
    # )  # ポート指定してDriverインスタンス化
    # ----------------------------------------------------------------------------------------------------------------------
    # MJSを起動-------------------------------------------------------------------------------------------------------------
    logger.debug("MJS起動: debug level log")
    try:
        MJSURL = r"C:\Program Files (x86)\MJS\MJSNXSVA\MJSDesktopNX.exe"
        app = ExeOpen(MJSURL)
        return app
    except:
        MJSURL = r"C:\Program Files (x86)\MJS\MJSNXSVB\MJSDesktopNX.exe"
        app = ExeOpen(MJSURL)
        return app


@wrapt_timeout_decorator.timeout(dec_timeout=TIMEOUT)
def tryFlow(app, ImgFolName):
    # 画像が出現するまで待機-------------------------------------------------------------------------------------------
    ImgFolName = ImgFolName + r"\MJSOpen"
    List = [r"\PassTxtBox.png", r"\PassTxtBox2.png"]
    conf = 0.9  # 画像認識感度
    LoopVal = 10000  # 検索回数
    ListCheck = ImgCheckForList(ImgFolName, List, conf)
    if ListCheck[0] is True:
        logger.debug("Pass入力開始: debug level log")
        MLI = ImgCheck(ImgFolName, r"\MyLogIn.png", 0.9, 10)
        if MLI[0] is True:
            ImgClick(ImgFolName, ListCheck[1], conf, LoopVal)  # 電子申告・申請タブを押す
            pg.write("051210561111111", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        else:
            LB = ImgCheckForList(ImgFolName, [r"\LoginBox.png", r"\LoginBox2.png"], 0.9)
            ImgClick(ImgFolName, LB[1], 0.9, 10)  # 電子申告・申請タブを押す
            pg.write("561", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
            ImgClick(ImgFolName, ListCheck[1], conf, LoopVal)  # 電子申告・申請タブを押す
            pg.write("051210561111111", interval=0.01)  # 直接SENDできないのでpyautoguiで入力
        ImgClick(ImgFolName, r"\LoginOKBtn.png", conf, LoopVal)  # 電子申告・申請タブを押す
        time.sleep(1)
        IC = ImgCheck(
            ImgFolName, r"\MJSOsiraseClose.png", conf, LoopVal
        )  # お知らせ画面があるかチェック
        if IC[0] is True:
            ImgClick(
                ImgFolName, r"\MJSOsiraseClose.png", conf, LoopVal
            )  # お知らせ画面があれば閉じるボタンをクリック
            return app  # driver
        else:
            logger.debug("MJSログイン完了: debug level log")
            return app  # driver
    time.sleep(1)
    # ----------------------------------------------------------------------------------------------------------------------


def MainFlow(BatUrl, FolURL2, ImgFolName):
    try:
        app = Flow(BatUrl, FolURL2, ImgFolName)
        f_app = tryFlow(app, ImgFolName)
        return f_app
    except TimeoutError:
        if app is not None:
            time.sleep(1)  # 子プロセスが起動していることを確認するまでの時間を確保
            print("killします")
            killcmd = "taskkill /F /PID {pid} /T".format(pid=app.pid)
            subprocess.run(killcmd, shell=True)
        return "TimeOut"
