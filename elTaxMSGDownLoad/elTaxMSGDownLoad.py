###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# pandasインポート
import pandas as pd

# 配列計算関数numpyインポート
import numpy as np

# timeインポート
import time

# seleniumインポート
from selenium import webdriver
from selenium.webdriver.chrome.options import Options  # ブラウザオプションを与える
from selenium.webdriver.support.ui import WebDriverWait  # 読込待機コマンドを与える
from selenium.webdriver.support import expected_conditions as EC  # 読込待機コマンドに条件式を与える
from webdriver_manager.chrome import ChromeDriverManager

# jsonインポート
import json

# osインポート
import os

# datetimeインポート
from datetime import datetime as dt

# 日付加減算インポート
from dateutil.relativedelta import relativedelta

# glob(フォルダファイルチェックコマンド)インポート
import glob

# shutil(フォルダファイル編集コマンド)インポート
import shutil

# 例外処理判定の為のtracebackインポート
import traceback

# logger設定------------------------------------------------------------------------------------------------------------
import logging.config

logging.config.fileConfig(r"LogConf\logging_debugelTaxLog.conf")
logger = logging.getLogger(__name__)
# ----------------------------------------------------------------------------------------------------------------------
class elTaxWeb:
    """
    概要: Webページクラス
    """

    def __init__(self, H_dfDataRow, InputYear, InputMonth):
        self.OKLog = []
        self.NGLog = []
        self.pyoklog = r"//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/PyOKLog.csv"
        self.pynglog = r"//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/PyNGLog.csv"
        # 入力日付をDate型に変換------------------------------------------------
        InputYear = InputYear.zfill(4)
        InputMonth = InputMonth.zfill(2)
        ID = InputYear + "/" + InputMonth + "/01 01:01:01"
        self.InputDate = dt.strptime(ID, "%Y/%m/%d %H:%M:%S")
        self.H_Today = self.InputDate + relativedelta(months=-0)  # 本日の日付
        self.H_dtToday = self.InputDate + relativedelta(months=+1)  # 1月後の日付
        # --------------------------------------------------------------------
        self.H_SCode = H_dfDataRow["SyanaiCode"]  # 関与先コード
        try:
            self.H_TKCName = H_dfDataRow["TKCName"].replace(
                "\u3000", " "
            )  # 空白\u3000を置換
        except:
            self.H_TKCName = H_dfDataRow["TKCName"]
        self.H_MSikibetu = str(H_dfDataRow["MirokuTihouzeiUserID"])
        self.H_TSikibetu = str(H_dfDataRow["TKCTihouzeiUserID"])
        # --------------------------------------------------------------------
        try:
            self.H_ePass = str(H_dfDataRow["etaxPass"]).replace(
                "\u3000", " "
            )  # 空白\u3000を置換
        except:
            self.H_ePass = str(H_dfDataRow["etaxPass"])
        try:
            self.H_elPass = str(H_dfDataRow["eltaxPass"]).replace(
                "\u3000", " "
            )  # 空白\u3000を置換
        except:
            self.H_elPass = str(H_dfDataRow["eltaxPass"])
        # --------------------------------------------------------------------
        logger.debug("WebDriver起動")
        # 作成した変数は頭にH_をつける事
        # ブラウザ閲覧時のオプションを指定するオブジェクト"options"を作成
        self.H_fold = os.getcwd().replace("\\", "/")
        self.H_options = Options()
        self.appState = {
            "recentDestinations": [
                {
                    "id": "Save as PDF",
                    "origin": "local",
                    "account": "",
                }
            ],
            "selectedDestinationId": "Save as PDF",
            "version": 2,
        }
        self.prefs = {
            "printing.print_preview_sticky_settings.appState": json.dumps(
                self.appState
            ),
            "savefile.default_directory": self.H_fold,  # "D:/PythonScript"
        }
        self.H_options.add_experimental_option("prefs", self.prefs)
        # 必要に応じてオプションを追加
        self.H_options.add_argument("--window-size=1024,768")
        self.H_options.add_argument("--kiosk-printing")
        # ドライバのpathを指定
        self.H_path = ChromeDriverManager().install()  # self.H_fold + "/chromedriver"
        # WEBURLの指定
        self.H_WEBurl = "https://www.portal.eltax.lta.go.jp/apa/web/webindexb#eLTAX"
        # ブラウザのウィンドウを表すオブジェクト"driver"を作成
        self.H_driver = webdriver.Chrome(
            executable_path=self.H_path, chrome_options=self.H_options
        )
        self.H_driver.get(self.H_WEBurl)
        WebDriverWait(self.H_driver, 30).until(
            EC.presence_of_all_elements_located
        )  # 要素が読み込まれるまで最大30秒待つ
        # 初めの拡張機能ポップアップを閉じる---------------------------------------------
        time.sleep(1)
        self.H_PopupCheck_btn = self.H_driver.find_element_by_xpath(
            "/html/body/div[3]/div[2]/div/form/div/div[1]/input"
        )
        self.H_driver.execute_script("arguments[0].click();", self.H_PopupCheck_btn)
        self.H_PopupClose_btn = self.H_driver.find_element_by_xpath(
            "/html/body/div[3]/div[2]/div/form/div/div[2]/div[1]/a"
        )  # 閉じるボタンの要素指定
        self.H_PopupClose_btn.click()
        # ---------------------------------------------------------------------------
        WebDriverWait(self.H_driver, 30).until(
            EC.presence_of_all_elements_located
        )  # 要素が読み込まれるまで最大30秒待つ
        self.LoginFlag = 1  # ログイン回数フラグ
        SWL = self.WebLogin()
        if SWL is True:
            # ログイン1回目成功
            logger.debug(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "DL処理完了")
            print(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "DL処理完了")
            self.LoginFlag = 0  # ログイン回数フラグ
        else:
            self.LoginFlag += 1  # ログイン回数フラグ加算
            self.H_driver = webdriver.Chrome(
                executable_path=self.H_path, chrome_options=self.H_options
            )
            self.H_driver.get(self.H_WEBurl)
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            # 初めの拡張機能ポップアップを閉じる---------------------------------------------
            time.sleep(1)
            self.H_PopupCheck_btn = self.H_driver.find_element_by_xpath(
                "/html/body/div[3]/div[2]/div/form/div/div[1]/input"
            )
            self.H_driver.execute_script("arguments[0].click();", self.H_PopupCheck_btn)
            self.H_PopupClose_btn = self.H_driver.find_element_by_xpath(
                "/html/body/div[3]/div[2]/div/form/div/div[2]/div[1]/a"
            )  # 閉じるボタンの要素指定
            self.H_PopupClose_btn.click()
            # ---------------------------------------------------------------------------
            SWL = self.WebLogin()
            if SWL is True:
                # ログイン2回目成功
                logger.debug(
                    str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "DL処理完了"
                )
                print(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "DL処理完了")
                self.LoginFlag = 0  # ログイン回数フラグ
            else:
                # ログイン2回目失敗
                logger.debug(
                    str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "ログイン2回目失敗"
                )
                print(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_" + "ログイン2回目失敗")
                self.LoginFlag = 0  # ログイン回数フラグ

    # -----------------------------------------------------------------------------------
    def WebLogin(self):
        # ログイン画面要素指定-------------------------------------------------------------
        # 利用者識別番号入力欄をxpathで取得
        self.H_Sikibetu_box1 = self.H_driver.find_element_by_xpath(
            "/html/body/div[2]/form/div/div/div[1]/div/ul[1]/li[1]/input"
        )
        # ログインPASS入力欄をxpathで取得
        self.H_LogPassBox = self.H_driver.find_element_by_xpath(
            "/html/body/div[2]/form/div/div/div[1]/div/ul[1]/li[2]/input"
        )
        self.H_Log_btn = self.H_driver.find_element_by_xpath(
            "/html/body/div[2]/form/div/div/div[1]/div/button"
        )
        # 識別番号とパスのnan判定-----------------------------------------------------------
        if self.H_ePass != "nan" and self.H_elPass != "nan":
            if self.H_MSikibetu == "nan":
                if self.H_ePass == "nan":
                    if self.LoginFlag == 1:
                        LL = self.LoginLoop(self.H_TSikibetu, self.H_elPass)
                    elif self.LoginFlag == 2:
                        LL = self.LoginLoop(self.H_TSikibetu, self.H_ePass)
                else:
                    if self.LoginFlag == 1:
                        LL = self.LoginLoop(self.H_TSikibetu, self.H_ePass)
                    elif self.LoginFlag == 2:
                        LL = self.LoginLoop(self.H_TSikibetu, self.H_elPass)
            else:
                if self.H_ePass == "nan":
                    if self.LoginFlag == 1:
                        LL = self.LoginLoop(self.H_MSikibetu, self.H_elPass)
                    elif self.LoginFlag == 2:
                        LL = self.LoginLoop(self.H_MSikibetu, self.H_ePass)
                else:
                    if self.LoginFlag == 1:
                        LL = self.LoginLoop(self.H_MSikibetu, self.H_ePass)
                    elif self.LoginFlag == 2:
                        LL = self.LoginLoop(self.H_MSikibetu, self.H_elPass)
            # ----------------------------------------------------------------------------
            if LL is True:
                # ログイン成功
                return True
            else:
                # ログイン失敗
                self.H_driver.close()
                self.H_driver.quit()
                return False
        else:
            # ログイン失敗
            self.H_driver.close()
            self.H_driver.quit()
            return False

    # ------------------------------------------------------------------------------------------
    def LoginLoop(self, H_id, H_pass):
        """
        概要: Webページログイン入力処理
        """
        try:
            self.H_id = str(H_id)  # 識別番号の変数代入
            self.H_pass = str(H_pass)
            # WEB操作---------------------------------------------------------------------------
            self.H_Sikibetu_box1.send_keys(self.H_id)
            self.H_LogPassBox.send_keys(self.H_pass)
            self.H_Log_btn.click()
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            try:
                self.H_H1 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[3]/form/div/div/div[1]/div/p"
                )  # ログイン確認用にH1要素を取得
                return False
            except Exception:
                try:
                    self.H_H1 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/div/div/div/div/div[1]/div[3]/ul/li[1]/a"
                    )  # ログイン確認用にH1要素を取得
                    SLR = self.LogReturn()
                    if SLR is True:
                        return True
                    else:
                        return False
                except:
                    return False
        except:
            return False

    # ---------------------------------------------------------------------------------------------
    def LogOnOuter(self):
        """
        検索結果のテーブル要素を取得
        """
        # ログイン成功の場合
        # WEB画面要素指定------------------------------------------------------------------------------------------------------
        try:
            self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[2]/div[1]"
            )  # メッセージボックスのテーブルをxpath取得
            # /html/body/div[2]/form/div/div/div[2]/div[2]/div[1]/div[2]/table
        except Exception:
            pass
        else:
            self.dfs = pd.read_html(
                self.H_driver.page_source, encoding="cp932"
            )  # pandasにWEBURLをぶっこむ
            self.H_MSG_row = len(self.dfs[2])  # テーブル行数を代入
            return True

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LogReturn(self):
        """
        ログイン後の処理
        """
        if "ログインしている利用者あてのメッセージを照会します。" in self.H_H1.text:  # ログイン成功の場合
            # WEB画面要素指定------------------------------------------------------------------------------------------------------
            LRS = self.LogReturnSub()
            if LRS is True:
                LOO = self.LogOnOuter()
                if LOO is True:
                    self.LoginLoopSub()
                    return True
            else:
                return False
        elif "認証エラー" in self.H_H1.text:  # ログイン失敗の場合
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            return "認証エラー"
        elif "ログインできませんでした。" in self.H_H1.text:  # ログイン失敗の場合
            return "ログインできませんでした。"
        elif "メッセージ照会" in self.H_H1.text:
            LRS = self.LogReturnSub()
            if LRS is True:
                LOO = self.LogOnOuter()
                if LOO is True:
                    self.LoginLoopSub()
                    return True
            else:
                return False
        else:  # その他の場合
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LogReturnSub(self):
        try:
            try:
                self.H_MSG_OpenMyNo = self.H_driver.find_element_by_xpath(
                    "/html/body/div[1]/form/div/div/div/div/div[1]/div[3]/ul/li[1]/a"
                )  # メッセージ照会ボタン
            except:
                self.H_MSG_OpenMyNo = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div/div/div[1]/div[3]/ul/li[1]/a"
                )  # メッセージ照会ボタン
            self.H_MSG_OpenMyNo.click()  # メッセージ照会ボタン押下
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            self.Hj = str(self.H_dtToday.year)  # 1月後の日付の年を文字列で代入
            self.Hjj = str("{0:02}".format(self.H_dtToday.month))  # 1月後の日付の月を文字列で代入
            self.H_Str = self.Hj + "/" + self.Hjj + "/01 01:01:01"  # 1月後のモデル日付を作成
            self.H_LT = dt.strptime(self.H_Str, "%Y/%m/%d %H:%M:%S")  # 1月後のモデル日付を日付型に代入
            self.H_LastToday = self.H_LT + relativedelta(days=-1)  # 1月後の月初日から1日前
            self.H_FirstMonth = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[1]/div/div[2]/input[1]"
            )  # 絞込み条件の開始日付ボックス
            self.H_EndMonth = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[1]/div/div[2]/input[4]"
            )  # 絞込み条件の終了日付ボックス
            self.H_FMon = (
                "{0:04}".format(self.H_Today.year)
                + "{0:02}".format(self.H_Today.month)
                + "{0:02}".format(1)
            )  # 開始日の日付文字列
            self.H_EMon = (
                "{0:04}".format(self.H_LastToday.year)
                + "{0:02}".format(self.H_LastToday.month)
                + "{0:02}".format(self.H_LastToday.day)
            )  # 終了日の日付文字列
            self.H_FirstMonth.send_keys(self.H_FMon)  # 絞込み条件の開始日付ボックスへ開始日の日付入力
            self.H_EndMonth.send_keys(self.H_EMon)  # 絞込み条件の終了日付ボックスへ終了日の日付入力
            self.H_KensakuBtn = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[1]/div/input"
            )  # 検索開始ボタン
            self.H_KensakuBtn.click()  # 検索開始ボタンクリック
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            try:
                self.H1SText = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/p[2]/span"
                ).text  # 検索結果テキストを抽出
                if "該当するデータはありませんでした。" == self.H1SText:
                    logger.debug(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_作成")
                    print(str(self.H_SCode) + "_" + str(self.H_TKCName) + "該当データ無し。")
                    return False
                else:
                    return True
            except:
                self.H1SText = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/p"
                ).text  # 検索結果テキストを抽出
                if "該当するデータはありませんでした。" == self.H1SText:
                    logger.debug(str(self.H_SCode) + "_" + str(self.H_TKCName) + "_作成")
                    print(str(self.H_SCode) + "_" + str(self.H_TKCName) + "該当データ無し。")
                    return False
                else:
                    return True
        except:
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LoginLoopSub(self):
        for H_MSG_rowItem in range(self.H_MSG_row):  # WEBテーブル行数分ループ
            self.H_Parameters = self.ParGet(H_MSG_rowItem)  # WEBテーブルから要素取得
            self.H_row = str(H_MSG_rowItem + 1)  # 重複タイトルの場合の末尾文字列
            SPDF = self.SortPDF()
            if SPDF is False or SPDF is None:
                logger.debug(
                    str(self.H_MSG_TableItem3)
                    + "_"
                    + str(self.H_MSG_TableItem4)
                    + "_"
                    + str(self.H_SCode)
                    + "_"
                    + str(self.H_TKCName)
                    + "_"
                    + str(self.H_MSG_TableItem1)
                    + "_"
                    + str(self.H_MSG_TableItem2)
                    + "_"
                    + str(self.H_row)
                    + "_"
                    + "未ダウンロード"
                )
                PIFS = self.PrintIFS()  # メッセージボックスを開き中身を印刷
                if PIFS is True:
                    RPDF = self.RenamePDF()  # PDF保存先フォルダー作成後リネーム&移動
                    if RPDF is True:
                        if H_MSG_rowItem == self.H_MSG_row - 1:
                            self.H_driver.close()
                            self.H_driver.quit()
                            time.sleep(1)
                        else:
                            self.H_driver.switch_to.window(
                                self.H_driver.window_handles[0]
                            )  # タブ移動する
                            WebDriverWait(self.H_driver, 30).until(
                                EC.presence_of_all_elements_located
                            )  # 要素が読み込まれるまで最大30秒待つ
                    else:
                        # リネーム&移動失敗
                        logger.debug(
                            str(self.H_MSG_TableItem3)
                            + "_"
                            + str(self.H_MSG_TableItem4)
                            + "_"
                            + str(self.H_SCode)
                            + "_"
                            + str(self.H_TKCName)
                            + "_"
                            + str(self.H_MSG_TableItem1)
                            + "_"
                            + str(self.H_MSG_TableItem2)
                            + "_"
                            + str(self.H_row)
                            + "_"
                            + "リネーム&移動失敗。"
                        )
                        self.H_driver.close()
                        self.H_driver.quit()
                        time.sleep(1)
                else:
                    self.H_driver.close()
                    self.H_driver.quit()
                    time.sleep(1)
            else:
                logger.debug(
                    str(self.H_MSG_TableItem3)
                    + "_"
                    + str(self.H_MSG_TableItem4)
                    + "_"
                    + str(self.H_SCode)
                    + "_"
                    + str(self.H_TKCName)
                    + "_"
                    + str(self.H_MSG_TableItem1)
                    + "_"
                    + str(self.H_MSG_TableItem2)
                    + "_"
                    + str(self.H_row)
                    + "_"
                    + "ファイルが存在します。"
                )
                print("ファイルが存在します。")
                if H_MSG_rowItem == self.H_MSG_row - 1:
                    self.H_driver.close()
                    self.H_driver.quit()
                    time.sleep(1)
                else:
                    self.H_BackBtn = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/footer/div[1]/div/div[1]/a"
                    )
                    self.H_BackBtn.click()
                    time.sleep(1)

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def ParGet(self, H_MSG_rowItem):
        """
        検索結果テーブル要素を1行ずつ変数代入
        """
        try:
            # WEBTable
            self.H_FirstMSG_Cbx = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[{}]/td[1]/div/label".format(
                    H_MSG_rowItem
                )
            )  # 一行上のチェックボックス
            self.H_MSG_Cbx = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[{}]/td[1]/div/label".format(
                    H_MSG_rowItem + 1
                )
            )  # チェックボックス
            self.H_driver.execute_script(
                "arguments[0].click();", self.H_FirstMSG_Cbx
            )  # チェックボックスクリック
            self.H_driver.execute_script(
                "arguments[0].click();", self._MSG_Cbx
            )  # チェックボックスクリック
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            self.H_driver.execute_script("hyouziAction()")  # Javaで表示
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
        except Exception:
            self.H_MSG_Cbx = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/div/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr[{}]/td[1]/div/label".format(
                    H_MSG_rowItem + 1
                )
            )  # チェックボックス
            self.H_driver.execute_script(
                "arguments[0].click();", self.H_MSG_Cbx
            )  # チェックボックスクリック
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            self.H_driver.execute_script("hyouziAction()")  # Javaで表示
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
        finally:
            try:
                # 要素取得
                self.H_MSG_TableItem1 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/section/table/tbody/tr[1]/td"
                ).text.replace(
                    "\u3000", " "
                )  # 発行元
                self.H_MSG_TableItem2 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/section/table/tbody/tr[2]/td[1]"
                ).text.replace(
                    "\u3000", " "
                )  # 発行元2
                self.H_MSG_TableItem3 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/section/table/tbody/tr[3]/td[1]"
                ).text.replace(
                    "\u3000", " "
                )  # 発行日時
                self.H_MSG_TableItem4 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/section/table/tbody/tr[4]/td"
                ).text.replace(
                    "\u3000", " "
                )  # 件名
                return True
            except:
                self.H_MSG_TableItem1 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[1]/td"
                ).text.replace(
                    "\u3000", " "
                )  # 発行元
                self.H_MSG_TableItem2 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[2]/td[1]"
                ).text.replace(
                    "\u3000", " "
                )  # 発行元2
                self.H_MSG_TableItem3 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[3]/td[1]"
                ).text.replace(
                    "\u3000", " "
                )  # 発行日時
                self.H_MSG_TableItem4 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[4]/td"
                ).text.replace(
                    "\u3000", " "
                )  # 件名
                if (
                    "/" not in self.H_MSG_TableItem3
                    and ":" not in self.H_MSG_TableItem3
                ):
                    self.H_MSG_TableItem1 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[1]/td"
                    ).text.replace(
                        "\u3000", " "
                    )  # 発行元
                    self.H_MSG_TableItem2 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[3]/td[1]"
                    ).text.replace(
                        "\u3000", " "
                    )  # 発行元2
                    self.H_MSG_TableItem3 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[2]/td[1]"
                    ).text.replace(
                        "\u3000", " "
                    )  # 発行日時
                    self.H_MSG_TableItem4 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/form/div/div/div[2]/table/tbody/tr[6]/td"
                    ).text.replace(
                        "\u3000", " "
                    )  # 件名
                    return True
                else:
                    return True

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def PrintIFS(self):
        """
        メッセージボックスの内容を印刷
        """
        try:
            self.H_driver.execute_script("window.print();")  # 開いたタブを印刷
            # time.sleep(4)
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            self.H_BackBtn = self.H_driver.find_element_by_xpath(
                "/html/body/div[2]/form/footer/div[1]/div/div[1]/a"
            )
            self.H_BackBtn.click()
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            logger.debug("印刷完了")
            return True
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def RenamePDF(self):
        self.tdy = dt.strptime(self.H_MSG_TableItem3, "%Y/%m/%d %H:%M:%S")  # 文字列を日付型に変換
        self.CfolName = str(self.tdy.year) + "-" + str(self.tdy.month)
        self.folders = glob.glob("//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/*-*")
        self.KanyoFolName = str(self.H_SCode) + "_" + self.H_TKCName
        for foldersItem in self.folders:
            if (
                foldersItem
                == "//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\" + self.CfolName
            ):
                self.Subfolders = glob.glob(foldersItem + "/*")  # フォルダーがあった場合
                if self.H_MSG_TableItem4 == "プレ申告データに関するお知らせ":
                    self.MotherFol = foldersItem + "\\eLTAX"
                    for SubfoldersItem in self.Subfolders:
                        if SubfoldersItem.replace("\u3000", " ") == self.MotherFol:
                            self.SubFFlag = 1
                            print(self.MotherFol + "あります")
                            break
                        else:
                            self.SubFFlag = 0
                    if self.SubFFlag == 0:
                        os.mkdir(foldersItem + "/eLTAX")
                        logger.debug(str(self.MotherFol) + "_作成")
                        print(self.MotherFol + "作りました。")
                        break
                    self.Tfolders = glob.glob(
                        foldersItem + "\\eLTAX" + "/*"
                    )  # フォルダーがあった場合
                    self.ChildFol = foldersItem + "\\eLTAX" + "\\" + self.KanyoFolName
                    self.TFFlag = 0
                    for TfoldersItem in self.Tfolders:
                        if TfoldersItem == self.ChildFol:
                            self.TFFlag = 1
                            print(self.ChildFol + "あります")
                            break
                        else:
                            self.TFFlag = 0
                    if self.TFFlag == 0:
                        os.mkdir(foldersItem + "/eLTAX" + "/" + self.KanyoFolName)
                        logger.debug(str(self.ChildFol) + "_作成")
                        print(self.ChildFol + "作りました。")
                        break
                    break
                else:
                    self.MotherFol = foldersItem + "\\eLTAX受信通知"
                    for SubfoldersItem in self.Subfolders:
                        if SubfoldersItem.replace("\u3000", " ") == self.MotherFol:
                            SubFFlag = 1
                            print(self.MotherFol + "あります")
                            self.ChildFol = self.MotherFol
                            break
                        else:
                            SubFFlag = 0
                    if SubFFlag == 0:
                        os.mkdir(foldersItem + "/eLTAX受信通知")
                        logger.debug(str(self.MotherFol) + "_作成")
                        print(self.MotherFol + "作りました。")
                        self.ChildFol = self.MotherFol
                        break
                    else:
                        break
            else:
                print("ありません")  # フォルダーがなかった場合
        self.PDFfolder = glob.glob(
            os.getcwd().replace("\\", "/") + "/" + "*.pdf"
        )  # フォルダーがあった場合
        for PDFfolderItem in self.PDFfolder:
            self.PDFSerch = "メッセージ照会_お知らせ" in PDFfolderItem
            if self.PDFSerch is False:
                self.PDFSerch = "メッセージ照会_受付通知（申告）" in PDFfolderItem
            if self.PDFSerch is False:
                self.PDFSerch = "メッセージ照会_受付通知（利用届出、申請・届出）" in PDFfolderItem
            if self.PDFSerch is False:
                self.PDFSerch = "メッセージ照会_申告書不受理通知" in PDFfolderItem
            if self.PDFSerch is False:
                self.PDFSerch = "メッセージ照会_納付情報発行依頼通知" in PDFfolderItem
            if self.H_MSG_TableItem4 == "プレ申告データに関するお知らせ":
                self.PDFName = (
                    self.KanyoFolName
                    + "_"
                    + self.H_MSG_TableItem1
                    + "_"
                    + self.H_MSG_TableItem2
                    + "_"
                    + self.H_MSG_TableItem4
                    + "["
                    + self.H_row
                    + "]"
                    + ".pdf"
                )
            else:
                self.PDFName = (
                    self.KanyoFolName
                    + "_"
                    + self.H_MSG_TableItem1
                    + "_"
                    + self.H_MSG_TableItem2
                    + "_"
                    + self.H_MSG_TableItem4
                    + ".pdf"
                )
            self.PDFPath = os.getcwd().replace("\\", "/") + "/" + self.PDFName
            self.PDFPath = self.PDFPath.replace("/", "\\")

        try:
            if self.PDFSerch is True:
                os.rename(PDFfolderItem, self.PDFPath)
                self.MovePDFPath = self.ChildFol + "/" + self.PDFName
                self.MovePDFPath = self.MovePDFPath.replace("/", "\\")
                shutil.move(self.PDFPath, self.MovePDFPath)
                logger.debug(str(self.MovePDFPath) + "_PDF作成成功")
                self.OKstr = self.MovePDFPath + "成功"
                self.OKstr = self.OKstr.replace("\uff0d", "-").replace("\xa0", "")
                self.OKLog.append(self.OKstr)
                return True
            # 'D:\\PythonScript\\国税電子申告・納税システム－SU00S100 メール詳細 (1).pdf'
            else:
                logger.debug(str(self.MovePDFPath) + "_リネーム失敗-メッセージ照会_お知らせが含まれないファイル名-")
                self.NGstr = self.MovePDFPath + "_リネーム失敗-メッセージ照会_お知らせが含まれないファイル名-"
                self.NGstr = self.NGstr.replace("\uff0d", "-").replace("\xa0", "")
                self.NGLog.append(self.NGstr)
                return False
        except:
            traceback.print_exc()
            logger.debug(str(self.MovePDFPath) + "_リネーム失敗-トレースバックエラー-")
            self.NGstr = self.MovePDFPath + "_リネーム失敗-トレースバックエラー-"
            self.NGstr = self.NGstr.replace("\uff0d", "-").replace("\xa0", "")
            self.NGLog.append(self.NGstr)
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def SortPDF(self):
        try:
            self.tdy = dt.strptime(
                self.H_MSG_TableItem3, "%Y/%m/%d %H:%M:%S"
            )  # 文字列を日付型に変換
            self.CfolName = str(self.tdy.year) + "-" + str(self.tdy.month)
            self.KanyoFolName = str(self.H_SCode) + "_" + self.H_TKCName
            if self.H_MSG_TableItem4 == "プレ申告データに関するお知らせ":
                self.PDFName = (
                    self.KanyoFolName
                    + "_"
                    + self.H_MSG_TableItem1
                    + "_"
                    + self.H_MSG_TableItem2
                    + "_"
                    + self.H_MSG_TableItem4
                    + "["
                    + self.H_row
                    + "]"
                    + ".pdf"
                )
                self.dir_path = (
                    "//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/"
                    + self.CfolName
                    + "/eLTAX//"
                    + self.KanyoFolName
                )
            else:
                self.PDFName = (
                    self.KanyoFolName
                    + "_"
                    + self.H_MSG_TableItem1
                    + "_"
                    + self.H_MSG_TableItem2
                    + "_"
                    + self.H_MSG_TableItem4
                    + ".pdf"
                )
                self.dir_path = (
                    "//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/"
                    + self.CfolName
                    + "/eLTAX受信通知"
                )
            for current_dir, sub_dirs, files_list in os.walk(self.dir_path):
                for file_name in files_list:
                    foldersItem = os.path.join(current_dir, file_name).replace(
                        "\u3000", ""
                    )
                    self.PDFSerch = self.PDFName in foldersItem
            try:
                if self.PDFSerch is True:
                    # 同名ファイルがあった場合
                    return True
                else:
                    # 同名ファイルがなかった場合
                    return False
            except Exception:
                return False
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------


# pandas(pd)で関与先データCSVを取得
if __name__ == "__main__":
    H_url = r"//nas-sv/A_共通/A8_ｼｽﾃﾑ資料/RPA/ALLDataBase/Heidi関与先DB.csv"
    H_df = pd.read_csv(H_url, encoding="utf-8")
    H_forCount = 0
    H_dfRow = np.array(H_df).shape[0]  # 配列行数取得
    H_dfCol = np.array(H_df).shape[1]  # 配列列数取得
    InputYear = input("取得年を西暦で入力してください。(例:2022)")
    InputMonth = input("取得月を西暦で入力してください。(例:9)")
    for x in range(H_dfRow):
        try:
            if x >= 785:
                # 関与先DB配列をループして識別番号とPassを取得
                H_dfDataRow = H_df.loc[x]
                WEB = elTaxWeb(H_dfDataRow, InputYear, InputMonth)
        except:
            traceback.print_exc()
