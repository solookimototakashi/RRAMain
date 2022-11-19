###########################################################################################################
# 稼働設定：解像度 1920*1080 表示スケール125%
###########################################################################################################
# pandasインポート
import pandas as pd

# 配列計算関数numpyインポート
import numpy as np

# timeインポート
import time

# reインポート
import re

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

# glob(フォルダファイルチェックコマンド)インポート
import glob

# shutil(フォルダファイル編集コマンド)インポート
import shutil

# 例外処理判定の為のtracebackインポート
import traceback

# -----------------------------------------------------------------------------------------------


class eTaxWeb:
    """
    概要: Webページクラス
    """

    def __init__(self, H_dfDataRow):
        self.OKLog = []
        self.NGLog = []
        self.pyoklog = r"//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/PyOKLog.csv"
        self.pynglog = r"//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/PyNGLog.csv"
        self.H_SCode = H_dfDataRow["SyanaiCode"]  # 関与先コード
        # --------------------------------------------------------------------
        try:
            self.H_TKCName = H_dfDataRow["TKCName"].replace(
                "\u3000", " "
            )  # 関与先名空白\u3000を置換
        except:
            self.H_TKCName = H_dfDataRow["TKCName"]  # 関与先
        self.H_MSikibetu = str(int(H_dfDataRow["MirokuKokuzeiUserCode"]))  # ミロク識別番号
        self.H_TSikibetu = str(int(H_dfDataRow["TKCKokuzeiUserCode"]))  # TKC識別番号
        # --------------------------------------------------------------------
        try:
            self.H_ePass = H_dfDataRow["etaxPass"].replace("\u3000", " ")  # etaxパスワード
        except:
            self.H_ePass = H_dfDataRow["etaxPass"]  # etaxパスワード
        try:
            self.H_elPass = H_dfDataRow["eltaxPass"].replace(
                "\u3000", " "
            )  # eltaxパスワード
        except:
            self.H_elPass = H_dfDataRow["eltaxPass"]  # eltaxパスワード
        # --------------------------------------------------------------------
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
        self.H_WEBurl = "https://uketsuke.e-tax.nta.go.jp/UF_APP/lnk/loginCtlKakutei"
        # ブラウザのウィンドウを表すオブジェクト"driver"を作成
        self.H_driver = webdriver.Chrome(
            executable_path=self.H_path, chrome_options=self.H_options
        )
        self.H_driver.get(self.H_WEBurl)
        # 初めの拡張機能ポップアップを閉じる
        try:
            self.H_PopupClose_btn = self.H_driver.find_element_by_xpath(
                "/html/body/div[4]/div[2]/div[2]/form/input[1]"
            )  # 閉じるボタンの要素指定
            self.H_PopupClose_btn.click()
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
        except:
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
        try:
            # Login関数処理
            self.LoginFlag = 1  # ログイン回数フラグ
            SWL = self.WebLogin()
            if SWL is True:
                SLR = self.LogReturn()
                if SLR is True:
                    # ログイン成功
                    SML = self.MsgLogin()
                    if SML is True:
                        print(self.H_TKCName + "DL完了")
                        return True
                else:
                    # 認証エラー失敗一回目
                    print(self.H_TKCName + "認証エラー")
                    self.H_driver.close()
                    self.H_driver.quit  # 画面を一回閉じる
                    # ブラウザのウィンドウを表すオブジェクト"driver"を作成
                    self.H_driver = webdriver.Chrome(
                        executable_path=self.H_path, chrome_options=self.H_options
                    )
                    self.H_driver.get(self.H_WEBurl)
                    # 初めの拡張機能ポップアップを閉じる
                    self.LoginFlag = 2  # ログイン回数フラグ
                    SWL = self.WebLogin()
                    if SWL is True:
                        SLR = self.LogReturn()
                        if SLR is True:
                            SML = self.MsgLogin()
                            if SML is True:
                                print(self.H_TKCName + "DL完了")
                                return True
                        else:
                            return False
                    else:
                        # 認証エラー失敗二回目
                        return False
            else:
                print(self.H_TKCName + "DLエラー")
                return False
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def WebLogin(self):
        """
        概要: Webページログイン処理
        """
        try:
            # ログイン画面要素指定------------------------------------------------------------------------------------------------------
            # 利用者識別番号入力欄をxpathで取得
            self.H_Sikibetu_box1 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[1]/table/tbody/tr/td/input[1]"
            )
            self.H_Sikibetu_box2 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[1]/table/tbody/tr/td/input[2]"
            )
            self.H_Sikibetu_box3 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[1]/table/tbody/tr/td/input[3]"
            )
            self.H_Sikibetu_box4 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[1]/table/tbody/tr/td/input[4]"
            )
            # ログインPASS入力欄をxpathで取得
            self.H_LogPassBox = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[2]/table/tbody/tr/td/span[1]/input"
            )
            self.H_Log_btn = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/div[2]/form[2]/p/input[1]"
            )
            # 識別番号とパスのnan判定--------------------------------------------------------------
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
            # -----------------------------------------------------------------------------------
            if LL is True:
                pd.DataFrame(self.OKLog).to_csv(
                    self.pyoklog,
                    encoding="shift-jis",
                )
                pd.DataFrame(self.NGLog).to_csv(
                    self.pynglog,
                    encoding="shift-jis",
                )
                self.LoginFlag += 1  # ログイン回数フラグ
                return True
            else:
                NGstr = str(self.H_SCode) + "_" + self.H_TKCName + "_" + "エラー"
                NGstr = NGstr.replace("\uff0d", "-").replace("\xa0", "")
                self.NGLog.append(NGstr)
                pd.DataFrame(self.NGLog).to_csv(
                    self.pynglog,
                    encoding="shift-jis",
                )
                self.H_driver.close()
                self.H_driver.quit()
                self.LoginFlag += 1  # ログイン回数フラグ
                return False
        except:
            self.LoginFlag += 1  # ログイン回数フラグ
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LoginLoop(self, H_id, H_pass):
        """
        概要: Webページログイン入力処理
        """
        try:
            self.H_id = str(int(H_id))  # 識別番号の変数代入
            self.H_pass = str(H_pass)
            # WEB操作--------------------------------------------------------------------------------------------------------------------
            self.H_idArray = [
                self.H_id[i : i + 4] for i in range(0, len(self.H_id), 4)
            ]  # 引数H_idを4文字づつ分割
            self.H_Sikibetu_box1.send_keys(self.H_idArray[0])
            self.H_Sikibetu_box2.send_keys(self.H_idArray[1])
            self.H_Sikibetu_box3.send_keys(self.H_idArray[2])
            self.H_Sikibetu_box4.send_keys(self.H_idArray[3])
            self.H_LogPassBox.send_keys(self.H_pass)
            self.H_Log_btn.click()
            WebDriverWait(self.H_driver, 30).until(
                EC.presence_of_all_elements_located
            )  # 要素が読み込まれるまで最大30秒待つ
            # ---------------------------------------------------------------------------------------------------------------------------
            try:
                self.H_H1 = self.H_driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/h1"
                )  # H1要素を取得
                return True
            except Exception:
                try:
                    self.H_H1 = self.H_driver.find_element_by_xpath(
                        "/html/body/div[1]/div[2]/form/h1"
                    )  # H1要素を取得
                    return True
                except:
                    return False
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LogOnOuter(self):
        """
        概要: ログイン後WEBテーブル取得
        """
        try:
            self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/p[1]"
            )  # メッセージボックスの中身ページのH1を指定
            self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[1]"
            )  # メッセージボックスのテーブルをxpath取得
        except Exception:
            pass
        else:
            self.dfs = pd.read_html(
                self.H_driver.page_source, encoding="cp932"
            )  # pandasにWEBURLをぶっこむ
            self.H_MSG_row = len(self.dfs[0])  # テーブル行数を代入
            return True

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def ParGet(self, H_MSG_rowItem):
        """
        概要: メッセージボックス情報取得
        """
        try:
            # WEBTable日付
            self.H_MSG_TableItem1 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[1]".format(
                    H_MSG_rowItem + 1
                )
            ).text
            self.H_MSG_TableItem2 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[2]".format(
                    H_MSG_rowItem + 1
                )
            ).text
            self.H_MSG_TableItem3 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[3]".format(
                    H_MSG_rowItem + 1
                )
            ).text
            # WEBTableメールタイトル
            self.H_MSG_TableItem4 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[4]".format(
                    H_MSG_rowItem + 1
                )
            ).text.replace("\u3000", " ")
            self.H_MSG_TableItem5 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[5]".format(
                    H_MSG_rowItem + 1
                )
            ).text
            self.H_MSG_TableItem6 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[6]".format(
                    H_MSG_rowItem + 1
                )
            ).text
            self.H_MSG_TableItem7 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/div/form/table/tbody/tr[{}]/td[7]".format(
                    H_MSG_rowItem + 1
                )
            ).text

        except Exception:
            pass
            # print(str(H_SC) + "_" + H_TN + "データ無")
        else:
            self.H_result = self.H_driver.page_source.split(
                "goDetail", self.H_MSG_row
            )  # ページソースからgoDetail区切りをテーブル行数分行う
            try:
                self.H_result = self.H_result[H_MSG_rowItem + 1].split(
                    ">"
                )  # goDetail区切り配列からループ回数に応じたJavaScript引数を抽出
                self.H_OnClickPar = (
                    self.H_result[00]
                    .replace("(", "")
                    .replace("/", "")
                    .replace(")", "")
                    .replace('"', "")
                )  # JavaScript引数を置換して成型
                self.H_driver.execute_script(
                    "javascript:goDetail({})".format(self.H_OnClickPar)
                )  # 抽出したJavaScript引数を利用して実行
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.switch_to.window(
                    self.H_driver.window_handles[1]
                )  # 開いたタブに移動する
                # -----------------------------------------ページ内条件分岐---------------------------------------------
                self.H_Title = self.H_driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/h2[1]"
                ).text  # 先頭H2タイトル取得
                return True
            except:
                # 鍵付きのMSGBOXしかない場合
                print("")
                return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def LogReturn(self):
        try:
            if self.H_H1.text == "メインメニュー":  # ログイン成功の場合
                # WEB画面要素指定------------------------------------------------------------------------------------------------------
                try:
                    self.H_MSG_OpenMyNo = self.H_driver.find_element_by_xpath(
                        "/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[3]/div/form/p/input"
                    )
                    self.H_MSG_OpenMyNo.click()
                    WebDriverWait(self.H_driver, 30).until(
                        EC.presence_of_all_elements_located
                    )  # 要素が読み込まれるまで最大30秒待つ
                    self.H_LogMSGAns = self.H_driver.find_element_by_xpath(
                        "/html/body/div[1]/div[2]/p[1]"
                    ).text  # メッセージボックス画面の文言を返す
                    return True
                except:
                    self.H_MSG_Open = self.H_driver.find_element_by_xpath(
                        "/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[2]/form/p/input"
                    )  # メッセージボックスを開くボタンをxpath指定
                    self.H_MSG_Open.click()
                    WebDriverWait(self.H_driver, 30).until(
                        EC.presence_of_all_elements_located
                    )  # 要素が読み込まれるまで最大30秒待つ
                    self.H_LogMSGAns = self.H_driver.find_element_by_xpath(
                        "/html/body/div[1]/div[2]/p[1]"
                    ).text  # メッセージボックス画面の文言を返す
                    return True
            elif self.H_H1.text == "認証エラー":  # ログイン失敗の場合
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                return False
            elif self.H_H1.text == "暗証番号の変更":  # ログイン失敗の場合
                try:
                    self.H_PassKousinBtn = self.H_driver.find_element_by_xpath(
                        "/html/body/div[2]/div/form[2]/input[3]"
                    )
                    self.H_PassKousinBtn.click()
                    WebDriverWait(self.H_driver, 30).until(
                        EC.presence_of_all_elements_located
                    )  # 要素が読み込まれるまで最大30秒待つ
                    try:
                        self.H_MSG_OpenMyNo = self.H_driver.find_element_by_xpath(
                            "/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[3]/div/form/p/input"
                        )
                        self.H_MSG_OpenMyNo.click()
                        WebDriverWait(self.H_driver, 30).until(
                            EC.presence_of_all_elements_located
                        )  # 要素が読み込まれるまで最大30秒待つ
                        self.H_LogMSGAns = self.H_driver.find_element_by_xpath(
                            "/html/body/div[1]/div[2]/p[1]"
                        ).text  # メッセージボックス画面の文言を返す
                        return True
                    except:
                        self.H_MSG_Open = self.H_driver.find_element_by_xpath(
                            "/html/body/div[1]/div[2]/div/div[1]/div/div[1]/div[2]/form/p/input"
                        )  # メッセージボックスを開くボタンをxpath指定
                        self.H_MSG_Open.click()
                        WebDriverWait(self.H_driver, 30).until(
                            EC.presence_of_all_elements_located
                        )  # 要素が読み込まれるまで最大30秒待つ
                        self.H_LogMSGAns = self.H_driver.find_element_by_xpath(
                            "/html/body/div[1]/div[2]/p[1]"
                        ).text  # メッセージボックス画面の文言を返す
                        return True
                except Exception:
                    pass
            else:  # その他の場合
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                return False
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def MsgLogin(self):
        try:
            if (
                not self.H_LogMSGAns == "表示するメッセージはありません。"
                and not self.H_LogMSGAns == "認証エラー"
            ):
                self.LogOnOuter()  # WEBテーブル取得
                # Loop------------------------------------------------------------------------------
                for H_MSG_rowItem in range(self.H_MSG_row):  # WEBテーブル行数分ループ
                    self.ParGet(H_MSG_rowItem)  # メッセージボックス情報取得
                    SPDF = SortPDF(
                        self.H_MSG_TableItem1,
                        self.H_MSG_TableItem4,
                        self.H_SCode,
                        self.H_TKCName,
                    )
                    # 振替納税の判定---------------------------------------
                    Hurikae = "振替納税のお知らせ" in self.H_MSG_TableItem4
                    # ----------------------------------------------------
                    if SPDF is False or SPDF is None or Hurikae is True:
                        # 開いたタブが未DLの場合------------------------------------------------------
                        self.PIFS = self.PrintIFS()  # メッセージボックスの内容に応じて処理分け
                        RPDF = self.RenamePDF(
                            self.H_MSG_TableItem1,
                            self.H_MSG_TableItem4,
                            self.H_SCode,
                            self.H_TKCName,
                        )  # PDF保存先フォルダー作成後リネーム&移動
                        if RPDF is True:
                            if H_MSG_rowItem == self.H_MSG_row - 1:
                                self.LoginFlag = 0
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
                            print(self.H_SCode + "_" + self.H_TKCName + "_" + "リネーム失敗")
                            self.LoginFlag = 0
                            self.H_driver.quit()
                            time.sleep(1)
                    else:
                        # 開いたタブがDL済の場合------------------------------------------------------
                        self.H_driver.close()  # タブを閉じる
                        print(
                            str(self.H_SCode)
                            + "_"
                            + self.H_TKCName
                            + "_"
                            + self.H_MSG_TableItem4
                            + "：重複有"
                        )
                        time.sleep(1)
                        if H_MSG_rowItem == self.H_MSG_row - 1:
                            # メッセージボックスにもうMSGが無ければ
                            self.LoginFlag = 0
                            self.H_driver.quit()  # ブラウザを閉じる
                            time.sleep(2)
                        else:
                            self.H_driver.switch_to.window(
                                self.H_driver.window_handles[0]
                            )  # タブ移動する
                            WebDriverWait(self.H_driver, 30).until(
                                EC.presence_of_all_elements_located
                            )  # 要素が読み込まれるまで最大30秒待つ
                        # ---------------------------------------------------------------------------
                return True
                # -----------------------------------------------------------------------------------
            elif self.H_LogMSGAns == "認証エラー":
                print(self.H_SCode + "_" + self.H_TKCName + "_" + "暗証番号の変更")
                return False
            else:
                self.LoginFlag = 0
                self.H_driver.close()
                self.H_driver.quit()
                return False
        except:
            return False

    # -----------------------------------------------------------------------------------------------------------------------------------------------------
    def PrintIFS(self):
        """
        pyファイルと同ディレクトリにリネーム無しで保存
        """
        try:
            if re.search("法定調書の提出について", self.H_Title):
                self.H_driver.execute_script("window.print();")  # 開いたタブを印刷
                # time.sleep(4)
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.close()
                self.H_driver.switch_to.window(
                    self.H_driver.window_handles[0]
                )  # タブ移動する
                return "法定調書"
            elif re.search("申告のお知らせ", self.H_Title):
                self.H_driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/form/p/input"
                ).click()  # 申告のお知らせボタンをクリック
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.execute_script("window.print();")  # 開いたタブを印刷
                # time.sleep(4)
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.close()
                self.H_driver.switch_to.window(
                    self.H_driver.window_handles[0]
                )  # タブ移動する
                return "申告のお知らせ"
            elif re.search("振替納税のお知らせ", self.H_Title):
                self.H_driver.find_element_by_xpath(
                    "/html/body/div[1]/div[2]/form/p/input"
                ).click()  # 申告のお知らせボタンをクリック
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.execute_script("window.print();")  # 開いたタブを印刷
                # time.sleep(4)
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.close()
                self.H_driver.switch_to.window(
                    self.H_driver.window_handles[0]
                )  # タブ移動する
                return "振替納税のお知らせ"
            else:
                self.H_driver.execute_script("window.print();")  # 開いたタブを印刷
                # time.sleep(4)
                WebDriverWait(self.H_driver, 30).until(
                    EC.presence_of_all_elements_located
                )  # 要素が読み込まれるまで最大30秒待つ
                self.H_driver.close()
                self.H_driver.switch_to.window(
                    self.H_driver.window_handles[0]
                )  # タブ移動する
                return self.H_Title
        except:
            return False

    # ----------------------------------------------------------------------------------------------------------------------------
    def RenamePDF(self, DownTime, MTitle, KanyoNo, KanyoName):
        """
        同名のPDFファイルが指定ディレクトリにあるか確認
        """
        try:
            tdy = dt.strptime(DownTime, "%Y/%m/%d %H:%M:%S")  # 文字列を日付型に変換
            CfolName = str(tdy.year) + "-" + str(tdy.month)
            folders = glob.glob("//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/*-*")
            KanyoFolName = str(KanyoNo) + "_" + KanyoName
            for foldersItem in folders:
                if foldersItem == "//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ\\" + CfolName:
                    Subfolders = glob.glob(foldersItem + "/*")  # フォルダーがあった場合
                    MotherFol = foldersItem + "\\eTAX"
                    for SubfoldersItem in Subfolders:
                        if SubfoldersItem.replace("\u3000", " ") == MotherFol:
                            SubFFlag = 1
                            print(MotherFol + "あります")
                            break
                        else:
                            SubFFlag = 0
                    if SubFFlag == 0:
                        os.mkdir(foldersItem + "/eTAX")
                        print(MotherFol + "作りました。")
                        break
                    Tfolders = glob.glob(foldersItem + "\\eTAX" + "/*")  # フォルダーがあった場合
                    ChildFol = foldersItem + "\\eTAX" + "\\" + KanyoFolName
                    TFFlag = 0
                    for TfoldersItem in Tfolders:
                        if TfoldersItem == ChildFol:
                            TFFlag = 1
                            print(ChildFol + "あります")
                            break
                        else:
                            TFFlag = 0
                    if TFFlag == 0:
                        os.mkdir(foldersItem + "/eTAX" + "/" + KanyoFolName)
                        print(ChildFol + "作りました。")
                        break
                    break
                else:
                    print("ありません")  # フォルダーがなかった場合
            PDFfolder = glob.glob(
                os.getcwd().replace("\\", "/") + "/" + "*.pdf"
            )  # フォルダーがあった場合
            PDFSerch = False
            for PDFfolderItem in PDFfolder:
                if "国税電子申告・納税システム" in PDFfolderItem:
                    DTime = (
                        "{0:04}".format(tdy.year)
                        + "{0:02}".format(tdy.month)
                        + "{0:02}".format(tdy.day)
                        + " "
                        + "{0:02}".format(tdy.hour)
                        + "{0:02}".format(tdy.minute)
                        + "{0:02}".format(tdy.second)
                    )
                    PDFName = KanyoFolName + "_" + MTitle + "_" + DTime + ".pdf"
                    PDFPath = os.getcwd().replace("\\", "/") + "/" + PDFName
                    PDFPath = PDFPath.replace("/", "\\")
                    PDFSerch = True
                    break
            try:
                if PDFSerch is True:
                    os.rename(PDFfolderItem, PDFPath)
                    MovePDFPath = ChildFol + "/" + PDFName
                    MovePDFPath = MovePDFPath.replace("/", "\\")
                    shutil.move(PDFPath, MovePDFPath)
                    OKstr = MovePDFPath + "成功"
                    OKstr = OKstr.replace("\uff0d", "-").replace("\xa0", "")
                    self.OKLog.append(OKstr)
                    return True
                else:
                    NGstr = MovePDFPath + "_リネーム失敗-国税電子申告・納税システムが含まれないファイル名-"
                    NGstr = NGstr.replace("\uff0d", "-").replace("\xa0", "")
                    self.NGLog.append(NGstr)
                    return False
            except:
                traceback.print_exc()
                NGstr = MovePDFPath + "_リネーム失敗-トレースバックエラー-"
                NGstr = NGstr.replace("\uff0d", "-").replace("\xa0", "")
                self.NGLog.append(NGstr)
                return False
        except:
            return False


# -----------------------------------------------------------------------------------------------------------------------------------------------------
def SortPDF(DownTime, MTitle, KanyoNo, KanyoName):
    """
    同名のPDFファイルが指定ディレクトリにあるか確認
    """
    try:
        tdy = dt.strptime(DownTime, "%Y/%m/%d %H:%M:%S")  # 文字列を日付型に変換
        CfolName = str(tdy.year) + "-" + str(tdy.month)
        KanyoFolName = str(KanyoNo) + "_" + KanyoName
        DTime = (
            "{0:04}".format(tdy.year)
            + "{0:02}".format(tdy.month)
            + "{0:02}".format(tdy.day)
            + " "
            + "{0:02}".format(tdy.hour)
            + "{0:02}".format(tdy.minute)
            + "{0:02}".format(tdy.second)
        )
        PDFName = KanyoFolName + "_" + MTitle + "_" + DTime + ".pdf"
        dir_path = "//nas-sv/B_監査etc/B2_電子ﾌｧｲﾙ/ﾒｯｾｰｼﾞﾎﾞｯｸｽ/" + CfolName + "/eTax"
        PDFSerch = False
        for current_dir, sub_dirs, files_list in os.walk(dir_path):
            for file_name in files_list:
                foldersItem = os.path.join(current_dir, file_name)
                if PDFName in foldersItem:
                    PDFSerch = True
        try:
            if PDFSerch is True:
                return True
            else:
                return False
        except:
            Exception
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
    for x in range(H_dfRow):
        try:
            if x >= 14:
                # 関与先DB配列をループして識別番号とPassを取得
                H_dfDataRow = H_df.loc[x]
                WEB = eTaxWeb(H_dfDataRow)

        except:
            traceback.print_exc()
