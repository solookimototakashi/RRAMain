# eTaxMSGDownLoad

* [機能](#機能)
* [利用ライブラリ](#利用ライブラリ)
* [モジュール関数](#モジュール関数)

# 機能 <a id="機能"></a>
#### 国税電子申告・納税システム（e-Tax）から
#### メッセージボックスをダウンロードし、ローカルフォルダへリネーム＆保存
<br>

# 利用ライブラリ <a id="利用ライブラリ"></a>

#### import pandas as pd
#### ・Pythonデータ分析ライブラリ
<br>

#### import numpy as np
#### ・Python高速計算ライブラリ
<br>

#### import time
#### ・Python時間操作ライブラリ
<br>

#### from selenium import webdriver
#### ・WEBRPAライブラリ(ドライバー)
<br>

#### from selenium.webdriver.chrome.options import Options
#### ・WEBRPAライブラリ(WEBオプション)
<br>

#### from selenium.webdriver.support.ui import WebDriverWait
#### ・WEBRPAライブラリ(WEBドライバー待機ライブラリ)
<br>

#### from selenium.webdriver.support import expected_conditions as EC
#### ・WEBRPAライブラリ(WEB要素操作ライブラリ)
<br>

#### import json
#### ・jsonデータ操作ライブラリ
<br>

#### import os
#### ・PythonOS操作ライブラリ
<br>

#### from datetime import datetime as dt
#### ・Python日付型操作ライブラリ
<br>

#### import glob
#### ・PythonOS操作ライブラリ
<br>

#### import shutil
#### ・PythonOS操作ライブラリ
<br>

#### import traceback
#### ・Pythonスタックトレース出力ライブラリ
<br>

# クラス関数
* [eTaxWeb](#eTaxWeb)
* [eTaxWebクラス変数](#eTaxWebクラス変数)
## eTaxWeb <a id="eTaxWeb"></a>
    def __init__(self):
    ==================================================================  
    機能：Webページクラス   
    ==================================================================
    引数：
        ・self：()
            GUIライブラリインスタンス自身
    ==================================================================
## eTaxWebクラス変数 <a id="eTaxWebクラス変数"></a>
* [Main](#Main)
* [フレーム](#フレーム)
  
## Main <a id="Main"></a>
    ・self.window_root
        自身のインスタンス(GUIウィンドウ)            
![](imgs/2022-09-07_09h13_32.png)

    <要素>
        ・self.H_fold
            このpyファイル

        ・self.H_options
            WEBブラウザーオプションのインスタンス

        ・self.appState
            WEBブラウザー設定(画像DL処理の設定)
            ・"id": "Save as PDF"
                画像ファイルをPDFでダウンロードする設定
            ・"origin": "local"           
                ダウンロードファイルの保存先

        ・self.prefs
            WEBブラウザープリフェッツインスタンス

        ・self.H_path
            WEBドライバ(chromedriver.exe)のpath

        ・self.H_WEBurl
            閲覧WEBサイトURL

        ・self.H_driver
            ブラウザウィンドウインスタンス

        ・self.H_PopupClose_btn
            ポップアップの閉じるボタン

        ・self.H_Sikibetu_box1
            利用者識別番号入力ボックス（1）

        ・self.H_Sikibetu_box2
            利用者識別番号入力ボックス（2）

        ・self.H_Sikibetu_box3
            利用者識別番号入力ボックス（3）
        
        ・self.H_Sikibetu_box4
            利用者識別番号入力ボックス（4）

        ・self.H_LogPassBox
            ログインPASS入力欄をxpathで取得

        ・self.H_Log_btn
            ログインボタン
            
        ・self.H_idArray
            利用者識別番号を4桁ずつ分割したリスト

        try:
            self.H_H1 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/h1"
            )  # H1要素を取得
            # return self.H_H1.text, self.H_driver
        except Exception:
            self.H_H1 = self.H_driver.find_element_by_xpath(
                "/html/body/div[1]/div[2]/form/h1"
            )  # H1要素を取得
        #     return self.H_H1.text, self.H_driver
        # else:
        #     pass

        ・self.window_root.geometry("1480x750")
            GUIウィンドウサイズ(横,縦)
        ・self.window_root.title("GUI Image Editor v0.90")
            GUIウィンドウタイトル
        ・self.control
            外部GUI操作ライブラリのインスタンス
        ・self.dir_path
            デフォルトのパス(pyファイルのディレクトリ) 
        ・self.file_list
            ファイルリストの表示テキスト
        ・self.clip_enable
            ファイルを開いたフラグ
[戻る](#Main)