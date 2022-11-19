# elTaxDLOpen
[機能](#機能)
<br>

[使用モジュール](#使用モジュール)
<br>

[関数](#関数)
<br>

# 機能
#### このRPAフローはelTax(PCDESKWEB版)にアクセスし、
#### CSVファイルから読取ったID・Passを元にログインし、
#### メッセージボックスのファイルを指定のローカルディレクトリへ保存するRPAフローです。
#### (フォルダ作成機能あり) 
<br>

# 使用モジュール
#### from appium import webdriver : Appium-Python-Client==1.3.0
#### <補足>RPAライブラリ：クラス要素へのアクション
<br>

#### import subprocess : Python3.7.8
#### <補足>組込関数：プログラム起動
<br>

#### import pyautogui as pg : PyAutoGUI==0.9.53
#### <補足>RPAライブラリ：画像要素へのアクション,キー・マウスアクション等
<br>

#### import time : Python3.7.8
#### <補足>組込関数：待機処理
<br>

#### import os : Python3.7.8
#### <補足>組込関数：ファイルチェック・作成,フォルダチェック・作成
<br>

# 関数

#### MainFlow(BatUrl, FolURL2, ImgFolName)
#### <機能>メイン処理
#### <引数1>BatUrl：elTax(PCDESKWEB版)のURL(str)
#### <引数2>FolURL2 ：保存場所の親ディレクトリ(str)
#### <引数3>ImgFolName：未定

