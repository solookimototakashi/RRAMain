# osインポート
import os

# loggerインポート
from logging import getLogger

logger = getLogger()
# ----------------------------------------------------------------------------------------
def SerchdirFolders(URL):  # 指定URL配下のサブフォルダを取得
    List = []
    for fd_path, sb_folder, sb_file in os.walk(URL):
        for fol in sb_folder:
            # print(fd_path + '\\' + fol)
            List.append([fd_path, fol])
    return List


# ----------------------------------------------------------------------------------------
def SerchdirFiles(URL):  # 指定URL配下のファイルを取得
    List = []
    for fd_path, sb_folder, sb_file in os.walk(URL):
        for fil in sb_file:
            # print(fd_path + '\\' + fil)
            List.append([fd_path, fil])
    return List


# ----------------------------------------------------------------------------------------
def CreFol(URL, FolderName):  # 指定フォルダ内に引数名のフォルダが無ければ作成する。
    dir_List = os.listdir(URL)  # 指定URL配下のサブフォルダを取得
    print(dir_List)
    if FolderName in dir_List:
        print("フォルダーあり。")
        return URL + "/" + FolderName
    else:
        os.mkdir(URL + "/" + FolderName)
        print("フォルダー作成しました。")
        return URL + "/" + FolderName


# ----------------------------------------------------------------------------------------
