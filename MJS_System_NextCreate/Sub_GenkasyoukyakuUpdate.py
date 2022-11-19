from ctypes import windll
import pyautogui as pg
import time
import RPA_Function as RPA
import pyperclip  # クリップボードへのコピーで使用
import wrapt_timeout_decorator

TIMEOUT = 600

# ------------------------------------------------------------------------------------------------------------------
@wrapt_timeout_decorator.timeout(dec_timeout=TIMEOUT)
def Flow(Job, Exc):
    """
    概要: 減価償却更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\GenkasyoukyakuUpdate"
        ErrStr = ""  # Rpaエラー判別変数
        # 減価償却フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\G_SyoukyakuFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        time.sleep(1)
        # 年度を最新に指定----------------------------------------------------
        IC = RPA.ImgCheck(URL, r"\Nendo_Saisin.png", 0.9, 10)
        if IC[0] is False:
            IC2 = RPA.ImgCheck(URL, r"\Nendo_All.png", 0.9, 10)
            if IC2[0] is False:
                print("年度選択がない")
            else:
                RPA.ImgClick(URL, r"\Nendo_All.png", 0.9, 10)
                pg.press("home")
                time.sleep(1)
                pg.press("down")
                time.sleep(1)
                pg.press("return")
                time.sleep(1)
        # ------------------------------------------------------------------
        # 関与先コード入力ボックスをクリック------------------------------------
        while pg.locateOnScreen(URL + r"\K_NoBox.png", confidence=0.9) is None:
            time.sleep(1)
        time.sleep(1)
        RPA.ImgClick(URL, r"\K_NoBox.png", 0.9, 10)
        while pg.locateOnScreen(URL + r"\K_AfterNoBox.png", confidence=0.9) is None:
            time.sleep(1)
        pg.write(str(Exc.row_data["関与先番号"]))
        pg.press(["return", "return", "return"])
        time.sleep(1)
        # 入力した関与先コードを取得------------
        pg.keyDown("shift")
        pg.press(["tab", "tab", "tab"])
        pg.keyUp("shift")
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()
        time.sleep(1)
        pg.hotkey("ctrl", "c")
        ThisNo = pyperclip.paste()
        # -----------------------------------
        pg.press("return")
        # 表示された年度を取得-----------------
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()
        time.sleep(1)
        pg.hotkey("ctrl", "c")
        ThisYear = pyperclip.paste()
        # -----------------------------------
        pg.press("return")
        # 表示された月を取得-------------------
        if windll.user32.OpenClipboard(None):
            windll.user32.EmptyClipboard()
            windll.user32.CloseClipboard()
        time.sleep(1)
        pg.hotkey("ctrl", "c")
        ThisMonth = pyperclip.paste()
        # -----------------------------------
        time.sleep(1)
        YearDiff = Job.Start_Year - int(ThisYear)
        if abs(YearDiff) > Job.Start_Year:
            YearDiff = 32 - int(ThisYear) + Job.Start_Year - 2
        if str(Exc.row_data["関与先番号"]) == ThisNo:
            print("関与先あり")
            if YearDiff == 1:  # 次年度更新か判定
                pg.press(["return", "return", "return"])
                time.sleep(1)
                # 減価償却メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\G_SyoukyakuMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    # アップデート情報画面が出たら閉じる-------------------------------
                    GSUM = RPA.ImgCheck(URL, r"\G_SyoukyakuUpMsg.png", 0.9, 10)
                    if GSUM[0] is True:
                        RPA.ImgClick(URL, r"\G_SyoukyakuUpMsgCansel.png", 0.9, 10)
                    # 顧問先情報変更ダイアログが表示されたら
                    CDQ = RPA.ImgCheck(
                        URL,
                        r"ChangeDataQ.png",
                        0.9,
                        10,
                    )
                    if CDQ[0] is True:
                        pg.press("y")  # yで決定
                        # 顧問先情報取込メニューが表示されるまで待機--------------------------
                        while (
                            RPA.ImgCheckForList(
                                URL,
                                [r"ChangeDataBtn.png", r"ChangeDataBtn2.png"],
                                0.9,
                                1,
                            )[0]
                            is False
                        ):
                            time.sleep(1)
                        CDB = RPA.ImgCheckForList(
                            URL,
                            [r"ChangeDataBtn.png", r"ChangeDataBtn2.png"],
                            0.9,
                            10,
                        )
                        RPA.ImgClick(URL, CDB[1], 0.9, 10)  # 顧問先情報取込ボタンをクリック
                # --------------------------------------------------------------------
                RPA.ImgClick(URL, r"\G_SyoukyakuMenu.png", 0.9, 10)  # 決算更新のアイコンをクリック
                RPA.ImgClick(URL, r"\M_G_Sonota.png", 0.9, 10)  # 5.その他処理アイコンをクリック
                # 決算更新メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\IkkatsuGenkaKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                KousinList = [r"\IkkatsuGenkaKousin.png", r"\IkkatsuGenkaKousin2.png"]
                while RPA.ImgCheckForList(URL, KousinList, 0.9, 10)[0] is False:
                    time.sleep(1)
                # --------------------------------------------------------------------
                IK = RPA.ImgCheckForList(URL, KousinList, 0.9, 10)
                if IK[0] is True:
                    RPA.ImgClick(URL, IK[1], 0.9, 10)  # 決算更新メニューのアイコンをクリック
                # 決算更新メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\IkkatuOpenGenkaFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                RPA.ImgClick(
                    URL, r"\IkkatuOpenGenkaFlag.png", 0.9, 10
                )  # 決算更新メニューのアイコンをクリック
                time.sleep(1)
                RPA.ImgClick(URL, r"\IkkatuGenakStart.png", 0.9, 10)  # 決算更新開始のアイコンをクリック
                # 確認ウィンドウが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SakuseiGenkaQ.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                pg.press("y")  # yで決定(nがキャンセル)
                # 処理終了ウィンドウが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SakuseiGenkaEnd.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    Nor = RPA.ImgCheck(URL, r"\No_Rendou.png", 0.9, 10)
                    if Nor[0] is True:
                        ErrStr = "Noren"
                        pg.press("y")  # yで決定(nがキャンセル)
                # --------------------------------------------------------------------
                pg.press("return")  # 決定
                # 決算更新のアイコンが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\G_SyoukyakuMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                # 減価償却フラグが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\G_SyoukyakuFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    al4c = RPA.ImgCheck(URL, r"\altf4Q.png", 0.9, 10)  # 終了確認が表示されたら
                    if al4c[0] is True:
                        pg.press("y")  # yで決定(nがキャンセル)
                # --------------------------------------------------------------------
                print("更新完了")
                if ErrStr == "":
                    # with open(Job.taskurl, "w") as f:
                    #     f.write([ThisNo, ThisYear, ThisMonth])
                    #     f.close()
                    return True, ThisNo, ThisYear, ThisMonth
                elif ErrStr == "Noren":
                    return False, ErrStr, "", ""
            elif str(Job.Start_Year) == ThisYear:
                return False, "該当年度有り", "", ""
            else:
                ID = IkkatuUpDate(Job, Exc, YearDiff)
                if ID is True:
                    return True, ThisNo, ThisYear, ThisMonth
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def IkkatuUpDate(Job, Exc, YearDiff):
    """
    減価償却:該当年度まで繰り返し更新
    """
    try:
        URL = Job.imgdir_url + r"\\GenkasyoukyakuUpdate"
        pg.press(["return", "return", "return"])
        time.sleep(1)
        # 減価償却メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\G_SyoukyakuMenu.png", confidence=0.9) is None:
            time.sleep(1)
            # アップデート情報画面が出たら閉じる-------------------------------
            GSUM = RPA.ImgCheck(URL, r"\G_SyoukyakuUpMsg.png", 0.9, 10)
            if GSUM[0] is True:
                RPA.ImgClick(URL, r"\G_SyoukyakuUpMsgCansel.png", 0.9, 10)
            # 顧問先情報変更ダイアログが表示されたら
            CDQ = RPA.ImgCheck(
                URL,
                r"ChangeDataQ.png",
                0.9,
                10,
            )
            if CDQ[0] is True:
                pg.press("y")  # yで決定
                # 顧問先情報取込メニューが表示されるまで待機--------------------------
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [r"ChangeDataBtn.png", r"ChangeDataBtn2.png"],
                        0.9,
                        1,
                    )[0]
                    is False
                ):
                    time.sleep(1)
                CDB = RPA.ImgCheckForList(
                    URL,
                    [r"ChangeDataBtn.png", r"ChangeDataBtn2.png"],
                    0.9,
                    10,
                )
                RPA.ImgClick(URL, CDB[1], 0.9, 10)  # 顧問先情報取込ボタンをクリック
        # --------------------------------------------------------------------
        # 次年まで繰り返し
        for i in range(YearDiff):
            RPA.ImgClick(URL, r"\G_SyoukyakuMenu.png", 0.9, 10)  # 減価償却メニューアイコンをクリック
            MGS = RPA.ImgCheck(URL, r"\M_G_Sonota.png", 0.9, 10)
            if MGS[0] is True:
                RPA.ImgClick(URL, r"\M_G_Sonota.png", 0.9, 10)  # 4.決算処理アイコンをクリック
            RPA.ImgClick(URL, r"\DataOpen_btn.png", 0.9, 10)  # データを開くをクリック
            while RPA.ImgCheck(URL, r"\DataOpen_flag.png", 0.9, 1)[0] is False:
                time.sleep(1)
            pg.write(str(Exc.row_data["関与先番号"]))
            pg.press("return")
            RPA.ImgClick(URL, r"\DataLastOpen_btn.png", 0.9, 10)  # データを開く
            while RPA.ImgCheck(URL, r"\DataOpen_flag.png", 0.9, 1)[0] is True:
                time.sleep(1)
            # 決算更新メニューが表示されるまで待機------------------------------------
            KousinList = [r"\IkkatsuGenkaKousin.png", r"\IkkatsuGenkaKousin2.png"]
            while RPA.ImgCheckForList(URL, KousinList, 0.9, 10)[0] is False:
                time.sleep(1)
            # --------------------------------------------------------------------
            IK = RPA.ImgCheckForList(URL, KousinList, 0.9, 10)
            if IK[0] is True:
                RPA.ImgClick(URL, IK[1], 0.9, 10)  # 決算更新メニューのアイコンをクリック
            # 決算更新メニューが表示されるまで待機------------------------------------
            while (
                pg.locateOnScreen(URL + r"\IkkatuOpenGenkaFlag.png", confidence=0.9)
                is None
            ):
                time.sleep(1)
            # --------------------------------------------------------------------
            RPA.ImgClick(
                URL, r"\IkkatuOpenGenkaFlag.png", 0.9, 10
            )  # 決算更新メニューのアイコンをクリック
            time.sleep(1)
            RPA.ImgClick(URL, r"\IkkatuGenakStart.png", 0.9, 10)  # 決算更新開始のアイコンをクリック
            # 確認ウィンドウが表示されるまで待機-------------------------------------
            while (
                pg.locateOnScreen(URL + r"\SakuseiGenkaQ.png", confidence=0.9) is None
            ):
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("y")  # yで決定(nがキャンセル)
            # 処理終了ウィンドウが表示されるまで待機----------------------------------
            while (
                pg.locateOnScreen(URL + r"\SakuseiGenkaEnd.png", confidence=0.9) is None
            ):
                time.sleep(1)
                Nor = RPA.ImgCheck(URL, r"\No_Rendou.png", 0.9, 10)
                if Nor[0] is True:
                    pg.press("y")  # yで決定(nがキャンセル)
            # --------------------------------------------------------------------
            pg.press("return")  # 決定
            # 決算更新のアイコンが表示されるまで待機----------------------------------
            while (
                pg.locateOnScreen(URL + r"\G_SyoukyakuMenu.png", confidence=0.9) is None
            ):
                time.sleep(1)
        # --------------------------------------------------------------------

        # 閉じる処理--------------------------
        pg.keyDown("alt")
        pg.press("f4")
        pg.keyUp("alt")
        # -----------------------------------
        # 減価償却フラグが表示されるまで待機-------------------------------------
        while pg.locateOnScreen(URL + r"\G_SyoukyakuFlag.png", confidence=0.9) is None:
            time.sleep(1)
            al4c = RPA.ImgCheck(URL, r"\altf4Q.png", 0.9, 10)  # 終了確認が表示されたら
            if al4c[0] is True:
                pg.press("y")  # yで決定(nがキャンセル)
        # --------------------------------------------------------------------
        print("更新完了")
        return True
    except:
        return False


# ------------------------------------------------------------------------------------------------------------------
def GenkasyoukyakuUpdate(Job, Exc):
    """
    main
    """
    try:
        f = Flow(Job, Exc)
        # プロセス待機時間
        time.sleep(3)
        return f
    except TimeoutError:
        return False, "TimeOut", "", ""
