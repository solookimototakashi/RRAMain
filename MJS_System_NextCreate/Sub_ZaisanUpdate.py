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
    概要: 財産評価更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\ZaisanUpdate"
        ErrStr = ""  # Rpaエラー判別変数

        # 財産評価フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\ZaisanhyoukaFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        time.sleep(1)
        # 他システムとメニューが違う-------------------------------------------------------
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
        pg.press(["return", "return"])
        time.sleep(1)
        # -----------------------------------
        if RPA.ImgCheck(URL, r"\NotData.png", 0.9, 10)[0] is True:
            # 入力した関与先コードを取得------------
            pg.press(["return", "return"])
            pg.keyDown("shift")
            pg.press(["tab", "tab", "tab", "tab"])
            pg.keyUp("shift")
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            time.sleep(1)
            pg.hotkey("ctrl", "c")
            ThisNo = pyperclip.paste()
            pg.press("return")
            # クリップボードをクリア----------------
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            # ------------------------------------
            # 表示された年度を取得-----------------
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            time.sleep(1)
            pg.hotkey("ctrl", "c")
            ThisYear = pyperclip.paste()
            # -----------------------------------
            pg.press("return")
            # 表示された申告種類を取得---------------
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            time.sleep(1)
            pg.hotkey("ctrl", "c")
            ThisMonth = pyperclip.paste()
            pg.press("return")
            # -----------------------------------
        else:
            # 入力した関与先コードを取得------------
            pg.press(["return", "return"])
            pg.keyDown("shift")
            pg.press(["tab", "tab", "tab", "tab"])
            pg.keyUp("shift")
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            time.sleep(1)
            pg.hotkey("ctrl", "c")
            ThisNo = pyperclip.paste()
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
            # 表示された申告種類を取得---------------
            if windll.user32.OpenClipboard(None):
                windll.user32.EmptyClipboard()
                windll.user32.CloseClipboard()
            time.sleep(1)
            pg.hotkey("ctrl", "c")
            ThisMonth = pyperclip.paste()
            pg.press("return")
            # -----------------------------------
        # 他システムとメニューが違う-------------------------------------------------------
        if str(Exc.row_data["関与先番号"]) == ThisNo:
            if str(Job.Start_Year + 1) == ThisYear:
                return False, "該当年度有り", "", ""
            else:
                print("関与先あり")
                pg.press(["return", "return", "return"])
                time.sleep(1)
                # 財産評価更新アイコンが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\ZaisanKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    CDQ = RPA.ImgCheck(
                        URL,
                        r"ZChangeDataQ.png",
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
                    PPP = RPA.ImgCheck(
                        URL,
                        r"Popup.png",
                        0.9,
                        10,
                    )
                    if PPP[0] is True:
                        pg.keyDown("alt")
                        pg.press("c")
                        pg.keyUp("alt")

                # データ基本情報を更新---------------------------------------------------
                RPA.ImgClick(URL, r"\DataIcon.png", 0.9, 10)
                while (
                    pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInIcon.png", 0.9, 10)
                while pg.locateOnScreen(URL + r"\DataInOK.png", confidence=0.9) is None:
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInOK.png", 0.9, 10)
                while (
                    pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                ME = RPA.ImgCheckForList(
                    URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                )
                if ME[0] is True:
                    RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                while pg.locateOnScreen(URL + r"\DataEndQ.png", confidence=0.9) is None:
                    time.sleep(1)
                pg.press("y")
                while (
                    pg.locateOnScreen(URL + r"\ZaisanKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                # --------------------------------------------------------------------
                RPA.ImgClick(URL, r"\ZaisanKousin.png", 0.9, 10)  # 一括更新のアイコンをクリック
                # 財産評価メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\ZaisanOpenFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(3)
                # 更新区分を次年繰越に変更---------------------------------------------
                RPA.ImgClick(URL, r"\Kousinkubun.png", 0.9, 10)
                pg.press("home")
                pg.press("return")
                # 更新区分フラグが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KousinkubunFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                FC = RPA.ImgCheckForList(
                    URL,
                    [
                        r"IkkatuFind.png",
                        r"IkkatuFind2.png",
                    ],
                    0.9,
                    10,
                )
                if FC[0] is True:
                    RPA.ImgClick(URL, FC[1], 0.9, 10)  # 一括更新メニューのアイコンをクリック
                    pyperclip.copy(str(Exc.row_data["関与先番号"]))
                    pg.hotkey("ctrl", "v")
                    # 検索ボタンまでエンター-------------------------------------
                    while RPA.ImgCheck(URL, r"ZFindFlag.png", 0.9, 1)[0] is False:
                        time.sleep(1)
                        pg.press("return")
                pg.press("return")
                time.sleep(1)
                pg.press("space")

                if ErrStr == "":
                    # チェックマークが表示されるまで待機-------------------------------------
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"IkkatuCheck.png",
                                r"ZaisanCheck.png",
                                r"NendCheck.png",
                                r"HouteiCheck.png",
                            ],
                            0.9,
                            1,
                        )[0]
                        is False
                    ):
                        time.sleep(1)
                        ZND = RPA.ImgCheck(URL, r"ZaisanNoData.png", 0.9, 10)
                        if ZND[0] is True:
                            ErrStr = "NoData"
                            break
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    if not ErrStr == "NoData":
                        RPA.ImgClick(
                            URL, r"\ZaisanStart.png", 0.9, 10
                        )  # 更新開始のアイコンをクリック
                        # 確認ウィンドウが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\ZaisanStartQ.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        pg.press("y")  # yで決定(nがキャンセル)
                        # 処理終了ウィンドウが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\ZaisanEnd.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        pg.press("return")
                        # チェックマークが表示されなくなるまで待機-------------------------------
                        while (
                            RPA.ImgCheckForList(
                                URL,
                                [
                                    r"IkkatuCheck.png",
                                    r"ZaisanCheck.png",
                                    r"NendCheck.png",
                                    r"HouteiCheck.png",
                                ],
                                0.9,
                                1,
                            )[0]
                            is True
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        ME = RPA.ImgCheckForList(
                            URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                        )
                        if ME[0] is True:
                            RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                        # 一括更新のアイコンが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\ZaisanOpenFlag.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        # 閉じる処理--------------------------
                        pg.keyDown("alt")
                        pg.press("f4")
                        pg.keyUp("alt")
                        # -----------------------------------
                        # 財産評価フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\ZaisanhyoukaFlag.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        # 初期画面で開封された財産評価項目を閉じる----------------------------------
                        HoujinList = [r"\Zaisanhyouka.png", r"\Zaisanhyouka2.png"]
                        HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                        if HLI[0] is True:
                            RPA.ImgClick(URL, HLI[1], 0.9, 10)
                        # --------------------------------------------------------------------
                        print("更新完了")
                    if ErrStr == "":
                        return True, ThisNo, ThisYear, ThisMonth
                    else:
                        # --------------------------------------------------------------------
                        DD = RPA.ImgCheck(URL, r"\DoubleDataQ.png", 0.9, 10)
                        if DD[0] is True:
                            pg.press("return")
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        ME = RPA.ImgCheckForList(
                            URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                        )
                        if ME[0] is True:
                            RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                        # 一括更新のアイコンが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\ZaisanOpenFlag.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        # 閉じる処理--------------------------
                        pg.keyDown("alt")
                        pg.press("f4")
                        pg.keyUp("alt")
                        # -----------------------------------
                        # 財産評価フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\ZaisanhyoukaFlag.png", confidence=0.9
                            )
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        # 初期画面で開封された財産評価項目を閉じる----------------------------------
                        HoujinList = [r"\Zaisanhyouka.png", r"\Zaisanhyouka2.png"]
                        HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                        if HLI[0] is True:
                            RPA.ImgClick(URL, HLI[1], 0.9, 10)
                        # --------------------------------------------------------------------
                        print("更新完了_更新対象年度無し")
                        return False, "更新対象年度無し", "", ""
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def ZaisanUpdate(Job, Exc):
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
