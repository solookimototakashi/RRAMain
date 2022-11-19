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
    概要: 所得税更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\SyotokuzeiUpdate"
        ErrStr = ""  # Rpaエラー判別変数

        # 所得税フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\SyotokuFlag.png", confidence=0.9) is None:
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
            pg.press("return")
            pg.keyDown("shift")
            pg.press(["tab", "tab", "tab"])
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
            pg.press("return")
            pg.keyDown("shift")
            pg.press(["tab", "tab", "tab"])
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
        YearDiff = Job.Start_Year - int(ThisYear)
        if abs(YearDiff) > Job.Start_Year:
            YearDiff = 32 - int(ThisYear) + Job.Start_Year - 2
        # 他システムとメニューが違う-------------------------------------------------------
        if str(Exc.row_data["関与先番号"]) == ThisNo:
            print("関与先あり")
            if YearDiff == 1:  # 次年度更新か判定
                pg.press(["return", "return", "return"])
                time.sleep(1)
                # 所得税メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuMenu.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                    # 顧問先情報変更ダイアログが表示されたら
                    CDQ = RPA.ImgCheck(
                        URL,
                        r"ChangeDataQ.png",
                        0.9,
                        1,
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
                            1,
                        )
                        RPA.ImgClick(URL, CDB[1], 0.9, 10)  # 顧問先情報取込ボタンをクリック
                    # 自治体情報変更ダイアログが表示されたら
                    THI = RPA.ImgCheck(
                        URL,
                        r"THI.png",
                        0.9,
                        10,
                    )
                    if THI[0] is True:
                        pg.press("return")
                # --------------------------------------------------------------------
                RPA.ImgClick(URL, r"\KojinKihon.png", 0.9, 10)  # 個人基本情報のアイコンをクリック
                SQ = RPA.ImgCheck(URL, r"SansyouQ.png", 0.9, 10)
                if SQ[0] is True:
                    pg.press("n")
                    time.sleep(1)
                    pg.press("y")
                    time.sleep(1)
                # 顧問先情報取り込みアイコンが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInIcon.png", 0.9, 10)  # 顧問先情報取り込みアイコンをクリック
                while pg.locateOnScreen(URL + r"\DataInOK.png", confidence=0.9) is None:
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInOK.png", 0.9, 10)  # 取り込むボタンをクリック
                # 一括更新のアイコンが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    # 自治体情報変更ダイアログが表示されたら
                    THI = RPA.ImgCheck(
                        URL,
                        r"THI.png",
                        0.9,
                        10,
                    )
                    if THI[0] is True:
                        pg.press("return")
                        time.sleep(1)
                        pg.press("return")
                        time.sleep(1)
                    SCH = RPA.ImgCheck(
                        URL,
                        r"SyotokuCheck.png",
                        0.9,
                        10,
                    )
                    if SCH[0] is True:
                        pg.press("return")
                        time.sleep(1)
                        pg.press("return")
                        time.sleep(1)
                    time.sleep(3)
                    pg.keyDown("alt")
                    pg.press("x")
                    pg.keyUp("alt")
                    time.sleep(1)
                    pg.press("y")

                # ----------------------------------------------------------------------
                SK = RPA.ImgCheckForList(
                    URL, [r"\SyotokuKousin.png", r"\SyotokuKousin2.png"], 0.9, 10
                )
                if SK[0] is True:
                    RPA.ImgClick(URL, SK[1], 0.9, 10)  # 一括更新のアイコンをクリック
                # 所得税メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuKMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                while (
                    RPA.ImgCheckForList(
                        URL, [r"IkkatuFind.png", r"IkkatuFind2.png"], 0.9, 1
                    )[0]
                    is False
                ):
                    time.sleep(1)
                # 検索メニューが表示されるまでループ------------------------------------
                while pg.locateOnScreen(URL + r"\Find.png", confidence=0.9) is None:
                    time.sleep(1)
                    pg.press("return")
                    time.sleep(1)
                    pg.keyDown("alt")
                    pg.press("s")
                    pg.keyUp("alt")
                time.sleep(3)
                pyperclip.copy(str(Exc.row_data["関与先番号"]))
                pg.hotkey("ctrl", "v")
                pg.press(["return", "return"])
                time.sleep(1)
                pg.press("space")
                time.sleep(1)
                # --------------------------------------------------------------------
                SNC = RPA.ImgCheck(URL, r"SyotokuNoCalc.png", 0.9, 10)
                if SNC[0] is True:
                    ErrStr = "Nocalc"
                    pg.press("y")  # yで決定(nがキャンセル)
                time.sleep(2)
                RPA.ImgClick(URL, r"\SyotokuStart.png", 0.9, 10)  # 更新開始のアイコンをクリック
                # 確認ウィンドウが表示されるまで待機-------------------------------------
                while pg.locateOnScreen(URL + r"\SyotokuQ.png", confidence=0.9) is None:
                    time.sleep(1)
                # --------------------------------------------------------------------
                pg.press("y")  # yで決定(nがキャンセル)
                # 処理終了ウィンドウが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuEnd.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                pg.press("return")  # 決定
                # 一括更新のアイコンが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuMenu.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                    ME = RPA.ImgCheckForList(
                        URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                    )
                    if ME[0] is True:
                        RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                # --------------------------------------------------------------------
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                # 所得税フラグが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SyotokuFlag.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                # 初期画面で開封された所得税項目を閉じる----------------------------------
                HoujinList = [
                    r"\Syotoku.png",
                    r"\Syotoku2.png",
                ]
                HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                if HLI[0] is True:
                    RPA.ImgClick(URL, HLI[1], 0.9, 10)
                # --------------------------------------------------------------------
                print("更新完了")
                if ErrStr == "":
                    return True, ThisNo, ThisYear, ThisMonth
                elif ErrStr == "Nocalc":
                    return True, "Nocalc", ThisYear, ThisMonth
            elif str(Job.Start_Year) == ThisYear:
                return False, "該当年度有り", "", ""
            else:
                ID = IkkatuUpDate(Job, Exc, YearDiff)
                if ID is True:
                    return True, ThisNo, ThisYear, ThisMonth
        else:
            print("関与先なし")
            # 初期画面で開封された所得税項目を閉じる----------------------------------
            HoujinList = [
                r"\Syotoku.png",
                r"\Syotoku2.png",
            ]
            HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
            if HLI[0] is True:
                RPA.ImgClick(URL, HLI[1], 0.9, 10)
            # --------------------------------------------------------------------
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def IkkatuUpDate(Job, Exc, YearDiff):
    """
    所得税申告:該当年度まで繰り返し更新
    """
    URL = Job.imgdir_url + r"\\SyotokuzeiUpdate"
    try:
        pg.press(["return", "return", "return"])
        time.sleep(1)
        # 所得税メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\SyotokuMenu.png", confidence=0.9) is None:
            time.sleep(1)
            # 顧問先情報変更ダイアログが表示されたら
            CDQ = RPA.ImgCheck(
                URL,
                r"ChangeDataQ.png",
                0.9,
                1,
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
                    1,
                )
                RPA.ImgClick(URL, CDB[1], 0.9, 10)  # 顧問先情報取込ボタンをクリック
            # 自治体情報変更ダイアログが表示されたら
            THI = RPA.ImgCheck(
                URL,
                r"THI.png",
                0.9,
                10,
            )
            if THI[0] is True:
                pg.press("return")
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\KojinKihon.png", 0.9, 10)  # 個人基本情報のアイコンをクリック
        SQ = RPA.ImgCheck(URL, r"SansyouQ.png", 0.9, 10)
        if SQ[0] is True:
            pg.press("n")
            time.sleep(1)
            pg.press("y")
            time.sleep(1)
        # 顧問先情報取り込みアイコンが表示されるまで待機--------------------------------
        while pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None:
            time.sleep(1)
        RPA.ImgClick(URL, r"\DataInIcon.png", 0.9, 10)  # 顧問先情報取り込みアイコンをクリック
        while pg.locateOnScreen(URL + r"\DataInOK.png", confidence=0.9) is None:
            time.sleep(1)
        RPA.ImgClick(URL, r"\DataInOK.png", 0.9, 10)  # 取り込むボタンをクリック
        # 一括更新のアイコンが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\SyotokuKousin.png", confidence=0.9) is None:
            time.sleep(1)
            # 自治体情報変更ダイアログが表示されたら
            THI = RPA.ImgCheck(
                URL,
                r"THI.png",
                0.9,
                10,
            )
            if THI[0] is True:
                pg.press("return")
                time.sleep(1)
                pg.press("return")
                time.sleep(1)
            SCH = RPA.ImgCheck(
                URL,
                r"SyotokuCheck.png",
                0.9,
                10,
            )
            if SCH[0] is True:
                pg.press("return")
                time.sleep(1)
                pg.press("return")
                time.sleep(1)
            time.sleep(3)
            pg.keyDown("alt")
            pg.press("x")
            pg.keyUp("alt")
            time.sleep(1)
            pg.press("y")
        # --------------------------------------------------------------------
        # 次年まで繰り返し
        for i in range(YearDiff):
            RPA.ImgClick(URL, r"\DataOpen_btn.png", 0.9, 10)  # データを開くをクリック
            while RPA.ImgCheck(URL, r"\DataOpen_flag.png", 0.9, 1)[0] is False:
                time.sleep(1)
            pg.write(str(Exc.row_data["関与先番号"]))
            pg.press("return")
            RPA.ImgClick(URL, r"\DataLastOpen_btn.png", 0.9, 10)  # データを開く
            while RPA.ImgCheck(URL, r"\DataOpen_flag.png", 0.9, 1)[0] is True:
                time.sleep(1)
            # ----------------------------------------------------------------------
            SK = RPA.ImgCheckForList(
                URL, [r"\SyotokuKousin.png", r"\SyotokuKousin2.png"], 0.9, 10
            )
            if SK[0] is True:
                RPA.ImgClick(URL, SK[1], 0.9, 10)  # 一括更新のアイコンをクリック
            # 所得税メニューが表示されるまで待機------------------------------------
            while pg.locateOnScreen(URL + r"\SyotokuKMenu.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            while (
                RPA.ImgCheckForList(
                    URL, [r"IkkatuFind.png", r"IkkatuFind2.png"], 0.9, 1
                )[0]
                is False
            ):
                time.sleep(1)
            # 検索メニューが表示されるまでループ------------------------------------
            while pg.locateOnScreen(URL + r"\Find.png", confidence=0.9) is None:
                time.sleep(1)
                pg.press("return")
                time.sleep(1)
                pg.keyDown("alt")
                pg.press("s")
                pg.keyUp("alt")
            time.sleep(3)
            pyperclip.copy(str(Exc.row_data["関与先番号"]))
            pg.hotkey("ctrl", "v")
            pg.press(["return", "return"])
            time.sleep(1)
            pg.press("space")
            time.sleep(1)
            # --------------------------------------------------------------------
            SNC = RPA.ImgCheck(URL, r"SyotokuNoCalc.png", 0.9, 10)
            if SNC[0] is True:
                pg.press("y")  # yで決定(nがキャンセル)
            time.sleep(2)
            RPA.ImgClick(URL, r"\SyotokuStart.png", 0.9, 10)  # 更新開始のアイコンをクリック
            # 確認ウィンドウが表示されるまで待機-------------------------------------
            while pg.locateOnScreen(URL + r"\SyotokuQ.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("y")  # yで決定(nがキャンセル)
            # 処理終了ウィンドウが表示されるまで待機----------------------------------
            while pg.locateOnScreen(URL + r"\SyotokuEnd.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("return")  # 決定
            # 一括更新のアイコンが表示されるまで待機----------------------------------
            while pg.locateOnScreen(URL + r"\SyotokuMenu.png", confidence=0.9) is None:
                time.sleep(1)
                ME = RPA.ImgCheckForList(
                    URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                )
                if ME[0] is True:
                    RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
            # --------------------------------------------------------------------
        # 閉じる処理--------------------------
        pg.keyDown("alt")
        pg.press("f4")
        pg.keyUp("alt")
        # -----------------------------------
        # 所得税フラグが表示されるまで待機-------------------------------------
        while pg.locateOnScreen(URL + r"\SyotokuFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # --------------------------------------------------------------------
        # 初期画面で開封された所得税項目を閉じる----------------------------------
        HoujinList = [
            r"\Syotoku.png",
            r"\Syotoku2.png",
        ]
        HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
        if HLI[0] is True:
            RPA.ImgClick(URL, HLI[1], 0.9, 10)
        # --------------------------------------------------------------------
        print("更新完了")
        return True
    except:
        return False


# ------------------------------------------------------------------------------------------------------------------
def SyotokuzeiUpdate(Job, Exc):
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
