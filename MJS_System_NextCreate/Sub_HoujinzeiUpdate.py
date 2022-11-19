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
    概要: 法人税更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\HoujinzeiUpdate"
        ErrStr = ""  # Rpaエラー判別変数

        # 法人税フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9) is None:
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
                # 法人税メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\HoujinzeiMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    HQ = RPA.ImgCheck(URL, r"\HoujinOpenQ.png", 0.9, 10)
                    if HQ[0] is True:
                        RPA.ImgClick(URL, r"\HoujinOpenQCansel.png", 0.9, 10)
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
                RPA.ImgClick(
                    URL, r"\IkkatsuHoujinKousin.png", 0.9, 10
                )  # 一括更新のアイコンをクリック
                # 法人税メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\HoujinzeiMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                # 再計算ウィンドウが表示されていないかチェック-----------------------------
                DRC = RPA.ImgCheck(URL, r"\DataReCalc.png", 0.9, 10)
                if DRC[0] is True:
                    pg.press("return")
                    ErrStr = "ReCalc"  # Rpaエラー判別変数
                # ------------------------------------------------------------------
                if ErrStr == "":
                    ListCheck = False  # 申告種類判定変数
                    # Excelの値から申告種類を判定------------------------------------
                    try:
                        ExPar = str(Exc.row_data["法人税申告書_繰越対象"])
                        if ExPar == "1":
                            TaxPngName = r"\KakuteiNext.png"
                        elif ExPar == "2":
                            TaxPngName = r"\CyuukanNext.png"
                        elif ExPar == "3":
                            TaxPngName = r"\YoteiNext.png"
                        elif ExPar == "4":
                            TaxPngName = r"\SyuuseiThis.png"
                        else:
                            TaxPngName = "Err"
                    except:
                        TaxPngName = "Err"
                    time.sleep(1)
                    if TaxPngName == "Err":
                        # 確認ウィンドウが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\HoujinEndQ.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                            # --------------------------------------------------------------------
                            ME = RPA.ImgCheckForList(
                                URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                            )
                            if ME[0] is True:
                                RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                        # --------------------------------------------------------------------
                        pg.press("y")
                        # 法人税メニューが表示されるまで待機------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\HoujinzeiMenu.png", confidence=0.9
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
                        # 法人税フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        # 初期画面で開封された法人税項目を閉じる----------------------------------
                        HoujinList = [
                            r"\FastMenuHoujinzei.png",
                            r"\FastMenuHoujinzei2.png",
                        ]
                        HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                        if HLI[0] is True:
                            RPA.ImgClick(URL, HLI[1], 0.9, 10)
                        # --------------------------------------------------------------------
                        print("更新完了")

                        return False, "要申告指定", "", ""
                    else:
                        # 申告種類に応じたリスト画像までループ----------------------------
                        while ListCheck is False:
                            pg.press("down")
                            TPN = RPA.ImgCheck(URL, TaxPngName, 0.9, 10)
                            if TPN[0] is True:
                                ListCheck = True
                                pg.press("return")
                        # ------------------------------------------------------------
                        time.sleep(1)
                        RPA.ImgClick(
                            URL, r"\HoujinStart.png", 0.9, 10
                        )  # 更新開始のアイコンをクリック
                        # 確認ウィンドウが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\HoujinQ.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        pg.press("y")  # yで決定(nがキャンセル)
                        # 処理終了ウィンドウが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\HoujinEnd.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        pg.press("return")  # 決定
                        # 一括更新のアイコンが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\HoujinzeiMenu.png", confidence=0.9
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
                        # 法人税フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                            HEQ2 = RPA.ImgCheck(URL, r"\HoujinEndQ2.png", 0.9, 10)
                            if HEQ2[0] is True:
                                pg.press("y")
                        # --------------------------------------------------------------------
                        # 初期画面で開封された法人税項目を閉じる----------------------------------
                        HoujinList = [
                            r"\FastMenuHoujinzei.png",
                            r"\FastMenuHoujinzei2.png",
                        ]
                        HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                        if HLI[0] is True:
                            RPA.ImgClick(URL, HLI[1], 0.9, 10)
                        # --------------------------------------------------------------------
                        print("更新完了")
                        if ErrStr == "":
                            return True, ThisNo, ThisYear, ThisMonth
                elif ErrStr == "ReCalc":
                    # 閉じる処理--------------------------
                    pg.keyDown("alt")
                    pg.press("f4")
                    pg.keyUp("alt")
                    # -----------------------------------
                    # 法人税フラグが表示されるまで待機-------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                    # 初期画面で開封された法人税項目を閉じる----------------------------------
                    HoujinList = [r"\FastMenuHoujinzei.png", r"\FastMenuHoujinzei2.png"]
                    HLI = RPA.ImgCheckForList(URL, HoujinList, 0.9, 10)
                    if HLI[0] is True:
                        RPA.ImgClick(URL, HLI[1], 0.9, 10)
                    # --------------------------------------------------------------------
                    return False, "要データ再計算", "", ""
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdate(Job, Exc):
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
