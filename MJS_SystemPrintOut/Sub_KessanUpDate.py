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
    概要: 決算内訳書更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        global URL, ThisNo, ThisYear, ThisMonth
        URL = Job.PrintOut_url + r"\\KessanUpDate"
        # 決算内訳書フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\Kessan_CFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        time.sleep(1)
        # 年度を全てに指定----------------------------------------------------
        IC2 = RPA.ImgCheck(URL, r"\Nendo_All.png", 0.9, 10)
        if IC2[0] is True:
            print("年度選択OK")
        # ------------------------------------------------------------------
        # 関与先コード入力ボックスをクリック------------------------------------
        Nob = RPA.ImgCheckForList(URL, [r"\K_NoBox.png", r"\K_AfterNoBox.png"], 0.9, 10)
        if Nob[0] is True:
            RPA.ImgClick(URL, Nob[1], 0.9, 10)
            while pg.locateOnScreen(URL + r"\K_AfterNoBox.png", confidence=0.9) is None:
                time.sleep(1)
        time.sleep(1)
        pg.write(str(Exc.row_kanyo_no))
        pg.press("return")
        pg.write(str(int(Exc.year) - 1))
        pg.press(["return", "return", "return"])
        pg.keyDown("shift")
        pg.press(["tab", "tab"])
        pg.keyUp("shift")
        if RPA.ImgCheck(URL, r"\NotData.png", 0.9, 10)[0] is False:
            # 入力した関与先コードを取得------------
            pg.keyDown("shift")
            pg.press(["tab", "tab"])
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
            # -----------------------------
            # ------
        else:
            return False, "NoData", ThisYear, "NoData"
        time.sleep(1)
        if str(Exc.row_kanyo_no) == ThisNo:
            if str(int(Exc.year) - 1) != ThisYear:
                return False, "年度なし", ThisYear, "NoData"
            print("関与先あり")
            pg.press(["return", "return", "return"])
            # 決算内訳書メニューが表示されるまで待機------------------------------------
            while pg.locateOnScreen(URL + r"\KessanMenu.png", confidence=0.9) is None:
                time.sleep(1)
                # 顧問先情報更新ダイアログ確認------------------------------------------
                KK = RPA.ImgCheck(URL, r"\KomonKoushin.png", 0.9, 1)
                if KK[0] is True:
                    pg.press("y")
                    while (
                        pg.locateOnScreen(
                            URL + r"\KomonKoushinBar.png",
                            confidence=0.9,
                        )
                        is None
                    ):
                        time.sleep(1)
                    RPA.ImgClick(URL, r"\KomonKoushinInput.png", 0.9, 10)
                DL = RPA.ImgCheck(URL, r"\DLCheck.png", 0.9, 10)
                if DL[0] is True:
                    pg.press("return")

            # --------------------------------------------------------------------
            # 内訳書印刷メニューが表示されるまで待機----------------------------------
            while pg.locateOnScreen(URL + r"\11U_Flag.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            # 内訳書印刷のアイコンを探す-------------------------------------------------
            ImgList = [
                r"\11U_Uchiwake.png",
                r"\11U_Uchiwake2.png",
            ]
            ICFL = RPA.ImgCheckForList(URL, ImgList, 0.9, 10)
            # -----------------------------------------------------------------------
            if ICFL[0] is True:
                RPA.ImgClick(URL, ICFL[1], 0.9, 10)  # 内訳書印刷アイコンをクリック
            # 印刷ボタンが表示されるまで待機---------------------------------
            while pg.locateOnScreen(URL + r"\11U_PrintBtn.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            RPA.ImgClick(URL, r"\11U_PrintBtn.png", 0.9, 10)  # 印刷ボタンをクリック
            PSQ = False, ""
            # 印刷設定が表示されるまで待機---------------------------------
            while pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9) is None:
                time.sleep(1)
                PSQ = RPA.ImgCheck(URL, r"\PrintStyleQ.png", 0.9, 1)
                if PSQ[0] is True:
                    pg.press("return")
                    break
            # --------------------------------------------------------------------
            if PSQ[0] is False:
                # 申告税一覧表印刷処理----------------------------------------------------
                FO = RPA.ImgCheckForList(
                    URL,
                    [
                        r"\FileOut.png",
                        r"\FileOut2.png",
                    ],
                    0.9,
                    10,
                )
                if FO[0] is True:
                    RPA.ImgClick(URL, FO[1], 0.9, 10)
                RPA.ImgClick(URL, r"\PDFBar.png", 0.9, 10)
                pg.press("return")
                pg.press("delete")
                pg.press("backspace")
                time.sleep(1)
                pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
                time.sleep(1)
                pg.hotkey("ctrl", "v")
                pg.press("return")
                time.sleep(1)
                RPA.ImgClick(URL, r"\PrintOut.png", 0.9, 10)
                time.sleep(2)
                #  印刷設定が表示されなくなるまで待機---------------------------------
                while (
                    pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                    FO = RPA.ImgCheck(URL, r"\FileOver.png", 0.9, 1)
                    if FO[0] is True:
                        pg.press("y")
                # --------------------------------------------------------------------
                #  印刷中が表示されるまで待機-------------------------------------------
                IC = 0
                while pg.locateOnScreen(URL + r"\NowPrint.png", confidence=0.9) is None:
                    time.sleep(1)
                    IC += 1
                    if IC == 5:
                        pg.press("tab")
                        break
                #  印刷中が表示されなくなるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NowPrint.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                #  確実に閉じる--------------------------------------------------------
                UED = RPA.ImgCheck(URL, r"\11U_End.png", 0.9, 10)
                if UED[0] is True:
                    RPA.ImgClick(URL, r"\11U_End.png", 0.9, 10)
                else:
                    pg.keyDown("alt")
                    pg.press("x")
                    pg.keyUp("alt")
                # --------------------------------------------------------------------
                time.sleep(1)
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [
                            r"\11U_Uchiwake.png",
                            r"\11U_Uchiwake2.png",
                        ],
                        0.9,
                        1,
                    )[0]
                    is False
                ):
                    time.sleep(1)
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                f = 0
                # 決算内訳書フラグが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\Kessan_CFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    f += 1
                    if f == 5:
                        pg.keyDown("alt")
                        pg.press("f4")
                        pg.keyUp("alt")
                        f = 0
                # ------------------------------------------------------------------
                return True, ThisNo, ThisYear, ThisMonth
            else:
                #  確実に閉じる--------------------------------------------------------
                UED = RPA.ImgCheck(URL, r"\11U_End.png", 0.9, 10)
                if UED[0] is True:
                    RPA.ImgClick(URL, r"\11U_End.png", 0.9, 10)
                else:
                    pg.keyDown("alt")
                    pg.press("x")
                    pg.keyUp("alt")
                # --------------------------------------------------------------------
                time.sleep(1)
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [
                            r"\11U_Uchiwake.png",
                            r"\11U_Uchiwake2.png",
                        ],
                        0.9,
                        1,
                    )[0]
                    is False
                ):
                    time.sleep(1)
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                f = 0
                # 決算内訳書フラグが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\Kessan_CFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    f += 1
                    if f == 5:
                        pg.keyDown("alt")
                        pg.press("f4")
                        pg.keyUp("alt")
                        f = 0
                # ------------------------------------------------------------------
                return False, "印刷様式未設定", "", ""
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def KessanUpDate(Job, Exc):
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
