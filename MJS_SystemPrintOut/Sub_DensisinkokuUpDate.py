# from ctypes import windll
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
    概要: 電子申告同意書印刷処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param Job.PrintOut_url : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        global URL, ThisNo, ThisYear, ThisMonth
        URL = Job.PrintOut_url + r"\\DensisinkokuUpDate"
        # 電子申告フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\DensiFlag.png", confidence=0.9) is None:
            time.sleep(1)
            UPT = RPA.ImgCheck(
                URL, r"\14D_UpDateText.png", 0.9, 1
            )  # アップデート情報ウィンドウがあるかチェック
            if UPT[0] is True:
                RPA.ImgClick(URL, r"\14D_UDTClose.png", 0.9, 10)  # 閉じるをクリック
        # ------------------------------------------------------------------
        time.sleep(1)
        # 利用同意書アイコンを探す-------------------------------------------------
        DDR = RPA.ImgCheckForList(URL, [r"\14Doui.png", r"\14Doui2.png"], 0.9, 10)
        # -----------------------------------------------------------------------
        if DDR[0] is True:  # 利用同意書アイコンがあれば
            RPA.ImgClick(URL, DDR[1], 0.9, 10)  # 利用同意書アイコンをクリック
        # 利用同意書メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\14D_Print.png", confidence=0.9) is None:
            time.sleep(1)
            # ------------------------------------------------------------------
            RPA.ImgClick(URL, r"\14D_Search.png", 0.9, 10)  # 利用同意書アイコンをクリック
        # 検索メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\14D_SearchFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        # キャンセルボタンが選択されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\CanselBtn.png", confidence=0.99999) is None:
            time.sleep(1)
            pg.keyDown("shift")
            pg.press("tab")
            pg.keyUp("shift")
        # ------------------------------------------------------------------
        # 関与先コードで指定--------------------------------------------------
        pyperclip.copy(str(Exc.row_kanyo_no))
        pg.press("tab")
        time.sleep(1)
        pg.hotkey("ctrl", "v")
        time.sleep(1)
        pg.press(["return", "return"])
        RPA.ImgClick(URL, r"\14D_SearchOK.png", 0.9, 10)  # OKボタンをクリック
        # ------------------------------------------------------------------
        # チェックボックスが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\14D_CheckBox.png", confidence=0.9) is None:
            time.sleep(1)
            # ------------------------------------------------------------------
        # チェックボックスが表示されなくなるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\14D_CheckBox.png", confidence=0.9) is not None:
            time.sleep(1)
            # ------------------------------------------------------------------
            RPA.ImgClick(URL, r"\14D_CheckBox.png", 0.9, 10)  # チェックボックスをクリック
        time.sleep(1)
        RPA.ImgClick(URL, r"\14D_Print.png", 0.9, 10)  # 印刷開始ボタンをクリック
        # 印刷設定が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9) is None:
            time.sleep(1)
            RNO = RPA.ImgCheck(URL, r"\RiyouNodata.png", 0.9, 10)
            if RNO[0] is True:

                time.sleep(1)
                #  確実に閉じる--------------------------------------------
                DED = RPA.ImgCheck(URL, r"\14D_End.png", 0.9, 10)
                if DED[0] is True:
                    RPA.ImgClick(URL, r"\14D_End.png", 0.9, 10)
                else:
                    pg.keyDown("alt")
                    pg.press("x")
                    pg.keyUp("alt")
                # ---------------------------------------------------------
                time.sleep(1)
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [
                            r"\14Doui.png",
                            r"\14Doui2.png",
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
                # 電子申告フラグが表示されるまで待機------------------------------------
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [
                            r"\DensiIcon.png",
                            r"\DensiIcon2.png",
                        ],
                        0.9,
                        1,
                    )[0]
                    is False
                ):
                    time.sleep(1)
                    f += 1
                    if f == 5:
                        pg.keyDown("alt")
                        pg.press("f4")
                        pg.keyUp("alt")
                        f = 0
                # ------------------------------------------------------------------

                return False, "電子申告起動失敗", "", ""
        # --------------------------------------------------------------------
        # 申告税一覧表印刷処理----------------------------------------------------
        FO = RPA.ImgCheckForList(
            URL,
            [
                r"\FileOut.png",
                r"\FileOut2.png",
            ],
            0.9,
            1,
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
        while pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9) is not None:
            time.sleep(1)
            FO = RPA.ImgCheck(URL, r"\FileOver.png", 0.9, 1)
            if FO[0] is True:
                pg.press("y")
        # --------------------------------------------------------------------
        time.sleep(1)
        #  確実に閉じる--------------------------------------------
        DED = RPA.ImgCheck(URL, r"\14D_End.png", 0.9, 10)
        if DED[0] is True:
            RPA.ImgClick(URL, r"\14D_End.png", 0.9, 10)
        else:
            pg.keyDown("alt")
            pg.press("x")
            pg.keyUp("alt")
        # ---------------------------------------------------------
        time.sleep(1)
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\14Doui.png",
                    r"\14Doui2.png",
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
        # 電子申告フラグが表示されるまで待機------------------------------------
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\DensiIcon.png",
                    r"\DensiIcon2.png",
                ],
                0.9,
                1,
            )[0]
            is False
        ):
            time.sleep(1)
            f += 1
            if f == 5:
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                f = 0
        # ------------------------------------------------------------------
        return True, str(Exc.row_kanyo_no), "ThisYear", "ThisMonth"
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def DensisinkokuUpDate(Job, Exc):
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
