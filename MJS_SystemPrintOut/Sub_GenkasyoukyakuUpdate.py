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
        global URL, ThisNo, ThisYear, ThisMonth
        URL = Job.PrintOut_url + r"\\GenkaSyoukyaku"
        # 減価償却フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\G_SyoukyakuFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        time.sleep(1)
        # 年度を全てに指定----------------------------------------------------
        IC2 = RPA.ImgCheck(URL, r"\Nendo_All.png", 0.9, 10)
        if IC2[0] is True:
            print("年度選択OK")
        # ------------------------------------------------------------------
        # 関与先コード入力ボックスをクリック------------------------------------
        RPA.ImgClick(URL, r"\K_NoBox.png", 0.9, 10)
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
        time.sleep(1)
        if str(Exc.row_kanyo_no) == ThisNo:
            if str(int(Exc.year) - 1) != ThisYear:
                return False, "年度なし", ThisYear, "NoData"
            print("関与先あり")
            pg.press(["return", "return", "return"])
            # 減価償却メニューが表示されるまで待機------------------------------------
            while (
                pg.locateOnScreen(URL + r"\G_SyoukyakuMenu.png", confidence=0.9) is None
            ):
                time.sleep(1)
                # アップデート情報画面が出たら閉じる-------------------------------
                GSUM = RPA.ImgCheck(URL, r"\G_SyoukyakuUpMsg.png", 0.9, 1)
                if GSUM[0] is True:
                    RPA.ImgClick(URL, r"\G_SyoukyakuUpMsgCansel.png", 0.9, 10)
                # 顧問先情報更新ダイアログ確認-----------------------------------------
                KK = RPA.ImgCheck(URL, r"\KomonKoushin.png", 0.9, 10)
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
                # 参照表示確認ダイアログ確認------------------------------------------
                KK = RPA.ImgCheck(URL, r"\KSansyouQ.png", 0.9, 10)
                if KK[0] is True:
                    pg.press("return")
                DL = RPA.ImgCheck(URL, r"\DLCheck.png", 0.9, 10)
                if DL[0] is True:
                    pg.press("return")
                    while (
                        pg.locateOnScreen(URL + r"\K_TaisyouMenu.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        # アップデート情報画面が出たら閉じる-------------------------------
                        GSUM = RPA.ImgCheck(URL, r"\G_SyoukyakuUpMsg.png", 0.9, 1)
                        if GSUM[0] is True:
                            RPA.ImgClick(URL, r"\G_SyoukyakuUpMsgCansel.png", 0.9, 10)
            # --------------------------------------------------------------------
            RPA.ImgClick(URL, r"\G_Insatu.png", 0.9, 10)  # 2.印刷処理アイコンをクリック
            # 印刷処理メニューが表示されるまで待機------------------------------------
            while pg.locateOnScreen(URL + r"\G_InsatuFlag.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            if Exc.PN == "固定資産台帳":
                GIF = RPA.ImgCheckForList(
                    URL,
                    [
                        r"\01G_Uchiwake.png",
                        r"\01G_Uchiwake2.png",
                    ],
                    0.9,
                    10,
                )
                if GIF[0] is True:
                    RPA.ImgClick(URL, GIF[1], 0.9, 10)
                # 出力条件ウィンドウが表示されるまで待機---------------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\01G_PrintWait.png",
                        confidence=0.9,
                    )
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(1)
                RPA.ImgClick(URL, r"\01G_PrintOK.png", 0.9, 10)  # 出力条件設定OKをクリック
                # --------------------------------------------------------------------
                NOB = False, ""
                NOB = RPA.ImgCheck(URL, r"\G_NOB.png", 0.9, 10)
                if NOB[0] is True:
                    pg.press("return")
                    # 出力条件ウィンドウが表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\01G_PrintWait.png",
                            confidence=0.9,
                        )
                        is None
                    ):
                        time.sleep(1)
                    RPA.ImgClick(URL, r"\G_NOBCan.png", 0.9, 10)
                    # --------------------------------------------------------------------
                # 出力条件ウィンドウが表示されるなくなるまで待機--------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\01G_PrintWait.png",
                        confidence=0.9,
                    )
                    is not None
                ):
                    time.sleep(1)
                if NOB[0] is False:
                    time.sleep(3)
                    # --------------------------------------------------------------------
                    RPA.ImgClick(URL, r"\01G_PrintBtn.png", 0.9, 10)  # 印刷ボタンをクリック
                    # 印刷設定が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
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
                    time.sleep(1)
                    # 確実に閉じる---------------------------------------------------
                    ZME = RPA.ImgCheck(URL, r"\01G_End.png", 0.9, 10)
                    if ZME[0] is True:
                        RPA.ImgClick(URL, r"\01G_End.png", 0.9, 10)
                    else:
                        pg.keyDown("alt")
                        pg.press("x")
                        pg.keyUp("alt")
                    # ---------------------------------------------------------------
                    time.sleep(1)
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\01G_Uchiwake.png",
                                r"\01G_Uchiwake2.png",
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
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    # 閉じる処理--------------------------
                    pg.keyDown("alt")
                    pg.press("f4")
                    pg.keyUp("alt")
                    # -----------------------------------
                    f = 0
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    return False, "物件無し", "", ""
            elif Exc.PN == "一括償却資産":
                Nod = ""
                GIF = RPA.ImgCheckForList(
                    URL,
                    [
                        r"\03G_Meisai.png",
                        r"\03G_Meisai2.png",
                    ],
                    0.9,
                    10,
                )
                if GIF[0] is True:
                    RPA.ImgClick(URL, GIF[1], 0.9, 10)
                # 出力条件ウィンドウが表示されるまで待機---------------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\03G_PrintWait.png",
                        confidence=0.9,
                    )
                    is None
                ):
                    time.sleep(1)
                    GN = RPA.ImgCheck(URL, r"\03G_Nodata.png", 0.9, 1)
                    if GN[0] is True:
                        pg.press("return")
                        Nod = "Nodata"
                    if Nod == "Nodata":
                        break
                # --------------------------------------------------------------------
                time.sleep(1)
                if Nod == "":
                    RPA.ImgClick(URL, r"\03G_PrintOK.png", 0.9, 10)  # 出力条件設定OKをクリック
                    # 出力条件ウィンドウが表示されるなくなるまで待機--------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\03G_PrintWait.png",
                            confidence=0.9,
                        )
                        is not None
                    ):
                        time.sleep(1)
                    time.sleep(3)
                    # --------------------------------------------------------------------
                    RPA.ImgClick(URL, r"\01G_PrintBtn.png", 0.9, 10)  # 印刷ボタンをクリック
                    # 印刷設定が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        SJS = RPA.ImgCheckForList(
                            URL,
                            [r"\Nodata.png", r"\NodataQ.png"],
                            0.9,
                            1,
                        )
                        if SJS[0] is True:
                            if "NodataQ.png" in SJS[1] is True:
                                pg.press("return")
                                Nod = "Nodata"
                            else:
                                pg.press("return")
                                RPA.ImgClick(URL, r"\NodataCan.png", 0.9, 10)
                                Nod = "Nodata"
                        if Nod == "Nodata":
                            break
                    # --------------------------------------------------------------------
                if Nod == "":
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
                    time.sleep(1)
                    # 確実に閉じる---------------------------------------------------
                    ZME = RPA.ImgCheck(URL, r"\01G_End.png", 0.9, 10)
                    if ZME[0] is True:
                        RPA.ImgClick(URL, r"\01G_End.png", 0.9, 10)
                    else:
                        pg.keyDown("alt")
                        pg.press("x")
                        pg.keyUp("alt")
                    # ---------------------------------------------------------------
                    time.sleep(1)
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\01G_Uchiwake.png",
                                r"\01G_Uchiwake2.png",
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
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\01G_Uchiwake.png",
                                r"\01G_Uchiwake2.png",
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
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    return False, ThisNo, ThisYear, Nod
            elif Exc.PN == "少額償却資産":
                Nod = ""
                GIF = RPA.ImgCheckForList(
                    URL,
                    [
                        r"\04G_Syougaku.png",
                        r"\04G_Syougaku2.png",
                    ],
                    0.9,
                    10,
                )
                if GIF[0] is True:
                    RPA.ImgClick(URL, GIF[1], 0.9, 10)
                # 出力条件ウィンドウが表示されるまで待機---------------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\04G_PrintWait.png",
                        confidence=0.9,
                    )
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(1)
                RPA.ImgClick(URL, r"\04G_PrintOK.png", 0.9, 10)  # 出力条件設定OKをクリック
                # 出力条件ウィンドウが表示されるなくなるまで待機--------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\04G_PrintWait.png",
                        confidence=0.9,
                    )
                    is not None
                ):
                    time.sleep(1)
                time.sleep(3)
                GN = RPA.ImgCheck(URL, r"\01G_Nodata.png", 0.9, 10)
                if GN[0] is True:
                    pg.press("return")
                    Nod = "Nodata"
                if Nod == "":
                    # --------------------------------------------------------------------
                    RPA.ImgClick(URL, r"\01G_PrintBtn.png", 0.9, 10)  # 印刷ボタンをクリック
                    # 印刷設定が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        SJS = RPA.ImgCheck(URL, r"\Nodata.png", 0.9, 10)
                        if SJS[0] is True:
                            pg.press("return")
                            RPA.ImgClick(URL, r"\NodataCan.png", 0.9, 10)
                            Nod = "Nodata"
                        if Nod == "Nodata":
                            break
                    # --------------------------------------------------------------------
                else:
                    RPA.ImgClick(URL, r"\03G_Cancel.png", 0.9, 10)
                if Nod == "":
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
                    time.sleep(1)
                    # 確実に閉じる---------------------------------------------------
                    ZME = RPA.ImgCheck(URL, r"\01G_End.png", 0.9, 10)
                    if ZME[0] is True:
                        RPA.ImgClick(URL, r"\01G_End.png", 0.9, 10)
                    else:
                        pg.keyDown("alt")
                        pg.press("x")
                        pg.keyUp("alt")
                    # ---------------------------------------------------------------
                    time.sleep(1)
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\01G_Uchiwake.png",
                                r"\01G_Uchiwake2.png",
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
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\01G_Uchiwake.png",
                                r"\01G_Uchiwake2.png",
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
                    # 減価償却フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\G_SyoukyakuFlag.png",
                            confidence=0.9,
                        )
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
                    return False, ThisNo, ThisYear, Nod
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""

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
