from ctypes import windll
import pyautogui as pg
import time
import RPA_Function as RPA
import pyperclip  # クリップボードへのコピーで使用
import MJSSPOPDFMarge as PDFM
import wrapt_timeout_decorator

# TIMEOUT = 600

# ------------------------------------------------------------------------------------------------------------------
# @wrapt_timeout_decorator.timeout(dec_timeout=TIMEOUT)
def Flow(Job, Exc):
    """
    概要: 法人税更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param Job.PrintOut_url : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        global URL, ThisNo, ThisYear, ThisMonth
        URL = Job.PrintOut_url + r"\\HoujinzeiUpdate"
        # 法人税フラグが表示されるまで待機---------------------------------n---
        while pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # ------------------------------------------------------------------
        time.sleep(1)
        # 他システムとメニューが違う-------------------------------------------------------
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
        pg.write(str(Exc.year))
        pg.press(["return", "return", "return"])
        pg.keyDown("shift")
        pg.press(["tab", "tab"])
        pg.keyUp("shift")
        # 申告種類が確定申告になっているか確認-----------------------------------
        KF = RPA.ImgCheck(URL, r"\\KakuteiFlag.png", 0.9, 10)
        if KF[0] is False:
            while RPA.ImgCheck(URL, r"\\KakuteiFlag.png", 0.9, 1)[0] is False:
                RPA.ImgClick(URL, r"\\SinkokuArrow.png", 0.9, 10)
                pg.press("down")
                pg.press("return")
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
        if str(Exc.row_kanyo_no) == ThisNo:
            if str(Exc.year) != ThisYear:
                return False, "年度なし", ThisYear, "NoData"
            print("関与先あり")
            pg.press(["return", "return", "return"])
            # 法人税メニューが表示されるまで待機------------------------------------
            while (
                pg.locateOnScreen(URL + r"\HoujinzeiMenu.png", confidence=0.9) is None
            ):
                time.sleep(1)
                HQ = RPA.ImgCheck(URL, r"\HoujinOpenQ.png", 0.9, 1)
                if HQ[0] is True:
                    RPA.ImgClick(URL, r"\HoujinOpenQCansel.png", 0.9, 10)
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
                # 新規別表追加選択-------------------------------------------------
                SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 10)
                if SB[0] is True:
                    RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
                DL = RPA.ImgCheck(URL, r"\DLCheck.png", 0.9, 10)
                if DL[0] is True:
                    pg.press("return")
                    while (
                        pg.locateOnScreen(URL + r"\K_TaisyouMenu.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)

            # 指示内容で処理分け----------------------------------------------------------
            if (
                Exc.PN == "申告税一覧表"
                or Exc.PN == "第6号様式（県）"
                or Exc.PN == "第6号様式別表9（県）"
                or Exc.PN == "第20号様式（市）"
                or Exc.PN == "第22号の2様式"
                or Exc.PN == "別表１　緑色"
            ):
                PNFlow = HoujinzeiUpdateSinkokuItiran(Job, Exc)
                return PNFlow[0], PNFlow[1], PNFlow[2], PNFlow[3]
            elif Exc.PN == "税務代理権限書":
                PNFlow = HoujinzeiUpdateZeimuDairi(Job, Exc)
                return PNFlow[0], PNFlow[1], PNFlow[2], PNFlow[3]
            elif Exc.PN == "書面添付　法人税":
                PNFlow = HoujinzeiUpdateSyomen(Job, Exc)
                return PNFlow[0], PNFlow[1], PNFlow[2], PNFlow[3]
            elif Exc.PN == "別表2-16":
                PNFlow = HoujinzeiUpdateBeppyou(Job, Exc)
                return PNFlow[0], PNFlow[1], PNFlow[2], PNFlow[3]
            elif Exc.PN == "法人事業概況説明書":
                PNFlow = HoujinzeiUpdateGaikyou(Job, Exc)
                return PNFlow[0], PNFlow[1], PNFlow[2], PNFlow[3]
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdateSinkokuItiran(Job, Exc):
    # 申告税一覧表印刷処理----------------------------------------------------
    OP = RPA.ImgCheckForList(
        URL,
        [
            r"\\01SinkokuNyuuryoku.png",
            r"\\01SinkokuNyuuryoku2.png",
        ],
        0.9,
        10,
    )
    if OP[0] is True:
        RPA.ImgClick(URL, OP[1], 0.9, 10)
        ll = 0
        # 法人税メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\SinkokuTab.png", confidence=0.9) is None:
            time.sleep(1)
            SYB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 1)
            if SYB[0] is True:
                time.sleep(1)
                SBP = RPA.ImgCheckForList(
                    URL,
                    [r"\\Saiyou_Btn.png", r"\\Saiyou_Btn2.png"],
                    0.9,
                    1,
                )
                if SBP[0] is True:
                    RPA.ImgClick(URL, SBP[1], 0.9, 1)
                    # 確実に閉じる---------------------------------------------------------
                    close()
            Nodatacheck()
            SCB = RPA.ImgCheck(URL, r"\\Saiyou_Check.png", 0.9, 1)
            if SCB[0] is True:
                pg.press("y")
                while (
                    pg.locateOnScreen(URL + r"\\Saiyou_Check.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
            ll += 1
            if ll == 5:
                RPA.ImgClick(URL, OP[1], 0.9, 10)
                ll = 0
        # --------------------------------------------------------------------
        if Exc.PN == "申告税一覧表":
            RPA.ImgClick(URL, r"\\DownPrint.png", 0.9, 10)
            time.sleep(1)
            pg.press("n")
            # 一覧表メニューが表示されるまで待機------------------------------------
            while pg.locateOnScreen(URL + r"\\ItiranFlag.png", confidence=0.9) is None:
                time.sleep(1)
                SRC = RPA.ImgCheck(URL, r"\\S_RendouCheck.png", 0.9, 1)
                if SRC[0] is True:
                    pg.press("return")
                SRCCCCC = RPA.ImgCheck(URL, r"\\S_RendouCheck2.png", 0.9, 1)
                if SRCCCCC[0] is True:
                    pg.press("return")
                SRCCCCCC = RPA.ImgCheck(URL, r"\\S_RendouCheck3.png", 0.9, 1)
                if SRCCCCCC[0] is True:
                    pg.press("return")
                SRCCCCCCC = RPA.ImgCheck(URL, r"\\S_RendouCheck4.png", 0.9, 1)
                if SRCCCCCCC[0] is True:
                    pg.press("return")
                SRCC = RPA.ImgCheck(URL, r"\\S_Rendou2.png", 0.9, 1)
                if SRCC[0] is True:
                    pg.press("n")
                SRCCC = RPA.ImgCheck(URL, r"\\S_Rendou3.png", 0.9, 1)
                if SRCCC[0] is True:
                    pg.press("n")
                SRCCCC = RPA.ImgCheck(URL, r"\\S_Rendou4.png", 0.9, 1)
                if SRCCCC[0] is True:
                    pg.press("y")
                # 新規別表追加選択-------------------------------------------------
                SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 1)
                if SB[0] is True:
                    RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
            # --------------------------------------------------------------------
            RPA.ImgClick(URL, r"\\Print.png", 0.9, 10)
            # 一覧表出力項目指定が表示されるまで待機---------------------------------
            while (
                pg.locateOnScreen(URL + r"\\ItiranSyutuQ.png", confidence=0.9) is None
            ):
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("return")
            # 一覧表出力項目指定が表示されるまで待機---------------------------------
            while pg.locateOnScreen(URL + r"\\PrintBar.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            # 申告税一覧表印刷処理----------------------------------------------------
            FO = RPA.ImgCheckForList(
                URL,
                [
                    r"\\FileOut.png",
                    r"\\FileOut2.png",
                ],
                0.9,
                10,
            )
            if FO[0] is True:
                RPA.ImgClick(URL, FO[1], 0.9, 10)
            RPA.ImgClick(URL, r"\\PDFBar.png", 0.9, 10)
            pg.press("return")
            pg.press("delete")
            pg.press("backspace")
            time.sleep(1)
            pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
            time.sleep(1)
            pg.hotkey("ctrl", "v")
            pg.press("return")
            time.sleep(1)
            RPA.ImgClick(URL, r"\\PrintOut.png", 0.9, 10)
            # 印刷中が表示されるまで待機---------------------------------
            IC = 0
            while pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9) is None:
                time.sleep(1)
                IC += 1
                if IC == 5:
                    pg.press("tab")
                    break
                FO = RPA.ImgCheck(URL, r"\\FileOver.png", 0.9, 1)
                if FO[0] is True:
                    pg.press("y")
            # --------------------------------------------------------------------
            # 印刷中が表示されなくなるまで待機---------------------------------
            while (
                pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9) is not None
            ):
                time.sleep(1)
            # --------------------------------------------------------------------
            time.sleep(1)
            pg.keyDown("alt")
            pg.press("x")
            pg.keyUp("alt")
            time.sleep(3)
            # 確実に閉じる---------------------------------------------------------
            close()
            while (
                RPA.ImgCheckForList(
                    URL,
                    [
                        r"\\01SinkokuNyuuryoku.png",
                        r"\\01SinkokuNyuuryoku2.png",
                    ],
                    0.9,
                    1,
                )[0]
                is False
            ):
                time.sleep(1)
                close()
            sinkokuend(Job)

            return True, ThisNo, ThisYear, ThisMonth
        elif (
            Exc.PN == "第6号様式（県）"
            or Exc.PN == "第6号様式別表9（県）"
            or Exc.PN == "第20号様式（市）"
            or Exc.PN == "第22号の2様式"
            or Exc.PN == "別表１　緑色"
        ):
            RPA.ImgClick(URL, r"\\HoujinzeiSelecter.png", 0.9, 10)
            # 申告書印刷メニューが表示されるまで待機------------------------------------
            while pg.locateOnScreen(URL + r"\\BeppyouIcon.png", confidence=0.9) is None:
                time.sleep(1)
                SLQ = RPA.ImgCheck(URL, r"\\SelectQ.png", 0.9, 10)
                if SLQ[0] is True:
                    RPA.ImgClick(URL, r"\\Kousin.png", 0.9, 10)
            # --------------------------------------------------------------------
            if Exc.PN == "第6号様式（県）":
                RPA.ImgClick(URL, r"\\HSCancel.png", 0.9, 10)
                time.sleep(1)
                pg.write("0600")
                pg.press("return")
            elif Exc.PN == "第6号様式別表9（県）":
                RPA.ImgClick(URL, r"\\HSCancel.png", 0.9, 10)
                time.sleep(1)
                pg.write("0690")
                pg.press("return")
            elif Exc.PN == "第20号様式（市）":
                RPA.ImgClick(URL, r"\\HSCancel.png", 0.9, 10)
                time.sleep(1)
                pg.write("2000")
                pg.press("return")
            elif Exc.PN == "第22号の2様式":
                RPA.ImgClick(URL, r"\\HSCancel.png", 0.9, 10)
                time.sleep(1)
                pg.write("2202")
                pg.press("return")
            elif Exc.PN == "別表１　緑色":
                pg.press("home")
                pg.press("return")

            if Exc.PN == "第20号様式（市）":
                # 市町村の数だけループ-------------------------------------------------------------
                # 画像に名前を付ける
                SSN = "20ScreenShot.png"
                SSN2 = "20ScreenShot2.png"
                Win = "Window.png"
                while RPA.ImgCheck(URL, r"\\All\\" + Win, 0.9, 1)[0] is True:
                    WC = RPA.ImgCheck(URL, r"\\All\\" + Win, 0.9, 10)
                    pg.click(WC[1], WC[2], button="right")
                    pg.press("x")
                time.sleep(1)
                RPA.ImgClick(URL, r"\\PreviewIcon.png", 0.9, 1)
                # プレビュー画面が表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\\HoujinOpen.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(3)
                # ロード完了まで待機----------------------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\\PreviewLoad.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                TPI = RPA.ImgCheckForList(
                    URL,
                    [r"\\ThisPIcon.png", r"\\ThisPIcon2.png"],
                    0.9,
                    10,
                )  # 現在項印刷アイコンを探す
                if TPI[0] is True:  # 現在項印刷アイコンがあれば
                    # pg.keyDown("alt")
                    # pg.press("c")
                    # pg.keyup("alt")
                    # RPA.ImgClick(Job.PrintOut_url, r"\HoujinFlag.png", 0.9, 10)
                    RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                    # 印刷ダイアログ待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\\ThisPMenu.png",
                            confidence=0.9,
                        )
                        is None
                    ):
                        time.sleep(1)
                        RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
                    time.sleep(1)
                    pg.hotkey("ctrl", "v")
                    pg.press("return")
                    # 印刷完了まで待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\\ThisPMenu2.png",
                            confidence=0.9,
                        )
                        is not None
                    ):
                        time.sleep(1)
                        TPOQ = RPA.ImgCheck(URL, r"\\ThisPOverQ.png", 0.9, 10)  # 上書き確認
                        if TPOQ[0] is True:
                            pg.press("y")  # yで上書き
                    # -------------------------------------------------------------------
                    # 印刷完了まで待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                # 市町村の数だけループ-------------------------------------------------------------
                for CNo in range(1, 100):
                    ICFL = RPA.ImgCheckForList(
                        URL,
                        [r"\\All\\TihouKey.png", r"\\All\\TihouKey.png"],
                        0.9,
                        10,
                    )
                    if ICFL[0] is True:
                        # 一回目のスクショ----------------------------------------------
                        s = pg.screenshot(region=(820, 250, 450, 40))  # スクショ
                        s.save(URL + r"\\All\\" + SSN)  # 保存
                        # -------------------------------------------------------------
                        RPA.ImgClick(URL, ICFL[1], 0.9, 10)
                        CNo += 1
                        pg.write(str(CNo))
                        pg.press("return")
                        # 二回目以降のスクショ------------------------------------------
                        s = pg.screenshot(region=(820, 250, 450, 40))  # スクショ
                        s.save(URL + r"\\All\\" + SSN2)  # 保存
                        # -------------------------------------------------------------
                    WC = RPA.ImgCheck(URL, r"\\All\\" + SSN, 0.9, 10)
                    if WC[0] is True:
                        break
                    else:
                        time.sleep(1)
                        RPA.ImgClick(URL, r"\\PreviewIcon.png", 0.9, 10)
                        # プレビュー画面が表示されるまで待機------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\\HoujinOpen.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        time.sleep(3)
                        # ロード完了まで待機----------------------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\\PreviewLoad.png", confidence=0.9
                            )
                            is not None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        TPI = RPA.ImgCheckForList(
                            URL,
                            [r"\\ThisPIcon.png", r"\\ThisPIcon2.png"],
                            0.9,
                            10,
                        )  # 現在項印刷アイコンを探す
                        if TPI[0] is True:  # 現在項印刷アイコンがあれば
                            # pg.keyDown("alt")
                            # pg.press("c")
                            # pg.keyup("alt")
                            # RPA.ImgClick(Job.PrintOut_url, r"\HoujinFlag.png", 0.9, 10)
                            RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                            # 印刷ダイアログ待機----------------------------------------------------
                            while (
                                pg.locateOnScreen(
                                    URL + r"\\ThisPMenu.png",
                                    confidence=0.9,
                                )
                                is None
                            ):
                                time.sleep(1)
                                RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                            # --------------------------------------------------------------------
                            Exc.Fname = (
                                Exc.Fname.replace(".pdf", "") + str(CNo) + ".pdf"
                            )
                            time.sleep(1)
                            pyperclip.copy(
                                Exc.Fname.replace("\\\\", "\\").replace("/", "\\")
                            )
                            time.sleep(1)
                            pg.hotkey("ctrl", "v")
                            pg.press("return")
                            # 印刷完了まで待機----------------------------------------------------
                            while (
                                pg.locateOnScreen(
                                    URL + r"\\ThisPMenu2.png",
                                    confidence=0.9,
                                )
                                is not None
                            ):
                                time.sleep(1)
                                TPOQ = RPA.ImgCheck(
                                    URL, r"\\ThisPOverQ.png", 0.9, 10
                                )  # 上書き確認
                                if TPOQ[0] is True:
                                    pg.press("y")  # yで上書き
                            # -------------------------------------------------------------------
                            # 印刷完了まで待機----------------------------------------------------
                            while (
                                pg.locateOnScreen(
                                    URL + r"\\NowPrint.png", confidence=0.9
                                )
                                is not None
                            ):
                                time.sleep(1)
                            # --------------------------------------------------------------------
                # 確実に閉じる---------------------------------------------------------
                close()
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [
                            r"\\01SinkokuNyuuryoku.png",
                            r"\\01SinkokuNyuuryoku2.png",
                        ],
                        0.9,
                        10,
                    )[0]
                    is False
                ):
                    time.sleep(1)
                    close()
                sinkokuend(Job)

                return True, ThisNo, ThisYear, ThisMonth
                # --------------------------------------------------------------------------------
            else:
                time.sleep(1)
                RPA.ImgClick(URL, r"\\PreviewIcon.png", 0.9, 10)
                # プレビュー画面が表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\\HoujinOpen.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(3)
                # ロード完了まで待機----------------------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\\PreviewLoad.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                TPI = RPA.ImgCheckForList(
                    URL,
                    [r"\\ThisPIcon.png", r"\\ThisPIcon2.png"],
                    0.9,
                    10,
                )  # 現在項印刷アイコンを探す
                if TPI[0] is True:  # 現在項印刷アイコンがあれば
                    # pg.keyDown("alt")
                    # pg.press("c")
                    # pg.keyup("alt")
                    # RPA.ImgClick(Job.PrintOut_url, r"\HoujinFlag.png", 0.9, 10)
                    RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                    # 印刷ダイアログ待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\\ThisPMenu.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        RPA.ImgClick(URL, TPI[1], 0.9, 10)  # 現在項印刷アイコンをクリック
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
                    time.sleep(1)
                    pg.hotkey("ctrl", "v")
                    pg.press("return")
                    # 印刷完了まで待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\\ThisPMenu2.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                        TPOQ = RPA.ImgCheck(URL, r"\\ThisPOverQ.png", 0.9, 1)  # 上書き確認
                        if TPOQ[0] is True:
                            pg.press("y")  # yで上書き
                    # -------------------------------------------------------------------
                    # 印刷完了まで待機----------------------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                    # 確実に閉じる---------------------------------------------------------
                    close()
                    while (
                        RPA.ImgCheckForList(
                            URL,
                            [
                                r"\\01SinkokuNyuuryoku.png",
                                r"\\01SinkokuNyuuryoku2.png",
                            ],
                            0.9,
                            10,
                        )[0]
                        is False
                    ):
                        time.sleep(1)
                        close()
                    sinkokuend(Job)

                    return True, ThisNo, ThisYear, ThisMonth


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdateZeimuDairi(Job, Exc):
    # 税務代理権限証書印刷処理----------------------------------------------------
    OP = RPA.ImgCheckForList(
        URL,
        [
            r"\\15Zeimudairi.png",
            r"\\15Zeimudairi2.png",
        ],
        0.9,
        10,
    )
    if OP[0] is True:
        RPA.ImgClick(URL, OP[1], 0.9, 10)
        # 法人税メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\HoujinOpen.png", confidence=0.9) is None:
            time.sleep(1)
            # 新規別表追加選択-------------------------------------------------
            SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 1)
            if SB[0] is True:
                RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\\ZeimuPrint.png", 0.9, 10)
        time.sleep(1)
        # 一覧表出力項目指定が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\PrintBar.png", confidence=0.9) is None:
            time.sleep(1)
        # --------------------------------------------------------------------
        # 申告税一覧表印刷処理----------------------------------------------------
        FO = RPA.ImgCheckForList(
            URL,
            [
                r"\\FileOut.png",
                r"\\FileOut2.png",
            ],
            0.9,
            10,
        )
        if FO[0] is True:
            RPA.ImgClick(URL, FO[1], 0.9, 10)
        RPA.ImgClick(URL, r"\\PDFBar.png", 0.9, 10)
        pg.press("return")
        pg.press("delete")
        pg.press("backspace")
        time.sleep(1)
        pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
        time.sleep(1)
        pg.hotkey("ctrl", "v")
        pg.press("return")
        time.sleep(1)
        RPA.ImgClick(URL, r"\\PrintOut.png", 0.9, 10)
        time.sleep(2)
        # 印刷中が表示されなくなるまで待機---------------------------------
        while (
            pg.locateOnScreen(URL + r"\\ZeimuPrintFlag.png", confidence=0.9) is not None
        ):
            time.sleep(1)
            FO = RPA.ImgCheck(URL, r"\\FileOver.png", 0.9, 1)
            if FO[0] is True:
                pg.press("y")
                time.sleep(2)
        # --------------------------------------------------------------------
        time.sleep(1)
        # 確実に閉じる---------------------------------------------------------
        close()
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\\15Zeimudairi.png",
                    r"\\15Zeimudairi2.png",
                ],
                0.9,
                10,
            )[0]
            is False
        ):
            time.sleep(1)
            close()
        sinkokuend(Job)

        return True, ThisNo, ThisYear, ThisMonth


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdateSyomen(Job, Exc):
    # 税務代理権限証書印刷処理----------------------------------------------------
    OP = RPA.ImgCheckForList(
        URL,
        [
            r"\\16TenpuSyomen.png",
            r"\\16TenpuSyomen2.png",
        ],
        0.9,
        10,
    )
    if OP[0] is True:
        RPA.ImgClick(URL, OP[1], 0.9, 10)
        # 法人税メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\HoujinOpen.png", confidence=0.9) is None:
            time.sleep(1)
            # 新規別表追加選択-------------------------------------------------
            SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 1)
            if SB[0] is True:
                RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
        # --------------------------------------------------------------------
        time.sleep(3)
        AYR = RPA.ImgCheck(URL, r"\\AYear.png", 0.9, 10)
        if AYR[0] is True:
            pg.press("n")
        RPA.ImgClick(URL, r"\\ZeimuPrint.png", 0.9, 10)
        time.sleep(1)
        FO = False, ""
        # 用紙選択が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\YousiSentaku.png", confidence=0.9) is None:
            time.sleep(1)
            FO = RPA.ImgCheckForList(
                URL,
                [
                    r"\\FileOut.png",
                    r"\\FileOut2.png",
                ],
                0.9,
                1,
            )
            if FO[0] is True:
                break
        # --------------------------------------------------------------------
        if FO[0] is False:
            RPA.ImgClick(URL, r"\\YousiOK.png", 0.9, 10)
        # 申告税一覧表印刷処理----------------------------------------------------
        FO = RPA.ImgCheckForList(
            URL,
            [
                r"\\FileOut.png",
                r"\\FileOut2.png",
            ],
            0.9,
            10,
        )
        if FO[0] is True:
            RPA.ImgClick(URL, FO[1], 0.9, 10)
        RPA.ImgClick(URL, r"\\PDFBar.png", 0.9, 10)
        pg.press("return")
        pg.press("delete")
        pg.press("backspace")
        time.sleep(1)
        pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
        time.sleep(1)
        pg.hotkey("ctrl", "v")
        pg.press("return")
        time.sleep(1)
        RPA.ImgClick(URL, r"\\PrintOut.png", 0.9, 10)
        time.sleep(2)
        # 操作ガイドが表示されなくなるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\HPCFlag.png", confidence=0.9) is not None:
            time.sleep(1)
            FO = RPA.ImgCheck(URL, r"\\FileOver.png", 0.9, 1)
            if FO[0] is True:
                pg.press("y")
        time.sleep(1)
        # 確実に閉じる---------------------------------------------------
        close()
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\\16TenpuSyomen.png",
                    r"\\16TenpuSyomen2.png",
                ],
                0.9,
                1,
            )[0]
            is False
        ):
            time.sleep(1)
        sinkokuend(Job)

        return True, ThisNo, ThisYear, ThisMonth


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdateBeppyou(Job, Exc):
    # 別表2-16印刷処理--------------------------------------------------------
    OP = RPA.ImgCheckForList(
        URL,
        [
            r"\\01SinkokuNyuuryoku.png",
            r"\\01SinkokuNyuuryoku2.png",
        ],
        0.9,
        10,
    )
    if OP[0] is True:
        RPA.ImgClick(URL, OP[1], 0.9, 10)
        # 法人税メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\SinkokuTab.png", confidence=0.9) is None:
            Nodatacheck()
            # 新規別表追加選択-------------------------------------------------
            SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 1)
            if SB[0] is True:
                RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\\DownPrint.png", 0.9, 10)
        time.sleep(1)
        pg.press("s")
        # 一覧表メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\SItiranFlag.png", confidence=0.9) is None:
            time.sleep(1)
            HEQ = RPA.ImgCheck(URL, r"\\HNoEntryQ.png", 0.9, 1)
            if HEQ[0] is True:  # 法人番号未登録ダイアログが表示されていたら
                pg.press("y")  # yで確定
            SSL = RPA.ImgCheck(URL, r"\\S_Rendou4.png", 0.9, 1)
            if SSL[0] is True:  # 法人番号未登録ダイアログが表示されていたら
                pg.press("y")  # yで確定
        # --------------------------------------------------------------------
        if RPA.ImgCheck(URL, r"\\Teisyutu.png", 0.99999, 10)[0] is True:
            RPA.ImgClick(URL, r"\\Teisyutu.png", 0.99999, 10)
        if RPA.ImgCheck(URL, r"\\Nyuuryoku.png", 0.99999, 10)[0] is True:
            RPA.ImgClick(URL, r"\\Nyuuryoku.png", 0.99999, 10)
        if RPA.ImgCheck(URL, r"\\Hikae.png", 0.99999, 10)[0] is True:
            RPA.ImgClick(URL, r"\\Hikae.png", 0.99999, 10)
        if RPA.ImgCheck(URL, r"\\Beppyou.png", 0.99999, 10)[0] is True:
            RPA.ImgClick(URL, r"\\Beppyou.png", 0.99999, 10)
            pg.press("space")
        RPA.ImgClick(URL, r"\\SPrint.png", 0.9, 10)
        # 一覧表出力項目指定が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\PrintBar.png", confidence=0.9) is None:
            time.sleep(1)
        # --------------------------------------------------------------------
        # 申告税一覧表印刷処理----------------------------------------------------
        FO = RPA.ImgCheckForList(
            URL,
            [
                r"\\FileOut.png",
                r"\\FileOut2.png",
            ],
            0.9,
            10,
        )
        if FO[0] is True:
            RPA.ImgClick(URL, FO[1], 0.9, 10)
        RPA.ImgClick(URL, r"\\PDFBar.png", 0.9, 10)
        pg.press("return")
        pg.press("delete")
        pg.press("backspace")
        time.sleep(1)
        pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
        time.sleep(1)
        pg.hotkey("ctrl", "v")
        pg.press("return")
        time.sleep(1)
        RPA.ImgClick(URL, r"\\PrintOut.png", 0.9, 10)
        # 印刷中が表示されるまで待機---------------------------------
        IC = 0
        while pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9) is None:
            time.sleep(1)
            IC += 1
            if IC == 5:
                pg.press("tab")
                break
            FO = RPA.ImgCheck(URL, r"\\FileOver.png", 0.9, 1)
            if FO[0] is True:
                pg.press("y")
                while (
                    pg.locateOnScreen(
                        URL + r"\\NowPrint.png",
                        confidence=0.9,
                    )
                    is None
                ):
                    time.sleep(1)
        # --------------------------------------------------------------------
        # 印刷中が表示されなくなるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9) is not None:
            time.sleep(1)
        # --------------------------------------------------------------------
        time.sleep(3)
        # 確実に閉じる---------------------------------------------------
        close()
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\\01SinkokuNyuuryoku.png",
                    r"\\01SinkokuNyuuryoku2.png",
                ],
                0.9,
                1,
            )[0]
            is False
        ):
            time.sleep(1)
        sinkokuend(Job)
        # --------------------------------------------------------------------
        PDFM.BeppyouPDFSplit(
            Exc.Fname.replace("\\\\", "\\").replace("/", "\\"), Job.Img_dir + r"\PDF"
        )
        return True, ThisNo, ThisYear, ThisMonth


# ------------------------------------------------------------------------------------------------------------------
def HoujinzeiUpdateGaikyou(Job, Exc):
    # 事業概況説明書印刷処理--------------------------------------------------------
    OP = RPA.ImgCheckForList(
        URL,
        [
            r"\\02JigyouGaikyou.png",
            r"\\02JigyouGaikyou2.png",
        ],
        0.9,
        10,
    )
    if OP[0] is True:
        RPA.ImgClick(URL, OP[1], 0.9, 10)
        # 事業概況説明書メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\\JigyouTab.png", confidence=0.9) is None:
            Nodatacheck()
            time.sleep(1)
            # 新規別表追加選択-------------------------------------------------
            SB = RPA.ImgCheck(URL, r"\\SaiyouBeppyou.png", 0.9, 10)
            if SB[0] is True:
                RPA.ImgClick(URL, r"\\SBKousin.png", 0.9, 10)
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\\GaikyouPrint.png", 0.9, 10)
        # 一覧表出力項目指定が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\PrintBar.png", confidence=0.9) is None:
            time.sleep(1)
            # 印刷条件設定----------------------------------------------------------
            GP = RPA.ImgCheck(URL, r"\\GPMenu.png", 0.9, 1)
            if GP[0] is True:
                # 白紙チェックボックス確認
                HP = RPA.ImgCheck(URL, r"\\HPCheck.png", 0.9, 1)
                if HP[0] is True:
                    # 白紙チェックボックスが未指定ならクリック
                    RPA.ImgClick(URL, r"\\HPCheck.png", 0.9, 10)
                    # 両面チェックボックス確認
                    RM = RPA.ImgCheck(URL, r"\\Ryoumen.png", 0.9, 1)
                    if RM[0] is True:
                        # 両面チェックボックスが未指定ならクリック
                        RPA.ImgClick(URL, r"\\Ryoumen.png", 0.9, 10)
                        # 印刷ボタンクリック
                        RPA.ImgClick(URL, r"\\GPPrint.png", 0.9, 10)
                    else:
                        # 両面チェックボックスが指定済の場合
                        # 印刷ボタンクリック
                        RPA.ImgClick(URL, r"\\GPPrint.png", 0.9, 10)
                else:
                    # 白紙チェックボックスが指定済の場合
                    # 両面チェックボックス確認
                    RM = RPA.ImgCheck(URL, r"\\Ryoumen.png", 0.9, 1)
                    if RM[0] is True:
                        # 両面チェックボックスが未指定ならクリック
                        RPA.ImgClick(URL, r"\\Ryoumen.png", 0.9, 10)
                        # 印刷ボタンクリック
                        RPA.ImgClick(URL, r"\\GPPrint.png", 0.9, 10)
                    else:
                        # 両面チェックボックスが指定済の場合
                        # 印刷ボタンクリック
                        RPA.ImgClick(URL, r"\\GPPrint.png", 0.9, 10)
        # --------------------------------------------------------------------------
        # 申告税一覧表印刷処理----------------------------------------------------
        FO = RPA.ImgCheckForList(
            URL,
            [
                r"\\FileOut.png",
                r"\\FileOut2.png",
            ],
            0.9,
            10,
        )
        if FO[0] is True:
            RPA.ImgClick(URL, FO[1], 0.9, 10)
        RPA.ImgClick(URL, r"\\PDFBar.png", 0.9, 10)
        pg.press("return")
        pg.press("delete")
        pg.press("backspace")
        time.sleep(1)
        pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
        time.sleep(1)
        pg.hotkey("ctrl", "v")
        pg.press("return")
        time.sleep(1)
        RPA.ImgClick(URL, r"\\PrintOut.png", 0.9, 10)
        # 印刷中が表示されるまで待機---------------------------------
        IC = 0
        while pg.locateOnScreen(URL + r"\\HPCFlag.png", confidence=0.9) is not None:
            time.sleep(1)
            IC += 1
            if IC == 5:
                pg.press("tab")
                break
            FO = RPA.ImgCheck(URL, r"\\FileOver.png", 0.9, 1)
            if FO[0] is True:
                pg.press("y")
        # --------------------------------------------------------------------
        # 印刷中が表示されなくなるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\NowPrint.png", confidence=0.9) is not None:
            time.sleep(1)
        # --------------------------------------------------------------------
        time.sleep(1)
        # 確実に閉じる---------------------------------------------------
        close()
        while (
            RPA.ImgCheckForList(
                URL,
                [
                    r"\\02JigyouGaikyou.png",
                    r"\\02JigyouGaikyou2.png",
                ],
                0.9,
                1,
            )[0]
            is False
        ):
            time.sleep(1)
        sinkokuend(Job)

        return True, ThisNo, ThisYear, ThisMonth


# ------------------------------------------------------------------------------------------------------------------
def Nodatacheck():
    """
    データ無判定関数
    """
    time.sleep(1)
    DNI = RPA.ImgCheck(URL, r"\\NoDataInQ.png", 0.9, 2)
    if DNI[0] is True:
        pg.press("y")
        while pg.locateOnScreen(URL + r"\\NoDataInQ.png", confidence=0.9) is not None:
            time.sleep(1)
    DNQ = RPA.ImgCheck(URL, r"\\NoDataInQ_K.png", 0.9, 2)
    if DNQ[0] is True:
        pg.press("y")
        while pg.locateOnScreen(URL + r"\\NoDataInQ_K.png", confidence=0.9) is not None:
            time.sleep(1)
    NDY = RPA.ImgCheck(URL, r"\NoDataYear.png", 0.9, 2)
    if NDY[0] is True:
        pg.press("return")
        while pg.locateOnScreen(URL + r"\\NoDataInQ_K.png", confidence=0.9) is not None:
            time.sleep(1)


def close():
    while (
        RPA.ImgCheckForList(
            URL,
            [r"\MenuEnd.png", r"\MenuEnd2.png", r"\MenuEnd3.png", r"\MenuEnd4.png"],
            0.9,
            2,
        )[0]
        is True
    ):
        KME = RPA.ImgCheckForList(
            URL,
            [r"\MenuEnd.png", r"\MenuEnd2.png", r"\MenuEnd3.png", r"\MenuEnd4.png"],
            0.9,
            2,
        )
        if KME[0] is True:
            RPA.ImgClick(URL, KME[1], 0.9, 2)
        HM = RPA.ImgCheckForList(
            URL,
            [r"\HoujinOpen.png", r"\HoujinzeiMenu.png"],
            0.9,
            2,
        )
        if HM[0] is True:
            if RPA.ImgCheck(URL, r"\SinkokuPrint.png", 0.9, 10)[0] is True:
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
        TS = RPA.ImgCheck(URL, r"\TaxSet.png", 0.9, 2)
        if TS[0] is True:
            pg.press("return")
        if RPA.ImgCheck(URL, r"\EndCheck.png", 0.9, 2)[0] is True:
            pg.press("return")
        GE = RPA.ImgCheck(
            URL,
            r"\GaikyouEnd.png",
            0.9,
            2,
        )
        if GE[0] is True:
            pg.press("y")
        TS = RPA.ImgCheck(URL, r"\TaxSet.png", 0.9, 2)
        if TS[0] is True:
            pg.press("return")
        if RPA.ImgCheck(URL, r"\TaxEnd.png", 0.9, 2)[0] is True:
            RPA.ImgClick(URL, r"\TaxEnd.png", 0.9, 2)
    return


# ------------------------------------------------------------------------------------------------------------------
def sinkokuend(Job):
    """
    申告書入力終了処理
    """
    time.sleep(1)
    SEQ = RPA.ImgCheck(URL, r"\\SinkokuEndQ2.png", 0.9, 10)
    if SEQ[0] is True:
        RPA.ImgClick(URL, r"\\SinkokuEndQ2Btn.png", 0.9, 10)
    SEQQ = RPA.ImgCheck(URL, r"\\SinkokuEndQ3.png", 0.9, 10)
    if SEQQ[0] is True:
        pg.press("return")
        # 地方税一覧入力が表示されるまで待機---------------------------------
        while pg.locateOnScreen(URL + r"\\SinkokuEndQ4.png", confidence=0.9) is None:
            time.sleep(1)
        RPA.ImgClick(URL, r"\\SinkokuEndQ4.png", 0.9, 10)
        # --------------------------------------------------------------------
    # 閉じる処理--------------------------
    pg.keyDown("alt")
    pg.press("f4")
    pg.keyUp("alt")
    # -----------------------------------
    f = 0
    # 法人税フラグが表示されるまで待機-------------------------------------
    while pg.locateOnScreen(URL + r"\HoujinFlag.png", confidence=0.9) is None:
        time.sleep(1)
        f += 1
        if f == 5:
            pg.keyDown("alt")
            pg.press("f4")
            pg.keyUp("alt")
            f = 0
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
