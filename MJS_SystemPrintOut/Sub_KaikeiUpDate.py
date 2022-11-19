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
    概要: 会計大将更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param Job.PrintOut_url : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        global URL, ThisNo, ThisYear, ThisMonth
        URL = Job.PrintOut_url + r"\\KaikeiUpDate"
        # 会計大将フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\Kaikei_CFlag.png", confidence=0.9) is None:
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

        if str(Exc.row_kanyo_no) == ThisNo:
            if str(int(Exc.year) - 1) != ThisYear:
                return False, "年度なし", ThisYear, "NoData"
            print("関与先あり")
            pg.press(["return", "return", "return"])
            # 会計大将メニューが表示されるまで待機------------------------------------
            while (
                pg.locateOnScreen(URL + r"\K_TaisyouMenu.png", confidence=0.9) is None
            ):
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
                # マスター再計算ダイアログ確認------------------------------------------
                MR = RPA.ImgCheck(URL, r"\MasterRecalcQ.png", 0.9, 10)
                if MR[0] is True:
                    pg.press("y")
                    while (
                        pg.locateOnScreen(URL + r"\MasterRecalcQ.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)

                DL = RPA.ImgCheck(URL, r"\DLCheck.png", 0.9, 10)
                if DL[0] is True:
                    pg.press("return")
                    while (
                        pg.locateOnScreen(URL + r"\K_TaisyouMenu.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
            # --------------------------------------------------------------------
            # 指示内容で処理分け----------------------------------------------------------
            if Exc.PN == "消費税確定申告書":
                RPA.ImgClick(URL, r"\KessanSinkoku.png", 0.9, 10)  # 決算申告書アイコンをクリック
                # 決算申告書が表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KessanFlag.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # サイドメニューより消費税申告書を確認------------------------------------------
                SSM = RPA.ImgCheckForList(
                    URL,
                    [r"\Syouhizei.png", r"\Syouhizei2.png"],
                    0.9,
                    10,
                )
                if SSM[0] is True:  # 消費税申告書があれば
                    RPA.ImgClick(URL, SSM[1], 0.9, 10)  # 消費税申告書アイコンをクリック
                    time.sleep(2)
                # 申告書・付表入力を確認------------------------------------------------
                SFI = RPA.ImgCheckForList(
                    URL,
                    [r"\SFIcon.png", r"\SFIcon2.png"],
                    0.9,
                    10,
                )
                if SFI[0] is True:  # 申告書・付表入力があれば
                    RPA.ImgClick(URL, SFI[1], 0.9, 10)  # 申告書・付表入力アイコンをクリック
                KKS = False, ""
                # 申告選択ウィンドウが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SinkokuSentaku.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    # 申告回数設定ウィンドウがひょうじされたら
                    SCQ = RPA.ImgCheck(URL, r"\ScountQ.png", 0.9, 1)
                    if SCQ[0] is True:
                        pg.press("n")  # nでいいえ
                    # 基本情報登録誘導ウィンドウがひょうじされたら
                    KKS = RPA.ImgCheck(URL, r"\KihonKoushin.png", 0.9, 1)
                    if KKS[0] is True:
                        pg.press("return")  # nでいいえ
                        break
                if KKS[0] is False:
                    # 申告種で処理分け----------------------------------------------------------------------------------
                    SL = RPA.ImgCheck(
                        URL, r"\SinkokuList.png", 0.9, 10
                    )  # 申告種類選択ボックスが開かれてるか確認
                    if SL[0] is True:
                        pg.press("return")
                    # ----------------------------------------------------------------------------------------------------
                    time.sleep(1)
                    # OKボタンにフォーカスするまでエンター押下-------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\SinkokuAllOK.png", confidence=0.9)
                        is None
                    ):
                        pg.press("return")
                    # ---------------------------------------------------------------------
                    pg.press("return")  # 確定
                    # 確認ウィンドウが表示されるまで待機-------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\SinkokuWin.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        SIQ = RPA.ImgCheck(URL, r"\SiwakeInputQ.png", 0.9, 1)
                        if SIQ[0] is True:
                            pg.press("y")
                        # 再計算確認ウィンドウがあるか確認
                        RCQ = RPA.ImgCheck(URL, r"\ReCalcQ.png", 0.9, 1)
                        if RCQ[0] is True:
                            pg.press("n")
                        # 消費税不明取引確認ウィンドウがあるか確認
                        KN = RPA.ImgCheck(URL, r"\S_Humei.png", 0.9, 1)
                        if KN[0] is True:
                            pg.press("y")
                        DIN = RPA.ImgCheck(URL, r"\Din.png", 0.9, 1)
                        if DIN[0] is True:
                            pg.press("y")
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    SKQ = RPA.ImgCheck(URL, r"\KakuninQ.png", 0.9, 10)
                    if SKQ[0] is True:
                        pg.press("return")
                    time.sleep(1)
                    SP_l = [
                        r"\SyouhiPrint.png",
                        r"\SyouhiPrint2.png",
                    ]
                    SP = RPA.ImgCheckForList(URL, SP_l, 0.9, 10)
                    RPA.ImgClick(URL, SP[1], 0.9, 10)
                    # 標準アイコンが隠されるまで待機
                    while (
                        pg.locateOnScreen(URL + r"\S_Huhyou.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                    c = 0
                    huhyou_flag = False  # 付表フラグ
                    # 一覧表出力項目指定が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                        is None
                    ):
                        # 再計算確認ウィンドウがあるか確認
                        RCQ = RPA.ImgCheck(URL, r"\ReCalcQ.png", 0.9, 1)
                        if RCQ[0] is True:
                            pg.press("n")
                            huhyou_flag = True
                        # 消費税情報変更ウィンドウがあるか確認
                        SC = RPA.ImgCheck(URL, r"\SyouhiChangeQ.png", 0.9, 1)
                        if SC[0] is True:
                            pg.press("n")
                            huhyou_flag = True
                        # 消費税不明取引確認ウィンドウがあるか確認
                        KN = RPA.ImgCheck(URL, r"\S_Humei.png", 0.9, 1)
                        if KN[0] is True:
                            pg.press("y")
                            huhyou_flag = True
                        c += 1
                        if c == 5:
                            SP = RPA.ImgCheckForList(URL, SP_l, 0.9, 1)
                            RPA.ImgClick(URL, SP[1], 0.9, 10)
                            c = 0
                        if huhyou_flag is True:  # 付表フラグ
                            # --------------------------------------------------------------------
                            time.sleep(1)
                            pg.press("p")  # 決定
                            time.sleep(1)
                        # 法人番号未登録確認が表示されているか確認
                        SAL = RPA.ImgCheck(URL, r"\S_Alert.png", 0.9, 10)
                        if SAL[0] is True:  # 法人番号未登録確認が表示された場合
                            pg.press("y")  # yで確定
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
                    # 印刷中が表示されるまで待機---------------------------------
                    IC = 0
                    while (
                        pg.locateOnScreen(
                            URL + r"\NowSyouhiPrint.png",
                            confidence=0.9,
                        )
                        is None
                    ):
                        time.sleep(1)
                        IC += 1
                        if IC == 5:
                            pg.press("tab")
                            break
                        FO = RPA.ImgCheck(URL, r"\FileOver.png", 0.9, 1)
                        if FO[0] is True:
                            pg.press("y")
                            while (
                                pg.locateOnScreen(
                                    URL + r"\NowSyouhiPrint.png",
                                    confidence=0.9,
                                )
                                is None
                            ):
                                time.sleep(1)
                    # --------------------------------------------------------------------
                    # 印刷中が表示されなくなるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\NowSyouhiPrint.png",
                            confidence=0.9,
                        )
                        is not None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    # 確実に閉じる---------------------------------------------------------------------------------
                    lastclose()
                    return True, ThisNo, ThisYear, ThisMonth
                else:
                    # 閉じる処理--------------------------
                    pg.keyDown("alt")
                    pg.press("f4")
                    pg.keyUp("alt")
                    # -----------------------------------
                    f = 0
                    # 会計大将フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\Kaikei_CFlag.png", confidence=0.9)
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
                    return False, "要消費税基本情報登録", "", ""
            elif Exc.PN == "書面添付　消費税":
                RPA.ImgClick(URL, r"\KessanSinkoku.png", 0.9, 10)  # 決算申告書アイコンをクリック
                # 決算申告書が表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KessanFlag.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # サイドメニューより消費税申告書を確認------------------------------------------
                SSM = RPA.ImgCheckForList(
                    URL,
                    [r"\Syouhizei.png", r"\Syouhizei2.png"],
                    0.9,
                    10,
                )
                if SSM[0] is True:  # 消費税申告書があれば
                    RPA.ImgClick(URL, SSM[1], 0.9, 10)  # 消費税申告書アイコンをクリック
                    time.sleep(2)
                # 添付書面を確認------------------------------------------------
                SFI = RPA.ImgCheckForList(
                    URL,
                    [r"\14Tenpu.png", r"\14Tenpu2.png"],
                    0.9,
                    10,
                )
                if SFI[0] is True:  # 添付書面があれば
                    RPA.ImgClick(URL, SFI[1], 0.9, 10)  # 添付書面アイコンをクリック
                # -----------------------------------------------------------------------
                time.sleep(1)
                NOD = False, ""
                # 添付書面申告種類選択ウィンドウが表示されるまで待機--------------------------
                while (
                    pg.locateOnScreen(URL + r"\TenpuBar.png", confidence=0.99999)
                    is None
                ):
                    time.sleep(1)
                    NOD = RPA.ImgCheck(URL, r"\NoSinkoku.png", 0.9, 1)
                    if NOD[0] is True:
                        pg.press("return")
                        break
                if NOD[0] is False:
                    pg.press(["return", "return"])
                    # 添付書面印刷ウィンドウが表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\TenpuFlag.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    time.sleep(3)
                    BEF = RPA.ImgCheck(URL, r"\BeforeQ.png", 0.9, 10)
                    if BEF[0] is True:
                        pg.press("n")
                    # ----------------------------------------------------------------------
                    WPC = RPA.ImgCheckForList(
                        URL,
                        [
                            r"\TenpuFlagPrint.png",
                            r"\TenpuFlagPrint2.png",
                        ],
                        0.9,
                        10,
                    )
                    if WPC[0] is True:
                        RPA.ImgClick(URL, WPC[1], 0.9, 10)
                    else:
                        pg.keyDown("alt")
                        pg.press("return")
                        pg.keyUp("alt")
                    # ----------------------------------------------------------------------
                    FO = False, ""
                    # 添付書面印刷サイズ選択が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(
                            URL + r"\TenpuPrintType.png",
                            confidence=0.9,
                        )
                        is None
                    ):
                        time.sleep(1)
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
                            break
                    if FO[0] is False:
                        A4T = RPA.ImgCheck(URL, r"\A4Box.png", 0.99999, 10)
                        if A4T[0] is True:
                            RPA.ImgClick(URL, r"\A4Box.png", 0.99999, 10)
                        RPA.ImgClick(URL, r"\A4BoxOK.png", 0.9, 10)
                        # 一覧表出力項目指定が表示されるまで待機---------------------------------
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
                    # 印刷中が表示されるまで待機---------------------------------
                    IC = 0
                    while (
                        pg.locateOnScreen(URL + r"\TenpuPBar.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                        IC += 1
                        if IC == 5:
                            pg.press("tab")
                            break
                        FO = RPA.ImgCheck(URL, r"\FileOver.png", 0.9, 1)
                        if FO[0] is True:
                            pg.press("y")
                    # --------------------------------------------------------------------
                    time.sleep(2)
                    # 確実に閉じる---------------------------------------------------------------------------------
                    lastclose()
                    return True, ThisNo, ThisYear, ThisMonth
                elif NOD[0] is True:
                    # 閉じる処理--------------------------
                    pg.keyDown("alt")
                    pg.press("f4")
                    pg.keyUp("alt")
                    # -----------------------------------
                    f = 0
                    # 終了確認が表示されるまで待機---------------------------------
                    while pg.locateOnScreen(URL + r"\EndQ.png", confidence=0.9) is None:
                        time.sleep(1)
                        f += 1
                        if f == 5:
                            pg.keyDown("alt")
                            pg.press("f4")
                            pg.keyUp("alt")
                            f = 0
                    # --------------------------------------------------------------------
                    pg.press("y")
                    # ------------------------------------------------------------------
                    # 会計大将フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\Kaikei_CFlag.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    # ------------------------------------------------------------------
                    return False, "出力履歴なし", "", ""
            elif Exc.PN == "決算報告書":
                CalcErr = ""
                kessansinkoku()  # 決算申告書が表示されるまで待機
                KJJ = RPA.ImgCheck(URL, r"\KessanJyunjyoQ.png", 0.9, 10)
                if KJJ[0] is True:
                    pg.press("return")
                    time.sleep(1)

                SSY = RPA.ImgCheck(URL, r"\Sansyou_F.png", 0.9, 10)
                if SSY[0] is True:
                    pg.press("return")  # 確定後参照表示確認

                if KJJ[0] is False:
                    RPA.ImgClick(URL, r"\KessansyoPrint.png", 0.9, 10)  # 印刷ボタンクリック
                    # 一覧表出力項目指定が表示されるまで待機---------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\PrintBar.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        # 金額エラーが表示されていないか確認
                        KCE = RPA.ImgCheck(URL, r"\K_CalcErr.png", 0.9, 1)
                        if KCE[0] is True:
                            CalcErr = "Err"  # 金額エラーフラグを立てる
                            pg.press("y")  # 金額エラー無視で印刷
                        # 貸借エラーが表示されていないか確認
                        CE = RPA.ImgCheck(URL, r"\CalcErr.png", 0.9, 1)
                        if CE[0] is True:
                            CalcErr = "T_Err"  # 金額エラーフラグを立てる
                            pg.press("y")  # 金額エラー無視で印刷
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
                    # 印刷ボタンが表示されるまで待機---------------------------------
                    #####################################################################
                    # NotEnd.pngはミロクウィンドウ右上のグレーアウトした閉じるボタン。
                    # うまく動作しないときはVS画面などが閉じるボタンに被さってないか確認。
                    #####################################################################
                    while (
                        pg.locateOnScreen(URL + r"\NotEnd.png", confidence=0.9)
                        is not None
                    ):
                        time.sleep(1)
                        KOW = RPA.ImgCheck(URL, r"\K_OverWrite2.png", 0.9, 1)
                        if KOW[0] is True:
                            pg.press("y")
                    # --------------------------------------------------------------------
                    time.sleep(4)
                    # 確実に閉じる---------------------------------------------------------------------------------
                    lastclose()
                    # ------------------------------------------------------------------
                    PDFM.BeppyouPDFSplit(
                        Exc.Fname.replace("\\\\", "\\").replace("/", "\\"),
                        Job.Img_dir + r"\PDF",
                    )
                    if CalcErr == "":
                        return True, ThisNo, ThisYear, ThisMonth
                    elif CalcErr == "T_Err":
                        return True, CalcErr, "", ""
                    else:
                        return False, ThisNo, ThisYear, CalcErr
                else:
                    # 閉じる処理--------------------------
                    pg.keyDown("alt")
                    pg.press("f4")
                    pg.keyUp("alt")
                    # -----------------------------------
                    f = 0
                    # 会計大将フラグが表示されるまで待機------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\Kaikei_CFlag.png", confidence=0.9)
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
                    return False, "決算順序未設定", "", ""
            elif Exc.PN == "株主（社員）資本変動計算書":
                CalcErr = ""
                kessansinkoku()  # 決算申告書が表示されるまで待機
                RPA.ImgClick(URL, r"\K_Preview.png", 0.9, 10)  # 印刷ボタンクリック
                # プレビュー画面が表示されるまで待機---------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_AllRight.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                    # 金額エラーが表示されていないか確認
                    KCE = RPA.ImgCheck(URL, r"\K_CalcErr.png", 0.9, 1)
                    if KCE[0] is True:
                        CalcErr = "Err"  # 金額エラーフラグを立てる
                        pg.press("y")  # 金額エラー無視で印刷
                # ---------------------------------------------------------------------
                RPA.ImgClick(URL, r"\K_AllRight.png", 0.9, 10)  # ページ最後尾へ
                # 左矢印が表示されるまで待機---------------------------------------------
                while pg.locateOnScreen(URL + r"\K_Left.png", confidence=0.9) is None:
                    time.sleep(1)
                # ---------------------------------------------------------------------
                RPA.ImgClick(URL, r"\K_Left.png", 0.9, 10)  # 1ページ戻る
                time.sleep(1)
                RPA.ImgClick(URL, r"\K_ThisPrint.png", 0.9, 10)  # 現在項印刷をクリック
                time.sleep(1)
                # 現在項印刷画面が表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_PrintBar.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # ---------------------------------------------------------------------
                time.sleep(1)
                pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
                time.sleep(1)
                pg.hotkey("ctrl", "v")
                pg.press("return")
                time.sleep(1)
                # 現在項印刷画面が表示されなくなるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_PrintBar2.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                    OVC = RPA.ImgCheck(URL, r"\K_OverWrite.png", 0.9, 1)
                    if OVC[0] is True:
                        pg.press("y")
                # ---------------------------------------------------------------------
                time.sleep(1)
                # 確実に閉じる---------------------------------------------------------------------------------
                lastclose()
                # ------------------------------------------------------------------
                if CalcErr == "":
                    return True, ThisNo, ThisYear, ThisMonth
                else:
                    return False, ThisNo, ThisYear, CalcErr
            elif Exc.PN == "個別注記表":
                CalcErr = ""
                kessansinkoku()  # 決算申告書が表示されるまで待機
                RPA.ImgClick(URL, r"\K_Preview.png", 0.9, 10)  # 印刷ボタンクリック
                # プレビュー画面が表示されるまで待機---------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_AllRight.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                    # 金額エラーが表示されていないか確認
                    KCE = RPA.ImgCheck(URL, r"\K_CalcErr.png", 0.9, 10)
                    if KCE[0] is True:
                        CalcErr = "Err"  # 金額エラーフラグを立てる
                        pg.press("y")  # 金額エラー無視で印刷
                # ---------------------------------------------------------------------
                RPA.ImgClick(URL, r"\K_AllRight.png", 0.9, 10)  # ページ最後尾へ
                # 左矢印が表示されるまで待機---------------------------------------------
                while pg.locateOnScreen(URL + r"\K_Left.png", confidence=0.9) is None:
                    time.sleep(1)
                # ---------------------------------------------------------------------
                time.sleep(1)
                RPA.ImgClick(URL, r"\K_ThisPrint.png", 0.9, 10)  # 現在項印刷をクリック
                time.sleep(1)
                # 現在項印刷画面が表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_PrintBar.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                # ---------------------------------------------------------------------
                time.sleep(1)
                pyperclip.copy(Exc.Fname.replace("\\\\", "\\").replace("/", "\\"))
                time.sleep(1)
                pg.hotkey("ctrl", "v")
                pg.press("return")
                time.sleep(1)
                # 現在項印刷画面が表示されなくなるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\K_PrintBar2.png", confidence=0.9)
                    is not None
                ):
                    time.sleep(1)
                    OVC = RPA.ImgCheck(URL, r"\K_OverWrite.png", 0.9, 1)
                    if OVC[0] is True:
                        pg.press("y")
                # ---------------------------------------------------------------------
                time.sleep(1)
                # 確実に閉じる---------------------------------------------------------------------------------
                lastclose()
                # ------------------------------------------------------------------
                if CalcErr == "":
                    return True, ThisNo, ThisYear, ThisMonth
                else:
                    return False, ThisNo, ThisYear, CalcErr
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ---------------------------------------------------------------------------------
def kessansinkoku():
    RPA.ImgClick(URL, r"\KessanSinkoku.png", 0.9, 10)  # 決算申告書アイコンをクリック
    # 決算申告書が表示されるまで待機----------------------------------
    while pg.locateOnScreen(URL + r"\KessanFlag.png", confidence=0.9) is None:
        time.sleep(1)
    # 01決算書アイコンを確認------------------------------------------
    KSP = RPA.ImgCheckForList(
        URL,
        [r"\01Kessansyo.png", r"\01Kessansyo2.png"],
        0.9,
        10,
    )
    if KSP[0] is True:  # 01決算書アイコンがあれば
        ll = 0
        RPA.ImgClick(URL, KSP[1], 0.9, 10)  # 01決算書アイコンをクリック
        while pg.locateOnScreen(URL + r"\K_PreviewFlag.png", confidence=0.9) is None:
            time.sleep(1)
            if RPA.ImgCheck(URL, r"\Kessankakutei.png", 0.9, 1)[0] is True:
                pg.press("return")
            ll += 1
            if ll == 5:
                RPA.ImgClick(URL, KSP[1], 0.9, 10)  # 01決算書アイコンをクリック


# ---------------------------------------------------------------------------------
def lastclose():

    time.sleep(1)
    f = 0
    while RPA.ImgCheck(URL, r"\HomeIcon.png", 0.9, 1)[0] is False:
        time.sleep(1)
        if RPA.ImgCheck(URL, r"\EndCheck.png", 0.9, 1)[0] is True:
            pg.press("y")
        else:
            # 閉じる処理--------------------------
            pg.keyDown("alt")
            pg.press("f4")
            pg.keyUp("alt")
            # -----------------------------------
            if RPA.ImgCheck(URL, r"\Last_End.png", 0.9, 1)[0] is True:
                pg.press("n")

            # 会計大将フラグが表示されるまで待機------------------------------------
            if RPA.ImgCheck(URL, r"\EndWindow.png", 0.9, 1)[0] is True:
                pg.press("y")

            if RPA.ImgCheck(URL, r"\CallErr.png", 0.9, 1)[0] is True:
                pg.press("return")

            if f == 10:
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                f = 0
                if RPA.ImgCheck(URL, r"\Last_End.png", 0.9, 1)[0] is True:
                    pg.press("n")
            f += 1
            # -----------------------------------
    if RPA.ImgCheck(URL, r"\Last_End.png", 0.9, 10)[0] is True:
        pg.press("n")
    return True


# ------------------------------------------------------------------------------------------------------------------
def KaikeiUpDate(Job, Exc):
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
