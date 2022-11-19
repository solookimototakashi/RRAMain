from ctypes import windll
import pyautogui as pg
import time
import RPA_Function as RPA
import pyperclip  # クリップボードへのコピーで使用
import wrapt_timeout_decorator

TIMEOUT = 600

# ------------------------------------------------------------------------------------------------------------------
# @wrapt_timeout_decorator.timeout(dec_timeout=TIMEOUT)
def Flow(Job, Exc):
    """
    概要: 年調更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\NencyouUpdate"
        ErrStr = ""  # Rpaエラー判別変数
        # 年調フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\NencyouFlag.png", confidence=0.9) is None:
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
        pg.press("return")
        time.sleep(1)
        YearDiff = Job.Start_Year - int(ThisYear)
        if abs(YearDiff) > Job.Start_Year:
            YearDiff = 32 - int(ThisYear) + Job.Start_Year - 2
        # 他システムとメニューが違う-------------------------------------------------------
        if str(Exc.row_data["関与先番号"]) == ThisNo:
            if YearDiff == 1:  # 次年度更新か判定
                print("関与先あり")
                pg.press(["return", "return", "return"])
                time.sleep(1)
                # 年調更新アイコンが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouOpenFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
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

                RPA.ImgClick(URL, r"\NencyouTopMenu.png", 0.9, 10)  # 印刷更新タブをクリック
                time.sleep(1)
                RPA.ImgClick(URL, r"\NenjiKakutei.png", 0.9, 10)  # 年長確定処理ボタンをクリック
                # 年長確定確認メニューが表示されるまで待機--------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouKakuteiQ.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                time.sleep(1)
                NK = RPA.ImgCheck(
                    URL, r"\NencyouKubun.png", 0.99999, 10
                )  # 年調チェックボックスをクリック
                if NK[0] is True:
                    RPA.ImgClick(URL, r"\NencyouKubun.png", 0.99999, 10)  # ＯＫアイコンをクリック
                time.sleep(1)
                RPA.ImgClick(URL, r"\NencyouOK.png", 0.9, 10)  # ＯＫアイコンをクリック
                time.sleep(1)
                pg.press("y")
                # 年調更新アイコンが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouOpenFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DounyuuTAB.png", 0.9, 10)  # 導入タブをクリック
                time.sleep(1)
                RPA.ImgClick(URL, r"\CamIcon.png", 0.9, 10)  # 会社基本情報アイコンをクリック
                # 顧問先情報取り込みアイコンが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None
                ):
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInIcon.png", 0.9, 10)  # 顧問先情報取り込みアイコンをクリック
                # 更新確認ダイアログが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KousinKakunin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                pg.press("y")
                # 更新完了ダイアログが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KousinKanryou.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                pg.press("return")
                # 取り込むボタンが表示されるまで待機--------------------------------
                while pg.locateOnScreen(URL + r"\DataInOK.png", confidence=0.9) is None:
                    time.sleep(1)
                RPA.ImgClick(URL, r"\DataInOK.png", 0.9, 10)  # 取り込むボタンをクリック
                time.sleep(3)
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
                time.sleep(1)
                # 印刷更新タブが表示されるまで待機--------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouTopMenu.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                RPA.ImgClick(URL, r"\NencyouTopMenu.png", 0.9, 10)  # 印刷更新タブをクリック
                # 年調メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                time.sleep(1)

                # 年調更新をクリック---------------------------------------------------
                RPA.ImgClick(URL, r"\NencyouKousin.png", 0.9, 10)
                # 更新区分フラグが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\NencyouKousinFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)

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
                    while RPA.ImgCheck(URL, r"ZNBtn.png", 0.9, 1)[0] is False:
                        time.sleep(1)
                        pg.press("return")
                pg.press("return")
                time.sleep(1)
                pg.press("space")

                # --------------------------------------------------------------------
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
                    # --------------------------------------------------------------------
                    time.sleep(1)
                    RPA.ImgClick(URL, r"\NencyouStart.png", 0.9, 10)  # 更新開始のアイコンをクリック
                    # 確認ウィンドウが表示されるまで待機-------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\NencyouQ.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                    pg.press("y")  # yで決定(nがキャンセル)
                    # 確認ウィンドウが表示されるまで待機-------------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\NencyouQ2.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                    # --------------------------------------------------------------------
                    pg.press("y")  # yで決定(nがキャンセル)
                    # 処理終了ウィンドウが表示されるまで待機----------------------------------
                    while (
                        pg.locateOnScreen(URL + r"\NencyouFQ.png", confidence=0.9)
                        is None
                    ):
                        time.sleep(1)
                        # 年調データが完了していない場合表示される----------------------------
                        ErrC = RPA.ImgCheck(URL, r"NencyouQ3ErrMsg.png", 0.9, 10)
                        if ErrC[0] is True:
                            pg.press("y")  # yで決定(nがキャンセル)
                        # 自然データがすでにある場合表示される----------------------------
                        NNQ = RPA.ImgCheck(URL, r"NextNoQ.png", 0.9, 10)
                        if NNQ[0] is True:
                            pg.press("n")  # yで決定(nがキャンセル)
                            ErrStr = "NoData"
                            break
                    if ErrStr != "NoData":
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
                        # 処理終了ウィンドウが表示されるまで待機----------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\NencyouEnd.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        pg.press("return")

                        # --------------------------------------------------------------------
                        ME = RPA.ImgCheckForList(
                            URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                        )
                        if ME[0] is True:
                            RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                        # 年調開始フラグが表示されるまで待機--------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\NencyouOpenFlag.png", confidence=0.9
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
                        # 年調フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\NencyouFlag.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        print("更新完了")
                    else:
                        # --------------------------------------------------------------------
                        ME = RPA.ImgCheckForList(
                            URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                        )
                        if ME[0] is True:
                            RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                        # 年調開始フラグが表示されるまで待機--------------------------------------
                        while (
                            pg.locateOnScreen(
                                URL + r"\NencyouOpenFlag.png", confidence=0.9
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
                        # 年調フラグが表示されるまで待機-------------------------------------
                        while (
                            pg.locateOnScreen(URL + r"\NencyouFlag.png", confidence=0.9)
                            is None
                        ):
                            time.sleep(1)
                        # --------------------------------------------------------------------
                        print("更新完了")
                    if ErrStr == "":
                        return True, ThisNo, ThisYear, "ThisMonth"
                    else:
                        return False, "該当年度有り", "", ""
            elif str(Job.Start_Year) == ThisYear:
                return False, "該当年度有り", "", ""
            else:
                ID = IkkatuUpDate(Job, Exc, YearDiff)
                if ID is True:
                    return True, ThisNo, ThisYear, "ThisMonth"

        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def IkkatuUpDate(Job, Exc, YearDiff):
    print("")


# ------------------------------------------------------------------------------------------------------------------
def NencyouUpdate(Job, Exc):
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
