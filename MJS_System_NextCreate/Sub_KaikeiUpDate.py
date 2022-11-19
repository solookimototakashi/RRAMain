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
    概要: 会計大将更新処理
    @param FolURL : ミロク起動関数のフォルダ(str)
    @param URL : このpyファイルのフォルダ(str)
    @param Exc.row_data : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        URL = Job.imgdir_url + r"\\KaikeiUpDate"
        # 会計大将フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\Kaikei_CFlag.png", confidence=0.9) is None:
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
        if str(Exc.row_data["関与先番号"]) == ThisNo:
            print("関与先あり")
            pg.press(["return", "return", "return"])
            time.sleep(1)
            # 会計大将メニューが表示されるまで待機------------------------------------
            while (
                pg.locateOnScreen(URL + r"\K_TaisyouMenu.png", confidence=0.9) is None
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
            if (
                Job.Start_Year < int(ThisYear) or Job.Start_Year - int(ThisYear) == 1
            ):  # 次年度更新か判定
                IUD = IkkatuUpDate(Job, Exc)
                time.sleep(3)
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
                # --------------------------------------------------------------------
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
            elif Job.Start_Year == int(ThisYear):  # 次年度更新か判定
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                IUD = "当年"
            else:
                IUD = IkkatuUpDate(Job, Exc)
                time.sleep(3)
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
                # --------------------------------------------------------------------
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
            if IUD is True:
                print("更新完了")
                return True, ThisNo, ThisYear, ThisMonth
            elif IUD == "当年":
                return False, "当年データ重複エラー", "", ""
            else:
                return False, "会計大将一括更新エラー", "", ""
        else:
            print("関与先なし")
            return False, "関与先なし", "", ""
    except:
        return False, "exceptエラー", "", ""


# ------------------------------------------------------------------------------------------------------------------
def CampanyUpDate(Job, Exc):
    """
    会計大将:会社情報更新
    """
    URL = Job.imgdir_url + r"\\KaikeiUpDate"
    RPA.ImgClick(URL, r"\D_TourokuTAB.png", 0.9, 10)  # 5.導入登録のアイコンをクリック
    CamList = [r"\CamIcon.png", r"\CamIcon2.png"]
    # 会社基本情報アイコンが表示されるまで待機--------------------------------------
    while RPA.ImgCheckForList(URL, CamList, 0.9, 1)[0] is False:
        time.sleep(1)
    CAL = RPA.ImgCheckForList(URL, CamList, 0.9, 1)
    if CAL[0] is True:
        RPA.ImgClick(URL, CAL[1], 0.9, 1)  # 会社基本情報アイコンをクリック
    # 顧問先情報取り込みアイコンが表示されるまで待機--------------------------------
    while pg.locateOnScreen(URL + r"\DataInIcon.png", confidence=0.9) is None:
        time.sleep(1)
    RPA.ImgClick(URL, r"\DataInIcon.png", 0.9, 10)  # 顧問先情報取り込みアイコンをクリック
    while pg.locateOnScreen(URL + r"\DataInOK.png", confidence=0.9) is None:
        time.sleep(1)
    RPA.ImgClick(URL, r"\DataInOK.png", 0.9, 10)  # 取り込むボタンをクリック
    time.sleep(3)
    pg.keyDown("alt")
    pg.press("x")
    pg.keyUp("alt")
    time.sleep(1)
    return


# ------------------------------------------------------------------------------------------------------------------
def IkkatuUpDate(Job, Exc):
    """
    会計大将:その他メニュー一括更新
    """
    try:
        CampanyUpDate(Job, Exc)  # 会社情報更新
        URL = Job.imgdir_url + r"\\KaikeiUpDate"
        RPA.ImgClick(URL, r"\M_Sonota.png", 0.9, 10)  # その他メニュ-のアイコンをクリック
        # 一括更新のアイコンが表示されるまで待機----------------------------------
        while pg.locateOnScreen(URL + r"\IkkatsuKousin.png", confidence=0.9) is None:
            time.sleep(1)
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\IkkatsuKousin.png", 0.9, 10)  # その他メニュ-のアイコンをクリック
        # 一括更新メニューが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\IkkatuOpenFlag.png", confidence=0.9) is None:
            time.sleep(1)
        # --------------------------------------------------------------------
        RPA.ImgClick(URL, r"\IkkatuOpenFlag.png", 0.9, 10)  # 一括更新メニューのアイコンをクリック
        while (
            RPA.ImgCheckForList(URL, [r"IkkatuFind.png", r"IkkatuFind2.png"], 0.9, 10)[
                0
            ]
            is False
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

        # ThisYearKey.pngが表示されるまで繰り返し
        while RPA.ImgCheck(URL, r"\ThisYearKey.png", 0.9, 1)[0] is False:
            pyperclip.copy(str(Exc.row_data["関与先番号"]))
            pg.hotkey("ctrl", "v")
            # 検索ボタンまでエンター-------------------------------------
            while RPA.ImgCheck(URL, r"FindFlag.png", 0.9, 1)[0] is False:
                time.sleep(1)
                pg.press("return")
            pg.press("return")
            time.sleep(1)
            pg.press("space")
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
            RPA.ImgClick(URL, r"\IkkatuStart.png", 0.9, 10)  # 一括更新開始のアイコンをクリック
            # 確認ウィンドウが表示されるまで待機-------------------------------------
            while pg.locateOnScreen(URL + r"\SakuseiQ.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("y")  # yで決定(nがキャンセル)
            # 確認ウィンドウが表示されるまで待機-------------------------------------
            while pg.locateOnScreen(URL + r"\SakuseiQ2.png", confidence=0.9) is None:
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("return")  # 決定
            # 処理終了ウィンドウが表示されるまで待機----------------------------------
            while (
                pg.locateOnScreen(URL + r"\IkkatuEndFlag.png", confidence=0.9) is None
            ):
                time.sleep(1)
            # --------------------------------------------------------------------
            pg.press("return")  # 決定
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
            while RPA.ImgCheck(URL, r"\IkkatuUpDateStopFlag.png", 0.9, 1)[0] is False:
                pg.press("tab")
            pg.press("tab")
        return True
    except:
        return False


# ------------------------------------------------------------------------------------------------
def KessanKakuteiErr(Job, Exc):
    URL = Job.imgdir_url + r"\\KaikeiUpDate"

    time.sleep(1)
    pg.press("return")
    time.sleep(1)

    p = pg.locateOnScreen(URL + r"\KessanKey.png", confidence=0.9)  # 決算月の画像
    K_x, K_y = pg.center(p)
    OverCount = 0
    while OverCount <= 6:
        try:
            # マスター更新後月次確定コメントがでたら
            MR = RPA.ImgCheck(URL, r"\MasterRecalcQ2.png", 0.9, 10)
            if MR[0] is True:
                while RPA.ImgCheck(URL, r"\GetusjiKakutei.png", 0.9, 1)[0] is True:
                    time.sleep(1)
                pg.press("return")
                time.sleep(3)
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
                RPQ = RPA.ImgCheck(URL, r"\Replace_Q.png", 0.9, 10)
                if RPQ[0] is True:
                    pg.press("n")
                time.sleep(1)
                # マスター更新------------------------------------------------------------------------
                RPA.ImgClick(URL, r"\MasterUp.png", 0.9, 10)  # マスター更新をクリック
                while (
                    RPA.ImgCheckForList(
                        URL,
                        [r"MasterUpStart.png", r"MasterUpStart2.png"],
                        0.9,
                        1,
                    )[0]
                    is True
                ):
                    time.sleep(1)
                    TL = RPA.ImgCheckForList(
                        URL,
                        [r"MasterUpStart.png", r"MasterUpStart2.png"],
                        0.9,
                        10,
                    )
                    RPA.ImgClick(URL, TL[1], 0.9, 10)  # マスター更新開始をクリック
                    break
                time.sleep(3)
                pg.press("y")
                time.sleep(3)
                pg.press("return")
                time.sleep(3)
                pg.keyDown("alt")
                pg.press("x")
                pg.keyUp("alt")
                return "Master"
            else:
                p = pg.locateOnScreen(URL + r"\UnsettledBox.png", confidence=0.8)
                x, y = pg.center(p)
                try:
                    # 確定取消判定
                    ka_p = pg.locateOnScreen(
                        URL + r"\UnsettledBox2.png", confidence=0.8
                    )
                    ka_x, ka_y = pg.center(ka_p)
                    KakuteiFlag = True
                except:
                    KakuteiFlag = False

                if KakuteiFlag is False:
                    if y <= (K_y + (OverCount * 2)):  # チェックボックスが決算月より上なら
                        pg.click(x, y)
                        BR = RPA.ImgCheck(
                            URL, r"\BranceErr.png", 0.9, 10
                        )  # バランスエラーが表示されたら
                        if BR[0] is True:
                            return False
                    else:
                        OverCount += 1
                else:
                    if ka_y <= (K_y + (OverCount * 2)):  # チェックボックスが決算月より上なら
                        pg.click(ka_x, ka_y)
                        BR = RPA.ImgCheck(
                            URL, r"\BranceErr.png", 0.9, 10
                        )  # バランスエラーが表示されたら
                        if BR[0] is True:
                            return False
                    else:
                        OverCount += 1
        except:
            print("失敗")
            OverCount += 1
    GK = False
    RPA.ImgClick(URL, r"\KakuteiLock.png", 0.9, 10)  # 月次処理確定アイコンをクリック
    time.sleep(1)
    pg.press("y")
    while pg.locateOnScreen(URL + r"\KessanKQ.png", confidence=0.9) is None:
        if GK is True:
            break
        time.sleep(2)
        pg.press("y")
        while pg.locateOnScreen(URL + r"\GetusjiKakutei.png", confidence=0.9) is None:
            time.sleep(1)
            pg.press("return")
            while (
                pg.locateOnScreen(URL + r"\GetusjiKakutei.png", confidence=0.9) is None
            ):
                time.sleep(1)
                pg.press("return")
                GK = True
                break
        if GK is True:
            break

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
