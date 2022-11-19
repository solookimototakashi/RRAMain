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
    @param ExRow : Excel抽出行(obj)
    @param driver : 画面操作ドライバー(obj)
    @return : bool,ミロク入力関与先コード, ミロク入力処理年, ミロク入力処理月
    """
    try:
        ErrStr = ""  # Rpaエラー判別変数
        URL = Job.imgdir_url + r"\\KessanUpDate"

        # 決算内訳書フラグが表示されるまで待機------------------------------------
        while pg.locateOnScreen(URL + r"\Kessan_CFlag.png", confidence=0.9) is None:
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
            if str(Job.Start_Year) == ThisYear:
                return False, "該当年度有り", "", ""
            else:
                print("関与先あり")
                pg.press(["return", "return", "return"])
                time.sleep(1)
                # 決算内訳書メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\KessanMenu.png", confidence=0.9) is None
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
                # 一括更新のアイコンが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\IkkatsuKessanKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                RPA.ImgClick(
                    URL, r"\IkkatsuKessanKousin.png", 0.9, 10
                )  # 一括更新のアイコンをクリック
                # 一括更新メニューが表示されるまで待機------------------------------------
                while (
                    pg.locateOnScreen(
                        URL + r"\IkkatuKessanOpenFlag.png", confidence=0.9
                    )
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                RPA.ImgClick(
                    URL, r"\IkkatuKessanOpenFlag.png", 0.9, 10
                )  # 一括更新メニューのアイコンをクリック
                time.sleep(1)
                RPA.ImgClick(
                    URL, r"\IkkatuKessanStart.png", 0.9, 10
                )  # 一括更新開始のアイコンをクリック
                # 確認ウィンドウが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\SakuseiKessanQ.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                # --------------------------------------------------------------------
                pg.press("y")  # yで決定(nがキャンセル)
                # 処理終了ウィンドウが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\IkkatuKessanEndFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    Noren = RPA.ImgCheckForList(
                        URL, [r"No_Rendou.png", r"No_Rendou2.png"], 0.9, 10
                    )
                    if Noren[0] is True:
                        # ErrStr = "Noren"
                        break
                        # RPA.ImgClick(URL, r"\No_Rendou_Cansel.png", 0.9, 10)
                # --------------------------------------------------------------------
                pg.press("return")  # 決定
                # 一括更新のアイコンが表示されるまで待機----------------------------------
                while (
                    pg.locateOnScreen(URL + r"\IkkatsuKessanKousin.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    ME = RPA.ImgCheckForList(
                        URL, [r"\MenuEnd.png", r"\MenuEnd2.png"], 0.9, 10
                    )
                    if ME[0] is True:
                        pg.press("return")  # 決定
                        # RPA.ImgClick(URL, ME[1], 0.9, 10)  # 終了アイコンをクリック
                # --------------------------------------------------------------------
                # 閉じる処理--------------------------
                pg.keyDown("alt")
                pg.press("f4")
                pg.keyUp("alt")
                # -----------------------------------
                # 決算内訳書フラグが表示されるまで待機-------------------------------------
                while (
                    pg.locateOnScreen(URL + r"\Kessan_CFlag.png", confidence=0.9)
                    is None
                ):
                    time.sleep(1)
                    al4c = RPA.ImgCheck(URL, r"\altf4Q.png", 0.9, 10)  # 終了確認が表示されたら
                    if al4c[0] is True:
                        pg.press("y")  # yで決定(nがキャンセル)
                # --------------------------------------------------------------------
                print("更新完了")
                if ErrStr == "":
                    return True, ThisNo, ThisYear, ThisMonth
                elif ErrStr == "Noren":
                    return False, ErrStr, "", ""
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
