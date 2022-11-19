import os
import numpy as np
import pandas as pd
import re


def DiffListPlus(CDict, Settingtoml, ColList, ScrList, Ers):
    try:
        LNList = Settingtoml["MASTER"]["ListNameList"]  # tomlから各種設定リスト名を抽出
        NewColList = []
        for LNListItem in LNList:
            NewColList = Settingtoml["CsvSaveEnc"][LNListItem]
            SColA = set(ColList) - set(NewColList)
            SColB = set(NewColList) - set(ColList)
            if len(SColA) == 0 and len(SColB) == 0:
                CDict[LNListItem].append(ScrList)
                return True, LNListItem
        print(ColList)
        print(
            "========================================================================================"
        )
        print(ScrList)
        if Ers == "Sub":
            print("サブテーブル取得エラー")
            CDict["SubErrList"].append(ScrList)
            return False, "SubErrList"
        else:
            print("指定列名での設定項目がありませんでした。")
            CDict["ErrList"].append(ScrList)
            return False, "ErrList"
    except:
        print(ColList)
        print(
            "========================================================================================"
        )
        print(ScrList)
        if Ers == "Sub":
            print("サブテーブル取得エラー")
            CDict["SubErrList"].append(ScrList)
            return False, "SubErrList"
        else:
            print("指定列名での設定項目がありませんでした。")
            CDict["ErrList"].append(ScrList)
            return False, "ErrList"


# -------------------------------------------------------------------------------------------------------
def ChangeBusyu(strs):
    try:
        MeUrl = os.getcwd().replace("\\", "/")  # 自分のパス
        df = pd.read_csv(MeUrl + "/Function/cmap置換.csv", encoding="utf8")
        npdf = np.array(df)
        lstr = len(strs)
        npdfRow = npdf.shape[0]
        Cst = ""
        for s in range(lstr):
            CF = False
            for d in range(npdfRow):
                Key = npdf[d, 3]
                Tar = npdf[d, 4]
                if strs[s] == Key:
                    strs.replace(strs[s], Tar)
                    CF = True
                    break
                else:
                    CF = False
            if CF is True:
                Cst += Tar
            else:
                Cst += strs[s]
        return Cst
    except:
        return ""


# --------------------------------------------------------------------------------------------------
def CellsAction(stt, cellsList, ColumList, TxtList):  # 主にeltax処理
    try:
        ColumFlag = False  # カラム対象フラグ
        TFlag = False  # テキスト取得対象フラグ
        txts = []
        Kint = 1
        PreFlag = False
        for cellsListItem in cellsList:  # セルループ
            ThreeCol = False
            txt = repr(cellsListItem[1].text)  # 改行コードの数でセル分割の判定
            if "＜プレ申告データの内容＞" in txt:
                PreFlag = True
            nc = txt.count(r"\n")  # テキスト内の改行コードの数
            if txt.endswith(r"\n") is True:  # テキスト内の最後の改行コードは省く
                nc -= 1
            if nc >= 2:  # テキスト内に改行コードが2つ以上ある場合
                if len(cellsListItem) >= 2:
                    txt = repr(cellsListItem[0].text)  # 改行コードの数でセル分割の判定
                    nc = txt.count(r"\n")  # テキスト内の改行コードの数
                    if nc >= 3:
                        ThreeCol = True
                txts = txt.split(r"\n")  # テキストを改行コードで分割する
                txtsc = len(txts)  # 改行コード分割後のリスト要素数
                for tx in range(txtsc):  # 改行コード分割後のリストを置換
                    txts[tx] = (
                        txts[tx]
                        .replace("'", "")
                        .replace('"', "")
                        .replace("\xa0", "")
                        .replace(r"\n", "")
                        .replace(r"\u2003", " ")
                        .replace(r"\u3000", " ")
                    )
                    txts[tx] = txts[tx].replace(" ", "")
                    txts[tx] = ChangeBusyu(txts[tx])
                and_list = set(txts) & set(stt)  # テキストリストとカラムリストで一致する要素を抽出
                if len(and_list) > 0:  # テキストリストとカラムリストで一致する要素が一つ以上なら
                    TFlag = True  # テキスト取得対象フラグを立てる
                if PreFlag is True:
                    for txtsItem in txts:
                        for sttItem in stt:
                            if sttItem in txtsItem:
                                print(txtsItem)
                                tisp = txtsItem.replace(
                                    sttItem, sttItem + "::"
                                )  # テキストを改行コードで分割する
                                tisp = tisp.split("::")
                                for sttItem in stt:
                                    if sttItem == tisp[0]:
                                        ColumList.append(tisp[0])
                                        TxtList.append(tisp[1])  # テキストリストに代入
                else:
                    # テキスト取得対象の処理--------------------------------------------------------
                    if TFlag is True:
                        for tc in range(txtsc):  # 分割後のリストループ
                            # ヘッダー判定後の処理--------------------------------------------------
                            if tc % 2 == 0:  # 2で割り切れたらヘッダー
                                # 処理中のテキストがカラムリストにあるか判定--------------------------
                                for sttItem in stt:
                                    st = txts[tc]
                                    if st == sttItem:
                                        ColumFlag = True  # カラム対象フラグを立てる
                                        break
                                # ----------------------------------------------------------------
                                # テキスト内の改行コード総数が3で割り切れる場合------------------------
                                if nc % 3 == 0 and tc == 0:  # テキスト内初回の処理
                                    if Kint == 1:
                                        ColumList.append("項目")  # 項目リストに代入
                                        Kint += 1
                                    else:
                                        ColumList.append("項目" + str(Kint))  # 項目リストに代入
                                        Kint += 1
                                if ColumFlag is True:  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                    ColumList.append(st)  # 項目リストに代入
                                    ColumFlag = False  # カラム対象フラグ解除
                                else:
                                    if TFlag is True:
                                        if not txts[tc] == "" and not txts[tc] == " ":
                                            TxtList.append(st)  # テキストリストに代入
                                # ----------------------------------------------------------------
                            else:
                                if (
                                    not txts[tc] == ""
                                    and not txts[tc] == " "
                                    and ThreeCol is False
                                ):
                                    tsc = txts[tc]
                                    if tsc.startswith(" ") is True:
                                        tsc = tsc.replace(" ", "")
                                    TxtList.append(tsc)  # テキストリストに代入
                            # ---------------------------------------------------------------------
                        TFlag = False
                    # -----------------------------------------------------------------------------
                    elif ThreeCol is True:
                        txt = repr(cellsListItem[0].text)  # 改行コードの数でセル分割の判定
                        txts = txt.split(r"\n")  # テキストを改行コードで分割する
                        txt2 = repr(cellsListItem[1].text)  # 改行コードの数でセル分割の判定
                        txts2 = txt2.split(r"\n")  # テキストを改行コードで分割する
                        txts.pop(0)
                        txtsc = len(txts)  # 改行コード分割後のリスト要素数
                        for tx in range(txtsc):  # 改行コード分割後のリストを置換
                            txts[tx] = (
                                txts[tx]
                                .replace("'", "")
                                .replace('"', "")
                                .replace("\xa0", "")
                                .replace(r"\n", "")
                                .replace(r"\u2003", " ")
                                .replace(r"\u3000", " ")
                            )
                            txts[tx] = txts[tx].replace(" ", "")
                            txts[tx] = ChangeBusyu(txts[tx])
                            if not txts[tx] == "":
                                ColumList.append(txts[tx])  # 項目リストに代入
                                txts2[tx] = (
                                    txts2[tx]
                                    .replace("'", "")
                                    .replace('"', "")
                                    .replace("\xa0", "")
                                    .replace(r"\n", "")
                                    .replace(r"\u2003", " ")
                                    .replace(r"\u3000", " ")
                                )
                                txts2[tx] = txts2[tx].replace(" ", "")
                                txts2[tx] = ChangeBusyu(txts2[tx])
                                TxtList.append(txts2[tx])  # 項目リストに代入
            else:
                if (
                    "税目１" in stt or "税目２" in stt or "優先税⽬" in stt or "第２税⽬" in stt
                ) and not cellsListItem[1].text == "":
                    txts = repr(
                        cellsListItem[1].text.replace(r"\n", "")
                    )  # 改行コード判定の為repr
                else:
                    txts = repr(
                        cellsListItem[0].text.replace(r"\n", "")
                    )  # 改行コード判定の為repr
                txts = (
                    txts.replace("'", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                txts = txts.replace(" ", "")
                txts = ChangeBusyu(txts)
                if txts.startswith(" ") is True:
                    txts.replace(" ", "")
                ColumList.append(txts)  # 項目リストに代入
                if "税目１" in stt or "税目２" in stt or "優先税⽬" in stt or "第２税⽬" in stt:
                    txts = repr(
                        cellsListItem[2].text.replace(r"\n", "")
                    )  # 改行コード判定の為repr
                else:
                    txts = repr(
                        cellsListItem[1].text.replace(r"\n", "")
                    )  # 改行コード判定の為repr
                txts = (
                    txts.replace("'", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                txts = txts.replace(" ", "")
                txts = ChangeBusyu(txts)
                if txts.startswith(" ") is True:
                    txts.replace(" ", "")
                TxtList.append(txts)  # テキストリストに代入
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsActionMJS(stt, rtt, cellsList, ColumList, TxtList, Sbtext):  # 主にeltax処理
    try:
        Sbtext = Sbtext.split("\n")
        print(Sbtext)
        for SbtextItem in Sbtext:  # セルループ
            SbtextItem = SbtextItem.replace("\u3000", "").replace(r"\u3000", "")
            # 処理中のテキストがカラムリストにあるか判定--------------------------
            for rttItem in rtt:
                if (
                    rttItem in SbtextItem
                    and not SbtextItem == ""
                    and "ＱＲコード" not in SbtextItem
                ):
                    if rttItem + ":" in SbtextItem:
                        SbtextItem = SbtextItem.replace(rttItem + ":", rttItem + "::")
                    elif rttItem + "：" in SbtextItem:
                        SbtextItem = SbtextItem.replace(rttItem + "：", rttItem + "::")
                    else:
                        SbtextItem = SbtextItem.replace(rttItem, rttItem + ":")
                        SbtextItem = SbtextItem.replace(rttItem + ":", rttItem + "::")
                        if "受付日::時：" in SbtextItem:
                            SbtextItem = SbtextItem.replace("受付日::時：", "受付日時::")
                        elif "受付日::時:" in SbtextItem:
                            SbtextItem = SbtextItem.replace("受付日::時:", "受付日時::")
                    SbtextItem = SbtextItem.replace(rttItem + "：", rttItem + "::")
                    SbtextItem = SbtextItem.replace("\u3000", "").replace(r"\u3000", "")
                    SbtextItem = SbtextItem.replace(" ", "").replace("円:", "円::")
                    SI = SbtextItem.split("::")
                    ColumList.append(SI[0])
                    TxtList.append(SI[1])
                    break

        ColumFlag = False  # カラム対象フラグ
        TFlag = False  # テキスト取得対象フラグ
        txts = []
        Kint = 1
        for cellsListItem in cellsList:  # セルループ
            txt = repr(cellsListItem[1].text)  # 改行コードの数でセル分割の判定
            nc = txt.count(r"\n")  # テキスト内の改行コードの数
            if txt.endswith(r"\n") is True:  # テキスト内の最後の改行コードは省く
                nc -= 1
            if nc >= 2:  # テキスト内に改行コードが2つ以上ある場合
                txts = txt.split(r"\n")  # テキストを改行コードで分割する
                txtsc = len(txts)  # 改行コード分割後のリスト要素数
                for tx in range(txtsc):  # 改行コード分割後のリストを置換
                    txts[tx] = (
                        txts[tx]
                        .replace("'", "")
                        .replace('"', "")
                        .replace("\xa0", "")
                        .replace(r"\n", "")
                        .replace(r"\u2003", " ")
                        .replace(r"\u3000", " ")
                    )
                    txts[tx] = txts[tx].replace(" ", "")
                    txts[tx] = ChangeBusyu(txts[tx])
                and_list = set(txts) & set(stt)  # テキストリストとカラムリストで一致する要素を抽出
                if len(and_list) > 0:  # テキストリストとカラムリストで一致する要素が一つ以上なら
                    TFlag = True  # テキスト取得対象フラグを立てる
                # テキスト取得対象の処理--------------------------------------------------------
                if TFlag is True:
                    for tc in range(txtsc):  # 分割後のリストループ
                        # ヘッダー判定後の処理--------------------------------------------------
                        if tc % 2 == 0:  # 2で割り切れたらヘッダー
                            # 処理中のテキストがカラムリストにあるか判定--------------------------
                            for sttItem in stt:
                                st = txts[tc]
                                if st == sttItem:
                                    ColumFlag = True  # カラム対象フラグを立てる
                                    break
                            # ----------------------------------------------------------------
                            # テキスト内の改行コード総数が3で割り切れる場合------------------------
                            if nc % 3 == 0 and tc == 0:  # テキスト内初回の処理
                                if Kint == 1:
                                    ColumList.append("項目")  # 項目リストに代入
                                    Kint += 1
                                else:
                                    ColumList.append("項目" + str(Kint))  # 項目リストに代入
                                    Kint += 1
                            if ColumFlag is True:  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                ColumList.append(st)  # 項目リストに代入
                                ColumFlag = False  # カラム対象フラグ解除
                            else:
                                if TFlag is True:
                                    if not txts[tc] == "" and not txts[tc] == " ":
                                        TxtList.append(st)  # テキストリストに代入
                            # ----------------------------------------------------------------
                        else:
                            if not txts[tc] == "" and not txts[tc] == " ":
                                tsc = txts[tc]
                                if tsc.startswith(" ") is True:
                                    tsc = tsc.replace(" ", "")
                                TxtList.append(tsc)  # テキストリストに代入
                        # ---------------------------------------------------------------------
                    TFlag = False
                # -----------------------------------------------------------------------------
            else:
                txts = repr(cellsListItem[0].text.replace(r"\n", ""))  # 改行コード判定の為repr
                txts = (
                    txts.replace("'", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                txts = txts.replace(" ", "")
                txts = ChangeBusyu(txts)
                if txts.startswith(" ") is True:
                    txts.replace(" ", "")
                ColumList.append(txts)  # 項目リストに代入
                txts = repr(cellsListItem[1].text.replace(r"\n", ""))  # 改行コード判定の為repr
                txts = (
                    txts.replace("'", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                txts = txts.replace(" ", "")
                txts = ChangeBusyu(txts)
                if txts.startswith(" ") is True:
                    txts.replace(" ", "")
                TxtList.append(txts)  # テキストリストに代入
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsActionJigyounendo(stt, cellsList, ColumList, TxtList):  # 主にeltax処理
    try:
        ColumFlag = False  # カラム対象フラグ
        TFlag = False  # テキスト取得対象フラグ
        txts = []
        cicount = 0
        for cellsListItem in cellsList:  # セルループ
            Skiprows = False
            for cLIItem in cellsListItem:  # セルループ
                if "月数換算" in cLIItem.text:
                    Skiprows = True
            if Skiprows is False:
                for cLIItem in cellsListItem:  # セルループ
                    if len(ColumList) == 4:
                        ColumList.append("前事業年度等")
                    txt = repr(cLIItem.text)  # 改行コードの数でセル分割の判定
                    # テキスト検出ミスの修正-------------------------------------------
                    if (
                        "日 修正・更正・決定の年月日\\n" in txt
                        or "円 前課税" in txt
                        or "円 同上" in txt
                        or "円 差引" in txt
                        or "円 納付" in txt
                    ):
                        txtsp = txt.split(" ")
                        for txtspItem in txtsp:
                            txtspItem = (
                                txtspItem.replace("'", "")
                                .replace("'", "")
                                .replace('"', "")
                                .replace("\xa0", "")
                                .replace(r"\n", "")
                                .replace(r"\u2003", " ")
                                .replace(r"\u3000", " ")
                            )
                            txtspItem = txtspItem.replace(" ", "")
                            txtspItem = ChangeBusyu(txtspItem)
                            # 処理中のテキストがカラムリストにあるか判定--------------------------
                            for sttItem in stt:
                                st = txtspItem
                                if st == sttItem:
                                    ColumFlag = True  # カラム対象フラグを立てる
                                    break
                            # ----------------------------------------------------------------
                            if not st == "" and not st == " ":
                                if st.startswith(" ") is True:
                                    st = st.replace(" ", "")
                                if ColumFlag is True:
                                    ColumList.append(st)
                                    ColumFlag = False
                                else:
                                    TxtList.append(st)  # テキストリストに代入
                    else:
                        # ----------------------------------------------------------------
                        nc = txt.count(r"\n")  # テキスト内の改行コードの数
                        if txt.endswith(r"\n") is True:  # テキスト内の最後の改行コードは省く
                            nc -= 1
                        if nc >= 2:  # テキスト内に改行コードが2つ以上ある場合
                            txts = txt.split(r"\n")  # テキストを改行コードで分割する
                            txtsc = len(txts)  # 改行コード分割後のリスト要素数
                            for tx in range(txtsc):  # 改行コード分割後のリストを置換
                                txts[tx] = (
                                    txts[tx]
                                    .replace("'", "")
                                    .replace('"', "")
                                    .replace("\xa0", "")
                                    .replace(r"\n", "")
                                    .replace(r"\u2003", " ")
                                    .replace(r"\u3000", " ")
                                )
                                txts[tx] = txts[tx].replace(" ", "")
                                txts[tx] = ChangeBusyu(txts[tx])
                            and_list = set(txts) & set(stt)  # テキストリストとカラムリストで一致する要素を抽出
                            if len(and_list) > 0:  # テキストリストとカラムリストで一致する要素が一つ以上なら
                                TFlag = True  # テキスト取得対象フラグを立てる
                            # テキスト取得対象の処理--------------------------------------------------------
                            if TFlag is True:
                                al = str(and_list).replace("{'", "").replace("'}", "")
                                if al == "同上のうち土地譲渡税額等及":
                                    # 処理中のテキストがカラムリストにあるか判定--------------------------
                                    for sttItem in stt:
                                        st = txts[0]
                                        if st == sttItem:
                                            ColumFlag = True  # カラム対象フラグを立てる
                                            break
                                    # ----------------------------------------------------------------
                                    if ColumFlag is True:  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                        ColumList.append(st)  # 項目リストに代入
                                        ColumFlag = False  # カラム対象フラグ解除
                                    else:
                                        if TFlag is True:
                                            if not txts[0] == "" and not txts[0] == " ":
                                                TxtList.append(st)  # テキストリストに代入
                                    # ----------------------------------------------------------------
                                else:
                                    for tc in range(txtsc):  # 分割後のリストループ
                                        # ヘッダー判定後の処理--------------------------------------------------
                                        if tc % 2 == 0:  # 2で割り切れたらヘッダー
                                            # 処理中のテキストがカラムリストにあるか判定--------------------------
                                            for sttItem in stt:
                                                st = txts[tc]
                                                if st == sttItem:
                                                    ColumFlag = True  # カラム対象フラグを立てる
                                                    break
                                            # ----------------------------------------------------------------
                                            if (
                                                ColumFlag is True
                                            ):  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                                ColumList.append(st)  # 項目リストに代入
                                                ColumFlag = False  # カラム対象フラグ解除
                                            else:
                                                if TFlag is True:
                                                    if (
                                                        not txts[tc] == ""
                                                        and not txts[tc] == " "
                                                    ):
                                                        TxtList.append(st)  # テキストリストに代入
                                            # ----------------------------------------------------------------
                                        else:
                                            if (
                                                not txts[tc] == ""
                                                and not txts[tc] == " "
                                            ):
                                                tsc = txts[tc]
                                                if tsc.startswith(" ") is True:
                                                    tsc = tsc.replace(" ", "")
                                                TxtList.append(tsc)  # テキストリストに代入
                                        # ---------------------------------------------------------------------
                                    TFlag = False
                                # -----------------------------------------------------------------------------
                        else:
                            txts = txt.split(r"\n")  # テキストを改行コードで分割する
                            txtsc = len(txts)  # 改行コード分割後のリスト要素数
                            for tx in range(txtsc):  # 改行コード分割後のリストを置換
                                txts[tx] = (
                                    txts[tx]
                                    .replace("'", "")
                                    .replace('"', "")
                                    .replace("\xa0", "")
                                    .replace(r"\n", "")
                                    .replace(r"\u2003", " ")
                                    .replace(r"\u3000", " ")
                                )
                                txts[tx] = txts[tx].replace(" ", "")
                                txts[tx] = ChangeBusyu(txts[tx])
                            and_list = set(txts) & set(stt)  # テキストリストとカラムリストで一致する要素を抽出
                            if len(and_list) > 0:  # テキストリストとカラムリストで一致する要素が一つ以上なら
                                TFlag = True  # テキスト取得対象フラグを立てる
                            # テキスト取得対象の処理--------------------------------------------------------
                            if TFlag is True:
                                al = str(and_list).replace("{'", "").replace("'}", "")
                                if al == "同上のうち土地譲渡税額等及":
                                    # 処理中のテキストがカラムリストにあるか判定--------------------------
                                    for sttItem in stt:
                                        st = txts[0]
                                        if st == sttItem:
                                            ColumFlag = True  # カラム対象フラグを立てる
                                            break
                                    # ----------------------------------------------------------------
                                    if ColumFlag is True:  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                        ColumList.append(st)  # 項目リストに代入
                                        ColumFlag = False  # カラム対象フラグ解除
                                    else:
                                        if TFlag is True:
                                            if not txts[0] == "" and not txts[0] == " ":
                                                TxtList.append(st)  # テキストリストに代入
                                    # ----------------------------------------------------------------
                                else:
                                    for tc in range(txtsc):  # 分割後のリストループ
                                        # ヘッダー判定後の処理--------------------------------------------------
                                        if tc % 2 == 0:  # 2で割り切れたらヘッダー
                                            # 処理中のテキストがカラムリストにあるか判定--------------------------
                                            for sttItem in stt:
                                                st = txts[tc]
                                                if st == sttItem:
                                                    ColumFlag = True  # カラム対象フラグを立てる
                                                    break
                                            # ----------------------------------------------------------------
                                            if (
                                                ColumFlag is True
                                            ):  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                                st = (
                                                    st.replace("'", "")
                                                    .replace('"', "")
                                                    .replace("\xa0", "")
                                                    .replace(r"\n", "")
                                                    .replace(r"\u2003", " ")
                                                    .replace(r"\u3000", " ")
                                                )
                                                st = st.replace(" ", "")
                                                if (
                                                    not st == ""
                                                    and not st == " "
                                                    and not st == "殿"
                                                    and not st == "法人税額の計算"
                                                    and not st == "地方法人税額の計算"
                                                ):
                                                    ColumList.append(st)  # 項目リストに代入
                                                    ColumFlag = False  # カラム対象フラグ解除
                                            else:
                                                if TFlag is True:
                                                    if (
                                                        not txts[tc] == ""
                                                        and not txts[tc] == " "
                                                    ):
                                                        TxtList.append(st)  # テキストリストに代入
                                            # ----------------------------------------------------------------
                                        else:
                                            if (
                                                not txts[tc] == ""
                                                and not txts[tc] == " "
                                                and ColumFlag is not False
                                            ):
                                                tsc = txts[tc]
                                                if tsc.startswith(" ") is True:
                                                    tsc = tsc.replace(" ", "")
                                                TxtList.append(tsc)  # テキストリストに代入
                                        # ---------------------------------------------------------------------
                    cicount += 1
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsActionOsirase(stt, cellsList, ColumList, TxtList):  # 申告のお知らせ処理
    try:
        ColumFlag = False  # カラム対象フラグ
        TFlag = False  # テキスト取得対象フラグ
        txts = []
        cicount = 0
        for cellsListItem in cellsList:  # セルループ
            if cicount >= 2:
                for cLIItem in cellsListItem:  # セルループ
                    if len(ColumList) == 4:
                        ColumList.append("項目")
                    elif len(ColumList) == 7:
                        ColumList.append("社名")
                    elif len(ColumList) == 8:
                        ColumList.append("氏名")
                    elif len(ColumList) == 9:
                        ColumList.append("相手氏名")
                    elif len(ColumList) == 10:
                        ColumList.append("件名")
                    elif len(TxtList) == 10:
                        TxtList.append("申告のお知らせ")
                    txt = repr(cLIItem.text)  # 改行コードの数でセル分割の判定
                    nc = txt.count(r"\n")  # テキスト内の改行コードの数
                    if txt.endswith(r"\n") is True:  # テキスト内の最後の改行コードは省く
                        nc -= 1
                    if nc >= 2:  # テキスト内に改行コードが2つ以上ある場合
                        txts = txt.split(r"\n")  # テキストを改行コードで分割する
                        txtsc = len(txts)  # 改行コード分割後のリスト要素数
                        for tx in range(txtsc):  # 改行コード分割後のリストを置換
                            txts[tx] = (
                                txts[tx]
                                .replace("'", "")
                                .replace('"', "")
                                .replace("\xa0", "")
                                .replace(r"\n", "")
                                .replace(r"\u2003", " ")
                                .replace(r"\u3000", " ")
                            )
                            txts[tx] = txts[tx].replace(" ", "")
                            txts[tx] = ChangeBusyu(txts[tx])
                        and_list = set(txts) & set(stt)  # テキストリストとカラムリストで一致する要素を抽出
                        if len(and_list) > 0:  # テキストリストとカラムリストで一致する要素が一つ以上なら
                            TFlag = True  # テキスト取得対象フラグを立てる
                        # テキスト取得対象の処理--------------------------------------------------------
                        if TFlag is True:
                            for tc in range(txtsc):  # 分割後のリストループ
                                # ヘッダー判定後の処理--------------------------------------------------
                                if tc % 2 == 0:  # 2で割り切れたらヘッダー
                                    # 処理中のテキストがカラムリストにあるか判定--------------------------
                                    for sttItem in stt:
                                        st = txts[tc]
                                        if st == sttItem:
                                            ColumFlag = True  # カラム対象フラグを立てる
                                            break
                                    # ----------------------------------------------------------------
                                    # テキスト内の改行コード総数が3で割り切れる場合------------------------
                                    if nc % 3 == 0 and tc == 0:  # テキスト内初回の処理
                                        ColumList.append("項目")  # 項目リストに代入
                                    if ColumFlag is True:  # テキスト内2回目以降でカラム対象フラグが立っている処理
                                        ColumList.append(st)  # 項目リストに代入
                                        ColumFlag = False  # カラム対象フラグ解除
                                    else:
                                        if TFlag is True:
                                            if (
                                                not txts[tc] == ""
                                                and not txts[tc] == " "
                                            ):
                                                TxtList.append(st)  # テキストリストに代入
                                    # ----------------------------------------------------------------
                                else:
                                    if not txts[tc] == "" and not txts[tc] == " ":
                                        tsc = txts[tc]
                                        if tsc.startswith(" ") is True:
                                            tsc = tsc.replace(" ", "")
                                        TxtList.append(tsc)  # テキストリストに代入
                                # ---------------------------------------------------------------------
                            TFlag = False
                        # -----------------------------------------------------------------------------
                    else:
                        if not len(TxtList) >= 10:
                            txts = repr(
                                cLIItem.text.replace(r"\n", "")
                            )  # 改行コード判定の為repr
                            txts = (
                                txts.replace("'", "")
                                .replace('"', "")
                                .replace("\xa0", "")
                                .replace(r"\n", "")
                                .replace(r"\u2003", " ")
                                .replace(r"\u3000", " ")
                            )
                            txts = txts.replace(" ", "")
                            txts = ChangeBusyu(txts)
                            if txts.startswith(" ") is True:
                                txts.replace(" ", "")
                            if not txts == "" and not txts == " " and not txts == "殿":
                                TxtList.append(txts)  # テキストリストに代入
                cicount += 1
            else:
                cicount += 1
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsActionTKC(stt, cellsList, ColumList, TxtList, Sbtext):  # 主にeltax処理
    try:
        Sbtext = Sbtext.split("\n")
        print(Sbtext)
        for SbtextItem in Sbtext:  # セルループ
            SbtextItem = SbtextItem.replace("\u3000", "")
            SbtextItem = SbtextItem.replace(" ", "")
            # 処理中のテキストがカラムリストにあるか判定--------------------------
            for sttItem in stt:
                if (
                    sttItem in SbtextItem
                    and not SbtextItem == ""
                    and "ＱＲコード" not in SbtextItem
                ):
                    if sttItem + ":" in SbtextItem:
                        SbtextItem = SbtextItem.replace(sttItem + ":", sttItem + "::")
                    elif sttItem + "：" in SbtextItem:
                        SbtextItem = SbtextItem.replace(sttItem + "：", sttItem + "::")
                    else:
                        SbtextItem = SbtextItem.replace(sttItem, sttItem + ":")
                        SbtextItem = SbtextItem.replace(sttItem + ":", sttItem + "::")
                        if "受付日::時：" in SbtextItem:
                            SbtextItem = SbtextItem.replace("受付日::時：", "受付日時::")
                        elif "受付日::時:" in SbtextItem:
                            SbtextItem = SbtextItem.replace("受付日::時:", "受付日時::")
                        elif "発行元::の" in SbtextItem:
                            SbtextItem = SbtextItem.replace("発行元::の", "発行元の")
                    SbtextItem = SbtextItem.replace(sttItem + "：", sttItem + "::")
                    SbtextItem = SbtextItem.replace("\u3000", "")
                    SbtextItem = SbtextItem.replace(" ", "")
                    if "::" in SbtextItem:
                        if (
                            "地方税ポータルシステム(eLTAX)のメッセージボックスに格納された受付通知は以下の通りです。:"
                            in SbtextItem
                        ):
                            SI = SbtextItem.split(
                                "地方税ポータルシステム(eLTAX)のメッセージボックスに格納された受付通知は以下の通りです。:"
                            )
                            SI.pop(0)
                            SI = SI[0].split("::")
                        else:
                            # tomlリストと抽出内容の文字数が一致する場合
                            Mojisa = len(SbtextItem) - len(sttItem)
                            if Mojisa >= 50:
                                print("Skip行")
                                break
                            else:
                                SI = SbtextItem.split("::")
                        ColumList.append(SI[0])
                        TxtList.append(SI[1])
                        break
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsMoveAction(stt, cellsList, ColumList, TxtList, Sbtext):  # 主にeltax処理
    try:
        ColumFlag = False  # カラム対象フラグ
        TFlag = False  # テキスト取得対象フラグ
        Getu = False  # 月数換算項目フラグ
        for cellsListItem in cellsList:  # セルループ
            for CLIItem in cellsListItem:
                CLIT = (
                    CLIItem.text.replace(" ", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                CLIT = CLIT.replace("\n", "")
                CLIT = CLIT.replace(" ", "")
                if CLIT == "月数換算":
                    Getu = True  # 月数換算項目フラグ
                # 処理中のテキストがカラムリストにあるか判定--------------------------
                for sttItem in stt:
                    if sttItem == CLIT and not CLIT == "":
                        ColumFlag = True
                        TFlag = False
                        break
                    else:
                        ColumFlag = False
                        TFlag = True
                if ColumFlag is True:
                    CLIT = CLIT.replace("\n", "")
                    CLIT = CLIT.replace(" ", "")
                    ColumList.append(CLIT)
                    ColumFlag = False
                elif TFlag is True:
                    if Getu is True:
                        CLITs = ""
                        CLITList = re.findall(r"\d+", CLIT)  # 数値のみ取り出すてリスト化
                        CLIRow = 0
                        for CLITListItem in CLITList:
                            if len(CLITList) == 3:
                                CLIsp = 1
                            else:
                                CLIsp = 2
                            if CLIRow == CLIsp:
                                CLITs = CLITs + "/" + CLITListItem
                            else:
                                CLITs += CLITListItem
                            CLIRow += 1
                        CLIT = CLITs
                        CLIT = (
                            CLIT.replace(" ", "")
                            .replace("'", "")
                            .replace('"', "")
                            .replace("\xa0", "")
                            .replace(r"\n", "")
                            .replace(r"\u2003", " ")
                            .replace(r"\u3000", " ")
                        )
                        CLIT = (
                            CLIT.replace(" ", "")
                            .replace("'", "")
                            .replace('"', "")
                            .replace("\xa0", "")
                            .replace("\n", "")
                            .replace("\u2003", "")
                            .replace("\u3000", "")
                        )
                        CLIT = CLIT.replace("\n", "")
                        CLIT = CLIT.replace(" ", "")
                        TxtList.append(CLIT)
                        Getu = False
                    else:
                        CLIT = (
                            CLIT.replace(" ", "")
                            .replace("'", "")
                            .replace('"', "")
                            .replace("\xa0", "")
                            .replace(r"\n", "")
                            .replace(r"\u2003", " ")
                            .replace(r"\u3000", " ")
                        )
                        CLIT = (
                            CLIT.replace(" ", "")
                            .replace("'", "")
                            .replace('"', "")
                            .replace("\xa0", "")
                            .replace("\n", "")
                            .replace("\u2003", "")
                            .replace("\u3000", "")
                        )
                        CLIT = CLIT.replace("\n", "")
                        CLIT = CLIT.replace(" ", "")
                        TxtList.append(CLIT)
                    TFlag = False
        return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CATKCList(stt, Kokuzei, Tihouzei, Mcells, ColumList, TxtList):
    try:
        if Kokuzei is True or Tihouzei is True:  # フラグが立っていれば
            for cellsListItem in Mcells:  # セルループ
                for CLIItem in cellsListItem:
                    CLIT = (
                        CLIItem.text.replace(" ", "")
                        .replace("'", "")
                        .replace('"', "")
                        .replace("\xa0", "")
                        .replace(r"\n", "")
                        .replace(r"\u2003", " ")
                        .replace(r"\u3000", " ")
                    )
                    CLIT = CLIT.replace(" ", "").replace("\n", "")
                    CLIT = CLIT.replace("\u3000", "")
                    # 処理中のテキストがカラムリストにあるか判定-----------
                    for sttItem in stt:
                        if sttItem in CLIT and not CLIT == "":
                            ColumFlag = True
                            TFlag = False
                            break
                        else:
                            ColumFlag = False
                            TFlag = True
                    # -------------------------------------------------
                    CLIT = CLIT.replace(":", "::").replace("：", "::")
                    CLITList = CLIT.split("::")
                    if ColumFlag is True:
                        ColumList.append(CLITList[0])
                        ColumFlag = False
                        TxtList.append(CLITList[1])
                        TFlag = False
            return True, ColumList, TxtList
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsActionTKCList(
    CDict,
    Settingtoml,
    stt,
    cellsList,
    ColumList,
    TxtList,
    Sbtext,
    path_pdf,
    page,
    SCode,
    DLCList,
    PageList,
):  # 主にeltax処理
    try:
        CellPosition = []
        Sbtext = Sbtext.split("\n")
        SBR = len(Sbtext) - 1
        for SbtextItem in reversed(Sbtext):
            if SbtextItem == "":
                Sbtext.pop(SBR)
            elif "国税受付システムからの「受信通知」の内容" in SbtextItem:
                Relist = SbtextItem.split("国税受付システムからの「受信通知」の内容")
                Sbtext[SBR] = Sbtext[SBR].replace(Relist[0], "")
            SBR -= 1
        print(Sbtext)
        sKR = 0
        for SbtextItem in Sbtext:  # セルループ
            strKey = (
                SbtextItem.replace(" ", "")
                .replace("'", "")
                .replace('"', "")
                .replace("\xa0", "")
                .replace(r"\n", "")
                .replace(r"\u2003", " ")
                .replace(r"\u3000", " ")
            )
            if "受信通知】" in strKey or "受付通知】" in strKey:
                CellPosition.append(sKR + 1)
            sKR += 1
        lcp = len(CellPosition)
        MSbList = []
        for lcpC in range(lcp):  # セルポジションリストループ
            FP = CellPosition[lcpC]
            try:
                LP = CellPosition[lcpC + 1]
            except:
                LP = len(Sbtext)
            for Sr in range(FP, LP):
                SBT = (
                    Sbtext[Sr]
                    .replace(" ", "")
                    .replace("'", "")
                    .replace('"', "")
                    .replace("\xa0", "")
                    .replace(r"\n", "")
                    .replace(r"\u2003", " ")
                    .replace(r"\u3000", " ")
                )
                SBT = SBT.replace("\u3000", "").replace(" ", "")
                MSbList.append(SBT)
            # 固定パラメーターを挿入------------------------------------------------
            ColumList.append("URL")
            ColumList.append("ページ")
            ColumList.append("コード")
            TxtList.append(path_pdf.replace("/", "\\"))
            if len(PageList) == 0:
                TxtList.append(str(page + 1) + "ページ")
            else:
                Pagestr = []
                Pagestr.append(str(page + 1))
                for PageListItem in PageList:
                    PagesPlus = str(PageListItem + page + 2)
                    Pagestr.append(PagesPlus)
                result = "-".join(s for s in Pagestr)
                TxtList.append(result + "ページ")
            TxtList.append(str(SCode))
            # --------------------------------------------------------------------
            for SbtextItem in MSbList:  # セルループ
                if SbtextItem == "申告の種類":
                    print(SbtextItem)
                # 処理中のテキストがカラムリストにあるか判定--------------------------
                for sttItem in stt:
                    if (
                        sttItem in SbtextItem
                        and not SbtextItem == ""
                        and "ＱＲコード" not in SbtextItem
                    ):
                        if sttItem + ":" in SbtextItem:
                            SbtextItem = SbtextItem.replace(
                                sttItem + ":", sttItem + "::"
                            )
                        elif sttItem + "：" in SbtextItem:
                            SbtextItem = SbtextItem.replace(
                                sttItem + "：", sttItem + "::"
                            )
                        else:
                            SbtextItem = SbtextItem.replace(sttItem, sttItem + ":")
                            SbtextItem = SbtextItem.replace(
                                sttItem + ":", sttItem + "::"
                            )
                            if "受付日::時：" in SbtextItem:
                                SbtextItem = SbtextItem.replace("受付日::時：", "受付日時::")
                            elif "受付日::時:" in SbtextItem:
                                SbtextItem = SbtextItem.replace("受付日::時:", "受付日時::")
                            elif "後日、発行元::の担当者から" in SbtextItem:
                                SbtextItem = SbtextItem.replace("発行元::", "発行元")
                            elif "円:" in SbtextItem:
                                SbtextItem = SbtextItem.replace("円:", "円::")
                        SbtextItem = SbtextItem.replace(sttItem + "：", sttItem + "::")
                        SbtextItem = SbtextItem.replace("\u3000", "")
                        SbtextItem = SbtextItem.replace(" ", "")
                        if ":提出先::" in SbtextItem:
                            SbtextItem = SbtextItem.replace(":提出先::", ":提出先:")
                        if "::" in SbtextItem:
                            SI = SbtextItem.split("::")
                            if SI[0] == sttItem or "納税者の氏名又は名称" == SI[0]:
                                SII = SI[0]
                                if len(SII) == 0:
                                    ColumList.append(SII)
                                else:
                                    ColumList.append(SII)
                            elif "特別法人事業税申告納付税額" in SbtextItem:
                                if sttItem == "法人事業税申告納付税額":
                                    ColumList.append("特別法人事業税申告納付税額")
                            elif "法人市民税（法人税割）申告納付税額" in SbtextItem:
                                if sttItem == "法人事業税申告納付税額":
                                    SIRow = len(SI)
                                    for SIItemRow in range(SIRow):
                                        if sttItem in SI[SIItemRow]:
                                            ColumList.append(SI[SIItemRow])
                                            SI[1] = SI[SIItemRow + 1]
                                else:
                                    SIRow = len(SI)
                                    for SIItemRow in range(SIRow):
                                        if sttItem in SI[SIItemRow]:
                                            ColumList.append(sttItem)
                            else:
                                SII = SI[0].split(sttItem)
                                if len(SII) == 0:
                                    ColumList.append(SII)
                                else:
                                    ColumList.append(SII[1])
                            TxtList.append(SI[1])
                            break
            DLP = DiffListPlus(CDict, Settingtoml, ColumList, TxtList, "")
            if DLP[0] is True:
                DLCList.append(DLP[1])
            ColumList = []
            TxtList = []
            MSbList = []
        return True, DLCList, ""
    except:
        return False, "", ""


# --------------------------------------------------------------------------------------------------
def CellsImport(
    CDict,
    Settingtoml,
    SCode,
    path_pdf,
    tables,
    page,
    TaxType,
    Sbtext,
    DLCList,
    NextFlag,
    PageList,
):
    try:
        # MeUrl = os.getcwd().replace("\\", "/")  # 自分のパス
        # # toml読込---------------------------------------------------------------------------------
        # with open(
        #     MeUrl + r"/RPAPhoto/PDFReadForList/Setting.toml", encoding="utf-8"
        # ) as f:
        #     Settingtoml = toml.load(f)
        #     print(Settingtoml)
        # # ----------------------------------------------------------------------------------------
        ColumList = []  # 項目リスト
        TxtList = []  # テキストリスト
        try:
            cellsList = list(tables._tables[0].cells)  # セル情報のリスト
        except:
            cellsList = []
        # 固定パラメーターを挿入------------------------------------------------
        ColumList.append("URL")
        ColumList.append("ページ")
        ColumList.append("コード")
        TxtList.append(path_pdf.replace("/", "\\"))
        TxtList.append(str(page + 1) + "ページ")
        TxtList.append(str(SCode))
        # --------------------------------------------------------------------
        if "etaxosirase" == TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsActionOsirase(stt, cellsList, ColumList, TxtList)
        elif "etax3retu" == TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsAction(stt, cellsList, ColumList, TxtList)
        elif "eltaxList" == TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsAction(stt, cellsList, ColumList, TxtList)
        elif "etaxList" == TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsAction(stt, cellsList, ColumList, TxtList)
        elif "etaxjigyounendo" == TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsActionJigyounendo(stt, cellsList, ColumList, TxtList)
        elif "TKC" in TaxType:
            if "TKC3" == TaxType:
                stt = Settingtoml["CsvSaveEnc"][TaxType]
                ColumList = []
                TxtList = []
                CA = CellsActionTKCList(
                    CDict,
                    Settingtoml,
                    stt,
                    cellsList,
                    ColumList,
                    TxtList,
                    Sbtext,
                    path_pdf,
                    page,
                    SCode,
                    DLCList,
                    PageList,
                )
            elif "TKC13" == TaxType:
                stt = Settingtoml["CsvSaveEnc"][TaxType]
                ColumList = []
                TxtList = []
                CA = CellsActionTKCList(
                    CDict,
                    Settingtoml,
                    stt,
                    cellsList,
                    ColumList,
                    TxtList,
                    Sbtext,
                    path_pdf,
                    page,
                    SCode,
                    DLCList,
                    PageList,
                )
            else:
                stt = Settingtoml["CsvSaveEnc"][TaxType]
                CA = CellsActionTKC(stt, cellsList, ColumList, TxtList, Sbtext)
        elif "etaxsyouhicyuukan" in TaxType:
            stt = Settingtoml["CsvSaveEnc"][TaxType]
            CA = CellsMoveAction(stt, cellsList, ColumList, TxtList, Sbtext)
        elif "MJS" in TaxType:
            if "MJSOutList" == TaxType:
                stt = Settingtoml["CsvSaveEnc"]["MJS"]
                rtt = Settingtoml["OUTLIST"]["MJSOutList"]
                CA = CellsActionMJS(stt, rtt, cellsList, ColumList, TxtList, Sbtext)
            else:
                stt = Settingtoml["CsvSaveEnc"][TaxType]
                CA = CellsAction(stt, cellsList, ColumList, TxtList)
        else:
            if TaxType == "No":
                print("テキスト内容からTaxTypeを判定できませんでした。")
                return False, "テキスト内容からTaxTypeを判定できませんでした。", ""
            else:
                stt = Settingtoml["CsvSaveEnc"][TaxType]
                CA = CellsAction(stt, cellsList, ColumList, TxtList)
        if CA[0] is True:
            return CA
        else:
            return CA
    except:
        print("CellsImport中のエラー")
        return False, "CellsImport中のエラー", ""


# ----------------------------------------------------------------------------------------------
