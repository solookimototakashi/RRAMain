import PyPDF2
import os
import pandas as pd
import csv
import mojimoji
import unicodedata
import WarekiHenkan as WK
import datetime
import re
from pdfminer.high_level import extract_text
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
import Progress


def dirCheck(Url):
    List = []
    for fd_path, sb_folder, sb_file in os.walk(Url):
        for fil in sb_file:
            List.append([fd_path + "\\" + fil, fil.replace(".pdf", "")])
    print(List)
    Ldf = pd.DataFrame(List)
    return Ldf


def Remoji(text):
    if unicodedata.east_asian_width(text) == "F":  # 全角。全角英数など。
        return mojimoji.zen_to_han(text)
    elif unicodedata.east_asian_width(text) == "H":  # 半角。半角カナなど。
        return text
    elif unicodedata.east_asian_width(text) == "W":  # 全角。漢字やかななど。
        return mojimoji.zen_to_han(text)
    elif unicodedata.east_asian_width(text) == "Na":  # 半角。半角英数など。
        return text
    elif unicodedata.east_asian_width(text) == "A":  # 曖昧。ギリシア文字など。
        return text
    elif unicodedata.east_asian_width(text) == "N":  # 中立。アラビア文字など。
        return text


def dirCreate(Url, No, YearT):
    for fd_path, sb_folder, sb_file in os.walk(Url):
        for fol in sb_folder:
            Target = fd_path + r"\\" + fol
            if (
                unicodedata.east_asian_width(fol[0]) == "A"
                or unicodedata.east_asian_width(fol[0]) == "N"
            ):
                f = 0
                Newfol = ""
                for folitem in fol:
                    if not f == 0:
                        Newfol += Remoji(folitem)
                    f += 1
                fol = Newfol
            if "～" in fol:
                FolSP = fol.split("～")
            elif "~" in fol:
                FolSP = fol.split("~")
            try:
                Serch = int(FolSP[0])
                if Serch >= 901:
                    LSerch = 9999
                else:
                    LSerch = Serch + 100
                if Serch <= No and LSerch >= No:
                    Nofol = False
                    for fd_path, sb_folder, sb_file in os.walk(Target):
                        for fol in sb_folder:
                            try:
                                fol = re.findall(r"\d+", fol)[0]
                                if int(str(fol)) == No:
                                    print("フォルダ有")
                                    TTarget = Target + r"\\" + fol
                                    for fd_path, sb_folder, sb_file in os.walk(TTarget):
                                        for fol in sb_folder:
                                            if YearT == fol:
                                                Nofol = True
                                                return TTarget + r"\\" + fol
                                        if Nofol is False:
                                            os.mkdir(TTarget + r"\\" + str(YearT))
                                            return TTarget + r"\\" + str(YearT)
                            except:
                                print("Err")
                    try:
                        if Nofol is False:
                            os.mkdir(Target + str(No))
                            TTarget = Target + r"\\" + str(fol)
                            os.mkdir(TTarget + r"\\" + str(YearT))
                            return TTarget + r"\\" + str(YearT)
                    except:
                        os.mkdir(Target + r"\\" + str(No))
                        TTarget = Target + r"\\" + str(No)
                        os.mkdir(TTarget + r"\\" + str(YearT))
                        return TTarget + r"\\" + str(YearT)
            except:
                return False


def dirMarge(Line, PDFUrl, SerchURL, PDFDir, Title, No, YearT):
    # １つのPDFファイルにまとめる
    Title = Title.replace(r"\\", "").replace("\\u3000", " ")
    PR = len(PDFDir)
    merger = PyPDF2.PdfFileMerger()
    for LineR in Line:
        for PRow in range(PR):
            PDFRow = PDFDir.iloc[PRow]
            PN = str(PDFRow[1])
            if str(LineR[0]) == PN:
                MU = str(PDFRow[0])
                print(MU)
                merger.append(MU)
    TURL = dirCreate(SerchURL, No, YearT)
    # 保存ファイル名（先頭と末尾のファイル名で作成）
    merged_file = TURL + r"\\" + Title + ".pdf"
    Matu = "2"
    if os.path.isfile(merged_file) is True:
        merged_file = merged_file.replace(".pdf", "") + Matu + ".pdf"
        while os.path.isfile(merged_file) is True:
            MatuF = Matu
            Matu = str(int(Matu) + 1)
            merged_file = merged_file.replace(MatuF + ".pdf", "") + Matu + ".pdf"
    # 保存
    merger.write(merged_file)
    merger.close()
    for PRow in range(PR):
        PDFRow = PDFDir.iloc[PRow]
        MU = str(PDFRow[0])
        os.remove(MU)
    return merged_file


def PDFMarge(Url, PDFUrl, SerchURL, Title, No, YearT):
    PB = ProgressBar()
    PB.start()
    with open(Url, "r", encoding="utf8") as f:
        LM = csv.reader(f)
        Line = [row for row in LM]
    print(Line)
    PDFDir = dirCheck(PDFUrl)
    print(PDFDir)
    MF = dirMarge(Line, PDFUrl, SerchURL, PDFDir, Title, No, YearT)
    PB.stop()
    return MF


def BeppyouPDFSplit(
    path_pdf,
    PDFDir,
):
    PB = ProgressBar(tk.Tk())
    PB.start()
    output = PyPDF2.PdfFileWriter()
    output2 = PyPDF2.PdfFileWriter()
    output3 = PyPDF2.PdfFileWriter()
    output4 = PyPDF2.PdfFileWriter()
    output5 = PyPDF2.PdfFileWriter()
    op = 0
    op2 = 0
    op3 = 0
    op4 = 0
    op5 = 0
    SPList = []
    fp = open(path_pdf, "rb")  # PDFファイルを読み込み
    parser = PDFParser(fp)  # PDFperserを作成。
    document = PDFDocument(parser)  # PDFperserを格納。
    num_pages = 0  # ページ数格納変数を初期化
    num_pagesList = []
    for page in PDFPage.create_pages(document):  # ページオブジェ分ループ
        num_pages += 1  # ページ数カウント
        num_pagesList.append(num_pages - 1)
    print(num_pages)  # ページ数確認
    # ------------------------------------------------------------------------------------
    try:
        # PDFのページ数分ループ---------------------------------------------------------------------------
        for y in range(num_pages):
            infile = PyPDF2.PdfFileReader(path_pdf, "rb")
            PL = []
            PL.append(y)
            Sbtext = extract_text(
                path_pdf, page_numbers=PL, maxpages=1, codec="utf-8"
            )  # テキストのみ取得できる
            Sbtext = (
                Sbtext.replace("\n", "")
                .replace("\u3000", "")
                .replace("\x0c", "")
                .replace(" ", "")
            )
            if (
                "第六号様式（提出用）" in Sbtext
                or "第六号様式（入力用）" in Sbtext
                or "第六号様式別表九（提出用）" in Sbtext
                or "第六号様式別表九（入力用）" in Sbtext
                or "第二十号様式（提出用）" in Sbtext
                or "第二十号様式（入力用）" in Sbtext
            ):
                print("不要")
            elif "第六号様式（控用）" in Sbtext:
                print(Sbtext)
                p = infile.getPage(y)
                output2.addPage(p)
                op2 += 1
            elif "第六号様式別表九（控用）" in Sbtext:
                print(Sbtext)
                p = infile.getPage(y)
                output3.addPage(p)
                op3 += 1
            elif "第二十号様式（控用）" in Sbtext:
                print(Sbtext)
                p = infile.getPage(y)
                output4.addPage(p)
                op4 += 1
            elif "第二十二号の二様式" in Sbtext:
                print(Sbtext)
                p = infile.getPage(y)
                output5.addPage(p)
                op5 += 1
            elif "個別注記表" in Sbtext:
                print(Sbtext)
                SP = Sbtext.split("個別注記表")
                SP = SP[1].split("Ⅰ.")
                SP = SP[0].replace("自", "").replace("至", "-")
                SP = SP.split("-")
                FD = WK.SeirekiSTRDate(SP[0])
                FD = datetime.datetime.strptime(FD, "%Y/%m/%d")
                SPList.append([y, FD])
                p = infile.getPage(y)
            else:
                print(Sbtext)
                p = infile.getPage(y)
                output.addPage(p)
                op += 1
        if not len(SPList) == 0:
            DcDiff = SPList[0][1] - SPList[1][1]
            if DcDiff.days > 0:
                p = infile.getPage(SPList[0][0])
                output.addPage(p)
                op += 1
            elif DcDiff.days < 0:
                p = infile.getPage(SPList[1][0])
                output.addPage(p)
                op += 1
        splitext = os.path.splitext(path_pdf)
        if not op == 0:
            with open(path_pdf, "wb") as output_file:
                output.write(output_file)
        if not op2 == 0:
            with open(PDFDir + r"\第6号様式（県）" + splitext[1], "wb") as output_file:
                output2.write(output_file)
        if not op3 == 0:
            with open(PDFDir + r"\第6号様式別表9（県）" + splitext[1], "wb") as output_file:
                output3.write(output_file)
        if not op4 == 0:
            with open(PDFDir + r"\第20号様式（市）" + splitext[1], "wb") as output_file:
                output4.write(output_file)
        if not op5 == 0:
            with open(PDFDir + r"\第22号の2様式" + splitext[1], "wb") as output_file:
                output5.write(output_file)
        PB.stop()
        return True
    except:
        PB.stop()
        return False


if __name__ == "__main__":
    PB = Progress.ProgressBar()
    print("")
    PB.mainloop()
    PB.stop()
