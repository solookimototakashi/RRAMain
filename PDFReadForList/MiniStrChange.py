import pandas as pd
import numpy as np
import os


def ChangeBusyu(strs):
    try:
        MeUrl = os.getcwd().replace("\\", "/")  # 自分のパス
        df = pd.read_csv(MeUrl + "/Function/cmap置換.csv", encoding="utf8")
        lstr = len(strs)
        dfRow = np.array(df).shape[0]
        Cst = ""
        for s in range(lstr):
            CF = False
            for d in range(dfRow):
                dfData = df.iloc[d]
                Key = dfData["検索"]
                Tar = dfData["置換"]
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
