# ------------------------------------------------------------------------------------
etaxsyotoku = []
etaxsyotoku2 = []
etaxsyotoku3 = []
etaxsyotoku4 = []
etaxsyotoku5 = []
etaxosirase = []
eltaxList = []  # eltaxリスト
eltaxList2 = []
eltaxList3 = []
eltaxList4 = []
eltaxList5 = []
eltaxList6 = []
eltaxList7 = []
eltaxList8 = []
eltaxList9 = []
eltaxList10 = []
eltaxList11 = []
eltaxList12 = []
eltaxPreList = []
eltaxPreList2 = []
eltaxPreList3 = []
etaxList = []
etaxList2 = []
etaxList3 = []
etaxList4 = []
etaxList5 = []
etaxList6 = []
etaxList7 = []
etaxList8 = []
etaxList9 = []
etaxList10 = []
etaxList11 = []
etaxList12 = []
etaxList13 = []
etaxList14 = []
etaxList15 = []
etaxList16 = []
etaxList17 = []
etax3retu = []
etax3retu2 = []
etax3retu3 = []
etax3retu4 = []
etaxsyouhitodoke = []
etaxjigyounendo = []
etaximage = []
etaxnouhu = []
etaxnouhu2 = []
etaxnouhu3 = []
etaxnouhu4 = []
etaxnouhu5 = []
etaxnouhu6 = []
etaxsyouhicyuukan = []
etaxsyouhisinkoku = []
etaxseikyuhakkou = []
Syoukyaku = []
KakusinOsirase = []
TKCuketuke = []
TKCuketuke2 = []
TKC = []
TKCL = []
TKC2 = []
TKC3 = []
TKC4 = []
TKC5 = []
TKC6 = []
TKC7 = []
TKC8 = []
TKC9 = []
TKC10 = []
TKC11 = []
TKC12 = []
TKC13 = []
TKC14 = []
TKC15 = []
TKC16 = []
TKC17 = []
TKC18 = []
TKC19 = []
TKC20 = []
TKC21 = []
TKC22 = []
TKC23 = []
TKC24 = []
TKC25 = []
TKC26 = []
TKC27 = []
TKC28 = []
TKC29 = []
TKC30 = []
TKC31 = []
TKC32 = []
TKC33 = []
TKC34 = []
TKC35 = []
MJS = []
MJS2 = []
MJS3 = []
MJS4 = []
MJS5 = []
MJS6 = []
MJS7 = []
MJS8 = []
MJS9 = []
MJS10 = []
MJS11 = []
MJS12 = []
MJS13 = []
MJS14 = []
MJS15 = []
MJS16 = []
MJS17 = []
MJS18 = []
MJS19 = []
MJS20 = []
MJS21 = []
MJS22 = []
MJS23 = []
MJS24 = []
MJS25 = []
MJS26 = []
MJS27 = []
MJS28 = []
MJS29 = []
MJS30 = []
MJS31 = []
MJSImage = []
ErrList = []  # エラーリスト
SubErrList = []
NotAccessList = []
# 関数の引数様にDict格納-------------------------------------------------------------------------
CSVIndexSortFuncArray = {
    "etaxsyotoku": etaxsyotoku,
    "etaxsyotoku2": etaxsyotoku2,
    "etaxsyotoku3": etaxsyotoku3,
    "etaxsyotoku4": etaxsyotoku4,
    "etaxsyotoku5": etaxsyotoku5,
    "etaxosirase": etaxosirase,
    "eltaxList": eltaxList,
    "eltaxList2": eltaxList2,
    "eltaxList3": eltaxList3,
    "eltaxList4": eltaxList4,
    "eltaxList5": eltaxList5,
    "eltaxList6": eltaxList6,
    "eltaxList7": eltaxList7,
    "eltaxList8": eltaxList8,
    "eltaxList9": eltaxList9,
    "eltaxList10": eltaxList10,
    "eltaxList11": eltaxList11,
    "eltaxList12": eltaxList12,
    "eltaxPreList": eltaxPreList,
    "eltaxPreList2": eltaxPreList2,
    "eltaxPreList3": eltaxPreList3,
    "etaxList": etaxList,
    "etaxList2": etaxList2,
    "etaxList3": etaxList3,
    "etaxList4": etaxList4,
    "etaxList5": etaxList5,
    "etaxList6": etaxList6,
    "etaxList7": etaxList7,
    "etaxList8": etaxList8,
    "etaxList9": etaxList9,
    "etaxList10": etaxList10,
    "etaxList11": etaxList11,
    "etaxList12": etaxList12,
    "etaxList13": etaxList13,
    "etaxList14": etaxList14,
    "etaxList15": etaxList15,
    "etaxList16": etaxList16,
    "etaxList17": etaxList17,
    "etax3retu": etax3retu,
    "etax3retu2": etax3retu2,
    "etax3retu3": etax3retu3,
    "etax3retu4": etax3retu4,
    "etaxsyouhitodoke": etaxsyouhitodoke,
    "etaxjigyounendo": etaxjigyounendo,
    "etaximage": etaximage,
    "etaxnouhu": etaxnouhu,
    "etaxnouhu2": etaxnouhu2,
    "etaxnouhu3": etaxnouhu3,
    "etaxnouhu4": etaxnouhu4,
    "etaxnouhu5": etaxnouhu5,
    "etaxnouhu6": etaxnouhu6,
    "etaxsyouhicyuukan": etaxsyouhicyuukan,
    "etaxsyouhisinkoku": etaxsyouhisinkoku,
    "etaxseikyuhakkou": etaxseikyuhakkou,
    "Syoukyaku": Syoukyaku,
    "KakusinOsirase": KakusinOsirase,
    "TKCuketuke": TKCuketuke,
    "TKCuketuke2": TKCuketuke2,
    "TKC": TKC,
    "TKCL": TKCL,
    "TKC2": TKC2,
    "TKC3": TKC3,
    "TKC4": TKC4,
    "TKC5": TKC5,
    "TKC6": TKC6,
    "TKC7": TKC7,
    "TKC8": TKC8,
    "TKC9": TKC9,
    "TKC10": TKC10,
    "TKC11": TKC11,
    "TKC12": TKC12,
    "TKC13": TKC13,
    "TKC14": TKC14,
    "TKC15": TKC15,
    "TKC16": TKC16,
    "TKC17": TKC17,
    "TKC18": TKC18,
    "TKC19": TKC19,
    "TKC20": TKC20,
    "TKC21": TKC21,
    "TKC22": TKC22,
    "TKC23": TKC23,
    "TKC24": TKC24,
    "TKC25": TKC25,
    "TKC26": TKC26,
    "TKC27": TKC27,
    "TKC28": TKC28,
    "TKC29": TKC29,
    "TKC30": TKC30,
    "TKC31": TKC31,
    "TKC32": TKC32,
    "TKC33": TKC33,
    "TKC34": TKC34,
    "TKC35": TKC35,
    "MJS": MJS,
    "MJS2": MJS2,
    "MJS3": MJS3,
    "MJS4": MJS4,
    "MJS5": MJS5,
    "MJS6": MJS6,
    "MJS7": MJS7,
    "MJS8": MJS8,
    "MJS9": MJS9,
    "MJS10": MJS10,
    "MJS11": MJS11,
    "MJS12": MJS12,
    "MJS13": MJS13,
    "MJS14": MJS14,
    "MJS15": MJS15,
    "MJS16": MJS16,
    "MJS17": MJS17,
    "MJS18": MJS18,
    "MJS19": MJS19,
    "MJS20": MJS20,
    "MJS21": MJS21,
    "MJS22": MJS22,
    "MJS23": MJS23,
    "MJS24": MJS24,
    "MJS25": MJS25,
    "MJS26": MJS26,
    "MJS27": MJS27,
    "MJS28": MJS28,
    "MJS29": MJS29,
    "MJS30": MJS30,
    "MJS31": MJS31,
    "MJSImage": MJSImage,
    "ErrList": ErrList,
    "SubErrList": SubErrList,
    "NotAccessList": NotAccessList,
}
