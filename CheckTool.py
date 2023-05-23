

#==========================================================================================
#   構造計算書の数値検査プログラムのサブルーチン（ver.3.00）
#
#           一般財団法人日本建築総合試験所
#
#               coded by T.Kanyama  2023/05
#
#==========================================================================================
"""
このプログラムは、構造判定センターに提出される構造計算書（PDF）の検定比（許容応力度に対する部材応力度の比）を精査し、
設定した閾値（デフォルトは0.95）を超える部材を検出するプログラムのツールである。

"""
# pip install pdfminer
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.pdfpage import PDFPage
# from pdfminer.layout import LAParams, LTTextContainer
from pdfminer.layout import LAParams, LTTextContainer, LTContainer, LTTextBox, LTTextLine, LTChar

# pip install pdfrw
from pdfrw import PdfReader
from pdfrw.buildxobj import pagexobj
from pdfrw.toreportlab import makerl

# pip install reportlab
from reportlab.pdfgen import canvas
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm

# pip install pypdf
from pypdf import PdfReader as PR2 # 名前が上とかぶるので別名を使用
import pypdf

# その他のimport
import os,time
import sys
import numpy as np
import logging
import glob
import threading
from multiprocessing import Process,Array
import shutil

kind = ""
version = ""

#@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
# 並列化 = False      # デバッグ時にはFalse,実行時はTrue
並列化 = True

#============================================================================
#  浮動小数点数値を表しているかどうかを判定する関数
#============================================================================
def isfloat(s):  
    try:
        float(s)  # 文字列を実際にfloat関数で変換してみる
    except ValueError:
        return False
    else:
        return True
    #end if
#end def

#============================================================================
#  整数を表しているかどうかを判定する関数
#============================================================================
def isint(s):  
    try:
        int(s)  # 文字列を実際にint関数で変換してみる
    except ValueError:
        return False
    else:
        return True
    #end if
#end def

#============================================================================
#
#   構造計算書のチェックを行うclass
#
#============================================================================

class CheckTool():
    #==================================================================================
    #   オブジェクトのインスタンス化および初期化
    #==================================================================================
    
    def __init__(self):
        # 源真ゴシック等幅フォント
        # GEN_SHIN_GOTHIC_MEDIUM_TTF = "/Library/Fonts/GenShinGothic-Monospace-Medium.ttf"
        GEN_SHIN_GOTHIC_MEDIUM_TTF = "./Fonts/GenShinGothic-Monospace-Medium.ttf"
        self.fontname1 = 'GenShinGothic'
        # IPAexゴシックフォント
        # IPAEXG_TTF = "/Library/Fonts/ipaexg.ttf"
        IPAEXG_TTF = "./Fonts/ipaexg.ttf"
        self.fontname2 = 'ipaexg'
        
        # フォント登録
        pdfmetrics.registerFont(TTFont(self.fontname1, GEN_SHIN_GOTHIC_MEDIUM_TTF))
        pdfmetrics.registerFont(TTFont(self.fontname2, IPAEXG_TTF))
    #end def
    #*********************************************************************************


    #==================================================================================
    #   表紙の文字から構造計算プログラムの種類とバージョンを読み取る関数
    #==================================================================================

    def CoverCheck(self, page, interpreter, device):
        global kind, version

        interpreter.process_page(page)
        # １文字ずつのレイアウトデータを取得
        layout = device.get_result()

        # CharData5 = []
        CharData = []
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                m1 = lt.matrix
                if m1[1] == 0.0 :  # 回転していない文字のみを抽出
                    CharData.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #next

        # その際、CharData2をY座標の高さ順に並び替えるためのリスト「CY」を作成
        CharData2=[]
        CY = []
        for cdata in CharData:
            char2 = cdata[0]
            x0 = cdata[1]
            x1 = cdata[2]
            y0 = cdata[3]
            y1 = cdata[4]
            CharData2.append(cdata)
            CY.append(int(y0))
        #next
        
        # リスト「CY」から降順の並び替えインデックッスを取得
        y=np.argsort(np.array(CY))[::-1]
        t1 = []
            
        if len(CharData2) > 0:  # リストが空でない場合に処理を行う
            CharData3 = []
            # インデックスを用いて並べ替えた「CharData3」を作成
            for i in range(len(y)):
                CharData3.append(CharData2[y[i]])
            #next

            # 同じ高さのY座標毎にデータをまとめる２次元のリストを作成
            CharData4 = []
            i = 0
            for f in CharData3:
                if i==0 :   # 最初の文字のY座標を基準値に採用し、仮のリストを初期化
                    Fline = []
                    Fline.append(f)
                    gy = int(f[3])
                else:
                    if int(f[3])== gy:   # 同じY座標の場合は、リストに文字を追加
                        Fline.append(f)
                    else:           # Y座標が異なる場合は、リストを「CharData4」を保存し、仮のリストを初期化
                        if len(Fline) >= 4:
                            CharData4.append(Fline)
                        gy = int(f[3])
                        Fline = []
                        Fline.append(f)
                    #end if
                #end if
                i += 1
            #next

            # 仮のリストが残っている場合は、リストを「CharData4」を保存
            if len(Fline) >= 4:
                CharData4.append(Fline)
            #end if

            # 次にX座標の順番にデータを並び替える（昇順）
            CharData5 = []
            for F1 in CharData4:    # Y座標が同じデータを抜き出す。                        
                CX = []         # 各データのX座標のデータリストを作成
                for F2 in F1:
                    CX.append(F2[1])
                #next
                
                # リスト「CX」から降順の並び替えインデックッスを取得
                x=np.argsort(np.array(CX))
                
                # インデックスを用いて並べ替えた「F3」を作成
                F3 = []
                t2 = ""
                for i in range(len(x)):
                    F3.append(F1[x[i]])
                    t3 = F1[x[i]][0]
                    t2 += t3
                #next
                # t1 += t2 + "\n"
                t1.append([t2])
                # print(t2,len(F3))
                CharData5.append(F3)
            #next
        #end if

        CharData2 = []
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                if lt.matrix[1] > 0.0 : # 正の回転している文字のみを抽出
                    CharData2.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #next
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                if lt.matrix[1] < 0.0 : # 正の回転している文字のみを抽出
                    CharData2.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #next

        fline = []
        Sflag = False
        tt2 = ""

        fline = []
        Sflag = False
        tt2 = ""
        for F1 in CharData2:
            if not Sflag:
                if F1[0] != " ":
                    fline.append(F1)
                    tt2 += F1[0]
                    Sflag = True
            else:
                if F1[0] == " ":
                    CharData5.append(fline)
                    t1.append([tt2])
                    fline = []
                    tt2 = ""
                    Sflag = False
                else:
                    fline.append(F1)
                    tt2 += F1[0]
                #end if
            #end if
        #next

        if len(fline)>0:
            CharData5.append(fline)
            t1.append([tt2])
        #end if
        kind ="不明"
        version = "不明"
        if len(t1)>0:
            for line in t1:
                # 全角の'：'と'／'を半角に置換
                t2 = line[0].replace(" ","").replace("：",":").replace("／","/")

                if "プログラムの名称" in t2:
                    n = t2.find(":",0)
                    kind = t2[n+1:]
                elif "プログラムバージョン" in t2:
                    n = t2.find(":",0)
                    version = t2[n+1:]
                    break
                #end if
            #next
        #end if
        return kind , version
    #end def
    #*********************************************************************************


    #==================================================================================
    #   各ページから１文字ずつの文字と座標データを抽出し、行毎の文字配列および座標配列を戻す関数
    #       検定比表に使用する関数
    #==================================================================================

    def MakeChar(self, page, interpreter, device):

        interpreter.process_page(page)
        # １文字ずつのレイアウトデータを取得
        layout = device.get_result()

        CharData = []
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                m1 = lt.matrix
                if m1[1] == 0.0 :  # 回転していない文字のみを抽出
                    CharData.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #next

        # その際、CharData2をY座標の高さ順に並び替えるためのリスト「CY」を作成
        CharData2=[]
        CY = []
        for cdata in CharData:
            char2 = cdata[0]
            x0 = cdata[1]
            x1 = cdata[2]
            y0 = cdata[3]
            y1 = cdata[4]
            
            CharData2.append(cdata)
            CY.append(int(y0))
        #next
        
        # リスト「CY」から降順の並び替えインデックッスを取得
        y=np.argsort(np.array(CY))[::-1]

        if len(CharData2) > 0:  # リストが空でない場合に処理を行う
            CharData3 = []
            # インデックスを用いて並べ替えた「CharData3」を作成
            for i in range(len(y)):
                CharData3.append(CharData2[y[i]])
            #next

            # 同じ高さのY座標毎にデータをまとめる２次元のリストを作成
            CharData4 = []
            i = 0
            for f in CharData3:
                if i==0 :   # 最初の文字のY座標を基準値に採用し、仮のリストを初期化
                    Fline = []
                    Fline.append(f)
                    gy = int(f[3])
                else:
                    if int(f[3])== gy:   # 同じY座標の場合は、リストに文字を追加
                        Fline.append(f)
                    else:           # Y座標が異なる場合は、リストを「CharData4」を保存し、仮のリストを初期化
                        if len(Fline) >= 4:
                            CharData4.append(Fline)
                        gy = int(f[3])
                        Fline = []
                        Fline.append(f)
                    #end if
                #end if
                i += 1
            #next
            # 仮のリストが残っている場合は、リストを「CharData4」を保存
            if len(Fline) >= 4:
                CharData4.append(Fline)
            #end if

            # 次にX座標の順番にデータを並び替える（昇順）
            t1 = []
            CharData5 = []
            for F1 in CharData4:    # Y座標が同じデータを抜き出す。                        
                CX = []         # 各データのX座標のデータリストを作成
                for F2 in F1:
                    CX.append(F2[1])
                #next
                
                # リスト「CX」から降順の並び替えインデックッスを取得
                x=np.argsort(np.array(CX))
                
                # インデックスを用いて並べ替えた「F3」を作成
                F3 = []
                t2 = ""
                for i in range(len(x)):
                    F3.append(F1[x[i]])
                    t3 = F1[x[i]][0]
                    t2 += t3
                #next
                # t1 += t2 + "\n"
                t1.append([t2])
                # print(t2,len(F3))
                CharData5.append(F3)
            #next
        #end if

        CharData2 = []
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                if lt.matrix[1] > 0.0 : # 正の回転している文字のみを抽出
                    CharData2.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #nexr
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                if lt.matrix[1] < 0.0 : # 正の回転している文字のみを抽出
                    CharData2.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end iuf
        #next
        
        fline = []
        Sflag = False
        tt2 = ""
        
        fline = []
        Sflag = False
        tt2 = ""
        for F1 in CharData2:
            if not Sflag:
                if F1[0] != " ":
                    fline.append(F1)
                    tt2 += F1[0]
                    Sflag = True
                #end if
            else:
                if F1[0] == " ":
                    CharData5.append(fline)
                    t1.append([tt2])
                    fline = []
                    tt2 = ""
                    Sflag = False
                else:
                    fline.append(F1)
                    tt2 += F1[0]
                #end if
            #end if
        #next

        if len(fline)>0:
            CharData5.append(fline)
            t1.append([tt2])
        #end if

        return t1 , CharData5
    #end def
    #*********************************************************************************

    #==================================================================================
    #   各ページから１文字ずつの文字と座標データを抽出し、行毎の文字配列および座標配列を戻す関数
    #       検定比図に使用する関数
    #==================================================================================

    def MakeChar2(self, page, interpreter, device):

        interpreter.process_page(page)
        # １文字ずつのレイアウトデータを取得
        layout = device.get_result()

        CharData = []
        CharData2 = []
        for lt in layout:
            if isinstance(lt, LTChar):  # レイアウトデータうち、LTCharのみを取得
                char1 = lt.get_text()   # レイアウトデータに含まれる全文字を取得
                m1 = lt.matrix
                if m1[1] == 0.0 :  # 回転していない文字のみを抽出
                    CharData.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                else:
                    CharData2.append([char1, lt.x0, lt.x1, lt.y0, lt.y1,lt.matrix])
                #end if
            #end if
        #next

        # n = len(CharData)
        # text1 = ""
        y1=CharData[0][3]
        cdata = []
        cline=[]
        tline = ""
        tbox = []
        for c in CharData:
            y=c[3]
            if y == y1:
                cline.append(c)
                tline += c[0]
            else:
                cdata.append(cline)
                tbox.append([tline])
                # text1 += tline + "\n"
                y1=c[3]
                cline=[]
                tline = ""
                cline.append(c)
                tline += c[0]
            #end if
        #next
        if tline != "":
            cdata.append(cline)
            tbox.append([tline])
            # text1 += tline + "\n"
        #end if

        cline=[]
        tline = ""
        flag = False
        for c in CharData2:
            if flag == False:
                if c[0] != " ":
                    cline.append(c)
                    tline += c[0]
                    flag = True
                #end if
            else:
                if c[0] != " ":
                    cline.append(c)
                    tline += c[0]
                    flag = True
                else:
                    cdata.append(cline)
                    tbox.append([tline])
                    cline=[]
                    tline = ""
                    flag = False
                #end if
            #end if
        #next
        if tline != "":
            cdata.append(cline)
            tbox.append([tline])
        #end if

        # print(text1)
        t1 = tbox
        CharData5 = cdata
        

        return t1 , CharData5
    #end def
    #*********************************************************************************

    #==================================================================================
    #   各ページの数値を検索し、閾値を超える数値を四角で囲んだPDFファイルを作成する関数
    #   （SS7用の関数）
    #==================================================================================

    def SS7(self, page, limit, interpreter, device,interpreter2, device2):
        
        #============================================================
        # 構造計算書がSS7の場合の処理
        #============================================================
        pageFlag = False
        ResultData = []
        limit1 = limit
        limit2 = limit
        limit3 = limit
        interpreter.process_page(page)
        layout = device.get_result()
        #
        #   このページに「柱の断面検定表」、「梁の断面検定表」、「壁の断面検定表」、「検定比図」の
        #   文字が含まれている場合のみ数値の検索を行う。
        #
        QDL_Flag = False
        検定表_Flag = False
        柱_Flag = False
        梁_Flag = False
        壁_Flag = False
        ブレース_Flag = False
        杭_Flag = False
        検定比図_Flag = False

        xd = 3      #  X座標の左右に加える余白のサイズ（ポイント）を設定

        mode = ""

        for lt in layout:
            # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎)
            if isinstance(lt, LTTextContainer):
                texts = lt.get_text()
                if "柱の断面検定表"in texts :
                    柱_Flag = True
                    break
                #end if
                if  "梁の断面検定表"in texts:
                    梁_Flag = True
                    break
                #end if
                if "壁の断面検定表"in texts :                               
                    壁_Flag = True
                    break
                #end if
                if "断面算定表"in texts and "杭基礎"in texts:
                    杭_Flag = True
                    break
                #end if
                if "ブレースの断面検定表"in texts :
                    ブレース_Flag = True
                    break
                #end if
                if "検定比図"in texts:
                    検定比図_Flag = True
                    break
                #end if
            #end if
        #next

            
        if 壁_Flag:
            i=0
            for lt in layout:
                # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎)
                if isinstance(lt, LTTextContainer):
                    texts = lt.get_text()
                    if "ブレースの断面検定表"in texts :
                        ブレース_Flag = True
                        壁_Flag = False
                        break
                    #end if
                #enf if
                i += 1
                if i>20:
                    break
                #end if
            #next
        #end if
            
        if 検定比図_Flag:
            mode = "検定比図"
        #end if
        if 柱_Flag :
            mode = "柱の検定表"
        #end if
        if 梁_Flag :
            mode = "梁の検定表"
        #end if
        if 壁_Flag :
            mode = "壁の検定表"
        #end if
        if 杭_Flag :
            mode = "杭の検定表"
        #end if
        if ブレース_Flag :
            mode = "ブレースの検定表"
        #end if


        i = 0
        B_kind = ""
            
        for lt in layout:
            # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎)
            if isinstance(lt, LTTextContainer):
                texts = lt.get_text()
                if "RC柱"in texts or "RC梁"in texts:
                    B_kind = "RC造"
                    break
                #end if
                if "SRC柱"in texts or "SRC梁"in texts:
                    B_kind = "SRC造"
                    break
                #end if
                if "S柱"in texts or "S梁"in texts:
                    B_kind = "S造"
                    break
                #end if
            #end if
        #next

        if mode == "" :     # 該当しない場合はこのページの処理は飛ばす。
            print("No Data")
            return False,[]
        else:
            print(mode)
        #end if

        #=================================================================================================
        #   検定比図のチェック
        #=================================================================================================
        
        if mode == "検定比図" :

            CharLines , CharData = self.MakeChar2(page, interpreter2,device2)

            if len(CharLines) > 0:
                i = -1
                for line in CharLines:
                    i += 1
                    t3 = line[0]
                    CharLine = CharData[i] # １行文のデータを読み込む
                    
                    # line = CharLines[i][0]
                    line2 = ""
                    xx= CharData[i][0][2]
                    for Char in CharData[i]:
                        if Char[1]>xx+3:
                            line2 += " "
                        line2 += Char[0]
                        xx = Char[2]
                    #next
                    items = line2.split()
                    # print(line)
                    # print(items)
                    st = 0
                    # t4 = t3.split()            # 文字列を空白で分割
                    t4 = items

                    if len(t4)>0 and t4[0] != "Super":    # 文字列配列が１個以上ある場合に処理
                        for t5 in t4:
                            t6 = t5.replace("(","").replace(")","").replace(" ","").replace("C","").replace("T","").replace("組","")     # 「検定比」と数値が一緒の場合は除去
                            # print(t6)
                            nn = t3.find(t6,st)   # 数値の文字位置を検索
                            ln = len(t6)

                            # カッコがある場合は左右１文字ずつ追加
                            if "(" in t5:
                                xn1 = 1
                                xn2 = 1
                            elif "C" in t5 or "T" in t5 or "組" in t5:
                                xn1 = 0
                                xn2 = 1
                            else:
                                xn1 = 0
                                xn2 = 0

                            if isfloat(t6) and len(t6)>=4:
                                a = float(t6)
                                if a>=limit1 and a<=1.0:
                                    # 数値がlimit以上の場合はデータに登録
                                    xxx0 = CharLine[nn-xn1][1]
                                    xxx1 = CharLine[nn+ln+xn2-1][2]
                                    if CharLine[nn][5][1] > 0.0:
                                        yyy0 = CharLine[nn][3] - 1.0
                                        yyy1 = CharLine[nn+ln+xn1-1][4] + 1.0
                                    elif CharLine[nn][5][1] < 0.0:
                                        yyy0 = CharLine[nn+ln+xn1-1][3] - 2.0
                                        yyy1 = CharLine[nn][4] + 2.0
                                    else:
                                        yyy0 = CharLine[nn][3]
                                        yyy1 = CharLine[nn][4]
                                    #end if

                                    if ln <=4 :
                                        xxx0 -= xd
                                        xxx1 += xd
                                    #end if
                                    width3 = xxx1 - xxx0
                                    height3 = yyy1 - yyy0
                                    ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                    flag = True
                                    pageFlag = True
                                    val = a
                                    print('val={:.2f}'.format(val))
                                #end if
                            #end if

                            # 数値を検索を開始するを文字数分移動
                            st = nn + ln 
                        #next
                    #end if
                #next
            #end if
                
        #=================================================================================================
        #   柱の検定表のチェック
        #=================================================================================================
                        
        elif mode == "柱の検定表" : 

            CharLines , CharData = self.MakeChar(page, interpreter2,device2)
            
            if B_kind == "RC造" or B_kind == "SRC造" or B_kind == "":
                # =======================================================
                #   RC造およびSRC造の柱の検定表
                # ======================================================= 
                if len(CharLines) > 0:
                    # lines =t1.splitlines()
                    i = -1
                    kmode = False
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        if not kmode :
                            if "検定比" in t3 : # 最初の「検定比」が現れたら「kmode」をTrue
                                kmode = True
                                # 「検定比」の下にある数値だけを検出するためのX座標を取得
                                n = t3.index("検定比")
                                c1 = CharData[i][n]
                                zx0 = c1[1]
                                c2 = CharData[i][n+2]
                                zx1 = c2[2]
                                # print(c1[0],c2[0], zx0, zx1)
                        else:
                            CharLine = CharData[i] # １行文のデータを読み込む
                            t4 = ""
                        
                            for char in CharLine:
                                # kmodeの時には「検定比」の下にある数値だけを検出する。
                                if char[1]>=zx0 and char[2]<=zx1:
                                    t4 += char[0]
                            # t4=t4.replace("検定比","")
                            if isfloat(t4): # 切り取った文字が数値の場合の処理
                                a = float(t4)
                                if a>=limit1 and a<=1.0:
                                    # 数値がlimit以上の場合はデータに登録
                                    nn = t3.index(t4)   # 数値の文字位置を検索
                                    xxx0 = CharLine[nn][1]
                                    xxx1 = CharLine[nn+3][2]
                                    yyy0 = CharLine[nn][3]
                                    yyy1 = CharLine[nn][4]
                                    xxx0 -= xd
                                    xxx1 += xd
                                    width3 = xxx1 - xxx0
                                    height3 = yyy1 - yyy0
                                    ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                    flag = True
                                    pageFlag = True
                                    val = a
                                    print('val={:.2f}'.format(val))

                    i = -1
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        # print(t3)
                        CharLine = CharData[i] # １行文のデータを読み込む
                        t4 = ""
                    
                        for char in CharLine:
                            # kmodeの時には「検定比」の下にある数値だけを検出する。
                            if char[1]>zx1:
                                t4 += char[0]
                        if "検定比" in t4:
                            st = 0
                            n = t3.find("検定比",st)
                            w0 = t4.split()
                            if len(w0)>=1:
                                st = n + 3
                                for w1 in w0:
                                    w2 = w1.replace("検定比","")
                                    if isfloat(w2): # 切り取った文字が数値の場合の処理
                                        a = float(w2)
                                        if a>=limit1 and a<=1.0:
                                            # 数値がlimit以上の場合はデータに登録
                                            n = t3.find(w2,st)   # 数値の文字位置を検索
                                            xxx0 = CharLine[n][1]
                                            xxx1 = CharLine[n+3][2]
                                            yyy0 = CharLine[n][3]
                                            yyy1 = CharLine[n][4]
                                            xxx0 -= xd
                                            xxx1 += xd
                                            width3 = xxx1 - xxx0
                                            height3 = yyy1 - yyy0
                                            ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                            flag = True
                                            pageFlag = True
                                            val = a
                                            print('val={:.2f}'.format(val))
                                        #end if
                                    #end if
                                    
                                    st = t3.find(w1,st)+ len(w1)
                                #next
                            #end if
                        #end if
                    #next
                #end if
            if B_kind == "S造":
                # =======================================================
                #   S造の柱の検定表
                # ======================================================= 
                if len(CharLines) > 0:
                    # lines =t1.splitlines()
                    i = -1
                    kmode = False
                    # fword = "σc/fc"
                    fword1 = "σc/fc"
                    fword2 = "σ/fc"
                    fword3 = "検定比"
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        if not kmode :
                            # if fword in t3 : # 最初の「検定比」が現れたら「kmode」をTrue
                            if fword1 in t3 or fword2 in t3 or fword3 in t3 : # 最初のfwordが現れたら「kmode」をTrue
                                kmode = True
                                if fword1 in t3:
                                    fword = fword1
                                elif fword2 in t3:
                                    fword = fword2
                                elif fword3 in t3:
                                    fword = fword3
                                #end i
                                kmode = True
                                # fwordより右側にある数値だけを検出するためのX座標を取得
                                n = t3.index(fword)
                                c1 = CharData[i][n]
                                zx0 = c1[1]
                            #end if
                        else:
                            if kmode :
                                
                                CharLine = CharData[i] # １行文のデータを読み込む
                                t4 = ""
                            
                                for char in CharLine:
                                    # kmodeの時には「検定比」の下にある数値だけを検出する。
                                    if char[1]>=zx0 :
                                        t4 += char[0]
                                        
                                t4 = t4.replace(fword,"") 
                                if t4 == "": # 
                                    kmode = False
                                else:
                                    st = 0
                                    w0 = t4.split()
                                    if len(w0)>=1:
                                        for w1 in w0:
                                            w2 = w1.replace(" ","")
                                            if isfloat(w2): # 切り取った文字が数値の場合の処理
                                                a = float(w2)
                                                if a>=limit1 and a<=1.0:
                                                    # 数値がlimit以上の場合はデータに登録
                                                    n = t3.find(w2,st)   # 数値の文字位置を検索
                                                    xxx0 = CharLine[n][1]
                                                    xxx1 = CharLine[n+3][2]
                                                    yyy0 = CharLine[n][3]
                                                    yyy1 = CharLine[n][4]
                                                    xxx0 -= xd
                                                    xxx1 += xd
                                                    width3 = xxx1 - xxx0
                                                    height3 = yyy1 - yyy0
                                                    ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                                    flag = True
                                                    pageFlag = True
                                                    val = a
                                                    print('val={:.2f}'.format(val))
                                                #end if
                                            #end if
                                            
                                            st = t3.find(w1,st)+ len(w1)
                                        #next
                                    #end if
                                #end if
                            #end if
                        #end if
                    #next
                #end if
            #end if


        #=================================================================================================
        #   梁の検定表のチェック
        #=================================================================================================
                            
        elif mode == "梁の検定表" : 

            CharLines , CharData = self.MakeChar(page, interpreter2,device2)
            if B_kind == "RC造" or B_kind == "SRC造" or B_kind == "":
                # =======================================================
                #   RC造およびSRC造の梁の検定表
                # ======================================================= 
                
                if len(CharLines) > 0:
                
                    # lines =t1.splitlines()
                    i = -1
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        CharLine = CharData[i] # １行文のデータを読み込む
                        
                        if "検定比" in t3 : # 「検定比」が現れた場合の処理
                            # print(t3)
                            st = 0
                            t4 = t3.split()            # 文字列を空白で分割
                            if len(t4)>0:    # 文字列配列が１個以上ある場合に処理
                                for t5 in t4:
                                    t6 = t5.replace("検定比","")    # 「検定比」と数値が一緒の場合は除去
                                    nn = t3.find(t6,st)   # 数値の文字位置を検索
                                    ln = len(t5)
                                    if isfloat(t6):
                                        a = float(t6)
                                        if a>=limit1 and a<=1.0:
                                            # 数値がlimit以上の場合はデータに登録
                                            xxx0 = CharLine[nn][1]
                                            xxx1 = CharLine[nn+3][2]
                                            yyy0 = CharLine[nn][3]
                                            yyy1 = CharLine[nn][4]
                                            xxx0 -= xd
                                            xxx1 += xd
                                            width3 = xxx1 - xxx0
                                            height3 = yyy1 - yyy0
                                            ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                            flag = True
                                            pageFlag = True
                                            val = a
                                            print('val={:.2f}'.format(val))
                                        #end if
                                    #end if

                                    # 数値を検索を開始するを文字数分移動
                                    st = t3.find(t5,st)+ len(t5)
                                    # st += ln
                                #next
                            #end if
                        #end if
                    #next
                #end if
            if B_kind == "S造":
                # =======================================================
                #   S造の梁の検定表
                # ======================================================= 
                
                if len(CharLines) > 0:
                    # lines =t1.splitlines()
                    i = -1
                    kmode = False
                    fword1 = "σb/fb"
                    fword2 = "σ/fb"
                    fword3 = "検定比"
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        if not kmode :
                            if fword1 in t3 or fword2 in t3 or fword3 in t3 : # 最初のfwordが現れたら「kmode」をTrue
                                kmode = True
                                if fword1 in t3:
                                    fword = fword1
                                elif fword2 in t3:
                                    fword = fword2
                                elif fword3 in t3:
                                    fword = fword3
                                #end i
                                # fwordより右側にある数値だけを検出するためのX座標を取得
                                n = t3.index(fword)  # + len(fword)
                                c1 = CharData[i][n]
                                zx0 = c1[1]
                            #end if
                        if kmode :
                            CharLine = CharData[i] # １行文のデータを読み込む
                            t4 = ""
                        
                            for char in CharLine:
                                # kfwordより右側にある数値だけを検出する。
                                if char[1]>=zx0 :
                                    t4 += char[0]

                            t4 = t4.replace(fword,"") 
                            #next
                            if t4 == "": # 
                                kmode = False
                            else:
                                st = 0
                                w0 = t4.split()
                                if len(w0)>=1:
                                    for w1 in w0:
                                        w2 = w1.replace(" ","")
                                        if isfloat(w2): # 切り取った文字が数値の場合の処理
                                            a = float(w2)
                                            if a>=limit1 and a<=1.0:
                                                # 数値がlimit以上の場合はデータに登録
                                                n = t3.find(w2,st)   # 数値の文字位置を検索
                                                xxx0 = CharLine[n][1]
                                                xxx1 = CharLine[n+3][2]
                                                yyy0 = CharLine[n][3]
                                                yyy1 = CharLine[n][4]
                                                xxx0 -= xd
                                                xxx1 += xd
                                                width3 = xxx1 - xxx0
                                                height3 = yyy1 - yyy0
                                                ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                                flag = True
                                                pageFlag = True
                                                val = a
                                                print('val={:.2f}'.format(val))
                                            #end if
                                        #end if
                                        st = t3.find(w1,st)+ len(w1)
                                    #next
                                #end if
                            #end if
                        #end if
                    #next
                #end if
            #end if
                                
        #=================================================================================================
        #   耐力壁の検定表のチェック
        #=================================================================================================

        elif mode == "壁の検定表":
            outtext1 , CharData1 = self.MakeChar(page, interpreter2,device2)
            
            if len(outtext1) > 0:
                i = -1
                tn = len(outtext1)

                while True:
                    i += 1
                    if i > tn-1 : break

                    t3 = outtext1[i][0]
                    # print(t3)
                    CharLine = CharData1[i]
                    if "QDL" in t3:
                        nn = t3.find("QDL",0)   # 数値の文字位置を検索
                        xxx0 = CharLine[nn][1]
                        yyy1 = CharLine[nn][4]
                        t4 = t3[nn+3:].replace(" ","")
                        if isfloat(t4):
                            A1 = float(t4)
                        else:
                            A1 = 0.0
                        
                        i += 1
                        t3 = outtext1[i][0]
                        CharLine = CharData1[i]
                        
                        nn  = t3.find("QAL",0) 
                        yyy0 = CharLine[nn][3]

                        t4 = t3[nn+3:].replace(" ","")
                        nn2 = len(t3[nn:])
                        
                        xxx1 = CharLine[nn+nn2-1][2]
                        yyy0 = CharLine[nn+nn2-1][3]
                        
                        if isfloat(t4):
                            A2 = float(t4)
                        else:
                            A2 = 10000.0
                        QDL_mode = False
                        
                        if A2 != 0.0:
                            a = abs(A1/A2)
                            if a>=limit2 and a<=1.0:
                                
                                xxx0 -= xd
                                xxx1 += xd
                                width3 = xxx1 - xxx0
                                height3 = yyy1 - yyy0
                                points = []
                                points.append((xxx0,yyy0,xxx1,yyy0))
                                points.append((xxx1,yyy0,xxx1,yyy1))
                                points.append((xxx1,yyy1,xxx0,yyy1))
                                points.append((xxx0,yyy1,xxx0,yyy0))
                                ResultData.append([a,[xxx0, yyy0, width3, height3],True,points])
                                flag = True
                                pageFlag = True
                                val = a
                                print('val={:.2f}'.format(val))

                        i += 1
                        t3 = outtext1[i][0]
                        # print(t3)
                        CharLine = CharData1[i]

                        nn = t3.find("QDS",0)   # 数値の文字位置を検索
                        xxx0 = CharLine[nn][1]
                        yyy1 = CharLine[nn][4]
                        t4 = t3[nn+3:].replace(" ","")
                        if isfloat(t4):
                            A1 = float(t4)
                        else:
                            A1 = 0.0
                        QDL_mode = True
                            
                    
                        i += 1
                        t3 = outtext1[i][0]
                        CharLine = CharData1[i]
                        
                        nn = t3.find("QAS",0)
                        yyy0 = CharLine[nn][3]

                        t4 = t3[nn+3:].split()[0]
                        nn2 = len(t3[nn:])
                        
                        xxx1 = CharLine[nn+nn2-1][2]
                        yyy0 = CharLine[nn+nn2-1][3]
                        
                        if isfloat(t4):
                            A2 = float(t4)
                        else:
                            A2 = 10000.0
                        QDL_mode = False
                        
                        if A2 != 0.0:
                            a = abs(A1/A2)
                            if a>=limit2 and a<=1.0:
                                
                                xxx0 -= xd
                                xxx1 += xd
                                width3 = xxx1 - xxx0
                                height3 = yyy1 - yyy0
                                ResultData.append([a,[xxx0, yyy0, width3, height3],True])
                                flag = True
                                pageFlag = True
                                val = a
                                print('val={:.2f}'.format(val))
                            #end if
                        #end if
                    #end if
                #end while
            #end if

        elif mode == "杭の検定表":
            pageFlag = False


        #=================================================================================================
        #   ブレースの検定表のチェック
        #=================================================================================================
                        
        elif mode == "ブレースの検定表" : 

            CharLines , CharData = self.MakeChar(page, interpreter2,device2)
            
            if len(CharLines) > 0:
                    # lines =t1.splitlines()
                    i = -1
                    kmode = False
                    fword1 = "Nt/Nat"
                    fword2 = "σt/ft"
                    fword3 = "検定比"
                    for line in CharLines:
                        i += 1
                        t3 = line[0]
                        if not kmode :
                            if fword1 in t3 or fword2 in t3 or fword3 in t3 : # 最初のfwordが現れたら「kmode」をTrue
                                kmode = True
                                if fword1 in t3:
                                    fword = fword1
                                elif fword2 in t3:
                                    fword = fword2
                                elif fword3 in t3:
                                    fword = fword3
                                #end i
                                # fwordより右側にある数値だけを検出するためのX座標を取得
                                n = t3.index(fword)  # + len(fword)
                                c1 = CharData[i][n]
                                zx0 = c1[1]
                            #end if
                        if kmode :
                            CharLine = CharData[i] # １行文のデータを読み込む
                            t4 = ""
                        
                            for char in CharLine:
                                # kfwordより右側にある数値だけを検出する。
                                if char[1]>=zx0 :
                                    t4 += char[0]

                            t4 = t4.replace(fword,"") 
                            #next
                            if t4 == "": # 
                                kmode = False
                            else:
                                st = 0
                                w0 = t4.split()
                                if len(w0)>=1:
                                    for w1 in w0:
                                        w2 = w1.replace(" ","")
                                        if isfloat(w2): # 切り取った文字が数値の場合の処理
                                            a = float(w2)
                                            if a>=limit1 and a<=1.0:
                                                # 数値がlimit以上の場合はデータに登録
                                                n = t3.find(w2,st)   # 数値の文字位置を検索
                                                xxx0 = CharLine[n][1]
                                                xxx1 = CharLine[n+3][2]
                                                yyy0 = CharLine[n][3]
                                                yyy1 = CharLine[n][4]
                                                xxx0 -= xd
                                                xxx1 += xd
                                                width3 = xxx1 - xxx0
                                                height3 = yyy1 - yyy0
                                                ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                                flag = True
                                                pageFlag = True
                                                val = a
                                                print('val={:.2f}'.format(val))
                                            #end if
                                        #end if
                                        st = t3.find(w1,st)+ len(w1)
                                    #next
                                #end if
                            #end if
                        #end if
    
        #==========================================================================
        #  検出結果を出力する
        return pageFlag, ResultData
    #end def
    #*********************************************************************************


    def OtherSheet(self, page, limit, interpreter, device,interpreter2, device2):
        
        #============================================================
        # 構造計算書が不明の場合の処理
        #============================================================
        pageFlag = False
        ResultData = []
        limit1 = limit
        limit2 = limit
        limit3 = limit
        interpreter.process_page(page)
        layout = device.get_result()
        #
        #   このページに「断面検定表」、「検定比図」の
        #   文字が含まれている場合のみ数値の検索を行う。
        #
        

        検定比_Flag = False

        xd = 3      #  X座標の左右に加える余白のサイズ（ポイント）を設定

        mode = ""
        
        for lt in layout:
            # LTTextContainerの場合だけ標準出力　断面算定表(杭基礎) 仕口 継手 付着
            if isinstance(lt, LTTextContainer):
                texts = lt.get_text()
                texts = texts.replace("\n","")                
                if "柱の断面検定表"in texts or "梁の断面検定表"in texts or "ブレースの断面検定表"in texts or "壁の断面検定表"in texts  or "検定比図" in texts : 
                    検定比_Flag = True
                    break                   
            #end if
        #next

        if not 検定比_Flag  :     # 該当しない場合はこのページの処理は飛ばす。
            print("No Data")
            return False,[]
        else:
            print("")
        #end if

        #=================================================================================================
        #   検定比図のチェック
        #=================================================================================================
        
        if 検定比_Flag  :

            CharLines , CharData = self.MakeChar2(page, interpreter2,device2)

            if len(CharLines) > 0:
                i = -1
                for line in CharLines:
                    # print(line)
                    i += 1
                    t3 = line[0]
                    CharLine = CharData[i] # １行文のデータを読み込む
                    
                    # line = CharLines[i][0]
                    line2 = ""
                    xx= CharData[i][0][2]
                    for Char in CharData[i]:
                        if Char[1]>xx+3:
                            line2 += " "
                        line2 += Char[0]
                        xx = Char[2]
                    #next
                    items = line2.split()
                    # print(line)
                    # print(items)
                    a=0

                    # if "検定比" in t3 : # 「検定比」が現れた場合の処理
                    # print(t3)
                    st = 0
                    # t4 = t3.split()            # 文字列を空白で分割
                    t4 = items
                    if len(t4)>0 and t4[0] != "Super":    # 文字列配列が１個以上ある場合に処理
                        for t5 in t4:
                            t6 = t5.replace("(","").replace(")","").replace(" ","").replace("C","").replace("T","").replace("組","")     # 「検定比」と数値が一緒の場合は除去
                            nn = t3.find(t6,st)   # 数値の文字位置を検索
                            ln = len(t6)

                            # カッコがある場合は左右１文字ずつ追加
                            if "(" in t5:
                                xn1 = 1
                                xn2 = 1
                            elif "C" in t5 or "T" in t5 or "組" in t5:
                                xn1 = 0
                                xn2 = 1
                            else:
                                xn1 = 0
                                xn2 = 0

                            if isfloat(t6) and len(t6)>=4:
                                a = float(t6)
                                if a>=limit1 and a<=1.0:
                                    # 数値がlimit以上の場合はデータに登録
                                    xxx0 = CharLine[nn-xn1][1]
                                    xxx1 = CharLine[nn+ln+xn2-1][2]
                                    if CharLine[nn][5][1] > 0.0:
                                        yyy0 = CharLine[nn][3] - 1.0
                                        yyy1 = CharLine[nn+ln+xn1-1][4] + 1.0
                                    elif CharLine[nn][5][1] < 0.0:
                                        yyy0 = CharLine[nn+ln+xn1-1][3] - 2.0
                                        yyy1 = CharLine[nn][4] + 2.0
                                    else:
                                        yyy0 = CharLine[nn][3]
                                        yyy1 = CharLine[nn][4]
                                    #end if

                                    if ln <=4 :
                                        xxx0 -= xd
                                        xxx1 += xd
                                    #end if
                                    width3 = xxx1 - xxx0
                                    height3 = yyy1 - yyy0
                                    ResultData.append([a,[xxx0, yyy0, width3, height3],False])
                                    flag = True
                                    pageFlag = True
                                    val = a
                                    print('val={:.2f}'.format(val))
                                #end if
                            #end if


                            # 数値を検索を開始するを文字数分移動
                            st = nn + ln
                        #next
                    #end if
                #next
            #end if
        # #end if
        
        #==========================================================================
        #  検出結果を出力する
        return pageFlag, ResultData
    #end def
    #*********************************************************************************


    #============================================================================
    #  表紙のチェック（外部から読み出す関数名）
    #============================================================================

    def TopPageCheckTool(self,filename, outdir,limit=0.95 ):
        global flag1, fname, dir1, dir2, dir3, dir4, dir5, folderName, paraFileName
        global ErrorFlag, ErrorMessage
        global kind, verion

        if filename =="" :
            return False
        #end if

        pdf_file = filename
        pdf_out_file = outdir + "/outfile0000.pdf"

        # PyPDF2のツールを使用してPDFのページ情報を読み取る。
        # PDFのページ数と各ページの用紙サイズを取得
        try:
            with open(pdf_file, "rb") as input:
                reader = PR2(input)
                self.PageMax = len(reader.pages)     # PDFのページ数
                PaperSize = []
                for page in reader.pages:       # 各ページの用紙サイズの読取り
                    p_size = page.mediabox
                    page_xmin = float(page.mediabox.lower_left[0])
                    page_ymin = float(page.mediabox.lower_left[1])
                    page_xmax = float(page.mediabox.upper_right[0])
                    page_ymax = float(page.mediabox.upper_right[1])
                    PaperSize.append([page_xmax - page_xmin , page_ymax - page_ymin])
            #end with
        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, kind, version
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False, kind, version
        #end try
        
        #=============================================================
        startpage = 1
        endpage = self.PageMax
        
        # PDFMinerのツールの準備
        resourceManager = PDFResourceManager()
        # PDFから単語を取得するためのデバイス
        device = PDFPageAggregator(resourceManager, laparams=LAParams())
        # PDFから１文字ずつを取得するためのデバイス
        device2 = PDFPageAggregator(resourceManager)

        pageResultData = []
        pageNo = []

        try:
            with open(pdf_file, 'rb') as fp:
                interpreter = PDFPageInterpreter(resourceManager, device)
                interpreter2 = PDFPageInterpreter(resourceManager, device2)
                pageI = 0
                
                for page in PDFPage.get_pages(fp):
                    pageI += 1

                    ResultData = []
                    print("page={}:".format(pageI), end="")
                    if pageI == 1 :
                        # flag1 = True
                        pageFlag = True
                        kind, version = self.CoverCheck(page, interpreter2, device2)
                        print()
                        print("プログラムの名称：{}".format(kind))
                        print("プログラムのバージョン：{}".format(version))

                        with open("./kind.txt", 'w', encoding="utf-8") as fp2:
                            print(kind, file=fp2)
                            print(version, file=fp2)
                            fp2.close()
                        #end with
                        # break
                    if pageFlag : 
                        pageNo.append(pageI)
                        pageResultData.append(ResultData)
                        break
                    #end if
                #next

                fp.close()
                folderName = ""
            # end with

        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        #end try


        # 使用したデバイスをクローズ
        device.close()
        device2.close()

        #============================================================================================
        #
        #   数値検出結果を用いて各ページに四角形を描画する
        #
        #============================================================================================
        
        try:
            in_path = pdf_file
            out_path = pdf_out_file

            # 保存先PDFデータを作成
            cc = canvas.Canvas(out_path)
            cc.setLineWidth(1)
            # PDFを読み込む
            pdf = PdfReader(in_path, decompress=False)

            i = 0
            for pageI in range(len(pageNo)):
                pageN = pageNo[pageI]
                pageSizeX = float(PaperSize[pageN-1][0])
                pageSizeY = float(PaperSize[pageN-1][1])
                page = pdf.pages[pageN - 1]
                ResultData = pageResultData[pageI]
                # PDFデータへのページデータの展開
                pp = pagexobj(page) #ページデータをXobjへの変換
                rl_obj = makerl(cc, pp) # ReportLabオブジェクトへの変換  
                cc.doForm(rl_obj) # 展開
                cc.setFillColor("red")
                font_name = "ipaexg"
                cc.setFont(font_name, 20)
                cc.drawString(20 * mm,  pageSizeY - 40 * mm, "検定比（{}以上）の検索結果".format(limit))

                
                # ページデータの確定
                cc.showPage()
            # next

            # PDFの保存
            cc.save()

            # time.sleep(1.0)
            # # すべての処理がエラーなく終了したのでTrueを返す。
            # return True

            return kind, version 
        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return "",""
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return "",""

        #end try

        # すべての処理がエラーなく終了したのでTrueを返す。
        return True, kind, version

    #end def    
    #*********************************************************************************




    #============================================================================
    #  表紙以外のページのチェック（外部から読み出す関数名）
    #============================================================================

    def PageCheck(self,filename, outdir, limit ,kind, version, psn, PageNumber,ProcessN):
        global flag1, fname, dir1, dir2, dir3, dir4, dir5, folderName, paraFileName
        global ErrorFlag, ErrorMessage
        
        if filename =="" :
            return False
        #end if

        pdf_file = filename

        # PyPDF2のツールを使用してPDFのページ情報を読み取る。
        # PDFのページ数と各ページの用紙サイズを取得
        PaperSize = []
        self.PaperRotate = []                
        try:
            with open(pdf_file, "rb") as input:
                reader = PR2(input)
                PageMax = len(reader.pages)     # PDFのページ数
                # PaperSize = []
                # PaperRotate = []
                for page in reader.pages:       # 各ページの用紙サイズの読取り
                    p_size = page.mediabox
                    rotate = page.get('/Rotate', 0)
                    self.PaperRotate.append(rotate)
                    page_xmin = float(page.mediabox.lower_left[0])
                    page_ymin = float(page.mediabox.lower_left[1])
                    page_xmax = float(page.mediabox.upper_right[0])
                    page_ymax = float(page.mediabox.upper_right[1])
                    PaperSize.append([page_xmax - page_xmin , page_ymax - page_ymin])
            #end with
        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        #end try
        
        #=============================================================
        # startpage = 1
        # endpage = PageMax
        
        # PDFMinerのツールの準備
        resourceManager = PDFResourceManager()
        # PDFから単語を取得するためのデバイス
        device = PDFPageAggregator(resourceManager, laparams=LAParams())
        # PDFから１文字ずつを取得するためのデバイス
        device2 = PDFPageAggregator(resourceManager)

        pageResultData = []
        pageNo = []
        try:
            with open(pdf_file, 'rb') as fp:
                interpreter = PDFPageInterpreter(resourceManager, device)
                interpreter2 = PDFPageInterpreter(resourceManager, device2)

                PageData = []
                for page in PDFPage.get_pages(fp):
                    PageData.append(page)
                #next
                pageI = 0
                # pageI2 = 0
                # PageN2 = len(PageNumber)
                flagPage = True

                while flagPage:
                    for i,p in enumerate(PageNumber):
                        flagPage = False
                        if p > 0:
                            pageI = i + 1
                            PageNumber[i] = 0
                            flagPage = True
                            page = PageData[i]
                            ProcessN[psn] += 1
                            break
                        #end if
                    #next
                    if flagPage == False:
                        break
                    #end if

                    # outfile = outdir + "/" + "outfile{:0=4}.pdf".format(pageI)
                    ResultData = []
                    print("ps={}:page={}:".format(psn,pageI), end="")

                    if kind == "SuperBuild/SS7":
                        #============================================================
                        # 構造計算書がSS7の場合の処理
                        #============================================================

                        pageFlag, ResultData = self.SS7(page, limit, interpreter, device, interpreter2, device2)

                    # 他の種類の構造計算書を処理する場合はここに追加
                    # elif kind == "****":
                    #     pageFlag, ResultData = self.***(page, limit, interpreter, device, interpreter2, device2)

                    else:
                        #============================================================
                        # 構造計算書の種類が不明の場合はフォーマットを無視して数値のみを検出
                        #============================================================

                        pageFlag, ResultData = self.OtherSheet(page, limit, interpreter, device, interpreter2, device2)

                        # return False
                    #end if

                    if pageFlag : 
                        pageNo.append(pageI)
                        pageResultData.append(ResultData)
                    #end if
                #next

                fp.close()
                folderName = ""
                # os.remove("./kind.txt")
            # end with

        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        #end try


        # 使用したデバイスをクローズ
        device.close()
        device2.close()

        #============================================================================================
        #
        #   数値検出結果を用いて各ページに四角形を描画する
        #
        #============================================================================================
        
        try:
            if len(pageNo)>0:
                in_path = pdf_file
                # out_path = pdf_out_file

                # # 保存先PDFデータを作成
                # cc = canvas.Canvas(out_path)
                # cc.setLineWidth(1)
                # PDFを読み込む
                pdf = PdfReader(in_path, decompress=False)

                i = 0
                for pageI in range(len(pageNo)):
                    pageN = pageNo[pageI]

                    out_path = outdir + "/" + "outfile{:0=4}.pdf".format(pageN)

                    # 保存先PDFデータを作成
                    cc = canvas.Canvas(out_path)
                    cc.setLineWidth(1)
                    pageR = self.PaperRotate[pageN-1]
                    if pageR ==0:
                        pageSizeX = float(PaperSize[pageN-1][0])
                        pageSizeY = float(PaperSize[pageN-1][1])
                    else:
                        pageSizeX = float(PaperSize[pageN-1][1])
                        pageSizeY = float(PaperSize[pageN-1][0])

                    page = pdf.pages[pageN - 1]
                    ResultData = pageResultData[pageI]
                    # PDFデータへのページデータの展開
                    pp = pagexobj(page) #ページデータをXobjへの変換
                    rl_obj = makerl(cc, pp) # ReportLabオブジェクトへの変換  
                    cc.doForm(rl_obj) # 展開

                    pn = len(ResultData)

                    # ページの左肩に検出個数を印字
                    cc.setFillColor("red")
                    font_name = "ipaexg"
                    cc.setFont(font_name, 12)
                    t2 = "検索個数 = {}".format(pn)
                    cc.drawString(20 * mm,  pageSizeY - 15 * mm, t2)

                    # 該当する座標に四角形を描画
                    for R1 in ResultData:
                        a = R1[0]
                        origin = R1[1]
                        flag = R1[2]
                        x0 = origin[0]
                        y0 = origin[1]
                        width = origin[2]
                        height = origin[3]

                        # 長方形の描画
                        cc.setFillColor("white", 0.5)
                        cc.setStrokeColorRGB(1.0, 0, 0)
                        cc.rect(x0, y0, width, height, fill=0)

                        if flag:    # "壁の検定表"の場合は、四角形の右肩に数値を印字
                            cc.setFillColor("red")
                            font_name = "ipaexg"
                            cc.setFont(font_name, 7)
                            t2 = " {:.2f}".format(a)
                            cc.drawString(origin[0]+origin[2], origin[1]+origin[3], t2)
                        #end if
                    #next

                    # ページデータの確定
                    cc.showPage()
                    cc.save()
                # next

                # PDFの保存
                # cc.save()

            #end if

        except OSError as e:
            print(e)
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False
        except:
            logging.exception(sys.exc_info())#エラーをlog.txtに書き込む
            return False

        #end try

        # すべての処理がエラーなく終了したのでTrueを返す。
        return True

    #end def    
    #*********************************************************************************

    #============================================================================
    #  このクラスを単独で実行する場合に呼び出す関数
    #       filename    :   計算書のファイル名
    #       outfilename :   結果ファイル名
    #       limit       :   数値の閾値（0.01〜0.99）
    #       startpage   :   開始ページ
    #       endpage     :   終了ページ
    #
    #       CT = CheckTool()
    #       pdf_out_file = os.path.splitext(filename)[0] + '[検出結果(閾値={:.2f}'.format(limit)+')].pdf'
    #       flag = CT.doCheck(filename,pdf_out_file,limit,stpage,edpage)
    #============================================================================
    
    def doCheck(self, filenname, outfilename, limit, startpage, endpage):

        # プログラムのあるディレクトリー
        fld = os.getcwd()

        # 分割ファイルを一時保存するディレクトリー（無ければ作成）
        dir1 = fld + "/pdf"
        if not os.path.isdir(dir1):
            os.mkdir(dir1)
        #end if

        # 結果ファイルを一時保存するディレクトリー（無ければ作成）
        dir2 = fld + "/out"
        if not os.path.isdir(dir2):
            os.mkdir(dir2)
        #end if

#       表示の読取り        
        kind, verison = self.TopPageCheckTool(filename,dir2,limit)

        ProcessN = [0]
        PageNumber = list (range(1, self.PageMax + 1))

        for i, p in enumerate(PageNumber):
            if p < startpage or p > endpage:
                PageNumber[i] = 0
            #end if
        #next

        self.PageCheck(filename,dir2,limit,kind,version,0,PageNumber,ProcessN)

        # 結果フォルダーにあるファイル名の読取り
        files = glob.glob(os.path.join(dir2, "*.pdf"))
        # ファイルのソート
        files.sort()

        # 結果ファイルを順番に結合し、１つの結果ファイルを保存
        merger = pypdf.PdfMerger()
        file1 = dir1 + "/f1.pdf"
        for i,file in enumerate(files):
            if i == 0:
                merger.append(file)
            else:
                n = int(file.replace(".pdf","")[-4:])
                if self.PaperRotate[n-1] != 0:
                    f1 = pypdf.PdfReader(open(file , 'rb'))
                    f2 = pypdf.PdfWriter()
                    for page in f1.pages:
                        rotate = self.PaperRotatea[n-1]
                        if rotate != 0:
                            page.rotate(rotate)
                        f2.add_page(page)
                    with open(file1, 'wb') as f:
                        f2.write(f)
                    merger.append(file1)
                else:
                    merger.append(file)
                #end if
            #end if
            if os.path.exists(file1):
                os.remove(file1)
            #
        #next
        merger.write(outfilename)
        merger.close()

        # 結果ファイルを消去
        for file in files:
            os.remove(file)

        # # 結果ファイルを消去
        # for file in self.fnames:
        #     os.remove(file)

        if os.path.exists(file1):
            os.remove(file1)

        return True
    #end def
    #*********************************************************************************


#==================================================================================
#   このクラスを単独でテストする場合のメインルーチン（シングルプロセス）
#==================================================================================

if __name__ == '__main__':
    
    time_sta = time.time()  # 開始時刻の記録

    stpage = 170
    edpage = 200
    limit = 0.90
    filename = "サンプル計算書(1).pdf"
    # stpage = 1
    # edpage = 0
    # limit = 0.70
    # filename = "SS7構造計算書（抜粋）のコピー.pdf"

    stpage = 2
    edpage = 32
    limit = 0.70
    filename = "法円坂(計算書)【抜粋】.pdf"

    CT = CheckTool()
    
    pdf_out_file = os.path.splitext(filename)[0] + '[検出結果(閾値={:.2f}'.format(limit)+')].pdf'

    flag = CT.doCheck(filename,pdf_out_file,limit,stpage,edpage)

    print(flag)

    t1 = time.time() - time_sta
    print("time = {} sec".format(t1))

#*********************************************************************************
