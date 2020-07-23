#!/usr/bin/env python3
# -*- coding:utf-8 -*-
#
# FGO周回カウンタタグカウンタ
# https://twitter.com/fgophi
#
# Twitter から7から10日前までの #FGO周回カウンタ のデータを集めExcel出力する
# tweepy と xlsxwriter と selenium と tqdm は標準では入っていない

import time
import re
import os
import sys
import configparser
import unicodedata
import webbrowser
import urllib
import csv
import copy
import tweepy
import xlsxwriter
import argparse
import datetime
import selenium
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from operator import attrgetter
import codecs

progname = "FGO周回カウンタタグカウンタ"
version = "0.9.16"

settingfile = os.path.join(os.path.dirname(__file__), 'setting.ini')

#keyの取得
config = configparser.ConfigParser()
try:
    config.read(settingfile)
    section1 = "auth_info"
    CONSUMER_KEY = config.get(section1, "CONSUMER_KEY")
    CONSUMER_SECRET = config.get(section1, "CONSUMER_SECRET")
except configparser.NoSectionError:
    print("[エラー] 設定ファイルに不備があります。setting.ini を見直してください。")
    sys.exit()


#OAuthHandlerクラスのインスタンスを作成
auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
#auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)

MAXSERCH = 100 #一回の検索でサーチするツイート数(最大100)
MAXLOOP = 5 #一回プログラムを実行するごとに行う検索数

NG_NAME = "yu6572134 ReiwaGFX 147185" #複数記述する場合は間にスペースをいれること
NG_ID = "798730121062457344 939162209824813056 2217382994" #複数記述する場合は間にスペースをいれること

use_resume = False #前回の続きから取得するか
nofavorited_only = False

##last_id = -1

freequest = {}
syurenquest = {}
misc_quest_data_list = []
sozai = {}
sozai_betsumei = {}
quest = {}
noclass = False
                                       
def normalize_item(s):
    """
    アイテム名を正規化する
    クエスト情報を読みこむときと周回データを読み込むときに使用する
    """
    # スキル石
    s=s.replace("剣の輝石", "剣輝")
    s=s.replace("弓の輝石", "弓輝")
    s=s.replace("槍の輝石", "槍輝")
    s=s.replace("騎の輝石", "騎輝")
    s=s.replace("術の輝石", "術輝")
    s=s.replace("殺の輝石", "殺輝")
    s=s.replace("狂の輝石", "狂輝")
    s=s.replace("剣の魔石", "剣魔")
    s=s.replace("弓の魔石", "弓魔")
    s=s.replace("槍の魔石", "槍魔")
    s=s.replace("騎の魔石", "騎魔")
    s=s.replace("術の魔石", "術魔")
    s=s.replace("殺の魔石", "殺魔")
    s=s.replace("狂の魔石", "狂魔")
    s=s.replace("剣の秘石", "剣秘")
    s=s.replace("弓の秘石", "弓秘")
    s=s.replace("槍の秘石", "槍秘")
    s=s.replace("騎の秘石", "騎秘")
    s=s.replace("術の秘石", "術秘")
    s=s.replace("殺の秘石", "殺秘")
    s=s.replace("狂の秘石", "狂秘")
    s=s.replace("輝石", "輝")
    s=s.replace("魔石", "魔")
    s=s.replace("秘石", "秘")
               
    #種火
    s=s.replace("叡智の灯火", "灯火")
    s=s.replace("叡智の大火", "大火")
    s=s.replace("叡智の猛火", "猛火")
    s=s.replace("の猛火", "猛火")
    s=s.replace("星4種火", "猛火")
    s=s.replace("の業火", "業火")
    s=s.replace("叡智の業火", "業火")
    s=s.replace("星5種火", "業火")
    s = re.sub("剣灯$","剣灯火",s)
    s = re.sub("弓灯$","弓灯火",s)
    s = re.sub("槍灯$","槍灯火",s)
    s = re.sub("騎灯$","騎灯火",s)
    s = re.sub("術灯$","術灯火",s)
    s = re.sub("殺灯$","殺灯火",s)
    s = re.sub("狂灯$","狂灯火",s)
    s = re.sub("剣大$","剣大火",s)
    s = re.sub("弓大$","弓大火",s)
    s = re.sub("槍大$","槍大火",s)
    s = re.sub("騎大$","騎大火",s)
    s = re.sub("術大$","術大火",s)
    s = re.sub("殺大$","殺大火",s)
    s = re.sub("狂大$","狂大火",s)
    s = re.sub("剣猛$","剣猛火",s)
    s = re.sub("弓猛$","弓猛火",s)
    s = re.sub("槍猛$","槍猛火",s)
    s = re.sub("騎猛$","騎猛火",s)
    s = re.sub("術猛$","術猛火",s)
    s = re.sub("殺猛$","殺猛火",s)
    s = re.sub("狂猛$","狂猛火",s)

##    #銅素材
##    s=s.replace("英雄の証", "証")
##    s=s.replace("凶骨", "骨")
##    s=s.replace("竜の牙", "牙")
##    s=s.replace("虚影の塵", "塵")
##    s=s.replace("愚者の鎖", "鎖")
##    s=s.replace("万死の毒針", "毒針")
##    s = re.sub("^針","毒針",s)
##    s=s.replace("魔術髄液", "髄液")
##    s=s.replace("宵哭きの鉄杭", "鉄杭")
##    s = re.sub("^杭","鉄杭",s)
##    s=s.replace("励振火薬", "火薬")
##
##    #銀素材
##    s=s.replace("世界樹の種", "種")
##    s=s.replace("ゴーストランタン", "ランタン")
##    s=s.replace("八連双晶", "八連")
##    s=s.replace("蛇の宝玉", "宝玉")
##    s=s.replace("鳳凰の羽根", "羽根")
##    s=s.replace("無間の歯車","歯車")
##    s=s.replace("禁断の頁", "頁")
##    s=s.replace("ホムンクルスベビー", "ホム")
##    s=s.replace("ホムベビ", "ホム")
##    s=s.replace("隕蹄鉄", "蹄鉄")
##    s=s.replace("大騎士勲章", "勲章")
##    s=s.replace("追憶の貝殻", "貝殻")
##    s=s.replace("枯淡勾玉", "勾玉")
##    s=s.replace("永遠結氷", "結氷")
##    s=s.replace("巨人の指輪", "指輪")
##    s=s.replace("オーロラ鋼", "オーロラ")
##    s=s.replace("閑古鈴", "鈴")
##    s=s.replace("禍罪の矢尻", "矢尻")
##
##    #金素材
##    s=s.replace("混沌の爪", "爪")
##    s=s.replace("蛮神の心臓", "心臓")
##    s=s.replace("竜の逆鱗", "逆鱗")
##    s = re.sub("^根","精霊根",s)
##    s=s.replace("戦馬の幼角", "幼角")
##    s=s.replace("血の涙石", "涙石")
##    s=s.replace("血涙", "涙石")
##    s = re.sub("^脂$","黒獣脂",s)
##    s=s.replace("封魔のランプ", "ランプ")
##    s=s.replace("智慧のスカラベ", "スカラベ")
##    s=s.replace("原初の産毛", "産毛")
##    s = re.sub("^毛$","産毛",s)
##    s=s.replace("呪獣胆石", "胆石")
##    s=s.replace("奇奇神酒", "神酒")
##    s = re.sub("^酒$","神酒",s)
##    s=s.replace("暁光炉心", "炉心")
##    s=s.replace("九十九鏡", "鏡")
##    s=s.replace("真理の卵", "卵")

##    if s in sozai_betsumei.keys():
##        s = re.sub("^" + s + "$", sozai_betsumei[s], s)    
## 正規表現に対応するよう修正
    for pattern in sozai_betsumei.keys():
        if re.match(pattern, s):
            s = re.sub("^" + s + "$", sozai_betsumei[pattern], s)
            break
    
    #ピース
    s=s.replace("セイバーピース", "剣ピ")
    s=s.replace("アーチャーピース", "弓ピ")
    s=s.replace("ランサーピース", "槍ピ")
    s=s.replace("ライダーピース", "騎ピ")
    s=s.replace("キャスターピース", "術ピ")
    s=s.replace("アサシンピース", "殺ピ")
    s=s.replace("バーサーカーピース", "狂ピ")
##    s=s.replace("剣ピース", "剣ピ")
##    s=s.replace("弓ピース", "弓ピ")
##    s=s.replace("槍ピース", "槍ピ")
##    s=s.replace("騎ピース", "騎ピ")
##    s=s.replace("術ース", "術ピ")
##    s=s.replace("殺ピース", "殺ピ")
##    s=s.replace("狂ピース", "狂ピ")
    s=s.replace("ピース", "ピ")

    #モニュ
    s=s.replace("セイバーモニュメント", "剣モ")
    s=s.replace("アーチャーモニュメント", "弓モ")
    s=s.replace("ランサーモニュメント", "槍モ")
    s=s.replace("ライダーモニュメント", "騎モ")
    s=s.replace("キャスターモニュメント", "術モ")
    s=s.replace("アサシンモニュメント", "殺モ")
    s=s.replace("バーサーカーモニュメント", "狂モ")
##    s=s.replace("剣モニュメント", "剣モ")
##    s=s.replace("弓モニュメント", "弓モ")
##    s=s.replace("槍モニュメント", "槍モ")
##    s=s.replace("騎モニュメント", "騎モ")
##    s=s.replace("術モニュメント", "術モ")
##    s=s.replace("殺モニュメント", "殺モ")
##    s=s.replace("狂モニュメント", "狂モ")
    s=s.replace("モニュメント", "モ")
    s=s.replace("モニュ", "モ")

    return s

class TweetUser:
    def __init__(self, name, screen_name):
        self.name = name
        self.screen_name = name

class TweetStatus:
    def __init__(self):
        self.time = None
        self.user = TweetUser(None, None)
        self.id = None
        self.id_str = None
        self.text = None
        self.full_text = None
    
class SyukaiReport:
    """
    #FGO周回カウンタ でリポートされた周回を収めるクラス
    """
    def __init__(self, report):
        """
        
        """
        self.original = report
        self.memo = []
##        self.source = "https://twitter.com/" + status.user.screen_name + "/status/" + status.id_str
##        self.time = status.created_at + datetime.timedelta(hours=9)
##        self.name = status.user.name
##        self.screen_name = status.user.screen_name
                        
        self.make_data(report)

    def make_data(self, report):
        self.category = None
        if "ツイ消し" in self.memo:
            self.category = "Error"
            
        ##複数報告検知
        place_pattern = "【(?P<place>[\s\S]+?)】"
        if len(re.findall(place_pattern, report)) > 1:
            self.desc = None
            self.num = None
            self.category = "Error"
            self.memo.append("1ツイートでの複数報告")
            return

        pattern = "【(?P<place>[\s\S]+)】(?P<num>[\s\S]+?)周(?P<num_after>.*?)\n(?P<items>[\s\S]+?)#FGO周回カウンタ"
        if len(re.findall(pattern, report)) < 1:
            self.desc = None
            self.num = None
            self.category = "Error"
            self.memo.append("文法エラー")
            return
        
        m = re.search(pattern, report)
        self.desc = m.group()

        #周回数
        num_str = re.sub(pattern, r"\g<num>", m.group())
        num_str = num_str.replace(",", "") #カンマは除く
        if not num_str.isdecimal(): #数字じゃないとき
            self.desc = None
            self.num = None
            self.category = "Error"
            self.memo.append("文法エラー")
            return
        
        num_pattern = "(?P<num_pre>[\S\D]*?)(?P<num>\-?[0-9.]+)"
        m1 = re.search(num_pattern, num_str)
        if not m1:
            self.num = None
            self.memo.append("周回数異常")
        else:
            num = re.sub(num_pattern, r"\g<num>", m1.group())
                
            if not num.isdigit():
##                num = None            
##                self.category = "Error"
                self.memo.append("周回数が整数でない")
            elif int(num) < 1:
                num = None
                self.category = "Error"
                self.memo.append("周回数が1以下")
            else:
                num = int(num)
            self.num = num

        num_pre = re.sub(num_pattern, r"\g<num_pre>", m1.group()).strip()
        num_after = re.sub(pattern, r"\g<num_after>", m.group()).strip()
        if len(num_pre) != 0 or len(num_after) != 0:
            self.memo.append("周回数に情報付与")            
       
        #アイテム記述部分
        items = re.sub(pattern, r"\g<items>", m.group())
        self.__make_itemdic(items)

        #周回場所 標準化
        place = re.sub(pattern, r"\g<place>", m.group())
        self.__normalize_place(place)
##        self.place = place

##        if self.items == {}:
##            self.category = "Error"
        
    def __make_itemdic(self, s):
        """
        入力テキストからドロップアイテムとその数を抽出
        辞書に保存
        """
        # 辞書にいれる
        self.items = {}
        s = unicodedata.normalize("NFKC", s)
        s = s.replace(",", "") #4桁の場合 "," をいれる人がいるので
        s = s.replace("×", "x") #x をバツと書く人問題に対処
        s = s.replace("QP(x", "QP(+") #QPの表記ぶれを修正 #7
        # 1行1アイテムに分割
        # chr(8211) : enダッシュ
        # chr(8711) : Minus Sign
##        splitpattern = "(-|" + chr(8711) + "|" +chr(8722) + "|\n)"
        splitpattern = "[-\n"+ chr(8711) +chr(8722) +"]"
        itemlist = re.split(splitpattern, s)

        #アイテムがでてくるまで行を飛ばす
        error_flag = False
        for item in itemlist:
    ##        # なぜかドロップ率をいれてくる人がいるので、カッコを除く
            item =re.sub("\([^\(\)]*\)$", "", item.strip()).strip()

            if len(item) == 0: #空行は無視
                continue
            if item.endswith("NaN"):
                if item.replace("NaN", "") in self.items.keys():
                    if self.items[item.replace("NaN", "")] != "NaN":
                        self.category = "Error"
                        self.items = {}
                        break                 
                self.items[item.replace("NaN", "")] = "NaN"
                continue
            if item.endswith("?"):
                self.category = "Error"
                self.items = {}
                self.memo.append("アイテム数「?」")
                break

            # ボーナスを表記してくるのをコメント扱いにする「糸+2」など
            if re.search("\+\d$", item):
                continue

            pattern = "(?P<name>^.*\D)(?P<num>\d+$)"
            m = re.search(pattern, item)
            if not m: #パターンに一致しない場合コメント扱いにする
                if item.isdigit(): #数字のみの記述
                    error_flag = True
                    if "数字のみの記述" not in self.memo:
                        self.memo.append("数字のみの記述")
                    continue
                if normalize_item(item) in sozai.keys():
                    error_flag = True
                    if "アイテム名のみの記述" not in self.memo:
                        self.memo.append("アイテム名のみの記述")
                    continue
                continue
            # 括弧つきのものは分離する 変換は括弧外のみで
            pattern_kakko = "(?P<name_k>^.*\D)(?P<kakko>\(.+?\)$)"
            item_k = re.sub(pattern, r"\g<name>", item).strip()
            m_kakko = re.search(pattern_kakko, item_k)
            if not m_kakko:
                tmpitem = normalize_item(item_k)
            else:
                tmpitem1 = normalize_item(re.sub(pattern_kakko, r"\g<name_k>", item_k).strip())
                tmpitem2 = re.sub(pattern_kakko, r"\g<kakko>", item_k)
                # 以下、QPの表記ぶれを修正 #7
                tmpitem2 = tmpitem2.replace("百", "00") 
                tmpitem2 = tmpitem2.replace("千", "000")
                tmpitem2 = tmpitem2.replace("k", "000")                 
                tmpitem2 = tmpitem2.replace("万", "0000")
                tmpitem = tmpitem1 + tmpitem2                
            
            #業火は例外にする
            if noclass == True:
                exlist = []
            else:
                exlist = ["モ", "ピ", "輝", "魔", "秘", "大火", "猛火"]

##            exlist = ["モ", "ピ", "輝", "魔", "秘", "大火", "猛火", "業火"]
##            error_flag = False
            for ex in exlist:
                if ex == tmpitem:
                    error_flag = True
                    self.memo.append("クラス指定無しアイテム")
                    continue
            exlist2 = ["剣", "弓", "槍", "騎", "術", "殺", "狂"]
            for ex in exlist2:
                if ex == tmpitem:
                    error_flag = True
                    self.memo.append("クラスのみ記述のアイテム")
                    continue
            if tmpitem.endswith("種火"):
                error_flag = True
                self.memo.append("『種火』報告")
                continue
##            if error_flag:
##                self.items = {}
##                self.category = "Error"
##                break
            if " " in tmpitem:
                self.items = {}
                self.category = "Error"
                self.memo.append("アイテム名中の空白")
                continue
            else:
                num = int(re.sub(pattern, r"\g<num>", item))
                if tmpitem in self.items.keys():
                    if self.items[tmpitem] != num:
                        self.category = "Error"
                        self.items = {}
                        self.memo.append("同名アイテムの重複")
                        continue                        
                self.items[tmpitem] = num
        #アイテム数0のとき
        if self.category != "Error" and len(self.items.keys()) == 0:
            self.category = "Error"
            self.memo.append("アイテム数0判定")
        if error_flag == True:
            self.items = {}
            self.category = "Error"
            
    def __normalize_place(self, place):
        if self.category == "Error":
            self.place = place
            return
        place_pattern1 = "(\n|　)" #改行と全角スペースは半角スペースにいったん変換
        place_pattern2 = " +" #半角スペースの連続は半角スペースに変換
        place = re.sub(place_pattern1, " ", place)
        place = re.sub(place_pattern2, " ", place)
        # 【】で囲まれた場合【】を除去
        if place.startswith("【") and place.endswith("】"):
            place = place[1:-1]

        if place in quest.keys():
            place = re.sub("^" + place + "$", quest[place], place)
        self.place = place
            
        tmpplace = place.replace('(', ' ')
        tmpplace = tmpplace.replace(')', ' ')
        tmp = tmpplace.split()
        if len(tmp) != 2:
            flag = False
            for pl in reversed(tmp):
                if pl in freequest.keys():
                    tmp = [freequest[pl]["場所"], pl]
                    flag = True
                    self.place = tmp[0] + " " +tmp[1]
                    break
                for p in freequest.keys():
                    if ("場所", pl) in freequest[p].items():  #フリクエデータとして認識
                        tmp = [pl, p]
                        flag = True
                        self.place = tmp[0] + " " +tmp[1]
                        break
            if flag == False:                    
                self.category = "その他クエスト"
                return

        place1 = tmp[0]
        place2 = tmp[1]
        if place2 in freequest.keys(): #フリクエデータとして認識 一つの場所にフリクエ二つある場合
            if place2 == "不夜城" and place1 == "アガルタ":
                place2 = "眠らぬ街"
            #アイテムチェック
            if self.__validitem(self.items.keys(), freequest[place2]["ドロップアイテム"].keys()) == False:
                self.category = "Error"
                return
            #ドロップ率チェック
            if self.__dropcheck(self.num, self.items) == False:
                self.category = "Error"
                return
            if freequest[place2]["ストーリー"] == "フリクエ1部":
                self.category = "フリクエ1部"
            elif freequest[place2]["ストーリー"] == "フリクエ1.5部":
                self.category = "フリクエ1.5部"
            else: #フリクエ2部
                self.category = "フリクエ2部"
            self.__make_freequest_data(place2)
        elif place in syurenquest.keys(): #修練クエストとして認識
            #アイテムチェック
            if self.__validitem(self.items.keys(), syurenquest[place]["ドロップアイテム"].keys()) == False:
                self.category = "Error"
                self.memo.append("非存在アイテム")
                return
            #ドロップ率チェック
            if self.__dropcheck(self.num, self.items)  == False:
                self.category = "Error"
                self.memo.append("ドロップ率異常")
                return
            self.category = "修練場"
            self.__make_syurenquest_data()
        else:
            check = 0
            for p in freequest.keys():
                if ("場所", place2) in freequest[p].items():  #フリクエデータとして認識
                    #アイテムチェック
                    if self.__validitem(self.items.keys(), freequest[p]["ドロップアイテム"].keys()) == False:
                        self.category = "Error"
                        self.memo.append("非存在アイテム")
                        return
                    #ドロップ率チェック
                    if self.__dropcheck(self.num, self.items) == False:
                        self.category = "Error"
                        self.memo.append("ドロップ率異常")
                        return                        

                    if freequest[p]["ストーリー"] == "フリクエ1部":
                        self.category = "フリクエ1部"
                    elif freequest[p]["ストーリー"] == "フリクエ1.5部":
                        self.category = "フリクエ1.5部"
                    else:
                        self.category = "フリクエ2部"
                    check = 1
                    self.__make_freequest_data(p)
                    break
            if check == 0:
                self.category = "その他クエスト"
                
    def __validitem(self, reportkey, fqkey):
        """
        常設フリクエ・修練場で変なアイテムが入ってないかチェック
        """

        tmpset = set(fqkey)
        tmpset.add("銀種火")
        tmpset.add("剣大火")
        tmpset.add("弓大火")
        tmpset.add("槍大火")
        tmpset.add("騎大火")
        tmpset.add("術大火")
        tmpset.add("殺大火")
        tmpset.add("狂大火")
        tmpset.add("剣灯火")
        tmpset.add("弓灯火")
        tmpset.add("槍灯火")
        tmpset.add("騎灯火")
        tmpset.add("術灯火")
        tmpset.add("殺灯火")
        tmpset.add("狂灯火")
        
        if len(set(reportkey) -tmpset) > 0:
            return False

        return True

    def __dropcheck(self, num, item_dict):
        """
        金素材：
        周回数に関わらず泥率100%より大きい報告はエラー。
        周回数100周以上の場合は泥率50%以上でエラー。

        銀素材：
        周回数に関わらず泥率200%以上でエラー。
        周回数100周以上の場合は泥率70%以上でエラー。

        銅素材：
        周回数に関わらず泥率300%以上でエラー。
        周回数100周以上の場合は泥率90%以上でエラー
        """
        if type(num) != int:
            if not num.isdigit():
                return True
        if num < 1:
            return True            
            
        for item in item_dict.keys():
            if item_dict[item] == "NaN":
                return True
            if item in sozai.keys():
                if sozai[item] == "金":
                    if item_dict[item] > num:
                        return False
                    if num >= 100:
                        if item_dict[item] > num * 0.5:
                            return False
                elif sozai[item] == "銀":
                    if item_dict[item] > num * 2:
                        return False
                    if num >= 100:
                        if item_dict[item] > num * 0.7:
                            return False            
                elif sozai[item] == "銅":
                    if item_dict[item] > num * 3:
                        return False
                    if num >= 100:
                        if item_dict[item] > num * 0.9:
                            return False
                else:
                    return False

        return True
    
    def __make_freequest_data(self, place):
        """
        各ツイートからフリクエのデータを作成
        """

##        freequest[place]["周回数"].append(self.num)
##        freequest[place]["id"].append(self.id)
##        freequest[place]["screen_name"].append(self.screen_name)
##        freequest[place]["メモ"].append(self.time)
##        for item in freequest[place]["ドロップアイテム"]:
##            if item in self.items.keys():
##                freequest[place]["ドロップアイテム"][item].append(self.items[item])
##            else:
##                freequest[place]["ドロップアイテム"][item].append(-1)
        report = {}
        report["周回数"] = self.num
        report["id"] = self.id
        report["screen_name"] = self.screen_name
        report["time"] = self.time
        report["items"] = self.items
        freequest[place]["report"].append(report)

    def __make_syurenquest_data(self):
        """
        各ツイートから修練場のデータを作成
        """
        syurenquest[self.place]["周回数"].append(self.num)
##        syurenquest[self.place]["id"].append(self.id)
##        syurenquest[self.place]["screen_name"].append(self.screen_name)
##        syurenquest[self.place]["メモ"].append(self.time)
##        for item in syurenquest[self.place]["ドロップアイテム"]:
##            if item in self.items.keys():
##                syurenquest[self.place]["ドロップアイテム"][item].append(self.items[item])
##            else:
##                syurenquest[self.place]["ドロップアイテム"][item].append(-1)
        report = {}
        report["周回数"] = self.num
        report["id"] = self.id
        report["screen_name"] = self.screen_name
        report["time"] = self.time
        report["items"] = self.items
        syurenquest[self.place]["report"].append(report)

    def __make_misc_quest_data(self):
        """
        各ツイートからフリクエにも修練場にも該当しないデータを作成
        """

        #create date	name	user name	url	場所	周回数	Data
        tmplist = []
        tmplist.append(self.time)
        tmplist.append(self.name)
        tmplist.append(self.screen_name)
        tmplist.append(self.source)
        tmplist.append(self.place)
        tmplist.append(self.num)
        for item in self.items.keys():
            tmplist.append(item)
            tmplist.append(item_dict[item])
        misc_quest_data_list.append(tmplist)

class ReportTweet(SyukaiReport):
    """
    Twitterから直接取得したものを扱う
    """
    def __init__(self,status):
        """
        
        """
        self.source = "https://twitter.com/" + status.user.screen_name + "/status/" + status.id_str
        self.time = status.created_at + datetime.timedelta(hours=9)
        self.name = status.user.name
        self.screen_name = status.user.screen_name
        self.id = status.id
        self.reply_count = None
        self.correction = False
        self.full_text = status.full_text
        super().__init__(status.full_text)    

class DeletedTweet(ReportTweet):
    """
    削除されたツイートを扱う
    """
    def __init__(self, status):
        """
        
        """
        super().__init__(status)    
        self.original = status.full_text
        self.memo = ["ツイ消し"]                        
##        self.make_data(status.full_text)

class PrivateTweet(ReportTweet):
    """
    鍵垢化されたツイートを扱う
    """
    def __init__(self, status):
        """
        
        """
        super().__init__(status)    
        self.original = status.full_text
        self.memo = ["鍵垢化"]                        
##        self.make_data(status.full_text)

class YahooTweets:
    """
    Yahoo!リアルタイム検索から取得したものを扱う
    """
    def __init__(self, reports, history, since_id, sleep_time, use_yahoo):
        """
        
        """
        self.reports = []
        self.unsearch_reports = []      
        
        if len(reports) != 0:
            if use_yahoo:
                old_time = id2time(since_id)                    
                #データを取得
                self.__get_yahoo_reports(old_time, sleep_time)
                #差分データを作成        
                self.__make_diff(reports, history)
                #データを結合
                self.__conbine(reports)
            else:
                self.reports = reports
        
    def __get_yahoo_reports(self, old_time, sleep_time):
        """
        Yahoo! リアルタイム検索からデータを取得し、疑似Tweetを作成
        """
        global time #timeがローカル変数と認識されるので必要
        
        URL=r"""https://search.yahoo.co.jp/realtime/"""
        chrome_option = webdriver.ChromeOptions()
    ##    chrome_option.add_argument('--headless')
    ##    chrome_option.add_argument('--disable-gpu')
        chrome_option.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_option.add_argument('--disable-features=NetworkService')
        driver = webdriver.Chrome(options=chrome_option, service_args=["--silent"])

        driver.get(URL)
        time.sleep(sleep_time)

        ### Yahoo!リアルタイム検索トップページから入らないと自動取得が働かない
        elem = driver.find_element_by_name("p")
        elem.clear()
        elem.send_keys("#FGO周回カウンタ")
        elem.send_keys(Keys.RETURN)

        html01=driver.page_source
        flag = False
        now = datetime.datetime.now()
        unix_now = int(time.mktime(now.timetuple()))
        unix_old = int(time.mktime(old_time.timetuple()))
        tmp_time = unix_now

        print("Yahoo!リアルタイム検索での周回報告をブラウザ上に展開中")
        while 1:
            pbar = tqdm(total=unix_now - unix_old)
            while 1:
            ##   all_time =  time.mktime((now - old_time).timetuple())


                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                item = driver.find_elements_by_css_selector(".cnt.cf:not(.TS2bh)")
                get_time = int(item[-1].get_attribute("data-time"))
                difftime = tmp_time - get_time
                tmp_time = get_time
                pbar.update(int(difftime))

#                if(datetime.datetime.now() - datetime.datetime.fromtimestamp(int(item[-1].get_attribute("data-time")))).days >= 10:
                if old_time > datetime.datetime.fromtimestamp(int(item[-1].get_attribute("data-time"))):
                    #print("Twitter APIで取得したデータより古いデータを取得したので終了します")
                    flag = True
                    pbar.close()
                    break
                #print(str(sleep_time) + "秒待ちます")
                time.sleep(sleep_time)
                html02=driver.page_source
                if html01!=html02: #サイズが違うとき
                    html01=html02
                    #print("新たなページの取得に成功しました")
                    try:
                        driver.find_element_by_link_text("もっと見る").click()
                        get_time = int(item[-1].get_attribute("data-time"))
                        difftime = tmp_time - get_time
                        tmp_time = get_time
                        pbar.update(int(difftime))
                         #print("「もっと見る」ボタンを押しました")
                    except selenium.common.exceptions.NoSuchElementException:
                        continue
                else:
                    break
            if flag == True:
                break
            print("「もっと見る」ボタンの表示待ちです")
            wait = WebDriverWait(driver, 5) # 最大5秒
            elem = wait.until( expected_conditions.element_to_be_clickable( (By.LINK_TEXT,"もっと見る")) )
            elem.click()

        pattern = "\n.+@(?P<name>.+)\n"
        #for s in  driver.find_elements_by_css_selector(".cnt.cf:not(.TS2bh)"):
        texts = driver.find_elements_by_css_selector(".cnt.cf")
        names = driver.find_elements_by_css_selector(".refname")
        screen_names = driver.find_elements_by_css_selector(".nam")


        print("ブラウザ上から周回データを取得中")

        pbar = tqdm(total=len(texts))
        self.yahooreports = []
        for text, name, screen_name in zip(texts, names, screen_names):
            status=TweetStatus()
            time = datetime.datetime.fromtimestamp(int(text.get_attribute("data-time")))
##            std_time = time - datetime.timedelta(hours=9)
            screen_name = screen_name.text.replace("@", "")
##            id = self.__get_id(screen_name, time)
            
            status.time = time
##            status["user"] = {}
            status.user.name = name
            status.user.screen_name = screen_name
##            status["id"] = int(id)
##            status["id_str"] = id
            status.text= text.text
##            status["full_text"] = self.__get_tweet_text(id)
            if time >= old_time:
                self.yahooreports.append(status)
            pbar.update(1)
        pbar.close()
        driver.close()
        driver.quit()

    def __report2report(self, report):
        """
        __make_diff 内で履歴とYahoo!リアルタイム検索の報告の内容を比較するために
        比較用形式に変換するためのもの
        スペースと改行を取り除く
        """
        pattern = "【(?P<text>[\s\S]+)#FGO周回カウンタ"
        m = re.search(pattern, report)
        if m:
            new_report = re.sub(pattern, r"\g<text>", m.group())
            new_report = re.sub(" ", "", new_report)
            new_report = re.sub("\n", "", new_report)
        else:
            new_report = ""

        return new_report

    def __make_diff(self, twitter_reports, history):
        """
        Yahooにしかないツイートを見つけ、full_text id データを加える
        """
##        output_dic = {}
        new_reports = []
        row = 1
        i = 0
        yahoo_only = 0
        not_in_history = 0
        end_time =twitter_reports[-1].time
        print("Yahoo!リアルタイム検索のデータをTwitter APIからのデータと比較中")
        pbar = tqdm(total=len(self.yahooreports))
        for yahoo_report in self.yahooreports:
            yahoo_text = yahoo_report.text
            if yahoo_text.startswith("RT @"):
                pbar.update(1)
                continue            
            yahoo_screen_name = yahoo_report.user.screen_name
            yahoo_time = yahoo_report.time
            #Twitter APIから取得した投稿に同じ時間の投稿があるか調査
            exists_flag = False
            for twitter_report in twitter_reports: #Twitterのデータと比較
                twitter_time = twitter_report.time
                twitter_screen_name = twitter_report.screen_name
                if twitter_time == yahoo_time:
                    #同じ時間(分)に連投された報告の場合この処理はうまくいかない
                    if twitter_screen_name == yahoo_screen_name:
                        exists_flag = True
                        break
            for h in history.keys(): #履歴と比較
                history_time = history[h]["time"]
                history_screen_name = history[h]["screen_name"]
                if history_time == yahoo_time:
                    #同じ時間(分)に連投された報告の場合この処理はうまくいかない
                    if history_screen_name == yahoo_screen_name:
                        exists_flag = True
                        break
                    # screen_name が変更になっているときtextで判断
                    elif self.__report2report(yahoo_text) == self.__report2report(history[h]["text"]):
                        exists_flag = True
                        break
            if exists_flag == False: #同じ時間の投稿がないとき
                try:
                    status = self.__get_status(yahoo_report.user.screen_name, yahoo_report.time)
                    if status != None: # ツイ消し以外
                        yahoo_report = ReportTweet(status)
    ##                    yahoo_report.full_text = self.__get_tweet_text(yahoo_report.id)
                        if "#FGO販売" not in yahoo_report.full_text:
                            new_reports.append(yahoo_report)
                            yahoo_only = yahoo_only + 1
                except tweepy.error.TweepError: #Not authorized.
                ## historyにもない鍵垢化データがYahooに出ていた場合はどうしようもないので無視
                    pbar.update(1)
                    continue
            pbar.update(1)
            i = i + 1
        pbar.close()
        print("履歴に無いYahoo!のみのデータ: ", end ="")
        print(yahoo_only, end = "件")
        print()
        
        self.unsearch_reports = new_reports        
             
    def __conbine(self, twitter_reports):
        new_list = self.unsearch_reports + twitter_reports
        new_list = sorted(new_list, key=attrgetter("time"), reverse=True)
        
        self.reports = new_list


##    def __get_id(self, user, tweet_time):
    def __get_status(self, user, tweet_time):
        """
        ユーザー名と投稿時間からツイートstatusを取得する
        """
        # OAuth認証
        api = tweepy.API(auth)
        
        max_id_option = -1
        max_loop = MAXLOOP
        status = None

        reports = []
        status = None
        flag = False
        delteted_flag = False
        for loop in range(max_loop):
            ### api.favorites は api.search と異なり、max_id に -1 をいれると
            ### 正常に動かないので　-1 のときの処理を作る必要がある

            if max_id_option == -1:
                for status in api.user_timeline(id=user, tweet_mode="extended",
                                                count=200):
                    if (status.created_at + datetime.timedelta(hours=9)) == tweet_time:
#                    if (status.created_at) == tweet_time:
##                        id = status.id
                        flag = True
                        break
                    ## 時間がすぎるまで検索
                    if (status.created_at + datetime.timedelta(hours=9)) < tweet_time:
#                    if (status.created_at) < tweet_time:
                        flag = True
                        delteted_flag = True #見つからなかった
                        break
            else:
                for status in api.user_timeline(id=user,
                                                max_id=max_id_option, tweet_mode="extended",
                                                count=200):
                    if (status.created_at + datetime.timedelta(hours=9)) == tweet_time:
#                    if (status.created_at) == tweet_time:
##                        id = status.id
                        flag = True
                        break
                    ## 時間がすぎるまで検索
                    if (status.created_at + datetime.timedelta(hours=9)) < tweet_time:
#                    if (status.created_a) < tweet_time:
                        flag = True
                        delteted_flag = True #見つからなかった
                        break
            if flag == True:
                break
                
            if status != None:
                max_id_option=status.id -1
        if delteted_flag == True:
            return None
        return status

    def __get_tweet_text(self, id):
        """
        ユーザー名と投稿時間からツイートidを取得する
        """
        # OAuth認証
        api = tweepy.API(auth)

        status = api.get_status(id, tweet_mode=extended)

        return status.text

def id2time(id):
    """
    idから投稿時間を取得する
    """
    # OAuth認証
    api = tweepy.API(auth)
    
    status = api.get_status(int(id))
        
    return status.created_at + datetime.timedelta(hours=9)

def id2screen_name(id):
    """
    投稿idから投稿者を取得する
    """
    # OAuth認証
    api = tweepy.API(auth)

    status = api.get_status(id)

    return status.user.screen_name

def userid2screen_name(id):
    """
    useridからscreen_nameを取得する
    """
    # OAuth認証
    api = tweepy.API(auth)
    
    user = api.get_user(id)

    return user.screen_name

class ExcelFile:
    def __init__(self, filename):
        story_list = ["フリクエ1部", "フリクエ1.5部", "フリクエ2部"]

        if filename.endswith(".xlsx") == True:
            filename = filename
        else:
            filename = filename + ".xlsx"
        self.wb = xlsxwriter.Workbook(filename)
        self.ws_all = self.wb.add_worksheet("全データ")
        self.ws_syuren = self.wb.add_worksheet("修練場")
        self.ws_fq1 = self.wb.add_worksheet("フリクエ1部")
        self.ws_fq15 = self.wb.add_worksheet("フリクエ1.5部")
        self.ws_fq2 = self.wb.add_worksheet("フリクエ2部")
        self.ws_misc = self.wb.add_worksheet("その他クエスト")
        self.ws_error = self.wb.add_worksheet("Error")
        self.ws_syuren_stat = self.wb.add_worksheet("統計【修練場】")
        for story in story_list:
            ws = self.wb.add_worksheet("統計【" + story + "】")

        self.make_header()

    def close(self):
        self.wb.close()

    def make_sheets(self, reports, history, resume_id, favlist):
        self.__make_all_sheets(reports, history, resume_id, favlist)
        self.__make_fq1_sheets(reports, history, resume_id, favlist)
        self.__make_fq15_sheets(reports, history, resume_id, favlist)
        self.__make_fq2_sheets(reports, history, resume_id, favlist)
        self.__make_syuren_sheets(reports, history, resume_id, favlist)
        self.__make_misc_sheets(reports, history, resume_id, favlist)
        self.__make_error_sheets(reports, history, resume_id, favlist)

    def make_header(self):
        """
        Excel のヘッダ行を作成
        """
        for ws in [self.ws_all, self.ws_fq1, self.ws_fq15, self.ws_fq2, self.ws_syuren, self.ws_misc, self.ws_error]:        
            ws.write(0, 0, "create date")
            ws.write(0, 1, "name")
            ws.write(0, 2, "screen_name")
            ws.write(0, 3, "reply")
            ws.write(0, 4, "memo")
            ws.write(0, 5, "url")
            
        for ws in [self.ws_all, self.ws_error]:        
            ws.write(0, 6, "tweet")
        
        for ws in [self.ws_fq1, self.ws_fq15, self.ws_fq2, self.ws_syuren, self.ws_misc]:        
            ws.write(0, 6, "場所")
            ws.write(0, 7, "周回数")
            ws.write(0, 8, "Data")

        
    def __make_all_sheets(self, reports, history, resume_id, favlist):
        """
        「全データ」シートをつくる
        初出のデータとresume_idより新しいデータを出力
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo or "ツイ消し" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= resume_id and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            self.write_col_header(row, self.ws_all, report)
            self.ws_all.write(row, 6, report.original)
            if report.reply_count != None:
                self.ws_all.write(row, 3, report.reply_count)
                if report.correction == True:
                    self.ws_all.set_row(row, None, format1)
            if "ツイ消し" in report.memo:
                self.ws_all.set_row(row, None, format1)                    
            row = row + 1

    def __make_fq1_sheets(self, reports, history, resume_id, favlist):
        """
        "フリクエ1部"シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= resume_id  and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "フリクエ1部":  
                self.write_col_header(row, self.ws_fq1, report)
                self.__write_quest_data(row, self.ws_fq1, report)
                if report.reply_count != None:
                    self.ws_fq1.write(row, 3, report.reply_count)
                    if report.correction == True:
                        self.ws_fq1.set_row(row, None, format1)
##                self.ws_fq1.write(row, 5, report.original)
                row = row + 1

    def __make_fq15_sheets(self, reports, history, resume_id, favlist):
        """
        "フリクエ1.5部"シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "フリクエ1.5部":  
                self.write_col_header(row, self.ws_fq15, report)
                self.__write_quest_data(row, self.ws_fq15, report)
                if report.reply_count != None:
                    self.ws_fq15.write(row, 3, report.reply_count)
                    if report.correction == True:
                        self.ws_fq15.set_row(row, None, format1)
##                self.ws_fq15.write(row, 5, report.original)
                row = row + 1

    def __make_fq2_sheets(self, reports, history, resume_id, favlist):
        """
        "フリクエ2部"シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "フリクエ2部":  
                self.write_col_header(row, self.ws_fq2, report)
                self.__write_quest_data(row, self.ws_fq2, report)
                if report.reply_count != None:
                    self.ws_fq2.write(row, 3, report.reply_count)
                    if report.correction == True:
                        self.ws_fq2.set_row(row, None, format1)
##                self.ws_fq2.write(row, 5, report.original)
                row = row + 1

    def __make_syuren_sheets(self, reports, history, resume_id, favlist):
        """
        「修練場」シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "修練場":  
                self.write_col_header(row, self.ws_syuren, report)
                self.__write_quest_data(row, self.ws_syuren, report)
                if report.reply_count != None:
                    self.ws_syuren.write(row, 3, report.reply_count)
                    if report.correction == True:
                        self.ws_syuren.set_row(row, None, format1)
##                self.ws_syuren.write(row, 5, report.original)
                row = row + 1

    def __make_misc_sheets(self, reports, history, resume_id, favlist):
        """
        「その他クエスト」シートをつくる
        """
        row = 1
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "その他クエスト":

                self.write_col_header(row, self.ws_misc, report)
                self.__write_quest_data(row, self.ws_misc, report)
                if report.reply_count != None:
                    self.ws_misc.write(row, 3, report.reply_count)
                    if report.correction == True:
                        self.ws_misc.set_row(row, None, format1)
#                self.ws_misc.write(row, 5, report.original, format1)
                row = row + 1

    def __make_error_sheets(self, reports, history, resume_id, favlist):
        """
        「エラー」シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo or "ツイ消し" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            if report.category == "Error":  
                self.write_col_header(row, self.ws_error, report)
                if report.reply_count != None:
                    self.ws_error.write(row, 3, report.reply_count)
                self.ws_error.write(row, 6, report.original)
                if "ツイ消し" in report.memo:
                    self.ws_error.set_row(row, None, format1)                    
                row = row + 1

    def write_col_header(self, row, ws, report):
        """
        Excel の先頭行に各ツイートのステータスを書き込む
        """
        # 書式を指定
        if report.correction == True or "ツイ消し" in report.memo:
            format1 = self.wb.add_format({'bg_color': 'yellow'})
            format2 = self.wb.add_format({'bg_color': 'yellow'})
            format3 = self.wb.add_format({'bg_color': 'yellow', 'font_color': 'blue', 'font_name': 'Calibri', 'underline': True})
        else:
            format1 = self.wb.add_format()
            format2 = self.wb.add_format()
            format3 = self.wb.add_format({'font_color': 'blue', 'font_name': 'Calibri', 'underline': True})
            
        format1.set_num_format('mm/dd hh:mm:ss') #日付
        ws.write(row, 0, report.time, format1)
        ws.set_column(0, 0, 13) # Width of columns A:A set to 13.
        ws.write(row, 1, report.name)
        ws.write(row, 2, report.screen_name)
        ws.set_column(3, 3, 5) # Width of columns D:D set to 5.
        memo_str = ""
        for memo in report.memo: # memo list をスペースで一行に
            memo_str = memo_str + memo + " "
        memo_str = memo_str.strip()
        ws.write(row, 4, memo_str)
        ws.write(row, 5, report.source, format3)
        
    def __write_quest_data(self, row, ws, report):
        """
        Excelファイルにデータを書き込む
        """
        # 書式を指定
        if report.correction == True:
            format2 = self.wb.add_format({'bg_color': 'yellow'})
        else:
            format2 = self.wb.add_format()
            
        format2.set_num_format('#,##0') #数値
        self.write_col_header(row, ws, report)
        ws.write(row, 6, report.place)
        ws.write(row, 7, report.num, format2)
        l = 8
        for item in report.items.keys():
            ws.write(row, l, item)
            ws.write(row, l+1, report.items[item], format2)
            l = l + 2

    def make_stats_sheets(self, history, use_number, resume_id, resume_time, favlist):
        story_list = ["フリクエ1部", "フリクエ1.5部", "フリクエ2部"]
        if resume_id == -1:
            # days=14の14に意味はない多めの時間 
            time = datetime.datetime.now() - datetime.timedelta(days=14) 
        else:
            if resume_time != "":
                time = datetime.datetime.strptime(resume_time, '%Y-%m-%d %H:%M:%S')
            else:
                ## resume_time が "" の状態でresume_idのツイートが消去されていると
                ## 動作しない
                try:
                    time = id2time(resume_id)
                except:
                    print("[エラー]前回終了時に取得した最後のツイートが消去されたようです")
                    print("setting.ini の resume_id を修正して再実行してください")
                    sys.exit()
        for story in story_list:
            ws = self.wb.get_worksheet_by_name("統計【" + story + "】")
            self.__make_fq_stats(ws, history, use_number, time, favlist)

        self.__make_syuren_stats(history, use_number, time, favlist)


    def __make_syuren_stats(self, history, use_number, time, favlist):
        """
        FGOアイテム効率劇場の楽屋風の修練クエの統計データを作成
        """
        format3 = self.wb.add_format()
        format3.set_num_format('mm/dd') #日付

        youbi = {"弓": "月", "槍":"火", "狂":"水", "騎":"木", "術":"金", "殺":"土", "剣":"日"}
        title = ""
        i = 0
        for syuren in syurenquest.keys():
            if title != youbi[syuren[0]]:
                title = youbi[syuren[0]]
                self.ws_syuren_stat.write(i, 0, title)
            tmp = syuren.split(" ")
            self.ws_syuren_stat.write(i + 1, 1, syuren[0] + tmp[1]) # 弓 + 超級
            self.ws_syuren_stat.write(i + 2, 1, "No.")
            self.ws_syuren_stat.write(i + 3, 1, "周回数")

##            # resume_id の時間より古い要素がどこまでか探る処理
##            # 新しいほうから検索する
##            index = 0
##            for index, t in enumerate(syurenquest[syuren]["メモ"]):
##                if t <= time:
##                    break
            ## リストのどの番号を出力するかのリストを作る
            ## 出力条件 resume_id より新しい or 新規出現
            indexes = []
            for index, id in enumerate(d.get("id") for d in syurenquest[syuren]["report"]):
##                if reversed(syurenquest[syuren]["メモ"])[index] < time or id not in history:
                if id in favlist:
                    continue
                elif len(history.keys()) == 0:
                    if syurenquest[syuren]["report"][index]["time"] > time: #history.csv が無い時
                        indexes.append(index)
                else:
                    if syurenquest[syuren]["report"][index]["time"] > time or id not in history:
                        indexes.append(index)    

##            s_index = -(len(syurenquest[syuren]["メモ"])-index) #リストのresumeする位置

            j = 4
##            for n in reversed(syurenquest[syuren]["周回数"][:s_index]):
            for index, n in enumerate(d.get("周回数") for d in syurenquest[syuren]["report"]):
                if index in indexes:
                    if use_number == True:
                        self.ws_syuren_stat.write(i + 2, j, j - 3)
                    self.ws_syuren_stat.write(i + 3, j, n)
                    j = j + 1
                
            i = i + 4 #縦
            for item in syurenquest[syuren]["ドロップアイテム"].keys():
                self.ws_syuren_stat.write(i, 1, item)
                j = 4 #横
##                for num in reversed(syurenquest[syuren]["ドロップアイテム"][item][:s_index]):
                for index, num in enumerate(d["items"].get(item) for d in syurenquest[syuren]["report"]):
                    if index in indexes:
                        if type(num) is int:
                            if num > -1:
                                self.ws_syuren_stat.write(i, j, num)
                            else:
                                self.ws_syuren_stat.write(i, j, "")
                        else:
                            self.ws_syuren_stat.write(i, j, num)
                        j = j + 1
                i = i + 1
            self.ws_syuren_stat.write(i, 1, "ソース")
            j = 4 #横
##            for url in reversed(syurenquest[syuren]["id"][:s_index]):
            for index, id in enumerate(d.get("id") for d in syurenquest[syuren]["report"]):
                if index in indexes:
                    url = "https://twitter.com/" + syurenquest[syuren]["report"][index]["screen_name"] + "/status/" + str(id)
                    self.ws_syuren_stat.write(i, j, url)
                    j = j + 1
            i = i + 1
            self.ws_syuren_stat.write(i, 1, "メモ")
            j = 4 #横
##            for t in reversed(syurenquest[syuren]["メモ"][:s_index]):
            for index, t in enumerate(d.get("time") for d in syurenquest[syuren]["report"]):
                if index in indexes:
                    self.ws_syuren_stat.write(i, j, t.date(), format3)
                    j = j + 1
            i = i + 2
            
    def __make_fq_stats(self, ws, history, use_number, time, favlist):
        """
        FGOアイテム効率劇場の楽屋風のフリクエの統計データを作成
        """
        format3 = self.wb.add_format()
        format3.set_num_format('mm/dd') #日付
        name = ws.get_name()
        pattern = "統計【(?P<story>.+)】"
        story = re.sub(pattern, r'\g<story>', name)
        tokuiten = ""
        i = 0
        for fq in freequest.keys():
            if freequest[fq]["ストーリー"]==story:
                if tokuiten !=  freequest[fq]["特異点"]:
                    tokuiten = freequest[fq]["特異点"]
                    ws.write(i, 0, tokuiten)
                ws.write(i + 1, 1, freequest[fq]["場所"])
                ws.write(i + 1, 7, fq)
                ws.write(i + 2, 1, "No.")
                ws.write(i + 3, 1, "周回数")

##                index = 0
##                # resume_id の時間より古い要素がどこまでか探る処理
##                # 新しいほうから検索する
##                for index, t in enumerate(freequest[fq]["メモ"]):
##                    if t <= time:
##                        break
##                s_index = -(len(freequest[fq]["メモ"])-index)
                ## リストのどの番号を出力するかのリストを作る
                ## 出力条件 resume_id より新しい or 新規出現
                indexes = []
                for index, id in enumerate(d.get("id") for d in freequest[fq]["report"]):
##                    if reversed(freequest[fq]["メモ"])[index] < time or id not in history:
##                    if freequest[fq]["メモ"][-index-1] < time or id not in history:
                    if id in favlist:
                        continue
                    elif len(history.keys()) == 0:
                        if freequest[fq]["report"][index]["time"] > time:  #history.csv が無い時
                            indexes.append(index)
                    else:
                        if freequest[fq]["report"][index]["time"] > time or id not in history:
                            indexes.append(index)
 
                j = 4
                for index, n in enumerate(d.get("周回数") for d in freequest[fq]["report"]):
                    if index in indexes:
                        if use_number == True:
                            ws.write(i + 2, j, j - 3)
                        ws.write(i + 3, j, n)
                        j = j + 1
                    
                i = i + 4
                for item in freequest[fq]["ドロップアイテム"].keys():
                    ws.write(i, 1, item)
                    j = 4
                    for index, num in enumerate(d["items"].get(item) for d in freequest[fq]["report"]):
                        if index in indexes:
                            if type(num) is int:
                                if num > -1:
                                    ws.write(i, j, num)
                                else:
                                    ws.write(i, j, "")
                            else:
                                ws.write(i, j, num)
                            j = j + 1
                    i = i + 1
                ws.write(i, 1, "ソース")
                j = 4
                for index, id in enumerate(d.get("id") for d in freequest[fq]["report"]):
                    if index in indexes:
                        url = "https://twitter.com/" + freequest[fq]["report"][index]["screen_name"] + "/status/" + str(id)
                        ws.write(i, j, url)
                        j = j + 1
                i = i + 1
                ws.write(i, 1, "メモ")
                j = 4
                for index, t in enumerate(d.get("time") for d in freequest[fq]["report"]):
                    if index in indexes:
                        ws.write(i, j, t.date(), format3)
                        j = j + 1
                i = i + 2

class NoserchExcelFile(ExcelFile):
    def __init__(self, filename):
        story_list = ["フリクエ1部", "フリクエ1.5部", "フリクエ2部"]

        if filename.endswith(".xlsx") == True:
            filename = filename
        else:
            filename = filename + ".xlsx"
        self.wb = xlsxwriter.Workbook(filename)
        self.ws_all = self.wb.add_worksheet("全データ")
        self.ws_syuren = self.wb.add_worksheet("修練場")
        self.ws_fq1 = self.wb.add_worksheet("フリクエ1部")
        self.ws_fq15 = self.wb.add_worksheet("フリクエ1.5部")
        self.ws_fq2 = self.wb.add_worksheet("フリクエ2部")
        self.ws_misc = self.wb.add_worksheet("その他クエスト")
        self.ws_nosearch = self.wb.add_worksheet("未検索") # ここだけ追加
        self.ws_error = self.wb.add_worksheet("Error")
        self.ws_syuren_stat = self.wb.add_worksheet("統計【修練場】")
        for story in story_list:
            ws = self.wb.add_worksheet("統計【" + story + "】")

        self.make_header()

    def make_header(self):
        """
        Excel のヘッダ行を作成
        """
        self.ws_nosearch.write(0, 0, "create date")
        self.ws_nosearch.write(0, 1, "name")
        self.ws_nosearch.write(0, 2, "user name")
        self.ws_nosearch.write(0, 3, "reply")
        self.ws_nosearch.write(0, 4, "memo")        
        self.ws_nosearch.write(0, 5, "url")
        self.ws_nosearch.write(0, 6, "tweet")        
        super().make_header()

    def make_noserch_sheets(self, reports, history, resume_id, favlist):
        """
        「未検索」シートをつくる
        """
        format1 = self.wb.add_format({'bg_color': 'yellow'})
        row = 1
        for report in reports:
            if "リプ数変化" in report.memo:
                pass
            elif report.id <= resume_id and len(history.keys()) == 0: #history.csv が無い時
                continue
            elif report.id <= int(resume_id) and report.id in history.keys():
                continue
            elif report.id in favlist:
                continue
            self.write_col_header(row, self.ws_nosearch, report)
            self.ws_nosearch.write(row, 6, report.original)
            if report.reply_count != None:
                self.ws_nosearch.write(row, 3, report.reply_count)
                if report.correction == True:
                    self.ws_nosearch.set_row(row, None, format1)
            row = row + 1
            
def add_reply_info(reports, replies):
    """
    各周回報告についてるリプライ件数を計上する
    訂正の言動があるかもチェックする
    """
    new_reports = []
    for report in reports:
        report.reply_count = 0
        for reply in replies[report.screen_name]:
            if report.id == reply.in_reply_to_status_id:
                report.reply_count = report.reply_count + 1
                pattern = "(訂正|修正|間違)"
                m = re.search(pattern, reply.full_text)
                if m:
                    report.correction = True
                    report.memo.append("訂正リプ有")
##                    report.category = "Error"
##                    report.items = {}
        new_reports.append(report)
    return new_reports

def make_replies(reports, favlist):
    """
    投稿者のreply を取得する
    
    """
    # OAuth認証
    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)
    api = tweepy.API(auth)
    max_loop = 5
    replies = {}

    print("投稿者のタイムラインを調査中")

    with tqdm(total=len(reports)) as pbar:

        for report in reversed(reports):
            max_id = -1
            status = None
            # 取得したReplyは複数報告の投稿者のために再利用するため、逆順(古い順)でソート
            if report.screen_name not in replies.keys():
                replies[report.screen_name] = []
                for loop in range(max_loop):
                    try:
                        if max_id == -1:
                            for status in api.user_timeline(screen_name=report.screen_name, tweet_mode="extended",
                                                            count=200,
                                                            since_id=report.id,
                                                            exclude_replies="false",
                                                            include_rts="false"):
                                if status.in_reply_to_screen_name == report.screen_name:
                                    if "media" not in status.entities and status.id not in favlist:
                                        replies[report.screen_name].append(status)
                                if status == None:
                                    flag = True
                                    break
                        else:
                            for status in api.user_timeline(screen_name=report.screen_name, tweet_mode="extended",
                                                            max_id=max_id -1,
                                                            count=200,
                                                            since_id=report.id,
                                                            exclude_replies="false",
                                                            include_rts="false"):
                                if status.in_reply_to_screen_name == report.screen_name:
                                    if "media" not in status.entities and status.id not in favlist:
                                        replies[report.screen_name].append(status)
                                if status == None:
                                    flag = True
                                    break
                        if status != None:
                            max_id=status.id
                    except tweepy.error.TweepError as err:
##                        print(err)
##                        print(report.screen_name)
##                        print(report.id)
                        if status != None:
                            max_id=status.id
                        continue
                    else:
                        break
            pbar.update(1)

    return replies
                    
##def make_replies2(reports):
##    """
##    reply を取得する
##    
##    """
##    # OAuth認証
##    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
##    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)
##    api = tweepy.API(auth)
##    max_loop = 5
##    replies = {}
##
##    print("投稿者の自身にリプライするツイートを検索中")
##
##    with tqdm(total=len(reports)) as pbar:
##
##        for report in reversed(reports):
##            max_id = -1
##            status = None
##
##            # 取得したReplyは複数投稿者のために再利用するため、逆順でソート
##            if report.screen_name not in replies.keys():
##                replies[report.screen_name] = []
##                for loop in range(max_loop):
##                    try:
##                        q_str = "to:" + report.screen_name + " from:" + report.screen_name
##    ##                        for status in api.search(q="'to:@" + report.screen_name + " from:@" + report.screen_name + "'",
##                        for status in api.search(q=q_str,
##                                                max_id=max_id,
##                                                tweet_mode="extended",
##                                                count=100,
##                                                since_id=report.id,
##                                                result_type="mixed"):
##                            if status.in_reply_to_screen_name == report.screen_name:
##                                replies[report.screen_name].append(status)
####                            else: # Retweet がここに入る
####                                print("https://twitter.com/", end="")
####                                print(status.user.screen_name, end="")
####                                print("/status/", end = "")
####                                print(status.id)
##                            if status == None:
##                                flag = True
##                                break
##                        if status != None:
##                            max_id=status.id - 1
##                        else:
##                            break
##                    except tweepy.error.TweepError:
##                        break
##            pbar.update(1)
##
##    return replies

def get_favlist(ACCESS_TOKEN, ACCESS_SECRET, since_id):
    """
    自身のいいねリストを取得する
    """
    # OAuth認証
    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)
    api = tweepy.API(auth)
    max_loop = 2
##    resume_id = last_id
    status = None
    max_id = -1

    ### 新しいほうからとれる
    ### 取得したidの古いところまでとれたら終了
    ###

    reports = []
##    with tqdm() as pbar:
    for loop in range(max_loop):
        if max_id == -1:
            for status in api.favorites(count=200, since_id=since_id -1):
                reports.append(status.id)
##                    pbar.update(1)
        else:
            for status in api.favorites(count=200, max_id=max_id, since_id=since_id -1):
                reports.append(status.id)
##                    pbar.update(1)
        if status != None:
            max_id=status.id -1
    return reports

def make_nofavreports(reports, favlist):
    new_reports = []
    for report in reports:
        if report.id not in favlist:
            new_reports.append(report)
    return new_reports
            
def get_tweet(ACCESS_TOKEN, ACCESS_SECRET):
    """
    #FGO周回カウンタ ハッシュタグのツイートを取得する
    """
##    global last_id
##    global since_id
    since_id = 999999999999999999999

    # OAuth認証
    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)
    api = tweepy.API(auth)
    
    max_id = 0
    max_loop = MAXLOOP
    # twitter内を検索し、結果をエクセルに書き込む
    resume_id = last_id
    status = None
    nglist = NG_NAME.split()
    ngidlist = NG_ID.split()
    reports = []
    print("Twitter API で周回データを検索中")
    with tqdm() as pbar:
        for loop in range(max_loop):
            for status in api.search(q='#FGO周回カウンタ',
                                     lang='ja',result_type='mixed',
                                     count=MAXSERCH,
                                     max_id=max_id -1,
                                     tweet_mode="extended"):
    #max_id
    #ページングに利用する。ツイートのIDを指定すると、これを含み、これより過去のツイートを取得できる。
                # NG name
                if status.user.screen_name not in nglist and status.user.id_str not in ngidlist:
                    if 'RT @' not in status.full_text \
                       and '#FGO買取' not in status.full_text \
                       and '#FGO販売' not in status.full_text :
                        reports.append(ReportTweet(status))
                        pbar.update(1)

                #どこまで検索したか記録
    ##            if int(max_id) < int(status.id):
    ##                max_id = int(status.id)               
                if int(since_id) > int(status.id):
                    since_id = int(status.id)
    
            if status != None:
                max_id=status.id
    
    return reports, since_id


def sort_quest():
    for fq in freequest.keys():
        freequest[fq]["report"].sort(key=lambda x: x['time'])
    for sq in syurenquest.keys():
        syurenquest[sq]["report"].sort(key=lambda x: x['time'])
        
def read_freequest():
    """
    CSV形式のフリークエストデータを読み込む
    """
    fqfile = os.path.join(os.path.dirname(__file__), 'freequest.csv')
    with open(fqfile, 'r', encoding="utf_8") as f:
        try:
            reader = csv.reader(f)
            header = next(reader)  # ヘッダーを読み飛ばしたい時

            for row in reader:
                q = {}
                q["ストーリー"] = row[0]
                q["特異点"] = row[1]
                q["場所"] = row[2]
                q["周回数"] = []
                d = {}
                for item in row[4:]:
                    if item == "":
                        break
                    d[normalize_item(item)] = []
                q["ドロップアイテム"] = d
##                q["id"] = []
##                q["screen_name"] = []                
##                q["メモ"] = []
                q["report"] = []
                freequest[row[3]] = q
        except UnicodeDecodeError:
            print("[エラー]freequest.csv の文字コードがおかしいようです。UTF-8で保存してください。")
            sys.exit()
        except IndexError:
            print("[エラー]freequest.csv がCSV形式でないようです。")
            sys.exit()
                

def read_syurenquest():
    """
    CSV形式の修練クエストデータを読み込む
    """
    syurenfile = os.path.join(os.path.dirname(__file__), 'syurenquest.csv')
    with open(syurenfile, 'r', encoding="utf_8") as f:
        try:
            reader = csv.reader(f)
            header = next(reader)  # ヘッダーを読み飛ばしたい時

            for row in reader:
                q = {}
                q["周回数"] = []
                d = {}
                for item in row[1:]:
                    if item == "":
                        break
                    d[normalize_item(item)] = []
                q["ドロップアイテム"] = d
##                q["id"] = []
##                q["screen_name"] = []
##                q["メモ"] = []
                q["report"] = []
                syurenquest[row[0]] = q
        except UnicodeDecodeError:
            print("[エラー]syurenquest.csv の文字コードがおかしいようです。UTF-8で保存してください。")
            sys.exit()
        except IndexError:
            print("[エラー]syurenquest.csv がCSV形式でないようです。")
            sys.exit()

def read_item():
    """
    CSV形式のアイテム変換データを読み込む
    """
    itemfile = os.path.join(os.path.dirname(__file__), 'item.csv')
    with open(itemfile, 'r' , encoding="utf_8") as f:
        try:
            reader = csv.reader(f)
            header = next(reader)  # ヘッダーを読み飛ばしたい時

            for row in reader:
##                q = {}
                for item in row[2:]:
                    if item == "":
                        break
                    sozai_betsumei[item] = row[1]
                sozai[row[1]] = row[0]
        except UnicodeDecodeError:
            print("[エラー]item.csv の文字コードがおかしいようです。UTF-8で保存してください。")
            sys.exit()
        except IndexError:
            print("[エラー]item.csv がCSV形式でないようです。")
            sys.exit()

def read_quest():
    """
    CSV形式のクエスト変換データを読み込む
    """
    itemfile = os.path.join(os.path.dirname(__file__), 'quest.csv')
    with open(itemfile, 'r' , encoding="utf_8") as f:
        try:
            reader = csv.reader(f)
            header = next(reader)  # ヘッダーを読み飛ばしたい時

            for row in reader:
##                q = {}
                for name in row:
                    if name == "":
                        break
                    quest[name] = row[0]
        except UnicodeDecodeError:
            print("[エラー]item.csv の文字コードがおかしいようです。UTF-8で保存してください。")
            sys.exit()
        except IndexError:
            print("[エラー]item.csv がCSV形式でないようです。")
            sys.exit()
                                                            
def read_history(NG_NAME, NG_ID):
    """
    CSV形式の履歴データを読み込む
    """
    history = {}
    itemfile = os.path.join(os.path.dirname(__file__), 'history.csv')

    if os.path.exists(itemfile) == False:
        print("history.csv ファイルを新規作成します")
        return history

    ngnamelist = NG_NAME.split()
    ngidlist = NG_ID.split()
    for ngid in ngidlist:
        #BANされた場合の処理
        try:
            ngnamelist.append(userid2screen_name(ngid))
        except tweepy.error.TweepError as err:
            if "User has been suspended." in str(err):
                print("setting.ini の ng_id に記述されている " + ngid + " はBANされました")

    with open(itemfile, 'r' , encoding="utf_8") as f:
        try:
            reader = csv.reader(f)
            header = next(reader)  # ヘッダーを読み飛ばしたい時

            for row in reader:
                q = {}
                for item in row:
                    q["time"] = datetime.datetime.strptime(row[0], '%Y-%m-%d %H:%M:%S')
                    if row[2] == "":
                        q["reply"] = None
                    else:
                        q["reply"] = int(row[2])
                    q["text"] = ""
                    if len(row) > 3:
                        q["name"] = row[3]
                        q["screen_name"] = row[4]
                        q["text"] = row[5]
                    else:
                        q["name"] = None
                        q["screen_name"] = None
                        q["text"] = ""
                if row[4] not in ngnamelist and '#FGO販売' not in row[5]:
                    history[int(row[1])] = q
                if row[1].isdigit() == False:
                    print("[エラー] history.csv がおかしいようです。削除してください。")
                    sys.exit()
        except UnicodeDecodeError:
            print("[エラー]history.csv の文字コードがおかしいようです。UTF-8で保存してください。")
            sys.exit()
        except (IndexError, ValueError) as err:
            print("[エラー]history.csv が正しいCSV形式でないようです。")
            print(err)
            sys.exit()
    return history

def write_history(history):
    """
    CSV形式の履歴データを書き込む
    """
    csvfile = os.path.join(os.path.dirname(__file__), 'history.csv')

    #リストを作成
    header = ["time", "id", "reply", "name", "screen_name", "text"]
    rows = []
    rows.append(header)
    for id in sorted(history, reverse=True):
        row = []
        row.append(history[id]["time"])
        row.append(id)
        row.append(history[id]["reply"])
        row.append(history[id]["name"])
        row.append(history[id]["screen_name"])
        row.append(history[id]["text"])
        rows.append(row)

    with open(csvfile, 'w', encoding='UTF-8') as f:
        writer = csv.writer(f, lineterminator='\n') # 改行コード（\n）を指定しておく
        writer.writerows(rows) # 2次元配列


def make_new_history(reports, history):
    new_history = {}
    for report in reports:
        if "ツイ消し" not in report.memo:
            q = {}
            q["time"] = report.time
            if report.reply_count == None:
                if report.id in history.keys():
                    q["reply"] = history[report.id]["reply"]
                else:
                    q["reply"] = None
            else:
                q["reply"] = report.reply_count
            q["name"] = report.name
            q["screen_name"] = report.screen_name
            q["text"] = report.original
            new_history[report.id] = q

    #ソート
    return new_history
        

def make_history_replies(yahoo_reports, history, replies, favlist):
    """
    history から replies にない投稿者のツイートを探す
    """
    # OAuth認証
    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)
    api = tweepy.API(auth)
    max_loop = 5
    new_replies = {}

    #yahoo_reportsのid
    id_yahoos = []
    old_id = 999999999999999999999
    for report in yahoo_reports:
        id_yahoos.append(report.id)
        if old_id > report.id:
            old_id = report.id
    id_yahoos = set(id_yahoos)

    id_history = []
    for key in history.keys():
        if old_id <= key:
            id_history.append(key)
    id_history = set(id_history)
    
    print("履歴のみに報告のある投稿者を調査中")
    with tqdm(total=len(id_history - id_yahoos)) as pbar:
        for id in (id_history - id_yahoos):
            max_id = -1
            status = None
            # id から投稿者を特定
            try:
                screen_name = id2screen_name(id)
            except tweepy.error.TweepError:
                # 投稿が削除されたときなど
                continue
            # replies に投稿者が無いとき
            if screen_name not in replies.keys():
                new_replies[screen_name] = []
                for loop in range(max_loop):
                    try:
                        if max_id == -1:
                            for status in api.user_timeline(screen_name=screen_name, tweet_mode="extended",
                                                            count=200,
                                                            since_id=id,
                                                            exclude_replies="false",
                                                            include_rts="false"):
                                if status.in_reply_to_screen_name == screen_name:
                                    if "media" not in status.entities and status.id not in favlist:
                                        new_replies[screen_name].append(status)
                                if status == None:
                                    flag = True
                                    break
                        else:
                            for status in api.user_timeline(screen_name=screen_name, tweet_mode="extended",
                                                            max_id=max_id -1,
                                                            count=200,
                                                            since_id=id,
                                                            exclude_replies="false",
                                                            include_rts="false"):
                                if status.in_reply_to_screen_name == screen_name:
                                    if "media" not in status.entities and status.id not in favlist:
                                        new_replies[screen_name].append(status)
                                if status == None:
                                    flag = True
                                    break
                    except tweepy.error.TweepError:
##                        print("\n[エラー]原因不明のTweepErrorです")
##                        print(screen_name)
##                        print(id)
                        continue
                    if status != None:
                        tmp_max_id=status.id
                    else:
                        break
            pbar.update(1)
    return new_replies


def check_history(yahoo_reports, history, replies, favlist):
    """
    履歴をチェック
    リプライ数に差がないか調べる
    """
    # yahoo_reports を脳死コピー
    api = tweepy.API(auth)
    ids = copy.copy(favlist)
##    for report in yahoo_reports:
##        h = {}
##        h["time"]=report.time
##        h["id"]=report.id
##        h["reply"]=report.reply
##        new_reports.append(h)
    # 2. yahoo_reportsと history で時間が重複したidと reply 数に変化がないかチェック
    # 変化していたら memo に追加
    old_id = yahoo_reports[-1].id
    if len(replies) != 0:
        for report in yahoo_reports:
            if report.id in history.keys() and len(replies.keys()) != 0:
                if history[report.id]["reply"] != None: #消去されたときなど
                    if report.reply_count != history[report.id]["reply"]:
                        report.memo.append("リプ数変化")
    ##                    report.correction = True
            ids.append(report.id)
    else:
        for report in yahoo_reports:
            ids.append(report.id)        
    #3. csv にしかないidもチェックしてコピー
    #(重複しない)検索した一番古いidから10日前までの新規情報を取得
    # 変化していたら memo に追加　reply数を更新
##    with tqdm(total=len(id_history - id_yahoos)) as pbar:
##    print("履歴のみにある投稿をチェック中")
    for id in history.keys():
        if id not in ids: #csvにしかないid
            # idから時間をゲットして10日以上前ならbreak
            try:
                if(datetime.datetime.now() - id2time(id)).days >= 10:
                    break                        
                if id < old_id:
                    break
            except tweepy.error.TweepError: #消されたときなど
                continue

            #リプライ数がhistoryと違ったらmemoに
            #リプライ数を取得
            screen_name = id2screen_name(id)
            reply_count = 0
            correction = False

            if len(replies.keys()) != 0:
                for reply in replies[screen_name]:
                    if id == reply.in_reply_to_status_id:
                        reply_count = reply_count + 1
                        pattern = "(訂正|修正|間違)"
                        m = re.search(pattern, reply.full_text)
                        if m:
                            correction = True

##            if reply_count != int(history[id]["reply"]) or correction == True:
            #リポートを yahoo_reportに加える    
            status = api.get_status(id, tweet_mode="extended")
            r = ReportTweet(status)
            if len(replies.keys()) != 0:
                r.reply_count = reply_count
                if history[id]["reply"] != None:
                    if reply_count !=  history[id]["reply"]:
                        r.memo.append("リプ数変化")
                if correction == True:
                    r.correcton = True
                    r.memo.append("訂正リプ有")
            yahoo_reports.append(r)

    #yahoo_reporsをソート
    yahoo_reports = sorted(yahoo_reports, key=attrgetter("time"), reverse=True)
                

    return yahoo_reports

def get_oauth_token(url:str)->str:
    querys = urllib.parse.urlparse(url).query
    querys_dict = urllib.parse.parse_qs(querys)
    return querys_dict["oauth_token"][0]

def compare_twitter_history(reports, history):
    """
    Twitter からのデータと履歴を比較
    次のデータをチェック
    1. 新規に投稿されたデータ
    2. 鍵垢から公開垢になって現れたデータ
    3. 削除されたデータ
    """
    if len(history) == 0:
        return reports, [], []
    
    new_count = 0
    restore_count = 0
    del_count = 0
    public_count = 0
    new_time = history[(list(history)[0])]["time"] #履歴にあるうちで最新の時間
    since_time = reports[-1].time #Twitterデータで最古の時間
    tw_list = []
    restore_id = []
    new_reports = []
##    print("Twitter からのデータと履歴を比較中")
    for report in reports:
        if report.id not in history.keys():
            if report.time > new_time:
                new_count = new_count +1
            elif report.time < new_time: #前回抜けてるものを発見
                restore_count = restore_count + 1
                restore_id.append(report.id)
                report.memo.append("前回未取得")
        # reports にあるid一覧を取得する
        tw_list.append(report.id)
        new_reports.append(report)
    print("Twitterからの新規取得データ: ", end = "")
    print(restore_count + new_count, end = "件")
    if restore_count > 0:
        print("(うち既取得データ時間以前のもの: ", end = "")
        print(restore_count, end = "件)")

    public_id = []
    deleted_id = []
    ## 削除されたデータを検出する
    for id in history.keys():
        if id not in tw_list and history[id]["time"] >= since_time:
            try:
                screen_name = id2screen_name(id)
            except tweepy.error.TweepError as err:
##                if err == "Not authorized.": #鍵垢化
                ## Sorry, you are not authorized to see this status.
                if "authorized" in str(err): #鍵垢化
                    public_count = public_count + 1
                    public_id.append(id)
                    # history から鍵垢化ツイートのデータを作成
                    status = TweetStatus()
                    status.created_at = history[id]["time"] - datetime.timedelta(hours=9)
                    status.user.name = history[id]["name"]
                    status.user.screen_name = history[id]["screen_name"]
                    status.id = id
                    status.id_str = str(id)
                    status.full_text = history[id]["text"]
                    status.text = history[id]["text"]
                    p = PrivateTweet(status)
                    new_reports.append(p)
                else:
                    del_count = del_count + 1
                    deleted_id.append(id)
                    # history から削除ツイートのデータを作成
                    status = TweetStatus()
                    status.created_at = history[id]["time"] - datetime.timedelta(hours=9)
                    status.user.name = history[id]["name"]
                    status.user.screen_name = history[id]["screen_name"]
                    status.id = id
                    status.id_str = str(id)
                    status.full_text = history[id]["text"]
                    status.text = history[id]["text"]
                    d = DeletedTweet(status)
                    new_reports.append(d)
                 
    if del_count > 0:
        print(", 削除データ: ", end = "")
        print(del_count, end = "件")
    if public_count > 0:
        print(", 鍵垢化データ: ", end = "")
        print(public_count, end = "件")
    print()

    
    return new_reports, restore_id, deleted_id
    

def create_access_key_secret():
#if __name__ == '__main__':

    auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)

    try:
        redirect_url = auth.get_authorization_url()
        print ("次のURLをウェブブラウザで開きます:",redirect_url)
    except tweepy.TweepError:
        print( "[エラー] リクエストされたトークンの取得に失敗しました。")


    oauth_token = get_oauth_token(redirect_url)
    print("oauth_token:", oauth_token)
    auth.request_token['oauth_token'] = oauth_token

    # Please confirm at twitter after login.
    webbrowser.open(redirect_url)

#    verifier = input("You can check Verifier on url parameter. Please input Verifier:")
    verifier = input("ウェブブラウザに表示されたPINコードを入力してください:")
    auth.request_token['oauth_token_secret'] = verifier

    try:
        auth.get_access_token(verifier)
    except tweepy.TweepError:
        print('[エラー] アクセストークンの取得に失敗しました。')

    print("access token key:",auth.access_token)
    print("access token secret:",auth.access_token_secret)

    config = configparser.ConfigParser()
    section1 = "auth_info"
    config.add_section(section1)
    config.set(section1, "ACCESS_TOKEN", auth.access_token)
    config.set(section1, "ACCESS_SECRET", auth.access_token_secret)

    settingfile = os.path.join(os.path.dirname(__file__), 'setting.ini')
    with open(settingfile, "w") as file:
        config.write(file)


    print("Twitterのアプリ認証は正常に終了しました。")

def check_reports(reports):
    for report in reports:
        if report.id == 1156582907227648002:
            print("【発見】消去ツイート")
        if report.id == 1152207853605855239:
            print("【発見】エラーツイート")

if __name__ == '__main__':
##    global noclass    
    last_id = -1
    last_time = ""
    use_number = False
    use_yahoo = False

    if os.path.exists(settingfile) == False:
        create_access_key_secret()
##        print("設定ファイルがありません。syutagcnt.py を実行してください。")
        sys.exit()

    config = configparser.ConfigParser()

    try:
        config.read(settingfile)
        section0 = "search"
        section1 = "auth_info"
        ACCESS_TOKEN = config.get(section1, "ACCESS_TOKEN")
        ACCESS_SECRET = config.get(section1, "ACCESS_SECRET")
        if section0 not in config.sections():
            config.add_section(section0)
        section0cfg = config[section0]
        MAXSERCH = section0cfg.getint("MAXSERCH", MAXSERCH)
        config.set(section0, "MAXSERCH", str(MAXSERCH))
        MAXLOOP = section0cfg.getint("MAXLOOP", MAXLOOP)
        config.set(section0, "MAXLOOP", str(MAXLOOP))
        NG_NAME = section0cfg.get("NG_NAME", NG_NAME)
        config.set(section0, "NG_NAME", NG_NAME)
        NG_ID = section0cfg.get("NG_ID", NG_ID)
        config.set(section0, "NG_ID", NG_ID)
        resume_id = section0cfg.getint("last_id", last_id)
        resume_time = section0cfg.get("last_time", last_time)
        with open(settingfile, "w") as file:
            config.write(file)

    except configparser.NoSectionError:
        print("[エラー] 設定ファイルに不備があります。setting.ini を消して再実行してください。")
        sys.exit()

    parser = argparse.ArgumentParser(description='FGO周回カウンタの報告を集めExcel出力する')
    # 3. parser.add_argumentで受け取る引数を追加していく
    parser.add_argument('filename', help='出力Excelファイル名')    # 必須の引数を追加
    parser.add_argument('-i', '--ignoreclass', help='クラス無しをエラーを無視する', action='store_true')
    parser.add_argument('-y', '--yahoo', help='Yahoo!リアルタイム検索からデータを取得する', action='store_true')
    parser.add_argument('-c', '--checkreply', help='リプライデータをチェックする', action='store_true')
    parser.add_argument('-r', '--resume', help='前回取得したツイートの続きから取得', action='store_true')
    parser.add_argument('-u', '--url', help='指定したURLより未来のツイートを取得')
    parser.add_argument('-a', '--asc', help='出力を昇順にする', action='store_true')
    parser.add_argument('-f', '--nofavorited', help='自分がふぁぼしていないツイートのみを取得', action='store_true')
    parser.add_argument('-n', '--number', help='統計シートの報告にNoをつける', action='store_true')
    parser.add_argument('-w', '--wait', help='ブラウザ操作時の待ち時間(秒)を指定する(デフォルト2秒)', type=int, default=2)
    parser.add_argument('--version', action='version', version=progname + " " + version)

    args = parser.parse_args()    # 4. 引数を解析
    if args.url != None: #https://twitter.com/ /status/1146270030495133697
        #パターンチェック
        tweet_pattern = "https://twitter.com/.+?/status/"
        if not re.match(tweet_pattern , args.url):
            print("[エラー] URLがTwitterのものではありません")
            sys.exit()
        resume_id = int(re.sub(tweet_pattern, "", args.url))
##        use_resume = True
    else: # URLが指定されないときだけ resume 可能
        if args.resume == False:
            resume_id = -1

    if args.ignoreclass == True:
        noclass = True
    if args.number == True:
        use_number = True
    if args.nofavorited == True:
        nofavorited_only = True
    if args.asc == True:
        ascending_order = True
    else:
        ascending_order = False        
    if args.yahoo == True:
        use_yahoo = True

    if args.wait < 1:
        print("[エラー] WAITの値は1以上でないといけません")
        sys.exit()

    # Excelファイルチェック
    # 処理が長くなるので、処理後に書き込めないといったことが無いようここでいったんチェックする
    try:
        f = ExcelFile(args.filename)            
        f.close()
    except PermissionError:        
        print("[エラー] Excelファイルに書き込めません。Excelで開いている場合は閉じるかファイル名を変更してください。")
        sys.exit()
    
    auth.set_access_token(ACCESS_TOKEN, ACCESS_SECRET)

    #各データCSVの存在チェック
    csvfiles = ["freequest.csv", "syurenquest.csv", "item.csv"]
    for csvfile in csvfiles:
        target_path = os.path.join(os.path.dirname(__file__), csvfile)
        if os.path.isfile(target_path) == False:
            print("[エラー] " + csvfile + " が見つかりません。 "+ os.path.dirname(__file__) +" に入れてください。")
            sys.exit()
    #各データCSVを読み込む
    read_item()
    read_freequest()
    read_syurenquest()
    read_quest()

    #メインの処理はここから
    history = read_history(NG_NAME, NG_ID)


    try:
##    yahoo_reports = get_tweet()
        #ツイッターからデータを取得
        reports, since_id = get_tweet(ACCESS_TOKEN, ACCESS_SECRET)
##        check_reports(reports)
        # history とツイッターデータを比較
        reports, restore_id, deleted_id = compare_twitter_history(reports, history)
##        check_reports(reports)
##    yahoo_reports = get_yahoo_reports()
        yt = YahooTweets(reports, history, since_id, args.wait, use_yahoo)
        yahoo_reports = yt.reports
##        check_reports(yahoo_reports)
        #いいねしたツイートの処理
        favlist = []
        if nofavorited_only == True:
            favlist = get_favlist(ACCESS_TOKEN, ACCESS_SECRET, since_id)
##            yahoo_reports = make_nofavreports(yahoo_reports, favlist)

        replies = {}
        if args.checkreply == True:
            replies = make_replies(yahoo_reports, favlist) #いいねしたリプライも除く
            # yahoo_reports に reply 情報を付与
            yahoo_reports = add_reply_info(yahoo_reports, replies)
            # history のみにある投稿者のリプライを加える
            # compare_twitter_history() で加えるので未使用
            replies.update(make_history_replies(yahoo_reports, history, replies, favlist))
            yahoo_reports = add_reply_info(yahoo_reports, replies)

        # 履歴をチェック・更新
        yahoo_reports = check_history(yahoo_reports, history, replies, favlist)            

    except selenium.common.exceptions.TimeoutException:
        print("[エラー] タイムアウトしました。WAIT 値が小さすぎる場合は、--wait でより大きな値を指定してください。")
        sys.exit()
    except tweepy.error.RateLimitError:
        print("[エラー] Twitter APIの使用制限に達しました。15分待ってから再実行してください。")
        sys.exit()


    last_id = yahoo_reports[0].id
    if ascending_order == True:
        yahoo_reports = yahoo_reports[::-1]


##    #履歴を保存したあとにいいねしていないツイートのみにする
##    if nofavorited_only == True:
##        yahoo_reports = make_nofavreports(yahoo_reports, favlist)

    #Excelファイルを生成
    try:
        if use_yahoo == True:
            f = NoserchExcelFile(args.filename)
            f.make_noserch_sheets(yt.unsearch_reports, history, resume_id, favlist)
        else:
            f = ExcelFile(args.filename)            
        f.make_sheets(yahoo_reports, history, resume_id, favlist)
        ## freequest を時間でソート
        ## syurenquest を時間でソート
        sort_quest()
        f.make_stats_sheets(history, use_number, resume_id, resume_time, favlist)
        f.close()
    except PermissionError:        
        print("[エラー] Excelファイルに書き込めません。Excelで開いている場合は閉じるかファイル名を変更してください。")
        sys.exit()

    history = make_new_history(yahoo_reports, history)
    write_history(history)
        
    #設定ファイルにどのIDまで検索したか記録
    config.set(section0, "last_id", str(last_id))
    last_time = id2time(last_id)
    config.set(section0, "last_time", str(last_time))
    with open(settingfile, "w") as file:
        config.write(file)


