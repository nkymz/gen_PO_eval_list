# -*- coding: utf-8 -*-

import os
import re
import datetime
import time

import openpyxl
import requests
from bs4 import BeautifulSoup
import pprint

def gen_POH_eval_list(HDN_eval_list, uma_eval_list):
    DICT_EVAL = {"Ｓhyouka_siba":"5T", "Ａhyouka_siba":"4T", "Ｂhyouka_siba":"3T", "Ｃhyouka_siba":"2T",
                 "Ｄhyouka_siba":"1T", "Ｅhyouka_siba":"0T", "Ｓhyouka_dirt":"5D", "Ａhyouka_dirt":"4D",
                 "Ｂhyouka_dirt":"3D", "Ｃhyouka_dirt":"2D", "Ｄhyouka_dirt":"1D", "Ｅhyouka_dirtsiba":"0D"}
    path = os.getenv("HOMEDRIVE", "None") + os.getenv("HOMEPATH", "None") + "/Dropbox/POG/"
    wbpath = (path + "PO_HorseEvalList.xlsx").replace("\\", "/")
    wb = openpyxl.load_workbook(wbpath)
    ws = wb["POHEvalList"]
    xlrow = 1

    while ws.cell(row=xlrow, column=1).value:
        horse_name = ws.cell(row=xlrow, column=1).value

        if ws.cell(row=xlrow, column=6).value == "HDN_eval_new":
            ws.cell(row=xlrow, column=6).value = "HDN_eval_exist"
            xlrow += 1
            continue
        elif ws.cell(row=xlrow, column=6).value == "HDN_eval_exist":
            xlrow += 1
            continue

        for row in HDN_eval_list:
            if horse_name != row[4]:
                continue
            score = 0
            for i in range(3):
                myeval = DICT_EVAL[row[i + 10] + row[i + 13]]
                ws.cell(row=xlrow,column=7+i).value = myeval
                score += int(myeval[0]) * 7
            ws.cell(row=xlrow, column=10).value = score
            ws.cell(row=xlrow, column=6).value = "HDN_eval_new"
            break

        xlrow += 1

    xlrow = 1
    while ws.cell(row=xlrow, column=1).value:
        horse_name = ws.cell(row=xlrow, column=1).value
        if ws.cell(row=xlrow, column=11).value == "UMA_eval_new":
            ws.cell(row=xlrow, column=11).value = "UMA_eval_exist"
            xlrow += 1
            continue
        elif ws.cell(row=xlrow, column=11).value == "UMA_eval_exist":
            xlrow += 1
            continue
        for row in uma_eval_list:
            if horse_name != row[0]:
                continue
            ws.cell(row=xlrow, column=12).value = row[1]
            ws.cell(row=xlrow, column=13).value = row[2]
            ws.cell(row=xlrow, column=14).value = row[2] + ws.cell(row=xlrow, column=10).value
            ws.cell(row=xlrow, column=11).value = "UMA_eval_new"
            break
        xlrow += 1

    wb.save(wbpath)

    xPOH_eval_list_all = [[cell.value for cell in row] for row in ws["A1:N" + str(xlrow - 1)]]
    return [row for row in xPOH_eval_list_all if row[5] != "HDN_eval_none" or row[10] != "UMA_eval_none"]

def gen_uma_eval_list():

    uma_url_1st_half = "http://umakeiba.com/post/category/%E5%84%AA%E9%A6%AC2%E6%AD%B3%E9%A6%AC%E3%83%81%E3%82%A7%E3%83%83%E3%82%AF/page/"

    uma_eval_list = []

    i = 1
    while True:
        target_url = uma_url_1st_half + str(i) + "/"

        time.sleep(1)
        r = requests.get(target_url)
        if r.status_code != requests.codes.ok:
            break
        soup = BeautifulSoup(r.content, 'lxml')

        for h2tag in soup.find_all("h2"):
            if not h2tag.find("a"):
                continue
            atag = h2tag.find("a")
            if not re.search(r"★評価一覧", atag.string):
                continue
            else:
                eval_page_url = atag.get("href")
                eval_page_no = int(eval_page_url.split("/")[-2])

            if eval_page_no < 6665:
                break
            time.sleep(1)
            r = requests.get(eval_page_url)
            if r.status_code != requests.codes.ok:
                break
            soup = BeautifulSoup(r.content, 'lxml')

            if len(soup.find_all("strong")) < 2:
                continue

            if soup.find_all("strong")[1].string[-1] == "点":
                eval_horses = [strong_tag.string for k, strong_tag in enumerate(soup.find_all("strong"))
                              if k % 2 == 0]
                eval_stars = [int(strong_tag.string[-2]) for k, strong_tag in enumerate(soup.find_all("strong"))
                              if k % 2 == 1]
            else:
                eval_horses = [strong_tag.string for k, strong_tag in enumerate(soup.find_all("strong"))
                               if k % 3 == 0]
                eval_stars = [int(strong_tag.string[-2]) for k, strong_tag in enumerate(soup.find_all("strong"))
                              if k % 3 == 2]
            for horse, star in zip(eval_horses, eval_stars):
                score = (star - 2 if star >= 2 else 0) * 5 * 3
                uma_eval_list.append([horse, star, score])
        else:
            i += 1
            continue
        break

    return uma_eval_list




def gen_HDN_eval_list():

    HDN_URL_1ST_HALF = "http://www.nikkankeiba.com/pog2018/hyouka/hyouka"
    url_2nd_half = ["01.html", "02.html", "03.html", "04.html", "05.html", "06.html", "07.html", "08.html,"
                    "09.html", "10.html", "11.html", "12.html"]
    HDN_eval_list = []

    for s in url_2nd_half:

        target_url = HDN_URL_1ST_HALF + s

        time.sleep(1)
        r = requests.get(target_url)
        if r.status_code != requests.codes.ok:
            break
        r.encoding = "euc-jp"
        soup = BeautifulSoup(r.text, 'lxml')

        mytable = soup.find("table").find_all("tr")

        for i, myrow in enumerate(mytable):
            if i == 0:
                continue
            eval_row = [element.string for element in myrow.find_all("td")]
            for j in range(3):
                eval_row.append(myrow.find_all("td")[j + 10].get("id"))

            HDN_eval_list.append(eval_row)

    return HDN_eval_list


def wrap_trtd(columns, tag_type):
    begin_tag = "<" + tag_type + ">"
    end_tag = "</" + tag_type + ">"
    t = "<tr>"
    for s in columns:
        t = t + begin_tag + s + end_tag
    t += "</tr>\n"

    return t


def out_POH_eval_list(xPOH_eval_list):
    path = os.getenv("HOMEDRIVE", "None") + os.getenv("HOMEPATH", "None") + "/Dropbox/POG/"
    htmlpath = path + "POH_eval_list.html"
    f = open(htmlpath, mode="w", encoding="utf-8")
    f.write('<table border="1" cellspacing="0" cellpadding="5" bordercolor="#333333">\n')
    f.write(wrap_trtd(["馬名", "オーナー", "性別", "指名順", "HDN1", "HDN2", "HDN3", "UMA"], "th"))
    for row in xPOH_eval_list:
        horse_name = row[0] if row[4] == "-" else "<s>" + row[0] + "</s>"
        f.write(wrap_trtd([horse_name, row[1], row[2], row[3], row[6], row[7], row[8], str(row[11])], "td"))
    f.write("\n")

    f.close()


if __name__ == "__main__":
    xPOH_eval_list = gen_POH_eval_list(gen_HDN_eval_list(), gen_uma_eval_list())
    xPOH_eval_list.sort(key=lambda x:x[13], reverse=True)
    out_POH_eval_list(xPOH_eval_list)





