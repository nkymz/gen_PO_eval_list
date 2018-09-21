# -*- coding: utf-8 -*-

import os
import re
import datetime
import time

import openpyxl
import requests
from bs4 import BeautifulSoup
import pprint

def gen_POH_eval_list(HDN_eval_list):
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

    wb.save(wbpath)

    xPOH_eval_list_all = [[cell.value for cell in row] for row in ws["A1:J" + str(xlrow - 1)]]
    return [row for row in xPOH_eval_list_all if row[5] != "HDN_eval_none"]


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

def wrap_trtd(columns):
    t = "|"
    for s in columns:
        t = t + s + "|"
    t += ">\n"

    return t

    # t = "<tr>"
    # for s in columns:
    #     t = t + "<td>" + s + "</td>"
    # t += "</tr>\n"
    #
    # return t


def out_POH_eval_list(xPOH_eval_list):
    path = os.getenv("HOMEDRIVE", "None") + os.getenv("HOMEPATH", "None") + "/Dropbox/POG/"
    htmlpath = path + "POH_eval_list.html"
    f = open(htmlpath, mode="w", encoding="utf-8")
    # f.write('<table border="1" cellspacing="0" cellpadding="5" bordercolor="#333333">\n')
    f.write("|-|-|-|-|-|-|-|")
    f.write(wrap_trtd(["馬名", "オーナー", "性別", "指名順", "HDN1", "HDN2", "HDN3"]))

    for row in xPOH_eval_list:
        f.write(wrap_trtd([row[0], row[1], row[2], row[3], row[6], row[7], row[8]]))

    f.write("\n")

    f.close()

if __name__ == "__main__":
    xPOH_eval_list = gen_POH_eval_list(gen_HDN_eval_list())
    xPOH_eval_list.sort(key=lambda x:x[9], reverse=True)
    out_POH_eval_list(xPOH_eval_list)


    # wb.save(wbpath)



