from datetime import date, timedelta, datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import datetime
from openpyxl import load_workbook

def gerar_meses():
    now = datetime.datetime.now()
    start_month = datetime.datetime(now.year, now.month, 1)
    date_on_next_month = start_month + datetime.timedelta(35)
    start_next_month = datetime.datetime(date_on_next_month.year, date_on_next_month.month, 1)
    last_day_month = start_next_month - datetime.timedelta(1)

    delta = last_day_month - start_month   # returns timedelta

    day1 = start_month + timedelta(days=0)
    day2 = start_month + timedelta(days=1)
    day3 = start_month + timedelta(days=2)
    day4 = start_month + timedelta(days=3)
    day5 = start_month + timedelta(days=4)
    day6 = start_month + timedelta(days=5)
    day7 = start_month + timedelta(days=6)
    day8 = start_month + timedelta(days=7)
    day9 = start_month + timedelta(days=8)
    day10 = start_month + timedelta(days=9)
    day11 = start_month + timedelta(days=10)
    day12 = start_month + timedelta(days=11)
    day13 = start_month + timedelta(days=12)
    day14 = start_month + timedelta(days=13)
    day15 = start_month + timedelta(days=14)
    day16 = start_month + timedelta(days=15)
    day17 = start_month + timedelta(days=16)
    day18 = start_month + timedelta(days=17)
    day19 = start_month + timedelta(days=18)
    day20 = start_month + timedelta(days=19)
    day21 = start_month + timedelta(days=20)
    day22 = start_month + timedelta(days=21)
    day23 = start_month + timedelta(days=22)
    day24 = start_month + timedelta(days=23)
    day25 = start_month + timedelta(days=24)
    day26 = start_month + timedelta(days=25)
    day27 = start_month + timedelta(days=26)
    day28 = start_month + timedelta(days=27)
    day29 = start_month + timedelta(days=28)
    day30 = start_month + timedelta(days=29)
    day31 = start_month + timedelta(days=30)

    # for i in range(delta.days + 1):
    #     day = start_month + timedelta(days=i)
    #     print(day)

    wb = load_workbook(filename='C:/Users/rafaelvilela/Desktop/MEGAsync/Code/planilha/months.xlsx')
    sheet = wb.active
    sheet['B42'] = day1
    sheet['B43'] = day2
    sheet['B44'] = day3
    sheet['B45'] = day4
    sheet['B46'] = day5
    sheet['B47'] = day6
    sheet['B48'] = day7
    sheet['B49'] = day8
    sheet['B50'] = day9
    sheet['B51'] = day10
    sheet['B52'] = day11
    sheet['B53'] = day12
    sheet['B54'] = day13
    sheet['B55'] = day14
    sheet['B56'] = day15
    sheet['B57'] = day16
    sheet['B58'] = day17
    sheet['B59'] = day18
    sheet['B60'] = day19
    sheet['B61'] = day20
    sheet['B62'] = day21
    sheet['B63'] = day22
    sheet['B64'] = day23
    sheet['B65'] = day24
    sheet['B66'] = day25
    sheet['B67'] = day26
    sheet['B68'] = day27
    sheet['B69'] = day28
    sheet['B70'] = day29
    sheet['B71'] = day30
    sheet['B72'] = day31
    wb.save('months.xlsx')
