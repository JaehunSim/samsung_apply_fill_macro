# -*- coding: utf-8 -*-
import time

import pandas as pd
import pyperclip

import macro
delay_time = 0.25 #seconds
bot = macro.ChatMacro(delay_time)
data = pd.read_excel("data.xlsx")
start_num = 0 #첫번째를 0으로 기준, 중간에 끊기면 이 숫자 조절

def next_form(bot):
    bot.press("tab")
    bot.press("tab")
    bot.press("enter")

time.sleep(3) #3초내로 과정 탭 누르고 엔터클릭
for index, row in data[start_num:].iterrows():
    bot.press("enter") #과정
    bot.press("enter")
    next_form(bot)
    bot.press("enter") #전공명
    next_form(bot)
    year = row["수강년도"] #수강년도
    count = 2019 - year
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    next_form(bot)
    semester = row["학기"] #학기
    if isinstance(semester, int):
        count = semester - 1
    elif semester.startswith("여름"):
        count = 3
    elif semester.startswith("겨울"):
        count = 4
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    next_form(bot)
    class_type = row["과목유형"] #과목유형
    if class_type.startswith("교양"):
        count = 1
    elif class_type.startswith("전공"):
        count = 0
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    next_form(bot)
    is_retaken = row["재수강"] #재수강여부
    if not is_retaken:
        count = 1
    else:
        count = 0
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    bot.press("tab")
    bot.press("tab")
    class_name = row["과목명"] #과목명
    pyperclip.copy(class_name)
    bot.pressHoldRelease("ctrl", "v")
    bot.press("tab")
    bot.press("enter")
    class_point = row["취득학점"] #취득학점
    count = class_point-1
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    next_form(bot)
    class_grade = row["성적"] #성적
    grade_dict = {"A+": 0,
                  "A": 1,
                  "A-": 2,
                  "B+": 3,
                  "B": 4,
                  "B-": 5,
                  "C+": 6,
                  "C": 7,
                  "C-": 8,
                  "D+": 9,
                  "D": 10,
                  "D-": 11,
                  "F": 12,
                  "PASS": 13,
                  "P": 13,
                  "FAIL": 14,
                  "기타": 15}
    count = grade_dict[class_grade]
    for i in range(count):
        bot.press("down_arrow")
    bot.press("enter")
    next_form(bot) #추가버튼
    if index == 29 or index == 49:
        bot.press("enter") #중간 저장 팝업 
    time.sleep(0.2) #추가버튼 지연
    bot.press("tab")
    bot.press("enter") #내용삭제
    for i in range(17): #원 상태로 복귀
        bot.pressHoldRelease("shift", "tab")
        
        