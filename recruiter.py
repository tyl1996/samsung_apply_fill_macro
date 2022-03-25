# -*- coding: utf-8 -*-
import time
import sys
import pandas as pd
import pyperclip
import yaml

import macro

def load_config():
    with open('config.yaml', encoding="utf-8") as f:
        config = yaml.safe_load(f)
    return config

arg = sys.argv[1]
print(arg)
config = load_config()
delay_time = config["delay_time"]
bot = macro.ChatMacro(delay_time)
data = pd.read_excel(arg, engine='openpyxl')
start_num = config["start_num"] #첫번째를 0으로 기준, 중간에 끊기면 이 숫자 조절
setup_delay_time = config["setup_delay_time"]
class_taken_year = config["class_taken_year"]
degree_type = config["degree_type"]

def next_form(bot):
    bot.press("tab")
    bot.press("tab")
    bot.press("enter")

time.sleep(setup_delay_time) #3초내로 과정 탭 누르고 엔터클릭
for index, row in data[start_num:].iterrows():

    class_name = row["과목명"] #과목명
    pyperclip.copy(class_name)
    bot.pressHoldRelease("ctrl", "v")
    bot.press("tab")

    class_point = row["취득학점"] #취득학점
    pyperclip.copy(class_point)
    bot.pressHoldRelease("ctrl", "v")
    bot.press("tab")

    class_grade = row["성적"] #성적
    grade_dict = {"A+": 4.5,
                  "A": 4,
                  "B+": 3.5,
                  "B": 3,
                  "C+": 2.5,
                  "C": 2,
                  "D+": 1.5,
                  "D": 1,
                  "F": 0,
                  "PASS": 'pass',
                  "P": 'pass',
                  "FAIL": 14,
                  "기타": 15}
    count = grade_dict[class_grade]
    if count == 'pass':
        pass
        bot.pressHoldRelease("shift", "tab")
        bot.pressHoldRelease("shift", "tab")
        bot.press("right_arrow")
        pyperclip.copy('('+count+')')
        bot.pressHoldRelease("ctrl", "v")
        bot.press("tab")
        bot.press("tab")
        pyperclip.copy('4.5')
        bot.pressHoldRelease("ctrl", "v")
    else:
        pass
        pyperclip.copy(count)
        bot.pressHoldRelease("ctrl", "v")
    bot.press("tab")
    for i in range(5):
        bot.press("down_arrow")

    bot.press("tab")
    bot.press("enter")
    bot.press("tab")
    bot.press("tab")

        
        