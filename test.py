"""
検査成績書入力用プログラム　
品番選択GUI表示
"""

import subprocess
import PySimpleGUI as sg

sg.theme('green')

layout = [
    [sg.Submit(button_text="AF-M-184030 NKメカ", size=(20,1)), sg.Submit(button_text='020-191ZZ791', size=(20,1))]
]

window = sg.Window('品番選択', layout)

while True:
    event, values = window.read()

    if event is None:
        print('終了します')
        break

    if event == "AF-M-184030 NKメカ":
        subprocess.run("D:\ドキュメント\python\inspection_sheet\input_gui\dist\input_afm184030\input_afm184030.exe")

        break

window.close