from tkinter import messagebox
import tkinter as tk
import sys
import PySimpleGUI as sg



class check_Data:
    def __init__(self, qty, cal_qty):
        self.qty = qty
        self.cal_qty = cal_qty

    def check_qty(self):
        if str(self.qty) != self.cal_qty:
            sg.theme('Bluemono')
            sg.popup_error("入力された出荷数とロット番号から計算した出荷数がちがいます。",  title="数量エラー")
            sys.exit()
            
        else:
            pass

if __name__ == "__main__":
    a = check_Data(qty=1000, cal_qty=2000)
    a.check_qty()
