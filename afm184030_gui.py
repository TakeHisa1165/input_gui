import PySimpleGUI as sg
import new_file_save
import sys


class input_gui:
    
    def __init__(self, wb, ws):
        self.wb = wb
        self.ws = ws
        self.d_inspect_day = ws.range("B1").value
        self.d_ship_day = ws.range("B2").value


    # テーマ設定
    def Gui(self):
        sg.theme('Bluemono')

        frame1 = [
            [sg.Radio('成績書を送る', 1, key="送る")],
            [sg.Radio('成績書を送らない', 1, key="送らない")],
        ]
        frame2 =[
            [sg.Submit(button_text="実行", size=(12,2))],
        ]
        frame3 =[
            [sg.Submit(button_text="途中保存", size=(12,2))]
        ]

        layout = [
            [sg.Text("検査日", size=(12, 1)), sg.InputText()],
            [sg.Text("出荷日", size=(12, 1)), sg.InputText()],
            [sg.Text("出荷数", size=(12, 1)), sg.InputText()],
            [sg.Text('指示書番号', size=(12, 1)), sg.InputText()],
            [sg.Text("引き当て番号", size=(12, 1)), sg.InputText()],
            [sg.Text("客先納期", size=(12, 1)), sg.InputText()],
            [sg.Text("設備番号、西暦", size=(12,1)), sg.InputText(size=(10,1)),sg.Text("ロット番号", size=(8, 1)), sg.InputText(size=(10, 1)),
                sg.Text("～", size=(1, 1)),sg.InputText(size=(10, 1))],
            [sg.Text("青島入荷日", size=(12, 1)), sg.InputText()],
            [sg.Text("測定日", size=(12, 1)), sg.InputText()],
            [sg.Frame("成績書送付", frame1),sg.Frame("入力",frame2),sg.Frame("途中保存", frame3)],
        ]


        window = sg.Window("AF-M-184030  成績書情報入力", layout)

        while True:
            event, values = window.read()
            if event is None:
                print('exit')
                break

            if event == "実行":
                value_dict = {}
                inspect_day = values[0]
                value_dict["inspect_day"] = inspect_day
                ship_day = values[1]
                value_dict["ship_day"] = ship_day
                qty = values[2]
                value_dict["qty"] = qty
                order_no = values[3]
                value_dict["order_no"] = order_no
                order_no2 = values[4]
                value_dict["order_no2"] = order_no2
                delivery_day = values[5]
                value_dict["delivery_day"] = delivery_day
                machine_no1 = values[6]
                value_dict["machine_no1"] = machine_no1
                lot1 = values[7]
                value_dict["lot1"] = lot1
                lot2 = values[8]
                value_dict["lot2"] = lot2
                value_dict["machine_no2"] = "-"
                value_dict["lot3"] = "-"
                value_dict["lot4"] = "-"
                arrivale_day = values[9]
                value_dict["arrivale_day"] = arrivale_day
                measure_day = values[10]
                value_dict["measure_day"] = measure_day

          

                # print(inspect_day)

                # sg.popup(inspect_day)
   
            
                if values["送る"]:
                    value_dict["send"] = "送る"
                
                elif values["送らない"]:
                    value_dict["send"] = '送らない'
                

                return value_dict
            
            if event == '途中保存':
                print('test')
                new_file_save.File_save(self.wb)


###########################################################################################################
# 保存画面
    def save_new_file(self):
        sg.theme('Bluemono')

        layout = [
            [sg.Text("ファイル名を入力してください　例)200318"), sg.InputText(size=(10,1))],
            [sg.Submit(button_text="保存", size=(10,1))],
        ]

        window = sg.Window('名前を付けて保存', layout)

        while True:
            event, values = window.read()
            if event is None:
                print('終了します')
                break

            if event == "保存":
                fn = values[0]
                return fn

        window.close()


######################################################################################################
# ロットの固まり数選択画面
    def select_gui(self):
        sg.theme('Bluemono')

        frame1 = [
            [sg.Radio('ロット1個', 1, key="1")],
            [sg.Radio('ロット2個', 1, key="2")],
        ]




        layout = [
            [sg.Text("ロットの固まり数を選択")],
            [sg.Frame('選択してください', frame1)],
            [sg.Submit('開く', size=(10,1))]
        ]

        window = sg.Window('選択してください', layout)

        while True:
            event, values = window.read()
            if event is None:
                print('終了します')
                sys.exit()

            if event == "開く":
                if values["1"]:
                    return "1"

                
                elif values["2"]:
                    return "2"


        window.close()


########################################################################################
# ロット2個バージョン

    def Gui2(self):
        sg.theme('Bluemono')

        frame1 = [
            [sg.Radio('成績書を送る', 1, key="送る")],
            [sg.Radio('成績書を送らない', 1, key="送らない")],
        ]
        frame2 =[
            [sg.Submit(button_text="実行", size=(12,2))],
        ]

        frame3 =[
            [sg.Submit(button_text="途中保存", size=(12,2))]
        ]

        layout = [
            [sg.Text('　　　空欄の場合は　- ハイフンを入力してください', size=(40,1), font=("メイリオ",12), text_color=("#ff0000"))],
            [sg.Text("検査日", size=(12, 1)), sg.InputText()],
            [sg.Text("出荷日", size=(12, 1)), sg.InputText()],
            [sg.Text("出荷数", size=(12, 1)), sg.InputText()],
            [sg.Text('指示書番号', size=(12, 1)), sg.InputText()],
            [sg.Text("引き当て番号", size=(12, 1)), sg.InputText()],
            [sg.Text("客先納期", size=(12, 1)), sg.InputText()],
            [sg.Text("設備番号、西暦", size=(12,1)), sg.InputText(size=(10,1)),sg.Text("ロット番号", size=(8, 1)), sg.InputText(size=(10, 1)),
                sg.Text("～", size=(1, 1)),sg.InputText(size=(10, 1))],
            [sg.Text("設備番号、西暦", size=(12,1)), sg.InputText(size=(10,1)),sg.Text("ロット番号", size=(8, 1)), sg.InputText(size=(10, 1)), 
                sg.Text("～", size=(1, 1)),sg.InputText(size=(10, 1))],
            [sg.Text("青島入荷日", size=(12, 1)), sg.InputText()],
            [sg.Text("測定日", size=(12, 1)), sg.InputText()],
            [sg.Text('入り数➀', size=(10,1)), sg.InputText(size=(10,1), key="box1"),
                sg.Text("入り数➁", size=(12,1)), sg.InputText(size=(10,1), key="box2")],
            [sg.Frame("成績書送付", frame1),sg.Frame("入力",frame2),sg.Frame("途中保存", frame3)],
        ]


        window = sg.Window("AF-M-184030  成績書情報入力", layout)

        while True:
            event, values = window.read()
            if event is None:
                print('exit')
                break

            if event == "実行":
                value_dict = {}
                inspect_day = values[0]
                value_dict["inspect_day"] = inspect_day
                ship_day = values[1]
                value_dict["ship_day"] = ship_day
                qty = values[2]
                value_dict["qty"] = qty
                order_no = values[3]
                value_dict["order_no"] = order_no
                order_no2 = values[4]
                value_dict["order_no2"] = order_no2
                delivery_day = values[5]
                value_dict["delivery_day"] = delivery_day
                machine_no1 = values[6]
                value_dict["machine_no1"] = machine_no1
                lot1 = values[7]
                value_dict["lot1"] = lot1
                lot2 = values[8]
                value_dict["lot2"] = lot2
                machine_no2 = values[9]
                value_dict["machine_no2"] = machine_no2
                lot3 = values[10]
                value_dict["lot3"] = lot3
                lot4 = values[11]
                value_dict["lot4"] = lot4
                arrivale_day = values[12]
                value_dict["arrivale_day"] = arrivale_day
                measure_day = values[13]
                value_dict["measure_day"] = measure_day
                box1 = values["box1"]
                value_dict["box1"] = box1
                box2 = values["box2"]
                value_dict["box2"] = box2

            

                # print(inspect_day)

                # sg.popup(inspect_day)
   
            
                if values["送る"]:
                    value_dict["send"] = "送る"
                
                elif values["送らない"]:
                    value_dict["send"] = '送らない'
                

                return value_dict






if __name__ == "__main__":
    pass



