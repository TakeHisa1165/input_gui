import PySimpleGUI as sg

sg.theme('dark')

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
        print("ファイル名は{}".format(fn))

window.close()

