# AF-M-184030　成績書入力　メインプログラム
import afm184030_gui
import input_to_inspectsheet
import xlwings as xw
import check_data
import new_file_save

wb = xw.Book(r"D:\デスクトップ\会社関係\検査成績書\検査成績書原紙\検査成績書AF-M-184030 (完全簡易）.xlsm")
ws = wb.sheets('検査成績(完全簡易）本社向け')
ws2 = wb.sheets('データ入力')
# GUI描写クラスのインスタンス
Gui_app = afm184030_gui.input_gui(wb, ws2)
value_list = Gui_app.Gui2()

# ロットの固まり数選択画面
# gui_choice = Gui_app.select_gui()
# guiの呼び出し


####### ロット数を選択する場合に使用する
# if gui_choice == "1":
#     value_list = Gui_app.Gui()
# else:
#     value_list = Gui_app.Gui2()

# guiで入力された値を表示
value_list["ws"] = ws2
print(value_list)

# ロットNo.から計算した数量と入力した数量を変数に格納
inp = input_to_inspectsheet.to_Inspectionsheet(**value_list)
inp.inpput_data()


# 計算した数量と入力した数量を比較して同じであるかを確認
# 別モジュールから計算した数量と入力した数量を取得
# inputdata class で作成
cal_qty = inp.cal_qty
qty = inp.qty
print("cal_qty = {}".format(cal_qty))
ck_data = check_data.check_Data(cal_qty, qty)
ck_data.check_qty()

# 保存用GUIの呼び出し
file_name_no = Gui_app.save_new_file()

# 保存用クラスのインスタンス
A_save = new_file_save.new_File_save(wb, file_name_no)
A_save.save_file()

