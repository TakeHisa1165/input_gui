# 名前を付けて保存

import xlwings as xw

class new_File_save:
    def __init__(self, wb, file_date):
        self.save_book = wb
        self.dir_path = "D:\\デスクトップ\\AF-M-184030検査成績書" + "\\"
        self.file_date = file_date

    def save_file(self):
        self.save_book.save()
        self.save_book.save(self.dir_path + "検査成績書AF-M-184030" + '(' + self.file_date + ').xlsm')


class File_save:
    def __init__(self, wb):
        wb.save()