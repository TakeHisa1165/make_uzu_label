import xlwings as xw
import PySimpleGUI as sg
import win32com.client

class InputExcel:
    def __init__(self, wb, ws, start_no, no_of_labels, part_no, path):
        self.label_book = wb
        self.label_sheet = ws
        self.start_no = int(start_no)
        self.no_of_labels = int(no_of_labels)
        self.part_no = part_no
        self.path = path
        self.input_data()


    def input_data(self):
        self.xlapp = win32com.client.Dispatch("Excel.Application")
        wb = self.xlapp.Workbooks.Open(self.path)
        ws = wb.Worksheets('ラベル印刷')
        ws.Unprotect()

        sheet = self.label_book.sheets('コピー元')
        for i in range(1, 93):
            sheet.range(i,1).value = 0
        for i in range(self.start_no, self.start_no + self.no_of_labels):
            sheet.range(i, 1).value = self.part_no

        ws.Protect()
        sg.popup('ラベル記入が完了しました。\n印刷してください。')








        # イコールで結ばない方法　模索中
        # if self.start_no % 4 != 0:
        #     start_no_row = (self.start_no // 4) + 1
        #     cal_row = (start_no_row - 1)
        #     start_no_row = start_no_row + cal_row
        #     start_no_col = (self.start_no % 4)
        #     if start_no_col == 2:
        #         start_no_col = start_no_col + 1

        #     elif start_no_col == 3:
        #         start_no_col = start_no_col + 2
        #     else:
        #         pass
        # elif self.start_no % 4 == 0:
        #     start_no_row = (self.start_no // 4)
        #     cal_row = (start_no_row - 1)
        #     start_no_row = start_no_row + cal_row
        #     start_no_col = (self.start_no % 4)
        #     start_no_col = 7

        # self.label_sheet.range(start_no_row, start_no_col).value = self.part_no

# class PrintOut:
#     def __init__(self, path):
#         self.xlapp = win32com.client.Dispatch("Excel.Application")
#         self.path = path
#         self.print_out()
#         PRINTER_NAME = "CubePDF"
#     def print_out(self):
#         wb = self.xlapp.Workbooks.Open(self.path)
#         ws = wb.Worksheets[0]
#         ws.printout





