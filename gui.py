import PySimpleGUI as sg
import xlwings as xw
import input_to_excel
import os
import csv
import w_csv
import sys
import win32com.client


class InputWindow:
    def __init__(self):
        #　ラベル作成エクセルブックを開く
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
        try:
            with open('path.csv', newline='') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    self.dir_path = row["dir_path"]

        except FileNotFoundError:
            sg.popup_ok('初期設定が必要です。\n設定画面から書き出しフォルダを設定してください。')
            SelectFile()

        try:
            self.label_file_path = self.dir_path
        except AttributeError:
            sg.popup_error('Excelファイルを開けません\n初期設定をやり直してください。')
            SelectFile()
        self.label_book = xw.Book(self.label_file_path)
        self.label_sheet = self.label_book.sheets("ラベル印刷")
        self.part_no_sheet = self.label_book.sheets('品番リスト')
        self.part_no_list = []
        self.item_dict = {}
        self.item_max_col = self.part_no_sheet.range(2, 10000).end('left').column
        # 納入先項目を辞書に格納　項目名と列番号
        for i in range(2, self.item_max_col + 1):
            self.item_dict[str(self.part_no_sheet.range(2, i).value)] = i

        print(self.item_dict)
        # 納入先リストを格納
        self.to_ship_Max_row = self.part_no_sheet.range(10000, 2).end("up").row
        self.to_ship_list =self.part_no_sheet.range((3, 2), (self.to_ship_Max_row, 2)).value




    def input_window(self):
        sg.theme("systemdefault")

        frame1 = [
            [sg.Text(text=('納入先を選んでください'),font=('メイリオ', 14)),],
            [sg.Listbox(self.to_ship_list, size=(15,3), key="-to_ship-", font=('メイリオ', 14)),
            sg.Listbox(self.part_no_list, size=(15,3),key="-part_no-", font=('メイリオ', 14))],
            [sg.Submit(button_text="納入先選択", size=(10,1), font=('メイリオ', 14), pad=((30,85),(0,0))),sg.Submit(button_text="品番選択", size=(10,1), font=('メイリオ', 14))],
        ]

        frame2 = [
            [sg.Text("選択した納入先", font=('メイリオ', 14), pad=((40, 130),(0,0))), sg.Text("選択した品番", font=('メイリオ', 14))],
            [sg.InputText(size=(20,1), key='-select_to_ship-', font=('メイリオ', 14)), sg.InputText(size=(20,1), key='-select_part_no-', font=('メイリオ', 14))]
        ]

        frame3 = [
            [sg.Text("ラベル作成開始位置", font=('メイリオ', 14), pad=((5, 80),(0,0))), sg.Text("必要枚数", font=('メイリオ', 14))],
            [sg.InputText(size=(15,1), key="-start_no-", font=('メイリオ', 14), pad=((5,30),(0,0))), sg.InputText(size=(15,1), key="-no_of_labels", font=('メイリオ', 14))],
        ]

        layout = [
            [sg.MenuBar([["設定",["フォルダ設定"]]], key="menu1")],
            [sg.Frame("品番選択", frame1,)],
            [sg.Frame('選択結果表示', frame2)],
            [sg.Frame("ラベル枚数設定", frame3)],
            [sg.Submit(button_text=("ラベル作成"), font=("メイリオ", 14),pad=((5, 200),(0,0))), sg.Submit(button_text="終了する", font=("メイリオ", 14))],
            ]


        window = sg.Window('ラベル作成', layout)

        while True:
            event, values = window.read()

            if event is None:
                print(exit)
                break

            if event == "納入先選択":
                name = values["-to_ship-"]
                print(name)
                try:
                    col = self.item_dict[name[0]]
                    part_no_Max_row = self.part_no_sheet.range(10000, col).end('up').row
                    self.part_no_list = self.part_no_sheet.range((3, col),(part_no_Max_row, col)).value
                    window["-part_no-"].update(self.part_no_list)
                except IndexError:
                    sg.popup_error('納入先を選択してください')

            if event == "品番選択":
                to_ship = values["-to_ship-"]
                part_no = values["-part_no-"]
                window['-select_to_ship-'].update(to_ship)
                window['-select_part_no-'].update(part_no)
                print(part_no)

            if event == "ラベル作成":
                part_no = values["-part_no-"]
                start_no = values["-start_no-"]
                no_of_labels = values["-no_of_labels"]
                try:
                    input_to_excel.InputExcel(wb=self.label_book, ws=self.label_sheet,
                                                start_no=start_no, no_of_labels=no_of_labels,
                                                part_no=part_no)

                except ValueError:
                    sg.popup_error("納入先、品番が選択されているか確認してください")

            if event == "終了する":
                sys.exit()
            # if event == "印刷":
            #     input_to_excel.PrintOut(self.label_file_path)

            if values["menu1"] == "フォルダ設定":
                SelectFile()


        window.close()

class SelectFile:
    def __init__(self):
        self.path_dict = self.select_file()

    def select_file(self):

        sg.theme("systemdefault")

        layout = [
            [sg.Text("ラベル作成ファイルを選んでください", size=(50, 1), font=('メイリオ', 14))], 
            [sg.InputText(font=('メイリオ', 14)),sg.FileBrowse('開く', key='File1', font=('メイリオ', 14))],
            [sg.Submit(button_text='設定', font=('メイリオ', 14)), sg.Submit(button_text="閉じる", font=('メイリオ', 14))]
        ]

        # セクション 2 - ウィンドウの生成z
        window = sg.Window('ファイル選択', layout)

        # セクション 3 - イベントループ
        while True:
            event, values = window.read()

            if event is None:
                print('exit')
                break

            if event == '設定':
                path_dict = {}
                dir_path = values[0]
                path_dict["dir_path"] = dir_path
                csv = w_csv.Write_csv()
                csv.write_csv(path_dict=path_dict)
                sg.popup('初期設定が完了しましたアプリを再起動してください\nアプリを終了します')
                sys.exit()


                return path_dict
            if event == '終了する':
                sys.exit()




        #  セクション 4 - ウィンドウの破棄と終了
        window.close()

if __name__ == "__main__":
    app = InputWindow()
    app.input_window()