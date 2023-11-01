from random import randint
import openpyxl
import typing
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


def ice(n: list) -> int:
    return len(n)-1


class IdeasListDict(typing.TypedDict):
    ns: list
    vs: list


class Excel:
    def __init__(self, source_path: str) -> None:
        self.source_path: str = source_path
        self.file: openpyxl.Workbook
        self.sheet: openpyxl.openpyxl.worksheet.worksheet
        self.row: int
        self.column: int
        self.loop_count: int
        if self.source_path != None:
            self.file = openpyxl.load_workbook(self.source_path)
            self.sheet = self.file["Sheet1"]
            self.row = self.sheet.max_row
            self.column = self.sheet.max_column

    def get_source_path(self) -> str:
        return self.source_path

    def get_row(self) -> int:
        return self.row

    def get_row_column(self) -> tuple[int, int]:
        return (self.row, self.column)

    def get_cell_value(self, row: int, column: int):
        return self.sheet.cell(row, column).value

    def set_loop_count(self, loop_count) -> None:
        self.loop_count = loop_count

    def read_data(self) -> IdeasListDict:
        (n, v) = ([], [])
        for i in range(3, self.get_row(), 1):
            n.append(self.get_cell_value(i, 1))
            v.append(self.get_cell_value(i, 2))
        return IdeasListDict(
            ns=[nouns for nouns in n if nouns != None],
            vs=[verb for verb in v if verb != None]
        )

    # @staticmethod
    def generate_pair(self):
        rdata = self.read_data()
        output_label['text'] = self.output(rdata)

    def output(self, idea: IdeasListDict):
        read_data = idea
        (nouns, verbs) = (read_data['ns'], read_data['vs'])
        data: IdeasListDict
        for _ in range(LOOP_COUNT):
            data = IdeasListDict(
                ns=nouns[randint(0, ice(nouns))],
                vs=verbs[randint(0, ice(verbs))]
            )
        return data


LOOP_COUNT = 20
if __name__ == "__main__":
    root = tk.Tk()
    root.title('アイデア出し')
    root.geometry('800x600')

    xlsx = filedialog.askopenfilename(
        title='*.xlsxを開く',
        filetypes=[('XLSX Worksheet', '*.xlsx')],
        initialdir='./'
    )
    excel = Excel(xlsx)

    name_label = tk.Label(root, text='選択しているファイル: ' + excel.source_path)
    name_label.pack()

    loop_input = tk.Entry(root)
    loop_input.insert(tk.END, 20)
    loop_input.pack()

    output_label = tk.Label()
    output_label.pack()
    run_btn = tk.Button(root, text='生成', command=excel.generate_pair)
    run_btn.pack()

    root.mainloop()
