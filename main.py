from random import randint

import openpyxl as pyxl
from openpyxl import (
    Workbook,
    worksheet,
)
from typing import (
    TypedDict,
    Any
)

import tkinter as tk
from tkinter import (
    # messagebox,
    filedialog
)


def ice(n: list) -> int:
    return randint(0, len(n)-1)


class IdeasListDict(TypedDict):
    nouns: list
    verbs: list


class Excel:
    def __init__(self, source_path: str) -> None:
        self.source_path: str = source_path
        self.file: Workbook
        self.sheet: worksheet
        self.row: int
        self.column: int
        self.count: int
        if self.source_path != None:
            self.file = pyxl.load_workbook(self.source_path)
            self.sheet = self.file["Sheet1"]
            self.row = self.sheet.max_row
            self.column = self.sheet.max_column

    def get_source_path(self) -> str:
        return self.source_path

    def get_row(self) -> int:
        return self.row

    def get_cell_value(self, row: int, column: int) -> Any:
        return self.sheet.cell(row, column).value

    def set_loop_count(self, loop_count):
        self.count = loop_count

    def get_loop_count(self) -> int:
        return self.count

    def read_data(self) -> IdeasListDict:
        (ns, vs) = ([], [])
        for i in range(3, self.get_row(), 1):
            ns.append(self.get_cell_value(i, 1))
            vs.append(self.get_cell_value(i, 2))
        return IdeasListDict(
            nouns=[n for n in ns if n != None], verbs=[v for v in vs if v != None])

    # @staticmethod
    def generate_pair(self) -> None:
        rdata: IdeasListDict = self.read_data()
        (nouns, verbs) = (rdata['nouns'], rdata['verbs'])
        (done_n, done_v) = ([], [])
        donestr = []
        # d:  IdeasListDict
        for _ in range(self.get_loop_count()):
            # done_n.append(nouns[ice(nouns)])
            # done_v.append(verbs[ice(verbs)])
            donestr.append(f'{nouns[ice(nouns)]} + {verbs[ice(verbs)]}')
            # done_n.append(randint(0, ice(nouns)))
            # done_v.append(randint(0, ice(verbs)))
            # d = IdeasListDict(
            # nouns=[randint(0, ice(nouns))], verbs=[randint(0, ice(verbs))])
        # donestr  # len(done_n), len(done_v)  # data
        output_label['text'] = str.join('\n', donestr)

    # def output(self, idea: IdeasListDict) -> IdeasListDict:
    #     read_data = idea
    #     (nouns, verbs) = (read_data['nouns'], read_data['verbs'])
    #     data: IdeasListDict
    #     for _ in range(self.get_loop_count()):
    #         data = IdeasListDict(
    #             noun_s=nouns[randint(0, ice(nouns))], verb_s=verbs[randint(0, ice(verbs))])
    #     return data


class App:
    def __init__(self, title: str, resolution: str) -> None:
        self.__root = tk.Tk()
        self.title = title
        self.resolution = resolution
        self.__root.title(self.title)
        self.__root.geometry(self.resolution)

    def root(self) -> tk.Tk:
        return self.__root

    def loop(self) -> None:
        self.__root.mainloop()


if __name__ == "__main__":
    # root = tk.Tk()
    # root.title('アイデア出し')
    # root.geometry('800x600')

    app = App(title='アイデア出し', resolution='800x600')

    xlsx = filedialog.askopenfilename(
        title='*.xlsxを開く',
        filetypes=[('XLSX Worksheet', '*.xlsx')],
        initialdir='./'
    )
    excel = Excel(xlsx)

    # name_label = tk.Label(root, text='選択しているファイル: ' + excel.source_path)
    name_label = tk.Label(app.root(), text='選択しているファイル: ' + excel.source_path)
    name_label.pack()

    # loop_input = tk.Entry(root)
    loop_input = tk.Entry(app.root())
    loop_input.insert(tk.END, 20)
    loop_input.pack()
    excel.set_loop_count(20)

    output_label = tk.Label()
    output_label.pack()
    # run_btn = tk.Button(root, text='生成', command=excel.generate_pair)
    run_btn = tk.Button(app.root(), text='生成', command=excel.generate_pair)
    run_btn.pack()

    # root.mainloop()
    app.loop()
