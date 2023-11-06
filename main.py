from random import (
    randint,
    uniform
)

import openpyxl as pyxl
from openpyxl import (
    Workbook,
    worksheet,
)
from typing import (
    TypedDict,
    Any,
    TypeVar,
    Generic,
    overload
)

import tkinter as tk
from tkinter import (
    # messagebox,
    filedialog
)

from functools import (
    singledispatch
)


def choice(n: list) -> int:
    return randint(0, len(n)-1)


TSubject = TypeVar('TSubject')
TWeight = TypeVar('TWeight')


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

    def set_loop_count(self, loop_count) -> None:
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
        donestr = []
        for _ in range(self.get_loop_count()):
            donestr.append(f'{nouns[choice(nouns)]} + {verbs[choice(verbs)]}')
        output_label['text'] = str.join('\n', donestr)


class MyApp(tk.Frame):
    def __init__(self, title: str, resolution: str) -> None:
        self.__root = tk.Tk()
        self.title = title
        self.resolution = resolution
        self.__root.title(self.title)
        self.__root.geometry(self.resolution)
        super().__init__(self.__root)

    def root(self) -> tk.Tk:
        return self.__root

    def loop(self) -> None:
        self.__root.mainloop()


if __name__ == "__main__":
    app = MyApp(title='アイデア出し', resolution='300x375')

    xlsx = filedialog.askopenfilename(
        title='*.xlsxを開く',
        filetypes=[('XLSX Worksheet', '*.xlsx')],
        initialdir='./'
    )
    excel = Excel(xlsx)

    name_label = tk.Label(app.root(), text='選択しているファイル: ' + excel.source_path)
    name_label.pack()

    loop_input = tk.Entry(app.root(), width=20)
    # loop_input.insert(tk.END, 20)
    loop_input.insert('end', 20)
    loop_input.pack()
    excel.set_loop_count(int(loop_input.get()))

    output_label = tk.Label()
    output_label.pack()
    run_btn = tk.Button(app.root(), text='生成', command=excel.generate_pair)
    run_btn.pack()

    app.loop()
