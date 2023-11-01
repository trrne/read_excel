import random
import openpyxl

source_path = "ideas.xlsx"
file = openpyxl.load_workbook(source_path, read_only=True)
sheet = file["Sheet1"]
row = sheet.max_row

nouns = [""]
verbs = [""]
choice_nouns = [""]
choice_verbs = [""]


def read_nouns():
    for i in range(3, row, 1):
        nouns.append(sheet.cell(i, 1).value)
    return [n for n in nouns if n != '' or n != None]


def read_verbs():
    for i in range(3, row, 1):
        verbs.append(sheet.cell(i, 2).value)
    return [v for v in verbs if v != None]
    # return [v for v in verbs if v != '' or v != None]
    # vv = []
    # for j in range(len(verbs)):
    # if (len(j) >= 1):
    # vv.append(verbs[j])
    # return vv


LOOP_COUNT = 20
if __name__ == "__main__":
    n = read_nouns()
    v = read_verbs()
    for _ in range(LOOP_COUNT):
        print(read_nouns()[random.randint(0, len(read_nouns())-1)],
              "+", read_verbs()[random.randint(0, len(read_verbs())-1)])
        # nn = n[random.randint(0, )]
        # print()
