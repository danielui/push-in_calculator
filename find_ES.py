from openpyxl import load_workbook


def Find_ES(d, holeTolerance):
    wb = load_workbook(
        filename=r"C:\Users\user\VisualStudioCode\push-in_calculations\tolerances.xlsx")
    sheet = wb["1"]
    cells = sheet['A1': 'L53']
    switch = {'H6': 'D', 'H7': 'E', 'H8': 'F', 'H9': 'G',
              'H10': 'H', 'H11': 'I', 'H12': 'J', 'H13': 'K', 'H14': 'L'}

    for i in range(5, 55):
        if d <= int(sheet["A" + str(i)].value):
            row = i-2
            return (sheet[str(switch[str(holeTolerance)]) + str(row)].value), i-1
            break
