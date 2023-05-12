import openpyxl
from openpyxl.styles import NamedStyle, PatternFill, Border, Side, Font, Alignment
import re, time
from deepdiff import DeepDiff
# Variable Input
workBookPath = r'../ComparisonCustLib.xlsx'
firstWsName = "Analyzed"
secondWsName = "Latest"

# Fill Input
wb = openpyxl.load_workbook(workBookPath)
ws1 = wb[firstWsName]
ws2 = wb[secondWsName]
wsDataSet1 = {
    "maxRows": len(ws1['B']),
    "columnA": 2, #FileName
    "columnB": 7, #RuleName
}
wsDataSet2 = {
    "maxRows": len(ws2['B']),
    "columnA": 2, #FileName
    "columnB": 6, #RuleName
}
dic1 = {}
dic2 = {}
def CreateDataSet(ws, wsDataSet):
    dic = {}
    flag = False
    firstRow = 0
    lastRow = 0
    for nRow in range(2, wsDataSet["maxRows"] + 1):
        cellValPrev = ws.cell(row=nRow - 1, column=wsDataSet["columnA"])
        cellValCurr = ws.cell(row=nRow, column=wsDataSet["columnA"])
        if cellValPrev.value == cellValCurr.value and not flag:
            flag = True
            firstRow = nRow - 1
        elif cellValPrev.value != cellValCurr.value:
            if flag:
                lastRow = nRow - 1
                flag = False
                if cellValPrev.value in dic:
                    dic[cellValPrev.value].append([firstRow, lastRow])
                else:
                    listAdd = [[firstRow, lastRow]]
                    dic[cellValPrev.value] = listAdd
            else:
                if cellValPrev.value in dic:
                    dic[cellValPrev.value].append([nRow - 1, nRow - 1])
                else:
                    listAdd = [[nRow - 1, nRow - 1]]
                    dic[cellValPrev.value] = listAdd
    # Add rules into the related files
    subDic = {}
    for key, value in dic.items():
        for elem in value:
            if elem[0] != elem[1]:
                for nRow2 in range(elem[0], elem[1] + 1):
                    cellVarB = ws.cell(row=nRow2, column=wsDataSet["columnB"])
                    if cellVarB.value not in subDic:
                        subDic.setdefault(cellVarB.value, 1)
                    else:
                        subDic[cellVarB.value] += 1
            else:
                nRow2 = elem[0]
                cellVarB = ws.cell(row=nRow2, column=wsDataSet["columnB"])
                subDic.setdefault(cellVarB.value, 1)
        dic[key] = subDic
        subDic = {}
    return dic

def CreateText(dic, filename):
    with open(f'{filename}.txt', 'w') as f:
        for x, y in dic.items():
            text = x + '\t' + ':' + str(y)
            f.write(text)
            f.write('\n')

def BorderCell(cell):
    pink = "00FF00FF"
    green = "00008000"
    black = "000e2e2a"
    thick = Side(border_style="thick", color=black)
    double = Side(border_style="double", color=green)
    cell.border = Border(top=thick, left=thick, right=thick, bottom=thick)

def FillBackGroundColor(cell):
    yellow = "00FFFF00"
    cell.fill = PatternFill(start_color=yellow, end_color=yellow, fill_type="solid")

def FontCell(cell):
    cell.font = Font(name="Arial", size=14, color="00FF0000")

def InsertData(data, ws):
    nCol = 1
    for key, value in data.items():
        if key == "dictionary_item_added":
            key = firstWsName + " removed"
        elif key == "dictionary_item_removed":
            key = firstWsName + " added"
        else:
            key = f"Values_changed : (new_value: {secondWsName}; old_value: {firstWsName})"
        cell = ws.cell(row=1, column=nCol, value=key)
        # Formatting cell
        BorderCell(cell)
        FillBackGroundColor(cell)
        FontCell(cell)
        value = str(value)
        value = value.replace('root', '')
        if not re.findall("value.*", key):
            value = value.split(',')
        else:
            value = value.split("},")
        nRow = 2
        for eachVal in value:
            cell = ws.cell(row=nRow, column=nCol, value=eachVal)
            BorderCell(cell)
            nRow += 1
        nCol += 1

if __name__ == "__main__":
    i = 0
    dic1 = CreateDataSet(ws1, wsDataSet1)
    dic2 = CreateDataSet(ws2, wsDataSet2)
    diff = DeepDiff(dic1, dic2)
    # CreateText(dic1, "J12s")
    # CreateText(dic2, "CustLibPlus")
    # with open(f'Diff.txt', 'w') as f:
    #     for x, y in diff.items():
    #         text = str(x) + '\t' + ':' + str(y)
    #         f.write(text)
    #         f.write('\n')
    try:
        ws3 = wb.create_sheet("Diff_CustLib")
        InsertData(diff, ws3)
        wb.save("../Result.xlsx")
        print("Saving the current workbook successfully")
    except PermissionError:
        print("Please close workbook before saving")

