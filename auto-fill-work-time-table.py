import argparse
import sys
import xlrd
import random
import datetime
from shutil import copyfile
from xlutils.filter import process, XLRDReader, XLWTWriter


def copy2(wb):
    w = XLWTWriter()
    process(
        XLRDReader(wb, 'unknown.xls'),
        w
    )
    return w.output[0][1], w.style_list


class CellStructure:
    def __init__(self, date, dayOfWeek, row, enable):
        self.date = date
        self.dayOfWeek = dayOfWeek
        self.row = row
        self.startTime = ""
        self.endTime = ""
        self.enable = enable


def getStyle(wb, sheet, row, col):
    return wb.xf_list[sheet.cell(row, col).xf_index]


def getSheetStructure(excelFileName, year, month):
    dateCol = 0
    dayOfWeekCol = 1
    workTimeCol = 2
    startRow = 3
    wb = xlrd.open_workbook(excelFileName, formatting_info=True)
    sheetStructure = []
    # sheet = wb.get_sheet(1)
    sheet = wb.sheet_by_name(year + month)
    for row in range(startRow, sheet.nrows):
        if(sheet.row_values(row)[dateCol] == "小計"):
            break
        if(sheet.row_values(row)[dayOfWeekCol]):
            fmt = getStyle(wb, sheet, row, workTimeCol)
            cellStructure = CellStructure(
                sheet.row_values(row)[dateCol],
                sheet.row_values(row)[dayOfWeekCol],
                row,
                fmt.background.background_colour_index == 65)
            sheetStructure.append(cellStructure)
    return sheetStructure


def autoFillTable(excelFileName, year, month, startTimeStr, workHour, offsetRange, yourName):
    sheetStructure = getSheetStructure(excelFileName, year, month)
    baseStartHour = int(str.split(startTimeStr, ":")[0])
    baseStartMinute = int(str.split(startTimeStr, ":")[1])

    workStartTimeCol = 2
    workEndTimeCol = 3
    sourceWb = xlrd.open_workbook(excelFileName, formatting_info=True)
    sourceSheet = sourceWb.sheet_by_name(year + month)

    wb, styleList = copy2(sourceWb)
    sheet = wb.get_sheet(year + month)
    xfIndex = sourceSheet.cell_xf_index(1, 2)
    style = styleList[xfIndex]
    sheet.write(1, 2, yourName, style)

    global startTimeOffset
    global endTimeOffset
    for structure in sheetStructure:
        if(structure.enable):
            startHour = baseStartHour
            startOffset = int(baseStartMinute + (random.randrange(
                -offsetRange, offsetRange) + random.randrange(
                -offsetRange, offsetRange))/2)

            if(startOffset < 0):
                startOffset = 60 + startOffset
                startHour -= 1
            if(startOffset >= 60):
                startOffset = startOffset % 60
                startHour += 1
            startTime = datetime.time(startHour, startOffset)

            endHour = startHour + workHour + 1
            endOffset = int((random.randrange(
                startOffset, startOffset + offsetRange*1.2) + random.randrange(
                startOffset, startOffset + offsetRange*1.2))/2)

            if(endOffset < 0):
                endOffset = 60 + endOffset
                endHour -= 1
            if(endOffset >= 60):
                endOffset = endOffset % 60
                endHour += 1
            endTime = datetime.time(endHour, endOffset)

            structure.startTime = startTime.strftime("%H:%M")
            structure.endTime = endTime.strftime("%H:%M")

        xfIndex = sourceSheet.cell_xf_index(structure.row, workStartTimeCol)
        style = styleList[xfIndex]
        sheet.write(structure.row, workStartTimeCol,
                    structure.startTime, style)
        sheet.write(structure.row, workEndTimeCol,
                    structure.endTime, style)
    wb.save(excelFileName)


def main():
    parser = argparse.ArgumentParser(
        description='身為一個工程師，手填工時表是不被允許的: ')
    parser.add_argument("year")
    parser.add_argument("month", type=int)
    parser.add_argument("name")
    parser.add_argument("-s", "--startTime",
                        help="default 09:20", nargs="?", default="09:20")
    parser.add_argument("-w", "--workHour", help="default 8",
                        nargs="?", default="8", type=int)
    parser.add_argument("-o", "--offsetRange", help="default 10", nargs="?",
                        default="10", type=int)
    args = parser.parse_args()
    if args.month < 10:
        args.month = "0"+str(args.month)
    else:
        args.month = str(args.month)
    excelFileName = "姓名_{}年度工時表(範本).xls".format(args.year)
    autoFillTable(excelFileName, args.year, args.month,
                  args.startTime, int(args.workHour), int(args.offsetRange), args.name)


if __name__ == "__main__":
    main()
