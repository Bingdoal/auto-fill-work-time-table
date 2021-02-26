import xlrd
import random
import datetime
from shutil import copyfile
from xlutils.filter import process, XLRDReader, XLWTWriter

yourName = "林宜霆"
year = "110"
month = "02"
startTimeOffset = 10
endTimeOffset = 10


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


def autoFillTable(excelFileName, outputFileName, year, month):
    sheetStructure = getSheetStructure(excelFileName, year, month)
    baseStartHour = 9
    baseEndHour = 18

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
            startOffset = random.randrange(-startTimeOffset, startTimeOffset)
            endHour = baseEndHour
            endOffset = random.randrange(
                startOffset + endTimeOffset, startOffset + endTimeOffset*2)

            if(startOffset < 0):
                startOffset = 60 + startOffset
                startHour -= 1
            startTime = datetime.time(startHour, startOffset)

            if(endOffset < 0):
                endOffset = 60 + endOffset
                endHour -= 1
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
    wb.save(outputFileName)


def main():
    excelFileName = "姓名_{}年度工時表(範本).xls".format(year)
    outputFileName = "{}_{}年度工時表.xls".format(yourName, year)
    autoFillTable(excelFileName, outputFileName, year, month)


if __name__ == "__main__":
    main()
