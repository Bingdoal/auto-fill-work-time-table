import xlrd
import random
import datetime
from shutil import copyfile


class CellStructure:
    def __init__(self, date, dayOfWeek, row):
        self.date = date
        self.dayOfWeek = dayOfWeek
        self.row = row
        self.startTime = ""
        self.endTime = ""


def getSheetStructure(excelFileName):
    dateCol = 0
    dayOfWeekCol = 1
    startRow = 3
    wb = xlrd.open_workbook(excelFileName)
    sheetStructure = []
    # sheet = wb.get_sheet(0)
    sheet = wb.sheet_by_index(0)
    for row in range(startRow, sheet.nrows):
        if(sheet.row_values(row)[dateCol] == "小計"):
            break
        if(sheet.row_values(row)[dayOfWeekCol]):
            cellStructure = CellStructure(
                sheet.row_values(row)[dateCol],
                sheet.row_values(row)[dayOfWeekCol],
                row)
            sheetStructure.append(cellStructure)
    return sheetStructure


def autoFillTable(excelFileName, outputFileName):
    sheetStructure = getSheetStructure(excelFileName)
    baseStartHour = 9
    baseEndHour = 18
    tempFile = open("tempData.csv", "w")
    for structure in sheetStructure:
        if(structure.dayOfWeek in ["一", "二", "三", "四", "五"]):
            startHour = baseStartHour
            startOffset = random.randrange(-10, 20)
            endHour = baseEndHour
            endOffset = random.randrange(startOffset + 10, startOffset + 20)

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
        tempFile.write("{}, {}, {}\n".format(
            structure.date, structure.startTime, structure.endTime))
    tempFile.close()

def main():
    yourName = "林宜霆"
    year = "109"
    month = "11"
    excelFileName = "姓名_{}{}工時表(範本).xls".format(year, month)
    outputFileName = "{}_{}{}工時表.xls".format(yourName, year, month)
    copyfile(excelFileName, outputFileName)
    autoFillTable(excelFileName, outputFileName)


if __name__ == "__main__":
    main()
