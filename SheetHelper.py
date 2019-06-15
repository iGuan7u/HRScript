from itertools import groupby

class SheetHelper:
    @staticmethod
    def getValueFromSheet(sheet, cellIndex):
        indexs = SheetHelper.cellColAndRowFromIndex(cellIndex)
        cellValue = sheet.cell_value(rowx=indexs[1], colx=indexs[0])
        print(cellValue)
        if type(cellValue) is str:
            cellValue = cellValue.replace('\n', '')        # if sheet.cell(rowx=indexs[1], colx=indexs[0]).ctype == 3: # 3 means 'xldate' , 1 means 'text'
        #     year, month, day, hour, minute, second = xlrd.xldate_as_tuple(cellValue, 
        #         0)
        #     print(year, month, day, hour, minute, second)
            # cellValue = datetime.datetime(year, month, day, hour, minute, nearest_second)
        return cellValue

    @staticmethod
    def cellColAndRowFromIndex(cellIndex):
        splitedIndex = [''.join(list(g)) for k, g in groupby(cellIndex, key=lambda x: x.isdigit())]
        col = splitedIndex[0]
        row = int(splitedIndex[1])
        if type(col) is not str:
            return [0, 0]

        colNum = 0
        power = 1

        for i in range(len(col)-1,-1,-1):
            ch = col[i]
            colNum += (ord(ch)-ord('A')+1)*power
            power *= 26
        return [colNum - 1, row - 1]

    @staticmethod
    def getValueFromBounds(sheet, startIndex, endIndex):
        startIndexs = SheetHelper.cellColAndRowFromIndex(startIndex)
        endIndexs = SheetHelper.cellColAndRowFromIndex(endIndex)
        values = []
        for leftIndex in range(startIndexs[1], endIndexs[1] + 1):
            subValues = []
            for rightIndex in range(startIndexs[0], endIndexs[0] + 1):
                cellValue = sheet.cell_value(rowx=leftIndex, colx=rightIndex)
                if type(cellValue) is str:
                    cellValue = cellValue.replace('\n', '')
                if cellValue == '' and rightIndex == startIndexs[0]:
                    return values
                if cellValue == '' and SheetHelper.isCellMerged(sheet, leftIndex, rightIndex):
                    continue
                subValues.append(cellValue)
            values.append(subValues)
        return values

    @staticmethod
    def isCellMerged(sheet, row, col):
        for crange in sheet.merged_cells:
            rlo, rhi, clo, chi = crange
            for rowx in range(rlo, rhi):
                for colx in range(clo, chi):
                    if rowx == row and colx == col:
                        return True
        return False