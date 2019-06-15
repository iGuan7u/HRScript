from itertools import groupby

class SheetHelper:
    @staticmethod
    def writeTheHeader(sheet):
        sheet.write(0, 0, '序号')
        sheet.write(0, 1, '公司')
        sheet.write(0, 2, '中心')
        sheet.write(0, 3, '部门')
        sheet.write(0, 4, '科')
        sheet.write(0, 5, '组')
        sheet.write(0, 6, '工作地点')
        sheet.write(0, 7, '一级主管')
        sheet.write(0, 8, '姓名')
        sheet.write(0, 9, 'OA用户名')
        sheet.write(0, 10, '英文名')
        sheet.write(0, 11, '在职状态')
        sheet.write(0, 12, '民族')
        sheet.write(0, 13, '政治面貌')
        sheet.write(0, 14, '性别')
        sheet.write(0, 15, '籍贯')
        sheet.write(0, 16, '婚姻状况')
        sheet.write(0, 17, '出生年月')
        sheet.write(0, 18, '最高学历')
        sheet.write(0, 19, '兴趣爱好')
        sheet.write(0, 20, '工作Email')
        sheet.write(0, 21, '入职时间')
        sheet.write(0, 22, '特长')
        sheet.write(0, 23, '身份证号码')
        sheet.write(0, 24, '户口所在地详细地址')
        sheet.write(0, 25, '户口性质')
        sheet.write(0, 26, '手机')
        sheet.write(0, 27, '家庭电话')
        sheet.write(0, 28, '现在住址')
        sheet.write(0, 29, '教育经历')
        sheet.write(0, 30, '培训经历')
        sheet.write(0, 31, '专业资格证书及获取时间')
        sheet.write(0, 32, '工作履历')
        sheet.write(0, 33, '家庭成员')
        sheet.write(0, 34, '紧急联系人')
        sheet.write(0, 35, '工资卡号')
        sheet.write(0, 36, '开户地')
        sheet.write(0, 37, '递交个人资料')
        sheet.write(0, 38, '参加工作时间')
        sheet.write(0, 39, '员工类型')
        sheet.write(0, 40, '岗位')
        sheet.write(0, 41, '岗位级别')
        sheet.write(0, 42, '职级')
        sheet.write(0, 43, '个人Email')
        sheet.write(0, 44, '合同期开始')
        sheet.write(0, 45, '合同期结束')
        sheet.write(0, 46, '合同到期提醒日期')
        sheet.write(0, 47, '试用期')
        sheet.write(0, 48, '转正日期')
        sheet.write(0, 49, '离职日期')
        sheet.write(0, 50, '离职原因')
        sheet.write(0, 51, '公积金购买日期')
        sheet.write(0, 52, '公积金比例')
        sheet.write(0, 53, '公积金类型')
        sheet.write(0, 54, '公积金封存日期')
        sheet.write(0, 55, '公积金基数')
        sheet.write(0, 56, '公积金号')
        sheet.write(0, 57, '社保增员日期')
        sheet.write(0, 58, '是否当地社保新开户')
        sheet.write(0, 59, '社保减员日期')
        sheet.write(0, 60, '社保基数')
        sheet.write(0, 61, '社保其他说明')
        sheet.write(0, 62, '商业保险')
        sheet.write(0, 63, '异动情况')
        sheet.write(0, 64, '奖惩记录')
        sheet.write(0, 65, '劳动合同签订情况')
        sheet.write(0, 66, '体检')
        sheet.write(0, 67, '其他')
        sheet.write(0, 68, '工号')
        sheet.write(0, 69, '员工生日')
        sheet.write(0, 70, '工龄开始日期')
        sheet.write(0, 71, '合同年限')
        sheet.write(0, 72, '年龄')
        sheet.write(0, 73, '子女姓名')
        sheet.write(0, 74, '子女性别')
        sheet.write(0, 75, '子女身份证号')
        sheet.write(0, 76, '考勤规则')
        sheet.write(0, 77, '子女姓名1')
        sheet.write(0, 78, '子女性别1')
        sheet.write(0, 79, '子女身份证号1')
        sheet.write(0, 80, '子女个数')
        sheet.write(0, 81, '档案编号')
    @staticmethod
    def getValueFromSheet(sheet, cellIndex):
        indexs = SheetHelper.cellColAndRowFromIndex(cellIndex)
        cellValue = sheet.cell_value(rowx=indexs[1], colx=indexs[0])
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