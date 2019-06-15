from SheetHelper import SheetHelper

class Staff:
    # 姓名
    name = ''
    # 用户名
    userName = ''
    # 英文名
    englishName = ''
    # 在职状态
    state = ''
    # 民族
    nation = ''
    # 政治面貌
    politcInfo = ''
    # 性别
    gender = 0
    # 籍贯
    nativePlace = ''
    # 婚姻状态
    mariageState = 0
    # 出生年月
    birthDate = 0
    # 最高学历
    education = ''
    # 兴趣爱好
    hobby = ''
    # 电子邮件
    email = ''
    # 入职时间
    entryTime = 0
    # 特长
    skill = ''
    # 身份证
    identityNumber = ''
    # 户口所在地
    householdPlace = ''
    # 户口性质
    householdType = ''
    # 手机号吗
    phoneNumber = 0
    # 家庭电话
    housePhoneNumber = 0
    # 现在住址
    livingPlace = ''
    # 教育经历
    educationInfo = ''
    # 培训经历
    trainInfo = ''
    # 专业资格证书...
    professionInfo = ''
    # 工作履历
    workHistory = ''
    # 家庭成员
    familyMember = ''
    # 紧急联系人
    urgentContact = ''
    # 工资卡号
    bankCard = ''
    # 银行卡开户地
    bankPosition = ''
    # 提交资料
    submitInfo = ''
    # 参加工作时间
    startWorkTime = ''
    # 员工类型
    staffType = ''
    # 职位
    position = ''

    def __init__(self, sheet):
        self.name = SheetHelper.getValueFromSheet(sheet, 'B4')
        
        self.englishName = SheetHelper.getValueFromSheet(sheet, 'D4')

        self.nation = SheetHelper.getValueFromSheet(sheet, 'F4')

        self.politcInfo = SheetHelper.getValueFromSheet(sheet, 'F5')

        self.gender = SheetHelper.getValueFromSheet(sheet, 'G5')

        self.nativePlace = SheetHelper.getValueFromSheet(sheet, 'B5')

        self.mariageState = SheetHelper.getValueFromSheet(sheet, 'B6')

        self.birthDate = SheetHelper.getValueFromSheet(sheet, 'D5')

        self.skill = SheetHelper.getValueFromSheet(sheet, 'B41')

        self.identityNumber = SheetHelper.getValueFromSheet(sheet, 'F8')

        self.householdPlace = SheetHelper.getValueFromSheet(sheet, 'F10')

        self.householdType = SheetHelper.getValueFromSheet(sheet, 'F6')

        self.phoneNumber = SheetHelper.getValueFromSheet(sheet, 'C9')

        # self.housePhoneNumber = SheetHelper.getValueFromSheet(sheet, 'F8')

        self.livingPlace = SheetHelper.getValueFromSheet(sheet, 'F9')

        educationInfos = SheetHelper.getValueFromBounds(sheet, 'B12', 'H14')
        educationInfosStrings = []
        for educationInfo in educationInfos:
            educationInfosStrings.append('起止时间：{0[0]}；学校：{0[1]}；专业：{0[2]}；证书：{0[3]}'.format(educationInfo))
        self.educationInfo = '；'.join(educationInfosStrings)

        trainInfos = SheetHelper.getValueFromBounds(sheet, 'B16', 'H18')
        trainInfosStrings = []
        for trainInfo in trainInfos:
            trainInfosStrings.append('起止时间：{0[0]}；培训机构：{0[1]}；培训课程：{0[2]};证书：{0[3]}'.format(trainInfo))
        self.trainInfo = '；'.join(trainInfosStrings)

        self.professionInfo = SheetHelper.getValueFromSheet(sheet, 'B19')

        workHistorys = SheetHelper.getValueFromBounds(sheet, 'B23', 'H26')
        workHistoryStrings = []
        for workHistory in workHistorys:
            workHistoryStrings.append('起止时间：{0[0]}；工作单位：{0[1]}；职位：{0[2]};薪资状况：{0[3]}；离职原因：{0[4]}'.format(workHistory))
        self.workHistory = '；'.join(workHistoryStrings)

        familyMembers = SheetHelper.getValueFromBounds(sheet, 'B28', 'H30')
        familyMemberStrings = []
        for familyMember in familyMembers:
            familyMemberStrings.append('姓名：{0[0]}；关系：{0[1]}；出生年月：{0[2]}；工作单位：{0[3]}；职务：{0[4]}'.format(familyMember))
        self.familyMember = '；'.join(familyMemberStrings)

        urgentContact = ''
        if SheetHelper.getValueFromSheet(sheet, 'B36') != '':
            urgentContact = '紧急联系人1：%s；双方关系：%s；联系电话：%s' % (SheetHelper.getValueFromSheet(sheet, 'B36'),SheetHelper.getValueFromSheet(sheet, 'D36'),
            SheetHelper.getValueFromSheet(sheet, 'G36'))
        if SheetHelper.getValueFromSheet(sheet, 'B37') != '':
            urgentContact += '；紧急联系人2：%s；双方关系：%s；联系电话：%s' % (SheetHelper.getValueFromSheet(sheet, 'B37'),SheetHelper.getValueFromSheet(sheet, 'D37'),
            SheetHelper.getValueFromSheet(sheet, 'G37'))
        self.urgentContact = urgentContact

        self.bankCard = SheetHelper.getValueFromSheet(sheet, 'C38')

        self.bankPosition = SheetHelper.getValueFromSheet(sheet, 'F38')

        self.bankCard = SheetHelper.getValueFromSheet(sheet, 'C38')

        self.position = SheetHelper.getValueFromSheet(sheet, 'D3')

    def writeToSheet(self, sheet, row):
        sheet.write(row, 0, self.name)
        
        sheet.write(row, 2, self.englishName)

        sheet.write(row, 4, self.nation)

        sheet.write(row, 5, self.politcInfo)

        sheet.write(row, 6, self.gender)

        sheet.write(row, 7, self.nativePlace)

        sheet.write(row, 8, self.mariageState)

        sheet.write(row, 9, self.birthDate)

        sheet.write(row, 10, self.education)

        sheet.write(row, 11, self.hobby)

        sheet.write(row, 12, self.email)

        sheet.write(row, 13, self.entryTime)

        sheet.write(row, 14, self.skill)

        sheet.write(row, 15, self.identityNumber)

        sheet.write(row, 16, self.householdPlace)

        sheet.write(row, 17, self.householdType)

        sheet.write(row, 18, self.phoneNumber)

        sheet.write(row, 19, self.housePhoneNumber)

        sheet.write(row, 20, self.livingPlace)

        sheet.write(row, 21, self.educationInfo)

        sheet.write(row, 22, self.trainInfo)

        sheet.write(row, 23, self.professionInfo)

        sheet.write(row, 24, self.workHistory)

        sheet.write(row, 25, self.familyMember)

        sheet.write(row, 26, self.urgentContact)

        sheet.write(row, 27, self.bankCard)

        sheet.write(row, 28, self.bankPosition)

        sheet.write(row, 29, self.submitInfo)

        sheet.write(row, 30, self.startWorkTime)

        sheet.write(row, 31, self.staffType)

        sheet.write(row, 32, self.position)
