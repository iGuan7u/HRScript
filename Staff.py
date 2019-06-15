from SheetHelper import SheetHelper

class Staff:
    # 姓名
    name = ''
    # 工作地点
    workPlace = ''
    # 中心
    center = ''
    # 部门
    department = ''
    # 科室
    administration = ''
    # 组
    group = ''
    # 一级主管
    leader = ''
    # 用户名
    userName = ''
    # OA
    OAName = ''
    # 英文名
    englishName = ''
    # 在职状态
    state = '在职'
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
    workEmail = ''
    # 入职时间
    entryTime = 0
    # 特长
    skill = ''
    # 身份证
    identityNumber = ''
    # 户口所在地详细地址
    householdPlace = ''
    # 户口性质
    householdType = ''
    # 手机
    phoneNumber = 0
    # 家庭电话
    housePhoneNumber = 0
    # 现在住址
    livingPlace = ''
    # 教育经历
    educationInfo = ''
    # 培训经历
    trainInfo = ''
    # 专业资格证书及获取
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
    submitInfo = 0
    # 参加工作时间
    startWorkTime = 0
    # 员工类型
    staffType = ''
    # 职位
    position = ''
    # 岗位级别
    positionLevel = ''
    # 职级
    grade = ''
    # 个人Email
    personEmail = ''
    # 合同期开始
    contractStartTime = ''
    # 合同期结束
    contractEndTime = ''
    # 合同到期提醒日期
    contractRemindTime = ''
    # 试用期
    probation = ''
    # 转正日期
    fullTime = ''
    # 离职日期
    leaveTime = ''
    # 离职原因
    leaveReason = ''
    # 公积金购买日期
    fundTime = ''
    # 公积金比例
    fundPercent = ''
    # 公积金类型
    fundType = ''
    # 公积金封存日期
    fundXXTime = ''
    # 公积金基数
    fundXXNum = ''
    # 公积金号
    fundNumber = ''
    # 社保增员日期
    insuranceDate = ''
    # 是否当地社保新开户
    insuranceType = ''
    # 社保减员日期
    insuranceDate2 = ''
    # 社保基数
    insuranceXXNum = ''
    # 社保其他说明
    insuranceMemo = ''
    # 商业保险
    businessInsurance = ''
    # 异动情况
    moveMemo = ''
    # 奖惩记录
    record = 0
    # 劳动合同签订情况
    laborContract = ''
    # 体检
    bodyCheck = ''
    # 其他
    other = ''
    # 工号
    workNum = ''
    # 工龄开始日期
    workingYearsStartTime = ''
    # 合同年限
    contractDuration = ''
    # 年龄
    age = ''
    # 子女姓名
    insuranceKid = ''
    # 子女性别
    insuranceKidGender = ''
    # 子女身份证号
    insuranceKidIdentifierNumber = ''
    # 考勤规则
    rule = ''
    # 子女姓名1
    kid1 = ''
    # 子女性别1
    kid1Gender = ''
    # 子女身份证号1
    kid1Identify = ''
    # 子女个数
    kidCount = ''
    # 档案编号
    fileNumber = ''

    @property
    def company(self):
        if self.workPlace == '广州':
            return '广州威尔森信息科技有限公司'
        elif self.workPlace == '北京' or self.workPlace == '长春':
            return '广州威尔森信息科技有限公司北京分公司'
        elif self.workPlace == '上海':
            return '广州威尔森信息科技有限公司上海分公司'
        else:
            return 'something is wrong'
    

    def __init__(self, sheet):
        self.name = SheetHelper.getValueFromSheet(sheet, 'B3')

        self.center = SheetHelper.getValueFromSheet(sheet, 'B53')

        self.department = SheetHelper.getValueFromSheet(sheet, 'D53')

        self.administration = SheetHelper.getValueFromSheet(sheet, 'F53')

        self.group = SheetHelper.getValueFromSheet(sheet, 'H53')

        self.workPlace = SheetHelper.getValueFromSheet(sheet, 'B56')

        self.leader = SheetHelper.getValueFromSheet(sheet, 'H54')
        
        self.englishName = SheetHelper.getValueFromSheet(sheet, 'D3')

        self.nation = SheetHelper.getValueFromSheet(sheet, 'F3')

        self.politcInfo = SheetHelper.getValueFromSheet(sheet, 'F5')

        self.gender = SheetHelper.getValueFromSheet(sheet, 'F4')

        self.nativePlace = SheetHelper.getValueFromSheet(sheet, 'B4')

        self.mariageState = SheetHelper.getValueFromSheet(sheet, 'B5')

        self.birthDate = SheetHelper.getValueFromSheet(sheet, 'D4')

        self.hobby = SheetHelper.getValueFromSheet(sheet, 'B6')

        self.workEmail = SheetHelper.getValueFromSheet(sheet, 'C56')

        self.entryTime = SheetHelper.getValueFromSheet(sheet, 'B55')

        self.education = SheetHelper.getValueFromSheet(sheet, 'D6')

        self.skill = SheetHelper.getValueFromSheet(sheet, 'B6')

        self.identityNumber = SheetHelper.getValueFromSheet(sheet, 'F7')

        self.householdPlace = SheetHelper.getValueFromSheet(sheet, 'F9')

        self.householdType = SheetHelper.getValueFromSheet(sheet, 'F6')

        self.phoneNumber = SheetHelper.getValueFromSheet(sheet, 'C8')

        self.livingPlace = SheetHelper.getValueFromSheet(sheet, 'F8')

        educationInfos = SheetHelper.getValueFromBounds(sheet, 'B11', 'H13')
        educationInfosStrings = []
        for educationInfo in educationInfos:
            educationInfosStrings.append('起止时间：{0[0]}；学校：{0[1]}；专业：{0[2]}；证书：{0[3]}'.format(educationInfo))
        self.educationInfo = '；'.join(educationInfosStrings)

        trainInfos = SheetHelper.getValueFromBounds(sheet, 'B15', 'H17')
        trainInfosStrings = []
        for trainInfo in trainInfos:
            trainInfosStrings.append('起止时间：{0[0]}；培训机构：{0[1]}；培训课程：{0[2]};证书：{0[3]}'.format(trainInfo))
        self.trainInfo = '；'.join(trainInfosStrings)

        self.professionInfo = SheetHelper.getValueFromSheet(sheet, 'B18')

        workHistorys = SheetHelper.getValueFromBounds(sheet, 'B22', 'H25')
        workHistoryStrings = []
        for workHistory in workHistorys:
            workHistoryStrings.append('起止时间：{0[0]}；工作单位：{0[1]}；职位：{0[2]};薪资状况：{0[3]}；离职原因：{0[4]}'.format(workHistory))
        self.workHistory = '；'.join(workHistoryStrings)

        familyMembers = SheetHelper.getValueFromBounds(sheet, 'B27', 'H29')
        familyMemberStrings = []
        for familyMember in familyMembers:
            familyMemberStrings.append('姓名：{0[0]}；关系：{0[1]}；出生年月：{0[2]}；工作单位：{0[3]}；职务：{0[4]}'.format(familyMember))
        self.familyMember = '；'.join(familyMemberStrings)

        urgentContact = ''
        if SheetHelper.getValueFromSheet(sheet, 'B35') != '':
            urgentContact = '紧急联系人1：%s；双方关系：%s；联系电话：%s' % (SheetHelper.getValueFromSheet(sheet, 'B35'),SheetHelper.getValueFromSheet(sheet, 'D35'),
            SheetHelper.getValueFromSheet(sheet, 'G35'))
        if SheetHelper.getValueFromSheet(sheet, 'B36') != '':
            urgentContact += '；紧急联系人2：%s；双方关系：%s；联系电话：%s' % (SheetHelper.getValueFromSheet(sheet, 'B36'),SheetHelper.getValueFromSheet(sheet, 'D36'),
            SheetHelper.getValueFromSheet(sheet, 'G36'))
        self.urgentContact = urgentContact

        self.bankCard = SheetHelper.getValueFromSheet(sheet, 'C37')

        self.bankPosition = SheetHelper.getValueFromSheet(sheet, 'F37')

        self.staffType = SheetHelper.getValueFromSheet(sheet, 'H55')

        self.position = SheetHelper.getValueFromSheet(sheet, 'B54')

        self.positionLevel = SheetHelper.getValueFromSheet(sheet, 'D54')

        self.grade = SheetHelper.getValueFromSheet(sheet, 'F54')

        self.personEmail = SheetHelper.getValueFromSheet(sheet, 'C9')

        self.contractStartTime = SheetHelper.getValueFromSheet(sheet, 'B55')

        self.probation = SheetHelper.getValueFromSheet(sheet, 'D55')

        self.fundPercent = SheetHelper.getValueFromSheet(sheet, 'C44')

        self.fundType = SheetHelper.getValueFromSheet(sheet, 'G43')

        self.fundNumber = SheetHelper.getValueFromSheet(sheet, 'G44')

        self.insuranceType = SheetHelper.getValueFromSheet(sheet, 'C43')

        self.insuranceMemo = SheetHelper.getValueFromSheet(sheet, 'F6')

        self.workingYearsStartTime = SheetHelper.getValueFromSheet(sheet, 'B55')

        self.contractDuration = SheetHelper.getValueFromSheet(sheet, 'F55')

        self.insuranceKid = SheetHelper.getValueFromSheet(sheet, 'C46')

        self.insuranceKidGender = SheetHelper.getValueFromSheet(sheet, 'G46')

        self.insuranceKidIdentifierNumber = SheetHelper.getValueFromSheet(sheet, 'C47')
        
        self.kid1 = SheetHelper.getValueFromSheet(sheet, 'B49')
        
        self.kid1Gender = SheetHelper.getValueFromSheet(sheet, 'C47')
        
        self.kid1Identify = SheetHelper.getValueFromSheet(sheet, 'F49')
        
        self.kidCount = SheetHelper.getValueFromSheet(sheet, 'D5')
        
        self.insuranceKidIdentifierNumber = SheetHelper.getValueFromSheet(sheet, 'C47')
        
        self.insuranceKidIdentifierNumber = SheetHelper.getValueFromSheet(sheet, 'C47')
        
        self.insuranceKidIdentifierNumber = SheetHelper.getValueFromSheet(sheet, 'C47')
        
        self.insuranceKidIdentifierNumber = SheetHelper.getValueFromSheet(sheet, 'C47')
    
    def writeTheHeader(self, sheet):
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

    def writeToSheet(self, sheet, row):
        sheet.write(row, 1, self.company)
        sheet.write(row, 2, self.center)
        sheet.write(row, 3, self.department)
        sheet.write(row, 4, self.administration)
        sheet.write(row, 5, self.group)
        sheet.write(row, 6, self.workPlace)
        sheet.write(row, 7, self.leader)
        sheet.write(row, 8, self.name)
        sheet.write(row, 9, self.OAName)
        sheet.write(row, 10, self.englishName)
        sheet.write(row, 11, self.state)
        sheet.write(row, 12, self.nation)
        sheet.write(row, 13, self.politcInfo)
        sheet.write(row, 14, self.gender)
        sheet.write(row, 15, self.nativePlace)
        sheet.write(row, 16, self.mariageState)
        sheet.write(row, 17, self.birthDate)
        sheet.write(row, 18, self.education)
        sheet.write(row, 19, self.hobby)
        sheet.write(row, 20, self.workEmail)
        sheet.write(row, 21, self.entryTime)
        sheet.write(row, 22, self.skill)
        sheet.write(row, 23, self.identityNumber)
        sheet.write(row, 24, self.householdPlace)
        sheet.write(row, 25, self.householdType)
        sheet.write(row, 26, self.phoneNumber)
        sheet.write(row, 27, self.housePhoneNumber)
        sheet.write(row, 28, self.livingPlace)
        sheet.write(row, 29, self.educationInfo)
        sheet.write(row, 30, self.trainInfo)
        sheet.write(row, 31, self.professionInfo)
        sheet.write(row, 32, self.workHistory)
        sheet.write(row, 33, self.familyMember)
        sheet.write(row, 34, self.urgentContact)
        sheet.write(row, 35, self.bankCard)
        sheet.write(row, 36, self.bankPosition)
        sheet.write(row, 37, self.submitInfo)
        sheet.write(row, 38, self.startWorkTime)
        sheet.write(row, 39, self.staffType)
        sheet.write(row, 40, self.position)
        sheet.write(row, 41, self.positionLevel)
        sheet.write(row, 42, self.grade)
        sheet.write(row, 43, self.personEmail)
        sheet.write(row, 44, self.contractStartTime)
        sheet.write(row, 45, self.contractEndTime)
        sheet.write(row, 46, self.contractRemindTime)
        sheet.write(row, 47, self.probation)
        sheet.write(row, 48, self.fullTime)
        sheet.write(row, 49, self.leaveTime)
        sheet.write(row, 50, self.leaveReason)
        sheet.write(row, 51, self.fundTime)
        sheet.write(row, 52, self.fundPercent)
        sheet.write(row, 53, self.fundType)
        sheet.write(row, 54, self.fundXXTime)
        sheet.write(row, 55, self.fundXXNum)
        sheet.write(row, 56, self.fundNumber)
        sheet.write(row, 57, self.insuranceDate)
        sheet.write(row, 58, self.insuranceType)
        sheet.write(row, 59, self.insuranceDate2)
        sheet.write(row, 60, self.insuranceXXNum)
        sheet.write(row, 61, self.insuranceMemo)
        sheet.write(row, 62, self.businessInsurance)
        sheet.write(row, 63, self.moveMemo)
        sheet.write(row, 64, self.record)
        sheet.write(row, 65, self.laborContract)
        sheet.write(row, 66, self.bodyCheck)
        sheet.write(row, 67, self.other)
        sheet.write(row, 68, self.workNum)
        # sheet.write(row, 69, '员工生日')
        sheet.write(row, 70, self.workingYearsStartTime)
        sheet.write(row, 71, self.contractDuration)
        sheet.write(row, 72, self.age)
        sheet.write(row, 73, self.insuranceKid)
        sheet.write(row, 74, self.insuranceKidGender)
        sheet.write(row, 75, self.insuranceKidIdentifierNumber)
        sheet.write(row, 76, self.rule)
        sheet.write(row, 77, self.kid1)
        sheet.write(row, 78, self.kid1Gender)
        sheet.write(row, 79, self.kid1Identify)
        sheet.write(row, 80, self.kidCount)
        sheet.write(row, 81, self.fileNumber)