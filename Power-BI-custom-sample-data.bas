Attribute VB_Name = "Power-BI-custom-sample-data"
Option Compare Database
Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'1、作者：焦棚子
'2、邮箱：jiaopengzi@qq.com
'3、博客：www.jiaopengzi.com
'4、CPU：12th Gen Intel(R) Core(TM) i9-12900KF   3.20 GHz
'5、内存：RAM 32.0 GB
'6、如上电脑配置 + ShopQuantity=300 的配置：大约需要 1000 秒，每秒按照业务逻辑生成约 1万行+ 数据；生成 1000 万行+ demo数据，基本满足实战学习所用。
'   如上电脑配置 + ShopQuantity=100 的配置：大约需要  350 秒，每秒按照业务逻辑生成约 1万行+ 数据；生成  360 万行+ demo数据，基本满足实战学习所用。
'   如上电脑配置 + ShopQuantity=10  的配置：大约需要   60 秒，每秒按照业务逻辑生成约 1万行+ 数据；生成   60 万行+ demo数据，基本满足实战学习所用。
'   如上电脑配置 + ShopQuantity=5   的配置：大约需要   20 秒，每秒按照业务逻辑生成约 1万行+ 数据；生成   20 万行+ demo数据，基本满足实战学习所用。

'=====================================================================================

Public productQuantity As Long   '产品数量；建议ShopQuantity∈[7,999]。
Public ShopQuantity As Long   '门店数量；建议ShopQuantity∈[1,390]。
Public MaxInventoryDays As Long   '入库间隔最大数；建议ShopQuantity∈[5,20]。

'=====================================================================================表名称管理
Public Const tbNameOrg As String = "D10_组织表"                         ' D10
Public Const tbNameRegion As String = "D20_大区表"                      ' D20
Public Const tbNameProvince As String = "D21_省份表"                    ' D21
Public Const tbNameCity As String = "D22_城市表"                        ' D22
Public Const tbNameDistrict As String = "D23_区县表"                    ' D23
Public Const tbNameProduct As String = "D30_产品表"                     ' D30
Public Const tbNameShop As String = "T10_门店表"                        ' T10
Public Const tbNameShopRental As String = "T11_门店表_租赁"             ' T11
Public Const tbNameShopDecoration As String = "T12_门店表_装修"         ' T12
Public Const tbNameCustomer As String = "T20_客户表"                    ' T20
Public Const tbNameStorage As String = "T30_入库信息表"                 ' T30
Public Const tbNameOrder As String = "T40_订单主表"                     ' T40
Public Const tbNameOrdersub As String = "T41_订单子表"                  ' T41
Public Const tbNameSaleTarget As String = "T50_销售目标表"              ' T50
Public Const tbNameEmployee As String = "T60_员工信息表"                ' T60
Public Const tbNameLaborCost As String = "T61_人工成本表"               ' T61

'=====================================================================================表的字典名称和生成SQL
Public Const fLaborCostOrgID As String = "组织ID"
Public Const fLaborCostMonth As String = "月份"
Public Const fLaborCostAmount As String = "人工成本金额_元"
Public Const createTbSqlLaborCost As String = "CREATE TABLE " & tbNameLaborCost & _
            "(                                                                  " & vbCrLf & _
            fLaborCostOrgID & "         INT,                                    " & vbCrLf & _
            fLaborCostMonth & "         DATE,                                   " & vbCrLf & _
            fLaborCostAmount & "        INT                                     " & vbCrLf & _
            ")"

Public Const fProductID As String = "产品ID"
Public Const fProductCategory As String = "产品分类"
Public Const fProductName As String = "产品名称"
Public Const fProductPrice As String = "产品销售价格"
Public Const fProductCostPrice = "产品成本价格"
Public Const createTbSqlProduct As String = "CREATE TABLE " & tbNameProduct & _
            "(                                                                  " & vbCrLf & _
            fProductID & "             VARCHAR(50) PRIMARY KEY,                 " & vbCrLf & _
            fProductCategory & "       VARCHAR(50),                             " & vbCrLf & _
            fProductName & "           VARCHAR(50),                             " & vbCrLf & _
            fProductPrice & "          INT,                                     " & vbCrLf & _
            fProductCostPrice & "      INT                                      " & vbCrLf & _
            ")"
            
'门店ID在创建是不设置ID主键,因为ID留空后置处理
Public Const fShopID As String = "门店组织ID"
Public Const fShopName As String = "门店名称"
Public Const fShopOpenDate As String = "开业日期"
Public Const fShopDistrictID As String = "区县ID"
Public Const fShopDistrict As String = "区县"
Public Const fShopLongitude As String = "纬度"
Public Const fShopLatitude As String = "经度"
Public Const fShopCloseDate As String = "闭店日期"
Public Const createTbSqlShop As String = "CREATE TABLE " & tbNameShop & _
            "(                                                                  " & vbCrLf & _
            fShopID & "                INT,                                     " & vbCrLf & _
            fShopName & "              VARCHAR(50),                             " & vbCrLf & _
            fShopOpenDate & "          DATE,                                    " & vbCrLf & _
            fShopDistrictID & "        INT,                                     " & vbCrLf & _
            fShopDistrict & "          VARCHAR(50),                             " & vbCrLf & _
            fShopLongitude & "         FLOAT,                                   " & vbCrLf & _
            fShopLatitude & "          FLOAT,                                   " & vbCrLf & _
            fShopCloseDate & "         DATE                                     " & vbCrLf & _
            ")"
            
Public Const fShopRentalShopID As String = "门店组织ID"
Public Const fShopRentalArea As String = "房屋面积_平方米"
Public Const fShopRentalPrice As String = "房屋租金_元每月每平方米"
Public Const fShopRentalStartDate As String = "起租日期"
Public Const fShopRentalEndDate As String = "止租日期"
Public Const fShopRentalIncrease As String = "年度租金涨幅"
Public Const createTbSqlShopRental As String = "CREATE TABLE " & tbNameShopRental & _
            "(                                                                  " & vbCrLf & _
            fShopRentalShopID & "      INT,                                     " & vbCrLf & _
            fShopRentalArea & "        FLOAT,                                   " & vbCrLf & _
            fShopRentalPrice & "       FLOAT,                                   " & vbCrLf & _
            fShopRentalStartDate & "   DATE,                                    " & vbCrLf & _
            fShopRentalEndDate & "     DATE,                                    " & vbCrLf & _
            fShopRentalIncrease & "    FLOAT                                    " & vbCrLf & _
            ")"
            
Public Const fShopDecorationShopID As String = "门店组织ID"
Public Const fShopDecorationStartDate As String = "装修开始日期"
Public Const fShopDecorationEndDate As String = "装修结束日期"
Public Const fShopDecorationAmount As String = "装修金额_元"
Public Const fShopDecorationYears As String = "装修折旧年限"
Public Const createTbSqlShopDecoration As String = "CREATE TABLE " & tbNameShopDecoration & _
            "(                                                                  " & vbCrLf & _
            fShopDecorationShopID & "  INT,                                     " & vbCrLf & _
            fShopDecorationStartDate & " DATE,                                  " & vbCrLf & _
            fShopDecorationEndDate & " DATE,                                    " & vbCrLf & _
            fShopDecorationAmount & "  FLOAT,                                   " & vbCrLf & _
            fShopDecorationYears & "   FLOAT                                    " & vbCrLf & _
            ")"
            
Public Const fCustomerID As String = "客户ID"
Public Const fCustomerName As String = "客户名称"
Public Const fCustomerBirthday As String = "客户生日"
Public Const fCustomerGender As String = "客户性别"
Public Const fCustomerRegister As String = "注册日期"
Public Const fCustomerIndustry As String = "客户行业"
Public Const fCustomerOccupation As String = "客户职业"
Public Const createTbSqlCustomer As String = "CREATE TABLE " & tbNameCustomer & _
            "(                                                                  " & vbCrLf & _
            fCustomerID & "            VARCHAR(50) PRIMARY KEY,                 " & vbCrLf & _
            fCustomerName & "          VARCHAR(50),                             " & vbCrLf & _
            fCustomerBirthday & "      DATE,                                    " & vbCrLf & _
            fCustomerGender & "        VARCHAR(50),                             " & vbCrLf & _
            fCustomerRegister & "      DATE,                                    " & vbCrLf & _
            fCustomerIndustry & "      VARCHAR(50),                             " & vbCrLf & _
            fCustomerOccupation & "    VARCHAR(50)                              " & vbCrLf & _
            ")"

Public Const fStorageProductID As String = "入库产品ID"
Public Const fStorageQuantity As String = "入库产品数量"
Public Const fStorageShopID As String = "入库门店组织ID"
Public Const fStorageDate As String = "入库日期"
Public Const createTbSqlStorage As String = "CREATE TABLE " & tbNameStorage & _
            "(                                                                  " & vbCrLf & _
            fStorageProductID & "      VARCHAR(50),                             " & vbCrLf & _
            fStorageQuantity & "       INT,                                     " & vbCrLf & _
            fStorageShopID & "         INT,                                     " & vbCrLf & _
            fStorageDate & "           DATE                                     " & vbCrLf & _
            ")"
         
Public Const fOrderID As String = "订单ID"
Public Const fOrderShopID As String = "门店组织ID"
Public Const fOrderDate As String = "下单日期"
Public Const fOrderSentDate As String = "送货日期"
Public Const fOrderCustomerID As String = "客户ID"
Public Const fOrderType As String = "销售渠道"
Public Const fOrderEmployeeID As String = "销售员工ID"
Public Const createTbSqlOrder As String = "CREATE TABLE " & tbNameOrder & _
            "(                                                                  " & vbCrLf & _
            fOrderID & "               VARCHAR(50) PRIMARY KEY,                 " & vbCrLf & _
            fOrderShopID & "           INT,                                     " & vbCrLf & _
            fOrderDate & "             DATE,                                    " & vbCrLf & _
            fOrderSentDate & "         DATE,                                    " & vbCrLf & _
            fOrderCustomerID & "       VARCHAR(50),                             " & vbCrLf & _
            fOrderType & "             VARCHAR(50),                             " & vbCrLf & _
            fOrderEmployeeID & "       INT                                      " & vbCrLf & _
            ")"

Public Const fOrdersubOrderID As String = "订单ID"
Public Const fOrdersubProductID As String = "产品ID"
Public Const fOrdersubPrice As String = "产品销售价格"
Public Const fOrdersubDiscount As String = "折扣比例"
Public Const fOrdersubQuantity As String = "产品销售数量"
Public Const fOrdersubAmount As String = "产品销售金额"
Public Const createTbSqlOrdersub As String = "CREATE TABLE " & tbNameOrdersub & _
            "(                                                                  " & vbCrLf & _
            fOrdersubOrderID & "       VARCHAR(50),                             " & vbCrLf & _
            fOrdersubProductID & "     VARCHAR(50),                             " & vbCrLf & _
            fOrdersubPrice & "         INT,                                     " & vbCrLf & _
            fOrdersubDiscount & "      FLOAT,                                   " & vbCrLf & _
            fOrdersubQuantity & "      INT,                                     " & vbCrLf & _
            fOrdersubAmount & "        FLOAT,                                   " & vbCrLf & _
            "CONSTRAINT PK_" & tbNameOrdersub & " PRIMARY KEY (订单ID, 产品ID)  " & vbCrLf & _
            ")"
            
Public Const fSaleTargetProvinceID As String = "省ID"
Public Const fSaleTargetProvinceName2 As String = "省简称"
Public Const fSaleTargetMonth As String = "月份"
Public Const fSaleTargetAmount As String = "销售目标"
Public Const createTbSqlSaleTarget As String = "CREATE TABLE " & tbNameSaleTarget & _
            "(                                                                  " & vbCrLf & _
            fSaleTargetProvinceID & "  INT,                                     " & vbCrLf & _
            fSaleTargetProvinceName2 & " VARCHAR(50),                           " & vbCrLf & _
            fSaleTargetMonth & "       DATE,                                    " & vbCrLf & _
            fSaleTargetAmount & "      FLOAT                                    " & vbCrLf & _
            ")"

Public Const fRegionID As String = "大区组织ID"
Public Const fRegionName As String = "简称"
Public Const fRegionCityID As String = "办公地城市ID"
Public Const fRegionCity As String = "办公地城市"
Public Const fRegionLongitude As String = "纬度"
Public Const fRegionLatitude As String = "经度"
Public Const createTbSqlRegion As String = "CREATE TABLE " & tbNameRegion & _
            "(                                                                  " & vbCrLf & _
            fRegionID & "              INT PRIMARY KEY,                         " & vbCrLf & _
            fRegionName & "            VARCHAR(50),                             " & vbCrLf & _
            fRegionCityID & "          INT,                                     " & vbCrLf & _
            fRegionCity & "            VARCHAR(50),                             " & vbCrLf & _
            fRegionLongitude & "       FLOAT,                                   " & vbCrLf & _
            fRegionLatitude & "        FLOAT                                    " & vbCrLf & _
            ")"

Public Const fProvinceRegionID As String = "大区组织ID"
Public Const fProvinceID As String = "省ID"
Public Const fProvinceNameAll As String = "省全称"
Public Const fProvinceName1 As String = "省简称1"
Public Const fProvinceName2 As String = "省简称2"
Public Const fProvinceLongitude As String = "纬度"
Public Const fProvinceLatitude As String = "经度"
Public Const createTbSqlProvince As String = "CREATE TABLE " & tbNameProvince & _
            "(                                                                  " & vbCrLf & _
            fProvinceRegionID & "      INT,                                     " & vbCrLf & _
            fProvinceID & "            INT,                                     " & vbCrLf & _
            fProvinceNameAll & "       VARCHAR(50),                             " & vbCrLf & _
            fProvinceName1 & "         VARCHAR(50),                             " & vbCrLf & _
            fProvinceName2 & "         VARCHAR(50),                             " & vbCrLf & _
            fProvinceLongitude & "     FLOAT,                                   " & vbCrLf & _
            fProvinceLatitude & "      FLOAT                                    " & vbCrLf & _
            ")"

Public Const fCityProvinceID As String = "省ID"
Public Const fCityID As String = "城市ID"
Public Const fCityName As String = "城市"
Public Const fCityLongitude As String = "纬度"
Public Const fCityLatitude As String = "经度"
Public Const createTbSqlCity As String = "CREATE TABLE " & tbNameCity & _
            "(                                                                  " & vbCrLf & _
            fCityProvinceID & "        INT,                                     " & vbCrLf & _
            fCityID & "                INT PRIMARY KEY,                         " & vbCrLf & _
            fCityName & "              VARCHAR(50),                             " & vbCrLf & _
            fCityLongitude & "         FLOAT,                                   " & vbCrLf & _
            fCityLatitude & "          FLOAT                                    " & vbCrLf & _
            ")"

Public Const fDistrictCityID As String = "城市ID"
Public Const fDistrictID As String = "区县ID"
Public Const fDistrictName As String = "区县"
Public Const fDistrictLongitude As String = "纬度"
Public Const fDistrictLatitude As String = "经度"
Public Const createTbSqlDistrict As String = "CREATE TABLE " & tbNameDistrict & _
            "(                                                                  " & vbCrLf & _
            fDistrictCityID & "        INT,                                     " & vbCrLf & _
            fDistrictID & "            INT PRIMARY KEY,                         " & vbCrLf & _
            fDistrictName & "          VARCHAR(50),                             " & vbCrLf & _
            fDistrictLongitude & "     FLOAT,                                   " & vbCrLf & _
            fDistrictLatitude & "      FLOAT                                    " & vbCrLf & _
            ")"

Public Const fOrgID As String = "组织ID"
Public Const fOrgNameAll As String = "组织名称"
Public Const fOrgParentID As String = "上级组织ID"
Public Const fOrgName As String = "组织简称"
Public Const fOrgEmployeeID As String = "负责人ID"
Public Const createTbSqlOrg As String = "CREATE TABLE " & tbNameOrg & _
            "(                                                                  " & vbCrLf & _
            fOrgID & "                 INT IDENTITY(1,1) PRIMARY KEY,           " & vbCrLf & _
            fOrgNameAll & "            VARCHAR(255),                            " & vbCrLf & _
            fOrgParentID & "           INT,                                     " & vbCrLf & _
            fOrgName & "               VARCHAR(255),                            " & vbCrLf & _
            fOrgEmployeeID & "         INT                                      " & vbCrLf & _
            ")"
            
Public Const fEmployeeID As String = "员工ID"
Public Const fEmployeeName As String = "姓名"
Public Const fEmployeeGender As String = "性别"
Public Const fEmployeeOrgID As String = "组织ID"
Public Const fEmployeeJobTitle As String = "职务"
Public Const fEmployeeGrade As String = "职级"
Public Const fEmployeeEdu As String = "学历"
Public Const fEmployeeBirthday As String = "出生日期"
Public Const fEmployeeEntryDate As String = "入职日期"
Public Const fEmployeeResignationDate As String = "离职日期"
Public Const fEmployeeResignationReason As String = "离职原因"
Public Const createTbSqlEmployee As String = "CREATE TABLE " & tbNameEmployee & _
            "(                                                                  " & vbCrLf & _
            fEmployeeID & "            INT IDENTITY(10001,1) PRIMARY KEY,       " & vbCrLf & _
            fEmployeeName & "          VARCHAR(50),                             " & vbCrLf & _
            fEmployeeGender & "        VARCHAR(20) DEFAULT 男,                  " & vbCrLf & _
            fEmployeeOrgID & "         INT,                                     " & vbCrLf & _
            fEmployeeJobTitle & "      VARCHAR(50),                             " & vbCrLf & _
            fEmployeeGrade & "         VARCHAR(50),                             " & vbCrLf & _
            fEmployeeEdu & "           VARCHAR(50),                             " & vbCrLf & _
            fEmployeeBirthday & "      DATE,                                    " & vbCrLf & _
            fEmployeeEntryDate & "     DATE,                                    " & vbCrLf & _
            fEmployeeResignationDate & " DATE NULL,                             " & vbCrLf & _
            fEmployeeResignationReason & " VARCHAR(255) NULL                    " & vbCrLf & _
            ")"
            

'=====================================================================================全局变量
Public TableNameDict As Object ' 表名称字典
Public MinDateOpen As Date ' 最早开业日期
Public ProvinceID2OrgIDDict As Object ' 省份区域ID的前两位与组织ID的映射字典
Public JobTitlesArr As Variant '职务
Public GradeArr As Variant '职级
Public EduArr As Variant '学历、
Public EduDict As Object
Public EduSalaryDict As Object
Public GradeDict As Object
Public GradeSalaryDict As Object
Public ResignationArr As Variant '离职原因


Public Function InitE()
    '初始化员工信息相关内容
    JobTitlesArr = Array("总经理", "总经理助理", "产品总监", "采购总监", "销售总监", "销售总监", "人力资源总监", "售后服务总监", "财务总监", "大区经理", "省区经理", "门店经理", "销售顾问", "售后专员")
    
    GradeArr = Array("总经理", "高级总监", "总监", "高级经理", "经理", "主管", "专员")
        
    EduArr = Array("研究生", "本科", "专科", "高中")
    
    ResignationArr = Array("个人发展", "工资原因", "工资强度", "工作内容与环境", "家庭原因", "身体原因", "违反规章制度", "劝离", "旷离", "其他原因", "试用期内解除") '试用期放在索引10
    Set EduDict = CreateObject("Scripting.Dictionary")
    With EduDict
        .Add "PD", "博士" '博士：Doctorate (PD)
        .Add "PG", "硕士" '研究生: Postgraduate (PG)
        .Add "UG", "本科" '本科: Undergraduate (UG)
        .Add "AD", "专科" '专科: Associate Degree(AD)
        .Add "HS", "高中" '高中: High School(HS)
        .Add "MS", "初中" '初中：Junior High School (JHS) 或 Middle School (MS)
        .Add "PS", "小学" '小学：Primary School (PS) 或 Elementary School (ES)
    End With
    
    Set EduSalaryDict = CreateObject("Scripting.Dictionary")
    With EduSalaryDict
        .Add "博士", 2 '薪资系数
        .Add "硕士", 1.2
        .Add "本科", 1.1
        .Add "专科", 1
        .Add "高中", 0.9
        .Add "初中", 0.8
        .Add "小学", 0.7
    End With
    
    Set GradeDict = CreateObject("Scripting.Dictionary")
    With GradeDict
        .Add "GM", "总经理"     'General Manager (GM)
        .Add "SD", "高级总监"   'Senior Director (SD)
        .Add "D", "总监"        'Director (D)
        .Add "SM", "高级经理"   'Senior Manager (SM)
        .Add "M", "经理"        'Manager (M)
        .Add "S", "主管"        'Supervisor (S)
        .Add "SP", "专员"       'Specialist (SP)
    End With
    
    Set GradeSalaryDict = CreateObject("Scripting.Dictionary")
    With GradeSalaryDict
        .Add "总经理", Array(50000, 100000) '薪资范围
        .Add "高级总监", Array(30000, 50000)
        .Add "总监", Array(20000, 30000)
        .Add "高级经理", Array(12000, 20000)
        .Add "经理", Array(8000, 12000)
        .Add "主管", Array(5000, 8000)
        .Add "专员", Array(3000, 5000)
    End With
End Function


Public Function InitPO()
    ' 初始化 省份区域ID的前两位与组织ID的映射字典
    
    Dim i As Long
    Dim ArrAddProvince
    Dim ArrAddProvinceRow
    Dim rows As Long
    
    ArrAddProvince = Split(AddressProvince, ";")

    ReDim ArrAddProvinceRow(0 To UBound(ArrAddProvince))

    For i = 0 To UBound(ArrAddProvince)
        ArrAddProvinceRow(i) = Split(ArrAddProvince(i), ",")
    Next

    rows = UBound(ArrAddProvinceRow)
    
    Set ProvinceID2OrgIDDict = CreateObject("Scripting.Dictionary")

    For i = 0 To rows
        ProvinceID2OrgIDDict.Add CInt(Left(Trim(ArrAddProvinceRow(i)(1)), 2)), i + 15 '15依据 D20里面的默认值确定的
    Next

     
End Function


Public Function InitTables()
    ' 初始化表名称字典
    Set TableNameDict = CreateObject("Scripting.Dictionary")
    
    ' 将表名称作为键，对应的表的创建sql语句作为值添加到字典中
    With TableNameDict
        .Add tbNameProduct, createTbSqlProduct
        .Add tbNameShop, createTbSqlShop
        .Add tbNameShopRental, createTbSqlShopRental
        .Add tbNameShopDecoration, createTbSqlShopDecoration
        .Add tbNameCustomer, createTbSqlCustomer
        .Add tbNameStorage, createTbSqlStorage
        .Add tbNameOrder, createTbSqlOrder
        .Add tbNameOrdersub, createTbSqlOrdersub
        .Add tbNameSaleTarget, createTbSqlSaleTarget
        .Add tbNameEmployee, createTbSqlEmployee
        .Add tbNameRegion, createTbSqlRegion
        .Add tbNameProvince, createTbSqlProvince
        .Add tbNameCity, createTbSqlCity
        .Add tbNameDistrict, createTbSqlDistrict
        .Add tbNameOrg, createTbSqlOrg
        .Add tbNameLaborCost, createTbSqlLaborCost
    End With
     
End Function

Public Function SQLDrop(tableName As String) As String
' 根据表名称删除表
    SQLDrop = "DROP TABLE " & tableName

End Function

Public Function TableADO(tableName As String, Sql_Drop As String, Sql_Create As String)
    ' 生成表
    Dim Cat As Object
    Dim cmd As Object
    
'    On Error GoTo ErrorHandler
    
    Set Cat = CreateObject("ADOX.Catalog")
    Set cmd = CreateObject("ADODB.Command")
    
    Set Cat.ActiveConnection = CurrentProject.Connection
    Set cmd.ActiveConnection = CurrentProject.Connection
    
    With cmd
        .CommandTimeout = 100
        
        ' 删除已存在的表
        If TableExists(tableName, Cat.tables) Then
            .CommandText = Sql_Drop
            .Execute
        End If
        
        ' 创建新表
        .CommandText = Sql_Create
        .Execute
    End With

CleanUp:
    Set cmd = Nothing
    Set Cat = Nothing
    Exit Function

'ErrorHandler:
'    ' 错误处理代码，可以根据需要进行相应的处理
'    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
'    Resume CleanUp
End Function

Public Function TableExists(tableName As String, tables As Object) As Boolean
    ' 检查表是否存在
    Dim tbl As Object
    
    For Each tbl In tables
        If tbl.name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next
    
    TableExists = False
End Function

Public Function ArrAddRegionDefault() As Variant
'大区默认信息
    Dim ArrAdd
    Dim ArrAddRow
    Dim Region As String
    Dim i As Long
    Region = "9, 东区,   310000, 上海,   31.231518,  121.471518;" & _
            "10, 西区,   510100, 成都,   30.659518,  104.065518;" & _
            "11, 南区,   440100, 广州,   23.125518,  113.280518;" & _
            "12, 北区,   210100, 沈阳,   41.796518,  123.429518;" & _
            "13, 中区,   110000, 北京,   39.901518,  116.401518;" & _
            "14, 港澳台, 810000, 香港,   22.320518,  114.173518"
    
    ArrAdd = Split(Region, ";")

    ReDim ArrAddRow(0 To UBound(ArrAdd))

    For i = 0 To UBound(ArrAdd)
        ArrAddRow(i) = Split(ArrAdd(i), ",")
    Next
    ArrAddRegionDefault = ArrAddRow
End Function

Public Function DataTableRegion()
' 根据业务逻辑生成 大区表
    Dim i As Long
    Dim ArrAddRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsRegion  As Object

    ArrAddRow = ArrAddRegionDefault()
    rows = UBound(ArrAddRow)
    
    Set conn = CreateConnection
    Set RsRegion = CreateRecordset(conn, tbNameRegion)

    For i = 0 To rows

        RsRegion.AddNew
            RsRegion.Fields(fRegionID) = Trim(ArrAddRow(i)(0))
            RsRegion.Fields(fRegionName) = Trim(ArrAddRow(i)(1))
            RsRegion.Fields(fRegionCityID) = Trim(ArrAddRow(i)(2))
            RsRegion.Fields(fRegionCity) = Trim(ArrAddRow(i)(3))
            RsRegion.Fields(fRegionLongitude) = Trim(ArrAddRow(i)(4))
            RsRegion.Fields(fRegionLatitude) = Trim(ArrAddRow(i)(5))
        RsRegion.Update

    Next
    
    CloseConnRs conn, RsRegion

End Function

Public Function DataTableProvince()
' 根据业务逻辑生成 省份表
    Dim i As Long
    Dim ArrAddProvince
    Dim ArrAddProvinceRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsProvince  As Object

    ArrAddProvince = Split(AddressProvince, ";")

    ReDim ArrAddProvinceRow(0 To UBound(ArrAddProvince))

    For i = 0 To UBound(ArrAddProvince)
        ArrAddProvinceRow(i) = Split(ArrAddProvince(i), ",")
    Next

    rows = UBound(ArrAddProvinceRow)

    Set conn = CreateConnection
    Set RsProvince = CreateRecordset(conn, tbNameProvince)

    For i = 0 To rows

        RsProvince.AddNew
            RsProvince.Fields(fProvinceRegionID) = ArrAddProvinceRow(i)(0)
            RsProvince.Fields(fProvinceID) = ArrAddProvinceRow(i)(1)
            RsProvince.Fields(fProvinceNameAll) = ArrAddProvinceRow(i)(2)
            RsProvince.Fields(fProvinceName1) = ArrAddProvinceRow(i)(3)
            RsProvince.Fields(fProvinceName2) = ArrAddProvinceRow(i)(4)
            RsProvince.Fields(fProvinceLongitude) = ArrAddProvinceRow(i)(5)
            RsProvince.Fields(fProvinceLatitude) = ArrAddProvinceRow(i)(6)
        RsProvince.Update
    Next

    CloseConnRs conn, RsProvince
            
End Function

Public Function DataTableCity()
' 根据业务逻辑生成 地市表
    Dim i As Long
    Dim ArrAddCity
    Dim ArrAddCityRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsCity  As Object

    ArrAddCity = Split(AddressCity, ";")

    ReDim ArrAddCityRow(0 To UBound(ArrAddCity))

    For i = 0 To UBound(ArrAddCity)
        ArrAddCityRow(i) = Split(ArrAddCity(i), ",")
    Next

    rows = UBound(ArrAddCityRow)

    Set conn = CreateConnection
    Set RsCity = CreateRecordset(conn, tbNameCity)

    For i = 0 To rows

        RsCity.AddNew
            RsCity.Fields(fCityProvinceID) = ArrAddCityRow(i)(0)
            RsCity.Fields(fCityID) = ArrAddCityRow(i)(1)
            RsCity.Fields(fCityName) = ArrAddCityRow(i)(2)
            RsCity.Fields(fCityLongitude) = ArrAddCityRow(i)(3)
            RsCity.Fields(fCityLatitude) = ArrAddCityRow(i)(4)
        RsCity.Update

    Next

    CloseConnRs conn, RsCity
End Function

Public Function DataTableDistrict()
' 根据业务逻辑生成 区县表

    Dim i As Long
    Dim ArrAddDistrictRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsDistrict  As Object

    ArrAddDistrictRow = ArrAddDistrictRowDefault()
    rows = UBound(ArrAddDistrictRow)
    
    Set conn = CreateConnection
    Set RsDistrict = CreateRecordset(conn, tbNameDistrict)
    
    For i = 0 To rows

        RsDistrict.AddNew
            RsDistrict.Fields(fDistrictCityID) = ArrAddDistrictRow(i)(0)
            RsDistrict.Fields(fDistrictID) = ArrAddDistrictRow(i)(1)
            RsDistrict.Fields(fDistrictName) = ArrAddDistrictRow(i)(2)
            RsDistrict.Fields(fDistrictLongitude) = ArrAddDistrictRow(i)(3)
            RsDistrict.Fields(fDistrictLatitude) = ArrAddDistrictRow(i)(4)
        RsDistrict.Update

    Next

    CloseConnRs conn, RsDistrict
        
End Function

Public Function ArrAddDistrictRowDefault() As Variant
    '区县默认信息
    Dim ArrAdd
    Dim ArrAddRow
    Dim Region As String
    Dim i As Long
    
    ArrAdd = Split(AddressDistrict, ";")

    ReDim ArrAddRow(0 To UBound(ArrAdd))

    For i = 0 To UBound(ArrAdd)
        ArrAddRow(i) = Split(ArrAdd(i), ",")
    Next
    ArrAddDistrictRowDefault = ArrAddRow
End Function

Public Function DataTableOrg()
' 根据业务逻辑生成 组织表
    Dim i As Long
    Dim ArrAddOrg
    Dim ArrAddOrgRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsOrg  As Object
    Dim RsShop As Object
    Dim RsEmployee As Object
    Dim myRnd As Double
    Dim maxOrgID As Long
    Dim dateOpen As Date
    Dim employeeName As String, employeeGender As String, employeeJobTitle As String, employeeGrade As String, employeeEdu As String, employeeOrgID As Long, employeeBirthday As Date, employeeEntryDate As Date
    Dim dictAllDate As Object, dictGender As Object
    
    Set dictGender = GenderDict() '前面 448 个人性根据头像锁定
    Set dictAllDate = DateStatusDict() '所有日期状态的字典
    
    InitE
    
    Set conn = CreateConnection
    Set RsOrg = CreateRecordset(conn, tbNameOrg)
    Set RsShop = CreateRecordset(conn, tbNameShop)
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)

'=====================================================================================
'一级部门 和 销售大区
    Const org As String = "焦棚子科技有限公司,   ,      总部,       10001;" & _
                          "总经理办公室,        1,      总经办,     10002;" & _
                          "产品研发中心,        1,      产品,       10003;" & _
                          "采购中心,            1,      采购,       10004;" & _
                          "销售中心,            1,      销售,       10005;" & _
                          "人力资源中心,        1,      人资,       10006;" & _
                          "售后服务中心,        1,      售后,       10007;" & _
                          "财务中心,            1,      财务,       10008;" & _
                          "东部销售大区,        5,      东区,       10009;" & _
                          "西部销售大区,        5,      西区,       10010;" & _
                          "南部销售大区,        5,      南区,       10011;" & _
                          "北部销售大区,        5,      北区,       10012;" & _
                          "中部销售大区,        5,      中区,       10013;" & _
                          "港澳台销售大区,      5,      港澳台,     10014"

    ArrAddOrg = Split(org, ";")

    ReDim ArrAddOrgRow(0 To UBound(ArrAddOrg))

    For i = 0 To UBound(ArrAddOrg)
        ArrAddOrgRow(i) = Split(ArrAddOrg(i), ",")
    Next

    rows = UBound(ArrAddOrgRow)

    For i = 0 To rows
        RsOrg.AddNew
            RsOrg.Fields(fOrgNameAll) = Trim(ArrAddOrgRow(i)(0))
            If Trim(ArrAddOrgRow(i)(1)) <> "" Then RsOrg.Fields(fOrgParentID) = Trim(ArrAddOrgRow(i)(1))
            RsOrg.Fields(fOrgName) = Trim(ArrAddOrgRow(i)(2))
            RsOrg.Fields(fOrgEmployeeID) = Trim(ArrAddOrgRow(i)(3))
        RsOrg.Update
    Next
'=====================================================================================
'省级销售区域
    
    ArrAddOrg = Split(AddressProvince, ";")

    ReDim ArrAddOrgRow(0 To UBound(ArrAddOrg))

    For i = 0 To UBound(ArrAddOrg)
        ArrAddOrgRow(i) = Split(ArrAddOrg(i), ",")
    Next

    rows = UBound(ArrAddOrgRow)
    
    For i = 0 To rows
        '===============================员工信息
        myRnd = Rnd()
        employeeName = generateName(myRnd)
        If myRnd < 0.7 Then employeeGender = "女" Else employeeGender = "男"
        employeeJobTitle = JobTitlesArr(10)
        If myRnd < 0.8 Then
            employeeGrade = GradeArr(3)
        Else
            employeeGrade = GradeArr(4)
        End If
        employeeEdu = EduArr(Round(Rnd() * 2, 0))
        employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 7500, 0)
        employeeEntryDate = MinDateOpen - Round(Rnd() * 50, 0)
        
        AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender '组织ID待定

        '===============================组织
        RsOrg.AddNew
            RsOrg.Fields(fOrgNameAll) = "省级销售区域" + Trim(ArrAddOrgRow(i)(4))
            RsOrg.Fields(fOrgParentID) = Trim(ArrAddOrgRow(i)(0))
            RsOrg.Fields(fOrgName) = Trim(ArrAddOrgRow(i)(4))
        RsOrg.Update
        
        '===============================交换ID
        RsOrg.Fields(fOrgEmployeeID) = RsEmployee.Fields(fEmployeeID)
        RsOrg.Update
        RsEmployee.Fields(fEmployeeOrgID) = RsOrg.Fields(fOrgID)
        RsEmployee.Update
        maxOrgID = RsOrg.Fields(fOrgID)
                
    Next
'=====================================================================================
'门店组织
    InitPO '初始化
    
    RsShop.MoveFirst
    Do Until RsShop.EOF
        '赋值门店的组织ID
        maxOrgID = maxOrgID + 1
        RsShop.Fields(fShopID) = maxOrgID
        dateOpen = RsShop.Fields(fShopOpenDate)
        RsShop.Update
        
        '组织新增
        RsOrg.AddNew
            RsOrg.Fields(fOrgNameAll) = "销售门店-" & RsShop.Fields(fShopName)
            RsOrg.Fields(fOrgParentID) = ProvinceID2OrgIDDict(CInt(Left(RsShop.Fields(fShopDistrictID), 2))) '通过门店 区县ID 的前两位获取 上级组织ID
            RsOrg.Fields(fOrgName) = RsShop.Fields(fShopName)
        RsOrg.Update
        
        '门店负责人新增
        myRnd = Rnd()
        employeeName = generateName(myRnd)
        If myRnd < 0.7 Then employeeGender = "女" Else employeeGender = "男"
        employeeJobTitle = JobTitlesArr(11)
        employeeGrade = GradeArr(4)
        employeeEdu = EduArr(1 + Round(Rnd() * 2, 0)) '学历要求降低
        employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 6000, 0) '更年轻化
        employeeEntryDate = dateOpen - Round(Rnd() * 30, 0)
        
        AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender '组织ID待定
   
        '===============================交换ID
        RsOrg.Fields(fOrgEmployeeID) = RsEmployee.Fields(fEmployeeID)
        RsOrg.Update
        
        RsEmployee.Fields(fEmployeeOrgID) = RsOrg.Fields(fOrgID)
        RsEmployee.Update
        
        RsShop.MoveNext
    Loop
    
    CloseConnRs conn, RsOrg, RsShop, RsEmployee
        
End Function

Public Function DataTableProduct()
' 根据业务逻辑生成 产品表

    Dim i As Long
    Dim myRnd As Double
    Dim price As Double
    Dim cost As Double

    Dim conn  As Object
    Dim RsProduct As Object
    
    Set conn = CreateConnection
    Set RsProduct = CreateRecordset(conn, tbNameProduct)
    
    For i = 1 To productQuantity
        RsProduct.AddNew
        RsProduct.Fields(fProductID) = "SKU_" & Format(i, "000000")
        Randomize
        myRnd = Rnd()
        
        RsProduct.Fields(fProductCategory) = Chr(Round(myRnd * 9, 0) + 65) & "类"
        
        RsProduct.Fields(fProductName) = "产品" & Chr(Round(myRnd * 9, 0) + 65) & "" & Format(i, "0000")
 
        price = 1000 + myRnd * 5000

        If myRnd < 0.28 Then
            cost = price * 0.28
        ElseIf myRnd > 0.7 Then
            cost = price * myRnd * 0.8
        Else
            cost = price * myRnd
        End If

        RsProduct.Fields(fProductPrice) = Round(price, 0)
        RsProduct.Fields(fProductCostPrice) = Round(cost, 0)
        RsProduct.Update
    Next

    CloseConnRs conn, RsProduct

End Function

Public Function DataTableShop()
' 根据业务逻辑生成 门店表
    Dim i As Long
    Dim k As Long
    Dim myRnd As Double
    Dim ArrAddressDistrict
    Dim ArrAddressDistrictRow
    Dim ArrDictName
    Dim ArrDefault7 '默认手动输入 直辖市+港澳台优先命中
    Dim DictName As Object

    Dim addUB0 As Long

    Dim dateOpen As Date
    Dim dateClose As Date
    Dim conn  As Object
    Dim RsShop As Object


    Set conn = CreateConnection
    Set RsShop = CreateRecordset(conn, tbNameShop)
    
    MinDateOpen = Format(Now, "YYYY-MM-DD")

    ArrAddressDistrictRow = ArrAddDistrictRowDefault()

    
    Set DictName = CreateObject("Scripting.Dictionary") '随机店名字，字典键名保证唯一不重复。
    For i = 1 To 17576 '26*26*26
        DictName(Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & "店") = i
        If DictName.Count = ShopQuantity Then
            Exit For
        End If
    Next
    
    ArrDictName = DictName.Keys
    Set DictName = Nothing
    
    ArrDefault7 = Array( _
                Array(110101, "东城区", 39.917548, 116.418758), _
                Array(120101, "和平区", 39.118328, 121.490318), _
                Array(310101, "黄浦区", 31.222778, 121.471518), _
                Array(500103, "渝中区", 29.556748, 106.562888), _
                Array(710000, "台湾", 25.044518, 121.509518), _
                Array(810001, "中西区", 22.28198088, 114.1543738), _
                Array(820001, "花地玛堂区", 22.207878, 113.5528958) _
                )
            
    '优先命中四个直辖市 + 港澳台
    If ShopQuantity < 8 Then
        For k = 0 To ShopQuantity - 1
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD")
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen '取最小是日期
            
            AddShopRecord RsShop, ArrDictName(k), dateOpen, CLng(ArrDefault7(k)(0)), CStr(ArrDefault7(k)(1)), Round(ArrDefault7(k)(2), 6), Round(ArrDefault7(k)(3), 6)
            
        Next
    End If
    
    '命中前面七个城市后在生成大于 7 的数据。
    If ShopQuantity > 7 Then
    
        For k = 0 To 6
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD")
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen '取最小是日期
            AddShopRecord RsShop, ArrDictName(k), dateOpen, CLng(ArrDefault7(k)(0)), CStr(ArrDefault7(k)(1)), Round(ArrDefault7(k)(2), 6), Round(ArrDefault7(k)(3), 6)

        Next
    
        For i = 8 To ShopQuantity
            Randomize
            myRnd = Rnd()
            
            Randomize
            addUB0 = Round(UBound(ArrAddressDistrictRow) * Rnd(), 0)
            Randomize
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD") '+28容错Dict3N
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen '取最小是日期
            Randomize
            dateClose = Format(dateOpen + 550 + 4320 * Rnd(), "YYYY-MM-DD") '550表示至少1.5年才能关店
    
            If dateClose > Now Then
                AddShopRecord RsShop, ArrDictName(i - 1), dateOpen, CLng(ArrAddressDistrictRow(addUB0)(1)), CStr(ArrAddressDistrictRow(addUB0)(2)), Round(ArrAddressDistrictRow(addUB0)(3) + Rnd() * 0.05, 6), Round(ArrAddressDistrictRow(addUB0)(4) + Rnd() * 0.05, 6)
            Else '闭店
                AddShopRecord RsShop, ArrDictName(i - 1), dateOpen, CLng(ArrAddressDistrictRow(addUB0)(1)), CStr(ArrAddressDistrictRow(addUB0)(2)), Round(ArrAddressDistrictRow(addUB0)(3) + Rnd() * 0.05, 6), Round(ArrAddressDistrictRow(addUB0)(4) + Rnd() * 0.05, 6), dateClose
            End If
        Next
    End If

    CloseConnRs conn, RsShop
        
End Function

Public Function AddShopRecord(ByRef RsShop As Object, ByVal ShopName As String, ShopOpenDate As Date, ShopDistrictID As Long, ShopDistrict As String, ShopLongitude As Double, ShopLatitude As Double, Optional ByVal ShopCloseDate As Date)
    '抽取门店新增函数
    RsShop.AddNew
        RsShop.Fields(fShopName) = ShopName
        RsShop.Fields(fShopOpenDate) = ShopOpenDate
        RsShop.Fields(fShopDistrictID) = ShopDistrictID
        RsShop.Fields(fShopDistrict) = ShopDistrict
        RsShop.Fields(fShopLongitude) = ShopLongitude
        RsShop.Fields(fShopLatitude) = ShopLatitude
        If ShopCloseDate <> CDate(0) Then RsShop.Fields(fShopCloseDate) = ShopCloseDate
    RsShop.Update
End Function

Public Function DataTableShopRD()
    '生成租赁和装修数据
    Dim ShopRentalArea As Double
    Dim ShopRentalPrice As Double
    Dim dateRentalStart As Date
    Dim dateRentalEnd As Date

    Dim dateRsShopDecorationStart As Date
    Dim dateRsShopDecorationEnd As Date
    Dim depreciationPeriod As Long
    Dim depreciationEndDate As Date
    Dim decorationAmount As Double
    
    Dim conn  As Object
    Dim RsShop As Object
    Dim RsShopRental As Object
    Dim RsShopDecoration As Object
    
    Set conn = CreateConnection
    Set RsShop = CreateRecordset(conn, tbNameShop)
    Set RsShopRental = CreateRecordset(conn, tbNameShopRental)
    Set RsShopDecoration = CreateRecordset(conn, tbNameShopDecoration)
    RsShop.MoveFirst
    
    Do Until RsShop.EOF
        dateRentalStart = RsShop.Fields(fShopOpenDate) - 30 - Round(Rnd * 15, 0) '首次租赁开始日期
        dateRsShopDecorationStart = dateRentalStart + Round(Rnd * 7, 0) '首次装修开始日期Format(Now, "YYYY-MM-DD")
        ShopRentalArea = 600 + Rnd * 600 '租赁面积
        ShopRentalPrice = 40 + Rnd * 40 '首次租赁价格
        decorationAmount = ShopRentalArea * 1000 * (0.8 + (Rnd() * 0.3)) '装修金额
        depreciationPeriod = 3 + Round(Rnd * 2, 0) '折旧年限
        
        If IsNull(RsShop.Fields(fShopCloseDate)) Then
Rental: '租赁
            dateRentalEnd = dateRentalStart + 3 * 365 '租赁到期日期
                AddRsShopRentalRecord RsShopRental, RsShop.Fields(fShopID), Round(ShopRentalArea, 2), Round(ShopRentalPrice, 2), dateRentalStart, dateRentalEnd, Round(Rnd * 0.05, 2)
            If dateRentalEnd < Now Then
                dateRentalStart = dateRentalEnd + 1
                ShopRentalPrice = ShopRentalPrice * 0.9 + (Rnd() * 0.2)
                GoTo Rental
            End If
        Else
            AddRsShopRentalRecord RsShopRental, RsShop.Fields(fShopID), Round(ShopRentalArea, 2), Round(ShopRentalPrice, 2), dateRentalStart, RsShop.Fields(fShopCloseDate), Round(Rnd * 0.05, 2)
        End If
        

        If IsNull(RsShop.Fields(fShopCloseDate)) Then '未关店
Decoration: '装修
            dateRsShopDecorationEnd = Format(dateRsShopDecorationStart + Round(45 * (0.8 + (Rnd() * 0.3)), 0), "YYYY-MM-DD") '装修结束日期
            depreciationEndDate = Format(dateRsShopDecorationEnd + depreciationPeriod * 365)  '装修折旧结束日期
            AddShopDecorationRecord RsShopDecoration, RsShop.Fields(fShopID), dateRsShopDecorationStart, dateRsShopDecorationEnd, Round(decorationAmount, 2), depreciationPeriod

            If depreciationEndDate < Now Then
                dateRsShopDecorationStart = depreciationEndDate + 1
                decorationAmount = decorationAmount * 0.9 + (Rnd() * 0.2)
                GoTo Decoration
            End If
        Else
            dateRsShopDecorationEnd = Format(dateRsShopDecorationStart + Round(45 * (0.8 + (Rnd() * 0.3)), 0), "YYYY-MM-DD") '装修结束日期
            depreciationPeriod = Int((RsShop.Fields(fShopCloseDate) - dateRsShopDecorationEnd) / 365)
            AddShopDecorationRecord RsShopDecoration, RsShop.Fields(fShopID), dateRsShopDecorationStart, dateRsShopDecorationEnd, Round(decorationAmount, 2), depreciationPeriod
        End If
            

    RsShop.MoveNext
    Loop
End Function

Public Function AddRsShopRentalRecord(ByRef RsShopRental As Object, ByVal shopID As Long, Area As Long, price As Double, startDate As Date, endDate As Date, Increase As Double)
    '抽取门店租赁新增函数
    RsShopRental.AddNew
        RsShopRental.Fields(fShopRentalShopID) = shopID
        RsShopRental.Fields(fShopRentalArea) = Area
        RsShopRental.Fields(fShopRentalPrice) = price
        RsShopRental.Fields(fShopRentalStartDate) = startDate
        RsShopRental.Fields(fShopRentalEndDate) = endDate
        RsShopRental.Fields(fShopRentalIncrease) = Increase
    RsShopRental.Update
End Function

Public Function AddShopDecorationRecord(ByRef RsShopDecoration As Object, ByVal shopID As Long, startDate As Date, endDate As Date, amount As Double, Years As Long)
    '抽取门店装修新增函数
    RsShopDecoration.AddNew
        RsShopDecoration.Fields(fShopDecorationShopID) = shopID
        RsShopDecoration.Fields(fShopDecorationStartDate) = startDate
        RsShopDecoration.Fields(fShopDecorationEndDate) = endDate
        RsShopDecoration.Fields(fShopDecorationAmount) = amount
        RsShopDecoration.Fields(fShopDecorationYears) = Years
    RsShopDecoration.Update
End Function

Public Function DataTableCustomer()
' 根据业务逻辑生成 客户表
    Dim registerDays As Long
    Dim row As Long
    Dim iCount As Long
    Dim i As Long
    Dim myRnd As Double
    Dim myRndHY As Double
    Dim myRndZY As Double
    Dim customerN As Double
    
    Dim arrHY '行业
    Dim arrZY '职业
    Dim arrRndNL '年龄分布
    
    Dim arrRndHY
    Dim arrRndZY
    
    Dim arrShop
    Dim Rrow As Long
    Dim Rcol As Long

    Dim name As String
    Dim gender As String
    Dim dateBirthday As Date
    Dim dateRegister As Date

    Dim conn  As Object
    Dim RsShop  As Object
    Dim RsCustomer  As Object

    arrHY = Array("建筑业", "制造业", "互联网", "农业", "餐饮", "物流", "汽车") '行业
    arrZY = Array("个体户", "HR", "运营", "IT", "财务", "销售", "研发") '职业
    arrRndHY = Array(0.2, 0.5, 0.5, 0.8, 0.8, 1, 0.9) '行业分布
    arrRndZY = Array(0.3, 0.7, 0.6, 1, 1, 0.8, 0.1) '职业分布
    arrRndNL = Array(0, 0.1, 0.2, 0.3, 0.3, 0.3, 0.3, 0.8, 0.8, 0.8, 0.9, 1) '年龄分布


    Set conn = CreateConnection
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '按照店铺规模注册客户
    
    Set RsShop = CreateRecordset(conn, tbNameShop)
    RsShop.MoveFirst

    Rrow = RsShop.RecordCount - 1
    Rcol = RsShop.Fields.Count - 1

    ReDim arrShop(0 To Rrow, 0 To Rcol)

    For row = 0 To Rrow
        For i = 0 To Rcol
            arrShop(row, i) = RsShop(i)
        Next
        RsShop.MoveNext
    Next
    RsClose RsShop

    '=====================================================================================
    
    Set RsCustomer = CreateRecordset(conn, tbNameCustomer)
    
    iCount = 0
    
    For row = 0 To UBound(arrShop)
        Randomize
        customerN = (3 + 8 * Rnd()) '客户数量控制系数
        If IsNull(arrShop(row, 7)) Then
            registerDays = Round((Now() - arrShop(row, 2)) * customerN, 0)
        Else
            registerDays = Round((arrShop(row, 7) - arrShop(row, 2)) * customerN, 0)
        End If

        For i = 1 To registerDays
            iCount = iCount + 1
            Randomize
            myRnd = Rnd()

            Randomize
            dateBirthday = Format(Now - 7500 - Round((arrRndNL(iCount Mod 12) + Rnd()) * 7000, 0), "YYYY-MM-DD")  '生日
            dateRegister = Format(Now - 1500 + Round((arrRndNL(iCount Mod 12) + Rnd()) * 750, 0), "YYYY-MM-DD") '注册时间，比开店少两天，不会业务逻辑溢出
    
            If myRnd < 0.8 Then
                name = generateName(myRnd)
                gender = "男"
            Else
                name = generateName(myRnd)
                gender = "女"
            End If
    
            RsCustomer.AddNew
            RsCustomer.Fields(fCustomerID) = "CC_" & Format(iCount, "0000000")
            RsCustomer.Fields(fCustomerName) = name
            RsCustomer.Fields(fCustomerBirthday) = dateBirthday
            RsCustomer.Fields(fCustomerGender) = gender
            RsCustomer.Fields(fCustomerRegister) = dateRegister
            Randomize
            RsCustomer.Fields(fCustomerIndustry) = arrHY(Round(Rnd() * arrRndHY(row Mod 7) * 6, 0))
            Randomize
            RsCustomer.Fields(fCustomerOccupation) = arrZY(Round(Rnd() * arrRndZY(row Mod 7) * 6, 0))
            RsCustomer.Update
        Next

    Next

    CloseConnRs conn, RsCustomer
    
End Function

Public Function DataTableSOS()
' 根据业务逻辑生成 入库表、订单主表、订单子表
    
        Dim iStorage As Long, rowShop As Long, iCustomer As Long, iStorageN As Long, dayN As Long
        Dim i As Long, k As Long, productQuantity As Long, discount As Double, n As Long, numOrder As Long, myRnd As Double

        Dim rowsProduct As Long
        Dim rowsShop As Long
        Dim rowsCustomer As Long

        Dim rowsProductRnd As Long
        Dim rowsCustomerRnd As Long

        Dim orderID As String
        Dim OcNumber As Long
        Dim Qd As String
        Dim Yyts As Long '营业天数
        Dim dateOpen As Date

        Dim arrProduct
        Dim arrShop
        Dim arrCustomer
        Dim arrHY '行业
        Dim arrZY '职业
        Dim arrDjjsxs '单均件数系数
        Dim arrDdslxsMonth '行业淡旺季趋势
        Dim arrDdslxsSC '区域系数
        Dim arrDictStorage
        Dim arrProductRnd
        Dim arrZK '折扣
        Dim arrZKMonth '折扣月份分布
        Dim arrRndKF '客户分布
        Dim Rrow As Long
        Dim Rcol As Long


        Dim conn As Object
        Dim RsProduct As Object         '产品
        Dim RsShop As Object            '门店
        Dim RsCustomer As Object        '客户
        Dim RsStorage As Object         '入库
        Dim RsOrder As Object           '订单
        Dim RsOrdersub As Object        '订单子表
        
        Dim dictCustomerHY As Object
        Dim dictCustomerZY As Object
        Dim dictStorageN As Object
        Dim dictStorage As Object
        Dim dictProductRnd As Object
        Dim arrPerson As Variant
        Dim orderDate As Date, pc As Long, personCount As Long, orderCount As Long
        
        Set conn = CreateConnection

    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Set RsProduct = CreateRecordset(conn, tbNameProduct)
        RsProduct.MoveFirst

        Rrow = RsProduct.RecordCount - 1
        Rcol = RsProduct.Fields.Count - 1

        ReDim arrProduct(0 To Rrow, 0 To Rcol)

        For i = 0 To Rrow
            For k = 0 To Rcol
                arrProduct(i, k) = RsProduct(k)
            Next
            RsProduct.MoveNext
        Next
        RsClose RsProduct
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        Set RsShop = CreateRecordset(conn, tbNameShop)
        RsShop.MoveFirst

        Rrow = RsShop.RecordCount - 1
        Rcol = RsShop.Fields.Count - 1

        ReDim arrShop(0 To Rrow, 0 To Rcol)

        For i = 0 To Rrow
            For k = 0 To Rcol
                arrShop(i, k) = RsShop(k)
            Next
            RsShop.MoveNext
        Next
        
        RsClose RsShop
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        arrHY = Array("保险", "互联网", "汽车", "制造业") '折扣行业准备
        arrZY = Array("HR", "财务", "销售", "运营")
        Set dictCustomerHY = CreateObject("Scripting.Dictionary") '行业
        Set dictCustomerZY = CreateObject("Scripting.Dictionary") '职业
        
        For i = 0 To UBound(arrHY)
            dictCustomerHY(arrHY(i)) = arrHY(i)
        Next
        
        For i = 0 To UBound(arrZY)
            dictCustomerZY(arrZY(i)) = arrZY(i)
        Next

        Set RsCustomer = CreateRecordset(conn, tbNameCustomer)
        RsCustomer.MoveFirst

        Rrow = RsCustomer.RecordCount - 1
        Rcol = RsCustomer.Fields.Count - 1

        ReDim arrCustomer(0 To Rrow, 0 To 3)

        For i = 0 To Rrow
            arrCustomer(i, 0) = RsCustomer.Fields(fCustomerID)
            arrCustomer(i, 1) = RsCustomer.Fields(fCustomerRegister)
            arrCustomer(i, 2) = RsCustomer.Fields(fCustomerIndustry)
            arrCustomer(i, 3) = RsCustomer.Fields(fCustomerOccupation)
            RsCustomer.MoveNext
        Next
        
        RsClose RsCustomer

        Set RsStorage = CreateRecordset(conn, tbNameStorage)
        Set RsOrder = CreateRecordset(conn, tbNameOrder)
        Set RsOrdersub = CreateRecordset(conn, tbNameOrdersub)

        '=====================================================================================
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        rowsProduct = UBound(arrProduct) '产品数量
        rowsShop = UBound(arrShop) '门店数量
        rowsCustomer = UBound(arrCustomer) '客户数量
        OcNumber = 0
        
        arrDjjsxs = Array(0.7, 0.8, 1, 1.2, 1.3) '单均件数系数，count=5
        arrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '行业淡旺季趋势，count=12
        arrDdslxsSC = Array(0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6) '区域订单系数正太分布，count=34
        arrZK = Array(1, 0.9, 0.8, 0.7, 0.7, 0.6) '折扣信息分布，count=6
        arrZKMonth = Array(0.95, 0.9, 1, 0.98, 0.95, 1, 0.98, 0.98, 0.9, 0.96, 0.92, 0.98) '行业淡旺季趋势，count=12
        arrRndKF = Array(0, 0.1, 0.5, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.8, 0.9, 1) '客户分布

    For rowShop = 0 To rowsShop

            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            '营业天数
            If IsNull(arrShop(rowShop, 7)) Then
                Yyts = Round(Now - arrShop(rowShop, 2), 0)
            Else
                Yyts = Round(arrShop(rowShop, 7) - arrShop(rowShop, 2), 0)
            End If
            '=====================================================================================
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            Set dictStorageN = CreateObject("Scripting.Dictionary") '入库的随机序列
            dictStorageN(1) = Yyts Mod MaxInventoryDays + 1
                    '入库随机时间
                    For i = 1 To Yyts
                        dictStorageN(i + 1) = dictStorageN.item(i) + Round(Rnd() * 2 + 5, 0)
                        If dictStorageN.item(i + 1) > Yyts Then
                            dictStorageN(i + 1) = Yyts
                            Exit For
                        End If
                    Next
            '=====================================================================================

            Set dictStorage = CreateObject("Scripting.Dictionary") '记录入库信息
            iStorageN = 1
            
           arrPerson = salePersonArr(conn, arrShop(rowShop, 0)) '门店销售人员 0 ID,1入职日期,2离职日期或当前日期

            For dayN = 1 To Yyts

                dateOpen = arrShop(rowShop, 2) + dayN - 1
                Randomize
                
                numOrder = Round(Rnd() * 10 * arrDdslxsMonth(Month(dateOpen) - 1) * arrDdslxsSC(arrShop(rowShop, 2) Mod (UBound(arrDdslxsSC) + 1)), 0) '每天订单数量

                If numOrder = 0 Then GoTo NoOrder '没有销售

                For i = 1 To numOrder '每天订单数
                    OcNumber = OcNumber + 1
                    orderID = "OC_" & Format(OcNumber, "0000000")
                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '订单主表写入
                    Randomize
                    myRnd = (Rnd() + arrRndKF(dayN * i Mod 12)) / 2

                    rowsCustomerRnd = Round(rowsCustomer * myRnd, 0)
                    
                '注册与购买分布
                If IsNull(arrShop(rowShop, 7)) Then '未关店
                    If arrCustomer(rowsCustomerRnd, 1) >= arrShop(rowShop, 2) And OcNumber Mod 13 > 6 Then
                        GoTo rowsCustomerRndLable
                    Else
                        For iCustomer = rowsCustomerRnd To rowsCustomer '往右前进
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 13 > 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 > 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                                
                        For iCustomer = rowsCustomerRnd To 0 Step -1 '往左前进
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 13 <= 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 <= 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                    End If
                Else '关店
                    If arrCustomer(rowsCustomerRnd, 1) >= arrShop(rowShop, 2) And arrCustomer(rowsCustomerRnd, 1) < arrShop(rowShop, 7) And OcNumber Mod 13 < 6 Then
                            GoTo rowsCustomerRndLable
                    Else
                        For iCustomer = rowsCustomerRnd To rowsCustomer '往右前进
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And arrCustomer(iCustomer, 1) < arrShop(rowShop, 7) And Month(arrCustomer(iCustomer, 1)) Mod 13 > 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 > 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                            
                        For iCustomer = rowsCustomerRnd To 0 Step -1 '往左前进
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And arrCustomer(iCustomer, 1) < arrShop(rowShop, 7) And Month(arrCustomer(iCustomer, 1)) Mod 13 <= 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And arrCustomer(iCustomer, 1) < arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 <= 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                    End If
                End If
     
rowsCustomerRndLable:

                    If myRnd < 0.7 Then
                        Qd = "线上"
                    Else
                        Qd = "线下"
                    End If
                    orderCount = 0
                    personCount = UBound(arrPerson, 2)
                    orderDate = arrShop(rowShop, 2) + dayN - 1
rndSalePerson: '随机销售人员
                    pc = Round(personCount * Rnd, 0)
                    If orderDate < arrPerson(1, pc) Or arrPerson(2, pc) < orderDate Then
                        
                        orderCount = orderCount + 1
                        If orderCount > 200 Then GoTo NoOrder '保证无死循环
                        GoTo rndSalePerson '需要满足当前日期在员工的在职期间

                    End If
                    
                    RsOrder.AddNew
                        RsOrder.Fields(fOrderID) = orderID
                        RsOrder.Fields(fOrderShopID) = arrShop(rowShop, 0)
                        RsOrder.Fields(fOrderDate) = orderDate
                        RsOrder.Fields(fOrderSentDate) = arrShop(rowShop, 2) + dayN + Round(4 * myRnd + 8, 0)
                        RsOrder.Fields(fOrderCustomerID) = arrCustomer(rowsCustomerRnd, 0)
                        RsOrder.Fields(fOrderType) = Qd
                        RsOrder.Fields(fOrderEmployeeID) = arrPerson(0, pc)
                    RsOrder.Update
                    '=====================================================================================

                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '订单子表

                    Randomize

                    k = Round(5 * Rnd(), 0) + 1 '计划每个订产品数量上限,均值为3

                    Set dictProductRnd = CreateObject("Scripting.Dictionary")
                    For n = 1 To k
                        If k < 4 Then
                            rowsProductRnd = Round(rowsProduct * Rnd() / 5, 0)  '往左偏移
                        Else
                            rowsProductRnd = Round(rowsProduct * Rnd(), 0)
                        End If
                        dictProductRnd(rowsProductRnd) = rowsProductRnd '字典去重sku 的 ID
                    Next

                    arrProductRnd = dictProductRnd.Keys

                    For n = 0 To dictProductRnd.Count - 1

                        productQuantity = Round(5 * Rnd() * arrDjjsxs(rowShop Mod 5) * arrDjjsxs(arrProductRnd(n) Mod 5), 0) + 1 '件数系数加权

                        'q产品折扣
                        
                        If rowShop Mod 40 > 30 Then '区域
                            discount = Round(arrZK(0) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf rowShop Mod 40 < 10 Then
                            discount = Round(arrZK(5) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf arrProductRnd(n) Mod 8 < 1 Then '产品
                            discount = Round(arrZK(1) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf arrProductRnd(n) Mod 8 > 5 Then
                            discount = Round(arrZK(3) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf dictCustomerHY.Exists(arrCustomer(rowsCustomerRnd, 2)) Then  '客户
                            discount = Round(arrZK(2) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf dictCustomerZY.Exists(arrCustomer(rowsCustomerRnd, 3)) Then
                            discount = Round(arrZK(4) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        Else
                            discount = 1
                        End If

ExitIFZhekou:
                        RsOrdersub.AddNew
                            RsOrdersub.Fields(fOrdersubOrderID) = orderID
                            RsOrdersub.Fields(fOrdersubProductID) = arrProduct(arrProductRnd(n), 0)
                            RsOrdersub.Fields(fOrdersubPrice) = arrProduct(arrProductRnd(n), 3)
                            RsOrdersub.Fields(fOrdersubDiscount) = discount
                            RsOrdersub.Fields(fOrdersubQuantity) = productQuantity
                            RsOrdersub.Fields(fOrdersubAmount) = Round(arrProduct(arrProductRnd(n), 3) * productQuantity * discount, 2)
                        RsOrdersub.Update
                        dictStorage(arrProduct(arrProductRnd(n), 0)) = dictStorage(arrProduct(arrProductRnd(n), 0)) + productQuantity
                    Next
                Set dictProductRnd = Nothing
                '=====================================================================================
            Next

NoOrder: '当日无订单跳转
                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                '生成入库信息
                If dictStorageN.item(iStorageN) = dayN And dayN < Yyts Then
                    iStorageN = iStorageN + 1

                    arrDictStorage = dictStorage.Keys

                    For iStorage = 0 To UBound(arrDictStorage)
                        RsStorage.AddNew
                            RsStorage.Fields(fStorageProductID) = arrDictStorage(iStorage)
                            RsStorage.Fields(fStorageQuantity) = dictStorage.item(arrDictStorage(iStorage))
                            RsStorage.Fields(fStorageShopID) = arrShop(rowShop, 0)
                            RsStorage.Fields(fStorageDate) = arrShop(rowShop, 2) + dayN - 1 '-1 保证有库存
                        RsStorage.Update
                    Next
                    Set dictStorage = Nothing
                    Set dictStorage = CreateObject("Scripting.Dictionary") '记录入库信息
                    GoTo Rk0
                    
                ElseIf dayN = Yyts Then  '保证最后一次入库累计大于0

                    iStorageN = iStorageN + 1

                    arrDictStorage = dictStorage.Keys

                    For iStorage = 0 To UBound(arrDictStorage)
                        RsStorage.AddNew
                            RsStorage.Fields(fStorageProductID) = arrDictStorage(iStorage)
                            RsStorage.Fields(fStorageQuantity) = dictStorage.item(arrDictStorage(iStorage)) + Round(Rnd() * 5, 0)
                            RsStorage.Fields(fStorageShopID) = arrShop(rowShop, 0)
                            RsStorage.Fields(fStorageDate) = arrShop(rowShop, 2) + dayN - 1 '-1 保证有库存
                        RsStorage.Update
                    Next
                    Set dictStorage = Nothing
                    Set dictStorage = CreateObject("Scripting.Dictionary") '记录入库信息
                    GoTo Rk0
                    
                End If
                '=====================================================================================
Rk0:
            Next

        Next

    CloseConnRs conn, RsStorage, RsOrder, RsOrdersub

End Function

Public Function salePersonArr(ByRef conn As Object, ByVal shopID As Long)
    '返回销售人员按照日期作为键名的人员字典，人员字典的键名为员工ID
    
    Dim EmployeeID As Long, i As Long
    Dim Arr As Variant
    Dim rows As Long, strSQL As String
    Dim RsEmployee As Object

    ' 构建 SQL 查询语句
    strSQL = "SELECT * FROM " & tbNameEmployee & " WHERE " & fEmployeeOrgID & " = " & shopID & ";"

    ' 创建记录集对象
    Set RsEmployee = CreateObject("ADODB.Recordset")

    ' 执行查询，并将结果存储在记录集对象中
    With RsEmployee
        .ActiveConnection = conn
        .source = strSQL
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    rows = RsEmployee.RecordCount - 1
    
    If RsEmployee.RecordCount >= 0 Then
    
        ReDim Arr(0 To 2, 0 To rows)
        RsEmployee.MoveFirst
        i = 0
        Do Until RsEmployee.EOF
            Arr(0, i) = RsEmployee.Fields(fEmployeeID)
            Arr(1, i) = RsEmployee.Fields(fEmployeeEntryDate)
            If IsNull(RsEmployee.Fields(fEmployeeResignationDate)) Then
                Arr(2, i) = CDate(Format(Now, "YYYY-MM-DD"))
            Else
                Arr(2, i) = RsEmployee.Fields(fEmployeeResignationDate)
            End If
            i = i + 1
        RsEmployee.MoveNext
        Loop
    End If
    salePersonArr = Arr

End Function
Public Function DataTableLaborCost()
    '根据业务逻辑生成人工成本

    Dim EmployeeID As Long, employeeOrgID As Long, employeeGrade As String, employeeEdu As String, employeeEntryDate As Date, employeeResignationDate As Date
    Dim daysFirst As Long, daysMonthFirst As Long, daysLast As Long, daysMonthLast As Long, numMonth As Long, i As Long, dateMonthStartAC As Date, dateMonthStartFirst As Date
    Dim salaryDown As Long, salaryUp As Long, yearsService As Long, eduN As Double, salaryRnd As Long, salaryRndBase As Long, salary As Long
    
    Dim ArrEmployee As Variant, ArrOrgIDDate As Variant
    Dim rows As Long, key As Variant
    Dim RsEmployee As Object, laborCostDict As Object
    Dim conn As Object, Employee As Object, RsLaborCost As Object
    Dim orgID As Long, monthStart As Date, amount As Long
    Const delimiter As String = "|"
    
    Set laborCostDict = CreateObject("Scripting.Dictionary")
    
    Set conn = CreateConnection
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)
    Set RsLaborCost = CreateRecordset(conn, tbNameLaborCost)
    
    InitE
    
    rows = RsEmployee.RecordCount - 1
    
    If RsEmployee.RecordCount >= 0 Then
    
        ReDim ArrEmployee(0 To 5, 0 To rows) '员工ID，组织ID，职级,学历,入职日期,离职日期
        RsEmployee.MoveFirst
        i = 0
        Do Until RsEmployee.EOF
            EmployeeID = RsEmployee.Fields(fEmployeeID)
            employeeOrgID = RsEmployee.Fields(fEmployeeOrgID)
            employeeGrade = RsEmployee.Fields(fEmployeeGrade)
            employeeEdu = RsEmployee.Fields(fEmployeeEdu)
            employeeEntryDate = RsEmployee.Fields(fEmployeeEntryDate)
            
            If IsNull(RsEmployee.Fields(fEmployeeResignationDate)) Then
                employeeResignationDate = GetMonthStart(CDate(Format(Now, "YYYY-MM-DD"))) - 1 '当前日期的上月底
            Else
                employeeResignationDate = RsEmployee.Fields(fEmployeeResignationDate)
            End If

        daysFirst = day(employeeEntryDate)
        daysMonthFirst = GetDaysInMonth(employeeEntryDate)
        
        daysLast = day(employeeResignationDate)
        daysMonthLast = GetDaysInMonth(employeeResignationDate)
        
        numMonth = DateDiffInMonths(employeeEntryDate, employeeResignationDate)
        dateMonthStartFirst = GetMonthStart(employeeEntryDate)
        
        salaryUp = GradeSalaryDict(employeeGrade)(0) '工资上限
        salaryDown = GradeSalaryDict(employeeGrade)(1) '工资下限
        eduN = EduSalaryDict(employeeEdu) '学历系数
        salaryRndBase = Round((salaryUp - salaryDown + 1) * Rnd + salaryDown, 0) '工资基数
        
        For i = 0 To numMonth
                    
            If i = 0 Then
                salaryRnd = Round(salaryRndBase * daysFirst / daysMonthFirst, 0) '首月判断工资天数
            ElseIf i = numMonth Then
                salaryRnd = Round(salaryRndBase * daysLast / daysMonthLast, 0) '末月判断工资天数
            Else
                salaryRnd = salaryRndBase * (0.8 + (Rnd * (1.2 - 0.8))) * 1.36  '社保公积金系数 0.8 到 1.2的浮动
            End If
            
            
            dateMonthStartAC = AddMonths(dateMonthStartFirst, i) '工资月份
            
            yearsService = Round(i / 12, 0) '司龄
            
            salary = salaryRnd * eduN * (1 + yearsService * 0.1) '当月工资

            key = CStr(employeeOrgID & delimiter & dateMonthStartAC) '键名

            AddDictByKey laborCostDict, key, salary '按照组织ID，月份累计成本
        Next
        
        RsEmployee.MoveNext
        Loop
    End If
    
    For Each key In laborCostDict.Keys
    
        ArrOrgIDDate = Split(key, delimiter)
        orgID = ArrOrgIDDate(0)
        monthStart = ArrOrgIDDate(1)
        amount = laborCostDict(key)
        If amount > 0 Then
            RsLaborCost.AddNew
                RsLaborCost.Fields(fLaborCostOrgID) = orgID
                RsLaborCost.Fields(fLaborCostMonth) = monthStart
                RsLaborCost.Fields(fLaborCostAmount) = amount
            RsLaborCost.Update
        End If
    Next key
    
End Function

Public Function DataTableSaleTarget()
' 根据业务逻辑生成 销售目标表
    Dim Sqlstr As String
    Dim ArrYQ
    Dim arrDdslxsMonth '行业淡旺季趋势
    Dim Qn As Double '最后三个月的系数和
    Dim B As Double
    Dim UP0 As Double '拉开区域差距
    
    Dim i As Long
    Dim Rrow As Long
    Dim Rcol As Long
    Dim k As Long
    
    Dim conn  As Object
    Dim RsSaleTarget  As Object

    Set conn = CreateConnection

    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '去年完成情况全年&去年Q4完成情况,月均取大
    Sqlstr = "SELECT" & vbCrLf
    Sqlstr = Sqlstr & "TY.*,TQ.fQuarter" & vbCrLf
    Sqlstr = Sqlstr & "FROM" & vbCrLf
    Sqlstr = Sqlstr & "(" & vbCrLf
    Sqlstr = Sqlstr & "SELECT" & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceID & "," & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceName2 & "," & vbCrLf
    Sqlstr = Sqlstr & "Sum(" & tbNameOrdersub & "." & fOrdersubAmount & ") AS fYear" & vbCrLf
    Sqlstr = Sqlstr & "FROM ((((" & tbNameOrdersub & " " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameOrder & " ON " & tbNameOrdersub & "." & fOrdersubOrderID & " = " & tbNameOrder & "." & fOrderID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameShop & " ON " & tbNameOrder & "." & fOrderShopID & " = " & tbNameShop & "." & fShopID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameDistrict & " ON " & tbNameShop & "." & fShopDistrictID & " = " & tbNameDistrict & "." & fDistrictID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameCity & " ON " & tbNameDistrict & "." & fDistrictCityID & " = " & tbNameCity & "." & fCityID & ")" & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameProvince & " ON " & tbNameCity & "." & fCityProvinceID & " = " & tbNameProvince & "." & fProvinceID & vbCrLf
    Sqlstr = Sqlstr & "WHERE " & tbNameOrder & "." & fOrderDate & ">#" & Format(Now(), "YYYY") - 2 & "-12-1# AND " & tbNameOrder & "." & fOrderDate & "<#" & Format(Now(), "YYYY") & "-1-1#" & vbCrLf
    Sqlstr = Sqlstr & "GROUP BY " & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceID & "," & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceName2 & vbCrLf
    Sqlstr = Sqlstr & ") TY" & vbCrLf
    Sqlstr = Sqlstr & "LEFT JOIN" & vbCrLf
    Sqlstr = Sqlstr & "(" & vbCrLf
    Sqlstr = Sqlstr & "SELECT" & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceID & "," & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceName2 & "," & vbCrLf
    Sqlstr = Sqlstr & "Sum(" & tbNameOrdersub & "." & fOrdersubAmount & ") AS fQuarter" & vbCrLf
    Sqlstr = Sqlstr & "FROM ((((" & tbNameOrdersub & " " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameOrder & " ON " & tbNameOrdersub & "." & fOrdersubOrderID & " = " & tbNameOrder & "." & fOrderID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameShop & " ON " & tbNameOrder & "." & fOrderShopID & " = " & tbNameShop & "." & fShopID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameDistrict & " ON " & tbNameShop & "." & fShopDistrictID & " = " & tbNameDistrict & "." & fDistrictID & ") " & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameCity & " ON " & tbNameDistrict & "." & fDistrictCityID & " = " & tbNameCity & "." & fCityID & ")" & vbCrLf
    Sqlstr = Sqlstr & "INNER JOIN " & tbNameProvince & " ON " & tbNameCity & "." & fCityProvinceID & " = " & tbNameProvince & "." & fProvinceID & vbCrLf
    Sqlstr = Sqlstr & "WHERE " & tbNameOrder & "." & fOrderDate & ">#" & Format(Now(), "YYYY") - 1 & "-9-1# AND " & tbNameOrder & "." & fOrderDate & "<#" & Format(Now(), "YYYY") & "-1-1#" & vbCrLf
    Sqlstr = Sqlstr & "GROUP BY " & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceID & "," & vbCrLf
    Sqlstr = Sqlstr & tbNameProvince & "." & fProvinceName2 & vbCrLf
    Sqlstr = Sqlstr & ") TQ" & vbCrLf
    Sqlstr = Sqlstr & "ON TY." & fProvinceID & "=TQ." & fProvinceID & " AND TY." & fProvinceName2 & "=TQ." & fProvinceName2 & vbCrLf
'    Debug.Print Sqlstr
    
    Set RsSaleTarget = CreateObject("ADODB.Recordset")
    With RsSaleTarget
        .ActiveConnection = conn
        .source = Sqlstr
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With

    Rrow = RsSaleTarget.RecordCount - 1
    Rcol = RsSaleTarget.Fields.Count - 1

    ReDim ArrYQ(0 To Rrow, 0 To Rcol)

    For i = 0 To Rrow
        For k = 0 To Rcol
            ArrYQ(i, k) = RsSaleTarget(k)
        Next
        RsSaleTarget.MoveNext
    Next
    
    RsClose RsSaleTarget

    '=====================================================================================
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    arrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '行业淡旺季趋势,归一，count=12;同上DataTableT345

    Set RsSaleTarget = CreateRecordset(conn, tbNameSaleTarget)

    For i = UBound(arrDdslxsMonth) - 2 To UBound(arrDdslxsMonth)
        Qn = arrDdslxsMonth(i) + Qn
    Next
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '生成去年目标
    For i = 0 To UBound(ArrYQ)

        If IsNull(ArrYQ(i, 3)) Then
            B = ArrYQ(i, 2) / 12
        ElseIf ArrYQ(i, 2) / 12 > ArrYQ(i, 3) / Qn Then '月均取大
            B = ArrYQ(i, 2) / 12
        Else
            B = ArrYQ(i, 3) / Qn
        End If
        
        '实际值都按照1，下方的UP0有拉开差距。
        UP0 = 1
        
        For k = 1 To 12
            RsSaleTarget.AddNew
            RsSaleTarget.Fields(fSaleTargetProvinceID) = ArrYQ(i, 0)
            RsSaleTarget.Fields(fSaleTargetProvinceName2) = ArrYQ(i, 1)
            RsSaleTarget.Fields(fSaleTargetMonth) = Format(Now(), "YYYY") - 1 & "-" & k & "-1"
            Randomize
            RsSaleTarget.Fields(fSaleTargetAmount) = Round(B * (0.7 + Rnd() * 0.1 * UP0) * arrDdslxsMonth(k - 1), 0) '下方今年目标的比例不一样，今年都按照实际浮动不大。
            RsSaleTarget.Update
        Next
    Next
    '=====================================================================================
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '今年目标
    For i = 0 To UBound(ArrYQ)

        If IsNull(ArrYQ(i, 3)) Then
            B = ArrYQ(i, 2) / 12
        ElseIf ArrYQ(i, 2) / 12 > ArrYQ(i, 3) / Qn Then '月均取大
            B = ArrYQ(i, 2) / 12
        Else
            B = ArrYQ(i, 3) / Qn
        End If
        
        If ArrYQ(i, 0) < 5 Then '拉开区域差距
            UP0 = 1.5
        ElseIf ArrYQ(i, 0) < 11 Then
            UP0 = 1.2
        Else
            UP0 = 1
        End If
        
        For k = 1 To 12
            RsSaleTarget.AddNew
            RsSaleTarget.Fields(fSaleTargetProvinceID) = ArrYQ(i, 0)
            RsSaleTarget.Fields(fSaleTargetProvinceName2) = ArrYQ(i, 1)
            RsSaleTarget.Fields(fSaleTargetMonth) = Format(Now(), "YYYY") & "-" & k & "-1"
            Randomize
            RsSaleTarget.Fields(fSaleTargetAmount) = Round(B * (0.6 + Rnd() * 0.2 * UP0) * arrDdslxsMonth(k - 1), 0)
            RsSaleTarget.Update
        Next
    Next


    CloseConnRs conn, RsSaleTarget
    
End Function

Public Function DataTableEmployeeExecutives()
' 根据业务逻辑生成 高管
    Dim i As Long
    Dim ArrAddEmployee
    Dim ArrAddEmployeeRow
    Dim rows As Long

    Dim conn  As Object
    Dim RsEmployee  As Object
    
    Dim EmployeeID As Long, employeeName As String, employeeGender As String, employeeJobTitle As String, employeeGrade As String, employeeEdu As String
    Dim employeeOrgID As Long
    Dim employeeBirthday As Date, employeeEntryDate As Date
    Dim dictAllDate As Object, dictGender As Object
    
    Set dictGender = GenderDict() '前面 448 个人性根据头像锁定
    Set dictAllDate = DateStatusDict() '所有日期状态的字典
    
    InitE '初始化员工信息相关内容
    Set conn = CreateConnection
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)

'=====================================================================================
'一级部门 和 销售大区

    '默认组织 出生日期、入职日期、离职日期、离职原因随机生成
    
    Const employeeStr As String = "总经理,     男, 1,      总经理,        总经理       ;" & _
                                  "巩耕,       男, 2,      总经理助理,    高级总监     ;" & _
                                  "施柴,       男, 3,      产品总监,      高级总监     ;" & _
                                  "幸堤玎,     女, 4,      采购总监,      高级总监     ;" & _
                                  "殷横,       男, 5,      销售总监,      高级总监     ;" & _
                                  "党赋,       男, 6,      人力资源总监,  高级总监     ;" & _
                                  "闻界基,     男, 7,      售后服务总监,  高级总监     ;" & _
                                  "闵款,       女, 8,      财务总监,      高级总监     ;" & _
                                  "欧阳经往,   男, 9,      大区经理,      总监         ;" & _
                                  "焦阿灰,     男, 10,     大区经理,      总监         ;" & _
                                  "左丘垂漫,   女, 11,     大区经理,      总监         ;" & _
                                  "左烈佐,     男, 12,     大区经理,      总监         ;" & _
                                  "焦仔耘,     女, 13,     大区经理,      总监         ;" & _
                                  "安修谊,     男, 14,     大区经理,      总监          "

    ArrAddEmployee = Split(employeeStr, ";")

    ReDim ArrAddEmployeeRow(0 To UBound(ArrAddEmployee))

    For i = 0 To UBound(ArrAddEmployee)
        ArrAddEmployeeRow(i) = Split(ArrAddEmployee(i), ",")
    Next

    rows = UBound(ArrAddEmployeeRow)

    For i = 0 To rows
        
        employeeName = Trim(ArrAddEmployeeRow(i)(0))
        employeeGender = Trim(ArrAddEmployeeRow(i)(1))
        employeeOrgID = Trim(ArrAddEmployeeRow(i)(2))
        employeeJobTitle = Trim(ArrAddEmployeeRow(i)(3))
        employeeGrade = Trim(ArrAddEmployeeRow(i)(4))
        employeeEdu = EduArr(Round(Rnd() * 3, 0))
        employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 8000, 0)
        employeeEntryDate = MinDateOpen - Round(Rnd() * 50, 0)

        AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender, employeeOrgID
    Next

    CloseConnRs conn, RsEmployee
    
End Function

Public Function AddEmployeeRecord(ByRef RsEmployee As Object, ByVal employeeName As String, employeeGender As String, employeeJobTitle As String, employeeGrade As String, employeeEdu As String, employeeBirthday As Date, employeeEntryDate As Date, dictAllDate As Object, dictGender As Object, Optional employeeOrgID As Long)
    '抽取员工新增函数
    Dim EmployeeID As Long, modxR As Long, mod2R As Long, dateStatus As Long, entryDate As Date

    RsEmployee.AddNew
        RsEmployee.Fields(fEmployeeName) = employeeName
        If employeeGender = "女" Then RsEmployee.Fields(fEmployeeGender) = "女"
        If employeeOrgID Then RsEmployee.Fields(fEmployeeOrgID) = employeeOrgID
        RsEmployee.Fields(fEmployeeJobTitle) = employeeJobTitle
        RsEmployee.Fields(fEmployeeGrade) = employeeGrade
        RsEmployee.Fields(fEmployeeEdu) = employeeEdu
        RsEmployee.Fields(fEmployeeBirthday) = employeeBirthday
        RsEmployee.Fields(fEmployeeEntryDate) = employeeEntryDate
    RsEmployee.Update
    
        EmployeeID = RsEmployee.Fields(fEmployeeID)
        mod2R = EmployeeID Mod 2
        
        '统一确认前面448人的性别和提前准备好的头像匹配
        If dictGender.Exists(EmployeeID) Then
            RsEmployee.Fields(fEmployeeGender) = dictGender(EmployeeID)
            RsEmployee.Update
        End If
        
        
        '入职时间不落在员工休息日
        entryDate = employeeEntryDate
        
EffectiveEntryDate: '确保有效的入职日期

        modxR = CLng(dictAllDate(entryDate)("modx"))
        dateStatus = CLng(dictAllDate(entryDate)("status"))

        If modxR <= 4 And mod2R = 1 And dateStatus < 3 Then '员工编号为奇数，余数为：0,1,2,3,4 工作日、补班
            RsEmployee.Fields(fEmployeeEntryDate) = entryDate
            RsEmployee.Update
        ElseIf modxR >= 2 And mod2R = 0 And dateStatus < 3 Then '员工编号为偶数，余数为：2,3,4,5,6 工作日、补班
            RsEmployee.Fields(fEmployeeEntryDate) = entryDate
            RsEmployee.Update
        Else
            entryDate = entryDate + 1
            GoTo EffectiveEntryDate
        End If
                
End Function

Public Function DataTableEmployeeRegular()
' 根据业务逻辑生成 门店销售人员

    Dim dateOpen As Date
    Dim dateClose As Date
    Dim dateResignation As Date
    Dim dateEntry As Date
    Dim conn  As Object
    Dim RsShop  As Object
    Dim RsEmployee  As Object
    Dim myRnd As Double, days As Long
    
    Dim i As Long, seed As Long, workDays As Long, Yyts As Long, numEmployee As Long
    Dim employeeName As String, employeeGender As String, employeeJobTitle As String, employeeGrade As String, employeeEdu As String, employeeOrgID As Long, employeeBirthday As Date, employeeEntryDate As Date
    Dim dictAllDate As Object, dictGender As Object
    
    Set dictGender = GenderDict() '前面448人员的性别字典
    Set dictAllDate = DateStatusDict() '所有日期状态的字典

    InitE '初始化员工信息相关内容
    Set conn = CreateConnection
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)
    Set RsShop = CreateRecordset(conn, tbNameShop)
    
    RsShop.MoveFirst
    Do Until RsShop.EOF
        SetShopRandomSeed RsShop.Fields(fShopDistrictID).value
        numEmployee = Round(2 + Rnd() * 7, 0)
    
        dateOpen = RsShop.Fields(fShopOpenDate)
        
        If IsNull(RsShop.Fields(fShopCloseDate)) Then
            Yyts = Round(Now - dateOpen, 0) '营业天数
            dateClose = Now
        Else
            Yyts = Round(dateClose - dateOpen, 0)
            dateClose = RsShop.Fields(fShopCloseDate)
        End If
    
    
        For i = 1 To numEmployee
            
            SetShopRandomSeed RsShop.Fields(fShopDistrictID).value
            Randomize
            myRnd = Rnd()
            dateEntry = RsShop.Fields(fShopOpenDate) - Round(14 * myRnd, 0) - 7
            If dateEntry > Now Then dateEntry = Format(Now, "YYYY-MM-DD")
            
            employeeName = generateName(myRnd)
            If myRnd < 0.7 Then employeeGender = "女" Else employeeGender = "男"
            employeeOrgID = RsShop.Fields(fShopID)
            employeeJobTitle = JobTitlesArr(12)
            employeeGrade = GradeArr(6)
            employeeEdu = EduArr(Round(1 + Rnd() * 2, 0))
            employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 8000, 0)
            employeeEntryDate = dateEntry
    
            AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender, employeeOrgID
                    
'           ResignationArr = Array("个人发展", "工资原因", "工资强度", "工作内容与环境", "家庭原因", "身体原因", "违反规章制度", "劝离", "旷离", "其他原因", "试用期内解除") '试用期放在索引10
            If RsEmployee.Fields(fEmployeeID) Mod 13 = 12 Then
                workDays = Round(Yyts * myRnd, 0)
                dateResignation = dateOpen + workDays
                
                days = 0
                '保证离职日期不落在休息日和假期
                Do Until CLng(dictAllDate(dateResignation)("status")) > 2 Or days > 100
                    days = days + 1
                    dateResignation = dateResignation + 1
                Loop
                
                RsEmployee.Fields(fEmployeeResignationDate) = dateOpen + Round(Yyts * myRnd, 0)
                If workDays < 90 Then
                    RsEmployee.Fields(fEmployeeResignationReason) = ResignationArr(10)
                Else
                    RsEmployee.Fields(fEmployeeResignationReason) = ResignationArr(Round(myRnd * 10, 0))
                End If

                RsEmployee.Update
                
                '离职补一个
                SetShopRandomSeed RsShop.Fields(fShopDistrictID).value
                Randomize
                myRnd = Rnd()
                dateEntry = dateResignation + Round(30 * myRnd - 15, 0)
                If dateEntry > Now Then dateEntry = Format(Now, "YYYY-MM-DD")
                
                employeeName = generateName(myRnd)
                If myRnd < 0.7 Then employeeGender = "女" Else employeeGender = "男"
                employeeOrgID = RsShop.Fields(fShopID)
                employeeJobTitle = JobTitlesArr(12)
                employeeGrade = GradeArr(6)
                employeeEdu = EduArr(Round(1 + Rnd() * 2, 0))
                employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 8000, 0)
                employeeEntryDate = dateEntry
        
                AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender, employeeOrgID
            End If
           
        Next
        
    RsShop.MoveNext
    Loop

    CloseConnRs conn, RsEmployee, RsShop
    
End Function

Public Function generateName(ByVal myRnd As Double) As String
    '随机生成姓名
    Dim ArrFN, ArrLN
    Dim fnUB0, fnUB1, lnUB As Long
    ArrFN = Split(FirstName(), ",")
    ArrLN = Split(LastName(), ",")
    fnUB0 = Round(UBound(ArrFN) * myRnd, 0)
    fnUB1 = Round(UBound(ArrFN) * (1 - myRnd), 0)
    lnUB = Round(UBound(ArrLN) * myRnd, 0)
    If myRnd < 0.66 Then
        generateName = ArrLN(lnUB) & ArrFN(fnUB0)
    Else
        generateName = ArrLN(lnUB) & ArrFN(fnUB0) & ArrFN(fnUB1)
    End If
End Function

Public Function CreateConnection() As Object
    ' 创建连接对象
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Set conn = CurrentProject.Connection
    Set CreateConnection = conn
End Function

Public Function CreateRecordset(ByRef conn As Object, source As String) As Object
    '创建 Recordset 对象
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    With rs
        .ActiveConnection = conn
        .source = source
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    Set CreateRecordset = rs
End Function

Public Function RsClose(ByRef rs As Object)
    ' 关闭 Recordset 对象
    If Not (rs Is Nothing) Then
        rs.Close
        Set rs = Nothing
    End If
End Function

Public Function ConnClose(ByRef conn As Object)
    ' 关闭 Connection 对象
    If Not (conn Is Nothing) Then
        conn.Close
        Set conn = Nothing
    End If
End Function

Public Function CloseConnRs(ByRef conn As Object, ParamArray rsList() As Variant)
    Dim i As Integer
    
    ' 关闭所有 Recordset 对象
    For i = LBound(rsList) To UBound(rsList)
        If Not (rsList(i) Is Nothing) Then
            rsList(i).Close
            Set rsList(i) = Nothing
        End If
    Next i
    
    ' 关闭 Connection 对象
    If Not (conn Is Nothing) Then
        conn.Close
        Set conn = Nothing
    End If
End Function

Public Function SetShopRandomSeed(ByVal RsShop3 As Long)
    Dim str As String
    Dim seed As Long, i As Long

    str = CStr(RsShop3)
    seed = 0

    For i = 1 To Len(str) ' 更改了循环起始值，并使用 Len(str) 动态设置结束值
        seed = seed + CInt(Mid(str, i, 1)) ' 更正为正确的 Mid 函数用法
    Next

    Randomize seed
End Function
Public Function FirstName() As String
' 生成姓名的名

    FirstName = "埃,艾,爱,安,谙,鞍,岸,按,案,昂,凹,敖,熬,奥,百,佰,斑,阪,板,半,扮,邦,棒,傍,包,褒,保,葆,杯,碑,辈,彼,币,闭,陛,弼,碧,编,便,冰,丙,稟,帛,捕,布,步,部,埠,瓿,菜,"
    FirstName = FirstName & "参,沧,曹,测,层,查,姹,柴,婵,蒇,昌,菖,尝,常,厂,畅,朝,潮,尘,晨,谌,诚,承,晟,橙,秤,侈,耻,冲,充,憧,虫,崇,绸,出,厨,楚,褚,处,畜,穿,传,船,床,炊,垂,春,纯,醇,绰,"
    FirstName = FirstName & "辍,祠,慈,磁,糍,伺,刺,赐,聪,从,凑,粗,促,崔,催,粹,翠,沓,妲,打,呆,代,岱,玳,带,贷,待,眈,旦,淡,惮,当,宕,导,倒,得,灯,低,堤,迪,嫡,底,弟,帝,第,滇,电,佃,甸,奠,"
    FirstName = FirstName & "殿,吊,钓,叮,玎,町,钉,酊,顶,鼎,订,东,冬,动,侗,洞,恫,豆,堵,杜,肚,度,渡,端,段,锻,堆,对,敦,顿,舵,俄,峨,蛾,恩,儿,尔,发,伐,罚,帆,番,幡,烦,樊标,返,范,贩,防,"
    FirstName = FirstName & "房,放,妃,非,淝,翡,废,费,坟,汾,份,枫,锋,逢,奉,扶,佛,服,艴,福,府,腑,妇,阜,复,副,赋,傅,腹,伽,该,改,盖,杆,肝,柑,敢,感,橄,干,淦,刚,纲,钢,岗,港,皋,高,膏,篙,"
    FirstName = FirstName & "糕,告,诰,哥,歌,革,格,个,各,铬,给,根,耕,耿,梗,更,攻,宫,恭,拱,珙,共,贡,沟,构,购,够,估,辜,谷,股,鼓,嘏,故,挂,冠,莞,馆尽,管,贯,光,圭,龟,规,皈,闺,轨,癸,贵,"
    FirstName = FirstName & "桂,崞,国,亥,酣,邯,涵,汉,杭,好,皓,合,贺,黑,珩,横,弘,泓,荭窥,侯,忽,壶,湖,瑚,虎,浒,琥,沪,戽,哗,滑,划,画,淮,缓,宦,换,涣,浣,黄,惶,灰,挥,珲,辉,麾,回,卉,汇,"
    FirstName = FirstName & "慧,昏,浑,混,或,获,机,姬,基,箕,吉,级,极,即,急,几,脊,掎,计,记,伎,纪,妓,迹,祭,寄,甲,贾,假,架,嫁,坚,肩,监,兼,剪,翦,见,建,健,渐,鉴,箭,江,将,姜,奖,姣,椒,蛟,"
    FirstName = FirstName & "角,侥,皎,叫,教,皆,接,节,捷,解,界,借,金,津,仅,紧,近,劲,晋,茎,经,菁,晶,兢,精,景,径,竟,婧,敬,靖,纠,究,九,玖,酒,臼,救,局,桔,莒,矩,巨,句,拒,具,钜,俱,捐,娟,"
    FirstName = FirstName & "绢,决,诀,珏,掘,军,均,君,俊,峻,开,凯,恺,楷,勘,坎,砍,看,康,伉,柯,科,克,客,肯,垦,扣,寇,苦,库,夸,块,款,匡,筐,况,奎,揆,魁,匮,愧,昆,琨,廓,拉,莱,崃,徕,岚,廊,"
    FirstName = FirstName & "榔,朗,浪,老,佬,烙,乐,雷,累,冷,厘,梨,李,里,历,立,吏,励,利,例,俐,莉,莅,栗,俩,奁,连,良,梁,两,亮,僚,料,廖,烈,林,琳,嶙,吝,伶,苓,玲,铃,凌,翎,绫,岭,领,另,溜,"
    FirstName = FirstName & "留,流,柳,娄,炉,路,伦,囵,纶,洛,侣,旅,屡,绿,略,麻,玛,码,买,麦,卖,脉,满,漫,慢,盲,冒,贸,枚,梅,媒,楣,煤,媚,门,盟,孟,梦,米,绵,棉,冕,描,淼,渺,庙,民,皿,闵,敏,"
    FirstName = FirstName & "明,茗,模,摩,末,漠,墨,牟,母,目,沐,牧,幕,睦,慕,暮,拿,那,纳,钠,娜,奶,耐,男,南,楠,瑙,嫩,尼,倪,你,拈,念,娘,鸟,妞,纽,农,弄,奴,努,暖,欧,偶,杷,琶,牌,盘,磐,袍,"
    FirstName = FirstName & "沛,朋,棚,丕,劈,枇,篇,票,漂,聘,平,评,坪,屏,坡,迫,仆,葡,朴,普,铺,妻,漆,齐,圻,岐,奇,俟,崎,骑,琪,琦,旗,杞,启,绮,气,器,洽,仟,谦,搴,前,乾,强,悄,侨,桥,樵,巧,"
    FirstName = FirstName & "俏,茄,且,妾,侵,亲,秦,琴,勤,沁,轻,卿,请,庆,穷,丘,秋,楸,求,区,曲,岖,屈,劬,取,娶,去,趣,泉,铨,却,确,然,冉,染,热,认,仞,妊,荣,蓉,溶,榕,融,柔,肉,儒,乳,锐,瑞,"
    FirstName = FirstName & "若,塞,三,桑,瑟,沙,纱,刹,砂,厦,珊,汕,商,赏,尚,裳,韶,邵,劭,绍,佘,设,谁,申,绅,神,审,肾,甚,慎,升,生,省,狮,施,石,实,食,莳,史,使,始,示,世,仕,市,势,事,侍,饰,"
    FirstName = FirstName & "是,室,誓,手,首,寿,受,售,兽,书,抒,纾,枢,姝,殊,输,塾,熟,署,蜀,术,戍,述,树润,竖,庶,数,墅,甩,帅,爽,水,睡,说,朔,硕,司,思,厮,巳,汜,饲,驷,嗣,松,菘,讼,诵,俗,"
    FirstName = FirstName & "素,速,宿,溯,算,岁,遂,所,索,他,塔,拓,台,汰,坛,炭,探,汤,唐,堂,棠,塘,倘,掏,桃,淘,讨,套,特,梯,提,体,倜,悌,替,添,田,恬,甜,佻,挑,条,窕,帖,贴,汀,廷,亭,庭,停,"
    FirstName = FirstName & "挺,梃,通,同,佟,彤,桐,桶,凸,突,图,徒,涂,湍,团,推,退,吞,托,陀,沱,柁,娃,瓦,哇,外,完,宛,菀,晚,婉,琬,碗,畹,万,汪,往,忘,旺,威,韦,唯,帷,伟,苇,尾,纬,玮,委,炜,"
    FirstName = FirstName & "卫,未,位,味,畏,胃,谓,尉,渭,慰,温,纹,玟,闻,蚊,吻,问,翁,蓊,我,沃,卧,握,渥,斡,巫,屋,吾,吴,梧,五,伍,妩,武,舞,兀,务,物,悟,吸,昔,析,犀,溪,僖,熹,习,喜,葸,系,"
    FirstName = FirstName & "细,侠,仙,贤,弦,娴,衔,舷,限,线,宪,相,香,湘,箱,享,向,巷,像,橡,逍,消,销,霄,晓,肖,校,啸,协,偕,谐,械,心,欣,新,信,星,刑,邢,行,型,荥,醒,兴,性,姓,兄,熊,修,羞,"
    FirstName = FirstName & "袖,需,旭,叙,绪,煦,宣,萱,喧,玄,旋,漩,璇,穴,学,勋,熏,询,押,芽,衙,亚,烟,菸,嫣,言,岩,沿,研,筵,衍,眼,琰,演,彦,宴,央,秧,鞅,扬,杨,佯,洋,仰,养,样,腰,姚,摇,瑶,"
    FirstName = FirstName & "杳,要,椰,耶,冶,野,业,叶,页,夜,谒,伊,依,仪,诒,怡,宜,移,颐,以,倚,弋,亿,义,屹,抑,邑,役,易,诣,奕,谊,逸,翌,肄,裔,意,溢,毅,熠,因,阴,音,姻,吟,银,寅,饮,印,胤,"
    FirstName = FirstName & "英,瑛,迎,盈,莹,萤,赢,郢,颍,影,映,庸,雍,壅,永,咏,泳,勇,涌,忧,幽,悠,由,邮,犹,油,游,酉,莠,右,幼,侑,柚,诱,于,余,鱼,於,俞,渔,逾,渝,瑜,榆,虞,与,宇,禹,语,圉,"
    FirstName = FirstName & "玉,郁,育,昱,钰,预,欲,谕,裕,煜,豫,园,员,沅,袁,原,圆,援,缘,源,苑,院,瑗,愿,约,岳,阅,悦,晕,芸,昀,耘,筠,运,熨,匝,哉,栽,仔,咱,皂,灶,造,则,责,增,甑,栅,宅,寨,"
    FirstName = FirstName & "旃,詹,展,崭,占,栈,战,站,张,彰,璋,长,掌,仗,杖,帐,嶂,招,昭,召,赵,哲,者,这,浙,贞,针,侦,珍,帧,真,桢,砧,祯,斟,甄,箴,轸,圳,振,朕,筝,拯,整,正,诤,芝,枝,知,直,"
    FirstName = FirstName & "值,植,殖,只,纸,祉,指,趾,至,志,帜,帙,质,治,峙,致,秩,痔,智,置,终,盅,种,重,舟,洲,妯,纣,侏,珠,株,竹,竺,主,助,住,杼,贮,注,柱,炷,祝,专,撰,妆,庄,装,壮,谆,准,"
    FirstName = FirstName & "桌,卓,酌,咨,姿,滋,紫,综,奏,租,足,卒,族,组,祖,醉,昨,左,佐"

End Function

Public Function LastName() As String
' 生成姓名的姓
    LastName = "艾,爱,安,敖,巴,白,百里,柏,班,包,薄,鲍,贝,贲,毕,边,卞,别,邴,伯,卜,步,蔡,苍,曹,岑,曾,查,柴,昌,常,晁,巢,车,陈,成,程,池,充,仇,储,楚,褚,淳于,从,崔,笪,戴,单,"
    LastName = LastName & "单于,澹台,党,邓,狄,翟,第五,刁,丁,东,东方,东郭,东门,董,都,钭,窦,督,堵,杜,端木,段,段干,鄂,佴,法,樊,范,方,房,费,丰,封,酆,冯,凤,伏,扶,符,福,傅,富,盖,甘,干,"
    LastName = LastName & "高,戈,葛,耿,弓,公良,公孙,公西,公羊,公冶,宫,龚,巩,贡,勾,缑,古,谷,谷梁,顾,关,管,广,归,桂,郭,国,哈,海,韩,杭,郝,何,和,贺,赫连,衡,弘,红,洪,侯,后,後,呼延,胡,"
    LastName = LastName & "扈,花,华,滑,怀,桓,宦,皇甫,黄,惠,霍,姬,嵇,吉,汲,籍,计,纪,季,蓟,暨,冀,夹谷,家,郏,贾,简,江,姜,蒋,焦,解,金,晋,靳,经,荆,井,景,居,鞠,阚,康,亢,柯,空,孔,寇,蒯,"
    LastName = LastName & "匡,况,夔,赖,蓝,郎,劳,乐,乐正,雷,冷,黎,李,厉,利,郦,连,廉,梁,梁丘,廖,林,蔺,凌,令狐,刘,柳,龙,隆,娄,卢,鲁,陆,逯,禄,路,栾,罗,骆,闾丘,吕,麻,马,满,毛,茅,梅,蒙,"
    LastName = LastName & "孟,糜,米,宓,苗,乜,闵,明,莫,墨,牟,缪,牧,慕,慕容,穆,那,南宫,南门,能,倪,年,聂,宁,牛,钮,农,欧,欧阳,潘,庞,逄,裴,彭,蓬,皮,平,蒲,濮,濮阳,浦,戚,漆雕,亓官,齐,祁,"
    LastName = LastName & "钱,强,乔,谯,钦,秦,琴,邱,秋,裘,曲,屈,璩,瞿,权,全,阙,冉,壤驷,饶,任,戎,荣,容,融,茹,汝,阮,芮,桑,沙,山,商,赏,上官,尚,韶,邵,佘,厍,申,申屠,莘,沈,慎,盛,师,施,石,"
    LastName = LastName & "时,史,寿,殳,舒,束,帅,双,水,司,司空,司寇,司马,司徒,松,宋,苏,宿,孙,索,拓跋,邰,太叔,谈,谭,汤,唐,陶,滕,田,通,佟,童,涂,屠,万,万俟,汪,王,危,微生,韦,隗,卫,尉迟,"
    LastName = LastName & "蔚,魏,温,文,闻,闻人,翁,沃,乌,邬,巫,巫马,毋,吴,伍,武,西门,郗,奚,习,席,郤,夏,夏侯,鲜于,咸,相,向,项,萧,谢,辛,邢,幸,熊,须,胥,徐,许,轩辕,宣,薛,荀,鄢,闫,严,言,"
    LastName = LastName & "阎,颜,晏,燕,羊,羊舌,阳,杨,仰,养,姚,叶,伊,易,羿,益,阴,殷,尹,印,应,雍,尤,游,有,于,余,鱼,於,俞,虞,宇文,禹,庾,郁,喻,元,袁,岳,越,云,宰,宰父,昝,臧,詹,湛,张,章,"
    LastName = LastName & "长孙,仉,赵,甄,郑,支,终,钟,钟离,仲,仲孙,周,朱,诸,诸葛,竺,祝,颛孙,庄,卓,子车,訾,宗,宗政,邹,祖,左,左丘"

End Function

Public Function AddressProvince() As String
' 所有省区数据，包含名称，坐标等
    '大区ID 省ID    省全称  省简称1  省简称2  纬度    经度
    AddressProvince = "13,110000,北京市,京,北京,39.904987,116.405289;"
    AddressProvince = AddressProvince & "13,120000,天津市,津,天津,39.125595,117.190186;"
    AddressProvince = AddressProvince & "13,130000,河北省,冀,河北,38.045475,114.502464;"
    AddressProvince = AddressProvince & "13,140000,山西省,晋,山西,37.857014,112.549248;"
    AddressProvince = AddressProvince & "12,150000,内蒙古自治区,蒙,内蒙古,40.81831,111.670799;"
    AddressProvince = AddressProvince & "12,210000,辽宁省,辽,辽宁,41.796768,123.429092;"
    AddressProvince = AddressProvince & "12,220000,吉林省,吉,吉林,43.886841,125.324501;"
    AddressProvince = AddressProvince & "12,230000,黑龙江省,黑,黑龙江,45.756966,126.642464;"
    AddressProvince = AddressProvince & "9,310000,上海市,沪,上海,31.231707,121.472641;"
    AddressProvince = AddressProvince & "9,320000,江苏省,苏,江苏,32.041546,118.76741;"
    AddressProvince = AddressProvince & "9,330000,浙江省,浙,浙江,30.287458,120.15358;"
    AddressProvince = AddressProvince & "9,340000,安徽省,皖,安徽,31.861191,117.283043;"
    AddressProvince = AddressProvince & "9,350000,福建省,闽,福建,26.075302,119.306236;"
    AddressProvince = AddressProvince & "9,360000,江西省,赣,江西,28.676493,115.892151;"
    AddressProvince = AddressProvince & "12,370000,山东省,鲁,山东,36.675808,117.000923;"
    AddressProvince = AddressProvince & "13,410000,河南省,豫,河南,34.757977,113.665413;"
    AddressProvince = AddressProvince & "13,420000,湖北省,鄂,湖北,30.584354,114.298569;"
    AddressProvince = AddressProvince & "11,430000,湖南省,湘,湖南,28.19409,112.982277;"
    AddressProvince = AddressProvince & "11,440000,广东省,粤,广东,23.125177,113.28064;"
    AddressProvince = AddressProvince & "11,450000,广西壮族自治区,桂,广西,22.82402,108.320007;"
    AddressProvince = AddressProvince & "11,460000,海南省,琼,海南,20.031971,110.331192;"
    AddressProvince = AddressProvince & "10,500000,重庆市,渝,重庆,29.533155,106.504959;"
    AddressProvince = AddressProvince & "10,510000,四川省,川,四川,30.659462,104.065735;"
    AddressProvince = AddressProvince & "11,520000,贵州省,黔,贵州,26.578342,106.713478;"
    AddressProvince = AddressProvince & "11,530000,云南省,滇,云南,25.040609,102.71225;"
    AddressProvince = AddressProvince & "10,540000,西藏自治区,藏,西藏,29.66036,91.13221;"
    AddressProvince = AddressProvince & "13,610000,陕西省,陕,陕西,34.263161,108.948021;"
    AddressProvince = AddressProvince & "10,620000,甘肃省,甘,甘肃,36.058041,103.823555;"
    AddressProvince = AddressProvince & "10,630000,青海省,青,青海,36.623177,101.778915;"
    AddressProvince = AddressProvince & "10,640000,宁夏回族自治区,宁,宁夏,38.46637,106.278175;"
    AddressProvince = AddressProvince & "10,650000,新疆维吾尔自治区,新,新疆,43.792816,87.617729;"
    AddressProvince = AddressProvince & "14,710000,台湾省,台,台湾,25.041618,121.501618;"
    AddressProvince = AddressProvince & "14,810000,香港特别行政区,港,香港,22.320047,114.173355;"
    AddressProvince = AddressProvince & "14,820000,澳门特别行政区,澳,澳门,22.198952,113.549088"


End Function

Public Function AddressCity() As String
' 所有地市数据，包含名称，坐标等。
    '省ID    城市ID  城市    纬度    经度
    AddressCity = "110000,110000,北京,39.904989,116.405285;"
    AddressCity = AddressCity & "120000,120000,天津,39.125596,117.190182;"
    AddressCity = AddressCity & "130000,130100,石家庄,38.045474,114.502461;"
    AddressCity = AddressCity & "130000,130200,唐山,39.635113,118.175393;"
    AddressCity = AddressCity & "130000,130300,秦皇岛,39.942531,119.586579;"
    AddressCity = AddressCity & "130000,130400,邯郸,36.612273,114.490686;"
    AddressCity = AddressCity & "130000,130500,邢台,37.0682,114.508851;"
    AddressCity = AddressCity & "130000,130600,保定,38.867657,115.482331;"
    AddressCity = AddressCity & "130000,130700,张家口,40.811901,114.884091;"
    AddressCity = AddressCity & "130000,130800,承德,40.976204,117.939152;"
    AddressCity = AddressCity & "130000,130900,沧州,38.310582,116.857461;"
    AddressCity = AddressCity & "130000,131000,廊坊,39.523927,116.704441;"
    AddressCity = AddressCity & "130000,131100,衡水,37.735097,115.665993;"
    AddressCity = AddressCity & "140000,140100,太原,37.857014,112.549248;"
    AddressCity = AddressCity & "140000,140200,大同,40.09031,113.295259;"
    AddressCity = AddressCity & "140000,140300,阳泉,37.861188,113.583285;"
    AddressCity = AddressCity & "140000,140400,长治,36.191112,113.113556;"
    AddressCity = AddressCity & "140000,140500,晋城,35.497553,112.851274;"
    AddressCity = AddressCity & "140000,140600,朔州,39.331261,112.433387;"
    AddressCity = AddressCity & "140000,140700,晋中,37.696495,112.736465;"
    AddressCity = AddressCity & "140000,140800,运城,35.022778,111.003957;"
    AddressCity = AddressCity & "140000,140900,忻州,38.41769,112.733538;"
    AddressCity = AddressCity & "140000,141000,临汾,36.08415,111.517973;"
    AddressCity = AddressCity & "140000,141100,吕梁,37.524366,111.134335;"
    AddressCity = AddressCity & "150000,150100,呼和浩特,40.818311,111.670801;"
    AddressCity = AddressCity & "150000,150200,包头,40.658168,109.840405;"
    AddressCity = AddressCity & "150000,150300,乌海,39.673734,106.825563;"
    AddressCity = AddressCity & "150000,150400,赤峰,42.275317,118.956806;"
    AddressCity = AddressCity & "150000,150500,通辽,43.617429,122.263119;"
    AddressCity = AddressCity & "150000,150600,鄂尔多斯,39.817179,109.99029;"
    AddressCity = AddressCity & "150000,150700,呼伦贝尔,49.215333,119.758168;"
    AddressCity = AddressCity & "150000,150800,巴彦淖尔,40.757402,107.416959;"
    AddressCity = AddressCity & "150000,150900,乌兰察布,41.034126,113.114543;"
    AddressCity = AddressCity & "150000,152200,兴安,46.076268,122.070317;"
    AddressCity = AddressCity & "150000,152500,锡林郭勒,43.944018,116.090996;"
    AddressCity = AddressCity & "150000,152900,阿拉善,38.844814,105.706422;"
    AddressCity = AddressCity & "210000,210100,沈阳,41.796767,123.429096;"
    AddressCity = AddressCity & "210000,210200,大连,38.91459,121.618622;"
    AddressCity = AddressCity & "210000,210300,鞍山,41.110626,122.995632;"
    AddressCity = AddressCity & "210000,210400,抚顺,41.875956,123.921109;"
    AddressCity = AddressCity & "210000,210500,本溪,41.297909,123.770519;"
    AddressCity = AddressCity & "210000,210600,丹东,40.124296,124.383044;"
    AddressCity = AddressCity & "210000,210700,锦州,41.119269,121.135742;"
    AddressCity = AddressCity & "210000,210800,营口,40.667432,122.235151;"
    AddressCity = AddressCity & "210000,210900,阜新,42.011796,121.648962;"
    AddressCity = AddressCity & "210000,211000,辽阳,41.269402,123.18152;"
    AddressCity = AddressCity & "210000,211100,盘锦,41.124484,122.06957;"
    AddressCity = AddressCity & "210000,211200,铁岭,42.290585,123.844279;"
    AddressCity = AddressCity & "210000,211300,朝阳,41.576758,120.451176;"
    AddressCity = AddressCity & "210000,211400,葫芦岛,40.755572,120.856394;"
    AddressCity = AddressCity & "220000,220100,长春,43.886841,125.3245;"
    AddressCity = AddressCity & "220000,220200,吉林,43.843577,126.55302;"
    AddressCity = AddressCity & "220000,220300,四平,43.170344,124.370785;"
    AddressCity = AddressCity & "220000,220400,辽源,42.902692,125.145349;"
    AddressCity = AddressCity & "220000,220500,通化,41.721177,125.936501;"
    AddressCity = AddressCity & "220000,220600,白山,41.942505,126.427839;"
    AddressCity = AddressCity & "220000,220700,松原,45.118243,124.823608;"
    AddressCity = AddressCity & "220000,220800,白城,45.619026,122.841114;"
    AddressCity = AddressCity & "220000,222400,延边朝鲜族,42.904823,129.513228;"
    AddressCity = AddressCity & "230000,230100,哈尔滨,45.756967,126.642464;"
    AddressCity = AddressCity & "230000,230200,齐齐哈尔,47.342081,123.95792;"
    AddressCity = AddressCity & "230000,230300,鸡西,45.300046,130.975966;"
    AddressCity = AddressCity & "230000,230400,鹤岗,47.332085,130.277487;"
    AddressCity = AddressCity & "230000,230500,双鸭山,46.643442,131.157304;"
    AddressCity = AddressCity & "230000,230600,大庆,46.590734,125.11272;"
    AddressCity = AddressCity & "230000,230700,伊春,47.724775,128.899396;"
    AddressCity = AddressCity & "230000,230800,佳木斯,46.809606,130.361634;"
    AddressCity = AddressCity & "230000,230900,七台河,45.771266,131.015584;"
    AddressCity = AddressCity & "230000,231000,牡丹江,44.582962,129.618602;"
    AddressCity = AddressCity & "230000,231100,黑河,50.249585,127.499023;"
    AddressCity = AddressCity & "230000,231200,绥化,46.637393,126.99293;"
    AddressCity = AddressCity & "230000,232700,大兴安岭,52.335262,124.711526;"
    AddressCity = AddressCity & "310000,310000,上海,31.231706,121.472644;"
    AddressCity = AddressCity & "320000,320100,南京,32.041544,118.767413;"
    AddressCity = AddressCity & "320000,320200,无锡,31.574729,120.301663;"
    AddressCity = AddressCity & "320000,320300,徐州,34.261792,117.184811;"
    AddressCity = AddressCity & "320000,320400,常州,31.772752,119.946973;"
    AddressCity = AddressCity & "320000,320500,苏州,31.299379,120.619585;"
    AddressCity = AddressCity & "320000,320600,南通,32.016212,120.864608;"
    AddressCity = AddressCity & "320000,320700,连云港,34.600018,119.178821;"
    AddressCity = AddressCity & "320000,320800,淮安,33.597506,119.021265;"
    AddressCity = AddressCity & "320000,320900,盐城,33.377631,120.139998;"
    AddressCity = AddressCity & "320000,321000,扬州,32.393159,119.421003;"
    AddressCity = AddressCity & "320000,321100,镇江,32.204402,119.452753;"
    AddressCity = AddressCity & "320000,321200,泰州,32.484882,119.915176;"
    AddressCity = AddressCity & "320000,321300,宿迁,33.963008,118.275162;"
    AddressCity = AddressCity & "330000,330100,杭州,30.287459,120.153576;"
    AddressCity = AddressCity & "330000,330200,宁波,29.868388,121.549792;"
    AddressCity = AddressCity & "330000,330300,温州,28.000575,120.672111;"
    AddressCity = AddressCity & "330000,330400,嘉兴,30.762653,120.750865;"
    AddressCity = AddressCity & "330000,330500,湖州,30.867198,120.102398;"
    AddressCity = AddressCity & "330000,330600,绍兴,29.997117,120.582112;"
    AddressCity = AddressCity & "330000,330700,金华,29.089524,119.649506;"
    AddressCity = AddressCity & "330000,330800,衢州,28.941708,118.87263;"
    AddressCity = AddressCity & "330000,330900,舟山,30.016028,122.106863;"
    AddressCity = AddressCity & "330000,331000,台州,28.661378,121.428599;"
    AddressCity = AddressCity & "330000,331100,丽水,28.451993,119.921786;"
    AddressCity = AddressCity & "340000,340100,合肥,31.86119,117.283042;"
    AddressCity = AddressCity & "340000,340200,芜湖,31.326319,118.376451;"
    AddressCity = AddressCity & "340000,340300,蚌埠,32.939667,117.363228;"
    AddressCity = AddressCity & "340000,340400,淮南,32.647574,117.018329;"
    AddressCity = AddressCity & "340000,340500,马鞍山,31.689362,118.507906;"
    AddressCity = AddressCity & "340000,340600,淮北,33.971707,116.794664;"
    AddressCity = AddressCity & "340000,340700,铜陵,30.929935,117.816576;"
    AddressCity = AddressCity & "340000,340800,安庆,30.50883,117.043551;"
    AddressCity = AddressCity & "340000,341000,黄山,29.709239,118.317325;"
    AddressCity = AddressCity & "340000,341100,滁州,32.303627,118.316264;"
    AddressCity = AddressCity & "340000,341200,阜阳,32.896969,115.819729;"
    AddressCity = AddressCity & "340000,341300,宿州,33.633891,116.984084;"
    AddressCity = AddressCity & "340000,341500,六安,31.752889,116.507676;"
    AddressCity = AddressCity & "340000,341600,亳州,33.869338,115.782939;"
    AddressCity = AddressCity & "340000,341700,池州,30.656037,117.489157;"
    AddressCity = AddressCity & "340000,341800,宣城,30.945667,118.757995;"
    AddressCity = AddressCity & "350000,350100,福州,26.075302,119.306239;"
    AddressCity = AddressCity & "350000,350200,厦门,24.490474,118.11022;"
    AddressCity = AddressCity & "350000,350300,莆田,25.431011,119.007558;"
    AddressCity = AddressCity & "350000,350400,三明,26.265444,117.635001;"
    AddressCity = AddressCity & "350000,350500,泉州,24.908853,118.589421;"
    AddressCity = AddressCity & "350000,350600,漳州,24.510897,117.661801;"
    AddressCity = AddressCity & "350000,350700,南平,26.635627,118.178459;"
    AddressCity = AddressCity & "350000,350800,龙岩,25.091603,117.02978;"
    AddressCity = AddressCity & "350000,350900,宁德,26.65924,119.527082;"
    AddressCity = AddressCity & "360000,360100,南昌,28.676493,115.892151;"
    AddressCity = AddressCity & "360000,360200,景德镇,29.29256,117.214664;"
    AddressCity = AddressCity & "360000,360300,萍乡,27.622946,113.852186;"
    AddressCity = AddressCity & "360000,360400,九江,29.712034,115.992811;"
    AddressCity = AddressCity & "360000,360500,新余,27.810834,114.930835;"
    AddressCity = AddressCity & "360000,360600,鹰潭,28.238638,117.033838;"
    AddressCity = AddressCity & "360000,360700,赣州,25.85097,114.940278;"
    AddressCity = AddressCity & "360000,360800,吉安,27.111699,114.986373;"
    AddressCity = AddressCity & "360000,360900,宜春,27.8043,114.391136;"
    AddressCity = AddressCity & "360000,361000,抚州,27.98385,116.358351;"
    AddressCity = AddressCity & "360000,361100,上饶,28.44442,117.971185;"
    AddressCity = AddressCity & "370000,370100,济南,36.675807,117.000923;"
    AddressCity = AddressCity & "370000,370200,青岛,36.082982,120.355173;"
    AddressCity = AddressCity & "370000,370300,淄博,36.814939,118.047648;"
    AddressCity = AddressCity & "370000,370400,枣庄,34.856424,117.557964;"
    AddressCity = AddressCity & "370000,370500,东营,37.434564,118.66471;"
    AddressCity = AddressCity & "370000,370600,烟台,37.539297,121.391382;"
    AddressCity = AddressCity & "370000,370700,潍坊,36.70925,119.107078;"
    AddressCity = AddressCity & "370000,370800,济宁,35.415393,116.587245;"
    AddressCity = AddressCity & "370000,370900,泰安,36.194968,117.129063;"
    AddressCity = AddressCity & "370000,371000,威海,37.509691,122.116394;"
    AddressCity = AddressCity & "370000,371100,日照,35.428588,119.461208;"
    AddressCity = AddressCity & "370000,371300,临沂,35.065282,118.326443;"
    AddressCity = AddressCity & "370000,371400,德州,37.453968,116.307428;"
    AddressCity = AddressCity & "370000,371500,聊城,36.456013,115.980367;"
    AddressCity = AddressCity & "370000,371600,滨州,37.383542,118.016974;"
    AddressCity = AddressCity & "370000,371700,菏泽,35.246531,115.469381;"
    AddressCity = AddressCity & "410000,410100,郑州,34.757975,113.665412;"
    AddressCity = AddressCity & "410000,410200,开封,34.797049,114.341447;"
    AddressCity = AddressCity & "410000,410300,洛阳,34.663041,112.434468;"
    AddressCity = AddressCity & "410000,410400,平顶山,33.735241,113.307718;"
    AddressCity = AddressCity & "410000,410500,安阳,36.103442,114.352482;"
    AddressCity = AddressCity & "410000,410600,鹤壁,35.748236,114.295444;"
    AddressCity = AddressCity & "410000,410700,新乡,35.302616,113.883991;"
    AddressCity = AddressCity & "410000,410800,焦作,35.23904,113.238266;"
    AddressCity = AddressCity & "410000,419001,济源,35.090378,112.590047;"
    AddressCity = AddressCity & "410000,410900,濮阳,35.768234,115.041299;"
    AddressCity = AddressCity & "410000,411000,许昌,34.022956,113.826063;"
    AddressCity = AddressCity & "410000,411100,漯河,33.575855,114.026405;"
    AddressCity = AddressCity & "410000,411200,三门峡,34.777338,111.194099;"
    AddressCity = AddressCity & "410000,411300,南阳,32.999082,112.540918;"
    AddressCity = AddressCity & "410000,411400,商丘,34.437054,115.650497;"
    AddressCity = AddressCity & "410000,411500,信阳,32.123274,114.075031;"
    AddressCity = AddressCity & "410000,411600,周口,33.620357,114.649653;"
    AddressCity = AddressCity & "410000,411700,驻马店,32.980169,114.024736;"
    AddressCity = AddressCity & "420000,420100,武汉,30.584355,114.298572;"
    AddressCity = AddressCity & "420000,420200,黄石,30.220074,115.077048;"
    AddressCity = AddressCity & "420000,420300,十堰,32.646907,110.787916;"
    AddressCity = AddressCity & "420000,420500,宜昌,30.702636,111.290843;"
    AddressCity = AddressCity & "420000,420600,襄阳,32.042426,112.144146;"
    AddressCity = AddressCity & "420000,420700,鄂州,30.396536,114.890593;"
    AddressCity = AddressCity & "420000,420800,荆门,31.03542,112.204251;"
    AddressCity = AddressCity & "420000,420900,孝感,30.926423,113.926655;"
    AddressCity = AddressCity & "420000,421000,荆州,30.326857,112.23813;"
    AddressCity = AddressCity & "420000,421100,黄冈,30.447711,114.879365;"
    AddressCity = AddressCity & "420000,421200,咸宁,29.832798,114.328963;"
    AddressCity = AddressCity & "420000,421300,随州,31.717497,113.37377;"
    AddressCity = AddressCity & "420000,422800,恩施,30.283114,109.48699;"
    AddressCity = AddressCity & "420000,429004,仙桃,30.364953,113.453974;"
    AddressCity = AddressCity & "420000,429005,潜江,30.421215,112.896866;"
    AddressCity = AddressCity & "420000,429006,天门,30.653061,113.165862;"
    AddressCity = AddressCity & "420000,429021,神农架,31.744449,110.671525;"
    AddressCity = AddressCity & "430000,430100,长沙,28.19409,112.982279;"
    AddressCity = AddressCity & "430000,430200,株洲,27.835806,113.151737;"
    AddressCity = AddressCity & "430000,430300,湘潭,27.82973,112.944052;"
    AddressCity = AddressCity & "430000,430400,衡阳,26.900358,112.607693;"
    AddressCity = AddressCity & "430000,430500,邵阳,27.237842,111.46923;"
    AddressCity = AddressCity & "430000,430600,岳阳,29.37029,113.132855;"
    AddressCity = AddressCity & "430000,430700,常德,29.040225,111.691347;"
    AddressCity = AddressCity & "430000,430800,张家界,29.127401,110.479921;"
    AddressCity = AddressCity & "430000,430900,益阳,28.570066,112.355042;"
    AddressCity = AddressCity & "430000,431000,郴州,25.793589,113.032067;"
    AddressCity = AddressCity & "430000,431100,永州,26.434516,111.608019;"
    AddressCity = AddressCity & "430000,431200,怀化,27.550082,109.97824;"
    AddressCity = AddressCity & "430000,431300,娄底,27.728136,112.008497;"
    AddressCity = AddressCity & "430000,433100,湘西,28.314296,109.739735;"
    AddressCity = AddressCity & "440000,440100,广州,23.125178,113.280637;"
    AddressCity = AddressCity & "440000,440200,韶关,24.801322,113.591544;"
    AddressCity = AddressCity & "440000,440300,深圳,22.547,114.085947;"
    AddressCity = AddressCity & "440000,440400,珠海,22.224979,113.553986;"
    AddressCity = AddressCity & "440000,440500,汕头,23.37102,116.708463;"
    AddressCity = AddressCity & "440000,440600,佛山,23.028762,113.122717;"
    AddressCity = AddressCity & "440000,440700,江门,22.590431,113.094942;"
    AddressCity = AddressCity & "440000,440800,湛江,21.274898,110.364977;"
    AddressCity = AddressCity & "440000,440900,茂名,21.659751,110.919229;"
    AddressCity = AddressCity & "440000,441200,肇庆,23.051546,112.472529;"
    AddressCity = AddressCity & "440000,441300,惠州,23.079404,114.412599;"
    AddressCity = AddressCity & "440000,441400,梅州,24.299112,116.117582;"
    AddressCity = AddressCity & "440000,441500,汕尾,22.774485,115.364238;"
    AddressCity = AddressCity & "440000,441600,河源,23.746266,114.697802;"
    AddressCity = AddressCity & "440000,441700,阳江,21.859222,111.975107;"
    AddressCity = AddressCity & "440000,441800,清远,23.685022,113.051227;"
    AddressCity = AddressCity & "440000,441900,东莞,23.046237,113.746262;"
    AddressCity = AddressCity & "440000,442000,中山,22.521113,113.382391;"
    AddressCity = AddressCity & "440000,445100,潮州,23.661701,116.632301;"
    AddressCity = AddressCity & "440000,445200,揭阳,23.543778,116.355733;"
    AddressCity = AddressCity & "440000,445300,云浮,22.929801,112.044439;"
    AddressCity = AddressCity & "450000,450100,南宁,22.82402,108.320004;"
    AddressCity = AddressCity & "450000,450200,柳州,24.314617,109.411703;"
    AddressCity = AddressCity & "450000,450300,桂林,25.274215,110.299121;"
    AddressCity = AddressCity & "450000,450400,梧州,23.474803,111.297604;"
    AddressCity = AddressCity & "450000,450500,北海,21.473343,109.119254;"
    AddressCity = AddressCity & "450000,450600,防城港,21.614631,108.345478;"
    AddressCity = AddressCity & "450000,450700,钦州,21.967127,108.624175;"
    AddressCity = AddressCity & "450000,450800,贵港,23.0936,109.602146;"
    AddressCity = AddressCity & "450000,450900,玉林,22.63136,110.154393;"
    AddressCity = AddressCity & "450000,451000,百色,23.897742,106.616285;"
    AddressCity = AddressCity & "450000,451100,贺州,24.414141,111.552056;"
    AddressCity = AddressCity & "450000,451200,河池,24.695899,108.062105;"
    AddressCity = AddressCity & "450000,451300,来宾,23.733766,109.229772;"
    AddressCity = AddressCity & "450000,451400,崇左,22.404108,107.353926;"
    AddressCity = AddressCity & "460000,460100,海口,20.031971,110.33119;"
    AddressCity = AddressCity & "460000,460200,三亚,18.247872,109.508268;"
    AddressCity = AddressCity & "460000,460300,三沙,16.831039,112.34882;"
    AddressCity = AddressCity & "460000,469001,五指山,18.776921,109.516662;"
    AddressCity = AddressCity & "460000,469002,琼海,19.246011,110.466785;"
    AddressCity = AddressCity & "460000,460400,儋州,19.517486,109.576782;"
    AddressCity = AddressCity & "460000,469005,文昌,19.612986,110.753975;"
    AddressCity = AddressCity & "460000,469006,万宁,18.796216,110.388793;"
    AddressCity = AddressCity & "460000,469007,东方,19.10198,108.653789;"
    AddressCity = AddressCity & "460000,469021,定安,19.684966,110.349235;"
    AddressCity = AddressCity & "460000,469022,屯昌,19.362916,110.102773;"
    AddressCity = AddressCity & "460000,469023,澄迈,19.737095,110.007147;"
    AddressCity = AddressCity & "460000,469024,临高,19.908293,109.687697;"
    AddressCity = AddressCity & "460000,469025,白沙,19.224584,109.452606;"
    AddressCity = AddressCity & "460000,469026,昌江,19.260968,109.053351;"
    AddressCity = AddressCity & "460000,469027,乐东,18.74758,109.175444;"
    AddressCity = AddressCity & "460000,469028,陵水,18.505006,110.037218;"
    AddressCity = AddressCity & "460000,469029,保亭,18.636371,109.70245;"
    AddressCity = AddressCity & "460000,469030,琼中,19.03557,109.839996;"
    AddressCity = AddressCity & "500000,500000,重庆,29.533155,106.504962;"
    AddressCity = AddressCity & "510000,510100,成都,30.659462,104.065735;"
    AddressCity = AddressCity & "510000,510300,自贡,29.352765,104.773447;"
    AddressCity = AddressCity & "510000,510400,攀枝花,26.580446,101.716007;"
    AddressCity = AddressCity & "510000,510500,泸州,28.889138,105.443348;"
    AddressCity = AddressCity & "510000,510600,德阳,31.127991,104.398651;"
    AddressCity = AddressCity & "510000,510700,绵阳,31.46402,104.741722;"
    AddressCity = AddressCity & "510000,510800,广元,32.433668,105.829757;"
    AddressCity = AddressCity & "510000,510900,遂宁,30.513311,105.571331;"
    AddressCity = AddressCity & "510000,511000,内江,29.58708,105.066138;"
    AddressCity = AddressCity & "510000,511100,乐山,29.582024,103.761263;"
    AddressCity = AddressCity & "510000,511300,南充,30.795281,106.082974;"
    AddressCity = AddressCity & "510000,511400,眉山,30.048318,103.831788;"
    AddressCity = AddressCity & "510000,511500,宜宾,28.760189,104.630825;"
    AddressCity = AddressCity & "510000,511600,广安,30.456398,106.633369;"
    AddressCity = AddressCity & "510000,511700,达州,31.209484,107.502262;"
    AddressCity = AddressCity & "510000,511800,雅安,29.987722,103.001033;"
    AddressCity = AddressCity & "510000,511900,巴中,31.858809,106.753669;"
    AddressCity = AddressCity & "510000,512000,资阳,30.122211,104.641917;"
    AddressCity = AddressCity & "510000,513200,阿坝,31.899792,102.221374;"
    AddressCity = AddressCity & "510000,513300,甘孜,30.050663,101.963815;"
    AddressCity = AddressCity & "510000,513400,凉山,27.886762,102.258746;"
    AddressCity = AddressCity & "520000,520100,贵阳,26.578343,106.713478;"
    AddressCity = AddressCity & "520000,520200,六盘水,26.584643,104.846743;"
    AddressCity = AddressCity & "520000,520300,遵义,27.706626,106.937265;"
    AddressCity = AddressCity & "520000,520400,安顺,26.245544,105.932188;"
    AddressCity = AddressCity & "520000,520600,铜仁,27.718346,109.191555;"
    AddressCity = AddressCity & "520000,522300,黔西南,25.08812,104.897971;"
    AddressCity = AddressCity & "520000,520500,毕节,27.301693,105.28501;"
    AddressCity = AddressCity & "520000,522600,黔东南,26.583352,107.977488;"
    AddressCity = AddressCity & "520000,522700,黔南,26.258219,107.517156;"
    AddressCity = AddressCity & "530000,530100,昆明,25.040609,102.712251;"
    AddressCity = AddressCity & "530000,530300,曲靖,25.501557,103.797851;"
    AddressCity = AddressCity & "530000,530400,玉溪,24.350461,102.543907;"
    AddressCity = AddressCity & "530000,530500,保山,25.111802,99.167133;"
    AddressCity = AddressCity & "530000,530600,昭通,27.336999,103.717216;"
    AddressCity = AddressCity & "530000,530700,丽江,26.872108,100.233026;"
    AddressCity = AddressCity & "530000,530800,普洱,22.777321,100.972344;"
    AddressCity = AddressCity & "530000,530900,临沧,23.886567,100.08697;"
    AddressCity = AddressCity & "530000,532300,楚雄,25.041988,101.546046;"
    AddressCity = AddressCity & "530000,532500,红河,23.366775,103.384182;"
    AddressCity = AddressCity & "530000,532600,文山,23.36951,104.24401;"
    AddressCity = AddressCity & "530000,532800,西双版纳,22.001724,100.797941;"
    AddressCity = AddressCity & "530000,532900,大理,25.589449,100.225668;"
    AddressCity = AddressCity & "530000,533100,德宏,24.436694,98.578363;"
    AddressCity = AddressCity & "530000,533300,怒江,25.850949,98.854304;"
    AddressCity = AddressCity & "530000,533400,迪庆,27.826853,99.706463;"
    AddressCity = AddressCity & "540000,540100,拉萨,29.660361,91.132212;"
    AddressCity = AddressCity & "540000,540300,昌都,31.136875,97.178452;"
    AddressCity = AddressCity & "540000,540500,山南,29.236023,91.766529;"
    AddressCity = AddressCity & "540000,540200,日喀则,29.267519,88.885148;"
    AddressCity = AddressCity & "540000,540600,那曲,31.476004,92.060214;"
    AddressCity = AddressCity & "540000,542500,阿里,32.503187,80.105498;"
    AddressCity = AddressCity & "540000,540400,林芝,29.654693,94.362348;"
    AddressCity = AddressCity & "610000,610100,西安,34.263161,108.948024;"
    AddressCity = AddressCity & "610000,610200,铜川,34.916582,108.979608;"
    AddressCity = AddressCity & "610000,610300,宝鸡,34.369315,107.14487;"
    AddressCity = AddressCity & "610000,610400,咸阳,34.333439,108.705117;"
    AddressCity = AddressCity & "610000,610500,渭南,34.499381,109.502882;"
    AddressCity = AddressCity & "610000,610600,延安,36.596537,109.49081;"
    AddressCity = AddressCity & "610000,610700,汉中,33.077668,107.028621;"
    AddressCity = AddressCity & "610000,610800,榆林,38.290162,109.741193;"
    AddressCity = AddressCity & "610000,610900,安康,32.6903,109.029273;"
    AddressCity = AddressCity & "610000,611000,商洛,33.868319,109.939776;"
    AddressCity = AddressCity & "620000,620100,兰州,36.058039,103.823557;"
    AddressCity = AddressCity & "620000,620200,嘉峪关,39.786529,98.277304;"
    AddressCity = AddressCity & "620000,620300,金昌,38.514238,102.187888;"
    AddressCity = AddressCity & "620000,620400,白银,36.54568,104.173606;"
    AddressCity = AddressCity & "620000,620500,天水,34.578529,105.724998;"
    AddressCity = AddressCity & "620000,620600,武威,37.929996,102.634697;"
    AddressCity = AddressCity & "620000,620700,张掖,38.932897,100.455472;"
    AddressCity = AddressCity & "620000,620800,平凉,35.54279,106.684691;"
    AddressCity = AddressCity & "620000,620900,酒泉,39.744023,98.510795;"
    AddressCity = AddressCity & "620000,621000,庆阳,35.734218,107.638372;"
    AddressCity = AddressCity & "620000,621100,定西,35.579578,104.626294;"
    AddressCity = AddressCity & "620000,621200,陇南,33.388598,104.929379;"
    AddressCity = AddressCity & "620000,622900,临夏,35.599446,103.212006;"
    AddressCity = AddressCity & "620000,623000,甘南,34.986354,102.911008;"
    AddressCity = AddressCity & "630000,630100,西宁,36.623178,101.778916;"
    AddressCity = AddressCity & "630000,630200,海东,36.502916,102.10327;"
    AddressCity = AddressCity & "630000,632200,海北,36.959435,100.901059;"
    AddressCity = AddressCity & "630000,632300,黄南,35.517744,102.019988;"
    AddressCity = AddressCity & "630000,632500,海南藏族,36.280353,100.619542;"
    AddressCity = AddressCity & "630000,632600,果洛,34.4736,100.242143;"
    AddressCity = AddressCity & "630000,632700,玉树,33.004049,97.008522;"
    AddressCity = AddressCity & "630000,632800,海西,37.374663,97.370785;"
    AddressCity = AddressCity & "640000,640100,银川,38.46637,106.278179;"
    AddressCity = AddressCity & "640000,640200,石嘴山,39.01333,106.376173;"
    AddressCity = AddressCity & "640000,640300,吴忠,37.986165,106.199409;"
    AddressCity = AddressCity & "640000,640400,固原,36.004561,106.285241;"
    AddressCity = AddressCity & "640000,640500,中卫,37.514951,105.189568;"
    AddressCity = AddressCity & "650000,650100,乌鲁木齐,43.792818,87.617733;"
    AddressCity = AddressCity & "650000,650200,克拉玛依,45.595886,84.873946;"
    AddressCity = AddressCity & "650000,650400,吐鲁番,42.947613,89.184078;"
    AddressCity = AddressCity & "650000,650500,哈密,42.833248,93.51316;"
    AddressCity = AddressCity & "650000,652300,昌吉,44.014577,87.304012;"
    AddressCity = AddressCity & "650000,652700,博尔塔拉,44.903258,82.074778;"
    AddressCity = AddressCity & "650000,652800,巴音郭楞,41.768552,86.150969;"
    AddressCity = AddressCity & "650000,652900,阿克苏,41.170712,80.265068;"
    AddressCity = AddressCity & "650000,653000,克孜勒苏柯尔克孜,39.713431,76.172825;"
    AddressCity = AddressCity & "650000,653100,喀什,39.467664,75.989138;"
    AddressCity = AddressCity & "650000,653200,和田,37.110687,79.92533;"
    AddressCity = AddressCity & "650000,654000,伊犁,43.92186,81.317946;"
    AddressCity = AddressCity & "650000,654200,塔城,46.746301,82.985732;"
    AddressCity = AddressCity & "650000,654300,阿勒泰,47.848393,88.13963;"
    AddressCity = AddressCity & "650000,659001,石河子,44.305886,86.041075;"
    AddressCity = AddressCity & "650000,659002,阿拉尔,40.541914,81.285884;"
    AddressCity = AddressCity & "650000,659003,图木舒克,39.867316,79.077978;"
    AddressCity = AddressCity & "650000,659004,五家渠,44.167401,87.526884;"
    '修正县级市和湾湾的层级数据
    AddressCity = AddressCity & "650000,659005,北屯,47.353177,87.824932;"
    AddressCity = AddressCity & "650000,659006,铁门关,41.827251,85.501218;"
    AddressCity = AddressCity & "650000,659007,双河,44.840524,82.353656;"
    AddressCity = AddressCity & "650000,659008,可克达拉,43.6832,80.63579;"
    AddressCity = AddressCity & "650000,659009,昆玉,37.207994,79.287372;"
    AddressCity = AddressCity & "650000,659010,胡杨河,44.69288853,84.8275959;"
    AddressCity = AddressCity & "710000,710000,台湾,25.044332,121.509062;"
    AddressCity = AddressCity & "810000,810000,香港,22.320048,114.173355;"
    AddressCity = AddressCity & "820000,820000,澳门,22.198951,113.54909"

End Function

Public Function AddressDistrict() As String
' 所有区县数据，包含名称，坐标等。
    '城市ID    区县ID  区县    纬度    经度
    AddressDistrict = "110000,110101,东城区,39.917544,116.418757;"
    AddressDistrict = AddressDistrict & "110000,110102,西城区,39.915309,116.366794;"
    AddressDistrict = AddressDistrict & "110000,110105,朝阳区,39.921489,116.486409;"
    AddressDistrict = AddressDistrict & "110000,110106,丰台区,39.863642,116.286968;"
    AddressDistrict = AddressDistrict & "110000,110107,石景山区,39.914601,116.195445;"
    AddressDistrict = AddressDistrict & "110000,110108,海淀区,39.956074,116.310316;"
    AddressDistrict = AddressDistrict & "110000,110109,门头沟区,39.937183,116.105381;"
    AddressDistrict = AddressDistrict & "110000,110111,房山区,39.735535,116.139157;"
    AddressDistrict = AddressDistrict & "110000,110112,通州区,39.902486,116.658603;"
    AddressDistrict = AddressDistrict & "110000,110113,顺义区,40.128936,116.653525;"
    AddressDistrict = AddressDistrict & "110000,110114,昌平区,40.218085,116.235906;"
    AddressDistrict = AddressDistrict & "110000,110115,大兴区,39.728908,116.338033;"
    AddressDistrict = AddressDistrict & "110000,110116,怀柔区,40.324272,116.637122;"
    AddressDistrict = AddressDistrict & "110000,110117,平谷区,40.144783,117.112335;"
    AddressDistrict = AddressDistrict & "110000,110118,密云区,40.377362,116.843352;"
    AddressDistrict = AddressDistrict & "110000,110119,延庆区,40.465325,115.985006;"
    AddressDistrict = AddressDistrict & "120000,120101,和平区,39.118327,117.195907;"
    AddressDistrict = AddressDistrict & "120000,120102,河东区,39.122125,117.226568;"
    AddressDistrict = AddressDistrict & "120000,120103,河西区,39.101897,117.217536;"
    AddressDistrict = AddressDistrict & "120000,120104,南开区,39.120474,117.164143;"
    AddressDistrict = AddressDistrict & "120000,120105,河北区,39.156632,117.201569;"
    AddressDistrict = AddressDistrict & "120000,120106,红桥区,39.175066,117.163301;"
    AddressDistrict = AddressDistrict & "120000,120110,东丽区,39.087764,117.313967;"
    AddressDistrict = AddressDistrict & "120000,120111,西青区,39.139446,117.012247;"
    AddressDistrict = AddressDistrict & "120000,120112,津南区,38.989577,117.382549;"
    AddressDistrict = AddressDistrict & "120000,120113,北辰区,39.225555,117.13482;"
    AddressDistrict = AddressDistrict & "120000,120114,武清区,39.376925,117.057959;"
    AddressDistrict = AddressDistrict & "120000,120115,宝坻区,39.716965,117.308094;"
    AddressDistrict = AddressDistrict & "120000,120116,滨海新区,39.032846,117.654173;"
    AddressDistrict = AddressDistrict & "120000,120117,宁河区,39.328886,117.82828;"
    AddressDistrict = AddressDistrict & "120000,120118,静海区,38.935671,116.925304;"
    AddressDistrict = AddressDistrict & "120000,120119,蓟州区,40.045342,117.407449;"
    AddressDistrict = AddressDistrict & "310000,310101,黄浦区,31.222771,121.490317;"
    AddressDistrict = AddressDistrict & "310000,310104,徐汇区,31.179973,121.43752;"
    AddressDistrict = AddressDistrict & "310000,310105,长宁区,31.218123,121.4222;"
    AddressDistrict = AddressDistrict & "310000,310106,静安区,31.229003,121.448224;"
    AddressDistrict = AddressDistrict & "310000,310107,普陀区,31.241701,121.392499;"
    AddressDistrict = AddressDistrict & "310000,310109,虹口区,31.26097,121.491832;"
    AddressDistrict = AddressDistrict & "310000,310110,杨浦区,31.270755,121.522797;"
    AddressDistrict = AddressDistrict & "310000,310112,闵行区,31.111658,121.375972;"
    AddressDistrict = AddressDistrict & "310000,310113,宝山区,31.398896,121.489934;"
    AddressDistrict = AddressDistrict & "310000,310114,嘉定区,31.383524,121.250333;"
    AddressDistrict = AddressDistrict & "310000,310115,浦东新区,31.245944,121.567706;"
    AddressDistrict = AddressDistrict & "310000,310116,金山区,30.724697,121.330736;"
    AddressDistrict = AddressDistrict & "310000,310117,松江区,31.03047,121.223543;"
    AddressDistrict = AddressDistrict & "310000,310118,青浦区,31.151209,121.113021;"
    AddressDistrict = AddressDistrict & "310000,310120,奉贤区,30.912345,121.458472;"
    AddressDistrict = AddressDistrict & "310000,310151,崇明区,31.626946,121.397516;"
    AddressDistrict = AddressDistrict & "500000,500101,万州区,30.807807,108.380246;"
    AddressDistrict = AddressDistrict & "500000,500102,涪陵区,29.703652,107.394905;"
    AddressDistrict = AddressDistrict & "500000,500103,渝中区,29.556742,106.56288;"
    AddressDistrict = AddressDistrict & "500000,500104,大渡口区,29.481002,106.48613;"
    AddressDistrict = AddressDistrict & "500000,500105,江北区,29.575352,106.532844;"
    AddressDistrict = AddressDistrict & "500000,500106,沙坪坝区,29.541224,106.4542;"
    AddressDistrict = AddressDistrict & "500000,500107,九龙坡区,29.523492,106.480989;"
    AddressDistrict = AddressDistrict & "500000,500108,南岸区,29.523992,106.560813;"
    AddressDistrict = AddressDistrict & "500000,500109,北碚区,29.82543,106.437868;"
    AddressDistrict = AddressDistrict & "500000,500110,綦江区,29.028091,106.651417;"
    AddressDistrict = AddressDistrict & "500000,500111,大足区,29.700498,105.715319;"
    AddressDistrict = AddressDistrict & "500000,500112,渝北区,29.601451,106.512851;"
    AddressDistrict = AddressDistrict & "500000,500113,巴南区,29.381919,106.519423;"
    AddressDistrict = AddressDistrict & "500000,500114,黔江区,29.527548,108.782577;"
    AddressDistrict = AddressDistrict & "500000,500115,长寿区,29.833671,107.074854;"
    AddressDistrict = AddressDistrict & "500000,500116,江津区,29.283387,106.253156;"
    AddressDistrict = AddressDistrict & "500000,500117,合川区,29.990993,106.265554;"
    AddressDistrict = AddressDistrict & "500000,500118,永川区,29.348748,105.894714;"
    AddressDistrict = AddressDistrict & "500000,500119,南川区,29.156646,107.098153;"
    AddressDistrict = AddressDistrict & "500000,500120,璧山区,29.593581,106.231126;"
    AddressDistrict = AddressDistrict & "500000,500151,铜梁区,29.839944,106.054948;"
    AddressDistrict = AddressDistrict & "500000,500152,潼南区,30.189554,105.841818;"
    AddressDistrict = AddressDistrict & "500000,500153,荣昌区,29.403627,105.594061;"
    AddressDistrict = AddressDistrict & "500000,500154,开州区,31.167735,108.413317;"
    AddressDistrict = AddressDistrict & "500000,500155,梁平区,30.672168,107.800034;"
    AddressDistrict = AddressDistrict & "500000,500156,武隆区,29.32376,107.75655;"
    AddressDistrict = AddressDistrict & "500000,500229,城口县,31.946293,108.6649;"
    AddressDistrict = AddressDistrict & "500000,500230,丰都县,29.866424,107.73248;"
    AddressDistrict = AddressDistrict & "500000,500231,垫江县,30.330012,107.348692;"
    AddressDistrict = AddressDistrict & "500000,500233,忠县,30.291537,108.037518;"
    AddressDistrict = AddressDistrict & "500000,500235,云阳县,30.930529,108.697698;"
    AddressDistrict = AddressDistrict & "500000,500236,奉节县,31.019967,109.465774;"
    AddressDistrict = AddressDistrict & "500000,500237,巫山县,31.074843,109.878928;"
    AddressDistrict = AddressDistrict & "500000,500238,巫溪县,31.3966,109.628912;"
    AddressDistrict = AddressDistrict & "500000,500240,石柱土家族自治县,29.99853,108.112448;"
    AddressDistrict = AddressDistrict & "500000,500241,秀山土家族苗族自治县,28.444772,108.996043;"
    AddressDistrict = AddressDistrict & "500000,500242,酉阳土家族苗族自治县,28.839828,108.767201;"
    AddressDistrict = AddressDistrict & "500000,500243,彭水苗族土家族自治县,29.293856,108.166551;"
    AddressDistrict = AddressDistrict & "810000,810001,中西区,22.28198083,114.1543731;"
    AddressDistrict = AddressDistrict & "810000,810002,湾仔区,22.27638889,114.1829153;"
    AddressDistrict = AddressDistrict & "810000,810003,东区,22.27969306,114.2260031;"
    AddressDistrict = AddressDistrict & "810000,810004,南区,22.24589667,114.1600117;"
    AddressDistrict = AddressDistrict & "810000,810005,油尖旺区,22.31170389,114.1733317;"
    AddressDistrict = AddressDistrict & "810000,810006,深水埗区,22.33385417,114.1632417;"
    AddressDistrict = AddressDistrict & "810000,810007,九龙城区,22.31251,114.1928467;"
    AddressDistrict = AddressDistrict & "810000,810008,黄大仙区,22.33632056,114.2038856;"
    AddressDistrict = AddressDistrict & "810000,810009,观塘区,22.32083778,114.2140542;"
    AddressDistrict = AddressDistrict & "810000,810010,荃湾区,22.36830667,114.1210792;"
    AddressDistrict = AddressDistrict & "810000,810011,屯门区,22.39384417,113.9765742;"
    AddressDistrict = AddressDistrict & "810000,810012,元朗区,22.44142833,114.0324381;"
    AddressDistrict = AddressDistrict & "810000,810013,北区,22.49610389,114.1473639;"
    AddressDistrict = AddressDistrict & "810000,810014,大埔区,22.44565306,114.1717431;"
    AddressDistrict = AddressDistrict & "810000,810015,西贡区,22.31421306,114.264645;"
    AddressDistrict = AddressDistrict & "810000,810016,沙田区,22.37953167,114.1953653;"
    AddressDistrict = AddressDistrict & "810000,810017,葵青区,22.36387667,114.1393194;"
    AddressDistrict = AddressDistrict & "810000,810018,离岛区,22.28640778,113.94612;"
    AddressDistrict = AddressDistrict & "820000,820001,花地玛堂区,22.20787,113.5528956;"
    AddressDistrict = AddressDistrict & "820000,820002,花王堂区,22.1992075,113.5489608;"
    AddressDistrict = AddressDistrict & "820000,820003,望德堂区,22.19372083,113.5501828;"
    AddressDistrict = AddressDistrict & "820000,820004,大堂区,22.18853944,113.5536475;"
    AddressDistrict = AddressDistrict & "820000,820005,风顺堂区,22.18736806,113.5419278;"
    AddressDistrict = AddressDistrict & "820000,820006,嘉模堂区,22.15375944,113.5587044;"
    AddressDistrict = AddressDistrict & "820000,820007,路凼填海区,22.13663,113.5695992;"
    AddressDistrict = AddressDistrict & "820000,820008,圣方济各堂区,22.12348639,113.5599542;"
    AddressDistrict = AddressDistrict & "130100,130102,长安区,38.047501,114.548151;"
    AddressDistrict = AddressDistrict & "130100,130104,桥西区,38.028383,114.462931;"
    AddressDistrict = AddressDistrict & "130100,130105,新华区,38.067142,114.465974;"
    AddressDistrict = AddressDistrict & "130100,130107,井陉矿区,38.069748,114.058178;"
    AddressDistrict = AddressDistrict & "130100,130108,裕华区,38.027696,114.533257;"
    AddressDistrict = AddressDistrict & "130100,130109,藁城区,38.033767,114.849647;"
    AddressDistrict = AddressDistrict & "130100,130110,鹿泉区,38.093994,114.321023;"
    AddressDistrict = AddressDistrict & "130100,130111,栾城区,37.886911,114.654281;"
    AddressDistrict = AddressDistrict & "130100,130121,井陉县,38.033614,114.144488;"
    AddressDistrict = AddressDistrict & "130100,130123,正定县,38.147835,114.569887;"
    AddressDistrict = AddressDistrict & "130100,130125,行唐县,38.437422,114.552734;"
    AddressDistrict = AddressDistrict & "130100,130126,灵寿县,38.306546,114.37946;"
    AddressDistrict = AddressDistrict & "130100,130127,高邑县,37.605714,114.610699;"
    AddressDistrict = AddressDistrict & "130100,130128,深泽县,38.18454,115.200207;"
    AddressDistrict = AddressDistrict & "130100,130129,赞皇县,37.660199,114.387756;"
    AddressDistrict = AddressDistrict & "130100,130130,无极县,38.176376,114.977845;"
    AddressDistrict = AddressDistrict & "130100,130131,平山县,38.259311,114.184144;"
    AddressDistrict = AddressDistrict & "130100,130132,元氏县,37.762514,114.52618;"
    AddressDistrict = AddressDistrict & "130100,130133,赵县,37.754341,114.775362;"
    AddressDistrict = AddressDistrict & "130100,130181,辛集市,37.92904,115.217451;"
    AddressDistrict = AddressDistrict & "130100,130183,晋州市,38.027478,115.044886;"
    AddressDistrict = AddressDistrict & "130100,130184,新乐市,38.344768,114.68578;"
    AddressDistrict = AddressDistrict & "130200,130202,路南区,39.615162,118.210821;"
    AddressDistrict = AddressDistrict & "130200,130203,路北区,39.628538,118.174736;"
    AddressDistrict = AddressDistrict & "130200,130204,古冶区,39.715736,118.45429;"
    AddressDistrict = AddressDistrict & "130200,130205,开平区,39.676171,118.264425;"
    AddressDistrict = AddressDistrict & "130200,130207,丰南区,39.56303,118.110793;"
    AddressDistrict = AddressDistrict & "130200,130208,丰润区,39.831363,118.155779;"
    AddressDistrict = AddressDistrict & "130200,130209,曹妃甸区,39.278277,118.446585;"
    AddressDistrict = AddressDistrict & "130200,130224,滦南县,39.506201,118.681552;"
    AddressDistrict = AddressDistrict & "130200,130225,乐亭县,39.42813,118.905341;"
    AddressDistrict = AddressDistrict & "130200,130227,迁西县,40.146238,118.305139;"
    AddressDistrict = AddressDistrict & "130200,130229,玉田县,39.887323,117.753665;"
    AddressDistrict = AddressDistrict & "130200,130281,遵化市,40.188616,117.965875;"
    AddressDistrict = AddressDistrict & "130200,130283,迁安市,40.012108,118.701933;"
    AddressDistrict = AddressDistrict & "130200,130284,滦州市,39.74485,118.699546;"
    AddressDistrict = AddressDistrict & "130300,130302,海港区,39.943458,119.596224;"
    AddressDistrict = AddressDistrict & "130300,130303,山海关区,39.998023,119.753591;"
    AddressDistrict = AddressDistrict & "130300,130304,北戴河区,39.825121,119.486286;"
    AddressDistrict = AddressDistrict & "130300,130306,抚宁区,39.887053,119.240651;"
    AddressDistrict = AddressDistrict & "130300,130321,青龙满族自治县,40.406023,118.954555;"
    AddressDistrict = AddressDistrict & "130300,130322,昌黎县,39.709729,119.164541;"
    AddressDistrict = AddressDistrict & "130300,130324,卢龙县,39.891639,118.881809;"
    AddressDistrict = AddressDistrict & "130400,130402,邯山区,36.603196,114.484989;"
    AddressDistrict = AddressDistrict & "130400,130403,丛台区,36.611082,114.494703;"
    AddressDistrict = AddressDistrict & "130400,130404,复兴区,36.615484,114.458242;"
    AddressDistrict = AddressDistrict & "130400,130406,峰峰矿区,36.420487,114.209936;"
    AddressDistrict = AddressDistrict & "130400,130407,肥乡区,36.555778,114.805154;"
    AddressDistrict = AddressDistrict & "130400,130408,永年区,36.776413,114.496162;"
    AddressDistrict = AddressDistrict & "130400,130423,临漳县,36.337604,114.610703;"
    AddressDistrict = AddressDistrict & "130400,130424,成安县,36.443832,114.680356;"
    AddressDistrict = AddressDistrict & "130400,130425,大名县,36.283316,115.152586;"
    AddressDistrict = AddressDistrict & "130400,130426,涉县,36.563143,113.673297;"
    AddressDistrict = AddressDistrict & "130400,130427,磁县,36.367673,114.38208;"
    AddressDistrict = AddressDistrict & "130400,130430,邱县,36.81325,115.168584;"
    AddressDistrict = AddressDistrict & "130400,130431,鸡泽县,36.914908,114.878517;"
    AddressDistrict = AddressDistrict & "130400,130432,广平县,36.483603,114.950859;"
    AddressDistrict = AddressDistrict & "130400,130433,馆陶县,36.539461,115.289057;"
    AddressDistrict = AddressDistrict & "130400,130434,魏县,36.354248,114.93411;"
    AddressDistrict = AddressDistrict & "130400,130435,曲周县,36.773398,114.957588;"
    AddressDistrict = AddressDistrict & "130400,130481,武安市,36.696115,114.194581;"
    AddressDistrict = AddressDistrict & "130500,130502,襄都区,37.064125,114.507131;"
    AddressDistrict = AddressDistrict & "130500,130503,信都区,37.068009,114.473687;"
    AddressDistrict = AddressDistrict & "130500,130505,任泽区,37.129952,114.684469;"
    AddressDistrict = AddressDistrict & "130500,130506,南和区,37.003812,114.691377;"
    AddressDistrict = AddressDistrict & "130500,130522,临城县,37.444009,114.506873;"
    AddressDistrict = AddressDistrict & "130500,130523,内丘县,37.287663,114.511523;"
    AddressDistrict = AddressDistrict & "130500,130524,柏乡县,37.483596,114.693382;"
    AddressDistrict = AddressDistrict & "130500,130525,隆尧县,37.350925,114.776348;"
    AddressDistrict = AddressDistrict & "130500,130528,宁晋县,37.618956,114.921027;"
    AddressDistrict = AddressDistrict & "130500,130529,巨鹿县,37.21768,115.038782;"
    AddressDistrict = AddressDistrict & "130500,130530,新河县,37.526216,115.247537;"
    AddressDistrict = AddressDistrict & "130500,130531,广宗县,37.075548,115.142797;"
    AddressDistrict = AddressDistrict & "130500,130532,平乡县,37.069404,115.029218;"
    AddressDistrict = AddressDistrict & "130500,130533,威县,36.983272,115.272749;"
    AddressDistrict = AddressDistrict & "130500,130534,清河县,37.059991,115.668999;"
    AddressDistrict = AddressDistrict & "130500,130535,临西县,36.8642,115.498684;"
    AddressDistrict = AddressDistrict & "130500,130581,南宫市,37.359668,115.398102;"
    AddressDistrict = AddressDistrict & "130500,130582,沙河市,36.861903,114.504902;"
    AddressDistrict = AddressDistrict & "130600,130602,竞秀区,38.88662,115.470659;"
    AddressDistrict = AddressDistrict & "130600,130606,莲池区,38.865005,115.500934;"
    AddressDistrict = AddressDistrict & "130600,130607,满城区,38.95138,115.32442;"
    AddressDistrict = AddressDistrict & "130600,130608,清苑区,38.771012,115.492221;"
    AddressDistrict = AddressDistrict & "130600,130609,徐水区,39.020395,115.64941;"
    AddressDistrict = AddressDistrict & "130600,130623,涞水县,39.393148,115.711985;"
    AddressDistrict = AddressDistrict & "130600,130624,阜平县,38.847276,114.198801;"
    AddressDistrict = AddressDistrict & "130600,130626,定兴县,39.266195,115.796895;"
    AddressDistrict = AddressDistrict & "130600,130627,唐县,38.748542,114.981241;"
    AddressDistrict = AddressDistrict & "130600,130628,高阳县,38.690092,115.778878;"
    AddressDistrict = AddressDistrict & "130600,130629,容城县,39.05282,115.866247;"
    AddressDistrict = AddressDistrict & "130600,130630,涞源县,39.35755,114.692567;"
    AddressDistrict = AddressDistrict & "130600,130631,望都县,38.707448,115.154009;"
    AddressDistrict = AddressDistrict & "130600,130632,安新县,38.929912,115.931979;"
    AddressDistrict = AddressDistrict & "130600,130633,易县,39.35297,115.501146;"
    AddressDistrict = AddressDistrict & "130600,130634,曲阳县,38.619992,114.704055;"
    AddressDistrict = AddressDistrict & "130600,130635,蠡县,38.496429,115.583631;"
    AddressDistrict = AddressDistrict & "130600,130636,顺平县,38.845127,115.132749;"
    AddressDistrict = AddressDistrict & "130600,130637,博野县,38.458271,115.461798;"
    AddressDistrict = AddressDistrict & "130600,130638,雄县,38.990819,116.107474;"
    AddressDistrict = AddressDistrict & "130600,130681,涿州市,39.485765,115.973409;"
    AddressDistrict = AddressDistrict & "130600,130682,定州市,38.517602,114.991389;"
    AddressDistrict = AddressDistrict & "130600,130683,安国市,38.421367,115.33141;"
    AddressDistrict = AddressDistrict & "130600,130684,高碑店市,39.327689,115.882704;"
    AddressDistrict = AddressDistrict & "130700,130702,桥东区,40.813875,114.885658;"
    AddressDistrict = AddressDistrict & "130700,130703,桥西区,40.824385,114.882127;"
    AddressDistrict = AddressDistrict & "130700,130705,宣化区,40.609368,115.0632;"
    AddressDistrict = AddressDistrict & "130700,130706,下花园区,40.488645,115.281002;"
    AddressDistrict = AddressDistrict & "130700,130708,万全区,40.765136,114.736131;"
    AddressDistrict = AddressDistrict & "130700,130709,崇礼区,40.971302,115.281652;"
    AddressDistrict = AddressDistrict & "130700,130722,张北县,41.151713,114.715951;"
    AddressDistrict = AddressDistrict & "130700,130723,康保县,41.850046,114.615809;"
    AddressDistrict = AddressDistrict & "130700,130724,沽源县,41.667419,115.684836;"
    AddressDistrict = AddressDistrict & "130700,130725,尚义县,41.080091,113.977713;"
    AddressDistrict = AddressDistrict & "130700,130726,蔚县,39.837181,114.582695;"
    AddressDistrict = AddressDistrict & "130700,130727,阳原县,40.113419,114.167343;"
    AddressDistrict = AddressDistrict & "130700,130728,怀安县,40.671274,114.422364;"
    AddressDistrict = AddressDistrict & "130700,130730,怀来县,40.405405,115.520846;"
    AddressDistrict = AddressDistrict & "130700,130731,涿鹿县,40.378701,115.219246;"
    AddressDistrict = AddressDistrict & "130700,130732,赤城县,40.912081,115.832708;"
    AddressDistrict = AddressDistrict & "130800,130802,双桥区,40.976204,117.939152;"
    AddressDistrict = AddressDistrict & "130800,130803,双滦区,40.959756,117.797485;"
    AddressDistrict = AddressDistrict & "130800,130804,鹰手营子矿区,40.546956,117.661154;"
    AddressDistrict = AddressDistrict & "130800,130821,承德县,40.768637,118.172496;"
    AddressDistrict = AddressDistrict & "130800,130822,兴隆县,40.418525,117.507098;"
    AddressDistrict = AddressDistrict & "130800,130824,滦平县,40.936644,117.337124;"
    AddressDistrict = AddressDistrict & "130800,130825,隆化县,41.316667,117.736343;"
    AddressDistrict = AddressDistrict & "130800,130826,丰宁满族自治县,41.209903,116.65121;"
    AddressDistrict = AddressDistrict & "130800,130827,宽城满族自治县,40.607981,118.488642;"
    AddressDistrict = AddressDistrict & "130800,130828,围场满族蒙古族自治县,41.949404,117.764086;"
    AddressDistrict = AddressDistrict & "130800,130881,平泉市,41.00561,118.690238;"
    AddressDistrict = AddressDistrict & "130900,130902,新华区,38.308273,116.873049;"
    AddressDistrict = AddressDistrict & "130900,130903,运河区,38.307405,116.840063;"
    AddressDistrict = AddressDistrict & "130900,130921,沧县,38.219856,117.007478;"
    AddressDistrict = AddressDistrict & "130900,130922,青县,38.569646,116.838384;"
    AddressDistrict = AddressDistrict & "130900,130923,东光县,37.88655,116.542062;"
    AddressDistrict = AddressDistrict & "130900,130924,海兴县,38.141582,117.496606;"
    AddressDistrict = AddressDistrict & "130900,130925,盐山县,38.056141,117.229814;"
    AddressDistrict = AddressDistrict & "130900,130926,肃宁县,38.427102,115.835856;"
    AddressDistrict = AddressDistrict & "130900,130927,南皮县,38.042439,116.709171;"
    AddressDistrict = AddressDistrict & "130900,130928,吴桥县,37.628182,116.391512;"
    AddressDistrict = AddressDistrict & "130900,130929,献县,38.189661,116.123844;"
    AddressDistrict = AddressDistrict & "130900,130930,孟村回族自治县,38.057953,117.105104;"
    AddressDistrict = AddressDistrict & "130900,130981,泊头市,38.073479,116.570163;"
    AddressDistrict = AddressDistrict & "130900,130982,任丘市,38.706513,116.106764;"
    AddressDistrict = AddressDistrict & "130900,130983,黄骅市,38.369238,117.343803;"
    AddressDistrict = AddressDistrict & "130900,130984,河间市,38.44149,116.089452;"
    AddressDistrict = AddressDistrict & "131000,131002,安次区,39.502569,116.694544;"
    AddressDistrict = AddressDistrict & "131000,131003,广阳区,39.521931,116.713708;"
    AddressDistrict = AddressDistrict & "131000,131022,固安县,39.436468,116.299894;"
    AddressDistrict = AddressDistrict & "131000,131023,永清县,39.319717,116.498089;"
    AddressDistrict = AddressDistrict & "131000,131024,香河县,39.757212,117.007161;"
    AddressDistrict = AddressDistrict & "131000,131025,大城县,38.699215,116.640735;"
    AddressDistrict = AddressDistrict & "131000,131026,文安县,38.866801,116.460107;"
    AddressDistrict = AddressDistrict & "131000,131028,大厂回族自治县,39.889266,116.986501;"
    AddressDistrict = AddressDistrict & "131000,131081,霸州市,39.117331,116.392021;"
    AddressDistrict = AddressDistrict & "131000,131082,三河市,39.982778,117.077018;"
    AddressDistrict = AddressDistrict & "131100,131102,桃城区,37.732237,115.694945;"
    AddressDistrict = AddressDistrict & "131100,131103,冀州区,37.542788,115.579173;"
    AddressDistrict = AddressDistrict & "131100,131121,枣强县,37.511512,115.726499;"
    AddressDistrict = AddressDistrict & "131100,131122,武邑县,37.803774,115.892415;"
    AddressDistrict = AddressDistrict & "131100,131123,武强县,38.03698,115.970236;"
    AddressDistrict = AddressDistrict & "131100,131124,饶阳县,38.232671,115.726577;"
    AddressDistrict = AddressDistrict & "131100,131125,安平县,38.233511,115.519627;"
    AddressDistrict = AddressDistrict & "131100,131126,故城县,37.350981,115.966747;"
    AddressDistrict = AddressDistrict & "131100,131127,景县,37.686622,116.258446;"
    AddressDistrict = AddressDistrict & "131100,131128,阜城县,37.869945,116.164727;"
    AddressDistrict = AddressDistrict & "131100,131182,深州市,38.00347,115.554596;"
    AddressDistrict = AddressDistrict & "140100,140105,小店区,37.817974,112.564273;"
    AddressDistrict = AddressDistrict & "140100,140106,迎泽区,37.855804,112.558851;"
    AddressDistrict = AddressDistrict & "140100,140107,杏花岭区,37.879291,112.560743;"
    AddressDistrict = AddressDistrict & "140100,140108,尖草坪区,37.939893,112.487122;"
    AddressDistrict = AddressDistrict & "140100,140109,万柏林区,37.862653,112.522258;"
    AddressDistrict = AddressDistrict & "140100,140110,晋源区,37.715619,112.477849;"
    AddressDistrict = AddressDistrict & "140100,140121,清徐县,37.60729,112.357961;"
    AddressDistrict = AddressDistrict & "140100,140122,阳曲县,38.058797,112.673818;"
    AddressDistrict = AddressDistrict & "140100,140123,娄烦县,38.066035,111.793798;"
    AddressDistrict = AddressDistrict & "140100,140181,古交市,37.908534,112.174353;"
    AddressDistrict = AddressDistrict & "140200,140212,新荣区,40.258269,113.141044;"
    AddressDistrict = AddressDistrict & "140200,140213,平城区,40.075667,113.298027;"
    AddressDistrict = AddressDistrict & "140200,140214,云冈区,40.005405,113.149693;"
    AddressDistrict = AddressDistrict & "140200,140215,云州区,40.040295,113.61244;"
    AddressDistrict = AddressDistrict & "140200,140221,阳高县,40.364927,113.749871;"
    AddressDistrict = AddressDistrict & "140200,140222,天镇县,40.421336,114.09112;"
    AddressDistrict = AddressDistrict & "140200,140223,广灵县,39.763051,114.279252;"
    AddressDistrict = AddressDistrict & "140200,140224,灵丘县,39.438867,114.23576;"
    AddressDistrict = AddressDistrict & "140200,140225,浑源县,39.699099,113.698091;"
    AddressDistrict = AddressDistrict & "140200,140226,左云县,40.012873,112.70641;"
    AddressDistrict = AddressDistrict & "140300,140302,城区,37.860938,113.586513;"
    AddressDistrict = AddressDistrict & "140300,140303,矿区,37.870085,113.559066;"
    AddressDistrict = AddressDistrict & "140300,140311,郊区,37.94096,113.58664;"
    AddressDistrict = AddressDistrict & "140300,140321,平定县,37.800289,113.631049;"
    AddressDistrict = AddressDistrict & "140300,140322,盂县,38.086131,113.41223;"
    AddressDistrict = AddressDistrict & "140400,140403,潞州区,36.187895,113.114107;"
    AddressDistrict = AddressDistrict & "140400,140404,上党区,36.052438,113.056679;"
    AddressDistrict = AddressDistrict & "140400,140405,屯留区,36.314072,112.892741;"
    AddressDistrict = AddressDistrict & "140400,140406,潞城区,36.332232,113.223245;"
    AddressDistrict = AddressDistrict & "140400,140423,襄垣县,36.532854,113.050094;"
    AddressDistrict = AddressDistrict & "140400,140425,平顺县,36.200202,113.438791;"
    AddressDistrict = AddressDistrict & "140400,140426,黎城县,36.502971,113.387366;"
    AddressDistrict = AddressDistrict & "140400,140427,壶关县,36.110938,113.206138;"
    AddressDistrict = AddressDistrict & "140400,140428,长子县,36.119484,112.884656;"
    AddressDistrict = AddressDistrict & "140400,140429,武乡县,36.834315,112.8653;"
    AddressDistrict = AddressDistrict & "140400,140430,沁县,36.757123,112.70138;"
    AddressDistrict = AddressDistrict & "140400,140431,沁源县,36.500777,112.340878;"
    AddressDistrict = AddressDistrict & "140500,140502,城区,35.496641,112.853106;"
    AddressDistrict = AddressDistrict & "140500,140521,沁水县,35.689472,112.187213;"
    AddressDistrict = AddressDistrict & "140500,140522,阳城县,35.482177,112.422014;"
    AddressDistrict = AddressDistrict & "140500,140524,陵川县,35.775614,113.278877;"
    AddressDistrict = AddressDistrict & "140500,140525,泽州县,35.617221,112.899137;"
    AddressDistrict = AddressDistrict & "140500,140581,高平市,35.791355,112.930691;"
    AddressDistrict = AddressDistrict & "140600,140602,朔城区,39.324525,112.428676;"
    AddressDistrict = AddressDistrict & "140600,140603,平鲁区,39.515603,112.295227;"
    AddressDistrict = AddressDistrict & "140600,140621,山阴县,39.52677,112.816396;"
    AddressDistrict = AddressDistrict & "140600,140622,应县,39.559187,113.187505;"
    AddressDistrict = AddressDistrict & "140600,140623,右玉县,39.988812,112.465588;"
    AddressDistrict = AddressDistrict & "140600,140681,怀仁市,39.820789,113.100511;"
    AddressDistrict = AddressDistrict & "140700,140702,榆次区,37.6976,112.740056;"
    AddressDistrict = AddressDistrict & "140700,140703,太谷区,37.424595,112.554103;"
    AddressDistrict = AddressDistrict & "140700,140721,榆社县,37.069019,112.973521;"
    AddressDistrict = AddressDistrict & "140700,140722,左权县,37.079672,113.377834;"
    AddressDistrict = AddressDistrict & "140700,140723,和顺县,37.327027,113.572919;"
    AddressDistrict = AddressDistrict & "140700,140724,昔阳县,37.60437,113.706166;"
    AddressDistrict = AddressDistrict & "140700,140725,寿阳县,37.891136,113.177708;"
    AddressDistrict = AddressDistrict & "140700,140727,祁县,37.358739,112.330532;"
    AddressDistrict = AddressDistrict & "140700,140728,平遥县,37.195474,112.174059;"
    AddressDistrict = AddressDistrict & "140700,140729,灵石县,36.847469,111.772759;"
    AddressDistrict = AddressDistrict & "140700,140781,介休市,37.027616,111.913857;"
    AddressDistrict = AddressDistrict & "140800,140802,盐湖区,35.025643,111.000627;"
    AddressDistrict = AddressDistrict & "140800,140821,临猗县,35.141883,110.77493;"
    AddressDistrict = AddressDistrict & "140800,140822,万荣县,35.417042,110.843561;"
    AddressDistrict = AddressDistrict & "140800,140823,闻喜县,35.353839,111.220306;"
    AddressDistrict = AddressDistrict & "140800,140824,稷山县,35.600412,110.978996;"
    AddressDistrict = AddressDistrict & "140800,140825,新绛县,35.613697,111.225205;"
    AddressDistrict = AddressDistrict & "140800,140826,绛县,35.49045,111.576182;"
    AddressDistrict = AddressDistrict & "140800,140827,垣曲县,35.298293,111.67099;"
    AddressDistrict = AddressDistrict & "140800,140828,夏县,35.140441,111.223174;"
    AddressDistrict = AddressDistrict & "140800,140829,平陆县,34.837256,111.212377;"
    AddressDistrict = AddressDistrict & "140800,140830,芮城县,34.694769,110.69114;"
    AddressDistrict = AddressDistrict & "140800,140881,永济市,34.865125,110.447984;"
    AddressDistrict = AddressDistrict & "140800,140882,河津市,35.59715,110.710268;"
    AddressDistrict = AddressDistrict & "140900,140902,忻府区,38.417743,112.734112;"
    AddressDistrict = AddressDistrict & "140900,140921,定襄县,38.484948,112.963231;"
    AddressDistrict = AddressDistrict & "140900,140922,五台县,38.725711,113.259012;"
    AddressDistrict = AddressDistrict & "140900,140923,代县,39.065138,112.962519;"
    AddressDistrict = AddressDistrict & "140900,140924,繁峙县,39.188104,113.267707;"
    AddressDistrict = AddressDistrict & "140900,140925,宁武县,39.001718,112.307936;"
    AddressDistrict = AddressDistrict & "140900,140926,静乐县,38.355947,111.940231;"
    AddressDistrict = AddressDistrict & "140900,140927,神池县,39.088467,112.200438;"
    AddressDistrict = AddressDistrict & "140900,140928,五寨县,38.912761,111.841015;"
    AddressDistrict = AddressDistrict & "140900,140929,岢岚县,38.705625,111.56981;"
    AddressDistrict = AddressDistrict & "140900,140930,河曲县,39.381895,111.146609;"
    AddressDistrict = AddressDistrict & "140900,140931,保德县,39.022576,111.085688;"
    AddressDistrict = AddressDistrict & "140900,140932,偏关县,39.442153,111.500477;"
    AddressDistrict = AddressDistrict & "140900,140981,原平市,38.729186,112.713132;"
    AddressDistrict = AddressDistrict & "141000,141002,尧都区,36.080366,111.522945;"
    AddressDistrict = AddressDistrict & "141000,141021,曲沃县,35.641387,111.475529;"
    AddressDistrict = AddressDistrict & "141000,141022,翼城县,35.738621,111.713508;"
    AddressDistrict = AddressDistrict & "141000,141023,襄汾县,35.876139,111.442932;"
    AddressDistrict = AddressDistrict & "141000,141024,洪洞县,36.255742,111.673692;"
    AddressDistrict = AddressDistrict & "141000,141025,古县,36.26855,111.920207;"
    AddressDistrict = AddressDistrict & "141000,141026,安泽县,36.146032,112.251372;"
    AddressDistrict = AddressDistrict & "141000,141027,浮山县,35.971359,111.850039;"
    AddressDistrict = AddressDistrict & "141000,141028,吉县,36.099355,110.682853;"
    AddressDistrict = AddressDistrict & "141000,141029,乡宁县,35.975402,110.857365;"
    AddressDistrict = AddressDistrict & "141000,141030,大宁县,36.46383,110.751283;"
    AddressDistrict = AddressDistrict & "141000,141031,隰县,36.692675,110.935809;"
    AddressDistrict = AddressDistrict & "141000,141032,永和县,36.760614,110.631276;"
    AddressDistrict = AddressDistrict & "141000,141033,蒲县,36.411682,111.09733;"
    AddressDistrict = AddressDistrict & "141000,141034,汾西县,36.653368,111.563021;"
    AddressDistrict = AddressDistrict & "141000,141081,侯马市,35.620302,111.371272;"
    AddressDistrict = AddressDistrict & "141000,141082,霍州市,36.57202,111.723103;"
    AddressDistrict = AddressDistrict & "141100,141102,离石区,37.524037,111.134462;"
    AddressDistrict = AddressDistrict & "141100,141121,文水县,37.436314,112.032595;"
    AddressDistrict = AddressDistrict & "141100,141122,交城县,37.555155,112.159154;"
    AddressDistrict = AddressDistrict & "141100,141123,兴县,38.464136,111.124816;"
    AddressDistrict = AddressDistrict & "141100,141124,临县,37.960806,110.995963;"
    AddressDistrict = AddressDistrict & "141100,141125,柳林县,37.431664,110.89613;"
    AddressDistrict = AddressDistrict & "141100,141126,石楼县,36.999426,110.837119;"
    AddressDistrict = AddressDistrict & "141100,141127,岚县,38.278654,111.671555;"
    AddressDistrict = AddressDistrict & "141100,141128,方山县,37.892632,111.238885;"
    AddressDistrict = AddressDistrict & "141100,141129,中阳县,37.342054,111.193319;"
    AddressDistrict = AddressDistrict & "141100,141130,交口县,36.983068,111.183188;"
    AddressDistrict = AddressDistrict & "141100,141181,孝义市,37.144474,111.781568;"
    AddressDistrict = AddressDistrict & "141100,141182,汾阳市,37.267742,111.785273;"
    AddressDistrict = AddressDistrict & "150100,150102,新城区,40.826225,111.685964;"
    AddressDistrict = AddressDistrict & "150100,150103,回民区,40.815149,111.662162;"
    AddressDistrict = AddressDistrict & "150100,150104,玉泉区,40.799421,111.66543;"
    AddressDistrict = AddressDistrict & "150100,150105,赛罕区,40.807834,111.698463;"
    AddressDistrict = AddressDistrict & "150100,150121,土默特左旗,40.720416,111.133615;"
    AddressDistrict = AddressDistrict & "150100,150122,托克托县,40.276729,111.197317;"
    AddressDistrict = AddressDistrict & "150100,150123,和林格尔县,40.380288,111.824143;"
    AddressDistrict = AddressDistrict & "150100,150124,清水河县,39.912479,111.67222;"
    AddressDistrict = AddressDistrict & "150100,150125,武川县,41.094483,111.456563;"
    AddressDistrict = AddressDistrict & "150200,150202,东河区,40.587056,110.026895;"
    AddressDistrict = AddressDistrict & "150200,150203,昆都仑区,40.661345,109.822932;"
    AddressDistrict = AddressDistrict & "150200,150204,青山区,40.668558,109.880049;"
    AddressDistrict = AddressDistrict & "150200,150205,石拐区,40.672094,110.272565;"
    AddressDistrict = AddressDistrict & "150200,150206,白云鄂博矿区,41.769246,109.97016;"
    AddressDistrict = AddressDistrict & "150200,150207,九原区,40.600581,109.968122;"
    AddressDistrict = AddressDistrict & "150200,150221,土默特右旗,40.566434,110.526766;"
    AddressDistrict = AddressDistrict & "150200,150222,固阳县,41.030004,110.063421;"
    AddressDistrict = AddressDistrict & "150200,150223,达尔罕茂明安联合旗,41.702836,110.438452;"
    AddressDistrict = AddressDistrict & "150300,150302,海勃湾区,39.673527,106.817762;"
    AddressDistrict = AddressDistrict & "150300,150303,海南区,39.44153,106.884789;"
    AddressDistrict = AddressDistrict & "150300,150304,乌达区,39.502288,106.722711;"
    AddressDistrict = AddressDistrict & "150400,150402,红山区,42.269732,118.961087;"
    AddressDistrict = AddressDistrict & "150400,150403,元宝山区,42.041168,119.289877;"
    AddressDistrict = AddressDistrict & "150400,150404,松山区,42.281046,118.938958;"
    AddressDistrict = AddressDistrict & "150400,150421,阿鲁科尔沁旗,43.87877,120.094969;"
    AddressDistrict = AddressDistrict & "150400,150422,巴林左旗,43.980715,119.391737;"
    AddressDistrict = AddressDistrict & "150400,150423,巴林右旗,43.528963,118.678347;"
    AddressDistrict = AddressDistrict & "150400,150424,林西县,43.605326,118.05775;"
    AddressDistrict = AddressDistrict & "150400,150425,克什克腾旗,43.256233,117.542465;"
    AddressDistrict = AddressDistrict & "150400,150426,翁牛特旗,42.937128,119.022619;"
    AddressDistrict = AddressDistrict & "150400,150428,喀喇沁旗,41.92778,118.708572;"
    AddressDistrict = AddressDistrict & "150400,150429,宁城县,41.598692,119.339242;"
    AddressDistrict = AddressDistrict & "150400,150430,敖汉旗,42.287012,119.906486;"
    AddressDistrict = AddressDistrict & "150500,150502,科尔沁区,43.617422,122.264042;"
    AddressDistrict = AddressDistrict & "150500,150521,科尔沁左翼中旗,44.127166,123.313873;"
    AddressDistrict = AddressDistrict & "150500,150522,科尔沁左翼后旗,42.954564,122.355155;"
    AddressDistrict = AddressDistrict & "150500,150523,开鲁县,43.602432,121.308797;"
    AddressDistrict = AddressDistrict & "150500,150524,库伦旗,42.734692,121.774886;"
    AddressDistrict = AddressDistrict & "150500,150525,奈曼旗,42.84685,120.662543;"
    AddressDistrict = AddressDistrict & "150500,150526,扎鲁特旗,44.555294,120.905275;"
    AddressDistrict = AddressDistrict & "150500,150581,霍林郭勒市,45.532361,119.657862;"
    AddressDistrict = AddressDistrict & "150600,150602,东胜区,39.81788,109.98945;"
    AddressDistrict = AddressDistrict & "150600,150603,康巴什区,39.607472,109.790076;"
    AddressDistrict = AddressDistrict & "150600,150621,达拉特旗,40.404076,110.040281;"
    AddressDistrict = AddressDistrict & "150600,150622,准格尔旗,39.865221,111.238332;"
    AddressDistrict = AddressDistrict & "150600,150623,鄂托克前旗,38.183257,107.48172;"
    AddressDistrict = AddressDistrict & "150600,150624,鄂托克旗,39.095752,107.982604;"
    AddressDistrict = AddressDistrict & "150600,150625,杭锦旗,39.831789,108.736324;"
    AddressDistrict = AddressDistrict & "150600,150626,乌审旗,38.596611,108.842454;"
    AddressDistrict = AddressDistrict & "150600,150627,伊金霍洛旗,39.604312,109.787402;"
    AddressDistrict = AddressDistrict & "150700,150702,海拉尔区,49.213889,119.764923;"
    AddressDistrict = AddressDistrict & "150700,150703,扎赉诺尔区,49.456567,117.716373;"
    AddressDistrict = AddressDistrict & "150700,150721,阿荣旗,48.130503,123.464615;"
    AddressDistrict = AddressDistrict & "150700,150722,莫力达瓦达斡尔族自治旗,48.478385,124.507401;"
    AddressDistrict = AddressDistrict & "150700,150723,鄂伦春自治旗,50.590177,123.725684;"
    AddressDistrict = AddressDistrict & "150700,150724,鄂温克族自治旗,49.143293,119.754041;"
    AddressDistrict = AddressDistrict & "150700,150725,陈巴尔虎旗,49.328422,119.437609;"
    AddressDistrict = AddressDistrict & "150700,150726,新巴尔虎左旗,48.216571,118.267454;"
    AddressDistrict = AddressDistrict & "150700,150727,新巴尔虎右旗,48.669134,116.825991;"
    AddressDistrict = AddressDistrict & "150700,150781,满洲里市,49.590788,117.455561;"
    AddressDistrict = AddressDistrict & "150700,150782,牙克石市,49.287024,120.729005;"
    AddressDistrict = AddressDistrict & "150700,150783,扎兰屯市,48.007412,122.744401;"
    AddressDistrict = AddressDistrict & "150700,150784,额尔古纳市,50.2439,120.178636;"
    AddressDistrict = AddressDistrict & "150700,150785,根河市,50.780454,121.532724;"
    AddressDistrict = AddressDistrict & "150800,150802,临河区,40.757092,107.417018;"
    AddressDistrict = AddressDistrict & "150800,150821,五原县,41.097639,108.270658;"
    AddressDistrict = AddressDistrict & "150800,150822,磴口县,40.330479,107.006056;"
    AddressDistrict = AddressDistrict & "150800,150823,乌拉特前旗,40.725209,108.656816;"
    AddressDistrict = AddressDistrict & "150800,150824,乌拉特中旗,41.57254,108.515255;"
    AddressDistrict = AddressDistrict & "150800,150825,乌拉特后旗,41.084307,107.074941;"
    AddressDistrict = AddressDistrict & "150800,150826,杭锦后旗,40.888797,107.147682;"
    AddressDistrict = AddressDistrict & "150900,150902,集宁区,41.034134,113.116453;"
    AddressDistrict = AddressDistrict & "150900,150921,卓资县,40.89576,112.577702;"
    AddressDistrict = AddressDistrict & "150900,150922,化德县,41.899335,114.01008;"
    AddressDistrict = AddressDistrict & "150900,150923,商都县,41.560163,113.560643;"
    AddressDistrict = AddressDistrict & "150900,150924,兴和县,40.872437,113.834009;"
    AddressDistrict = AddressDistrict & "150900,150925,凉城县,40.531627,112.500911;"
    AddressDistrict = AddressDistrict & "150900,150926,察哈尔右翼前旗,40.786859,113.211958;"
    AddressDistrict = AddressDistrict & "150900,150927,察哈尔右翼中旗,41.274212,112.633563;"
    AddressDistrict = AddressDistrict & "150900,150928,察哈尔右翼后旗,41.447213,113.1906;"
    AddressDistrict = AddressDistrict & "150900,150929,四子王旗,41.528114,111.70123;"
    AddressDistrict = AddressDistrict & "150900,150981,丰镇市,40.437534,113.163462;"
    AddressDistrict = AddressDistrict & "152200,152201,乌兰浩特市,46.077238,122.068975;"
    AddressDistrict = AddressDistrict & "152200,152202,阿尔山市,47.177,119.943656;"
    AddressDistrict = AddressDistrict & "152200,152221,科尔沁右翼前旗,46.076497,121.957544;"
    AddressDistrict = AddressDistrict & "152200,152222,科尔沁右翼中旗,45.059645,121.472818;"
    AddressDistrict = AddressDistrict & "152200,152223,扎赉特旗,46.725136,122.909332;"
    AddressDistrict = AddressDistrict & "152200,152224,突泉县,45.380986,121.564856;"
    AddressDistrict = AddressDistrict & "152500,152501,二连浩特市,43.652895,111.97981;"
    AddressDistrict = AddressDistrict & "152500,152502,锡林浩特市,43.944301,116.091903;"
    AddressDistrict = AddressDistrict & "152500,152522,阿巴嘎旗,44.022728,114.970618;"
    AddressDistrict = AddressDistrict & "152500,152523,苏尼特左旗,43.854108,113.653412;"
    AddressDistrict = AddressDistrict & "152500,152524,苏尼特右旗,42.746662,112.65539;"
    AddressDistrict = AddressDistrict & "152500,152525,东乌珠穆沁旗,45.510307,116.980022;"
    AddressDistrict = AddressDistrict & "152500,152526,西乌珠穆沁旗,44.586147,117.615249;"
    AddressDistrict = AddressDistrict & "152500,152527,太仆寺旗,41.895199,115.28728;"
    AddressDistrict = AddressDistrict & "152500,152528,镶黄旗,42.239229,113.843869;"
    AddressDistrict = AddressDistrict & "152500,152529,正镶白旗,42.286807,115.031423;"
    AddressDistrict = AddressDistrict & "152500,152530,正蓝旗,42.245895,116.003311;"
    AddressDistrict = AddressDistrict & "152500,152531,多伦县,42.197962,116.477288;"
    AddressDistrict = AddressDistrict & "152900,152921,阿拉善左旗,38.847241,105.70192;"
    AddressDistrict = AddressDistrict & "152900,152922,阿拉善右旗,39.21159,101.671984;"
    AddressDistrict = AddressDistrict & "152900,152923,额济纳旗,41.958813,101.06944;"
    AddressDistrict = AddressDistrict & "210100,210102,和平区,41.788074,123.406664;"
    AddressDistrict = AddressDistrict & "210100,210103,沈河区,41.795591,123.445696;"
    AddressDistrict = AddressDistrict & "210100,210104,大东区,41.808503,123.469956;"
    AddressDistrict = AddressDistrict & "210100,210105,皇姑区,41.822336,123.405677;"
    AddressDistrict = AddressDistrict & "210100,210106,铁西区,41.787808,123.350664;"
    AddressDistrict = AddressDistrict & "210100,210111,苏家屯区,41.665904,123.341604;"
    AddressDistrict = AddressDistrict & "210100,210112,浑南区,41.741946,123.458981;"
    AddressDistrict = AddressDistrict & "210100,210113,沈北新区,42.052312,123.521471;"
    AddressDistrict = AddressDistrict & "210100,210114,于洪区,41.795833,123.310829;"
    AddressDistrict = AddressDistrict & "210100,210115,辽中区,41.512725,122.731269;"
    AddressDistrict = AddressDistrict & "210100,210123,康平县,42.741533,123.352703;"
    AddressDistrict = AddressDistrict & "210100,210124,法库县,42.507045,123.416722;"
    AddressDistrict = AddressDistrict & "210100,210181,新民市,41.996508,122.828868;"
    AddressDistrict = AddressDistrict & "210200,210202,中山区,38.921553,121.64376;"
    AddressDistrict = AddressDistrict & "210200,210203,西岗区,38.914266,121.616112;"
    AddressDistrict = AddressDistrict & "210200,210204,沙河口区,38.912859,121.593702;"
    AddressDistrict = AddressDistrict & "210200,210211,甘井子区,38.975148,121.582614;"
    AddressDistrict = AddressDistrict & "210200,210212,旅顺口区,38.812043,121.26713;"
    AddressDistrict = AddressDistrict & "210200,210213,金州区,39.052745,121.789413;"
    AddressDistrict = AddressDistrict & "210200,210214,普兰店区,39.401555,121.9705;"
    AddressDistrict = AddressDistrict & "210200,210224,长海县,39.272399,122.587824;"
    AddressDistrict = AddressDistrict & "210200,210281,瓦房店市,39.63065,122.002656;"
    AddressDistrict = AddressDistrict & "210200,210283,庄河市,39.69829,122.970612;"
    AddressDistrict = AddressDistrict & "210300,210302,铁东区,41.110344,122.994475;"
    AddressDistrict = AddressDistrict & "210300,210303,铁西区,41.11069,122.971834;"
    AddressDistrict = AddressDistrict & "210300,210304,立山区,41.150622,123.024806;"
    AddressDistrict = AddressDistrict & "210300,210311,千山区,41.068909,122.949298;"
    AddressDistrict = AddressDistrict & "210300,210321,台安县,41.38686,122.429736;"
    AddressDistrict = AddressDistrict & "210300,210323,岫岩满族自治县,40.281509,123.28833;"
    AddressDistrict = AddressDistrict & "210300,210381,海城市,40.852533,122.752199;"
    AddressDistrict = AddressDistrict & "210400,210402,新抚区,41.86082,123.902858;"
    AddressDistrict = AddressDistrict & "210400,210403,东洲区,41.866829,124.047219;"
    AddressDistrict = AddressDistrict & "210400,210404,望花区,41.851803,123.801509;"
    AddressDistrict = AddressDistrict & "210400,210411,顺城区,41.881132,123.917165;"
    AddressDistrict = AddressDistrict & "210400,210421,抚顺县,41.922644,124.097979;"
    AddressDistrict = AddressDistrict & "210400,210422,新宾满族自治县,41.732456,125.037547;"
    AddressDistrict = AddressDistrict & "210400,210423,清原满族自治县,42.10135,124.927192;"
    AddressDistrict = AddressDistrict & "210500,210502,平山区,41.291581,123.761231;"
    AddressDistrict = AddressDistrict & "210500,210503,溪湖区,41.330056,123.765226;"
    AddressDistrict = AddressDistrict & "210500,210504,明山区,41.302429,123.763288;"
    AddressDistrict = AddressDistrict & "210500,210505,南芬区,41.104093,123.748381;"
    AddressDistrict = AddressDistrict & "210500,210521,本溪满族自治县,41.300344,124.126156;"
    AddressDistrict = AddressDistrict & "210500,210522,桓仁满族自治县,41.268997,125.359195;"
    AddressDistrict = AddressDistrict & "210600,210602,元宝区,40.136483,124.397814;"
    AddressDistrict = AddressDistrict & "210600,210603,振兴区,40.102801,124.361153;"
    AddressDistrict = AddressDistrict & "210600,210604,振安区,40.158557,124.427709;"
    AddressDistrict = AddressDistrict & "210600,210624,宽甸满族自治县,40.730412,124.784867;"
    AddressDistrict = AddressDistrict & "210600,210681,东港市,39.883467,124.149437;"
    AddressDistrict = AddressDistrict & "210600,210682,凤城市,40.457567,124.071067;"
    AddressDistrict = AddressDistrict & "210700,210702,古塔区,41.115719,121.130085;"
    AddressDistrict = AddressDistrict & "210700,210703,凌河区,41.114662,121.151304;"
    AddressDistrict = AddressDistrict & "210700,210711,太和区,41.105378,121.107297;"
    AddressDistrict = AddressDistrict & "210700,210726,黑山县,41.691804,122.117915;"
    AddressDistrict = AddressDistrict & "210700,210727,义县,41.537224,121.242831;"
    AddressDistrict = AddressDistrict & "210700,210781,凌海市,41.171738,121.364236;"
    AddressDistrict = AddressDistrict & "210700,210782,北镇市,41.598764,121.795962;"
    AddressDistrict = AddressDistrict & "210800,210802,站前区,40.669949,122.253235;"
    AddressDistrict = AddressDistrict & "210800,210803,西市区,40.663086,122.210067;"
    AddressDistrict = AddressDistrict & "210800,210804,鲅鱼圈区,40.263646,122.127242;"
    AddressDistrict = AddressDistrict & "210800,210811,老边区,40.682723,122.382584;"
    AddressDistrict = AddressDistrict & "210800,210881,盖州市,40.405234,122.355534;"
    AddressDistrict = AddressDistrict & "210800,210882,大石桥市,40.633973,122.505894;"
    AddressDistrict = AddressDistrict & "210900,210902,海州区,42.011162,121.657639;"
    AddressDistrict = AddressDistrict & "210900,210903,新邱区,42.086603,121.790541;"
    AddressDistrict = AddressDistrict & "210900,210904,太平区,42.011145,121.677575;"
    AddressDistrict = AddressDistrict & "210900,210905,清河门区,41.780477,121.42018;"
    AddressDistrict = AddressDistrict & "210900,210911,细河区,42.019218,121.654791;"
    AddressDistrict = AddressDistrict & "210900,210921,阜新蒙古族自治县,42.058607,121.743125;"
    AddressDistrict = AddressDistrict & "210900,210922,彰武县,42.384823,122.537444;"
    AddressDistrict = AddressDistrict & "211000,211002,白塔区,41.26745,123.172611;"
    AddressDistrict = AddressDistrict & "211000,211003,文圣区,41.266765,123.188227;"
    AddressDistrict = AddressDistrict & "211000,211004,宏伟区,41.205747,123.200461;"
    AddressDistrict = AddressDistrict & "211000,211005,弓长岭区,41.157831,123.431633;"
    AddressDistrict = AddressDistrict & "211000,211011,太子河区,41.251682,123.185336;"
    AddressDistrict = AddressDistrict & "211000,211021,辽阳县,41.216479,123.079674;"
    AddressDistrict = AddressDistrict & "211000,211081,灯塔市,41.427836,123.325864;"
    AddressDistrict = AddressDistrict & "211100,211102,双台子区,41.190365,122.055733;"
    AddressDistrict = AddressDistrict & "211100,211103,兴隆台区,41.122423,122.071624;"
    AddressDistrict = AddressDistrict & "211100,211104,大洼区,40.994428,122.071708;"
    AddressDistrict = AddressDistrict & "211100,211122,盘山县,41.240701,121.98528;"
    AddressDistrict = AddressDistrict & "211200,211202,银州区,42.292278,123.844877;"
    AddressDistrict = AddressDistrict & "211200,211204,清河区,42.542978,124.14896;"
    AddressDistrict = AddressDistrict & "211200,211221,铁岭县,42.223316,123.725669;"
    AddressDistrict = AddressDistrict & "211200,211223,西丰县,42.738091,124.72332;"
    AddressDistrict = AddressDistrict & "211200,211224,昌图县,42.784441,124.11017;"
    AddressDistrict = AddressDistrict & "211200,211281,调兵山市,42.450734,123.545366;"
    AddressDistrict = AddressDistrict & "211200,211282,开原市,42.542141,124.045551;"
    AddressDistrict = AddressDistrict & "211300,211302,双塔区,41.579389,120.44877;"
    AddressDistrict = AddressDistrict & "211300,211303,龙城区,41.576749,120.413376;"
    AddressDistrict = AddressDistrict & "211300,211321,朝阳县,41.526342,120.404217;"
    AddressDistrict = AddressDistrict & "211300,211322,建平县,41.402576,119.642363;"
    AddressDistrict = AddressDistrict & "211300,211324,喀喇沁左翼蒙古族自治县,41.125428,119.744883;"
    AddressDistrict = AddressDistrict & "211300,211381,北票市,41.803286,120.766951;"
    AddressDistrict = AddressDistrict & "211300,211382,凌源市,41.243086,119.404789;"
    AddressDistrict = AddressDistrict & "211400,211402,连山区,40.755143,120.85937;"
    AddressDistrict = AddressDistrict & "211400,211403,龙港区,40.709991,120.838569;"
    AddressDistrict = AddressDistrict & "211400,211404,南票区,41.098813,120.752314;"
    AddressDistrict = AddressDistrict & "211400,211421,绥中县,40.328407,120.342112;"
    AddressDistrict = AddressDistrict & "211400,211422,建昌县,40.812871,119.807776;"
    AddressDistrict = AddressDistrict & "211400,211481,兴城市,40.619413,120.729365;"
    AddressDistrict = AddressDistrict & "220100,220102,南关区,43.890235,125.337237;"
    AddressDistrict = AddressDistrict & "220100,220103,宽城区,43.903823,125.342828;"
    AddressDistrict = AddressDistrict & "220100,220104,朝阳区,43.86491,125.318042;"
    AddressDistrict = AddressDistrict & "220100,220105,二道区,43.870824,125.384727;"
    AddressDistrict = AddressDistrict & "220100,220106,绿园区,43.892177,125.272467;"
    AddressDistrict = AddressDistrict & "220100,220112,双阳区,43.525168,125.659018;"
    AddressDistrict = AddressDistrict & "220100,220113,九台区,44.157155,125.844682;"
    AddressDistrict = AddressDistrict & "220100,220122,农安县,44.431258,125.175287;"
    AddressDistrict = AddressDistrict & "220100,220182,榆树市,44.827642,126.550107;"
    AddressDistrict = AddressDistrict & "220100,220183,德惠市,44.533909,125.703327;"
    AddressDistrict = AddressDistrict & "220100,220184,公主岭市,43.509474,124.817588;"
    AddressDistrict = AddressDistrict & "220200,220202,昌邑区,43.851118,126.570766;"
    AddressDistrict = AddressDistrict & "220200,220203,龙潭区,43.909755,126.561429;"
    AddressDistrict = AddressDistrict & "220200,220204,船营区,43.843804,126.55239;"
    AddressDistrict = AddressDistrict & "220200,220211,丰满区,43.816594,126.560759;"
    AddressDistrict = AddressDistrict & "220200,220221,永吉县,43.667416,126.501622;"
    AddressDistrict = AddressDistrict & "220200,220281,蛟河市,43.720579,127.342739;"
    AddressDistrict = AddressDistrict & "220200,220282,桦甸市,42.972093,126.745445;"
    AddressDistrict = AddressDistrict & "220200,220283,舒兰市,44.410906,126.947813;"
    AddressDistrict = AddressDistrict & "220200,220284,磐石市,42.942476,126.059929;"
    AddressDistrict = AddressDistrict & "220300,220302,铁西区,43.176263,124.360894;"
    AddressDistrict = AddressDistrict & "220300,220303,铁东区,43.16726,124.388464;"
    AddressDistrict = AddressDistrict & "220300,220322,梨树县,43.30831,124.335802;"
    AddressDistrict = AddressDistrict & "220300,220323,伊通满族自治县,43.345464,125.303124;"
    AddressDistrict = AddressDistrict & "220300,220382,双辽市,43.518275,123.505283;"
    AddressDistrict = AddressDistrict & "220400,220402,龙山区,42.902702,125.145164;"
    AddressDistrict = AddressDistrict & "220400,220403,西安区,42.920415,125.151424;"
    AddressDistrict = AddressDistrict & "220400,220421,东丰县,42.675228,125.529623;"
    AddressDistrict = AddressDistrict & "220400,220422,东辽县,42.927724,124.991995;"
    AddressDistrict = AddressDistrict & "220500,220502,东昌区,41.721233,125.936716;"
    AddressDistrict = AddressDistrict & "220500,220503,二道江区,41.777564,126.045987;"
    AddressDistrict = AddressDistrict & "220500,220521,通化县,41.677918,125.753121;"
    AddressDistrict = AddressDistrict & "220500,220523,辉南县,42.683459,126.042821;"
    AddressDistrict = AddressDistrict & "220500,220524,柳河县,42.281484,125.740536;"
    AddressDistrict = AddressDistrict & "220500,220581,梅河口市,42.530002,125.687336;"
    AddressDistrict = AddressDistrict & "220500,220582,集安市,41.126276,126.186204;"
    AddressDistrict = AddressDistrict & "220600,220602,浑江区,41.943065,126.428035;"
    AddressDistrict = AddressDistrict & "220600,220605,江源区,42.048109,126.584229;"
    AddressDistrict = AddressDistrict & "220600,220621,抚松县,42.332643,127.273796;"
    AddressDistrict = AddressDistrict & "220600,220622,靖宇县,42.389689,126.808386;"
    AddressDistrict = AddressDistrict & "220600,220623,长白朝鲜族自治县,41.419361,128.203384;"
    AddressDistrict = AddressDistrict & "220600,220681,临江市,41.810689,126.919296;"
    AddressDistrict = AddressDistrict & "220700,220702,宁江区,45.176498,124.827851;"
    AddressDistrict = AddressDistrict & "220700,220721,前郭尔罗斯蒙古族自治县,45.116288,124.826808;"
    AddressDistrict = AddressDistrict & "220700,220722,长岭县,44.276579,123.985184;"
    AddressDistrict = AddressDistrict & "220700,220723,乾安县,45.006846,124.024361;"
    AddressDistrict = AddressDistrict & "220700,220781,扶余市,44.986199,126.042758;"
    AddressDistrict = AddressDistrict & "220800,220802,洮北区,45.619253,122.842499;"
    AddressDistrict = AddressDistrict & "220800,220821,镇赉县,45.846089,123.202246;"
    AddressDistrict = AddressDistrict & "220800,220822,通榆县,44.80915,123.088543;"
    AddressDistrict = AddressDistrict & "220800,220881,洮南市,45.339113,122.783779;"
    AddressDistrict = AddressDistrict & "220800,220882,大安市,45.507648,124.291512;"
    AddressDistrict = AddressDistrict & "222400,222401,延吉市,42.906964,129.51579;"
    AddressDistrict = AddressDistrict & "222400,222402,图们市,42.966621,129.846701;"
    AddressDistrict = AddressDistrict & "222400,222403,敦化市,43.366921,128.22986;"
    AddressDistrict = AddressDistrict & "222400,222404,珲春市,42.871057,130.365787;"
    AddressDistrict = AddressDistrict & "222400,222405,龙井市,42.771029,129.425747;"
    AddressDistrict = AddressDistrict & "222400,222406,和龙市,42.547004,129.008748;"
    AddressDistrict = AddressDistrict & "222400,222424,汪清县,43.315426,129.766161;"
    AddressDistrict = AddressDistrict & "222400,222426,安图县,43.110994,128.901865;"
    AddressDistrict = AddressDistrict & "230100,230102,道里区,45.762035,126.612532;"
    AddressDistrict = AddressDistrict & "230100,230103,南岗区,45.755971,126.652098;"
    AddressDistrict = AddressDistrict & "230100,230104,道外区,45.78454,126.648838;"
    AddressDistrict = AddressDistrict & "230100,230108,平房区,45.605567,126.629257;"
    AddressDistrict = AddressDistrict & "230100,230109,松北区,45.814656,126.563066;"
    AddressDistrict = AddressDistrict & "230100,230110,香坊区,45.713067,126.667049;"
    AddressDistrict = AddressDistrict & "230100,230111,呼兰区,45.98423,126.603302;"
    AddressDistrict = AddressDistrict & "230100,230112,阿城区,45.538372,126.972726;"
    AddressDistrict = AddressDistrict & "230100,230113,双城区,45.377942,126.308784;"
    AddressDistrict = AddressDistrict & "230100,230123,依兰县,46.315105,129.565594;"
    AddressDistrict = AddressDistrict & "230100,230124,方正县,45.839536,128.836131;"
    AddressDistrict = AddressDistrict & "230100,230125,宾县,45.759369,127.48594;"
    AddressDistrict = AddressDistrict & "230100,230126,巴彦县,46.081889,127.403602;"
    AddressDistrict = AddressDistrict & "230100,230127,木兰县,45.949826,128.042675;"
    AddressDistrict = AddressDistrict & "230100,230128,通河县,45.977618,128.747786;"
    AddressDistrict = AddressDistrict & "230100,230129,延寿县,45.455648,128.331886;"
    AddressDistrict = AddressDistrict & "230100,230183,尚志市,45.214953,127.968539;"
    AddressDistrict = AddressDistrict & "230100,230184,五常市,44.919418,127.15759;"
    AddressDistrict = AddressDistrict & "230200,230202,龙沙区,47.341736,123.957338;"
    AddressDistrict = AddressDistrict & "230200,230203,建华区,47.354494,123.955888;"
    AddressDistrict = AddressDistrict & "230200,230204,铁锋区,47.339499,123.973555;"
    AddressDistrict = AddressDistrict & "230200,230205,昂昂溪区,47.156867,123.813181;"
    AddressDistrict = AddressDistrict & "230200,230206,富拉尔基区,47.20697,123.638873;"
    AddressDistrict = AddressDistrict & "230200,230207,碾子山区,47.51401,122.887972;"
    AddressDistrict = AddressDistrict & "230200,230208,梅里斯达斡尔族区,47.311113,123.754599;"
    AddressDistrict = AddressDistrict & "230200,230221,龙江县,47.336388,123.187225;"
    AddressDistrict = AddressDistrict & "230200,230223,依安县,47.890098,125.307561;"
    AddressDistrict = AddressDistrict & "230200,230224,泰来县,46.39233,123.41953;"
    AddressDistrict = AddressDistrict & "230200,230225,甘南县,47.917838,123.506034;"
    AddressDistrict = AddressDistrict & "230200,230227,富裕县,47.797172,124.469106;"
    AddressDistrict = AddressDistrict & "230200,230229,克山县,48.034342,125.874355;"
    AddressDistrict = AddressDistrict & "230200,230230,克东县,48.03732,126.249094;"
    AddressDistrict = AddressDistrict & "230200,230231,拜泉县,47.607363,126.091911;"
    AddressDistrict = AddressDistrict & "230200,230281,讷河市,48.481133,124.882172;"
    AddressDistrict = AddressDistrict & "230300,230302,鸡冠区,45.30034,130.974374;"
    AddressDistrict = AddressDistrict & "230300,230303,恒山区,45.213242,130.910636;"
    AddressDistrict = AddressDistrict & "230300,230304,滴道区,45.348812,130.846823;"
    AddressDistrict = AddressDistrict & "230300,230305,梨树区,45.092195,130.697781;"
    AddressDistrict = AddressDistrict & "230300,230306,城子河区,45.338248,131.010501;"
    AddressDistrict = AddressDistrict & "230300,230307,麻山区,45.209607,130.481126;"
    AddressDistrict = AddressDistrict & "230300,230321,鸡东县,45.250892,131.148907;"
    AddressDistrict = AddressDistrict & "230300,230381,虎林市,45.767985,132.973881;"
    AddressDistrict = AddressDistrict & "230300,230382,密山市,45.54725,131.874137;"
    AddressDistrict = AddressDistrict & "230400,230402,向阳区,47.345372,130.292478;"
    AddressDistrict = AddressDistrict & "230400,230403,工农区,47.331678,130.276652;"
    AddressDistrict = AddressDistrict & "230400,230404,南山区,47.31324,130.275533;"
    AddressDistrict = AddressDistrict & "230400,230405,兴安区,47.252911,130.236169;"
    AddressDistrict = AddressDistrict & "230400,230406,东山区,47.337385,130.31714;"
    AddressDistrict = AddressDistrict & "230400,230407,兴山区,47.35997,130.30534;"
    AddressDistrict = AddressDistrict & "230400,230421,萝北县,47.577577,130.829087;"
    AddressDistrict = AddressDistrict & "230400,230422,绥滨县,47.289892,131.860526;"
    AddressDistrict = AddressDistrict & "230500,230502,尖山区,46.642961,131.15896;"
    AddressDistrict = AddressDistrict & "230500,230503,岭东区,46.591076,131.163675;"
    AddressDistrict = AddressDistrict & "230500,230505,四方台区,46.594347,131.333181;"
    AddressDistrict = AddressDistrict & "230500,230506,宝山区,46.573366,131.404294;"
    AddressDistrict = AddressDistrict & "230500,230521,集贤县,46.72898,131.13933;"
    AddressDistrict = AddressDistrict & "230500,230522,友谊县,46.775159,131.810622;"
    AddressDistrict = AddressDistrict & "230500,230523,宝清县,46.328781,132.206415;"
    AddressDistrict = AddressDistrict & "230500,230524,饶河县,46.801288,134.021162;"
    AddressDistrict = AddressDistrict & "230600,230602,萨尔图区,46.596356,125.114643;"
    AddressDistrict = AddressDistrict & "230600,230603,龙凤区,46.573948,125.145794;"
    AddressDistrict = AddressDistrict & "230600,230604,让胡路区,46.653254,124.868341;"
    AddressDistrict = AddressDistrict & "230600,230605,红岗区,46.403049,124.889528;"
    AddressDistrict = AddressDistrict & "230600,230606,大同区,46.034304,124.818509;"
    AddressDistrict = AddressDistrict & "230600,230621,肇州县,45.708685,125.273254;"
    AddressDistrict = AddressDistrict & "230600,230622,肇源县,45.518832,125.081974;"
    AddressDistrict = AddressDistrict & "230600,230623,林甸县,47.186411,124.877742;"
    AddressDistrict = AddressDistrict & "230600,230624,杜尔伯特蒙古族自治县,46.865973,124.446259;"
    AddressDistrict = AddressDistrict & "230700,230717,伊美区,47.728171,128.907303;"
    AddressDistrict = AddressDistrict & "230700,230718,乌翠区,47.726728,128.669859;"
    AddressDistrict = AddressDistrict & "230700,230719,友好区,47.853778,128.84075;"
    AddressDistrict = AddressDistrict & "230700,230722,嘉荫县,48.891378,130.397684;"
    AddressDistrict = AddressDistrict & "230700,230723,汤旺县,48.454651,129.571108;"
    AddressDistrict = AddressDistrict & "230700,230724,丰林县,48.290455,129.5336;"
    AddressDistrict = AddressDistrict & "230700,230725,大箐山县,47.028397,129.020793;"
    AddressDistrict = AddressDistrict & "230700,230726,南岔县,47.137314,129.28246;"
    AddressDistrict = AddressDistrict & "230700,230751,金林区,47.413074,129.429117;"
    AddressDistrict = AddressDistrict & "230700,230781,铁力市,46.985772,128.030561;"
    AddressDistrict = AddressDistrict & "230800,230803,向阳区,46.809645,130.361786;"
    AddressDistrict = AddressDistrict & "230800,230804,前进区,46.812345,130.377684;"
    AddressDistrict = AddressDistrict & "230800,230805,东风区,46.822476,130.403297;"
    AddressDistrict = AddressDistrict & "230800,230811,郊区,46.80712,130.351588;"
    AddressDistrict = AddressDistrict & "230800,230822,桦南县,46.240118,130.570112;"
    AddressDistrict = AddressDistrict & "230800,230826,桦川县,47.023039,130.723713;"
    AddressDistrict = AddressDistrict & "230800,230828,汤原县,46.730048,129.904463;"
    AddressDistrict = AddressDistrict & "230800,230881,同江市,47.651131,132.510119;"
    AddressDistrict = AddressDistrict & "230800,230882,富锦市,47.250747,132.037951;"
    AddressDistrict = AddressDistrict & "230800,230883,抚远市,48.364707,134.294501;"
    AddressDistrict = AddressDistrict & "230900,230902,新兴区,45.794258,130.889482;"
    AddressDistrict = AddressDistrict & "230900,230903,桃山区,45.771217,131.015848;"
    AddressDistrict = AddressDistrict & "230900,230904,茄子河区,45.776587,131.071561;"
    AddressDistrict = AddressDistrict & "230900,230921,勃利县,45.751573,130.575025;"
    AddressDistrict = AddressDistrict & "231000,231002,东安区,44.582399,129.623292;"
    AddressDistrict = AddressDistrict & "231000,231003,阳明区,44.596328,129.634645;"
    AddressDistrict = AddressDistrict & "231000,231004,爱民区,44.595443,129.601232;"
    AddressDistrict = AddressDistrict & "231000,231005,西安区,44.581032,129.61311;"
    AddressDistrict = AddressDistrict & "231000,231025,林口县,45.286645,130.268402;"
    AddressDistrict = AddressDistrict & "231000,231081,绥芬河市,44.396864,131.164856;"
    AddressDistrict = AddressDistrict & "231000,231083,海林市,44.574149,129.387902;"
    AddressDistrict = AddressDistrict & "231000,231084,宁安市,44.346836,129.470019;"
    AddressDistrict = AddressDistrict & "231000,231085,穆棱市,44.91967,130.527085;"
    AddressDistrict = AddressDistrict & "231000,231086,东宁市,44.063578,131.125296;"
    AddressDistrict = AddressDistrict & "231100,231102,爱辉区,50.249027,127.497639;"
    AddressDistrict = AddressDistrict & "231100,231123,逊克县,49.582974,128.476152;"
    AddressDistrict = AddressDistrict & "231100,231124,孙吴县,49.423941,127.327315;"
    AddressDistrict = AddressDistrict & "231100,231181,北安市,48.245437,126.508737;"
    AddressDistrict = AddressDistrict & "231100,231182,五大连池市,48.512688,126.197694;"
    AddressDistrict = AddressDistrict & "231100,231183,嫩江市,49.177461,125.229904;"
    AddressDistrict = AddressDistrict & "231200,231202,北林区,46.634912,126.990665;"
    AddressDistrict = AddressDistrict & "231200,231221,望奎县,46.83352,126.484191;"
    AddressDistrict = AddressDistrict & "231200,231222,兰西县,46.259037,126.289315;"
    AddressDistrict = AddressDistrict & "231200,231223,青冈县,46.686596,126.112268;"
    AddressDistrict = AddressDistrict & "231200,231224,庆安县,46.879203,127.510024;"
    AddressDistrict = AddressDistrict & "231200,231225,明水县,47.183527,125.907544;"
    AddressDistrict = AddressDistrict & "231200,231226,绥棱县,47.247195,127.111121;"
    AddressDistrict = AddressDistrict & "231200,231281,安达市,46.410614,125.329926;"
    AddressDistrict = AddressDistrict & "231200,231282,肇东市,46.069471,125.991402;"
    AddressDistrict = AddressDistrict & "231200,231283,海伦市,47.460428,126.969383;"
    AddressDistrict = AddressDistrict & "232700,232701,漠河市,52.972074,122.536256;"
    AddressDistrict = AddressDistrict & "232700,232718,加格达奇区,50.424654,124.126716;"
    AddressDistrict = AddressDistrict & "232700,232721,呼玛县,51.726998,126.662105;"
    AddressDistrict = AddressDistrict & "232700,232722,塔河县,52.335229,124.710516;"
    AddressDistrict = AddressDistrict & "320100,320102,玄武区,32.050678,118.792199;"
    AddressDistrict = AddressDistrict & "320100,320104,秦淮区,32.033818,118.786088;"
    AddressDistrict = AddressDistrict & "320100,320105,建邺区,32.004538,118.732688;"
    AddressDistrict = AddressDistrict & "320100,320106,鼓楼区,32.066966,118.769739;"
    AddressDistrict = AddressDistrict & "320100,320111,浦口区,32.05839,118.625307;"
    AddressDistrict = AddressDistrict & "320100,320113,栖霞区,32.102147,118.808702;"
    AddressDistrict = AddressDistrict & "320100,320114,雨花台区,31.995946,118.77207;"
    AddressDistrict = AddressDistrict & "320100,320115,江宁区,31.953418,118.850621;"
    AddressDistrict = AddressDistrict & "320100,320116,六合区,32.340655,118.85065;"
    AddressDistrict = AddressDistrict & "320100,320117,溧水区,31.653061,119.028732;"
    AddressDistrict = AddressDistrict & "320100,320118,高淳区,31.327132,118.87589;"
    AddressDistrict = AddressDistrict & "320200,320205,锡山区,31.585559,120.357298;"
    AddressDistrict = AddressDistrict & "320200,320206,惠山区,31.681019,120.303543;"
    AddressDistrict = AddressDistrict & "320200,320211,滨湖区,31.550228,120.266053;"
    AddressDistrict = AddressDistrict & "320200,320213,梁溪区,31.575706,120.296595;"
    AddressDistrict = AddressDistrict & "320200,320214,新吴区,31.550966,120.352782;"
    AddressDistrict = AddressDistrict & "320200,320281,江阴市,31.910984,120.275891;"
    AddressDistrict = AddressDistrict & "320200,320282,宜兴市,31.364384,119.820538;"
    AddressDistrict = AddressDistrict & "320300,320302,鼓楼区,34.269397,117.192941;"
    AddressDistrict = AddressDistrict & "320300,320303,云龙区,34.254805,117.194589;"
    AddressDistrict = AddressDistrict & "320300,320305,贾汪区,34.441642,117.450212;"
    AddressDistrict = AddressDistrict & "320300,320311,泉山区,34.262249,117.182225;"
    AddressDistrict = AddressDistrict & "320300,320312,铜山区,34.19288,117.183894;"
    AddressDistrict = AddressDistrict & "320300,320321,丰县,34.696946,116.592888;"
    AddressDistrict = AddressDistrict & "320300,320322,沛县,34.729044,116.937182;"
    AddressDistrict = AddressDistrict & "320300,320324,睢宁县,33.899222,117.95066;"
    AddressDistrict = AddressDistrict & "320300,320381,新沂市,34.368779,118.345828;"
    AddressDistrict = AddressDistrict & "320300,320382,邳州市,34.314708,117.963923;"
    AddressDistrict = AddressDistrict & "320400,320402,天宁区,31.779632,119.963783;"
    AddressDistrict = AddressDistrict & "320400,320404,钟楼区,31.78096,119.948388;"
    AddressDistrict = AddressDistrict & "320400,320411,新北区,31.824664,119.974654;"
    AddressDistrict = AddressDistrict & "320400,320412,武进区,31.718566,119.958773;"
    AddressDistrict = AddressDistrict & "320400,320413,金坛区,31.744399,119.573395;"
    AddressDistrict = AddressDistrict & "320400,320481,溧阳市,31.427081,119.487816;"
    AddressDistrict = AddressDistrict & "320500,320505,虎丘区,31.294845,120.566833;"
    AddressDistrict = AddressDistrict & "320500,320506,吴中区,31.270839,120.624621;"
    AddressDistrict = AddressDistrict & "320500,320507,相城区,31.396684,120.618956;"
    AddressDistrict = AddressDistrict & "320500,320508,姑苏区,31.311414,120.622249;"
    AddressDistrict = AddressDistrict & "320500,320509,吴江区,31.160404,120.641601;"
    AddressDistrict = AddressDistrict & "320500,320581,常熟市,31.658156,120.74852;"
    AddressDistrict = AddressDistrict & "320500,320582,张家港市,31.865553,120.543441;"
    AddressDistrict = AddressDistrict & "320500,320583,昆山市,31.381925,120.958137;"
    AddressDistrict = AddressDistrict & "320500,320585,太仓市,31.452568,121.112275;"
    AddressDistrict = AddressDistrict & "320600,320602,崇川区,32.015278,120.86635;"
    AddressDistrict = AddressDistrict & "320600,320612,通州区,32.084287,121.073171;"
    AddressDistrict = AddressDistrict & "320600,320623,如东县,32.311832,121.186088;"
    AddressDistrict = AddressDistrict & "320600,320681,启东市,31.810158,121.659724;"
    AddressDistrict = AddressDistrict & "320600,320682,如皋市,32.391591,120.566324;"
    AddressDistrict = AddressDistrict & "320600,320684,海门区,31.893528,121.176609;"
    AddressDistrict = AddressDistrict & "320600,320685,海安市,32.540288,120.465995;"
    AddressDistrict = AddressDistrict & "320700,320703,连云区,34.739529,119.366487;"
    AddressDistrict = AddressDistrict & "320700,320706,海州区,34.601584,119.179793;"
    AddressDistrict = AddressDistrict & "320700,320707,赣榆区,34.839154,119.128774;"
    AddressDistrict = AddressDistrict & "320700,320722,东海县,34.522859,118.766489;"
    AddressDistrict = AddressDistrict & "320700,320723,灌云县,34.298436,119.255741;"
    AddressDistrict = AddressDistrict & "320700,320724,灌南县,34.092553,119.352331;"
    AddressDistrict = AddressDistrict & "320800,320803,淮安区,33.507499,119.14634;"
    AddressDistrict = AddressDistrict & "320800,320804,淮阴区,33.622452,119.020817;"
    AddressDistrict = AddressDistrict & "320800,320812,清江浦区,33.603234,119.019454;"
    AddressDistrict = AddressDistrict & "320800,320813,洪泽区,33.294975,118.867875;"
    AddressDistrict = AddressDistrict & "320800,320826,涟水县,33.771308,119.266078;"
    AddressDistrict = AddressDistrict & "320800,320830,盱眙县,33.00439,118.493823;"
    AddressDistrict = AddressDistrict & "320800,320831,金湖县,33.018162,119.016936;"
    AddressDistrict = AddressDistrict & "320900,320902,亭湖区,33.383912,120.136078;"
    AddressDistrict = AddressDistrict & "320900,320903,盐都区,33.341288,120.139753;"
    AddressDistrict = AddressDistrict & "320900,320904,大丰区,33.199531,120.470324;"
    AddressDistrict = AddressDistrict & "320900,320921,响水县,34.19996,119.579573;"
    AddressDistrict = AddressDistrict & "320900,320922,滨海县,33.989888,119.828434;"
    AddressDistrict = AddressDistrict & "320900,320923,阜宁县,33.78573,119.805338;"
    AddressDistrict = AddressDistrict & "320900,320924,射阳县,33.773779,120.257444;"
    AddressDistrict = AddressDistrict & "320900,320925,建湖县,33.472621,119.793105;"
    AddressDistrict = AddressDistrict & "320900,320981,东台市,32.853174,120.314101;"
    AddressDistrict = AddressDistrict & "321000,321002,广陵区,32.392154,119.442267;"
    AddressDistrict = AddressDistrict & "321000,321003,邗江区,32.377899,119.397777;"
    AddressDistrict = AddressDistrict & "321000,321012,江都区,32.426564,119.567481;"
    AddressDistrict = AddressDistrict & "321000,321023,宝应县,33.23694,119.321284;"
    AddressDistrict = AddressDistrict & "321000,321081,仪征市,32.271965,119.182443;"
    AddressDistrict = AddressDistrict & "321000,321084,高邮市,32.785164,119.443842;"
    AddressDistrict = AddressDistrict & "321100,321102,京口区,32.206191,119.454571;"
    AddressDistrict = AddressDistrict & "321100,321111,润州区,32.213501,119.414877;"
    AddressDistrict = AddressDistrict & "321100,321112,丹徒区,32.128972,119.433883;"
    AddressDistrict = AddressDistrict & "321100,321181,丹阳市,31.991459,119.581911;"
    AddressDistrict = AddressDistrict & "321100,321182,扬中市,32.237266,119.828054;"
    AddressDistrict = AddressDistrict & "321100,321183,句容市,31.947355,119.167135;"
    AddressDistrict = AddressDistrict & "321200,321202,海陵区,32.488406,119.920187;"
    AddressDistrict = AddressDistrict & "321200,321203,高港区,32.315701,119.88166;"
    AddressDistrict = AddressDistrict & "321200,321204,姜堰区,32.508483,120.148208;"
    AddressDistrict = AddressDistrict & "321200,321281,兴化市,32.938065,119.840162;"
    AddressDistrict = AddressDistrict & "321200,321282,靖江市,32.018168,120.26825;"
    AddressDistrict = AddressDistrict & "321200,321283,泰兴市,32.168784,120.020228;"
    AddressDistrict = AddressDistrict & "321300,321302,宿城区,33.937726,118.278984;"
    AddressDistrict = AddressDistrict & "321300,321311,宿豫区,33.941071,118.330012;"
    AddressDistrict = AddressDistrict & "321300,321322,沭阳县,34.129097,118.775889;"
    AddressDistrict = AddressDistrict & "321300,321323,泗阳县,33.711433,118.681284;"
    AddressDistrict = AddressDistrict & "321300,321324,泗洪县,33.456538,118.211824;"
    AddressDistrict = AddressDistrict & "330100,330102,上城区,30.250236,120.171465;"
    AddressDistrict = AddressDistrict & "330100,330105,拱墅区,30.314697,120.150053;"
    AddressDistrict = AddressDistrict & "330100,330106,西湖区,30.272934,120.147376;"
    AddressDistrict = AddressDistrict & "330100,330108,滨江区,30.206615,120.21062;"
    AddressDistrict = AddressDistrict & "330100,330109,萧山区,30.162932,120.27069;"
    AddressDistrict = AddressDistrict & "330100,330110,余杭区,30.27365,119.978959;"
    AddressDistrict = AddressDistrict & "330100,330111,富阳区,30.049871,119.949869;"
    AddressDistrict = AddressDistrict & "330100,330112,临安区,30.231153,119.715101;"
    AddressDistrict = AddressDistrict & "330100,330114,钱塘区,30.322904,120.493972;"
    AddressDistrict = AddressDistrict & "330100,330113,临平区,30.419025,120.299376;"
    AddressDistrict = AddressDistrict & "330100,330122,桐庐县,29.797437,119.685045;"
    AddressDistrict = AddressDistrict & "330100,330127,淳安县,29.604177,119.044276;"
    AddressDistrict = AddressDistrict & "330100,330182,建德市,29.472284,119.279089;"
    AddressDistrict = AddressDistrict & "330200,330203,海曙区,29.874452,121.539698;"
    AddressDistrict = AddressDistrict & "330200,330205,江北区,29.888361,121.559282;"
    AddressDistrict = AddressDistrict & "330200,330206,北仑区,29.90944,121.831303;"
    AddressDistrict = AddressDistrict & "330200,330211,镇海区,29.952107,121.713162;"
    AddressDistrict = AddressDistrict & "330200,330212,鄞州区,29.831662,121.558436;"
    AddressDistrict = AddressDistrict & "330200,330213,奉化区,29.662348,121.41089;"
    AddressDistrict = AddressDistrict & "330200,330225,象山县,29.470206,121.877091;"
    AddressDistrict = AddressDistrict & "330200,330226,宁海县,29.299836,121.432606;"
    AddressDistrict = AddressDistrict & "330200,330281,余姚市,30.045404,121.156294;"
    AddressDistrict = AddressDistrict & "330200,330282,慈溪市,30.177142,121.248052;"
    AddressDistrict = AddressDistrict & "330300,330302,鹿城区,28.003352,120.674231;"
    AddressDistrict = AddressDistrict & "330300,330303,龙湾区,27.970254,120.763469;"
    AddressDistrict = AddressDistrict & "330300,330304,瓯海区,28.006444,120.637145;"
    AddressDistrict = AddressDistrict & "330300,330305,洞头区,27.836057,121.156181;"
    AddressDistrict = AddressDistrict & "330300,330324,永嘉县,28.153886,120.690968;"
    AddressDistrict = AddressDistrict & "330300,330326,平阳县,27.6693,120.564387;"
    AddressDistrict = AddressDistrict & "330300,330327,苍南县,27.507743,120.406256;"
    AddressDistrict = AddressDistrict & "330300,330328,文成县,27.789133,120.09245;"
    AddressDistrict = AddressDistrict & "330300,330329,泰顺县,27.557309,119.71624;"
    AddressDistrict = AddressDistrict & "330300,330381,瑞安市,27.779321,120.646171;"
    AddressDistrict = AddressDistrict & "330300,330382,乐清市,28.116083,120.967147;"
    AddressDistrict = AddressDistrict & "330300,330383,龙港市,27.578156,120.553039;"
    AddressDistrict = AddressDistrict & "330400,330402,南湖区,30.764652,120.749953;"
    AddressDistrict = AddressDistrict & "330400,330411,秀洲区,30.763323,120.720431;"
    AddressDistrict = AddressDistrict & "330400,330421,嘉善县,30.841352,120.921871;"
    AddressDistrict = AddressDistrict & "330400,330424,海盐县,30.522223,120.942017;"
    AddressDistrict = AddressDistrict & "330400,330481,海宁市,30.525544,120.688821;"
    AddressDistrict = AddressDistrict & "330400,330482,平湖市,30.698921,121.014666;"
    AddressDistrict = AddressDistrict & "330400,330483,桐乡市,30.629065,120.551085;"
    AddressDistrict = AddressDistrict & "330500,330502,吴兴区,30.867252,120.101416;"
    AddressDistrict = AddressDistrict & "330500,330503,南浔区,30.872742,120.417195;"
    AddressDistrict = AddressDistrict & "330500,330521,德清县,30.534927,119.967662;"
    AddressDistrict = AddressDistrict & "330500,330522,长兴县,31.00475,119.910122;"
    AddressDistrict = AddressDistrict & "330500,330523,安吉县,30.631974,119.687891;"
    AddressDistrict = AddressDistrict & "330600,330602,越城区,29.996993,120.585315;"
    AddressDistrict = AddressDistrict & "330600,330603,柯桥区,30.078038,120.476075;"
    AddressDistrict = AddressDistrict & "330600,330604,上虞区,30.016769,120.874185;"
    AddressDistrict = AddressDistrict & "330600,330624,新昌县,29.501205,120.905665;"
    AddressDistrict = AddressDistrict & "330600,330681,诸暨市,29.713662,120.244326;"
    AddressDistrict = AddressDistrict & "330600,330683,嵊州市,29.586606,120.82888;"
    AddressDistrict = AddressDistrict & "330700,330702,婺城区,29.082607,119.652579;"
    AddressDistrict = AddressDistrict & "330700,330703,金东区,29.095835,119.681264;"
    AddressDistrict = AddressDistrict & "330700,330723,武义县,28.896563,119.819159;"
    AddressDistrict = AddressDistrict & "330700,330726,浦江县,29.451254,119.893363;"
    AddressDistrict = AddressDistrict & "330700,330727,磐安县,29.052627,120.44513;"
    AddressDistrict = AddressDistrict & "330700,330781,兰溪市,29.210065,119.460521;"
    AddressDistrict = AddressDistrict & "330700,330782,义乌市,29.306863,120.074911;"
    AddressDistrict = AddressDistrict & "330700,330783,东阳市,29.262546,120.23334;"
    AddressDistrict = AddressDistrict & "330700,330784,永康市,28.895293,120.036328;"
    AddressDistrict = AddressDistrict & "330800,330802,柯城区,28.944539,118.873041;"
    AddressDistrict = AddressDistrict & "330800,330803,衢江区,28.973195,118.957683;"
    AddressDistrict = AddressDistrict & "330800,330822,常山县,28.900039,118.521654;"
    AddressDistrict = AddressDistrict & "330800,330824,开化县,29.136503,118.414435;"
    AddressDistrict = AddressDistrict & "330800,330825,龙游县,29.031364,119.172525;"
    AddressDistrict = AddressDistrict & "330800,330881,江山市,28.734674,118.627879;"
    AddressDistrict = AddressDistrict & "330900,330902,定海区,30.016423,122.108496;"
    AddressDistrict = AddressDistrict & "330900,330903,普陀区,29.945614,122.301953;"
    AddressDistrict = AddressDistrict & "330900,330921,岱山县,30.242865,122.201132;"
    AddressDistrict = AddressDistrict & "330900,330922,嵊泗县,30.727166,122.457809;"
    AddressDistrict = AddressDistrict & "331000,331002,椒江区,28.67615,121.431049;"
    AddressDistrict = AddressDistrict & "331000,331003,黄岩区,28.64488,121.262138;"
    AddressDistrict = AddressDistrict & "331000,331004,路桥区,28.581799,121.37292;"
    AddressDistrict = AddressDistrict & "331000,331022,三门县,29.118955,121.376429;"
    AddressDistrict = AddressDistrict & "331000,331023,天台县,29.141126,121.031227;"
    AddressDistrict = AddressDistrict & "331000,331024,仙居县,28.849213,120.735074;"
    AddressDistrict = AddressDistrict & "331000,331081,温岭市,28.368781,121.373611;"
    AddressDistrict = AddressDistrict & "331000,331082,临海市,28.845441,121.131229;"
    AddressDistrict = AddressDistrict & "331000,331083,玉环市,28.12842,121.232337;"
    AddressDistrict = AddressDistrict & "331100,331102,莲都区,28.451103,119.922293;"
    AddressDistrict = AddressDistrict & "331100,331121,青田县,28.135247,120.291939;"
    AddressDistrict = AddressDistrict & "331100,331122,缙云县,28.654208,120.078965;"
    AddressDistrict = AddressDistrict & "331100,331123,遂昌县,28.5924,119.27589;"
    AddressDistrict = AddressDistrict & "331100,331124,松阳县,28.449937,119.485292;"
    AddressDistrict = AddressDistrict & "331100,331125,云和县,28.111077,119.569458;"
    AddressDistrict = AddressDistrict & "331100,331126,庆元县,27.618231,119.067233;"
    AddressDistrict = AddressDistrict & "331100,331127,景宁畲族自治县,27.977247,119.634669;"
    AddressDistrict = AddressDistrict & "331100,331181,龙泉市,28.069177,119.132319;"
    AddressDistrict = AddressDistrict & "340100,340102,瑶海区,31.86961,117.315358;"
    AddressDistrict = AddressDistrict & "340100,340103,庐阳区,31.869011,117.283776;"
    AddressDistrict = AddressDistrict & "340100,340104,蜀山区,31.855868,117.262072;"
    AddressDistrict = AddressDistrict & "340100,340111,包河区,31.82956,117.285751;"
    AddressDistrict = AddressDistrict & "340100,340121,长丰县,32.478548,117.164699;"
    AddressDistrict = AddressDistrict & "340100,340122,肥东县,31.883992,117.463222;"
    AddressDistrict = AddressDistrict & "340100,340123,肥西县,31.719646,117.166118;"
    AddressDistrict = AddressDistrict & "340100,340124,庐江县,31.251488,117.289844;"
    AddressDistrict = AddressDistrict & "340100,340181,巢湖市,31.600518,117.874155;"
    AddressDistrict = AddressDistrict & "340200,340202,镜湖区,31.32559,118.376343;"
    AddressDistrict = AddressDistrict & "340200,340207,鸠江区,31.362716,118.400174;"
    AddressDistrict = AddressDistrict & "340200,340209,弋江区,31.313394,118.377476;"
    AddressDistrict = AddressDistrict & "340200,340210,湾沚区,31.145262,118.572301;"
    AddressDistrict = AddressDistrict & "340200,340211,繁昌区,31.080896,118.201349;"
    AddressDistrict = AddressDistrict & "340200,340223,南陵县,30.919638,118.337104;"
    AddressDistrict = AddressDistrict & "340200,340281,无为市,31.303075,117.911432;"
    AddressDistrict = AddressDistrict & "340300,340302,龙子湖区,32.950452,117.382312;"
    AddressDistrict = AddressDistrict & "340300,340303,蚌山区,32.938066,117.355789;"
    AddressDistrict = AddressDistrict & "340300,340304,禹会区,32.931933,117.35259;"
    AddressDistrict = AddressDistrict & "340300,340311,淮上区,32.963147,117.34709;"
    AddressDistrict = AddressDistrict & "340300,340321,怀远县,32.956934,117.200171;"
    AddressDistrict = AddressDistrict & "340300,340322,五河县,33.146202,117.888809;"
    AddressDistrict = AddressDistrict & "340300,340323,固镇县,33.318679,117.315962;"
    AddressDistrict = AddressDistrict & "340400,340402,大通区,32.632066,117.052927;"
    AddressDistrict = AddressDistrict & "340400,340403,田家庵区,32.644342,117.018318;"
    AddressDistrict = AddressDistrict & "340400,340404,谢家集区,32.598289,116.865354;"
    AddressDistrict = AddressDistrict & "340400,340405,八公山区,32.628229,116.841111;"
    AddressDistrict = AddressDistrict & "340400,340406,潘集区,32.782117,116.816879;"
    AddressDistrict = AddressDistrict & "340400,340421,凤台县,32.705382,116.722769;"
    AddressDistrict = AddressDistrict & "340400,340422,寿县,32.577304,116.785349;"
    AddressDistrict = AddressDistrict & "340500,340503,花山区,31.69902,118.511308;"
    AddressDistrict = AddressDistrict & "340500,340504,雨山区,31.685912,118.493104;"
    AddressDistrict = AddressDistrict & "340500,340506,博望区,31.562321,118.843742;"
    AddressDistrict = AddressDistrict & "340500,340521,当涂县,31.556167,118.489873;"
    AddressDistrict = AddressDistrict & "340500,340522,含山县,31.727758,118.105545;"
    AddressDistrict = AddressDistrict & "340500,340523,和县,31.716634,118.362998;"
    AddressDistrict = AddressDistrict & "340600,340602,杜集区,33.991218,116.833925;"
    AddressDistrict = AddressDistrict & "340600,340603,相山区,33.970916,116.790775;"
    AddressDistrict = AddressDistrict & "340600,340604,烈山区,33.889529,116.809465;"
    AddressDistrict = AddressDistrict & "340600,340621,濉溪县,33.916407,116.767435;"
    AddressDistrict = AddressDistrict & "340700,340705,铜官区,30.927613,117.816167;"
    AddressDistrict = AddressDistrict & "340700,340706,义安区,30.952338,117.792288;"
    AddressDistrict = AddressDistrict & "340700,340711,郊区,30.908927,117.80707;"
    AddressDistrict = AddressDistrict & "340700,340722,枞阳县,30.700615,117.222027;"
    AddressDistrict = AddressDistrict & "340800,340802,迎江区,30.506375,117.044965;"
    AddressDistrict = AddressDistrict & "340800,340803,大观区,30.505632,117.034512;"
    AddressDistrict = AddressDistrict & "340800,340811,宜秀区,30.541323,117.070003;"
    AddressDistrict = AddressDistrict & "340800,340822,怀宁县,30.734994,116.828664;"
    AddressDistrict = AddressDistrict & "340800,340825,太湖县,30.451869,116.305225;"
    AddressDistrict = AddressDistrict & "340800,340826,宿松县,30.158327,116.120204;"
    AddressDistrict = AddressDistrict & "340800,340827,望江县,30.12491,116.690927;"
    AddressDistrict = AddressDistrict & "340800,340828,岳西县,30.848502,116.360482;"
    AddressDistrict = AddressDistrict & "340800,340881,桐城市,31.050576,116.959656;"
    AddressDistrict = AddressDistrict & "340800,340882,潜山市,30.638222,116.573665;"
    AddressDistrict = AddressDistrict & "341000,341002,屯溪区,29.709186,118.317354;"
    AddressDistrict = AddressDistrict & "341000,341003,黄山区,30.294517,118.136639;"
    AddressDistrict = AddressDistrict & "341000,341004,徽州区,29.825201,118.339743;"
    AddressDistrict = AddressDistrict & "341000,341021,歙县,29.867748,118.428025;"
    AddressDistrict = AddressDistrict & "341000,341022,休宁县,29.788878,118.188531;"
    AddressDistrict = AddressDistrict & "341000,341023,黟县,29.923812,117.942911;"
    AddressDistrict = AddressDistrict & "341000,341024,祁门县,29.853472,117.717237;"
    AddressDistrict = AddressDistrict & "341100,341102,琅琊区,32.303797,118.316475;"
    AddressDistrict = AddressDistrict & "341100,341103,南谯区,32.329841,118.296955;"
    AddressDistrict = AddressDistrict & "341100,341122,来安县,32.450231,118.433293;"
    AddressDistrict = AddressDistrict & "341100,341124,全椒县,32.09385,118.268576;"
    AddressDistrict = AddressDistrict & "341100,341125,定远县,32.527105,117.683713;"
    AddressDistrict = AddressDistrict & "341100,341126,凤阳县,32.867146,117.562461;"
    AddressDistrict = AddressDistrict & "341100,341181,天长市,32.6815,119.011212;"
    AddressDistrict = AddressDistrict & "341100,341182,明光市,32.781206,117.998048;"
    AddressDistrict = AddressDistrict & "341200,341202,颍州区,32.891238,115.813914;"
    AddressDistrict = AddressDistrict & "341200,341203,颍东区,32.908861,115.858747;"
    AddressDistrict = AddressDistrict & "341200,341204,颍泉区,32.924797,115.804525;"
    AddressDistrict = AddressDistrict & "341200,341221,临泉县,33.062698,115.261688;"
    AddressDistrict = AddressDistrict & "341200,341222,太和县,33.16229,115.627243;"
    AddressDistrict = AddressDistrict & "341200,341225,阜南县,32.638102,115.590534;"
    AddressDistrict = AddressDistrict & "341200,341226,颍上县,32.637065,116.259122;"
    AddressDistrict = AddressDistrict & "341200,341282,界首市,33.26153,115.362117;"
    AddressDistrict = AddressDistrict & "341300,341302,埇桥区,33.633853,116.983309;"
    AddressDistrict = AddressDistrict & "341300,341321,砀山县,34.426247,116.351113;"
    AddressDistrict = AddressDistrict & "341300,341322,萧县,34.183266,116.945399;"
    AddressDistrict = AddressDistrict & "341300,341323,灵璧县,33.540629,117.551493;"
    AddressDistrict = AddressDistrict & "341300,341324,泗县,33.47758,117.885443;"
    AddressDistrict = AddressDistrict & "341500,341502,金安区,31.754491,116.503288;"
    AddressDistrict = AddressDistrict & "341500,341503,裕安区,31.750692,116.494543;"
    AddressDistrict = AddressDistrict & "341500,341504,叶集区,31.84768,115.913594;"
    AddressDistrict = AddressDistrict & "341500,341522,霍邱县,32.341305,116.278875;"
    AddressDistrict = AddressDistrict & "341500,341523,舒城县,31.462848,116.944088;"
    AddressDistrict = AddressDistrict & "341500,341524,金寨县,31.681624,115.878514;"
    AddressDistrict = AddressDistrict & "341500,341525,霍山县,31.402456,116.333078;"
    AddressDistrict = AddressDistrict & "341600,341602,谯城区,33.869284,115.781214;"
    AddressDistrict = AddressDistrict & "341600,341621,涡阳县,33.502831,116.211551;"
    AddressDistrict = AddressDistrict & "341600,341622,蒙城县,33.260814,116.560337;"
    AddressDistrict = AddressDistrict & "341600,341623,利辛县,33.143503,116.207782;"
    AddressDistrict = AddressDistrict & "341700,341702,贵池区,30.657378,117.488342;"
    AddressDistrict = AddressDistrict & "341700,341721,东至县,30.096568,117.021476;"
    AddressDistrict = AddressDistrict & "341700,341722,石台县,30.210324,117.482907;"
    AddressDistrict = AddressDistrict & "341700,341723,青阳县,30.63818,117.857395;"
    AddressDistrict = AddressDistrict & "341800,341802,宣州区,30.946003,118.758412;"
    AddressDistrict = AddressDistrict & "341800,341821,郎溪县,31.127834,119.185024;"
    AddressDistrict = AddressDistrict & "341800,341823,泾县,30.685975,118.412397;"
    AddressDistrict = AddressDistrict & "341800,341824,绩溪县,30.065267,118.594705;"
    AddressDistrict = AddressDistrict & "341800,341825,旌德县,30.288057,118.543081;"
    AddressDistrict = AddressDistrict & "341800,341881,宁国市,30.626529,118.983407;"
    AddressDistrict = AddressDistrict & "341800,341882,广德市,30.893116,119.417521;"
    AddressDistrict = AddressDistrict & "350100,350102,鼓楼区,26.082284,119.29929;"
    AddressDistrict = AddressDistrict & "350100,350103,台江区,26.058616,119.310156;"
    AddressDistrict = AddressDistrict & "350100,350104,仓山区,26.038912,119.320988;"
    AddressDistrict = AddressDistrict & "350100,350105,马尾区,25.991975,119.458725;"
    AddressDistrict = AddressDistrict & "350100,350111,晋安区,26.078837,119.328597;"
    AddressDistrict = AddressDistrict & "350100,350112,长乐区,25.960583,119.510849;"
    AddressDistrict = AddressDistrict & "350100,350121,闽侯县,26.148567,119.145117;"
    AddressDistrict = AddressDistrict & "350100,350122,连江县,26.202109,119.538365;"
    AddressDistrict = AddressDistrict & "350100,350123,罗源县,26.487234,119.552645;"
    AddressDistrict = AddressDistrict & "350100,350124,闽清县,26.223793,118.868416;"
    AddressDistrict = AddressDistrict & "350100,350125,永泰县,25.864825,118.939089;"
    AddressDistrict = AddressDistrict & "350100,350128,平潭县,25.503672,119.791197;"
    AddressDistrict = AddressDistrict & "350100,350181,福清市,25.720402,119.376992;"
    AddressDistrict = AddressDistrict & "350200,350203,思明区,24.462059,118.087828;"
    AddressDistrict = AddressDistrict & "350200,350205,海沧区,24.492512,118.036364;"
    AddressDistrict = AddressDistrict & "350200,350206,湖里区,24.512764,118.10943;"
    AddressDistrict = AddressDistrict & "350200,350211,集美区,24.572874,118.100869;"
    AddressDistrict = AddressDistrict & "350200,350212,同安区,24.729333,118.150455;"
    AddressDistrict = AddressDistrict & "350200,350213,翔安区,24.637479,118.242811;"
    AddressDistrict = AddressDistrict & "350300,350302,城厢区,25.433737,119.001028;"
    AddressDistrict = AddressDistrict & "350300,350303,涵江区,25.459273,119.119102;"
    AddressDistrict = AddressDistrict & "350300,350304,荔城区,25.430047,119.020047;"
    AddressDistrict = AddressDistrict & "350300,350305,秀屿区,25.316141,119.092607;"
    AddressDistrict = AddressDistrict & "350300,350322,仙游县,25.356529,118.694331;"
    AddressDistrict = AddressDistrict & "350400,350403,三元区,26.234191,117.607418;"
    AddressDistrict = AddressDistrict & "350400,350421,明溪县,26.357375,117.201845;"
    AddressDistrict = AddressDistrict & "350400,350423,清流县,26.17761,116.815821;"
    AddressDistrict = AddressDistrict & "350400,350424,宁化县,26.259932,116.659725;"
    AddressDistrict = AddressDistrict & "350400,350425,大田县,25.690803,117.849355;"
    AddressDistrict = AddressDistrict & "350400,350426,尤溪县,26.169261,118.188577;"
    AddressDistrict = AddressDistrict & "350400,350427,沙县区,26.397361,117.789095;"
    AddressDistrict = AddressDistrict & "350400,350428,将乐县,26.728667,117.473558;"
    AddressDistrict = AddressDistrict & "350400,350429,泰宁县,26.897995,117.177522;"
    AddressDistrict = AddressDistrict & "350400,350430,建宁县,26.831398,116.845832;"
    AddressDistrict = AddressDistrict & "350400,350481,永安市,25.974075,117.364447;"
    AddressDistrict = AddressDistrict & "350500,350502,鲤城区,24.907645,118.588929;"
    AddressDistrict = AddressDistrict & "350500,350503,丰泽区,24.896041,118.605147;"
    AddressDistrict = AddressDistrict & "350500,350504,洛江区,24.941153,118.670312;"
    AddressDistrict = AddressDistrict & "350500,350505,泉港区,25.126859,118.912285;"
    AddressDistrict = AddressDistrict & "350500,350521,惠安县,25.028718,118.798954;"
    AddressDistrict = AddressDistrict & "350500,350524,安溪县,25.056824,118.186014;"
    AddressDistrict = AddressDistrict & "350500,350525,永春县,25.320721,118.29503;"
    AddressDistrict = AddressDistrict & "350500,350526,德化县,25.489004,118.242986;"
    AddressDistrict = AddressDistrict & "350500,350527,金门县,24.436417,118.323221;"
    AddressDistrict = AddressDistrict & "350500,350581,石狮市,24.731978,118.628402;"
    AddressDistrict = AddressDistrict & "350500,350582,晋江市,24.807322,118.577338;"
    AddressDistrict = AddressDistrict & "350500,350583,南安市,24.959494,118.387031;"
    AddressDistrict = AddressDistrict & "350600,350602,芗城区,24.509955,117.656461;"
    AddressDistrict = AddressDistrict & "350600,350603,龙文区,24.515656,117.671387;"
    AddressDistrict = AddressDistrict & "350600,350622,云霄县,23.950486,117.340946;"
    AddressDistrict = AddressDistrict & "350600,350623,漳浦县,24.117907,117.614023;"
    AddressDistrict = AddressDistrict & "350600,350624,诏安县,23.710834,117.176083;"
    AddressDistrict = AddressDistrict & "350600,350625,长泰区,24.621475,117.755913;"
    AddressDistrict = AddressDistrict & "350600,350626,东山县,23.702845,117.427679;"
    AddressDistrict = AddressDistrict & "350600,350627,南靖县,24.516425,117.365462;"
    AddressDistrict = AddressDistrict & "350600,350628,平和县,24.366158,117.313549;"
    AddressDistrict = AddressDistrict & "350600,350629,华安县,25.001416,117.53631;"
    AddressDistrict = AddressDistrict & "350600,350681,龙海区,24.445341,117.817292;"
    AddressDistrict = AddressDistrict & "350700,350702,延平区,26.636079,118.178918;"
    AddressDistrict = AddressDistrict & "350700,350703,建阳区,27.332067,118.12267;"
    AddressDistrict = AddressDistrict & "350700,350721,顺昌县,26.792851,117.80771;"
    AddressDistrict = AddressDistrict & "350700,350722,浦城县,27.920412,118.536822;"
    AddressDistrict = AddressDistrict & "350700,350723,光泽县,27.542803,117.337897;"
    AddressDistrict = AddressDistrict & "350700,350724,松溪县,27.525785,118.783491;"
    AddressDistrict = AddressDistrict & "350700,350725,政和县,27.365398,118.858661;"
    AddressDistrict = AddressDistrict & "350700,350781,邵武市,27.337952,117.491544;"
    AddressDistrict = AddressDistrict & "350700,350782,武夷山市,27.751733,118.032796;"
    AddressDistrict = AddressDistrict & "350700,350783,建瓯市,27.03502,118.321765;"
    AddressDistrict = AddressDistrict & "350800,350802,新罗区,25.0918,117.030721;"
    AddressDistrict = AddressDistrict & "350800,350803,永定区,24.720442,116.732691;"
    AddressDistrict = AddressDistrict & "350800,350821,长汀县,25.842278,116.361007;"
    AddressDistrict = AddressDistrict & "350800,350823,上杭县,25.050019,116.424774;"
    AddressDistrict = AddressDistrict & "350800,350824,武平县,25.08865,116.100928;"
    AddressDistrict = AddressDistrict & "350800,350825,连城县,25.708506,116.756687;"
    AddressDistrict = AddressDistrict & "350800,350881,漳平市,25.291597,117.42073;"
    AddressDistrict = AddressDistrict & "350900,350902,蕉城区,26.659253,119.527225;"
    AddressDistrict = AddressDistrict & "350900,350921,霞浦县,26.882068,120.005214;"
    AddressDistrict = AddressDistrict & "350900,350922,古田县,26.577491,118.743156;"
    AddressDistrict = AddressDistrict & "350900,350923,屏南县,26.910826,118.987544;"
    AddressDistrict = AddressDistrict & "350900,350924,寿宁县,27.457798,119.506733;"
    AddressDistrict = AddressDistrict & "350900,350925,周宁县,27.103106,119.338239;"
    AddressDistrict = AddressDistrict & "350900,350926,柘荣县,27.236163,119.898226;"
    AddressDistrict = AddressDistrict & "350900,350981,福安市,27.084246,119.650798;"
    AddressDistrict = AddressDistrict & "350900,350982,福鼎市,27.318884,120.219761;"
    AddressDistrict = AddressDistrict & "360100,360102,东湖区,28.682988,115.889675;"
    AddressDistrict = AddressDistrict & "360100,360103,西湖区,28.662901,115.91065;"
    AddressDistrict = AddressDistrict & "360100,360104,青云谱区,28.635724,115.907292;"
    AddressDistrict = AddressDistrict & "360100,360111,青山湖区,28.689292,115.949044;"
    AddressDistrict = AddressDistrict & "360100,360112,新建区,28.690788,115.820806;"
    AddressDistrict = AddressDistrict & "360100,360113,红谷滩区,28.69819928,115.8580521;"
    AddressDistrict = AddressDistrict & "360100,360121,南昌县,28.543781,115.942465;"
    AddressDistrict = AddressDistrict & "360100,360123,安义县,28.841334,115.553109;"
    AddressDistrict = AddressDistrict & "360100,360124,进贤县,28.365681,116.267671;"
    AddressDistrict = AddressDistrict & "360200,360202,昌江区,29.288465,117.195023;"
    AddressDistrict = AddressDistrict & "360200,360203,珠山区,29.292812,117.214814;"
    AddressDistrict = AddressDistrict & "360200,360222,浮梁县,29.352251,117.217611;"
    AddressDistrict = AddressDistrict & "360200,360281,乐平市,28.967361,117.129376;"
    AddressDistrict = AddressDistrict & "360300,360302,安源区,27.625826,113.855044;"
    AddressDistrict = AddressDistrict & "360300,360313,湘东区,27.639319,113.7456;"
    AddressDistrict = AddressDistrict & "360300,360321,莲花县,27.127807,113.955582;"
    AddressDistrict = AddressDistrict & "360300,360322,上栗县,27.877041,113.800525;"
    AddressDistrict = AddressDistrict & "360300,360323,芦溪县,27.633633,114.041206;"
    AddressDistrict = AddressDistrict & "360400,360402,濂溪区,29.676175,115.99012;"
    AddressDistrict = AddressDistrict & "360400,360403,浔阳区,29.72465,115.995947;"
    AddressDistrict = AddressDistrict & "360400,360404,柴桑区,29.610264,115.892977;"
    AddressDistrict = AddressDistrict & "360400,360423,武宁县,29.260182,115.105646;"
    AddressDistrict = AddressDistrict & "360400,360424,修水县,29.032729,114.573428;"
    AddressDistrict = AddressDistrict & "360400,360425,永修县,29.018212,115.809055;"
    AddressDistrict = AddressDistrict & "360400,360426,德安县,29.327474,115.762611;"
    AddressDistrict = AddressDistrict & "360400,360428,都昌县,29.275105,116.205114;"
    AddressDistrict = AddressDistrict & "360400,360429,湖口县,29.7263,116.244313;"
    AddressDistrict = AddressDistrict & "360400,360430,彭泽县,29.898865,116.55584;"
    AddressDistrict = AddressDistrict & "360400,360481,瑞昌市,29.676599,115.669081;"
    AddressDistrict = AddressDistrict & "360400,360482,共青城市,29.247884,115.805712;"
    AddressDistrict = AddressDistrict & "360400,360483,庐山市,29.456169,116.043743;"
    AddressDistrict = AddressDistrict & "360500,360502,渝水区,27.819171,114.923923;"
    AddressDistrict = AddressDistrict & "360500,360521,分宜县,27.811301,114.675262;"
    AddressDistrict = AddressDistrict & "360600,360602,月湖区,28.239076,117.034112;"
    AddressDistrict = AddressDistrict & "360600,360603,余江区,28.206177,116.822763;"
    AddressDistrict = AddressDistrict & "360600,360681,贵溪市,28.283693,117.212103;"
    AddressDistrict = AddressDistrict & "360700,360702,章贡区,25.851367,114.93872;"
    AddressDistrict = AddressDistrict & "360700,360703,南康区,25.661721,114.756933;"
    AddressDistrict = AddressDistrict & "360700,360704,赣县区,25.865432,115.018461;"
    AddressDistrict = AddressDistrict & "360700,360722,信丰县,25.38023,114.930893;"
    AddressDistrict = AddressDistrict & "360700,360723,大余县,25.395937,114.362243;"
    AddressDistrict = AddressDistrict & "360700,360724,上犹县,25.794284,114.540537;"
    AddressDistrict = AddressDistrict & "360700,360725,崇义县,25.687911,114.307348;"
    AddressDistrict = AddressDistrict & "360700,360726,安远县,25.134591,115.392328;"
    AddressDistrict = AddressDistrict & "360700,360728,定南县,24.774277,115.03267;"
    AddressDistrict = AddressDistrict & "360700,360729,全南县,24.742651,114.531589;"
    AddressDistrict = AddressDistrict & "360700,360730,宁都县,26.472054,116.018782;"
    AddressDistrict = AddressDistrict & "360700,360731,于都县,25.955033,115.411198;"
    AddressDistrict = AddressDistrict & "360700,360732,兴国县,26.330489,115.351896;"
    AddressDistrict = AddressDistrict & "360700,360733,会昌县,25.599125,115.791158;"
    AddressDistrict = AddressDistrict & "360700,360734,寻乌县,24.954136,115.651399;"
    AddressDistrict = AddressDistrict & "360700,360735,石城县,26.326582,116.342249;"
    AddressDistrict = AddressDistrict & "360700,360781,瑞金市,25.875278,116.034854;"
    AddressDistrict = AddressDistrict & "360700,360783,龙南市,24.90476,114.792657;"
    AddressDistrict = AddressDistrict & "360800,360802,吉州区,27.112367,114.987331;"
    AddressDistrict = AddressDistrict & "360800,360803,青原区,27.105879,115.016306;"
    AddressDistrict = AddressDistrict & "360800,360821,吉安县,27.040042,114.905117;"
    AddressDistrict = AddressDistrict & "360800,360822,吉水县,27.213445,115.134569;"
    AddressDistrict = AddressDistrict & "360800,360823,峡江县,27.580862,115.319331;"
    AddressDistrict = AddressDistrict & "360800,360824,新干县,27.755758,115.399294;"
    AddressDistrict = AddressDistrict & "360800,360825,永丰县,27.321087,115.435559;"
    AddressDistrict = AddressDistrict & "360800,360826,泰和县,26.790164,114.901393;"
    AddressDistrict = AddressDistrict & "360800,360827,遂川县,26.323705,114.51689;"
    AddressDistrict = AddressDistrict & "360800,360828,万安县,26.462085,114.784694;"
    AddressDistrict = AddressDistrict & "360800,360829,安福县,27.382746,114.61384;"
    AddressDistrict = AddressDistrict & "360800,360830,永新县,26.944721,114.242534;"
    AddressDistrict = AddressDistrict & "360800,360881,井冈山市,26.745919,114.284421;"
    AddressDistrict = AddressDistrict & "360900,360902,袁州区,27.800117,114.387379;"
    AddressDistrict = AddressDistrict & "360900,360921,奉新县,28.700672,115.389899;"
    AddressDistrict = AddressDistrict & "360900,360922,万载县,28.104528,114.449012;"
    AddressDistrict = AddressDistrict & "360900,360923,上高县,28.234789,114.932653;"
    AddressDistrict = AddressDistrict & "360900,360924,宜丰县,28.388289,114.787381;"
    AddressDistrict = AddressDistrict & "360900,360925,靖安县,28.86054,115.361744;"
    AddressDistrict = AddressDistrict & "360900,360926,铜鼓县,28.520956,114.37014;"
    AddressDistrict = AddressDistrict & "360900,360981,丰城市,28.191584,115.786005;"
    AddressDistrict = AddressDistrict & "360900,360982,樟树市,28.055898,115.543388;"
    AddressDistrict = AddressDistrict & "360900,360983,高安市,28.420951,115.381527;"
    AddressDistrict = AddressDistrict & "361000,361002,临川区,27.981919,116.361404;"
    AddressDistrict = AddressDistrict & "361000,361003,东乡区,28.2325,116.605341;"
    AddressDistrict = AddressDistrict & "361000,361021,南城县,27.55531,116.63945;"
    AddressDistrict = AddressDistrict & "361000,361022,黎川县,27.292561,116.91457;"
    AddressDistrict = AddressDistrict & "361000,361023,南丰县,27.210132,116.532994;"
    AddressDistrict = AddressDistrict & "361000,361024,崇仁县,27.760907,116.059109;"
    AddressDistrict = AddressDistrict & "361000,361025,乐安县,27.420101,115.838432;"
    AddressDistrict = AddressDistrict & "361000,361026,宜黄县,27.546512,116.223023;"
    AddressDistrict = AddressDistrict & "361000,361027,金溪县,27.907387,116.778751;"
    AddressDistrict = AddressDistrict & "361000,361028,资溪县,27.70653,117.066095;"
    AddressDistrict = AddressDistrict & "361000,361030,广昌县,26.838426,116.327291;"
    AddressDistrict = AddressDistrict & "361100,361102,信州区,28.445378,117.970522;"
    AddressDistrict = AddressDistrict & "361100,361103,广丰区,28.440285,118.189852;"
    AddressDistrict = AddressDistrict & "361100,361104,广信区,28.453897,117.90612;"
    AddressDistrict = AddressDistrict & "361100,361123,玉山县,28.673479,118.244408;"
    AddressDistrict = AddressDistrict & "361100,361124,铅山县,28.310892,117.711906;"
    AddressDistrict = AddressDistrict & "361100,361125,横峰县,28.415103,117.608247;"
    AddressDistrict = AddressDistrict & "361100,361126,弋阳县,28.402391,117.435002;"
    AddressDistrict = AddressDistrict & "361100,361127,余干县,28.69173,116.691072;"
    AddressDistrict = AddressDistrict & "361100,361128,鄱阳县,28.993374,116.673748;"
    AddressDistrict = AddressDistrict & "361100,361129,万年县,28.692589,117.07015;"
    AddressDistrict = AddressDistrict & "361100,361130,婺源县,29.254015,117.86219;"
    AddressDistrict = AddressDistrict & "361100,361181,德兴市,28.945034,117.578732;"
    AddressDistrict = AddressDistrict & "370100,370102,历下区,36.664169,117.03862;"
    AddressDistrict = AddressDistrict & "370100,370103,市中区,36.657354,116.99898;"
    AddressDistrict = AddressDistrict & "370100,370104,槐荫区,36.668205,116.947921;"
    AddressDistrict = AddressDistrict & "370100,370105,天桥区,36.693374,116.996086;"
    AddressDistrict = AddressDistrict & "370100,370112,历城区,36.681744,117.063744;"
    AddressDistrict = AddressDistrict & "370100,370113,长清区,36.561049,116.74588;"
    AddressDistrict = AddressDistrict & "370100,370114,章丘区,36.71209,117.54069;"
    AddressDistrict = AddressDistrict & "370100,370115,济阳区,36.976771,117.176035;"
    AddressDistrict = AddressDistrict & "370100,370116,莱芜区,36.214395,117.675808;"
    AddressDistrict = AddressDistrict & "370100,370117,钢城区,36.058038,117.82033;"
    AddressDistrict = AddressDistrict & "370100,370124,平阴县,36.286923,116.455054;"
    AddressDistrict = AddressDistrict & "370100,370126,商河县,37.310544,117.156369;"
    AddressDistrict = AddressDistrict & "370200,370202,市南区,36.070892,120.395966;"
    AddressDistrict = AddressDistrict & "370200,370203,市北区,36.083819,120.355026;"
    AddressDistrict = AddressDistrict & "370200,370211,黄岛区,35.875138,119.995518;"
    AddressDistrict = AddressDistrict & "370200,370212,崂山区,36.102569,120.467393;"
    AddressDistrict = AddressDistrict & "370200,370213,李沧区,36.160023,120.421236;"
    AddressDistrict = AddressDistrict & "370200,370214,城阳区,36.306833,120.389135;"
    AddressDistrict = AddressDistrict & "370200,370215,即墨区,36.390847,120.447352;"
    AddressDistrict = AddressDistrict & "370200,370281,胶州市,36.285878,120.006202;"
    AddressDistrict = AddressDistrict & "370200,370283,平度市,36.788828,119.959012;"
    AddressDistrict = AddressDistrict & "370200,370285,莱西市,36.86509,120.526226;"
    AddressDistrict = AddressDistrict & "370300,370302,淄川区,36.647272,117.967696;"
    AddressDistrict = AddressDistrict & "370300,370303,张店区,36.807049,118.053521;"
    AddressDistrict = AddressDistrict & "370300,370304,博山区,36.497567,117.85823;"
    AddressDistrict = AddressDistrict & "370300,370305,临淄区,36.816657,118.306018;"
    AddressDistrict = AddressDistrict & "370300,370306,周村区,36.803699,117.851036;"
    AddressDistrict = AddressDistrict & "370300,370321,桓台县,36.959773,118.101556;"
    AddressDistrict = AddressDistrict & "370300,370322,高青县,37.169581,117.829839;"
    AddressDistrict = AddressDistrict & "370300,370323,沂源县,36.186282,118.166161;"
    AddressDistrict = AddressDistrict & "370400,370402,市中区,34.856651,117.557281;"
    AddressDistrict = AddressDistrict & "370400,370403,薛城区,34.79789,117.265293;"
    AddressDistrict = AddressDistrict & "370400,370404,峄城区,34.767713,117.586316;"
    AddressDistrict = AddressDistrict & "370400,370405,台儿庄区,34.564815,117.734747;"
    AddressDistrict = AddressDistrict & "370400,370406,山亭区,35.096077,117.458968;"
    AddressDistrict = AddressDistrict & "370400,370481,滕州市,35.088498,117.162098;"
    AddressDistrict = AddressDistrict & "370500,370502,东营区,37.461567,118.507543;"
    AddressDistrict = AddressDistrict & "370500,370503,河口区,37.886015,118.529613;"
    AddressDistrict = AddressDistrict & "370500,370505,垦利区,37.588679,118.551314;"
    AddressDistrict = AddressDistrict & "370500,370522,利津县,37.493365,118.248854;"
    AddressDistrict = AddressDistrict & "370500,370523,广饶县,37.05161,118.407522;"
    AddressDistrict = AddressDistrict & "370600,370602,芝罘区,37.540925,121.385877;"
    AddressDistrict = AddressDistrict & "370600,370611,福山区,37.496875,121.264741;"
    AddressDistrict = AddressDistrict & "370600,370612,牟平区,37.388356,121.60151;"
    AddressDistrict = AddressDistrict & "370600,370613,莱山区,37.473549,121.448866;"
    AddressDistrict = AddressDistrict & "370600,370614,蓬莱区,37.811045,120.759074;"
    AddressDistrict = AddressDistrict & "370600,370681,龙口市,37.648446,120.528328;"
    AddressDistrict = AddressDistrict & "370600,370682,莱阳市,36.977037,120.711151;"
    AddressDistrict = AddressDistrict & "370600,370683,莱州市,37.182725,119.942135;"
    AddressDistrict = AddressDistrict & "370600,370685,招远市,37.364919,120.403142;"
    AddressDistrict = AddressDistrict & "370600,370686,栖霞市,37.305854,120.834097;"
    AddressDistrict = AddressDistrict & "370600,370687,海阳市,36.780657,121.168392;"
    AddressDistrict = AddressDistrict & "370700,370702,潍城区,36.710062,119.103784;"
    AddressDistrict = AddressDistrict & "370700,370703,寒亭区,36.772103,119.207866;"
    AddressDistrict = AddressDistrict & "370700,370704,坊子区,36.654616,119.166326;"
    AddressDistrict = AddressDistrict & "370700,370705,奎文区,36.709494,119.137357;"
    AddressDistrict = AddressDistrict & "370700,370724,临朐县,36.516371,118.539876;"
    AddressDistrict = AddressDistrict & "370700,370725,昌乐县,36.703253,118.839995;"
    AddressDistrict = AddressDistrict & "370700,370781,青州市,36.697855,118.484693;"
    AddressDistrict = AddressDistrict & "370700,370782,诸城市,35.997093,119.403182;"
    AddressDistrict = AddressDistrict & "370700,370783,寿光市,36.874411,118.736451;"
    AddressDistrict = AddressDistrict & "370700,370784,安丘市,36.427417,119.206886;"
    AddressDistrict = AddressDistrict & "370700,370785,高密市,36.37754,119.757033;"
    AddressDistrict = AddressDistrict & "370700,370786,昌邑市,36.854937,119.394502;"
    AddressDistrict = AddressDistrict & "370800,370811,任城区,35.414828,116.595261;"
    AddressDistrict = AddressDistrict & "370800,370812,兖州区,35.556445,116.828996;"
    AddressDistrict = AddressDistrict & "370800,370826,微山县,34.809525,117.12861;"
    AddressDistrict = AddressDistrict & "370800,370827,鱼台县,34.997706,116.650023;"
    AddressDistrict = AddressDistrict & "370800,370828,金乡县,35.06977,116.310364;"
    AddressDistrict = AddressDistrict & "370800,370829,嘉祥县,35.398098,116.342885;"
    AddressDistrict = AddressDistrict & "370800,370830,汶上县,35.721746,116.487146;"
    AddressDistrict = AddressDistrict & "370800,370831,泗水县,35.653216,117.273605;"
    AddressDistrict = AddressDistrict & "370800,370832,梁山县,35.801843,116.08963;"
    AddressDistrict = AddressDistrict & "370800,370881,曲阜市,35.592788,116.991885;"
    AddressDistrict = AddressDistrict & "370800,370883,邹城市,35.405259,116.96673;"
    AddressDistrict = AddressDistrict & "370900,370902,泰山区,36.189313,117.129984;"
    AddressDistrict = AddressDistrict & "370900,370911,岱岳区,36.1841,117.04353;"
    AddressDistrict = AddressDistrict & "370900,370921,宁阳县,35.76754,116.799297;"
    AddressDistrict = AddressDistrict & "370900,370923,东平县,35.930467,116.461052;"
    AddressDistrict = AddressDistrict & "370900,370982,新泰市,35.910387,117.766092;"
    AddressDistrict = AddressDistrict & "370900,370983,肥城市,36.1856,116.763703;"
    AddressDistrict = AddressDistrict & "371000,371002,环翠区,37.510754,122.116189;"
    AddressDistrict = AddressDistrict & "371000,371003,文登区,37.196211,122.057139;"
    AddressDistrict = AddressDistrict & "371000,371082,荣成市,37.160134,122.422896;"
    AddressDistrict = AddressDistrict & "371000,371083,乳山市,36.919622,121.536346;"
    AddressDistrict = AddressDistrict & "371100,371102,东港区,35.426152,119.457703;"
    AddressDistrict = AddressDistrict & "371100,371103,岚山区,35.119794,119.315844;"
    AddressDistrict = AddressDistrict & "371100,371121,五莲县,35.751936,119.206745;"
    AddressDistrict = AddressDistrict & "371100,371122,莒县,35.588115,118.832859;"
    AddressDistrict = AddressDistrict & "371300,371302,兰山区,35.061631,118.327667;"
    AddressDistrict = AddressDistrict & "371300,371311,罗庄区,34.997204,118.284795;"
    AddressDistrict = AddressDistrict & "371300,371312,河东区,35.085004,118.398296;"
    AddressDistrict = AddressDistrict & "371300,371321,沂南县,35.547002,118.455395;"
    AddressDistrict = AddressDistrict & "371300,371322,郯城县,34.614741,118.342963;"
    AddressDistrict = AddressDistrict & "371300,371323,沂水县,35.787029,118.634543;"
    AddressDistrict = AddressDistrict & "371300,371324,兰陵县,34.855573,118.049968;"
    AddressDistrict = AddressDistrict & "371300,371325,费县,35.269174,117.968869;"
    AddressDistrict = AddressDistrict & "371300,371326,平邑县,35.511519,117.631884;"
    AddressDistrict = AddressDistrict & "371300,371327,莒南县,35.175911,118.838322;"
    AddressDistrict = AddressDistrict & "371300,371328,蒙阴县,35.712435,117.943271;"
    AddressDistrict = AddressDistrict & "371300,371329,临沭县,34.917062,118.648379;"
    AddressDistrict = AddressDistrict & "371400,371402,德城区,37.453923,116.307076;"
    AddressDistrict = AddressDistrict & "371400,371403,陵城区,37.332848,116.574929;"
    AddressDistrict = AddressDistrict & "371400,371422,宁津县,37.649619,116.79372;"
    AddressDistrict = AddressDistrict & "371400,371423,庆云县,37.777724,117.390507;"
    AddressDistrict = AddressDistrict & "371400,371424,临邑县,37.192044,116.867028;"
    AddressDistrict = AddressDistrict & "371400,371425,齐河县,36.795497,116.758394;"
    AddressDistrict = AddressDistrict & "371400,371426,平原县,37.164465,116.433904;"
    AddressDistrict = AddressDistrict & "371400,371427,夏津县,36.950501,116.003816;"
    AddressDistrict = AddressDistrict & "371400,371428,武城县,37.209527,116.078627;"
    AddressDistrict = AddressDistrict & "371400,371481,乐陵市,37.729115,117.216657;"
    AddressDistrict = AddressDistrict & "371400,371482,禹城市,36.934485,116.642554;"
    AddressDistrict = AddressDistrict & "371500,371502,东昌府区,36.45606,115.980023;"
    AddressDistrict = AddressDistrict & "371500,371503,茌平区,36.591934,116.25335;"
    AddressDistrict = AddressDistrict & "371500,371521,阳谷县,36.113708,115.784287;"
    AddressDistrict = AddressDistrict & "371500,371522,莘县,36.237597,115.667291;"
    AddressDistrict = AddressDistrict & "371500,371524,东阿县,36.336004,116.248855;"
    AddressDistrict = AddressDistrict & "371500,371525,冠县,36.483753,115.444808;"
    AddressDistrict = AddressDistrict & "371500,371526,高唐县,36.859755,116.229662;"
    AddressDistrict = AddressDistrict & "371500,371581,临清市,36.842598,115.713462;"
    AddressDistrict = AddressDistrict & "371600,371602,滨城区,37.384842,118.020149;"
    AddressDistrict = AddressDistrict & "371600,371603,沾化区,37.698456,118.129902;"
    AddressDistrict = AddressDistrict & "371600,371621,惠民县,37.483876,117.508941;"
    AddressDistrict = AddressDistrict & "371600,371622,阳信县,37.640492,117.581326;"
    AddressDistrict = AddressDistrict & "371600,371623,无棣县,37.740848,117.616325;"
    AddressDistrict = AddressDistrict & "371600,371625,博兴县,37.147002,118.123096;"
    AddressDistrict = AddressDistrict & "371600,371681,邹平市,36.87803,117.736807;"
    AddressDistrict = AddressDistrict & "371700,371702,牡丹区,35.24311,115.470946;"
    AddressDistrict = AddressDistrict & "371700,371703,定陶区,35.072701,115.569601;"
    AddressDistrict = AddressDistrict & "371700,371721,曹县,34.823253,115.549482;"
    AddressDistrict = AddressDistrict & "371700,371722,单县,34.790851,116.08262;"
    AddressDistrict = AddressDistrict & "371700,371723,成武县,34.947366,115.897349;"
    AddressDistrict = AddressDistrict & "371700,371724,巨野县,35.390999,116.089341;"
    AddressDistrict = AddressDistrict & "371700,371725,郓城县,35.594773,115.93885;"
    AddressDistrict = AddressDistrict & "371700,371726,鄄城县,35.560257,115.51434;"
    AddressDistrict = AddressDistrict & "371700,371728,东明县,35.289637,115.098412;"
    AddressDistrict = AddressDistrict & "410100,410102,中原区,34.748286,113.611576;"
    AddressDistrict = AddressDistrict & "410100,410103,二七区,34.730936,113.645422;"
    AddressDistrict = AddressDistrict & "410100,410104,管城回族区,34.746453,113.685313;"
    AddressDistrict = AddressDistrict & "410100,410105,金水区,34.775838,113.686037;"
    AddressDistrict = AddressDistrict & "410100,410106,上街区,34.808689,113.298282;"
    AddressDistrict = AddressDistrict & "410100,410108,惠济区,34.828591,113.61836;"
    AddressDistrict = AddressDistrict & "410100,410122,中牟县,34.721976,114.022521;"
    AddressDistrict = AddressDistrict & "410100,410181,巩义市,34.75218,112.98283;"
    AddressDistrict = AddressDistrict & "410100,410182,荥阳市,34.789077,113.391523;"
    AddressDistrict = AddressDistrict & "410100,410183,新密市,34.537846,113.380616;"
    AddressDistrict = AddressDistrict & "410100,410184,新郑市,34.394219,113.73967;"
    AddressDistrict = AddressDistrict & "410100,410185,登封市,34.459939,113.037768;"
    AddressDistrict = AddressDistrict & "410200,410202,龙亭区,34.799833,114.353348;"
    AddressDistrict = AddressDistrict & "410200,410203,顺河回族区,34.800459,114.364875;"
    AddressDistrict = AddressDistrict & "410200,410204,鼓楼区,34.792383,114.3485;"
    AddressDistrict = AddressDistrict & "410200,410205,禹王台区,34.779727,114.350246;"
    AddressDistrict = AddressDistrict & "410200,410212,祥符区,34.756476,114.437622;"
    AddressDistrict = AddressDistrict & "410200,410221,杞县,34.554585,114.770472;"
    AddressDistrict = AddressDistrict & "410200,410222,通许县,34.477302,114.467734;"
    AddressDistrict = AddressDistrict & "410200,410223,尉氏县,34.412256,114.193927;"
    AddressDistrict = AddressDistrict & "410200,410225,兰考县,34.829899,114.820572;"
    AddressDistrict = AddressDistrict & "410300,410302,老城区,34.682945,112.477298;"
    AddressDistrict = AddressDistrict & "410300,410303,西工区,34.667847,112.443232;"
    AddressDistrict = AddressDistrict & "410300,410304,瀍河回族区,34.684738,112.491625;"
    AddressDistrict = AddressDistrict & "410300,410305,涧西区,34.654251,112.399243;"
    AddressDistrict = AddressDistrict & "410300,410306,孟津区,34.826485,112.443892;"
    AddressDistrict = AddressDistrict & "410300,410311,洛龙区,34.618557,112.456634;"
    AddressDistrict = AddressDistrict & "410300,410323,新安县,34.728679,112.141403;"
    AddressDistrict = AddressDistrict & "410300,410324,栾川县,33.783195,111.618386;"
    AddressDistrict = AddressDistrict & "410300,410325,嵩县,34.131563,112.087765;"
    AddressDistrict = AddressDistrict & "410300,410326,汝阳县,34.15323,112.473789;"
    AddressDistrict = AddressDistrict & "410300,410327,宜阳县,34.516478,112.179989;"
    AddressDistrict = AddressDistrict & "410300,410328,洛宁县,34.387179,111.655399;"
    AddressDistrict = AddressDistrict & "410300,410329,伊川县,34.423416,112.429384;"
    AddressDistrict = AddressDistrict & "410300,410381,偃师区,34.723042,112.787739;"
    AddressDistrict = AddressDistrict & "410400,410402,新华区,33.737579,113.299061;"
    AddressDistrict = AddressDistrict & "410400,410403,卫东区,33.739285,113.310327;"
    AddressDistrict = AddressDistrict & "410400,410404,石龙区,33.901538,112.889885;"
    AddressDistrict = AddressDistrict & "410400,410411,湛河区,33.725681,113.320873;"
    AddressDistrict = AddressDistrict & "410400,410421,宝丰县,33.866359,113.066812;"
    AddressDistrict = AddressDistrict & "410400,410422,叶县,33.621252,113.358298;"
    AddressDistrict = AddressDistrict & "410400,410423,鲁山县,33.740325,112.906703;"
    AddressDistrict = AddressDistrict & "410400,410425,郏县,33.971993,113.220451;"
    AddressDistrict = AddressDistrict & "410400,410481,舞钢市,33.302082,113.52625;"
    AddressDistrict = AddressDistrict & "410400,410482,汝州市,34.167408,112.845336;"
    AddressDistrict = AddressDistrict & "410500,410502,文峰区,36.098101,114.352562;"
    AddressDistrict = AddressDistrict & "410500,410503,北关区,36.10978,114.352646;"
    AddressDistrict = AddressDistrict & "410500,410505,殷都区,36.108974,114.300098;"
    AddressDistrict = AddressDistrict & "410500,410506,龙安区,36.095568,114.323522;"
    AddressDistrict = AddressDistrict & "410500,410522,安阳县,36.130585,114.130207;"
    AddressDistrict = AddressDistrict & "410500,410523,汤阴县,35.922349,114.362357;"
    AddressDistrict = AddressDistrict & "410500,410526,滑县,35.574628,114.524;"
    AddressDistrict = AddressDistrict & "410500,410527,内黄县,35.953702,114.904582;"
    AddressDistrict = AddressDistrict & "410500,410581,林州市,36.063403,113.823767;"
    AddressDistrict = AddressDistrict & "410600,410602,鹤山区,35.936128,114.166551;"
    AddressDistrict = AddressDistrict & "410600,410603,山城区,35.896058,114.184202;"
    AddressDistrict = AddressDistrict & "410600,410611,淇滨区,35.748382,114.293917;"
    AddressDistrict = AddressDistrict & "410600,410621,浚县,35.671282,114.550162;"
    AddressDistrict = AddressDistrict & "410600,410622,淇县,35.609478,114.200379;"
    AddressDistrict = AddressDistrict & "410700,410702,红旗区,35.302684,113.878158;"
    AddressDistrict = AddressDistrict & "410700,410703,卫滨区,35.304905,113.866065;"
    AddressDistrict = AddressDistrict & "410700,410704,凤泉区,35.379855,113.906712;"
    AddressDistrict = AddressDistrict & "410700,410711,牧野区,35.312974,113.89716;"
    AddressDistrict = AddressDistrict & "410700,410721,新乡县,35.190021,113.806186;"
    AddressDistrict = AddressDistrict & "410700,410724,获嘉县,35.261685,113.657249;"
    AddressDistrict = AddressDistrict & "410700,410725,原阳县,35.054001,113.965966;"
    AddressDistrict = AddressDistrict & "410700,410726,延津县,35.149515,114.200982;"
    AddressDistrict = AddressDistrict & "410700,410727,封丘县,35.04057,114.423405;"
    AddressDistrict = AddressDistrict & "410700,410781,卫辉市,35.404295,114.065855;"
    AddressDistrict = AddressDistrict & "410700,410782,辉县市,35.461318,113.802518;"
    AddressDistrict = AddressDistrict & "410700,410783,长垣市,35.19615,114.673807;"
    AddressDistrict = AddressDistrict & "410800,410802,解放区,35.241353,113.226126;"
    AddressDistrict = AddressDistrict & "410800,410803,中站区,35.236145,113.175485;"
    AddressDistrict = AddressDistrict & "410800,410804,马村区,35.265453,113.321703;"
    AddressDistrict = AddressDistrict & "410800,410811,山阳区,35.21476,113.26766;"
    AddressDistrict = AddressDistrict & "410800,410821,修武县,35.229923,113.447465;"
    AddressDistrict = AddressDistrict & "410800,410822,博爱县,35.170351,113.069313;"
    AddressDistrict = AddressDistrict & "410800,410823,武陟县,35.09885,113.408334;"
    AddressDistrict = AddressDistrict & "410800,410825,温县,34.941233,113.079118;"
    AddressDistrict = AddressDistrict & "410800,410882,沁阳市,35.08901,112.934538;"
    AddressDistrict = AddressDistrict & "410800,410883,孟州市,34.90963,112.78708;"
    AddressDistrict = AddressDistrict & "410900,410902,华龙区,35.760473,115.03184;"
    AddressDistrict = AddressDistrict & "410900,410922,清丰县,35.902413,115.107287;"
    AddressDistrict = AddressDistrict & "410900,410923,南乐县,36.075204,115.204336;"
    AddressDistrict = AddressDistrict & "410900,410926,范县,35.851977,115.504212;"
    AddressDistrict = AddressDistrict & "410900,410927,台前县,35.996474,115.855681;"
    AddressDistrict = AddressDistrict & "410900,410928,濮阳县,35.710349,115.023844;"
    AddressDistrict = AddressDistrict & "411000,411002,魏都区,34.02711,113.828307;"
    AddressDistrict = AddressDistrict & "411000,411003,建安区,34.005018,113.842898;"
    AddressDistrict = AddressDistrict & "411000,411024,鄢陵县,34.100502,114.188507;"
    AddressDistrict = AddressDistrict & "411000,411025,襄城县,33.855943,113.493166;"
    AddressDistrict = AddressDistrict & "411000,411081,禹州市,34.154403,113.471316;"
    AddressDistrict = AddressDistrict & "411000,411082,长葛市,34.219257,113.768912;"
    AddressDistrict = AddressDistrict & "411100,411102,源汇区,33.565441,114.017948;"
    AddressDistrict = AddressDistrict & "411100,411103,郾城区,33.588897,114.016813;"
    AddressDistrict = AddressDistrict & "411100,411104,召陵区,33.567555,114.051686;"
    AddressDistrict = AddressDistrict & "411100,411121,舞阳县,33.436278,113.610565;"
    AddressDistrict = AddressDistrict & "411100,411122,临颍县,33.80609,113.938891;"
    AddressDistrict = AddressDistrict & "411200,411202,湖滨区,34.77812,111.19487;"
    AddressDistrict = AddressDistrict & "411200,411203,陕州区,34.720244,111.103851;"
    AddressDistrict = AddressDistrict & "411200,411221,渑池县,34.763487,111.762992;"
    AddressDistrict = AddressDistrict & "411200,411224,卢氏县,34.053995,111.052649;"
    AddressDistrict = AddressDistrict & "411200,411281,义马市,34.746868,111.869417;"
    AddressDistrict = AddressDistrict & "411200,411282,灵宝市,34.521264,110.88577;"
    AddressDistrict = AddressDistrict & "411300,411302,宛城区,32.994857,112.544591;"
    AddressDistrict = AddressDistrict & "411300,411303,卧龙区,32.989877,112.528789;"
    AddressDistrict = AddressDistrict & "411300,411321,南召县,33.488617,112.435583;"
    AddressDistrict = AddressDistrict & "411300,411322,方城县,33.255138,113.010933;"
    AddressDistrict = AddressDistrict & "411300,411323,西峡县,33.302981,111.485772;"
    AddressDistrict = AddressDistrict & "411300,411324,镇平县,33.036651,112.232722;"
    AddressDistrict = AddressDistrict & "411300,411325,内乡县,33.046358,111.843801;"
    AddressDistrict = AddressDistrict & "411300,411326,淅川县,33.136106,111.489026;"
    AddressDistrict = AddressDistrict & "411300,411327,社旗县,33.056126,112.938279;"
    AddressDistrict = AddressDistrict & "411300,411328,唐河县,32.687892,112.838492;"
    AddressDistrict = AddressDistrict & "411300,411329,新野县,32.524006,112.365624;"
    AddressDistrict = AddressDistrict & "411300,411330,桐柏县,32.367153,113.406059;"
    AddressDistrict = AddressDistrict & "411300,411381,邓州市,32.681642,112.092716;"
    AddressDistrict = AddressDistrict & "411400,411402,梁园区,34.436553,115.65459;"
    AddressDistrict = AddressDistrict & "411400,411403,睢阳区,34.390536,115.653813;"
    AddressDistrict = AddressDistrict & "411400,411421,民权县,34.648455,115.148146;"
    AddressDistrict = AddressDistrict & "411400,411422,睢县,34.428433,115.070109;"
    AddressDistrict = AddressDistrict & "411400,411423,宁陵县,34.449299,115.320055;"
    AddressDistrict = AddressDistrict & "411400,411424,柘城县,34.075277,115.307433;"
    AddressDistrict = AddressDistrict & "411400,411425,虞城县,34.399634,115.863811;"
    AddressDistrict = AddressDistrict & "411400,411426,夏邑县,34.240894,116.13989;"
    AddressDistrict = AddressDistrict & "411400,411481,永城市,33.931318,116.449672;"
    AddressDistrict = AddressDistrict & "411500,411502,浉河区,32.123274,114.075031;"
    AddressDistrict = AddressDistrict & "411500,411503,平桥区,32.098395,114.126027;"
    AddressDistrict = AddressDistrict & "411500,411521,罗山县,32.203206,114.533414;"
    AddressDistrict = AddressDistrict & "411500,411522,光山县,32.010398,114.903577;"
    AddressDistrict = AddressDistrict & "411500,411523,新县,31.63515,114.87705;"
    AddressDistrict = AddressDistrict & "411500,411524,商城县,31.799982,115.406297;"
    AddressDistrict = AddressDistrict & "411500,411525,固始县,32.183074,115.667328;"
    AddressDistrict = AddressDistrict & "411500,411526,潢川县,32.134024,115.050123;"
    AddressDistrict = AddressDistrict & "411500,411527,淮滨县,32.452639,115.415451;"
    AddressDistrict = AddressDistrict & "411500,411528,息县,32.344744,114.740713;"
    AddressDistrict = AddressDistrict & "411600,411602,川汇区,33.614836,114.652136;"
    AddressDistrict = AddressDistrict & "411600,411603,淮阳区,33.732547,114.870166;"
    AddressDistrict = AddressDistrict & "411600,411621,扶沟县,34.054061,114.392008;"
    AddressDistrict = AddressDistrict & "411600,411622,西华县,33.784378,114.530067;"
    AddressDistrict = AddressDistrict & "411600,411623,商水县,33.543845,114.60927;"
    AddressDistrict = AddressDistrict & "411600,411624,沈丘县,33.395514,115.078375;"
    AddressDistrict = AddressDistrict & "411600,411625,郸城县,33.643852,115.189;"
    AddressDistrict = AddressDistrict & "411600,411627,太康县,34.065312,114.853834;"
    AddressDistrict = AddressDistrict & "411600,411628,鹿邑县,33.861067,115.486386;"
    AddressDistrict = AddressDistrict & "411600,411681,项城市,33.443085,114.899521;"
    AddressDistrict = AddressDistrict & "411700,411702,驿城区,32.977559,114.029149;"
    AddressDistrict = AddressDistrict & "411700,411721,西平县,33.382315,114.026864;"
    AddressDistrict = AddressDistrict & "411700,411722,上蔡县,33.264719,114.266892;"
    AddressDistrict = AddressDistrict & "411700,411723,平舆县,32.955626,114.637105;"
    AddressDistrict = AddressDistrict & "411700,411724,正阳县,32.601826,114.38948;"
    AddressDistrict = AddressDistrict & "411700,411725,确山县,32.801538,114.026679;"
    AddressDistrict = AddressDistrict & "411700,411726,泌阳县,32.725129,113.32605;"
    AddressDistrict = AddressDistrict & "411700,411727,汝南县,33.004535,114.359495;"
    AddressDistrict = AddressDistrict & "411700,411728,遂平县,33.14698,114.00371;"
    AddressDistrict = AddressDistrict & "411700,411729,新蔡县,32.749948,114.975246;"
    AddressDistrict = AddressDistrict & "420100,420102,江岸区,30.594911,114.30304;"
    AddressDistrict = AddressDistrict & "420100,420103,江汉区,30.578771,114.283109;"
    AddressDistrict = AddressDistrict & "420100,420104,硚口区,30.57061,114.264568;"
    AddressDistrict = AddressDistrict & "420100,420105,汉阳区,30.549326,114.265807;"
    AddressDistrict = AddressDistrict & "420100,420106,武昌区,30.546536,114.307344;"
    AddressDistrict = AddressDistrict & "420100,420107,青山区,30.634215,114.39707;"
    AddressDistrict = AddressDistrict & "420100,420111,洪山区,30.504259,114.400718;"
    AddressDistrict = AddressDistrict & "420100,420112,东西湖区,30.622467,114.142483;"
    AddressDistrict = AddressDistrict & "420100,420113,汉南区,30.309637,114.08124;"
    AddressDistrict = AddressDistrict & "420100,420114,蔡甸区,30.582186,114.029341;"
    AddressDistrict = AddressDistrict & "420100,420115,江夏区,30.349045,114.313961;"
    AddressDistrict = AddressDistrict & "420100,420116,黄陂区,30.874155,114.374025;"
    AddressDistrict = AddressDistrict & "420100,420117,新洲区,30.842149,114.802108;"
    AddressDistrict = AddressDistrict & "420200,420202,黄石港区,30.212086,115.090164;"
    AddressDistrict = AddressDistrict & "420200,420203,西塞山区,30.205365,115.093354;"
    AddressDistrict = AddressDistrict & "420200,420204,下陆区,30.177845,114.975755;"
    AddressDistrict = AddressDistrict & "420200,420205,铁山区,30.20601,114.901366;"
    AddressDistrict = AddressDistrict & "420200,420222,阳新县,29.841572,115.212883;"
    AddressDistrict = AddressDistrict & "420200,420281,大冶市,30.098804,114.974842;"
    AddressDistrict = AddressDistrict & "420300,420302,茅箭区,32.644463,110.78621;"
    AddressDistrict = AddressDistrict & "420300,420303,张湾区,32.652516,110.772365;"
    AddressDistrict = AddressDistrict & "420300,420304,郧阳区,32.838267,110.812099;"
    AddressDistrict = AddressDistrict & "420300,420322,郧西县,32.991457,110.426472;"
    AddressDistrict = AddressDistrict & "420300,420323,竹山县,32.22586,110.2296;"
    AddressDistrict = AddressDistrict & "420300,420324,竹溪县,32.315342,109.717196;"
    AddressDistrict = AddressDistrict & "420300,420325,房县,32.055002,110.741966;"
    AddressDistrict = AddressDistrict & "420300,420381,丹江口市,32.538839,111.513793;"
    AddressDistrict = AddressDistrict & "420500,420502,西陵区,30.702476,111.295468;"
    AddressDistrict = AddressDistrict & "420500,420503,伍家岗区,30.679053,111.307215;"
    AddressDistrict = AddressDistrict & "420500,420504,点军区,30.692322,111.268163;"
    AddressDistrict = AddressDistrict & "420500,420505,猇亭区,30.530744,111.427642;"
    AddressDistrict = AddressDistrict & "420500,420506,夷陵区,30.770199,111.326747;"
    AddressDistrict = AddressDistrict & "420500,420525,远安县,31.059626,111.64331;"
    AddressDistrict = AddressDistrict & "420500,420526,兴山县,31.34795,110.754499;"
    AddressDistrict = AddressDistrict & "420500,420527,秭归县,30.823908,110.976785;"
    AddressDistrict = AddressDistrict & "420500,420528,长阳土家族自治县,30.466534,111.198475;"
    AddressDistrict = AddressDistrict & "420500,420529,五峰土家族自治县,30.199252,110.674938;"
    AddressDistrict = AddressDistrict & "420500,420581,宜都市,30.387234,111.454367;"
    AddressDistrict = AddressDistrict & "420500,420582,当阳市,30.824492,111.793419;"
    AddressDistrict = AddressDistrict & "420500,420583,枝江市,30.425364,111.751799;"
    AddressDistrict = AddressDistrict & "420600,420602,襄城区,32.015088,112.150327;"
    AddressDistrict = AddressDistrict & "420600,420606,樊城区,32.058589,112.13957;"
    AddressDistrict = AddressDistrict & "420600,420607,襄州区,32.085517,112.197378;"
    AddressDistrict = AddressDistrict & "420600,420624,南漳县,31.77692,111.844424;"
    AddressDistrict = AddressDistrict & "420600,420625,谷城县,32.262676,111.640147;"
    AddressDistrict = AddressDistrict & "420600,420626,保康县,31.873507,111.262235;"
    AddressDistrict = AddressDistrict & "420600,420682,老河口市,32.385438,111.675732;"
    AddressDistrict = AddressDistrict & "420600,420683,枣阳市,32.123083,112.765268;"
    AddressDistrict = AddressDistrict & "420600,420684,宜城市,31.709203,112.261441;"
    AddressDistrict = AddressDistrict & "420700,420702,梁子湖区,30.098191,114.681967;"
    AddressDistrict = AddressDistrict & "420700,420703,华容区,30.534468,114.74148;"
    AddressDistrict = AddressDistrict & "420700,420704,鄂城区,30.39669,114.890012;"
    AddressDistrict = AddressDistrict & "420800,420802,东宝区,31.033461,112.204804;"
    AddressDistrict = AddressDistrict & "420800,420804,掇刀区,30.980798,112.198413;"
    AddressDistrict = AddressDistrict & "420800,420822,沙洋县,30.70359,112.595218;"
    AddressDistrict = AddressDistrict & "420800,420881,钟祥市,31.165573,112.587267;"
    AddressDistrict = AddressDistrict & "420800,420882,京山市,31.022457,113.114595;"
    AddressDistrict = AddressDistrict & "420900,420902,孝南区,30.925966,113.925849;"
    AddressDistrict = AddressDistrict & "420900,420921,孝昌县,31.251618,113.988964;"
    AddressDistrict = AddressDistrict & "420900,420922,大悟县,31.565483,114.126249;"
    AddressDistrict = AddressDistrict & "420900,420923,云梦县,31.021691,113.750616;"
    AddressDistrict = AddressDistrict & "420900,420981,应城市,30.939038,113.573842;"
    AddressDistrict = AddressDistrict & "420900,420982,安陆市,31.26174,113.690401;"
    AddressDistrict = AddressDistrict & "420900,420984,汉川市,30.652165,113.835301;"
    AddressDistrict = AddressDistrict & "421000,421002,沙市区,30.315895,112.257433;"
    AddressDistrict = AddressDistrict & "421000,421003,荆州区,30.350674,112.195354;"
    AddressDistrict = AddressDistrict & "421000,421022,公安县,30.059065,112.230179;"
    AddressDistrict = AddressDistrict & "421000,421023,监利市,29.820079,112.904344;"
    AddressDistrict = AddressDistrict & "421000,421024,江陵县,30.033919,112.41735;"
    AddressDistrict = AddressDistrict & "421000,421081,石首市,29.716437,112.40887;"
    AddressDistrict = AddressDistrict & "421000,421083,洪湖市,29.81297,113.470304;"
    AddressDistrict = AddressDistrict & "421000,421087,松滋市,30.176037,111.77818;"
    AddressDistrict = AddressDistrict & "421100,421102,黄州区,30.447435,114.878934;"
    AddressDistrict = AddressDistrict & "421100,421121,团风县,30.63569,114.872029;"
    AddressDistrict = AddressDistrict & "421100,421122,红安县,31.284777,114.615095;"
    AddressDistrict = AddressDistrict & "421100,421123,罗田县,30.781679,115.398984;"
    AddressDistrict = AddressDistrict & "421100,421124,英山县,30.735794,115.67753;"
    AddressDistrict = AddressDistrict & "421100,421125,浠水县,30.454837,115.26344;"
    AddressDistrict = AddressDistrict & "421100,421126,蕲春县,30.234927,115.433964;"
    AddressDistrict = AddressDistrict & "421100,421127,黄梅县,30.075113,115.942548;"
    AddressDistrict = AddressDistrict & "421100,421181,麻城市,31.177906,115.02541;"
    AddressDistrict = AddressDistrict & "421100,421182,武穴市,29.849342,115.56242;"
    AddressDistrict = AddressDistrict & "421200,421202,咸安区,29.824716,114.333894;"
    AddressDistrict = AddressDistrict & "421200,421221,嘉鱼县,29.973363,113.921547;"
    AddressDistrict = AddressDistrict & "421200,421222,通城县,29.246076,113.814131;"
    AddressDistrict = AddressDistrict & "421200,421223,崇阳县,29.54101,114.049958;"
    AddressDistrict = AddressDistrict & "421200,421224,通山县,29.604455,114.493163;"
    AddressDistrict = AddressDistrict & "421200,421281,赤壁市,29.716879,113.88366;"
    AddressDistrict = AddressDistrict & "421300,421303,曾都区,31.717521,113.374519;"
    AddressDistrict = AddressDistrict & "421300,421321,随县,31.854246,113.301384;"
    AddressDistrict = AddressDistrict & "421300,421381,广水市,31.617731,113.826601;"
    AddressDistrict = AddressDistrict & "422800,422801,恩施市,30.282406,109.486761;"
    AddressDistrict = AddressDistrict & "422800,422802,利川市,30.294247,108.943491;"
    AddressDistrict = AddressDistrict & "422800,422822,建始县,30.601632,109.723822;"
    AddressDistrict = AddressDistrict & "422800,422823,巴东县,31.041403,110.336665;"
    AddressDistrict = AddressDistrict & "422800,422825,宣恩县,29.98867,109.482819;"
    AddressDistrict = AddressDistrict & "422800,422826,咸丰县,29.678967,109.15041;"
    AddressDistrict = AddressDistrict & "422800,422827,来凤县,29.506945,109.408328;"
    AddressDistrict = AddressDistrict & "422800,422828,鹤峰县,29.887298,110.033699;"
    AddressDistrict = AddressDistrict & "430100,430102,芙蓉区,28.193106,112.988094;"
    AddressDistrict = AddressDistrict & "430100,430103,天心区,28.192375,112.97307;"
    AddressDistrict = AddressDistrict & "430100,430104,岳麓区,28.213044,112.911591;"
    AddressDistrict = AddressDistrict & "430100,430105,开福区,28.201336,112.985525;"
    AddressDistrict = AddressDistrict & "430100,430111,雨花区,28.109937,113.016337;"
    AddressDistrict = AddressDistrict & "430100,430112,望城区,28.347458,112.819549;"
    AddressDistrict = AddressDistrict & "430100,430121,长沙县,28.237888,113.080098;"
    AddressDistrict = AddressDistrict & "430100,430181,浏阳市,28.141112,113.633301;"
    AddressDistrict = AddressDistrict & "430100,430182,宁乡市,28.253928,112.553182;"
    AddressDistrict = AddressDistrict & "430200,430202,荷塘区,27.833036,113.162548;"
    AddressDistrict = AddressDistrict & "430200,430203,芦淞区,27.827246,113.155169;"
    AddressDistrict = AddressDistrict & "430200,430204,石峰区,27.871945,113.11295;"
    AddressDistrict = AddressDistrict & "430200,430211,天元区,27.826909,113.136252;"
    AddressDistrict = AddressDistrict & "430200,430212,渌口区,27.705844,113.146175;"
    AddressDistrict = AddressDistrict & "430200,430223,攸县,27.000071,113.345774;"
    AddressDistrict = AddressDistrict & "430200,430224,茶陵县,26.789534,113.546509;"
    AddressDistrict = AddressDistrict & "430200,430225,炎陵县,26.489459,113.776884;"
    AddressDistrict = AddressDistrict & "430200,430281,醴陵市,27.657873,113.507157;"
    AddressDistrict = AddressDistrict & "430300,430302,雨湖区,27.86077,112.907427;"
    AddressDistrict = AddressDistrict & "430300,430304,岳塘区,27.828854,112.927707;"
    AddressDistrict = AddressDistrict & "430300,430321,湘潭县,27.778601,112.952829;"
    AddressDistrict = AddressDistrict & "430300,430381,湘乡市,27.734918,112.525217;"
    AddressDistrict = AddressDistrict & "430300,430382,韶山市,27.922682,112.52848;"
    AddressDistrict = AddressDistrict & "430400,430405,珠晖区,26.891063,112.626324;"
    AddressDistrict = AddressDistrict & "430400,430406,雁峰区,26.893694,112.612241;"
    AddressDistrict = AddressDistrict & "430400,430407,石鼓区,26.903908,112.607635;"
    AddressDistrict = AddressDistrict & "430400,430408,蒸湘区,26.89087,112.570608;"
    AddressDistrict = AddressDistrict & "430400,430412,南岳区,27.240536,112.734147;"
    AddressDistrict = AddressDistrict & "430400,430421,衡阳县,26.962388,112.379643;"
    AddressDistrict = AddressDistrict & "430400,430422,衡南县,26.739973,112.677459;"
    AddressDistrict = AddressDistrict & "430400,430423,衡山县,27.234808,112.86971;"
    AddressDistrict = AddressDistrict & "430400,430424,衡东县,27.083531,112.950412;"
    AddressDistrict = AddressDistrict & "430400,430426,祁东县,26.787109,112.111192;"
    AddressDistrict = AddressDistrict & "430400,430481,耒阳市,26.414162,112.847215;"
    AddressDistrict = AddressDistrict & "430400,430482,常宁市,26.406773,112.396821;"
    AddressDistrict = AddressDistrict & "430500,430502,双清区,27.240001,111.479756;"
    AddressDistrict = AddressDistrict & "430500,430503,大祥区,27.233593,111.462968;"
    AddressDistrict = AddressDistrict & "430500,430511,北塔区,27.245688,111.452315;"
    AddressDistrict = AddressDistrict & "430500,430522,新邵县,27.311429,111.459762;"
    AddressDistrict = AddressDistrict & "430500,430523,邵阳县,26.989713,111.2757;"
    AddressDistrict = AddressDistrict & "430500,430524,隆回县,27.116002,111.038785;"
    AddressDistrict = AddressDistrict & "430500,430525,洞口县,27.062286,110.579212;"
    AddressDistrict = AddressDistrict & "430500,430527,绥宁县,26.580622,110.155075;"
    AddressDistrict = AddressDistrict & "430500,430528,新宁县,26.438912,110.859115;"
    AddressDistrict = AddressDistrict & "430500,430529,城步苗族自治县,26.363575,110.313226;"
    AddressDistrict = AddressDistrict & "430500,430581,武冈市,26.732086,110.636804;"
    AddressDistrict = AddressDistrict & "430500,430582,邵东市,27.257273,111.743168;"
    AddressDistrict = AddressDistrict & "430600,430602,岳阳楼区,29.366784,113.120751;"
    AddressDistrict = AddressDistrict & "430600,430603,云溪区,29.473395,113.27387;"
    AddressDistrict = AddressDistrict & "430600,430611,君山区,29.438062,113.004082;"
    AddressDistrict = AddressDistrict & "430600,430621,岳阳县,29.144843,113.116073;"
    AddressDistrict = AddressDistrict & "430600,430623,华容县,29.524107,112.559369;"
    AddressDistrict = AddressDistrict & "430600,430624,湘阴县,28.677498,112.889748;"
    AddressDistrict = AddressDistrict & "430600,430626,平江县,28.701523,113.593751;"
    AddressDistrict = AddressDistrict & "430600,430681,汨罗市,28.803149,113.079419;"
    AddressDistrict = AddressDistrict & "430600,430682,临湘市,29.471594,113.450809;"
    AddressDistrict = AddressDistrict & "430700,430702,武陵区,29.040477,111.690718;"
    AddressDistrict = AddressDistrict & "430700,430703,鼎城区,29.014426,111.685327;"
    AddressDistrict = AddressDistrict & "430700,430721,安乡县,29.414483,112.172289;"
    AddressDistrict = AddressDistrict & "430700,430722,汉寿县,28.907319,111.968506;"
    AddressDistrict = AddressDistrict & "430700,430723,澧县,29.64264,111.761682;"
    AddressDistrict = AddressDistrict & "430700,430724,临澧县,29.443217,111.645602;"
    AddressDistrict = AddressDistrict & "430700,430725,桃源县,28.902734,111.484503;"
    AddressDistrict = AddressDistrict & "430700,430726,石门县,29.584703,111.379087;"
    AddressDistrict = AddressDistrict & "430700,430781,津市市,29.630867,111.879609;"
    AddressDistrict = AddressDistrict & "430800,430802,永定区,29.125961,110.484559;"
    AddressDistrict = AddressDistrict & "430800,430811,武陵源区,29.347827,110.54758;"
    AddressDistrict = AddressDistrict & "430800,430821,慈利县,29.423876,111.132702;"
    AddressDistrict = AddressDistrict & "430800,430822,桑植县,29.399939,110.164039;"
    AddressDistrict = AddressDistrict & "430900,430902,资阳区,28.592771,112.33084;"
    AddressDistrict = AddressDistrict & "430900,430903,赫山区,28.568327,112.360946;"
    AddressDistrict = AddressDistrict & "430900,430921,南县,29.372181,112.410399;"
    AddressDistrict = AddressDistrict & "430900,430922,桃江县,28.520993,112.139732;"
    AddressDistrict = AddressDistrict & "430900,430923,安化县,28.377421,111.221824;"
    AddressDistrict = AddressDistrict & "430900,430981,沅江市,28.839713,112.361088;"
    AddressDistrict = AddressDistrict & "431000,431002,北湖区,25.792628,113.032208;"
    AddressDistrict = AddressDistrict & "431000,431003,苏仙区,25.793157,113.038698;"
    AddressDistrict = AddressDistrict & "431000,431021,桂阳县,25.737447,112.734466;"
    AddressDistrict = AddressDistrict & "431000,431022,宜章县,25.394345,112.947884;"
    AddressDistrict = AddressDistrict & "431000,431023,永兴县,26.129392,113.114819;"
    AddressDistrict = AddressDistrict & "431000,431024,嘉禾县,25.587309,112.370618;"
    AddressDistrict = AddressDistrict & "431000,431025,临武县,25.279119,112.564589;"
    AddressDistrict = AddressDistrict & "431000,431026,汝城县,25.553759,113.685686;"
    AddressDistrict = AddressDistrict & "431000,431027,桂东县,26.073917,113.945879;"
    AddressDistrict = AddressDistrict & "431000,431028,安仁县,26.708625,113.27217;"
    AddressDistrict = AddressDistrict & "431000,431081,资兴市,25.974152,113.23682;"
    AddressDistrict = AddressDistrict & "431100,431102,零陵区,26.223347,111.626348;"
    AddressDistrict = AddressDistrict & "431100,431103,冷水滩区,26.434364,111.607156;"
    AddressDistrict = AddressDistrict & "431100,431121,祁阳市,26.585929,111.85734;"
    AddressDistrict = AddressDistrict & "431100,431122,东安县,26.397278,111.313035;"
    AddressDistrict = AddressDistrict & "431100,431123,双牌县,25.959397,111.662146;"
    AddressDistrict = AddressDistrict & "431100,431124,道县,25.518444,111.591614;"
    AddressDistrict = AddressDistrict & "431100,431125,江永县,25.268154,111.346803;"
    AddressDistrict = AddressDistrict & "431100,431126,宁远县,25.584112,111.944529;"
    AddressDistrict = AddressDistrict & "431100,431127,蓝山县,25.375255,112.194195;"
    AddressDistrict = AddressDistrict & "431100,431128,新田县,25.906927,112.220341;"
    AddressDistrict = AddressDistrict & "431100,431129,江华瑶族自治县,25.182596,111.577276;"
    AddressDistrict = AddressDistrict & "431200,431202,鹤城区,27.548474,109.982242;"
    AddressDistrict = AddressDistrict & "431200,431221,中方县,27.43736,109.948061;"
    AddressDistrict = AddressDistrict & "431200,431222,沅陵县,28.455554,110.399161;"
    AddressDistrict = AddressDistrict & "431200,431223,辰溪县,28.005474,110.196953;"
    AddressDistrict = AddressDistrict & "431200,431224,溆浦县,27.903802,110.593373;"
    AddressDistrict = AddressDistrict & "431200,431225,会同县,26.870789,109.720785;"
    AddressDistrict = AddressDistrict & "431200,431226,麻阳苗族自治县,27.865991,109.802807;"
    AddressDistrict = AddressDistrict & "431200,431227,新晃侗族自治县,27.359897,109.174443;"
    AddressDistrict = AddressDistrict & "431200,431228,芷江侗族自治县,27.437996,109.687777;"
    AddressDistrict = AddressDistrict & "431200,431229,靖州苗族侗族自治县,26.573511,109.691159;"
    AddressDistrict = AddressDistrict & "431200,431230,通道侗族自治县,26.158349,109.783359;"
    AddressDistrict = AddressDistrict & "431200,431281,洪江市,27.201876,109.831765;"
    AddressDistrict = AddressDistrict & "431300,431302,娄星区,27.726643,112.008486;"
    AddressDistrict = AddressDistrict & "431300,431321,双峰县,27.459126,112.198245;"
    AddressDistrict = AddressDistrict & "431300,431322,新化县,27.737456,111.306747;"
    AddressDistrict = AddressDistrict & "431300,431381,冷水江市,27.685759,111.434674;"
    AddressDistrict = AddressDistrict & "431300,431382,涟源市,27.692301,111.670847;"
    AddressDistrict = AddressDistrict & "433100,433101,吉首市,28.314827,109.738273;"
    AddressDistrict = AddressDistrict & "433100,433122,泸溪县,28.214516,110.214428;"
    AddressDistrict = AddressDistrict & "433100,433123,凤凰县,27.948308,109.599191;"
    AddressDistrict = AddressDistrict & "433100,433124,花垣县,28.581352,109.479063;"
    AddressDistrict = AddressDistrict & "433100,433125,保靖县,28.709605,109.651445;"
    AddressDistrict = AddressDistrict & "433100,433126,古丈县,28.616973,109.949592;"
    AddressDistrict = AddressDistrict & "433100,433127,永顺县,28.998068,109.853292;"
    AddressDistrict = AddressDistrict & "433100,433130,龙山县,29.453438,109.441189;"
    AddressDistrict = AddressDistrict & "440100,440103,荔湾区,23.124943,113.243038;"
    AddressDistrict = AddressDistrict & "440100,440104,越秀区,23.125624,113.280714;"
    AddressDistrict = AddressDistrict & "440100,440105,海珠区,23.103131,113.262008;"
    AddressDistrict = AddressDistrict & "440100,440106,天河区,23.13559,113.335367;"
    AddressDistrict = AddressDistrict & "440100,440111,白云区,23.162281,113.262831;"
    AddressDistrict = AddressDistrict & "440100,440112,黄埔区,23.103239,113.450761;"
    AddressDistrict = AddressDistrict & "440100,440113,番禺区,22.938582,113.364619;"
    AddressDistrict = AddressDistrict & "440100,440114,花都区,23.39205,113.211184;"
    AddressDistrict = AddressDistrict & "440100,440115,南沙区,22.794531,113.53738;"
    AddressDistrict = AddressDistrict & "440100,440117,从化区,23.545283,113.587386;"
    AddressDistrict = AddressDistrict & "440100,440118,增城区,23.290497,113.829579;"
    AddressDistrict = AddressDistrict & "440200,440203,武江区,24.80016,113.588289;"
    AddressDistrict = AddressDistrict & "440200,440204,浈江区,24.803977,113.599224;"
    AddressDistrict = AddressDistrict & "440200,440205,曲江区,24.680195,113.605582;"
    AddressDistrict = AddressDistrict & "440200,440222,始兴县,24.948364,114.067205;"
    AddressDistrict = AddressDistrict & "440200,440224,仁化县,25.088226,113.748627;"
    AddressDistrict = AddressDistrict & "440200,440229,翁源县,24.353887,114.131289;"
    AddressDistrict = AddressDistrict & "440200,440232,乳源瑶族自治县,24.776109,113.278417;"
    AddressDistrict = AddressDistrict & "440200,440233,新丰县,24.055412,114.207034;"
    AddressDistrict = AddressDistrict & "440200,440281,乐昌市,25.128445,113.352413;"
    AddressDistrict = AddressDistrict & "440200,440282,南雄市,25.115328,114.311231;"
    AddressDistrict = AddressDistrict & "440300,440303,罗湖区,22.555341,114.123885;"
    AddressDistrict = AddressDistrict & "440300,440304,福田区,22.541009,114.05096;"
    AddressDistrict = AddressDistrict & "440300,440305,南山区,22.531221,113.92943;"
    AddressDistrict = AddressDistrict & "440300,440306,宝安区,22.754741,113.828671;"
    AddressDistrict = AddressDistrict & "440300,440307,龙岗区,22.721511,114.251372;"
    AddressDistrict = AddressDistrict & "440300,440308,盐田区,22.555069,114.235366;"
    AddressDistrict = AddressDistrict & "440300,440309,龙华区,22.691963,114.044346;"
    AddressDistrict = AddressDistrict & "440300,440310,坪山区,22.69423,114.338441;"
    AddressDistrict = AddressDistrict & "440300,440311,光明区,22.748816,113.935895;"
    AddressDistrict = AddressDistrict & "440400,440402,香洲区,22.271249,113.55027;"
    AddressDistrict = AddressDistrict & "440400,440403,斗门区,22.209117,113.297739;"
    AddressDistrict = AddressDistrict & "440400,440404,金湾区,22.139122,113.345071;"
    AddressDistrict = AddressDistrict & "440500,440507,龙湖区,23.373754,116.732015;"
    AddressDistrict = AddressDistrict & "440500,440511,金平区,23.367071,116.703583;"
    AddressDistrict = AddressDistrict & "440500,440512,濠江区,23.279345,116.729528;"
    AddressDistrict = AddressDistrict & "440500,440513,潮阳区,23.262336,116.602602;"
    AddressDistrict = AddressDistrict & "440500,440514,潮南区,23.249798,116.423607;"
    AddressDistrict = AddressDistrict & "440500,440515,澄海区,23.46844,116.76336;"
    AddressDistrict = AddressDistrict & "440500,440523,南澳县,23.419562,117.027105;"
    AddressDistrict = AddressDistrict & "440600,440604,禅城区,23.019643,113.112414;"
    AddressDistrict = AddressDistrict & "440600,440605,南海区,23.031562,113.145577;"
    AddressDistrict = AddressDistrict & "440600,440606,顺德区,22.75851,113.281826;"
    AddressDistrict = AddressDistrict & "440600,440607,三水区,23.16504,112.899414;"
    AddressDistrict = AddressDistrict & "440600,440608,高明区,22.893855,112.882123;"
    AddressDistrict = AddressDistrict & "440700,440703,蓬江区,22.59677,113.07859;"
    AddressDistrict = AddressDistrict & "440700,440704,江海区,22.572211,113.120601;"
    AddressDistrict = AddressDistrict & "440700,440705,新会区,22.520247,113.038584;"
    AddressDistrict = AddressDistrict & "440700,440781,台山市,22.250713,112.793414;"
    AddressDistrict = AddressDistrict & "440700,440783,开平市,22.366286,112.692262;"
    AddressDistrict = AddressDistrict & "440700,440784,鹤山市,22.768104,112.961795;"
    AddressDistrict = AddressDistrict & "440700,440785,恩平市,22.182956,112.314051;"
    AddressDistrict = AddressDistrict & "440800,440802,赤坎区,21.273365,110.361634;"
    AddressDistrict = AddressDistrict & "440800,440803,霞山区,21.194229,110.406382;"
    AddressDistrict = AddressDistrict & "440800,440804,坡头区,21.24441,110.455632;"
    AddressDistrict = AddressDistrict & "440800,440811,麻章区,21.265997,110.329167;"
    AddressDistrict = AddressDistrict & "440800,440823,遂溪县,21.376915,110.255321;"
    AddressDistrict = AddressDistrict & "440800,440825,徐闻县,20.326083,110.175718;"
    AddressDistrict = AddressDistrict & "440800,440881,廉江市,21.611281,110.284961;"
    AddressDistrict = AddressDistrict & "440800,440882,雷州市,20.908523,110.088275;"
    AddressDistrict = AddressDistrict & "440800,440883,吴川市,21.428453,110.780508;"
    AddressDistrict = AddressDistrict & "440900,440902,茂南区,21.660425,110.920542;"
    AddressDistrict = AddressDistrict & "440900,440904,电白区,21.507219,111.007264;"
    AddressDistrict = AddressDistrict & "440900,440981,高州市,21.915153,110.853251;"
    AddressDistrict = AddressDistrict & "440900,440982,化州市,21.654953,110.63839;"
    AddressDistrict = AddressDistrict & "440900,440983,信宜市,22.352681,110.941656;"
    AddressDistrict = AddressDistrict & "441200,441202,端州区,23.052662,112.472329;"
    AddressDistrict = AddressDistrict & "441200,441203,鼎湖区,23.155822,112.565249;"
    AddressDistrict = AddressDistrict & "441200,441204,高要区,23.027694,112.460846;"
    AddressDistrict = AddressDistrict & "441200,441223,广宁县,23.631486,112.440419;"
    AddressDistrict = AddressDistrict & "441200,441224,怀集县,23.913072,112.182466;"
    AddressDistrict = AddressDistrict & "441200,441225,封开县,23.434731,111.502973;"
    AddressDistrict = AddressDistrict & "441200,441226,德庆县,23.141711,111.78156;"
    AddressDistrict = AddressDistrict & "441200,441284,四会市,23.340324,112.695028;"
    AddressDistrict = AddressDistrict & "441300,441302,惠城区,23.079883,114.413978;"
    AddressDistrict = AddressDistrict & "441300,441303,惠阳区,22.78851,114.469444;"
    AddressDistrict = AddressDistrict & "441300,441322,博罗县,23.167575,114.284254;"
    AddressDistrict = AddressDistrict & "441300,441323,惠东县,22.983036,114.723092;"
    AddressDistrict = AddressDistrict & "441300,441324,龙门县,23.723894,114.259986;"
    AddressDistrict = AddressDistrict & "441400,441402,梅江区,24.302593,116.12116;"
    AddressDistrict = AddressDistrict & "441400,441403,梅县区,24.267825,116.083482;"
    AddressDistrict = AddressDistrict & "441400,441422,大埔县,24.351587,116.69552;"
    AddressDistrict = AddressDistrict & "441400,441423,丰顺县,23.752771,116.184419;"
    AddressDistrict = AddressDistrict & "441400,441424,五华县,23.925424,115.775004;"
    AddressDistrict = AddressDistrict & "441400,441426,平远县,24.569651,115.891729;"
    AddressDistrict = AddressDistrict & "441400,441427,蕉岭县,24.653313,116.170531;"
    AddressDistrict = AddressDistrict & "441400,441481,兴宁市,24.138077,115.731648;"
    AddressDistrict = AddressDistrict & "441500,441502,城区,22.776227,115.363667;"
    AddressDistrict = AddressDistrict & "441500,441521,海丰县,22.971042,115.337324;"
    AddressDistrict = AddressDistrict & "441500,441523,陆河县,23.302682,115.657565;"
    AddressDistrict = AddressDistrict & "441500,441581,陆丰市,22.946104,115.644203;"
    AddressDistrict = AddressDistrict & "441600,441602,源城区,23.746255,114.696828;"
    AddressDistrict = AddressDistrict & "441600,441621,紫金县,23.633744,115.184383;"
    AddressDistrict = AddressDistrict & "441600,441622,龙川县,24.101174,115.256415;"
    AddressDistrict = AddressDistrict & "441600,441623,连平县,24.364227,114.495952;"
    AddressDistrict = AddressDistrict & "441600,441624,和平县,24.44318,114.941473;"
    AddressDistrict = AddressDistrict & "441600,441625,东源县,23.789093,114.742711;"
    AddressDistrict = AddressDistrict & "441700,441702,江城区,21.859182,111.968909;"
    AddressDistrict = AddressDistrict & "441700,441704,阳东区,21.864728,112.011267;"
    AddressDistrict = AddressDistrict & "441700,441721,阳西县,21.75367,111.617556;"
    AddressDistrict = AddressDistrict & "441700,441781,阳春市,22.169598,111.7905;"
    AddressDistrict = AddressDistrict & "441800,441802,清城区,23.688976,113.048698;"
    AddressDistrict = AddressDistrict & "441800,441803,清新区,23.736949,113.015203;"
    AddressDistrict = AddressDistrict & "441800,441821,佛冈县,23.866739,113.534094;"
    AddressDistrict = AddressDistrict & "441800,441823,阳山县,24.470286,112.634019;"
    AddressDistrict = AddressDistrict & "441800,441825,连山壮族瑶族自治县,24.567271,112.086555;"
    AddressDistrict = AddressDistrict & "441800,441826,连南瑶族自治县,24.719097,112.290808;"
    AddressDistrict = AddressDistrict & "441800,441881,英德市,24.18612,113.405404;"
    AddressDistrict = AddressDistrict & "441800,441882,连州市,24.783966,112.379271;"
    AddressDistrict = AddressDistrict & "445100,445102,湘桥区,23.664675,116.63365;"
    AddressDistrict = AddressDistrict & "445100,445103,潮安区,23.461012,116.67931;"
    AddressDistrict = AddressDistrict & "445100,445122,饶平县,23.668171,117.00205;"
    AddressDistrict = AddressDistrict & "445200,445202,榕城区,23.535524,116.357045;"
    AddressDistrict = AddressDistrict & "445200,445203,揭东区,23.569887,116.412947;"
    AddressDistrict = AddressDistrict & "445200,445222,揭西县,23.4273,115.838708;"
    AddressDistrict = AddressDistrict & "445200,445224,惠来县,23.029834,116.295832;"
    AddressDistrict = AddressDistrict & "445200,445281,普宁市,23.29788,116.165082;"
    AddressDistrict = AddressDistrict & "445300,445302,云城区,22.930827,112.04471;"
    AddressDistrict = AddressDistrict & "445300,445303,云安区,23.073152,112.005609;"
    AddressDistrict = AddressDistrict & "445300,445321,新兴县,22.703204,112.23083;"
    AddressDistrict = AddressDistrict & "445300,445322,郁南县,23.237709,111.535921;"
    AddressDistrict = AddressDistrict & "445300,445381,罗定市,22.765415,111.578201;"
    AddressDistrict = AddressDistrict & "450100,450102,兴宁区,22.819511,108.320189;"
    AddressDistrict = AddressDistrict & "450100,450103,青秀区,22.816614,108.346113;"
    AddressDistrict = AddressDistrict & "450100,450105,江南区,22.799593,108.310478;"
    AddressDistrict = AddressDistrict & "450100,450107,西乡塘区,22.832779,108.306903;"
    AddressDistrict = AddressDistrict & "450100,450108,良庆区,22.75909,108.322102;"
    AddressDistrict = AddressDistrict & "450100,450109,邕宁区,22.756598,108.484251;"
    AddressDistrict = AddressDistrict & "450100,450110,武鸣区,23.157163,108.280717;"
    AddressDistrict = AddressDistrict & "450100,450123,隆安县,23.174763,107.688661;"
    AddressDistrict = AddressDistrict & "450100,450124,马山县,23.711758,108.172903;"
    AddressDistrict = AddressDistrict & "450100,450125,上林县,23.431769,108.603937;"
    AddressDistrict = AddressDistrict & "450100,450126,宾阳县,23.216884,108.816735;"
    AddressDistrict = AddressDistrict & "450100,450127,横州市,22.68743,109.270987;"
    AddressDistrict = AddressDistrict & "450200,450202,城中区,24.312324,109.411749;"
    AddressDistrict = AddressDistrict & "450200,450203,鱼峰区,24.303848,109.415364;"
    AddressDistrict = AddressDistrict & "450200,450204,柳南区,24.287013,109.395936;"
    AddressDistrict = AddressDistrict & "450200,450205,柳北区,24.359145,109.406577;"
    AddressDistrict = AddressDistrict & "450200,450206,柳江区,24.257512,109.334503;"
    AddressDistrict = AddressDistrict & "450200,450222,柳城县,24.655121,109.245812;"
    AddressDistrict = AddressDistrict & "450200,450223,鹿寨县,24.483405,109.740805;"
    AddressDistrict = AddressDistrict & "450200,450224,融安县,25.214703,109.403621;"
    AddressDistrict = AddressDistrict & "450200,450225,融水苗族自治县,25.068812,109.252744;"
    AddressDistrict = AddressDistrict & "450200,450226,三江侗族自治县,25.78553,109.614846;"
    AddressDistrict = AddressDistrict & "450300,450302,秀峰区,25.278544,110.292445;"
    AddressDistrict = AddressDistrict & "450300,450303,叠彩区,25.301334,110.300783;"
    AddressDistrict = AddressDistrict & "450300,450304,象山区,25.261986,110.284882;"
    AddressDistrict = AddressDistrict & "450300,450305,七星区,25.254339,110.317577;"
    AddressDistrict = AddressDistrict & "450300,450311,雁山区,25.077646,110.305667;"
    AddressDistrict = AddressDistrict & "450300,450312,临桂区,25.246257,110.205487;"
    AddressDistrict = AddressDistrict & "450300,450321,阳朔县,24.77534,110.494699;"
    AddressDistrict = AddressDistrict & "450300,450323,灵川县,25.408541,110.325712;"
    AddressDistrict = AddressDistrict & "450300,450324,全州县,25.929897,111.072989;"
    AddressDistrict = AddressDistrict & "450300,450325,兴安县,25.609554,110.670783;"
    AddressDistrict = AddressDistrict & "450300,450326,永福县,24.986692,109.989208;"
    AddressDistrict = AddressDistrict & "450300,450327,灌阳县,25.489098,111.160248;"
    AddressDistrict = AddressDistrict & "450300,450328,龙胜各族自治县,25.796428,110.009423;"
    AddressDistrict = AddressDistrict & "450300,450329,资源县,26.0342,110.642587;"
    AddressDistrict = AddressDistrict & "450300,450330,平乐县,24.632216,110.642821;"
    AddressDistrict = AddressDistrict & "450300,450332,恭城瑶族自治县,24.833612,110.82952;"
    AddressDistrict = AddressDistrict & "450300,450381,荔浦市,24.497786,110.400149;"
    AddressDistrict = AddressDistrict & "450400,450403,万秀区,23.471318,111.315817;"
    AddressDistrict = AddressDistrict & "450400,450405,长洲区,23.4777,111.275678;"
    AddressDistrict = AddressDistrict & "450400,450406,龙圩区,23.40996,111.246035;"
    AddressDistrict = AddressDistrict & "450400,450421,苍梧县,23.845097,111.544008;"
    AddressDistrict = AddressDistrict & "450400,450422,藤县,23.373963,110.931826;"
    AddressDistrict = AddressDistrict & "450400,450423,蒙山县,24.199829,110.5226;"
    AddressDistrict = AddressDistrict & "450400,450481,岑溪市,22.918406,110.998114;"
    AddressDistrict = AddressDistrict & "450500,450502,海城区,21.468443,109.107529;"
    AddressDistrict = AddressDistrict & "450500,450503,银海区,21.444909,109.118707;"
    AddressDistrict = AddressDistrict & "450500,450512,铁山港区,21.5928,109.450573;"
    AddressDistrict = AddressDistrict & "450500,450521,合浦县,21.663554,109.200695;"
    AddressDistrict = AddressDistrict & "450600,450602,港口区,21.614406,108.346281;"
    AddressDistrict = AddressDistrict & "450600,450603,防城区,21.764758,108.358426;"
    AddressDistrict = AddressDistrict & "450600,450621,上思县,22.151423,107.982139;"
    AddressDistrict = AddressDistrict & "450600,450681,东兴市,21.541172,107.97017;"
    AddressDistrict = AddressDistrict & "450700,450702,钦南区,21.966808,108.626629;"
    AddressDistrict = AddressDistrict & "450700,450703,钦北区,22.132761,108.44911;"
    AddressDistrict = AddressDistrict & "450700,450721,灵山县,22.418041,109.293468;"
    AddressDistrict = AddressDistrict & "450700,450722,浦北县,22.268335,109.556341;"
    AddressDistrict = AddressDistrict & "450800,450802,港北区,23.107677,109.59481;"
    AddressDistrict = AddressDistrict & "450800,450803,港南区,23.067516,109.604665;"
    AddressDistrict = AddressDistrict & "450800,450804,覃塘区,23.132815,109.415697;"
    AddressDistrict = AddressDistrict & "450800,450821,平南县,23.544546,110.397485;"
    AddressDistrict = AddressDistrict & "450800,450881,桂平市,23.382473,110.074668;"
    AddressDistrict = AddressDistrict & "450900,450902,玉州区,22.632132,110.154912;"
    AddressDistrict = AddressDistrict & "450900,450903,福绵区,22.58163,110.054155;"
    AddressDistrict = AddressDistrict & "450900,450921,容县,22.856435,110.552467;"
    AddressDistrict = AddressDistrict & "450900,450922,陆川县,22.321054,110.264842;"
    AddressDistrict = AddressDistrict & "450900,450923,博白县,22.271285,109.980004;"
    AddressDistrict = AddressDistrict & "450900,450924,兴业县,22.74187,109.877768;"
    AddressDistrict = AddressDistrict & "450900,450981,北流市,22.701648,110.348052;"
    AddressDistrict = AddressDistrict & "451000,451002,右江区,23.897675,106.615727;"
    AddressDistrict = AddressDistrict & "451000,451003,田阳区,23.736079,106.904315;"
    AddressDistrict = AddressDistrict & "451000,451022,田东县,23.600444,107.12426;"
    AddressDistrict = AddressDistrict & "451000,451024,德保县,23.321464,106.618164;"
    AddressDistrict = AddressDistrict & "451000,451026,那坡县,23.400785,105.833553;"
    AddressDistrict = AddressDistrict & "451000,451027,凌云县,24.345643,106.56487;"
    AddressDistrict = AddressDistrict & "451000,451028,乐业县,24.782204,106.559638;"
    AddressDistrict = AddressDistrict & "451000,451029,田林县,24.290262,106.235047;"
    AddressDistrict = AddressDistrict & "451000,451030,西林县,24.492041,105.095025;"
    AddressDistrict = AddressDistrict & "451000,451031,隆林各族自治县,24.774318,105.342363;"
    AddressDistrict = AddressDistrict & "451000,451081,靖西市,23.134766,106.417549;"
    AddressDistrict = AddressDistrict & "451000,451082,平果市,23.320479,107.580403;"
    AddressDistrict = AddressDistrict & "451100,451102,八步区,24.412446,111.551991;"
    AddressDistrict = AddressDistrict & "451100,451103,平桂区,24.417148,111.524014;"
    AddressDistrict = AddressDistrict & "451100,451121,昭平县,24.172958,110.810865;"
    AddressDistrict = AddressDistrict & "451100,451122,钟山县,24.528566,111.303629;"
    AddressDistrict = AddressDistrict & "451100,451123,富川瑶族自治县,24.81896,111.277228;"
    AddressDistrict = AddressDistrict & "451200,451202,金城江区,24.695625,108.062131;"
    AddressDistrict = AddressDistrict & "451200,451203,宜州区,24.492193,108.653965;"
    AddressDistrict = AddressDistrict & "451200,451221,南丹县,24.983192,107.546605;"
    AddressDistrict = AddressDistrict & "451200,451222,天峨县,24.985964,107.174939;"
    AddressDistrict = AddressDistrict & "451200,451223,凤山县,24.544561,107.044592;"
    AddressDistrict = AddressDistrict & "451200,451224,东兰县,24.509367,107.373696;"
    AddressDistrict = AddressDistrict & "451200,451225,罗城仫佬族自治县,24.779327,108.902453;"
    AddressDistrict = AddressDistrict & "451200,451226,环江毛南族自治县,24.827628,108.258669;"
    AddressDistrict = AddressDistrict & "451200,451227,巴马瑶族自治县,24.139538,107.253126;"
    AddressDistrict = AddressDistrict & "451200,451228,都安瑶族自治县,23.934964,108.102761;"
    AddressDistrict = AddressDistrict & "451200,451229,大化瑶族自治县,23.739596,107.9945;"
    AddressDistrict = AddressDistrict & "451300,451302,兴宾区,23.732926,109.230541;"
    AddressDistrict = AddressDistrict & "451300,451321,忻城县,24.064779,108.667361;"
    AddressDistrict = AddressDistrict & "451300,451322,象州县,23.959824,109.684555;"
    AddressDistrict = AddressDistrict & "451300,451323,武宣县,23.604162,109.66287;"
    AddressDistrict = AddressDistrict & "451300,451324,金秀瑶族自治县,24.134941,110.188556;"
    AddressDistrict = AddressDistrict & "451300,451381,合山市,23.81311,108.88858;"
    AddressDistrict = AddressDistrict & "451400,451402,江州区,22.40469,107.354443;"
    AddressDistrict = AddressDistrict & "451400,451421,扶绥县,22.635821,107.911533;"
    AddressDistrict = AddressDistrict & "451400,451422,宁明县,22.131353,107.067616;"
    AddressDistrict = AddressDistrict & "451400,451423,龙州县,22.343716,106.857502;"
    AddressDistrict = AddressDistrict & "451400,451424,大新县,22.833369,107.200803;"
    AddressDistrict = AddressDistrict & "451400,451425,天等县,23.082484,107.142441;"
    AddressDistrict = AddressDistrict & "451400,451481,凭祥市,22.108882,106.759038;"
    AddressDistrict = AddressDistrict & "460100,460105,秀英区,20.008145,110.282393;"
    AddressDistrict = AddressDistrict & "460100,460106,龙华区,20.031026,110.330373;"
    AddressDistrict = AddressDistrict & "460100,460107,琼山区,20.001051,110.354722;"
    AddressDistrict = AddressDistrict & "460100,460108,美兰区,20.03074,110.356566;"
    AddressDistrict = AddressDistrict & "460200,460202,海棠区,18.407516,109.760778;"
    AddressDistrict = AddressDistrict & "460200,460203,吉阳区,18.247436,109.512081;"
    AddressDistrict = AddressDistrict & "460200,460204,天涯区,18.24734,109.506357;"
    AddressDistrict = AddressDistrict & "460200,460205,崖州区,18.352192,109.174306;"
    AddressDistrict = AddressDistrict & "460300,460301,西沙区,16.8310066,112.3386402;"
    AddressDistrict = AddressDistrict & "460300,460302,南沙区,9.543575,112.891018;"
    AddressDistrict = AddressDistrict & "510100,510104,锦江区,30.657689,104.080989;"
    AddressDistrict = AddressDistrict & "510100,510105,青羊区,30.667648,104.055731;"
    AddressDistrict = AddressDistrict & "510100,510106,金牛区,30.692058,104.043487;"
    AddressDistrict = AddressDistrict & "510100,510107,武侯区,30.630862,104.05167;"
    AddressDistrict = AddressDistrict & "510100,510108,成华区,30.660275,104.103077;"
    AddressDistrict = AddressDistrict & "510100,510112,龙泉驿区,30.56065,104.269181;"
    AddressDistrict = AddressDistrict & "510100,510113,青白江区,30.883438,104.25494;"
    AddressDistrict = AddressDistrict & "510100,510114,新都区,30.824223,104.16022;"
    AddressDistrict = AddressDistrict & "510100,510115,温江区,30.697996,103.836776;"
    AddressDistrict = AddressDistrict & "510100,510116,双流区,30.573243,103.922706;"
    AddressDistrict = AddressDistrict & "510100,510117,郫都区,30.808752,103.887842;"
    AddressDistrict = AddressDistrict & "510100,510118,新津区,30.414284,103.812449;"
    AddressDistrict = AddressDistrict & "510100,510121,金堂县,30.858417,104.415604;"
    AddressDistrict = AddressDistrict & "510100,510129,大邑县,30.586602,103.522397;"
    AddressDistrict = AddressDistrict & "510100,510131,蒲江县,30.194359,103.511541;"
    AddressDistrict = AddressDistrict & "510100,510181,都江堰市,30.99114,103.627898;"
    AddressDistrict = AddressDistrict & "510100,510182,彭州市,30.985161,103.941173;"
    AddressDistrict = AddressDistrict & "510100,510183,邛崃市,30.413271,103.46143;"
    AddressDistrict = AddressDistrict & "510100,510184,崇州市,30.631478,103.671049;"
    AddressDistrict = AddressDistrict & "510100,510185,简阳市,30.390666,104.550339;"
    AddressDistrict = AddressDistrict & "510300,510302,自流井区,29.343231,104.778188;"
    AddressDistrict = AddressDistrict & "510300,510303,贡井区,29.345675,104.714372;"
    AddressDistrict = AddressDistrict & "510300,510304,大安区,29.367136,104.783229;"
    AddressDistrict = AddressDistrict & "510300,510311,沿滩区,29.272521,104.876417;"
    AddressDistrict = AddressDistrict & "510300,510321,荣县,29.454851,104.423932;"
    AddressDistrict = AddressDistrict & "510300,510322,富顺县,29.181282,104.984256;"
    AddressDistrict = AddressDistrict & "510400,510402,东区,26.580887,101.715134;"
    AddressDistrict = AddressDistrict & "510400,510403,西区,26.596776,101.637969;"
    AddressDistrict = AddressDistrict & "510400,510411,仁和区,26.497185,101.737916;"
    AddressDistrict = AddressDistrict & "510400,510421,米易县,26.887474,102.109877;"
    AddressDistrict = AddressDistrict & "510400,510422,盐边县,26.677619,101.851848;"
    AddressDistrict = AddressDistrict & "510500,510502,江阳区,28.882889,105.445131;"
    AddressDistrict = AddressDistrict & "510500,510503,纳溪区,28.77631,105.37721;"
    AddressDistrict = AddressDistrict & "510500,510504,龙马潭区,28.897572,105.435228;"
    AddressDistrict = AddressDistrict & "510500,510521,泸县,29.151288,105.376335;"
    AddressDistrict = AddressDistrict & "510500,510522,合江县,28.810325,105.834098;"
    AddressDistrict = AddressDistrict & "510500,510524,叙永县,28.167919,105.437775;"
    AddressDistrict = AddressDistrict & "510500,510525,古蔺县,28.03948,105.813359;"
    AddressDistrict = AddressDistrict & "510600,510603,旌阳区,31.130428,104.389648;"
    AddressDistrict = AddressDistrict & "510600,510604,罗江区,31.303281,104.507126;"
    AddressDistrict = AddressDistrict & "510600,510623,中江县,31.03681,104.677831;"
    AddressDistrict = AddressDistrict & "510600,510681,广汉市,30.97715,104.281903;"
    AddressDistrict = AddressDistrict & "510600,510682,什邡市,31.126881,104.173653;"
    AddressDistrict = AddressDistrict & "510600,510683,绵竹市,31.343084,104.200162;"
    AddressDistrict = AddressDistrict & "510700,510703,涪城区,31.463557,104.740971;"
    AddressDistrict = AddressDistrict & "510700,510704,游仙区,31.484772,104.770006;"
    AddressDistrict = AddressDistrict & "510700,510705,安州区,31.53894,104.560341;"
    AddressDistrict = AddressDistrict & "510700,510722,三台县,31.090909,105.090316;"
    AddressDistrict = AddressDistrict & "510700,510723,盐亭县,31.22318,105.391991;"
    AddressDistrict = AddressDistrict & "510700,510725,梓潼县,31.635225,105.16353;"
    AddressDistrict = AddressDistrict & "510700,510726,北川羌族自治县,31.615863,104.468069;"
    AddressDistrict = AddressDistrict & "510700,510727,平武县,32.407588,104.530555;"
    AddressDistrict = AddressDistrict & "510700,510781,江油市,31.776386,104.744431;"
    AddressDistrict = AddressDistrict & "510800,510802,利州区,32.432276,105.826194;"
    AddressDistrict = AddressDistrict & "510800,510811,昭化区,32.322788,105.964121;"
    AddressDistrict = AddressDistrict & "510800,510812,朝天区,32.642632,105.88917;"
    AddressDistrict = AddressDistrict & "510800,510821,旺苍县,32.22833,106.290426;"
    AddressDistrict = AddressDistrict & "510800,510822,青川县,32.585655,105.238847;"
    AddressDistrict = AddressDistrict & "510800,510823,剑阁县,32.286517,105.527035;"
    AddressDistrict = AddressDistrict & "510800,510824,苍溪县,31.732251,105.939706;"
    AddressDistrict = AddressDistrict & "510900,510903,船山区,30.502647,105.582215;"
    AddressDistrict = AddressDistrict & "510900,510904,安居区,30.346121,105.459383;"
    AddressDistrict = AddressDistrict & "510900,510921,蓬溪县,30.774883,105.713699;"
    AddressDistrict = AddressDistrict & "510900,510923,大英县,30.581571,105.252187;"
    AddressDistrict = AddressDistrict & "510900,510981,射洪市,30.868752,105.381849;"
    AddressDistrict = AddressDistrict & "511000,511002,市中区,29.585265,105.065467;"
    AddressDistrict = AddressDistrict & "511000,511011,东兴区,29.600107,105.067203;"
    AddressDistrict = AddressDistrict & "511000,511024,威远县,29.52686,104.668327;"
    AddressDistrict = AddressDistrict & "511000,511025,资中县,29.775295,104.852463;"
    AddressDistrict = AddressDistrict & "511000,511083,隆昌市,29.338162,105.288074;"
    AddressDistrict = AddressDistrict & "511100,511102,市中区,29.588327,103.75539;"
    AddressDistrict = AddressDistrict & "511100,511111,沙湾区,29.416536,103.549961;"
    AddressDistrict = AddressDistrict & "511100,511112,五通桥区,29.406186,103.816837;"
    AddressDistrict = AddressDistrict & "511100,511113,金口河区,29.24602,103.077831;"
    AddressDistrict = AddressDistrict & "511100,511123,犍为县,29.209782,103.944266;"
    AddressDistrict = AddressDistrict & "511100,511124,井研县,29.651645,104.06885;"
    AddressDistrict = AddressDistrict & "511100,511126,夹江县,29.741019,103.578862;"
    AddressDistrict = AddressDistrict & "511100,511129,沐川县,28.956338,103.90211;"
    AddressDistrict = AddressDistrict & "511100,511132,峨边彝族自治县,29.230271,103.262148;"
    AddressDistrict = AddressDistrict & "511100,511133,马边彝族自治县,28.838933,103.546851;"
    AddressDistrict = AddressDistrict & "511100,511181,峨眉山市,29.597478,103.492488;"
    AddressDistrict = AddressDistrict & "511300,511302,顺庆区,30.795572,106.084091;"
    AddressDistrict = AddressDistrict & "511300,511303,高坪区,30.781809,106.108996;"
    AddressDistrict = AddressDistrict & "511300,511304,嘉陵区,30.762976,106.067027;"
    AddressDistrict = AddressDistrict & "511300,511321,南部县,31.349407,106.061138;"
    AddressDistrict = AddressDistrict & "511300,511322,营山县,31.075907,106.564893;"
    AddressDistrict = AddressDistrict & "511300,511323,蓬安县,31.027978,106.413488;"
    AddressDistrict = AddressDistrict & "511300,511324,仪陇县,31.271261,106.297083;"
    AddressDistrict = AddressDistrict & "511300,511325,西充县,30.994616,105.893021;"
    AddressDistrict = AddressDistrict & "511300,511381,阆中市,31.580466,105.975266;"
    AddressDistrict = AddressDistrict & "511400,511402,东坡区,30.048128,103.831553;"
    AddressDistrict = AddressDistrict & "511400,511403,彭山区,30.192298,103.8701;"
    AddressDistrict = AddressDistrict & "511400,511421,仁寿县,29.996721,104.147646;"
    AddressDistrict = AddressDistrict & "511400,511423,洪雅县,29.904867,103.375006;"
    AddressDistrict = AddressDistrict & "511400,511424,丹棱县,30.012751,103.518333;"
    AddressDistrict = AddressDistrict & "511400,511425,青神县,29.831469,103.846131;"
    AddressDistrict = AddressDistrict & "511500,511502,翠屏区,28.760179,104.630231;"
    AddressDistrict = AddressDistrict & "511500,511503,南溪区,28.839806,104.981133;"
    AddressDistrict = AddressDistrict & "511500,511504,叙州区,28.695678,104.541489;"
    AddressDistrict = AddressDistrict & "511500,511523,江安县,28.728102,105.068697;"
    AddressDistrict = AddressDistrict & "511500,511524,长宁县,28.577271,104.921116;"
    AddressDistrict = AddressDistrict & "511500,511525,高县,28.435676,104.519187;"
    AddressDistrict = AddressDistrict & "511500,511526,珙县,28.449041,104.712268;"
    AddressDistrict = AddressDistrict & "511500,511527,筠连县,28.162017,104.507848;"
    AddressDistrict = AddressDistrict & "511500,511528,兴文县,28.302988,105.236549;"
    AddressDistrict = AddressDistrict & "511500,511529,屏山县,28.64237,104.162617;"
    AddressDistrict = AddressDistrict & "511600,511602,广安区,30.456462,106.632907;"
    AddressDistrict = AddressDistrict & "511600,511603,前锋区,30.4963,106.893277;"
    AddressDistrict = AddressDistrict & "511600,511621,岳池县,30.533538,106.444451;"
    AddressDistrict = AddressDistrict & "511600,511622,武胜县,30.344291,106.292473;"
    AddressDistrict = AddressDistrict & "511600,511623,邻水县,30.334323,106.934968;"
    AddressDistrict = AddressDistrict & "511600,511681,华蓥市,30.380574,106.777882;"
    AddressDistrict = AddressDistrict & "511700,511702,通川区,31.213522,107.501062;"
    AddressDistrict = AddressDistrict & "511700,511703,达川区,31.199062,107.507926;"
    AddressDistrict = AddressDistrict & "511700,511722,宣汉县,31.355025,107.722254;"
    AddressDistrict = AddressDistrict & "511700,511723,开江县,31.085537,107.864135;"
    AddressDistrict = AddressDistrict & "511700,511724,大竹县,30.736289,107.20742;"
    AddressDistrict = AddressDistrict & "511700,511725,渠县,30.836348,106.970746;"
    AddressDistrict = AddressDistrict & "511700,511781,万源市,32.06777,108.037548;"
    AddressDistrict = AddressDistrict & "511800,511802,雨城区,29.981831,103.003398;"
    AddressDistrict = AddressDistrict & "511800,511803,名山区,30.084718,103.112214;"
    AddressDistrict = AddressDistrict & "511800,511822,荥经县,29.795529,102.844674;"
    AddressDistrict = AddressDistrict & "511800,511823,汉源县,29.349915,102.677145;"
    AddressDistrict = AddressDistrict & "511800,511824,石棉县,29.234063,102.35962;"
    AddressDistrict = AddressDistrict & "511800,511825,天全县,30.059955,102.763462;"
    AddressDistrict = AddressDistrict & "511800,511826,芦山县,30.152907,102.924016;"
    AddressDistrict = AddressDistrict & "511800,511827,宝兴县,30.369026,102.813377;"
    AddressDistrict = AddressDistrict & "511900,511902,巴州区,31.858366,106.753671;"
    AddressDistrict = AddressDistrict & "511900,511903,恩阳区,31.816336,106.486515;"
    AddressDistrict = AddressDistrict & "511900,511921,通江县,31.91212,107.247621;"
    AddressDistrict = AddressDistrict & "511900,511922,南江县,32.353164,106.843418;"
    AddressDistrict = AddressDistrict & "511900,511923,平昌县,31.562814,107.101937;"
    AddressDistrict = AddressDistrict & "512000,512002,雁江区,30.121686,104.642338;"
    AddressDistrict = AddressDistrict & "512000,512021,安岳县,30.099206,105.336764;"
    AddressDistrict = AddressDistrict & "512000,512022,乐至县,30.275619,105.031142;"
    AddressDistrict = AddressDistrict & "513200,513201,马尔康市,31.899761,102.221187;"
    AddressDistrict = AddressDistrict & "513200,513221,汶川县,31.47463,103.580675;"
    AddressDistrict = AddressDistrict & "513200,513222,理县,31.436764,103.165486;"
    AddressDistrict = AddressDistrict & "513200,513223,茂县,31.680407,103.850684;"
    AddressDistrict = AddressDistrict & "513200,513224,松潘县,32.63838,103.599177;"
    AddressDistrict = AddressDistrict & "513200,513225,九寨沟县,33.262097,104.236344;"
    AddressDistrict = AddressDistrict & "513200,513226,金川县,31.476356,102.064647;"
    AddressDistrict = AddressDistrict & "513200,513227,小金县,30.999016,102.363193;"
    AddressDistrict = AddressDistrict & "513200,513228,黑水县,32.061721,102.990805;"
    AddressDistrict = AddressDistrict & "513200,513230,壤塘县,32.264887,100.979136;"
    AddressDistrict = AddressDistrict & "513200,513231,阿坝县,32.904223,101.700985;"
    AddressDistrict = AddressDistrict & "513200,513232,若尔盖县,33.575934,102.963726;"
    AddressDistrict = AddressDistrict & "513200,513233,红原县,32.793902,102.544906;"
    AddressDistrict = AddressDistrict & "513300,513301,康定市,30.050738,101.964057;"
    AddressDistrict = AddressDistrict & "513300,513322,泸定县,29.912482,102.233225;"
    AddressDistrict = AddressDistrict & "513300,513323,丹巴县,30.877083,101.886125;"
    AddressDistrict = AddressDistrict & "513300,513324,九龙县,29.001975,101.506942;"
    AddressDistrict = AddressDistrict & "513300,513325,雅江县,30.03225,101.015735;"
    AddressDistrict = AddressDistrict & "513300,513326,道孚县,30.978767,101.123327;"
    AddressDistrict = AddressDistrict & "513300,513327,炉霍县,31.392674,100.679495;"
    AddressDistrict = AddressDistrict & "513300,513328,甘孜县,31.61975,99.991753;"
    AddressDistrict = AddressDistrict & "513300,513329,新龙县,30.93896,100.312094;"
    AddressDistrict = AddressDistrict & "513300,513330,德格县,31.806729,98.57999;"
    AddressDistrict = AddressDistrict & "513300,513331,白玉县,31.208805,98.824343;"
    AddressDistrict = AddressDistrict & "513300,513332,石渠县,32.975302,98.100887;"
    AddressDistrict = AddressDistrict & "513300,513333,色达县,32.268777,100.331657;"
    AddressDistrict = AddressDistrict & "513300,513334,理塘县,29.991807,100.269862;"
    AddressDistrict = AddressDistrict & "513300,513335,巴塘县,30.005723,99.109037;"
    AddressDistrict = AddressDistrict & "513300,513336,乡城县,28.930855,99.799943;"
    AddressDistrict = AddressDistrict & "513300,513337,稻城县,29.037544,100.296689;"
    AddressDistrict = AddressDistrict & "513300,513338,得荣县,28.71134,99.288036;"
    AddressDistrict = AddressDistrict & "513400,513401,西昌市,27.885786,102.258758;"
    AddressDistrict = AddressDistrict & "513400,513422,木里藏族自治县,27.926859,101.280184;"
    AddressDistrict = AddressDistrict & "513400,513423,盐源县,27.423415,101.508909;"
    AddressDistrict = AddressDistrict & "513400,513424,德昌县,27.403827,102.178845;"
    AddressDistrict = AddressDistrict & "513400,513425,会理市,26.658702,102.249548;"
    AddressDistrict = AddressDistrict & "513400,513426,会东县,26.630713,102.578985;"
    AddressDistrict = AddressDistrict & "513400,513427,宁南县,27.065205,102.757374;"
    AddressDistrict = AddressDistrict & "513400,513428,普格县,27.376828,102.541082;"
    AddressDistrict = AddressDistrict & "513400,513429,布拖县,27.709062,102.808801;"
    AddressDistrict = AddressDistrict & "513400,513430,金阳县,27.695916,103.248704;"
    AddressDistrict = AddressDistrict & "513400,513431,昭觉县,28.010554,102.843991;"
    AddressDistrict = AddressDistrict & "513400,513432,喜德县,28.305486,102.412342;"
    AddressDistrict = AddressDistrict & "513400,513433,冕宁县,28.550844,102.170046;"
    AddressDistrict = AddressDistrict & "513400,513434,越西县,28.639632,102.508875;"
    AddressDistrict = AddressDistrict & "513400,513435,甘洛县,28.977094,102.775924;"
    AddressDistrict = AddressDistrict & "513400,513436,美姑县,28.327946,103.132007;"
    AddressDistrict = AddressDistrict & "513400,513437,雷波县,28.262946,103.571584;"
    AddressDistrict = AddressDistrict & "520100,520102,南明区,26.573743,106.715963;"
    AddressDistrict = AddressDistrict & "520100,520103,云岩区,26.58301,106.713397;"
    AddressDistrict = AddressDistrict & "520100,520111,花溪区,26.410464,106.670791;"
    AddressDistrict = AddressDistrict & "520100,520112,乌当区,26.630928,106.762123;"
    AddressDistrict = AddressDistrict & "520100,520113,白云区,26.676849,106.633037;"
    AddressDistrict = AddressDistrict & "520100,520115,观山湖区,26.646358,106.626323;"
    AddressDistrict = AddressDistrict & "520100,520121,开阳县,27.056793,106.969438;"
    AddressDistrict = AddressDistrict & "520100,520122,息烽县,27.092665,106.737693;"
    AddressDistrict = AddressDistrict & "520100,520123,修文县,26.840672,106.599218;"
    AddressDistrict = AddressDistrict & "520100,520181,清镇市,26.551289,106.470278;"
    AddressDistrict = AddressDistrict & "520200,520201,钟山区,26.584805,104.846244;"
    AddressDistrict = AddressDistrict & "520200,520203,六枝特区,26.210662,105.474235;"
    AddressDistrict = AddressDistrict & "520200,520221,水城区,26.540478,104.95685;"
    AddressDistrict = AddressDistrict & "520200,520281,盘州市,25.706966,104.468367;"
    AddressDistrict = AddressDistrict & "520300,520302,红花岗区,27.694395,106.943784;"
    AddressDistrict = AddressDistrict & "520300,520303,汇川区,27.706626,106.937265;"
    AddressDistrict = AddressDistrict & "520300,520304,播州区,27.535288,106.831668;"
    AddressDistrict = AddressDistrict & "520300,520322,桐梓县,28.131559,106.826591;"
    AddressDistrict = AddressDistrict & "520300,520323,绥阳县,27.951342,107.191024;"
    AddressDistrict = AddressDistrict & "520300,520324,正安县,28.550337,107.441872;"
    AddressDistrict = AddressDistrict & "520300,520325,道真仡佬族苗族自治县,28.880088,107.605342;"
    AddressDistrict = AddressDistrict & "520300,520326,务川仡佬族苗族自治县,28.521567,107.887857;"
    AddressDistrict = AddressDistrict & "520300,520327,凤冈县,27.960858,107.722021;"
    AddressDistrict = AddressDistrict & "520300,520328,湄潭县,27.765839,107.485723;"
    AddressDistrict = AddressDistrict & "520300,520329,余庆县,27.221552,107.892566;"
    AddressDistrict = AddressDistrict & "520300,520330,习水县,28.327826,106.200954;"
    AddressDistrict = AddressDistrict & "520300,520381,赤水市,28.587057,105.698116;"
    AddressDistrict = AddressDistrict & "520300,520382,仁怀市,27.803377,106.412476;"
    AddressDistrict = AddressDistrict & "520400,520402,西秀区,26.248323,105.946169;"
    AddressDistrict = AddressDistrict & "520400,520403,平坝区,26.40608,106.259942;"
    AddressDistrict = AddressDistrict & "520400,520422,普定县,26.305794,105.745609;"
    AddressDistrict = AddressDistrict & "520400,520423,镇宁布依族苗族自治县,26.056096,105.768656;"
    AddressDistrict = AddressDistrict & "520400,520424,关岭布依族苗族自治县,25.944248,105.618454;"
    AddressDistrict = AddressDistrict & "520400,520425,紫云苗族布依族自治县,25.751567,106.084515;"
    AddressDistrict = AddressDistrict & "520500,520502,七星关区,27.302085,105.284852;"
    AddressDistrict = AddressDistrict & "520500,520521,大方县,27.143521,105.609254;"
    AddressDistrict = AddressDistrict & "520500,520522,黔西市,27.024923,106.038299;"
    AddressDistrict = AddressDistrict & "520500,520523,金沙县,27.459693,106.222103;"
    AddressDistrict = AddressDistrict & "520500,520524,织金县,26.668497,105.768997;"
    AddressDistrict = AddressDistrict & "520500,520525,纳雍县,26.769875,105.375322;"
    AddressDistrict = AddressDistrict & "520500,520526,威宁彝族回族苗族自治县,26.859099,104.286523;"
    AddressDistrict = AddressDistrict & "520500,520527,赫章县,27.119243,104.726438;"
    AddressDistrict = AddressDistrict & "520600,520602,碧江区,27.718745,109.192117;"
    AddressDistrict = AddressDistrict & "520600,520603,万山区,27.51903,109.21199;"
    AddressDistrict = AddressDistrict & "520600,520621,江口县,27.691904,108.848427;"
    AddressDistrict = AddressDistrict & "520600,520622,玉屏侗族自治县,27.238024,108.917882;"
    AddressDistrict = AddressDistrict & "520600,520623,石阡县,27.519386,108.229854;"
    AddressDistrict = AddressDistrict & "520600,520624,思南县,27.941331,108.255827;"
    AddressDistrict = AddressDistrict & "520600,520625,印江土家族苗族自治县,27.997976,108.405517;"
    AddressDistrict = AddressDistrict & "520600,520626,德江县,28.26094,108.117317;"
    AddressDistrict = AddressDistrict & "520600,520627,沿河土家族自治县,28.560487,108.495746;"
    AddressDistrict = AddressDistrict & "520600,520628,松桃苗族自治县,28.165419,109.202627;"
    AddressDistrict = AddressDistrict & "522300,522301,兴义市,25.088599,104.897982;"
    AddressDistrict = AddressDistrict & "522300,522302,兴仁市,25.431378,105.192778;"
    AddressDistrict = AddressDistrict & "522300,522323,普安县,25.786404,104.955347;"
    AddressDistrict = AddressDistrict & "522300,522324,晴隆县,25.832881,105.218773;"
    AddressDistrict = AddressDistrict & "522300,522325,贞丰县,25.385752,105.650133;"
    AddressDistrict = AddressDistrict & "522300,522326,望谟县,25.166667,106.091563;"
    AddressDistrict = AddressDistrict & "522300,522327,册亨县,24.983338,105.81241;"
    AddressDistrict = AddressDistrict & "522300,522328,安龙县,25.108959,105.471498;"
    AddressDistrict = AddressDistrict & "522600,522601,凯里市,26.582964,107.977541;"
    AddressDistrict = AddressDistrict & "522600,522622,黄平县,26.896973,107.901337;"
    AddressDistrict = AddressDistrict & "522600,522623,施秉县,27.034657,108.12678;"
    AddressDistrict = AddressDistrict & "522600,522624,三穗县,26.959884,108.681121;"
    AddressDistrict = AddressDistrict & "522600,522625,镇远县,27.050233,108.423656;"
    AddressDistrict = AddressDistrict & "522600,522626,岑巩县,27.173244,108.816459;"
    AddressDistrict = AddressDistrict & "522600,522627,天柱县,26.909684,109.212798;"
    AddressDistrict = AddressDistrict & "522600,522628,锦屏县,26.680625,109.20252;"
    AddressDistrict = AddressDistrict & "522600,522629,剑河县,26.727349,108.440499;"
    AddressDistrict = AddressDistrict & "522600,522630,台江县,26.669138,108.314637;"
    AddressDistrict = AddressDistrict & "522600,522631,黎平县,26.230636,109.136504;"
    AddressDistrict = AddressDistrict & "522600,522632,榕江县,25.931085,108.521026;"
    AddressDistrict = AddressDistrict & "522600,522633,从江县,25.747058,108.912648;"
    AddressDistrict = AddressDistrict & "522600,522634,雷山县,26.381027,108.079613;"
    AddressDistrict = AddressDistrict & "522600,522635,麻江县,26.494803,107.593172;"
    AddressDistrict = AddressDistrict & "522600,522636,丹寨县,26.199497,107.794808;"
    AddressDistrict = AddressDistrict & "522700,522701,都匀市,26.258205,107.517021;"
    AddressDistrict = AddressDistrict & "522700,522702,福泉市,26.702508,107.513508;"
    AddressDistrict = AddressDistrict & "522700,522722,荔波县,25.412239,107.8838;"
    AddressDistrict = AddressDistrict & "522700,522723,贵定县,26.580807,107.233588;"
    AddressDistrict = AddressDistrict & "522700,522725,瓮安县,27.066339,107.478417;"
    AddressDistrict = AddressDistrict & "522700,522726,独山县,25.826283,107.542757;"
    AddressDistrict = AddressDistrict & "522700,522727,平塘县,25.831803,107.32405;"
    AddressDistrict = AddressDistrict & "522700,522728,罗甸县,25.429894,106.750006;"
    AddressDistrict = AddressDistrict & "522700,522729,长顺县,26.022116,106.447376;"
    AddressDistrict = AddressDistrict & "522700,522730,龙里县,26.448809,106.977733;"
    AddressDistrict = AddressDistrict & "522700,522731,惠水县,26.128637,106.657848;"
    AddressDistrict = AddressDistrict & "522700,522732,三都水族自治县,25.985183,107.87747;"
    AddressDistrict = AddressDistrict & "530100,530102,五华区,25.042165,102.704412;"
    AddressDistrict = AddressDistrict & "530100,530103,盘龙区,25.070239,102.729044;"
    AddressDistrict = AddressDistrict & "530100,530111,官渡区,25.021211,102.723437;"
    AddressDistrict = AddressDistrict & "530100,530112,西山区,25.02436,102.705904;"
    AddressDistrict = AddressDistrict & "530100,530113,东川区,26.08349,103.182;"
    AddressDistrict = AddressDistrict & "530100,530114,呈贡区,24.889275,102.801382;"
    AddressDistrict = AddressDistrict & "530100,530115,晋宁区,24.666944,102.594987;"
    AddressDistrict = AddressDistrict & "530100,530124,富民县,25.219667,102.497888;"
    AddressDistrict = AddressDistrict & "530100,530125,宜良县,24.918215,103.145989;"
    AddressDistrict = AddressDistrict & "530100,530126,石林彝族自治县,24.754545,103.271962;"
    AddressDistrict = AddressDistrict & "530100,530127,嵩明县,25.335087,103.038777;"
    AddressDistrict = AddressDistrict & "530100,530128,禄劝彝族苗族自治县,25.556533,102.46905;"
    AddressDistrict = AddressDistrict & "530100,530129,寻甸回族彝族自治县,25.559474,103.257588;"
    AddressDistrict = AddressDistrict & "530100,530181,安宁市,24.921785,102.485544;"
    AddressDistrict = AddressDistrict & "530300,530302,麒麟区,25.501269,103.798054;"
    AddressDistrict = AddressDistrict & "530300,530303,沾益区,25.600878,103.819262;"
    AddressDistrict = AddressDistrict & "530300,530304,马龙区,25.429451,103.578755;"
    AddressDistrict = AddressDistrict & "530300,530322,陆良县,25.022878,103.655233;"
    AddressDistrict = AddressDistrict & "530300,530323,师宗县,24.825681,103.993808;"
    AddressDistrict = AddressDistrict & "530300,530324,罗平县,24.885708,104.309263;"
    AddressDistrict = AddressDistrict & "530300,530325,富源县,25.67064,104.25692;"
    AddressDistrict = AddressDistrict & "530300,530326,会泽县,26.412861,103.300041;"
    AddressDistrict = AddressDistrict & "530300,530381,宣威市,26.227777,104.09554;"
    AddressDistrict = AddressDistrict & "530400,530402,红塔区,24.350753,102.543468;"
    AddressDistrict = AddressDistrict & "530400,530403,江川区,24.291006,102.749839;"
    AddressDistrict = AddressDistrict & "530400,530423,通海县,24.112205,102.760039;"
    AddressDistrict = AddressDistrict & "530400,530424,华宁县,24.189807,102.928982;"
    AddressDistrict = AddressDistrict & "530400,530425,易门县,24.669598,102.16211;"
    AddressDistrict = AddressDistrict & "530400,530426,峨山彝族自治县,24.173256,102.404358;"
    AddressDistrict = AddressDistrict & "530400,530427,新平彝族傣族自治县,24.0664,101.990903;"
    AddressDistrict = AddressDistrict & "530400,530428,元江哈尼族彝族傣族自治县,23.597618,101.999658;"
    AddressDistrict = AddressDistrict & "530400,530481,澄江市,24.669679,102.916652;"
    AddressDistrict = AddressDistrict & "530500,530502,隆阳区,25.112144,99.165825;"
    AddressDistrict = AddressDistrict & "530500,530521,施甸县,24.730847,99.183758;"
    AddressDistrict = AddressDistrict & "530500,530523,龙陵县,24.591912,98.693567;"
    AddressDistrict = AddressDistrict & "530500,530524,昌宁县,24.823662,99.612344;"
    AddressDistrict = AddressDistrict & "530500,530581,腾冲市,25.01757,98.497292;"
    AddressDistrict = AddressDistrict & "530600,530602,昭阳区,27.336636,103.717267;"
    AddressDistrict = AddressDistrict & "530600,530621,鲁甸县,27.191637,103.549333;"
    AddressDistrict = AddressDistrict & "530600,530622,巧家县,26.9117,102.929284;"
    AddressDistrict = AddressDistrict & "530600,530623,盐津县,28.106923,104.23506;"
    AddressDistrict = AddressDistrict & "530600,530624,大关县,27.747114,103.891608;"
    AddressDistrict = AddressDistrict & "530600,530625,永善县,28.231526,103.63732;"
    AddressDistrict = AddressDistrict & "530600,530626,绥江县,28.599953,103.961095;"
    AddressDistrict = AddressDistrict & "530600,530627,镇雄县,27.436267,104.873055;"
    AddressDistrict = AddressDistrict & "530600,530628,彝良县,27.627425,104.048492;"
    AddressDistrict = AddressDistrict & "530600,530629,威信县,27.843381,105.04869;"
    AddressDistrict = AddressDistrict & "530600,530681,水富市,28.629688,104.415376;"
    AddressDistrict = AddressDistrict & "530700,530702,古城区,26.872229,100.234412;"
    AddressDistrict = AddressDistrict & "530700,530721,玉龙纳西族自治县,26.830593,100.238312;"
    AddressDistrict = AddressDistrict & "530700,530722,永胜县,26.685623,100.750901;"
    AddressDistrict = AddressDistrict & "530700,530723,华坪县,26.628834,101.267796;"
    AddressDistrict = AddressDistrict & "530700,530724,宁蒗彝族自治县,27.281109,100.852427;"
    AddressDistrict = AddressDistrict & "530800,530802,思茅区,22.776595,100.973227;"
    AddressDistrict = AddressDistrict & "530800,530821,宁洱哈尼族彝族自治县,23.062507,101.04524;"
    AddressDistrict = AddressDistrict & "530800,530822,墨江哈尼族自治县,23.428165,101.687606;"
    AddressDistrict = AddressDistrict & "530800,530823,景东彝族自治县,24.448523,100.840011;"
    AddressDistrict = AddressDistrict & "530800,530824,景谷傣族彝族自治县,23.500278,100.701425;"
    AddressDistrict = AddressDistrict & "530800,530825,镇沅彝族哈尼族拉祜族自治县,24.005712,101.108512;"
    AddressDistrict = AddressDistrict & "530800,530826,江城哈尼族彝族自治县,22.58336,101.859144;"
    AddressDistrict = AddressDistrict & "530800,530827,孟连傣族拉祜族佤族自治县,22.325924,99.585406;"
    AddressDistrict = AddressDistrict & "530800,530828,澜沧拉祜族自治县,22.553083,99.931201;"
    AddressDistrict = AddressDistrict & "530800,530829,西盟佤族自治县,22.644423,99.594372;"
    AddressDistrict = AddressDistrict & "530900,530902,临翔区,23.886562,100.086486;"
    AddressDistrict = AddressDistrict & "530900,530921,凤庆县,24.592738,99.91871;"
    AddressDistrict = AddressDistrict & "530900,530922,云县,24.439026,100.125637;"
    AddressDistrict = AddressDistrict & "530900,530923,永德县,24.028159,99.253679;"
    AddressDistrict = AddressDistrict & "530900,530924,镇康县,23.761415,98.82743;"
    AddressDistrict = AddressDistrict & "530900,530925,双江拉祜族佤族布朗族傣族自治县,23.477476,99.824419;"
    AddressDistrict = AddressDistrict & "530900,530926,耿马傣族佤族自治县,23.534579,99.402495;"
    AddressDistrict = AddressDistrict & "530900,530927,沧源佤族自治县,23.146887,99.2474;"
    AddressDistrict = AddressDistrict & "532300,532301,楚雄市,25.040912,101.546145;"
    AddressDistrict = AddressDistrict & "532300,532322,双柏县,24.685094,101.63824;"
    AddressDistrict = AddressDistrict & "532300,532323,牟定县,25.312111,101.543044;"
    AddressDistrict = AddressDistrict & "532300,532324,南华县,25.192408,101.274991;"
    AddressDistrict = AddressDistrict & "532300,532325,姚安县,25.505403,101.238399;"
    AddressDistrict = AddressDistrict & "532300,532326,大姚县,25.722348,101.323602;"
    AddressDistrict = AddressDistrict & "532300,532327,永仁县,26.056316,101.671175;"
    AddressDistrict = AddressDistrict & "532300,532328,元谋县,25.703313,101.870837;"
    AddressDistrict = AddressDistrict & "532300,532329,武定县,25.5301,102.406785;"
    AddressDistrict = AddressDistrict & "532300,532331,禄丰市,25.14327,102.075694;"
    AddressDistrict = AddressDistrict & "532500,532501,个旧市,23.360383,103.154752;"
    AddressDistrict = AddressDistrict & "532500,532502,开远市,23.713832,103.258679;"
    AddressDistrict = AddressDistrict & "532500,532503,蒙自市,23.366843,103.385005;"
    AddressDistrict = AddressDistrict & "532500,532504,弥勒市,24.40837,103.436988;"
    AddressDistrict = AddressDistrict & "532500,532523,屏边苗族自治县,22.987013,103.687229;"
    AddressDistrict = AddressDistrict & "532500,532524,建水县,23.618387,102.820493;"
    AddressDistrict = AddressDistrict & "532500,532525,石屏县,23.712569,102.484469;"
    AddressDistrict = AddressDistrict & "532500,532527,泸西县,24.532368,103.759622;"
    AddressDistrict = AddressDistrict & "532500,532528,元阳县,23.219773,102.837056;"
    AddressDistrict = AddressDistrict & "532500,532529,红河县,23.369191,102.42121;"
    AddressDistrict = AddressDistrict & "532500,532530,金平苗族瑶族傣族自治县,22.779982,103.228359;"
    AddressDistrict = AddressDistrict & "532500,532531,绿春县,22.99352,102.39286;"
    AddressDistrict = AddressDistrict & "532500,532532,河口瑶族自治县,22.507563,103.961593;"
    AddressDistrict = AddressDistrict & "532600,532601,文山市,23.369216,104.244277;"
    AddressDistrict = AddressDistrict & "532600,532622,砚山县,23.612301,104.343989;"
    AddressDistrict = AddressDistrict & "532600,532623,西畴县,23.437439,104.675711;"
    AddressDistrict = AddressDistrict & "532600,532624,麻栗坡县,23.124202,104.701899;"
    AddressDistrict = AddressDistrict & "532600,532625,马关县,23.011723,104.398619;"
    AddressDistrict = AddressDistrict & "532600,532626,丘北县,24.040982,104.194366;"
    AddressDistrict = AddressDistrict & "532600,532627,广南县,24.050272,105.056684;"
    AddressDistrict = AddressDistrict & "532600,532628,富宁县,23.626494,105.62856;"
    AddressDistrict = AddressDistrict & "532800,532801,景洪市,22.002087,100.797947;"
    AddressDistrict = AddressDistrict & "532800,532822,勐海县,21.955866,100.448288;"
    AddressDistrict = AddressDistrict & "532800,532823,勐腊县,21.479449,101.567051;"
    AddressDistrict = AddressDistrict & "532900,532901,大理市,25.593067,100.241369;"
    AddressDistrict = AddressDistrict & "532900,532922,漾濞彝族自治县,25.669543,99.95797;"
    AddressDistrict = AddressDistrict & "532900,532923,祥云县,25.477072,100.554025;"
    AddressDistrict = AddressDistrict & "532900,532924,宾川县,25.825904,100.578957;"
    AddressDistrict = AddressDistrict & "532900,532925,弥渡县,25.342594,100.490669;"
    AddressDistrict = AddressDistrict & "532900,532926,南涧彝族自治县,25.041279,100.518683;"
    AddressDistrict = AddressDistrict & "532900,532927,巍山彝族回族自治县,25.230909,100.30793;"
    AddressDistrict = AddressDistrict & "532900,532928,永平县,25.461281,99.533536;"
    AddressDistrict = AddressDistrict & "532900,532929,云龙县,25.884955,99.369402;"
    AddressDistrict = AddressDistrict & "532900,532930,洱源县,26.111184,99.951708;"
    AddressDistrict = AddressDistrict & "532900,532931,剑川县,26.530066,99.905887;"
    AddressDistrict = AddressDistrict & "532900,532932,鹤庆县,26.55839,100.173375;"
    AddressDistrict = AddressDistrict & "533100,533102,瑞丽市,24.010734,97.855883;"
    AddressDistrict = AddressDistrict & "533100,533103,芒市,24.436699,98.577608;"
    AddressDistrict = AddressDistrict & "533100,533122,梁河县,24.80742,98.298196;"
    AddressDistrict = AddressDistrict & "533100,533123,盈江县,24.709541,97.93393;"
    AddressDistrict = AddressDistrict & "533100,533124,陇川县,24.184065,97.794441;"
    AddressDistrict = AddressDistrict & "533300,533301,泸水市,25.851142,98.854063;"
    AddressDistrict = AddressDistrict & "533300,533323,福贡县,26.902738,98.867413;"
    AddressDistrict = AddressDistrict & "533300,533324,贡山独龙族怒族自治县,27.738054,98.666141;"
    AddressDistrict = AddressDistrict & "533300,533325,兰坪白族普米族自治县,26.453839,99.421378;"
    AddressDistrict = AddressDistrict & "533400,533401,香格里拉市,27.825804,99.708667;"
    AddressDistrict = AddressDistrict & "533400,533422,德钦县,28.483272,98.91506;"
    AddressDistrict = AddressDistrict & "533400,533423,维西傈僳族自治县,27.180948,99.286355;"
    AddressDistrict = AddressDistrict & "540100,540102,城关区,29.659472,91.132911;"
    AddressDistrict = AddressDistrict & "540100,540103,堆龙德庆区,29.647347,91.002823;"
    AddressDistrict = AddressDistrict & "540100,540104,达孜区,29.670314,91.350976;"
    AddressDistrict = AddressDistrict & "540100,540121,林周县,29.895754,91.261842;"
    AddressDistrict = AddressDistrict & "540100,540122,当雄县,30.474819,91.103551;"
    AddressDistrict = AddressDistrict & "540100,540123,尼木县,29.431346,90.165545;"
    AddressDistrict = AddressDistrict & "540100,540124,曲水县,29.349895,90.738051;"
    AddressDistrict = AddressDistrict & "540100,540127,墨竹工卡县,29.834657,91.731158;"
    AddressDistrict = AddressDistrict & "540200,540202,桑珠孜区,29.267003,88.88667;"
    AddressDistrict = AddressDistrict & "540200,540221,南木林县,29.680459,89.099434;"
    AddressDistrict = AddressDistrict & "540200,540222,江孜县,28.908845,89.605044;"
    AddressDistrict = AddressDistrict & "540200,540223,定日县,28.656667,87.123887;"
    AddressDistrict = AddressDistrict & "540200,540224,萨迦县,28.901077,88.023007;"
    AddressDistrict = AddressDistrict & "540200,540225,拉孜县,29.085136,87.63743;"
    AddressDistrict = AddressDistrict & "540200,540226,昂仁县,29.294758,87.23578;"
    AddressDistrict = AddressDistrict & "540200,540227,谢通门县,29.431597,88.260517;"
    AddressDistrict = AddressDistrict & "540200,540228,白朗县,29.106627,89.263618;"
    AddressDistrict = AddressDistrict & "540200,540229,仁布县,29.230299,89.843207;"
    AddressDistrict = AddressDistrict & "540200,540230,康马县,28.554719,89.683406;"
    AddressDistrict = AddressDistrict & "540200,540231,定结县,28.36409,87.767723;"
    AddressDistrict = AddressDistrict & "540200,540232,仲巴县,29.768336,84.032826;"
    AddressDistrict = AddressDistrict & "540200,540233,亚东县,27.482772,88.906806;"
    AddressDistrict = AddressDistrict & "540200,540234,吉隆县,28.852416,85.298349;"
    AddressDistrict = AddressDistrict & "540200,540235,聂拉木县,28.15595,85.981953;"
    AddressDistrict = AddressDistrict & "540200,540236,萨嘎县,29.328194,85.234622;"
    AddressDistrict = AddressDistrict & "540200,540237,岗巴县,28.274371,88.518903;"
    AddressDistrict = AddressDistrict & "540300,540302,卡若区,31.137035,97.178255;"
    AddressDistrict = AddressDistrict & "540300,540321,江达县,31.499534,98.218351;"
    AddressDistrict = AddressDistrict & "540300,540322,贡觉县,30.859206,98.271191;"
    AddressDistrict = AddressDistrict & "540300,540323,类乌齐县,31.213048,96.601259;"
    AddressDistrict = AddressDistrict & "540300,540324,丁青县,31.410681,95.597748;"
    AddressDistrict = AddressDistrict & "540300,540325,察雅县,30.653038,97.565701;"
    AddressDistrict = AddressDistrict & "540300,540326,八宿县,30.053408,96.917893;"
    AddressDistrict = AddressDistrict & "540300,540327,左贡县,29.671335,97.840532;"
    AddressDistrict = AddressDistrict & "540300,540328,芒康县,29.686615,98.596444;"
    AddressDistrict = AddressDistrict & "540300,540329,洛隆县,30.741947,95.823418;"
    AddressDistrict = AddressDistrict & "540300,540330,边坝县,30.933849,94.707504;"
    AddressDistrict = AddressDistrict & "540400,540402,巴宜区,29.653732,94.360987;"
    AddressDistrict = AddressDistrict & "540400,540421,工布江达县,29.88447,93.246515;"
    AddressDistrict = AddressDistrict & "540400,540422,米林县,29.213811,94.213679;"
    AddressDistrict = AddressDistrict & "540400,540423,墨脱县,29.32573,95.332245;"
    AddressDistrict = AddressDistrict & "540400,540424,波密县,29.858771,95.768151;"
    AddressDistrict = AddressDistrict & "540400,540425,察隅县,28.660244,97.465002;"
    AddressDistrict = AddressDistrict & "540400,540426,朗县,29.0446,93.073429;"
    AddressDistrict = AddressDistrict & "540500,540502,乃东区,29.236106,91.76525;"
    AddressDistrict = AddressDistrict & "540500,540521,扎囊县,29.246476,91.338;"
    AddressDistrict = AddressDistrict & "540500,540522,贡嘎县,29.289078,90.985271;"
    AddressDistrict = AddressDistrict & "540500,540523,桑日县,29.259774,92.015732;"
    AddressDistrict = AddressDistrict & "540500,540524,琼结县,29.025242,91.683753;"
    AddressDistrict = AddressDistrict & "540500,540525,曲松县,29.063656,92.201066;"
    AddressDistrict = AddressDistrict & "540500,540526,措美县,28.437353,91.432347;"
    AddressDistrict = AddressDistrict & "540500,540527,洛扎县,28.385765,90.858243;"
    AddressDistrict = AddressDistrict & "540500,540528,加查县,29.140921,92.591043;"
    AddressDistrict = AddressDistrict & "540500,540529,隆子县,28.408548,92.463309;"
    AddressDistrict = AddressDistrict & "540500,540530,错那县,27.991707,91.960132;"
    AddressDistrict = AddressDistrict & "540500,540531,浪卡子县,28.96836,90.398747;"
    AddressDistrict = AddressDistrict & "540600,540602,色尼区,31.475756,92.061862;"
    AddressDistrict = AddressDistrict & "540600,540621,嘉黎县,30.640846,93.232907;"
    AddressDistrict = AddressDistrict & "540600,540622,比如县,31.479917,93.68044;"
    AddressDistrict = AddressDistrict & "540600,540623,聂荣县,32.107855,92.303659;"
    AddressDistrict = AddressDistrict & "540600,540624,安多县,32.260299,91.681879;"
    AddressDistrict = AddressDistrict & "540600,540625,申扎县,30.929056,88.709777;"
    AddressDistrict = AddressDistrict & "540600,540626,索县,31.886173,93.784964;"
    AddressDistrict = AddressDistrict & "540600,540627,班戈县,31.394578,90.011822;"
    AddressDistrict = AddressDistrict & "540600,540628,巴青县,31.918691,94.054049;"
    AddressDistrict = AddressDistrict & "540600,540629,尼玛县,31.784979,87.236646;"
    AddressDistrict = AddressDistrict & "540600,540630,双湖县,33.18698,88.838578;"
    AddressDistrict = AddressDistrict & "542500,542521,普兰县,30.291896,81.177588;"
    AddressDistrict = AddressDistrict & "542500,542522,札达县,31.478587,79.803191;"
    AddressDistrict = AddressDistrict & "542500,542523,噶尔县,32.503373,80.105005;"
    AddressDistrict = AddressDistrict & "542500,542524,日土县,33.382454,79.731937;"
    AddressDistrict = AddressDistrict & "542500,542525,革吉县,32.389192,81.142896;"
    AddressDistrict = AddressDistrict & "542500,542526,改则县,32.302076,84.062384;"
    AddressDistrict = AddressDistrict & "542500,542527,措勤县,31.016774,85.159254;"
    AddressDistrict = AddressDistrict & "610100,610102,新城区,34.26927,108.959903;"
    AddressDistrict = AddressDistrict & "610100,610103,碑林区,34.251061,108.946994;"
    AddressDistrict = AddressDistrict & "610100,610104,莲湖区,34.2656,108.933194;"
    AddressDistrict = AddressDistrict & "610100,610111,灞桥区,34.267453,109.067261;"
    AddressDistrict = AddressDistrict & "610100,610112,未央区,34.30823,108.946022;"
    AddressDistrict = AddressDistrict & "610100,610113,雁塔区,34.213389,108.926593;"
    AddressDistrict = AddressDistrict & "610100,610114,阎良区,34.662141,109.22802;"
    AddressDistrict = AddressDistrict & "610100,610115,临潼区,34.372065,109.213986;"
    AddressDistrict = AddressDistrict & "610100,610116,长安区,34.157097,108.941579;"
    AddressDistrict = AddressDistrict & "610100,610117,高陵区,34.535065,109.088896;"
    AddressDistrict = AddressDistrict & "610100,610118,鄠邑区,34.108668,108.607385;"
    AddressDistrict = AddressDistrict & "610100,610122,蓝田县,34.156189,109.317634;"
    AddressDistrict = AddressDistrict & "610100,610124,周至县,34.161532,108.216465;"
    AddressDistrict = AddressDistrict & "610200,610202,王益区,35.069098,109.075862;"
    AddressDistrict = AddressDistrict & "610200,610203,印台区,35.111927,109.100814;"
    AddressDistrict = AddressDistrict & "610200,610204,耀州区,34.910206,108.962538;"
    AddressDistrict = AddressDistrict & "610200,610222,宜君县,35.398766,109.118278;"
    AddressDistrict = AddressDistrict & "610300,610302,渭滨区,34.371008,107.144467;"
    AddressDistrict = AddressDistrict & "610300,610303,金台区,34.375192,107.149943;"
    AddressDistrict = AddressDistrict & "610300,610304,陈仓区,34.352747,107.383645;"
    AddressDistrict = AddressDistrict & "610300,610322,凤翔区,34.521668,107.400577;"
    AddressDistrict = AddressDistrict & "610300,610323,岐山县,34.44296,107.624464;"
    AddressDistrict = AddressDistrict & "610300,610324,扶风县,34.375497,107.891419;"
    AddressDistrict = AddressDistrict & "610300,610326,眉县,34.272137,107.752371;"
    AddressDistrict = AddressDistrict & "610300,610327,陇县,34.893262,106.857066;"
    AddressDistrict = AddressDistrict & "610300,610328,千阳县,34.642584,107.132987;"
    AddressDistrict = AddressDistrict & "610300,610329,麟游县,34.677714,107.796608;"
    AddressDistrict = AddressDistrict & "610300,610330,凤县,33.912464,106.525212;"
    AddressDistrict = AddressDistrict & "610300,610331,太白县,34.059215,107.316533;"
    AddressDistrict = AddressDistrict & "610400,610402,秦都区,34.329801,108.698636;"
    AddressDistrict = AddressDistrict & "610400,610403,杨陵区,34.27135,108.086348;"
    AddressDistrict = AddressDistrict & "610400,610404,渭城区,34.336847,108.730957;"
    AddressDistrict = AddressDistrict & "610400,610422,三原县,34.613996,108.943481;"
    AddressDistrict = AddressDistrict & "610400,610423,泾阳县,34.528493,108.83784;"
    AddressDistrict = AddressDistrict & "610400,610424,乾县,34.527261,108.247406;"
    AddressDistrict = AddressDistrict & "610400,610425,礼泉县,34.482583,108.428317;"
    AddressDistrict = AddressDistrict & "610400,610426,永寿县,34.692619,108.143129;"
    AddressDistrict = AddressDistrict & "610400,610428,长武县,35.206122,107.795835;"
    AddressDistrict = AddressDistrict & "610400,610429,旬邑县,35.112234,108.337231;"
    AddressDistrict = AddressDistrict & "610400,610430,淳化县,34.79797,108.581173;"
    AddressDistrict = AddressDistrict & "610400,610431,武功县,34.259732,108.212857;"
    AddressDistrict = AddressDistrict & "610400,610481,兴平市,34.297134,108.488493;"
    AddressDistrict = AddressDistrict & "610400,610482,彬州市,35.034233,108.083674;"
    AddressDistrict = AddressDistrict & "610500,610502,临渭区,34.501271,109.503299;"
    AddressDistrict = AddressDistrict & "610500,610503,华州区,34.511958,109.76141;"
    AddressDistrict = AddressDistrict & "610500,610522,潼关县,34.544515,110.24726;"
    AddressDistrict = AddressDistrict & "610500,610523,大荔县,34.795011,109.943123;"
    AddressDistrict = AddressDistrict & "610500,610524,合阳县,35.237098,110.147979;"
    AddressDistrict = AddressDistrict & "610500,610525,澄城县,35.184,109.937609;"
    AddressDistrict = AddressDistrict & "610500,610526,蒲城县,34.956034,109.589653;"
    AddressDistrict = AddressDistrict & "610500,610527,白水县,35.177291,109.594309;"
    AddressDistrict = AddressDistrict & "610500,610528,富平县,34.746679,109.187174;"
    AddressDistrict = AddressDistrict & "610500,610581,韩城市,35.475238,110.452391;"
    AddressDistrict = AddressDistrict & "610500,610582,华阴市,34.565359,110.08952;"
    AddressDistrict = AddressDistrict & "610600,610602,宝塔区,36.596291,109.49069;"
    AddressDistrict = AddressDistrict & "610600,610603,安塞区,36.86441,109.325341;"
    AddressDistrict = AddressDistrict & "610600,610621,延长县,36.578306,110.012961;"
    AddressDistrict = AddressDistrict & "610600,610622,延川县,36.882066,110.190314;"
    AddressDistrict = AddressDistrict & "610600,610625,志丹县,36.823031,108.768898;"
    AddressDistrict = AddressDistrict & "610600,610626,吴起县,36.924852,108.176976;"
    AddressDistrict = AddressDistrict & "610600,610627,甘泉县,36.277729,109.34961;"
    AddressDistrict = AddressDistrict & "610600,610628,富县,35.996495,109.384136;"
    AddressDistrict = AddressDistrict & "610600,610629,洛川县,35.762133,109.435712;"
    AddressDistrict = AddressDistrict & "610600,610630,宜川县,36.050391,110.175537;"
    AddressDistrict = AddressDistrict & "610600,610631,黄龙县,35.583276,109.83502;"
    AddressDistrict = AddressDistrict & "610600,610632,黄陵县,35.580165,109.262469;"
    AddressDistrict = AddressDistrict & "610600,610681,子长市,37.14207,109.675968;"
    AddressDistrict = AddressDistrict & "610700,610702,汉台区,33.077674,107.028233;"
    AddressDistrict = AddressDistrict & "610700,610703,南郑区,33.003341,106.942393;"
    AddressDistrict = AddressDistrict & "610700,610722,城固县,33.153098,107.329887;"
    AddressDistrict = AddressDistrict & "610700,610723,洋县,33.223283,107.549962;"
    AddressDistrict = AddressDistrict & "610700,610724,西乡县,32.987961,107.765858;"
    AddressDistrict = AddressDistrict & "610700,610725,勉县,33.155618,106.680175;"
    AddressDistrict = AddressDistrict & "610700,610726,宁强县,32.830806,106.25739;"
    AddressDistrict = AddressDistrict & "610700,610727,略阳县,33.329638,106.153899;"
    AddressDistrict = AddressDistrict & "610700,610728,镇巴县,32.535854,107.89531;"
    AddressDistrict = AddressDistrict & "610700,610729,留坝县,33.61334,106.924377;"
    AddressDistrict = AddressDistrict & "610700,610730,佛坪县,33.520745,107.988582;"
    AddressDistrict = AddressDistrict & "610800,610802,榆阳区,38.299267,109.74791;"
    AddressDistrict = AddressDistrict & "610800,610803,横山区,37.964048,109.292596;"
    AddressDistrict = AddressDistrict & "610800,610822,府谷县,39.029243,111.069645;"
    AddressDistrict = AddressDistrict & "610800,610824,靖边县,37.596084,108.80567;"
    AddressDistrict = AddressDistrict & "610800,610825,定边县,37.59523,107.601284;"
    AddressDistrict = AddressDistrict & "610800,610826,绥德县,37.507701,110.265377;"
    AddressDistrict = AddressDistrict & "610800,610827,米脂县,37.759081,110.178683;"
    AddressDistrict = AddressDistrict & "610800,610828,佳县,38.021597,110.493367;"
    AddressDistrict = AddressDistrict & "610800,610829,吴堡县,37.451925,110.739315;"
    AddressDistrict = AddressDistrict & "610800,610830,清涧县,37.087702,110.12146;"
    AddressDistrict = AddressDistrict & "610800,610831,子洲县,37.611573,110.03457;"
    AddressDistrict = AddressDistrict & "610800,610881,神木市,38.835641,110.497005;"
    AddressDistrict = AddressDistrict & "610900,610902,汉滨区,32.690817,109.029098;"
    AddressDistrict = AddressDistrict & "610900,610921,汉阴县,32.891121,108.510946;"
    AddressDistrict = AddressDistrict & "610900,610922,石泉县,33.038512,108.250512;"
    AddressDistrict = AddressDistrict & "610900,610923,宁陕县,33.312184,108.313714;"
    AddressDistrict = AddressDistrict & "610900,610924,紫阳县,32.520176,108.537788;"
    AddressDistrict = AddressDistrict & "610900,610925,岚皋县,32.31069,108.900663;"
    AddressDistrict = AddressDistrict & "610900,610926,平利县,32.387933,109.361865;"
    AddressDistrict = AddressDistrict & "610900,610927,镇坪县,31.883395,109.526437;"
    AddressDistrict = AddressDistrict & "610900,610928,旬阳市,32.833567,109.368149;"
    AddressDistrict = AddressDistrict & "610900,610929,白河县,32.809484,110.114186;"
    AddressDistrict = AddressDistrict & "611000,611002,商州区,33.869208,109.937685;"
    AddressDistrict = AddressDistrict & "611000,611021,洛南县,34.088502,110.145716;"
    AddressDistrict = AddressDistrict & "611000,611022,丹凤县,33.694711,110.33191;"
    AddressDistrict = AddressDistrict & "611000,611023,商南县,33.526367,110.885437;"
    AddressDistrict = AddressDistrict & "611000,611024,山阳县,33.530411,109.880435;"
    AddressDistrict = AddressDistrict & "611000,611025,镇安县,33.423981,109.151075;"
    AddressDistrict = AddressDistrict & "611000,611026,柞水县,33.682773,109.111249;"
    AddressDistrict = AddressDistrict & "620100,620102,城关区,36.049115,103.841032;"
    AddressDistrict = AddressDistrict & "620100,620103,七里河区,36.06673,103.784326;"
    AddressDistrict = AddressDistrict & "620100,620104,西固区,36.100369,103.622331;"
    AddressDistrict = AddressDistrict & "620100,620105,安宁区,36.10329,103.724038;"
    AddressDistrict = AddressDistrict & "620100,620111,红古区,36.344177,102.861814;"
    AddressDistrict = AddressDistrict & "620100,620121,永登县,36.734428,103.262203;"
    AddressDistrict = AddressDistrict & "620100,620122,皋兰县,36.331254,103.94933;"
    AddressDistrict = AddressDistrict & "620100,620123,榆中县,35.84443,104.114975;"
    AddressDistrict = AddressDistrict & "620300,620302,金川区,38.513793,102.187683;"
    AddressDistrict = AddressDistrict & "620300,620321,永昌县,38.247354,101.971957;"
    AddressDistrict = AddressDistrict & "620400,620402,白银区,36.545649,104.17425;"
    AddressDistrict = AddressDistrict & "620400,620403,平川区,36.72921,104.819207;"
    AddressDistrict = AddressDistrict & "620400,620421,靖远县,36.561424,104.686972;"
    AddressDistrict = AddressDistrict & "620400,620422,会宁县,35.692486,105.054337;"
    AddressDistrict = AddressDistrict & "620400,620423,景泰县,37.193519,104.066394;"
    AddressDistrict = AddressDistrict & "620500,620502,秦州区,34.578645,105.724477;"
    AddressDistrict = AddressDistrict & "620500,620503,麦积区,34.563504,105.897631;"
    AddressDistrict = AddressDistrict & "620500,620521,清水县,34.75287,106.139878;"
    AddressDistrict = AddressDistrict & "620500,620522,秦安县,34.862354,105.6733;"
    AddressDistrict = AddressDistrict & "620500,620523,甘谷县,34.747327,105.332347;"
    AddressDistrict = AddressDistrict & "620500,620524,武山县,34.721955,104.891696;"
    AddressDistrict = AddressDistrict & "620500,620525,张家川回族自治县,34.993237,106.212416;"
    AddressDistrict = AddressDistrict & "620600,620602,凉州区,37.93025,102.634492;"
    AddressDistrict = AddressDistrict & "620600,620621,民勤县,38.624621,103.090654;"
    AddressDistrict = AddressDistrict & "620600,620622,古浪县,37.470571,102.898047;"
    AddressDistrict = AddressDistrict & "620600,620623,天祝藏族自治县,36.971678,103.142034;"
    AddressDistrict = AddressDistrict & "620700,620702,甘州区,38.931774,100.454862;"
    AddressDistrict = AddressDistrict & "620700,620721,肃南裕固族自治县,38.837269,99.617086;"
    AddressDistrict = AddressDistrict & "620700,620722,民乐县,38.434454,100.816623;"
    AddressDistrict = AddressDistrict & "620700,620723,临泽县,39.152151,100.166333;"
    AddressDistrict = AddressDistrict & "620700,620724,高台县,39.376308,99.81665;"
    AddressDistrict = AddressDistrict & "620700,620725,山丹县,38.784839,101.088442;"
    AddressDistrict = AddressDistrict & "620800,620802,崆峒区,35.54173,106.684223;"
    AddressDistrict = AddressDistrict & "620800,620821,泾川县,35.335283,107.365218;"
    AddressDistrict = AddressDistrict & "620800,620822,灵台县,35.064009,107.620587;"
    AddressDistrict = AddressDistrict & "620800,620823,崇信县,35.304533,107.031253;"
    AddressDistrict = AddressDistrict & "620800,620825,庄浪县,35.203428,106.041979;"
    AddressDistrict = AddressDistrict & "620800,620826,静宁县,35.525243,105.733489;"
    AddressDistrict = AddressDistrict & "620800,620881,华亭市,35.215341,106.649308;"
    AddressDistrict = AddressDistrict & "620900,620902,肃州区,39.743858,98.511155;"
    AddressDistrict = AddressDistrict & "620900,620921,金塔县,39.983036,98.902959;"
    AddressDistrict = AddressDistrict & "620900,620922,瓜州县,40.516525,95.780591;"
    AddressDistrict = AddressDistrict & "620900,620923,肃北蒙古族自治县,39.51224,94.87728;"
    AddressDistrict = AddressDistrict & "620900,620924,阿克塞哈萨克族自治县,39.631642,94.337642;"
    AddressDistrict = AddressDistrict & "620900,620981,玉门市,40.28682,97.037206;"
    AddressDistrict = AddressDistrict & "620900,620982,敦煌市,40.141119,94.664279;"
    AddressDistrict = AddressDistrict & "621000,621002,西峰区,35.733713,107.638824;"
    AddressDistrict = AddressDistrict & "621000,621021,庆城县,36.013504,107.885664;"
    AddressDistrict = AddressDistrict & "621000,621022,环县,36.569322,107.308754;"
    AddressDistrict = AddressDistrict & "621000,621023,华池县,36.457304,107.986288;"
    AddressDistrict = AddressDistrict & "621000,621024,合水县,35.819005,108.019865;"
    AddressDistrict = AddressDistrict & "621000,621025,正宁县,35.490642,108.361068;"
    AddressDistrict = AddressDistrict & "621000,621026,宁县,35.50201,107.921182;"
    AddressDistrict = AddressDistrict & "621000,621027,镇原县,35.677806,107.195706;"
    AddressDistrict = AddressDistrict & "621100,621102,安定区,35.579764,104.62577;"
    AddressDistrict = AddressDistrict & "621100,621121,通渭县,35.208922,105.250102;"
    AddressDistrict = AddressDistrict & "621100,621122,陇西县,35.003409,104.637554;"
    AddressDistrict = AddressDistrict & "621100,621123,渭源县,35.133023,104.211742;"
    AddressDistrict = AddressDistrict & "621100,621124,临洮县,35.376233,103.862186;"
    AddressDistrict = AddressDistrict & "621100,621125,漳县,34.848642,104.466756;"
    AddressDistrict = AddressDistrict & "621100,621126,岷县,34.439105,104.039882;"
    AddressDistrict = AddressDistrict & "621200,621202,武都区,33.388155,104.929866;"
    AddressDistrict = AddressDistrict & "621200,621221,成县,33.739863,105.734434;"
    AddressDistrict = AddressDistrict & "621200,621222,文县,32.942171,104.682448;"
    AddressDistrict = AddressDistrict & "621200,621223,宕昌县,34.042655,104.394475;"
    AddressDistrict = AddressDistrict & "621200,621224,康县,33.328266,105.609534;"
    AddressDistrict = AddressDistrict & "621200,621225,西和县,34.013718,105.299737;"
    AddressDistrict = AddressDistrict & "621200,621226,礼县,34.189387,105.181616;"
    AddressDistrict = AddressDistrict & "621200,621227,徽县,33.767785,106.085632;"
    AddressDistrict = AddressDistrict & "621200,621228,两当县,33.910729,106.306959;"
    AddressDistrict = AddressDistrict & "622900,622901,临夏市,35.59941,103.211634;"
    AddressDistrict = AddressDistrict & "622900,622921,临夏县,35.49236,102.993873;"
    AddressDistrict = AddressDistrict & "622900,622922,康乐县,35.371906,103.709852;"
    AddressDistrict = AddressDistrict & "622900,622923,永靖县,35.938933,103.319871;"
    AddressDistrict = AddressDistrict & "622900,622924,广河县,35.481688,103.576188;"
    AddressDistrict = AddressDistrict & "622900,622925,和政县,35.425971,103.350357;"
    AddressDistrict = AddressDistrict & "622900,622926,东乡族自治县,35.66383,103.389568;"
    AddressDistrict = AddressDistrict & "622900,622927,积石山保安族东乡族撒拉族自治县,35.712906,102.877473;"
    AddressDistrict = AddressDistrict & "623000,623001,合作市,34.985973,102.91149;"
    AddressDistrict = AddressDistrict & "623000,623021,临潭县,34.69164,103.353054;"
    AddressDistrict = AddressDistrict & "623000,623022,卓尼县,34.588165,103.508508;"
    AddressDistrict = AddressDistrict & "623000,623023,舟曲县,33.782964,104.370271;"
    AddressDistrict = AddressDistrict & "623000,623024,迭部县,34.055348,103.221009;"
    AddressDistrict = AddressDistrict & "623000,623025,玛曲县,33.998068,102.075767;"
    AddressDistrict = AddressDistrict & "623000,623026,碌曲县,34.589591,102.488495;"
    AddressDistrict = AddressDistrict & "623000,623027,夏河县,35.200853,102.520743;"
    AddressDistrict = AddressDistrict & "630100,630102,城东区,36.616043,101.796095;"
    AddressDistrict = AddressDistrict & "630100,630103,城中区,36.621181,101.784554;"
    AddressDistrict = AddressDistrict & "630100,630104,城西区,36.628323,101.763649;"
    AddressDistrict = AddressDistrict & "630100,630105,城北区,36.648448,101.761297;"
    AddressDistrict = AddressDistrict & "630100,630106,湟中区,36.500419,101.569475;"
    AddressDistrict = AddressDistrict & "630100,630121,大通回族土族自治县,36.931343,101.684183;"
    AddressDistrict = AddressDistrict & "630100,630123,湟源县,36.684818,101.263435;"
    AddressDistrict = AddressDistrict & "630200,630202,乐都区,36.480291,102.402431;"
    AddressDistrict = AddressDistrict & "630200,630203,平安区,36.502714,102.104295;"
    AddressDistrict = AddressDistrict & "630200,630222,民和回族土族自治县,36.329451,102.804209;"
    AddressDistrict = AddressDistrict & "630200,630223,互助土族自治县,36.83994,101.956734;"
    AddressDistrict = AddressDistrict & "630200,630224,化隆回族自治县,36.098322,102.262329;"
    AddressDistrict = AddressDistrict & "630200,630225,循化撒拉族自治县,35.847247,102.486534;"
    AddressDistrict = AddressDistrict & "632200,632221,门源回族自治县,37.376627,101.618461;"
    AddressDistrict = AddressDistrict & "632200,632222,祁连县,38.175409,100.249778;"
    AddressDistrict = AddressDistrict & "632200,632223,海晏县,36.959542,100.90049;"
    AddressDistrict = AddressDistrict & "632200,632224,刚察县,37.326263,100.138417;"
    AddressDistrict = AddressDistrict & "632300,632301,同仁市,35.516337,102.017604;"
    AddressDistrict = AddressDistrict & "632300,632322,尖扎县,35.938205,102.031953;"
    AddressDistrict = AddressDistrict & "632300,632323,泽库县,35.036842,101.469343;"
    AddressDistrict = AddressDistrict & "632300,632324,河南蒙古族自治县,34.734522,101.611877;"
    AddressDistrict = AddressDistrict & "632500,632521,共和县,36.280286,100.619597;"
    AddressDistrict = AddressDistrict & "632500,632522,同德县,35.254492,100.579465;"
    AddressDistrict = AddressDistrict & "632500,632523,贵德县,36.040456,101.431856;"
    AddressDistrict = AddressDistrict & "632500,632524,兴海县,35.58909,99.986963;"
    AddressDistrict = AddressDistrict & "632500,632525,贵南县,35.587085,100.74792;"
    AddressDistrict = AddressDistrict & "632600,632621,玛沁县,34.473386,100.243531;"
    AddressDistrict = AddressDistrict & "632600,632622,班玛县,32.931589,100.737955;"
    AddressDistrict = AddressDistrict & "632600,632623,甘德县,33.966987,99.902589;"
    AddressDistrict = AddressDistrict & "632600,632624,达日县,33.753259,99.651715;"
    AddressDistrict = AddressDistrict & "632600,632625,久治县,33.430217,101.484884;"
    AddressDistrict = AddressDistrict & "632600,632626,玛多县,34.91528,98.211343;"
    AddressDistrict = AddressDistrict & "632700,632701,玉树市,33.00393,97.008762;"
    AddressDistrict = AddressDistrict & "632700,632722,杂多县,32.891886,95.293423;"
    AddressDistrict = AddressDistrict & "632700,632723,称多县,33.367884,97.110893;"
    AddressDistrict = AddressDistrict & "632700,632724,治多县,33.852322,95.616843;"
    AddressDistrict = AddressDistrict & "632700,632725,囊谦县,32.203206,96.479797;"
    AddressDistrict = AddressDistrict & "632700,632726,曲麻莱县,34.12654,95.800674;"
    AddressDistrict = AddressDistrict & "632800,632801,格尔木市,36.401541,94.905777;"
    AddressDistrict = AddressDistrict & "632800,632802,德令哈市,37.374555,97.370143;"
    AddressDistrict = AddressDistrict & "632800,632803,茫崖市,38.247117,90.855955;"
    AddressDistrict = AddressDistrict & "632800,632821,乌兰县,36.930389,98.479852;"
    AddressDistrict = AddressDistrict & "632800,632822,都兰县,36.298553,98.089161;"
    AddressDistrict = AddressDistrict & "632800,632823,天峻县,37.29906,99.02078;"
    AddressDistrict = AddressDistrict & "632800,632825,海西蒙古族藏族自治州直辖,37.853631,95.357233;"
    AddressDistrict = AddressDistrict & "640100,640104,兴庆区,38.46747,106.278393;"
    AddressDistrict = AddressDistrict & "640100,640105,西夏区,38.492424,106.132116;"
    AddressDistrict = AddressDistrict & "640100,640106,金凤区,38.477353,106.228486;"
    AddressDistrict = AddressDistrict & "640100,640121,永宁县,38.28043,106.253781;"
    AddressDistrict = AddressDistrict & "640100,640122,贺兰县,38.554563,106.345904;"
    AddressDistrict = AddressDistrict & "640100,640181,灵武市,38.094058,106.334701;"
    AddressDistrict = AddressDistrict & "640200,640202,大武口区,39.014158,106.376651;"
    AddressDistrict = AddressDistrict & "640200,640205,惠农区,39.230094,106.775513;"
    AddressDistrict = AddressDistrict & "640200,640221,平罗县,38.90674,106.54489;"
    AddressDistrict = AddressDistrict & "640300,640302,利通区,37.985967,106.199419;"
    AddressDistrict = AddressDistrict & "640300,640303,红寺堡区,37.421616,106.067315;"
    AddressDistrict = AddressDistrict & "640300,640323,盐池县,37.784222,107.40541;"
    AddressDistrict = AddressDistrict & "640300,640324,同心县,36.9829,105.914764;"
    AddressDistrict = AddressDistrict & "640300,640381,青铜峡市,38.021509,106.075395;"
    AddressDistrict = AddressDistrict & "640400,640402,原州区,36.005337,106.28477;"
    AddressDistrict = AddressDistrict & "640400,640422,西吉县,35.965384,105.731801;"
    AddressDistrict = AddressDistrict & "640400,640423,隆德县,35.618234,106.12344;"
    AddressDistrict = AddressDistrict & "640400,640424,泾源县,35.49344,106.338674;"
    AddressDistrict = AddressDistrict & "640400,640425,彭阳县,35.849975,106.641512;"
    AddressDistrict = AddressDistrict & "640500,640502,沙坡头区,37.514564,105.190536;"
    AddressDistrict = AddressDistrict & "640500,640521,中宁县,37.489736,105.675784;"
    AddressDistrict = AddressDistrict & "640500,640522,海原县,36.562007,105.647323;"
    AddressDistrict = AddressDistrict & "650100,650102,天山区,43.796428,87.620116;"
    AddressDistrict = AddressDistrict & "650100,650103,沙依巴克区,43.788872,87.596639;"
    AddressDistrict = AddressDistrict & "650100,650104,新市区,43.870882,87.560653;"
    AddressDistrict = AddressDistrict & "650100,650105,水磨沟区,43.816747,87.613093;"
    AddressDistrict = AddressDistrict & "650100,650106,头屯河区,43.876053,87.425823;"
    AddressDistrict = AddressDistrict & "650100,650107,达坂城区,43.36181,88.30994;"
    AddressDistrict = AddressDistrict & "650100,650109,米东区,43.960982,87.691801;"
    AddressDistrict = AddressDistrict & "650100,650121,乌鲁木齐县,43.982546,87.505603;"
    AddressDistrict = AddressDistrict & "650200,650202,独山子区,44.327207,84.882267;"
    AddressDistrict = AddressDistrict & "650200,650203,克拉玛依区,45.600477,84.868918;"
    AddressDistrict = AddressDistrict & "650200,650204,白碱滩区,45.689021,85.129882;"
    AddressDistrict = AddressDistrict & "650200,650205,乌尔禾区,46.08776,85.697767;"
    AddressDistrict = AddressDistrict & "650400,650402,高昌区,42.947627,89.182324;"
    AddressDistrict = AddressDistrict & "650400,650421,鄯善县,42.865503,90.212692;"
    AddressDistrict = AddressDistrict & "650400,650422,托克逊县,42.793536,88.655771;"
    AddressDistrict = AddressDistrict & "650500,650502,伊州区,42.833888,93.509174;"
    AddressDistrict = AddressDistrict & "650500,650521,巴里坤哈萨克自治县,43.599032,93.021795;"
    AddressDistrict = AddressDistrict & "650500,650522,伊吾县,43.252012,94.692773;"
    AddressDistrict = AddressDistrict & "652300,652301,昌吉市,44.013183,87.304112;"
    AddressDistrict = AddressDistrict & "652300,652302,阜康市,44.152153,87.98384;"
    AddressDistrict = AddressDistrict & "652300,652323,呼图壁县,44.189342,86.888613;"
    AddressDistrict = AddressDistrict & "652300,652324,玛纳斯县,44.305625,86.217687;"
    AddressDistrict = AddressDistrict & "652300,652325,奇台县,44.021996,89.591437;"
    AddressDistrict = AddressDistrict & "652300,652327,吉木萨尔县,43.997162,89.181288;"
    AddressDistrict = AddressDistrict & "652300,652328,木垒哈萨克自治县,43.832442,90.282833;"
    AddressDistrict = AddressDistrict & "652700,652701,博乐市,44.903087,82.072237;"
    AddressDistrict = AddressDistrict & "652700,652702,阿拉山口市,45.16777,82.569389;"
    AddressDistrict = AddressDistrict & "652700,652722,精河县,44.605645,82.892938;"
    AddressDistrict = AddressDistrict & "652700,652723,温泉县,44.973751,81.03099;"
    AddressDistrict = AddressDistrict & "652800,652801,库尔勒市,41.763122,86.145948;"
    AddressDistrict = AddressDistrict & "652800,652822,轮台县,41.781266,84.248542;"
    AddressDistrict = AddressDistrict & "652800,652823,尉犁县,41.337428,86.263412;"
    AddressDistrict = AddressDistrict & "652800,652824,若羌县,39.023807,88.168807;"
    AddressDistrict = AddressDistrict & "652800,652825,且末县,38.138562,85.532629;"
    AddressDistrict = AddressDistrict & "652800,652826,焉耆回族自治县,42.064349,86.5698;"
    AddressDistrict = AddressDistrict & "652800,652827,和静县,42.31716,86.391067;"
    AddressDistrict = AddressDistrict & "652800,652828,和硕县,42.268863,86.864947;"
    AddressDistrict = AddressDistrict & "652800,652829,博湖县,41.980166,86.631576;"
    AddressDistrict = AddressDistrict & "652900,652901,阿克苏市,41.171272,80.2629;"
    AddressDistrict = AddressDistrict & "652900,652902,库车市,41.717141,82.96304;"
    AddressDistrict = AddressDistrict & "652900,652922,温宿县,41.272995,80.243273;"
    AddressDistrict = AddressDistrict & "652900,652924,沙雅县,41.226268,82.78077;"
    AddressDistrict = AddressDistrict & "652900,652925,新和县,41.551176,82.610828;"
    AddressDistrict = AddressDistrict & "652900,652926,拜城县,41.796101,81.869881;"
    AddressDistrict = AddressDistrict & "652900,652927,乌什县,41.21587,79.230805;"
    AddressDistrict = AddressDistrict & "652900,652928,阿瓦提县,40.638422,80.378426;"
    AddressDistrict = AddressDistrict & "652900,652929,柯坪县,40.50624,79.04785;"
    AddressDistrict = AddressDistrict & "653000,653001,阿图什市,39.712898,76.173939;"
    AddressDistrict = AddressDistrict & "653000,653022,阿克陶县,39.147079,75.945159;"
    AddressDistrict = AddressDistrict & "653000,653023,阿合奇县,40.937567,78.450164;"
    AddressDistrict = AddressDistrict & "653000,653024,乌恰县,39.716633,75.25969;"
    AddressDistrict = AddressDistrict & "653100,653101,喀什市,39.467861,75.98838;"
    AddressDistrict = AddressDistrict & "653100,653121,疏附县,39.378306,75.863075;"
    AddressDistrict = AddressDistrict & "653100,653122,疏勒县,39.399461,76.053653;"
    AddressDistrict = AddressDistrict & "653100,653123,英吉沙县,38.929839,76.174292;"
    AddressDistrict = AddressDistrict & "653100,653124,泽普县,38.191217,77.273593;"
    AddressDistrict = AddressDistrict & "653100,653125,莎车县,38.414499,77.248884;"
    AddressDistrict = AddressDistrict & "653100,653126,叶城县,37.884679,77.420353;"
    AddressDistrict = AddressDistrict & "653100,653127,麦盖提县,38.903384,77.651538;"
    AddressDistrict = AddressDistrict & "653100,653128,岳普湖县,39.235248,76.7724;"
    AddressDistrict = AddressDistrict & "653100,653129,伽师县,39.494325,76.741982;"
    AddressDistrict = AddressDistrict & "653100,653130,巴楚县,39.783479,78.55041;"
    AddressDistrict = AddressDistrict & "653100,653131,塔什库尔干塔吉克自治县,37.775437,75.228068;"
    AddressDistrict = AddressDistrict & "653200,653201,和田市,37.108944,79.927542;"
    AddressDistrict = AddressDistrict & "653200,653221,和田县,37.120031,79.81907;"
    AddressDistrict = AddressDistrict & "653200,653222,墨玉县,37.271511,79.736629;"
    AddressDistrict = AddressDistrict & "653200,653223,皮山县,37.616332,78.282301;"
    AddressDistrict = AddressDistrict & "653200,653224,洛浦县,37.074377,80.184038;"
    AddressDistrict = AddressDistrict & "653200,653225,策勒县,37.001672,80.803572;"
    AddressDistrict = AddressDistrict & "653200,653226,于田县,36.854628,81.667845;"
    AddressDistrict = AddressDistrict & "653200,653227,民丰县,37.064909,82.692354;"
    AddressDistrict = AddressDistrict & "654000,654002,伊宁市,43.922209,81.316343;"
    AddressDistrict = AddressDistrict & "654000,654003,奎屯市,44.423445,84.901602;"
    AddressDistrict = AddressDistrict & "654000,654004,霍尔果斯市,44.201669,80.420759;"
    AddressDistrict = AddressDistrict & "654000,654021,伊宁县,43.977876,81.524671;"
    AddressDistrict = AddressDistrict & "654000,654022,察布查尔锡伯自治县,43.838883,81.150874;"
    AddressDistrict = AddressDistrict & "654000,654023,霍城县,44.049912,80.872508;"
    AddressDistrict = AddressDistrict & "654000,654024,巩留县,43.481618,82.227044;"
    AddressDistrict = AddressDistrict & "654000,654025,新源县,43.434249,83.258493;"
    AddressDistrict = AddressDistrict & "654000,654026,昭苏县,43.157765,81.126029;"
    AddressDistrict = AddressDistrict & "654000,654027,特克斯县,43.214861,81.840058;"
    AddressDistrict = AddressDistrict & "654000,654028,尼勒克县,43.789737,82.504119;"
    AddressDistrict = AddressDistrict & "654200,654201,塔城市,46.746281,82.983988;"
    AddressDistrict = AddressDistrict & "654200,654202,乌苏市,44.430115,84.677624;"
    AddressDistrict = AddressDistrict & "654200,654221,额敏县,46.522555,83.622118;"
    AddressDistrict = AddressDistrict & "654200,654223,沙湾市,44.329544,85.622508;"
    AddressDistrict = AddressDistrict & "654200,654224,托里县,45.935863,83.60469;"
    AddressDistrict = AddressDistrict & "654200,654225,裕民县,46.202781,82.982157;"
    AddressDistrict = AddressDistrict & "654200,654226,和布克赛尔蒙古自治县,46.793001,85.733551;"
    AddressDistrict = AddressDistrict & "654300,654301,阿勒泰市,47.848911,88.138743;"
    AddressDistrict = AddressDistrict & "654300,654321,布尔津县,47.70453,86.86186;"
    AddressDistrict = AddressDistrict & "654300,654322,富蕴县,46.993106,89.524993;"
    AddressDistrict = AddressDistrict & "654300,654323,福海县,47.113128,87.494569;"
    AddressDistrict = AddressDistrict & "654300,654324,哈巴河县,48.059284,86.418964;"
    AddressDistrict = AddressDistrict & "654300,654325,青河县,46.672446,90.381561;"
    AddressDistrict = AddressDistrict & "654300,654326,吉木乃县,47.434633,85.876064;"
    '修正县级市和湾湾的层级数据
    AddressDistrict = AddressDistrict & "710000,710000,台湾,25.044332,121.509062;"
    AddressDistrict = AddressDistrict & "419001,419001,济源,35.090378,112.590047;"
    AddressDistrict = AddressDistrict & "429004,429004,仙桃,30.364953,113.453974;"
    AddressDistrict = AddressDistrict & "429005,429005,潜江,30.421215,112.896866;"
    AddressDistrict = AddressDistrict & "429006,429006,天门,30.653061,113.165862;"
    AddressDistrict = AddressDistrict & "429021,429021,神农架林区,31.744449,110.671525;"
    AddressDistrict = AddressDistrict & "441900,441900,东莞,23.046237,113.746262;"
    AddressDistrict = AddressDistrict & "442000,442000,中山,22.521113,113.382391;"
    AddressDistrict = AddressDistrict & "460400,460400,儋州,19.517486,109.576782;"
    AddressDistrict = AddressDistrict & "469001,469001,五指山,18.776921,109.516662;"
    AddressDistrict = AddressDistrict & "469002,469002,琼海,19.246011,110.466785;"
    AddressDistrict = AddressDistrict & "469005,469005,文昌,19.612986,110.753975;"
    AddressDistrict = AddressDistrict & "469006,469006,万宁,18.796216,110.388793;"
    AddressDistrict = AddressDistrict & "469007,469007,东方,19.10198,108.653789;"
    AddressDistrict = AddressDistrict & "469021,469021,定安县,19.684966,110.349235;"
    AddressDistrict = AddressDistrict & "469022,469022,屯昌县,19.362916,110.102773;"
    AddressDistrict = AddressDistrict & "469023,469023,澄迈县,19.737095,110.007147;"
    AddressDistrict = AddressDistrict & "469024,469024,临高县,19.908293,109.687697;"
    AddressDistrict = AddressDistrict & "469025,469025,白沙黎族自治县,19.224584,109.452606;"
    AddressDistrict = AddressDistrict & "469026,469026,昌江黎族自治县,19.260968,109.053351;"
    AddressDistrict = AddressDistrict & "469027,469027,乐东黎族自治县,18.74758,109.175444;"
    AddressDistrict = AddressDistrict & "469028,469028,陵水黎族自治县,18.505006,110.037218;"
    AddressDistrict = AddressDistrict & "469029,469029,保亭黎族苗族自治县,18.636371,109.70245;"
    AddressDistrict = AddressDistrict & "469030,469030,琼中黎族苗族自治县,19.03557,109.839996;"
    AddressDistrict = AddressDistrict & "620200,620200,嘉峪关,39.786529,98.277304;"
    AddressDistrict = AddressDistrict & "659001,659001,石河子,44.305886,86.041075;"
    AddressDistrict = AddressDistrict & "659002,659002,阿拉尔,40.541914,81.285884;"
    AddressDistrict = AddressDistrict & "659003,659003,图木舒克,39.867316,79.077978;"
    AddressDistrict = AddressDistrict & "659004,659004,五家渠,44.167401,87.526884;"
    AddressDistrict = AddressDistrict & "659005,659005,北屯,47.353177,87.824932;"
    AddressDistrict = AddressDistrict & "659006,659006,铁门关,41.827251,85.501218;"
    AddressDistrict = AddressDistrict & "659007,659007,双河,44.840524,82.353656;"
    AddressDistrict = AddressDistrict & "659008,659008,可克达拉,43.6832,80.63579;"
    AddressDistrict = AddressDistrict & "659009,659009,昆玉,37.207994,79.287372;"
    AddressDistrict = AddressDistrict & "659010,659010,胡杨河,44.69288853,84.8275959"

End Function

Public Function DateStatusDict(Optional ByVal intX As Long = 7) As Object
    '将日志状态写入字典
    '字典key为日期，值为行字典 包含 index：索引从0开始； status：日期状态；modx：索引 与 参数 intX 取模后的余数；weekday：周几 1-7 表示 日-六
    'status：1=>工作日；2=>补班；3=>周末；4=>假期；5=>法定节日
    '2016-2023年日期状态,每行表示一年,每年定了假期增加在后面
    Const dateStatusList As String = "5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,5,5,5,4,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4,5,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,5,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,5,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,5,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,5,5,5,4,4,4,4,2,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4," & _
                                     "5,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,4,5,5,5,4,4,4,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4,5,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4,5,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4,5,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,5,5,5,5,4,4,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4," & _
                                     "5,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,4,5,5,5,4,4,4,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,5,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4,5,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4,5,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4,5,1,1,1,1,2,2,5,5,5,4,4,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4," & _
                                     "5,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,2,4,5,5,5,4,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,5,4,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,5,5,5,4,4,4,4,1,1,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1," & _
                                     "5,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,4,5,5,5,4,4,4,4,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,5,4,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,5,4,4,4,4,1,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,5,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,5,5,5,5,4,4,4,4,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1," & _
                                     "5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,4,5,5,5,4,4,4,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,5,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,1,5,4,4,4,4,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,4,5,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4,5,1,1,1,3,2,1,1,1,1,5,5,5,4,4,4,4,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1," & _
                                     "5,4,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,2,4,5,5,5,4,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,2,4,4,5,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,1,4,5,4,4,4,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,5,4,4,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,5,4,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,5,5,5,4,4,4,4,2,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4," & _
                                     "5,4,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,4,5,5,5,4,4,4,2,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,5,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,2,1,1,1,1,1,4,4,5,4,4,1,1,2,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,5,4,4,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,5,4,5,5,5,4,4,4,2,2,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3,1,1,1,1,1,3,3"

    Const DateStart As Date = #1/1/2016#

    Dim dateAC As Date
    Dim ArrDateStatus
    Dim d As Long, days As Long, index As Long
    Dim rowDict As Object
    
    Set DateStatusDict = CreateObject("Scripting.Dictionary") ' 初始化表名称字典
    ArrDateStatus = Split(dateStatusList, ",")
    days = UBound(ArrDateStatus)
    
    For d = 0 To days
    
        dateAC = DateStart + d
        
        Set rowDict = CreateObject("Scripting.Dictionary")
        
        rowDict.Add "index", d '索引从0开始
        rowDict.Add "status", CLng(ArrDateStatus(d)) '日期状态
        rowDict.Add "modx", d Mod intX '索引 与 参数 intX 取模后的余数 默认为 7
        rowDict.Add "weekday", Weekday(dateAC) '周几 1-7 表示 日-六

        DateStatusDict.Add dateAC, rowDict '添加到主字典中
        
        Set rowDict = Nothing
    Next

End Function

Public Function GenderDict() As Object
    '前面 448 个 ID 的性别根据头像已经确认。
    Const GenderList As String = "1,1,1,0,1,1,1,0,1,1,0,1,0,1,0,1,0,0,0,0,0,1,0,1,1,0,1,0,1,0,1,1,0,0,1,0,1,1,1,0,0,0,0,1,0,1,1,1,1,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,1,1,1,0,1,0,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,1,1,1,0,1,0,1,0,0,1,0,0,0,0,1,0,0,0,0,1,1,1,1,0,0,0,0,0,0,1,0,1,1,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,1,0,0,0,0,0,0,1,1,1,0,0,1,1,0,0,0,1,1,0,0,1,0,0,0,1,0,1,0,0,1,0,1,0,0,1,0,1,1,0,1,0,0,0,0,0,0,1,0,0,1,0,0,0,1,0,0,0,0,1,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,1,1,0,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,1,0,0,1,0,0,1,1,1,0,0,0,1,0,0,0,0,1,1,1,0,0,0,0,0,1,0,1,0,1,0,1,1,0,0,0,0,0,0,0,1,1,0,0,0,1,1,0,1,0,0,1,0,1,0,0,1,0,0,1,1,1,0,1,1,0,0,0,1,1,0,1,1,0,0,0,0,1,0,0,0,1,0,1,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1,1,0,0,0,0,1,0,0,1,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,1,0,1,0,1,1,0,0,0,0,0,0,0,1,1,0,0,0,1,1,0,1,0,0,1,0,1,0,0,1,0,0,0,0,1,0,1,1,0,0,0,1,1,0,1,1,0,0,0,0,1,0,0,0,1,0,1,0,0,0,0,0,1,1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1"

    Const IDStart As Long = 10001

    Dim ID As Long, gender As String
    Dim ArrGender
    Dim n As Long, num As Long
    
    Set GenderDict = CreateObject("Scripting.Dictionary") ' 初始化表名称字典
    ArrGender = Split(GenderList, ",")
    num = UBound(ArrGender)
    
    For n = 0 To num
        ID = IDStart + n
        gender = "男"
        If ArrGender(n) = 0 Then gender = "女"
        GenderDict.Add ID, gender '添加到主字典中
    Next

End Function

Public Function AddMonths(d As Date, n As Long) As Date
    '日期增加月份
    AddMonths = DateSerial(Year(d), Month(d) + n, day(d))
End Function

Public Function GetDaysInMonth(d As Date) As Long
    '获取当月有多少天
    GetDaysInMonth = day(DateSerial(Year(d), Month(d) + 1, 1) - 1)
End Function

Public Function GetMonthStart(d As Date) As Date
    '根据日期获取月初日期
    GetMonthStart = DateSerial(Year(d), Month(d), 1)
End Function

Public Function DateDiffInMonths(DateStart As Date, DateEnd As Date) As Long
    '两个日期间的月份差异数量
    Dim YearsDiff As Long
    Dim MonthsDiff As Long
    
    YearsDiff = Year(DateEnd) - Year(DateStart)
    MonthsDiff = Month(DateEnd) - Month(DateStart)
    
    DateDiffInMonths = (YearsDiff * 12) + MonthsDiff
    
    ' 根据需要调整天数差异处理逻辑
    If day(DateEnd) < day(DateStart) Then
        DateDiffInMonths = DateDiffInMonths - 1
    End If
End Function

Public Function AddDictByKey(ByRef dictTarget As Object, ByVal KeyTarget As String, ByVal newValue As Long) As Object
    '根据字典值累加
    Dim oldValue As Long
    If dictTarget.Exists(KeyTarget) Then
        oldValue = dictTarget(KeyTarget)
        dictTarget(KeyTarget) = oldValue + newValue
    Else
        dictTarget.Add KeyTarget, newValue
    End If
    
End Function

Public Function Main()
    ' 入口函数 生成数据

    Dim t As Double
    t = timer
    Dim i As Long
    Dim pbRndInt  As Integer
    Dim pbLeftInt  As Integer
    Dim key As Variant
    Dim keyStr As String
    Dim valueStr As String
    
    productQuantity = 200       '产品数量；建议ShopQuantity∈[7,1688]。
    ShopQuantity = 1            '门店数量；建议ShopQuantity∈[1,390]。
    MaxInventoryDays = 14       '入库间隔最大数；建议ShopQuantity∈[5,20]。
    
    InitTables
    ' 遍历字典的键和值
    For Each key In TableNameDict.Keys
        
        keyStr = CStr(key)
        valueStr = CStr(TableNameDict(key))
'        Debug.Print valueStr
        
        ' ADO 新建表
        Call TableADO(keyStr, SQLDrop(keyStr), valueStr)
        
    Next key

    DataTableRegion             ' 大区
    DataTableProvince           ' 省份
    DataTableCity               ' 城市
    DataTableDistrict           ' 区县
    DataTableProduct            ' 产品
    DataTableShop               ' 门店
    DataTableEmployeeExecutives ' 员工表高管
    DataTableOrg                ' 组织
    DataTableShopRD             ' 门店租赁和装修
    DataTableEmployeeRegular    ' 员工表一线
    DataTableCustomer           ' 客户
    DataTableSOS                ' 入库、订单主表、订单子表
    DataTableSaleTarget         ' 销售预算
    DataTableLaborCost          ' 人工成本


    Application.RefreshDatabaseWindow

    MsgBox "完成，用时：" & Round(timer - t, 2) & "秒！"

End Function


