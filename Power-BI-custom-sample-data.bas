Attribute VB_Name = "demo_jiaopengzi_data"
Option Compare Database
Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'1、作者：焦棚子
'2、邮箱：jiaopengzi@qq.com
'3、博客：www.jiaopengzi.com
'4、CPU：Intel(R) Core(TM) i7-8750H CPU @ 2.20GHz   2.21 GHz
'5、内存：RAM 24.0 GB
'6、电脑配置 + N1=300的配置：大约需要1111秒，每秒按照业务逻辑生成约3500行数据；构成388+万行demo数据，基本满足实战学习所用。
'=====================================================================================

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'数据配置，代码行数1926行。
Public Const N1 As Long = 30 '门店数量；建议N1∈[5,390]。
'01、NewData            配置好上述三个参数后，调用所有函数生成demodata数据，建议第一次尝试按照 N1=5，大约20秒。
'02、TableNameN         所有表的命名管理。
'03、SqlCN              创建所有的sql。
'04、SqlDN              删除所有表的sql。
'05、TableADO           ADO创建表。
'06、DataTableD0        生成大区表。
'07、DataTableD1        生成省份表。
'08、DataTableD2        生成城市表。
'09、DataTableT0        生成产品表。
'10、DataTableT1        生成客户表，与N1相关。
'11、DataTableT2        生成客户表。
'12、DataTableT345      生成入库表、订单主表、订单子表。
'13、DataTableT6        生成销售目标表。
'14、FirstName          生成随机姓名的名。
'17、LastName           生成随机姓名的姓
'16、AddressProvince    所有省区数据，包含名称，坐标等。
'17、AddressCity        所有地市数据，包含名称，坐标等。
'=====================================================================================


Public Function NewData()

    Dim t As Double
    t = Timer
    Dim i As Long

    For i = 0 To 9
        Call TableADO(TableNameN(i), SqlDN(i), SqlCN(i))
    Next
    
    Call DataTableD0
    Call DataTableD1
    Call DataTableD2
    Call DataTableT0
    Call DataTableT1
    Call DataTableT2
    Call DataTableT345
    Call DataTableT6

    Application.RefreshDatabaseWindow
    MsgBox "完成，用时：" & Round(Timer - t, 2) & "秒！"

End Function


Public Function TableNameN(N As Long) As String

    Select Case N
        Case 0
            TableNameN = "T00_产品表"
        Case 1
            TableNameN = "T01_门店表"
        Case 2
            TableNameN = "T02_客户表"
        Case 3
            TableNameN = "T03_入库信息表"
        Case 4
            TableNameN = "T04_订单主表"
        Case 5
            TableNameN = "T05_订单子表"
        Case 6
            TableNameN = "T06_销售目标表"
        Case 7
            TableNameN = "D00_大区表"
        Case 8
            TableNameN = "D01_省份表"
        Case 9
            TableNameN = "D02_城市表"
    End Select

End Function

Public Function SqlCN(N As Long) As String

    Select Case N
        '产品
        Case 0
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_产品编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_产品分类            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_产品名称            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_产品销售价格        FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_产品成本价格        FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '门店
        Case 1
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_门店编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_门店名称            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_门店负责人          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_开店日期            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_城市ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_城市                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_纬度                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_08_经度                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_09_关店日期            DATE          NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '客户
        Case 2
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_客户编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_客户名称            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_客户生日            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_客户性别            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_注册日期            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_客户行业            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_客户职业            VARCHAR(50)   NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '入库
        Case 3
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_入库产品编号        VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_入库产品数量        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_入库门店编号        VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_入库日期            DATE          NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '订单主表
        Case 4
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_订单编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_门店编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_下单日期            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_送货日期            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_客户编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_销售渠道            VARCHAR(50)   NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '订单子表
        Case 5
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_订单编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_产品编号            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_产品销售价格        FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_折扣比例            FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_产品销售数量        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_产品销售金额        FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '销售目标
        Case 6
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_省ID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_省简称              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_月份                DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_销售目标            FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '大区
        Case 7
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_大区ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_大区                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_大区负责人          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_办公地城市ID        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_办公地城市          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_纬度                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_经度                FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '省份
        Case 8
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_大区ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_省ID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_省全称              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_省简称              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_纬度                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_经度                FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '城市
        Case 9
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_自动编号            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_省ID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_城市ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_城市                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_纬度                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_经度                FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
    End Select

End Function


Public Function SqlDN(N As Long) As String

    SqlDN = "DROP TABLE " & TableNameN(N)

End Function


Public Function TableADO(TableName As String, Sql_Drop As String, Sql_Creat As String)
    
    Dim Cat As Object
    Dim Cmd As Object
    Dim Tbls As Object
    Dim Tbl As Object
    
    Set Cat = CreateObject("ADOX.Catalog")
    Set Cmd = CreateObject("ADODB.Command")
    
    Set Cat.ActiveConnection = CurrentProject.Connection
    Set Tbls = Cat.Tables
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        For Each Tbl In Tbls
            If Tbl.Name = TableName Then
                .CommandText = Sql_Drop
                .Execute
                
                .CommandText = Sql_Creat
                .Execute
                
                    Set Tbl = Nothing
                    Set Tbls = Nothing
                    Set Cat = Nothing
                    Set Cmd = Nothing
                Exit Function
            End If
        Next
        .CommandText = Sql_Creat
        .Execute
    End With
    
        Set Tbl = Nothing
        Set Tbls = Nothing
        Set Cat = Nothing
        Set Cmd = Nothing
    
End Function


Public Function DataTableD0()

    Dim Tnn As String
    Tnn = TableNameN(7)

    Dim i As Long
    Dim ArrAdd0
    Dim ArrAdd1
    Dim fnUB1 As Long

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs  As Object
    Dim Region0 As String

    Region0 = "1,东区,欧阳经往,73,上海,31.231618,121.471618;"
    Region0 = Region0 & "2,西区,焦阿灰,256,成都,30.651618,104.061618;"
    Region0 = Region0 & "3,南区,左丘垂漫,200,广州,23.121618,113.281618;"
    Region0 = Region0 & "4,北区,左烈佐,37,沈阳,41.791618,123.421618;"
    Region0 = Region0 & "5,中区,焦仔耘,1,北京,39.901618,116.401618;"
    Region0 = Region0 & "6,港澳台,安修谊,386,香港,22.321618,114.171618"

    ArrAdd0 = Split(Region0, ";")

    ReDim ArrAdd1(0 To UBound(ArrAdd0))

    For i = 0 To UBound(ArrAdd0)
        ArrAdd1(i) = Split(ArrAdd0(i), ",")
    Next

    fnUB1 = UBound(ArrAdd1)
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '省份表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With

    For i = 0 To fnUB1

        Rs.AddNew
        Rs(1) = ArrAdd1(i)(0)
        Rs(2) = ArrAdd1(i)(1)
        Rs(3) = ArrAdd1(i)(2)
        Rs(4) = ArrAdd1(i)(3)
        Rs(5) = ArrAdd1(i)(4)
        Rs(6) = ArrAdd1(i)(5)
        Rs(7) = ArrAdd1(i)(6)
        Rs.Update

    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
        
End Function


Public Function DataTableD1()
    Dim Tnn As String
    Tnn = TableNameN(8)

    Dim i As Long
    Dim ArrAdd0
    Dim ArrAdd1
    Dim fnUB1 As Long

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs  As Object

    ArrAdd0 = Split(AddressProvince, ";")

    ReDim ArrAdd1(0 To UBound(ArrAdd0))

    For i = 0 To UBound(ArrAdd0)
        ArrAdd1(i) = Split(ArrAdd0(i), ",")
    Next

    fnUB1 = UBound(ArrAdd1)
    
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '省份表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With

    For i = 0 To fnUB1

        Rs.AddNew
        Rs(1) = ArrAdd1(i)(0)
        Rs(2) = ArrAdd1(i)(1)
        Rs(3) = ArrAdd1(i)(2)
        Rs(4) = ArrAdd1(i)(3)
        Rs(5) = ArrAdd1(i)(4)
        Rs(6) = ArrAdd1(i)(5)
        Rs.Update

    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
        
End Function


Public Function DataTableD2()
    Dim Tnn As String
    Tnn = TableNameN(9)

    Dim i As Long
    Dim ArrAdd0
    Dim ArrAdd1
    Dim fnUB1 As Long

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs  As Object

    ArrAdd0 = Split(AddressCity, ";")

    ReDim ArrAdd1(0 To UBound(ArrAdd0))

    For i = 0 To UBound(ArrAdd0)
        ArrAdd1(i) = Split(ArrAdd0(i), ",")
    Next

    fnUB1 = UBound(ArrAdd1)
    
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '城市表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With

    For i = 0 To fnUB1

        Rs.AddNew
        Rs(1) = ArrAdd1(i)(0)
        Rs(2) = ArrAdd1(i)(1)
        Rs(3) = ArrAdd1(i)(2)
        Rs(4) = ArrAdd1(i)(3)
        Rs(5) = ArrAdd1(i)(4)
        Rs.Update

    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
        
End Function


Public Function DataTableT0()
    Dim Tnn As String
    Dim N00 As Long
        Randomize
    N00 = 1000 + Round(500 * Rnd(), 0)
    Tnn = TableNameN(0)
    Dim i As Long
    Dim Sj As Double
    Dim R4 As Double
    Dim R5 As Double

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs As Object
    
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        
        Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '产品表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    For i = 1 To N00
        Rs.AddNew
        Rs(1) = "SKU_" & Format(i, "000000")
        Randomize
        Sj = Rnd()
        
        Rs(2) = Chr(Round(Sj * 9, 0) + 65) & "类"
        Rs(3) = "产品" & Chr(Round(Sj * 9, 0) + 65) & "" & Format(i, "0000")
 
        R4 = 5000 + Sj * 5000

        If Sj < 0.28 Then
            R5 = R4 * 0.18
        Else
            R5 = R4 * Sj
        End If

        Rs(4) = Round(R4, 2)
        Rs(5) = Round(R5, 2)
        Rs.Update
    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
        
End Function


Public Function DataTableT1()
    Dim Tnn As String
    Tnn = TableNameN(1)
    Dim N10 As Long
    N10 = N1
    Dim i As Long
    Dim Sj As Double
    Dim ArrFN
    Dim ArrLN
    Dim ArrAdd0
    Dim ArrAdd1
    Dim ArrDict1
    Dim Dict1 As Object
    Dim fnUB0 As Long
    Dim fnUB1 As Long
    Dim addUB0 As Long
    Dim lnUB As Long
    Dim XingMing As String
    Dim dateKD As Date
    Dim dateGD As Date

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs  As Object

    ArrFN = Split(FirstName(), ",")
    ArrLN = Split(LastName(), ",")

    ArrAdd0 = Split(AddressCity, ";")

    ReDim ArrAdd1(0 To UBound(ArrAdd0))

    For i = 0 To UBound(ArrAdd0)
        ArrAdd1(i) = Split(ArrAdd0(i), ",")
    Next
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '门店表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    Set Dict1 = CreateObject("Scripting.Dictionary") '随机店名字，保证唯一不重复。
    For i = 1 To 17576 '26*26*26
        Dict1(Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & "店") = i
        If Dict1.Count = N10 Then
            Exit For
        End If
    Next
    ArrDict1 = Dict1.Keys
    Set Dict1 = Nothing

    '命中四个直辖市
    Rs.AddNew: Rs(1) = "SC_0001": Rs(2) = ArrDict1(0): Rs(3) = "焦阿大": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 1: Rs(6) = "北京": Rs(7) = 39.901618: Rs(8) = 116.401618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0002": Rs(2) = ArrDict1(1): Rs(3) = "焦阿二": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 2: Rs(6) = "天津": Rs(7) = 39.121618: Rs(8) = 117.191618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0003": Rs(2) = ArrDict1(2): Rs(3) = "焦阿三": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 73: Rs(6) = "上海": Rs(7) = 31.231618: Rs(8) = 121.471618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0004": Rs(2) = ArrDict1(3): Rs(3) = "焦阿四": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 255: Rs(6) = "重庆": Rs(7) = 29.531618: Rs(8) = 106.501618: Rs.Update

    For i = 5 To N10
        Randomize
        Sj = Rnd()
        fnUB0 = Round(UBound(ArrFN) * Sj, 0)
        fnUB1 = Round(UBound(ArrFN) * (1 - Sj), 0)
        lnUB = Round(UBound(ArrLN) * Sj, 0)
        Randomize
        addUB0 = Round(UBound(ArrAdd0) * Rnd(), 0)
        Randomize
        dateKD = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD") '+28容错Dict3N
        Randomize
        dateGD = Format(dateKD + 550 + 4320 * Rnd(), "YYYY-MM-DD") '550表示至少1.5年才能关店

        If Sj < 0.66 Then
            XingMing = ArrLN(lnUB) & ArrFN(fnUB0)
        Else
            XingMing = ArrLN(lnUB) & ArrFN(fnUB0) & ArrFN(fnUB1)
        End If

        If dateGD > Now Then

            Rs.AddNew
            Rs(1) = "SC_" & Format(i, "0000")
            Randomize
            Rs(2) = ArrDict1(i - 1)
            Rs(3) = XingMing
            Rs(4) = dateKD
            Rs(5) = ArrAdd1(addUB0)(1)
            Rs(6) = ArrAdd1(addUB0)(2)
            Rs(7) = Round(ArrAdd1(addUB0)(3) + Rnd() * 0.05, 6) '相同城市偏移，不会同一个点。
            Rs(8) = Round(ArrAdd1(addUB0)(4) + Rnd() * 0.05, 6)
            Rs.Update

        Else

            Rs.AddNew
            Rs(1) = "SC_" & Format(i, "0000")
            Randomize
            Rs(2) = ArrDict1(i - 1)
            Rs(3) = XingMing
            Rs(4) = dateKD
            Rs(5) = ArrAdd1(addUB0)(1)
            Rs(6) = ArrAdd1(addUB0)(2)
            Rs(7) = Round(ArrAdd1(addUB0)(3) + Rnd() * 0.05, 6)
            Rs(8) = Round(ArrAdd1(addUB0)(4) + Rnd() * 0.05, 6)
            Rs(9) = dateGD
            Rs.Update

        End If
    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
        
End Function


Public Function DataTableT2()
    Dim Tnn As String
    Tnn = TableNameN(2)
    Dim N2 As Long
    Dim i As Long
    Dim ii As Long
    Dim k As Long
    Dim Sj As Double
    Dim SjHY As Double
    Dim SjZY As Double
    Dim ArrFN
    Dim ArrLN
    Dim ArrHY '行业
    Dim ArrZY '职业
    Dim ArrSjNL '年龄分布
    
    Dim ArrSjHY
    Dim ArrSjZY
    
    Dim Arr1
    Dim Rrow As Long
    Dim Rcol As Long

    Dim fnUB0 As Long
    Dim fnUB1 As Long

    Dim lnUB As Long
    Dim XingMing As String
    Dim sex As String
    Dim dateSR As Date
    Dim dateZC As Date
    Dim dateZZKD As Date

    Dim Cmd  As Object
    Dim Conn  As Object
    Dim Rs  As Object
    Dim Rs1  As Object

    ArrHY = Array("建筑业", "制造业", "互联网", "农业", "餐饮", "物流", "汽车") '行业
    ArrZY = Array("个体户", "HR", "运营", "IT", "财务", "销售", "研发") '职业
    ArrSjHY = Array(0.2, 0.5, 0.5, 0.8, 0.8, 1, 0.9) '行业分布
    ArrSjZY = Array(0.3, 0.7, 0.6, 1, 1, 0.8, 0.1) '职业分布
    ArrSjNL = Array(0, 0.1, 0.2, 0.3, 0.3, 0.3, 0.3, 0.8, 0.8, 0.8, 0.9, 1) '年龄分布

    ArrFN = Split(FirstName(), ",")
    ArrLN = Split(LastName(), ",")
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
    Set Cmd = Nothing
        
    Set Conn = CreateObject("ADODB.Connection")
    Set Conn = CurrentProject.Connection
           
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '按照店铺规模注册客户
    Set Rs1 = CreateObject("ADODB.Recordset")
    With Rs1
            .ActiveConnection = Conn
            .Source = TableNameN(1) '门店表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
            .MoveFirst
        End With

        Rrow = Rs1.RecordCount - 1
        Rcol = Rs1.Fields.Count - 1

        ReDim Arr1(0 To Rrow, 0 To Rcol)

        For i = 0 To Rrow
            For k = 0 To Rcol
                Arr1(i, k) = Rs1(k)
            Next
            Rs1.MoveNext
        Next
        Rs1.Close
    Set Rs1 = Nothing
    '=====================================================================================
    
    Set Rs = CreateObject("ADODB.Recordset")
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '客户表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    ii = 0
    
    For i = 0 To UBound(Arr1)
        Randomize
        If IsNull(Arr1(i, 9)) Then
            N2 = Int((Now() - Arr1(i, 4)) * (1.2 + Rnd()))
        Else
            N2 = Int((Arr1(i, 4) - Arr1(i, 4)) * (1.2 + Rnd()))
        End If

        For k = 1 To N2
            ii = ii + 1
            Randomize
            Sj = Rnd()
            fnUB0 = Round(UBound(ArrFN) * Sj, 0)
            fnUB1 = Round(UBound(ArrFN) * (1 - Sj), 0)
            lnUB = Round(UBound(ArrLN) * Sj, 0)
            Randomize
            dateSR = Format(Now - 7500 - Round((ArrSjNL(ii Mod 12) + Rnd()) * 7000, 0), "YYYY-MM-DD")  '生日
            dateZC = Format(Now - 1500 + Round((ArrSjNL(ii Mod 12) + Rnd()) * 750, 0), "YYYY-MM-DD") '注册时间，比开店少两天，不会业务逻辑溢出
    
            If Sj < 0.8 Then
                XingMing = ArrLN(lnUB) & ArrFN(fnUB0)
                sex = "男"
            Else
                XingMing = ArrLN(lnUB) & ArrFN(fnUB0) & ArrFN(fnUB1)
                sex = "女"
            End If
    
            Rs.AddNew
            Rs(1) = "CC_" & Format(ii, "0000000")
            Rs(2) = XingMing
            Rs(3) = dateSR
            Rs(4) = sex
            Rs(5) = dateZC
            Randomize
            Rs(6) = ArrHY(Round(Rnd() * ArrSjHY(i Mod 7) * 6, 0))
            Randomize
            Rs(7) = ArrZY(Round(Rnd() * ArrSjZY(i Mod 7) * 6, 0))
            Rs.Update
        Next

    Next

    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
            
End Function


    Public Function DataTableT345()
    
        Dim i0 As Long
        Dim i1 As Long
        Dim i2 As Long
        Dim i3 As Long
        Dim i4 As Long

        Dim i As Long
        Dim k As Long
        Dim p As Long
        Dim q As Double
        Dim N As Long
        Dim ND As Long
        Dim Sj As Double

        Dim UB0 As Long
        Dim UB1 As Long
        Dim UB2 As Long

        Dim UB0n As Long
        Dim UB2n As Long

        Dim Oc As String
        Dim OcNumber As Long
        Dim Qd As String
        Dim Yyts As Long '营业天数
        Dim dateDD As Date

        Dim Arr0
        Dim Arr1
        Dim Arr2
        Dim ArrHY '行业
        Dim ArrZY '职业
        Dim ArrDjjsxs '单均件数系数
        Dim ArrDdslxsMonth '行业淡旺季趋势
        Dim ArrDdslxsSC '区域系数
        Dim ArrDict3
        Dim ArrDict5
        Dim ArrZK '折扣
        Dim ArrZKMonth '折扣月份分布
        Dim ArrSjKF '客户分布
        Dim Rrow As Long
        Dim Rcol As Long

        Dim Cmd As Object
        Dim Conn As Object
        Dim Rs0 As Object 'T00
        Dim Rs1 As Object 'T01
        Dim Rs2 As Object 'T02
        Dim Rs3 As Object 'T03
        Dim Rs4 As Object 'T04
        Dim Rs5 As Object 'T05
        Dim Dict2HY As Object
        Dim Dict2ZY As Object
        Dim Dict3N As Object
        Dim Dict3 As Object
        Dim Dict5 As Object
        
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

        With Cmd
            .CommandTimeout = 100
            .CommandText = "DELETE FROM " & TableNameN(3)
            .Execute
            .CommandText = "DELETE FROM " & TableNameN(4)
            .Execute
            .CommandText = "DELETE FROM " & TableNameN(5)
            .Execute
        End With
        Set Cmd = Nothing
        
        Set Conn = CreateObject("ADODB.Connection")
        Set Conn = CurrentProject.Connection
        
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs0 = CreateObject("ADODB.Recordset")
        With Rs0
            .ActiveConnection = Conn
            .Source = TableNameN(0) '产品表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
            .MoveFirst
        End With

        Rrow = Rs0.RecordCount - 1
        Rcol = Rs0.Fields.Count - 1

        ReDim Arr0(0 To Rrow, 0 To Rcol)

        For i = 0 To Rrow
            For k = 0 To Rcol
                Arr0(i, k) = Rs0(k)
            Next
            Rs0.MoveNext
        Next
        Rs0.Close
    Set Rs0 = Nothing
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs1 = CreateObject("ADODB.Recordset")
        With Rs1
            .ActiveConnection = Conn
            .Source = TableNameN(1) '门店表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
            .MoveFirst
        End With

        Rrow = Rs1.RecordCount - 1
        Rcol = Rs1.Fields.Count - 1

        ReDim Arr1(0 To Rrow, 0 To Rcol)

        For i = 0 To Rrow
            For k = 0 To Rcol
                Arr1(i, k) = Rs1(k)
            Next
            Rs1.MoveNext
        Next
        Rs1.Close
    Set Rs1 = Nothing
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ArrHY = Array("保险", "互联网", "汽车", "制造业") '折扣行业准备
    ArrZY = Array("HR", "财务", "销售", "运营")
    Set Dict2HY = CreateObject("Scripting.Dictionary") '行业
    Set Dict2ZY = CreateObject("Scripting.Dictionary") '职业
    
    For i = 0 To UBound(ArrHY)
        Dict2HY(ArrHY(i)) = ArrHY(i)
    Next

    For i = 0 To UBound(ArrZY)
        Dict2ZY(ArrZY(i)) = ArrZY(i)
    Next

    Set Rs2 = CreateObject("ADODB.Recordset")
        With Rs2
            .ActiveConnection = Conn
            .Source = TableNameN(2) '客户表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
            .MoveFirst
        End With

        Rrow = Rs2.RecordCount - 1
        Rcol = Rs2.Fields.Count - 1

        ReDim Arr2(0 To Rrow, 0 To 3)

        For i = 0 To Rrow
            Arr2(i, 0) = Rs2(1)
            Arr2(i, 1) = Rs2(5)
            Arr2(i, 2) = Rs2(6)
            Arr2(i, 3) = Rs2(7)
            Rs2.MoveNext
        Next
        Rs2.Close
    Set Rs2 = Nothing
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs3 = CreateObject("ADODB.Recordset")
        With Rs3
            .ActiveConnection = Conn
            .Source = TableNameN(3) '入库表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs4 = CreateObject("ADODB.Recordset")
    With Rs4
            .ActiveConnection = Conn
            .Source = TableNameN(4) '订单主表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs5 = CreateObject("ADODB.Recordset")
        With Rs5
            .ActiveConnection = Conn
            .Source = TableNameN(5) '订单子表
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
        '=====================================================================================
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        UB0 = UBound(Arr0) '产品数量
        UB1 = UBound(Arr1) '门店数量
        UB2 = UBound(Arr2) '客户数量
        OcNumber = 0
        
        
        ArrDjjsxs = Array(0.7, 0.8, 1, 1.2, 1.3) '单均件数系数，count=5
        ArrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '行业淡旺季趋势，count=12
        ArrDdslxsSC = Array(0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6) '区域订单系数正太分布，count=34
        ArrZK = Array(1, 0.9, 0.8, 0.7, 0.6, 0.5) '折扣信息分布，count=6
        ArrZKMonth = Array(0.95, 0.9, 1, 0.98, 0.85, 1, 0.98, 0.88, 0.8, 0.86, 0.92, 0.98) '行业淡旺季趋势，count=12
        ArrSjKF = Array(0, 0.1, 0.5, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.8, 0.9, 1) '客户分布

        
        
    For i1 = 0 To UB1

            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            '营业天数
            If IsNull(Arr1(i1, 9)) Then
                Yyts = Round(Now - Arr1(i1, 4), 0)
            Else
                Yyts = Round(Arr1(i1, 9) - Arr1(i1, 4), 0)
            End If
            '=====================================================================================
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            Set Dict3N = CreateObject("Scripting.Dictionary") '入库的随机序列
            Dict3N(1) = Yyts Mod 6 + 1
                    '入库随机时间
                    For i = 1 To Yyts
                        Dict3N(i + 1) = Dict3N.Item(i) + Round(Rnd() * 2 + 5, 0)
                        If Dict3N.Item(i + 1) > Yyts Then
                            Dict3N(i + 1) = Yyts
                            Exit For
                        End If
                    Next
            '=====================================================================================

            Set Dict3 = CreateObject("Scripting.Dictionary") '记录入库信息
            i3 = 1

            For i4 = 1 To Yyts

                dateDD = Arr1(i1, 4) + i4 - 1
                Randomize
                
                ND = Round(Rnd() * 4 * ArrDdslxsMonth(Month(dateDD) - 1) * ArrDdslxsSC(Arr1(i1, 5) Mod (UBound(ArrDdslxsSC) + 1)), 0) '每天订单

                If ND = 0 Then GoTo Dd0 '没有销售

                For i = 1 To ND '每天订单数
                    OcNumber = OcNumber + 1
                    Oc = "OC_" & Format(OcNumber, "0000000")
                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '订单主表写入
                    Randomize
                    Sj = (Rnd() + ArrSjKF(i4 * i Mod 12)) / 2

                    UB2n = Round(UB2 * Sj, 0)
                    
                '注册与购买分布
                If IsNull(Arr1(i1, 9)) Then '未关店
                    If Arr2(UB2n, 1) >= Arr1(i1, 4) And OcNumber Mod 13 > 6 Then
                        GoTo UB2nlable
                    Else
                        For i2 = UB2n To UB2 '往右前进
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 13 > 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 > 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                                
                        For i2 = UB2n To 0 Step -1 '往左前进
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 13 <= 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 <= 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                    End If
                Else '关店
                    If Arr2(UB2n, 1) >= Arr1(i1, 4) And Arr2(UB2n, 1) < Arr1(i1, 9) And OcNumber Mod 13 < 6 Then
                            GoTo UB2nlable
                    Else
                        For i2 = UB2n To UB2 '往右前进
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Arr2(i2, 1) < Arr1(i1, 9) And Month(Arr2(i2, 1)) Mod 13 > 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 > 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                            
                        For i2 = UB2n To 0 Step -1 '往左前进
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Arr2(i2, 1) < Arr1(i1, 9) And Month(Arr2(i2, 1)) Mod 13 <= 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Arr2(i2, 1) < Arr1(i1, 9) And Month(Arr2(i2, 1)) Mod 3 <= 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                    End If
                End If
     
UB2nlable:

                    If Sj < 0.7 Then
                        Qd = "线上"
                    Else
                        Qd = "线下"
                    End If
                    Rs4.AddNew
                    Rs4(1) = Oc
                    Rs4(2) = Arr1(i1, 1)
                    Rs4(3) = Arr1(i1, 4) + i4 - 1
                    Rs4(4) = Arr1(i1, 4) + i4 + Round(4 * Sj + 8, 0)
                    Rs4(5) = Arr2(UB2n, 0)
                    Rs4(6) = Qd
                    Rs4.Update
                    '=====================================================================================

                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '订单子表

                    Randomize

                    k = Round(5 * Rnd(), 0) + 1 '计划每个订产品数量上限,均值为3

                    Set Dict5 = CreateObject("Scripting.Dictionary")
                    For N = 1 To k
                        If k < 4 Then
                            UB0n = Round(UB0 * Rnd() / 5, 0)  '往左偏移
                        Else
                            UB0n = Round(UB0 * Rnd(), 0)
                        End If
                        Dict5(UB0n) = UB0n '字典去重sku
                    Next

                    ArrDict5 = Dict5.Keys

                    For N = 0 To Dict5.Count - 1

                        p = Round(5 * Rnd() * ArrDjjsxs(i1 Mod 5) * ArrDjjsxs(ArrDict5(N) Mod 5), 0) + 1 '件数系数加权

                        'q产品折扣
                        
                        If i1 Mod 40 > 30 Then '区域
                            q = ArrZK(0) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf i1 Mod 40 < 10 Then
                            q = ArrZK(5) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf ArrDict5(N) Mod 8 < 1 Then '产品
                            q = ArrZK(1) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf ArrDict5(N) Mod 8 > 5 Then
                            q = ArrZK(3) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf Dict2HY.exists(Arr2(UB2n, 2)) Then  '客户
                            q = ArrZK(2) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf Dict2ZY.exists(Arr2(UB2n, 3)) Then
                            q = ArrZK(4) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        Else
                            q = 1
                        End If

ExitIFzk:
                        Rs5.AddNew
                        Rs5(1) = Oc
                        Rs5(2) = Arr0(ArrDict5(N), 1)
                        Rs5(3) = Arr0(ArrDict5(N), 4)
                        Rs5(4) = Round(q, 2)
                        Rs5(5) = p
                        Rs5(6) = Round(Arr0(ArrDict5(N), 4) * p * q, 2)
                        Rs5.Update
                        Dict3(Arr0(ArrDict5(N), 1)) = Dict3(Arr0(ArrDict5(N), 1)) + p
                    Next
                Set Dict5 = Nothing
                '=====================================================================================
            Next

Dd0: '当日无订单跳转
                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                '生成入库信息
                If Dict3N.Item(i3) = i4 And i4 < Yyts Then
                    i3 = i3 + 1

                    ArrDict3 = Dict3.Keys

                    For i0 = 0 To UBound(ArrDict3)
                        Rs3.AddNew
                        Rs3(1) = ArrDict3(i0)
                        Rs3(2) = Dict3.Item(ArrDict3(i0))
                        Rs3(3) = Arr1(i1, 1)
                        Rs3(4) = Arr1(i1, 4) + i4 - 14 '-14保证有库存
                        Rs3.Update
                    Next
                    Set Dict3 = Nothing
                    Set Dict3 = CreateObject("Scripting.Dictionary") '记录入库信息
                    GoTo Rk0
                    
                ElseIf i4 = Yyts Then  '保证最后一次入库累计大于0

                    i3 = i3 + 1

                    ArrDict3 = Dict3.Keys

                    For i0 = 0 To UBound(ArrDict3)
                        Rs3.AddNew
                        Rs3(1) = ArrDict3(i0)
                        Rs3(2) = Dict3.Item(ArrDict3(i0)) + Round(Rnd() * 5, 0)
                        Rs3(3) = Arr1(i1, 1)
                        Rs3(4) = Arr1(i1, 4) + i4 - 14 '-14保证有库存
                        Rs3.Update
                    Next
                    Set Dict3 = Nothing
                    Set Dict3 = CreateObject("Scripting.Dictionary") '记录入库信息
                    GoTo Rk0
                    
                End If
                '=====================================================================================
Rk0:
            Next

        Next
        Rs3.Close
        Rs4.Close
        Rs5.Close
        Conn.Close
    Set Rs3 = Nothing
    Set Rs4 = Nothing
    Set Rs5 = Nothing
    Set Conn = Nothing

End Function


Public Function DataTableT6()

    Dim Tnn As String
    Tnn = TableNameN(6)

    Dim Sqlstr As String
    Dim ArrYQ
    Dim ArrDdslxsMonth '行业淡旺季趋势
    Dim Qn As Double '最后三个月的系数和
    Dim B As Double
    Dim UP0 As Double '拉开区域差距
    
    Dim i As Long
    Dim Rrow As Long
    Dim Rcol As Long
    Dim k As Long
    
    Dim Conn  As Object
    Dim Rs  As Object
    Dim Cmd  As Object
    Set Cmd = CreateObject("ADODB.Command")
    Set Cmd.ActiveConnection = CurrentProject.Connection

    With Cmd
        .CommandTimeout = 100
        .CommandText = "DELETE FROM " & Tnn
        .Execute
    End With
    Set Cmd = Nothing
    
    Set Conn = CreateObject("ADODB.Connection")
    Set Conn = CurrentProject.Connection
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '去年完成情况全年&去年Q4完成情况,月均取大
    Sqlstr = "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "TY.*,TQ.A3Q" & Chr(13)
    Sqlstr = Sqlstr & "FROM" & Chr(13)
    Sqlstr = Sqlstr & "(" & Chr(13)
    Sqlstr = Sqlstr & "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "D01_省份表.F_02_省ID AS A0省ID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_省份表.F_04_省简称 AS A1省简称" & Chr(13)
    Sqlstr = Sqlstr & ", Sum(T05_订单子表.F_06_产品销售金额) AS A2Y" & Chr(13)
    Sqlstr = Sqlstr & "FROM (((T05_订单子表 " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T04_订单主表 ON T05_订单子表.F_01_订单编号 = T04_订单主表.F_01_订单编号) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T01_门店表 ON T04_订单主表.F_02_门店编号 = T01_门店表.F_01_门店编号) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D02_城市表 ON T01_门店表.F_05_城市ID = D02_城市表.F_02_城市ID) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D01_省份表 ON D02_城市表.F_01_省ID = D01_省份表.F_02_省ID" & Chr(13)
    Sqlstr = Sqlstr & "WHERE T04_订单主表.[F_03_下单日期]>#" & Format(Now(), "YYYY") - 2 & "-12-1# AND T04_订单主表.[F_03_下单日期]<#" & Format(Now(), "YYYY") & "-1-1#" & Chr(13)
    Sqlstr = Sqlstr & "GROUP BY " & Chr(13)
    Sqlstr = Sqlstr & "D01_省份表.F_02_省ID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_省份表.F_04_省简称" & Chr(13)
    Sqlstr = Sqlstr & ") TY" & Chr(13)
    Sqlstr = Sqlstr & "LEFT JOIN" & Chr(13)
    Sqlstr = Sqlstr & "(" & Chr(13)
    Sqlstr = Sqlstr & "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "D01_省份表.F_02_省ID AS A0省ID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_省份表.F_04_省简称 AS A1省简称" & Chr(13)
    Sqlstr = Sqlstr & ", Sum(T05_订单子表.F_06_产品销售金额) AS A3Q" & Chr(13)
    Sqlstr = Sqlstr & "FROM (((T05_订单子表 " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T04_订单主表 ON T05_订单子表.F_01_订单编号 = T04_订单主表.F_01_订单编号) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T01_门店表 ON T04_订单主表.F_02_门店编号 = T01_门店表.F_01_门店编号) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D02_城市表 ON T01_门店表.F_05_城市ID = D02_城市表.F_02_城市ID) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D01_省份表 ON D02_城市表.F_01_省ID = D01_省份表.F_02_省ID" & Chr(13)
    Sqlstr = Sqlstr & "WHERE T04_订单主表.[F_03_下单日期]>#" & Format(Now(), "YYYY") - 1 & "-9-1# AND T04_订单主表.[F_03_下单日期]<#" & Format(Now(), "YYYY") & "-1-1#" & Chr(13)
    Sqlstr = Sqlstr & "GROUP BY " & Chr(13)
    Sqlstr = Sqlstr & "D01_省份表.F_02_省ID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_省份表.F_04_省简称" & Chr(13)
    Sqlstr = Sqlstr & ") TQ" & Chr(13)
    Sqlstr = Sqlstr & "ON TY.A0省ID=TQ.A0省ID AND TY.A1省简称=TQ.A1省简称" & Chr(13)

    Set Rs = CreateObject("ADODB.Recordset")
    With Rs
        .ActiveConnection = Conn
        .Source = Sqlstr
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With

    Rrow = Rs.RecordCount - 1
    Rcol = Rs.Fields.Count - 1

    ReDim ArrYQ(0 To Rrow, 0 To Rcol)

    For i = 0 To Rrow
        For k = 0 To Rcol
            ArrYQ(i, k) = Rs(k)
        Next
        Rs.MoveNext
    Next
    Rs.Close
    Set Rs = Nothing

    '=====================================================================================
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    ArrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '行业淡旺季趋势,归一，count=12;同上DataTableT345

    Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '销售目标表
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    
    For i = UBound(ArrDdslxsMonth) - 2 To UBound(ArrDdslxsMonth)
        Qn = ArrDdslxsMonth(i) + Qn
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
            Rs.AddNew
            Rs(1) = ArrYQ(i, 0)
            Rs(2) = ArrYQ(i, 1)
            Rs(3) = Format(Now(), "YYYY") - 1 & "-" & k & "-1"
            Randomize
            Rs(4) = Round(B * (0.7 + Rnd() * 0.1 * UP0) * ArrDdslxsMonth(k - 1), 0) '下方今年目标的比例不一样，今年都按照实际浮动不大。
            Rs.Update
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
            Rs.AddNew
            Rs(1) = ArrYQ(i, 0)
            Rs(2) = ArrYQ(i, 1)
            Rs(3) = Format(Now(), "YYYY") & "-" & k & "-1"
            Randomize
            Rs(4) = Round(B * (0.6 + Rnd() * 0.2 * UP0) * ArrDdslxsMonth(k - 1), 0)
            Rs.Update
        Next
    Next
    '=====================================================================================
    Rs.Close
    Conn.Close
    Set Rs = Nothing
    Set Conn = Nothing
    '=====================================================================================
    
End Function


Public Function FirstName() As String

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

    '大区ID 省ID    省全称  省简称  纬度    经度
    AddressProvince = "5,1,北京市,北京,39.904987,116.405289;"
    AddressProvince = AddressProvince & "5,2,天津市,天津,39.125595,117.190186;"
    AddressProvince = AddressProvince & "5,3,河北省,河北,38.045475,114.502464;"
    AddressProvince = AddressProvince & "5,4,山西省,山西,37.857014,112.549248;"
    AddressProvince = AddressProvince & "4,5,内蒙古自治区,内蒙古,40.81831,111.670799;"
    AddressProvince = AddressProvince & "4,6,辽宁省,辽宁,41.796768,123.429092;"
    AddressProvince = AddressProvince & "4,7,吉林省,吉林,43.886841,125.324501;"
    AddressProvince = AddressProvince & "4,8,黑龙江省,黑龙江,45.756966,126.642464;"
    AddressProvince = AddressProvince & "1,9,上海市,上海,31.231707,121.472641;"
    AddressProvince = AddressProvince & "1,10,江苏省,江苏,32.041546,118.76741;"
    AddressProvince = AddressProvince & "1,11,浙江省,浙江,30.287458,120.15358;"
    AddressProvince = AddressProvince & "1,12,安徽省,安徽,31.861191,117.283043;"
    AddressProvince = AddressProvince & "1,13,福建省,福建,26.075302,119.306236;"
    AddressProvince = AddressProvince & "1,14,江西省,江西,28.676493,115.892151;"
    AddressProvince = AddressProvince & "4,15,山东省,山东,36.675808,117.000923;"
    AddressProvince = AddressProvince & "5,16,河南省,河南,34.757977,113.665413;"
    AddressProvince = AddressProvince & "5,17,湖北省,湖北,30.584354,114.298569;"
    AddressProvince = AddressProvince & "3,18,湖南省,湖南,28.19409,112.982277;"
    AddressProvince = AddressProvince & "3,19,广东省,广东,23.125177,113.28064;"
    AddressProvince = AddressProvince & "3,20,广西壮族自治区,广西,22.82402,108.320007;"
    AddressProvince = AddressProvince & "3,21,海南省,海南,20.031971,110.331192;"
    AddressProvince = AddressProvince & "2,22,重庆市,重庆,29.533155,106.504959;"
    AddressProvince = AddressProvince & "2,23,四川省,四川,30.659462,104.065735;"
    AddressProvince = AddressProvince & "3,24,贵州省,贵州,26.578342,106.713478;"
    AddressProvince = AddressProvince & "3,25,云南省,云南,25.040609,102.71225;"
    AddressProvince = AddressProvince & "2,26,西藏自治区,西藏,29.66036,91.13221;"
    AddressProvince = AddressProvince & "5,27,陕西省,陕西,34.263161,108.948021;"
    AddressProvince = AddressProvince & "2,28,甘肃省,甘肃,36.058041,103.823555;"
    AddressProvince = AddressProvince & "2,29,青海省,青海,36.623177,101.778915;"
    AddressProvince = AddressProvince & "2,30,宁夏回族自治区,宁夏,38.46637,106.278175;"
    AddressProvince = AddressProvince & "2,31,新疆维吾尔自治区,新疆,43.792816,87.617729;"
    AddressProvince = AddressProvince & "6,32,台湾省,台湾,25.041618,121.501618;"
    AddressProvince = AddressProvince & "6,33,香港特别行政区,香港,22.320047,114.173355;"
    AddressProvince = AddressProvince & "6,34,澳门特别行政区,澳门,22.198952,113.549088"

End Function


Public Function AddressCity() As String

    '省ID    城市ID  城市    纬度    经度
    AddressCity = "1,1,北京,39.904987,116.405289;"
    AddressCity = AddressCity & "2,2,天津,39.125595,117.190186;"
    AddressCity = AddressCity & "3,3,石家庄,38.045475,114.502464;"
    AddressCity = AddressCity & "3,4,唐山,39.635113,118.175392;"
    AddressCity = AddressCity & "3,5,秦皇岛,39.942532,119.586578;"
    AddressCity = AddressCity & "3,6,邯郸,36.612274,114.490685;"
    AddressCity = AddressCity & "3,7,邢台,37.068199,114.50885;"
    AddressCity = AddressCity & "3,8,保定,38.867657,115.48233;"
    AddressCity = AddressCity & "3,9,张家口,40.811901,114.884094;"
    AddressCity = AddressCity & "3,10,承德,40.976204,117.939156;"
    AddressCity = AddressCity & "3,11,沧州,38.310581,116.85746;"
    AddressCity = AddressCity & "3,12,廊坊,39.523926,116.704437;"
    AddressCity = AddressCity & "3,13,衡水,37.735096,115.665993;"
    AddressCity = AddressCity & "4,14,太原,37.857014,112.549248;"
    AddressCity = AddressCity & "4,15,大同,40.090309,113.295258;"
    AddressCity = AddressCity & "4,16,阳泉,37.861187,113.583282;"
    AddressCity = AddressCity & "4,17,长治,36.191113,113.113556;"
    AddressCity = AddressCity & "4,18,晋城,35.497555,112.851273;"
    AddressCity = AddressCity & "4,19,朔州,39.331261,112.433388;"
    AddressCity = AddressCity & "4,20,晋中,37.696495,112.736465;"
    AddressCity = AddressCity & "4,21,运城,35.022778,111.00396;"
    AddressCity = AddressCity & "4,22,忻州,38.41769,112.733536;"
    AddressCity = AddressCity & "4,23,临汾,36.084148,111.517975;"
    AddressCity = AddressCity & "4,24,吕梁,37.524364,111.134338;"
    AddressCity = AddressCity & "5,25,呼和浩特,40.81831,111.670799;"
    AddressCity = AddressCity & "5,26,包头,40.658169,109.840408;"
    AddressCity = AddressCity & "5,27,乌海,39.673733,106.825562;"
    AddressCity = AddressCity & "5,28,赤峰,42.275318,118.956802;"
    AddressCity = AddressCity & "5,29,通辽,43.617428,122.263123;"
    AddressCity = AddressCity & "5,30,鄂尔多斯,39.817181,109.990288;"
    AddressCity = AddressCity & "5,31,呼伦贝尔,49.215332,119.758171;"
    AddressCity = AddressCity & "5,32,巴彦淖尔,40.757401,107.416962;"
    AddressCity = AddressCity & "5,33,乌兰察布,41.034126,113.11454;"
    AddressCity = AddressCity & "5,34,兴安,46.076267,122.07032;"
    AddressCity = AddressCity & "5,35,锡林郭勒,43.944019,116.090996;"
    AddressCity = AddressCity & "5,36,阿拉善,38.844814,105.706421;"
    AddressCity = AddressCity & "6,37,沈阳,41.796768,123.429092;"
    AddressCity = AddressCity & "6,38,大连,38.914589,121.618622;"
    AddressCity = AddressCity & "6,39,鞍山,41.110626,122.995628;"
    AddressCity = AddressCity & "6,40,抚顺,41.875957,123.921112;"
    AddressCity = AddressCity & "6,41,本溪,41.297909,123.770515;"
    AddressCity = AddressCity & "6,42,丹东,40.124294,124.383041;"
    AddressCity = AddressCity & "6,43,锦州,41.11927,121.135742;"
    AddressCity = AddressCity & "6,44,营口,40.667431,122.235153;"
    AddressCity = AddressCity & "6,45,阜新,42.011795,121.648964;"
    AddressCity = AddressCity & "6,46,辽阳,41.269402,123.181519;"
    AddressCity = AddressCity & "6,47,盘锦,41.124485,122.069572;"
    AddressCity = AddressCity & "6,48,铁岭,42.290585,123.844276;"
    AddressCity = AddressCity & "6,49,朝阳,41.576759,120.45118;"
    AddressCity = AddressCity & "6,50,葫芦岛,40.755573,120.856392;"
    AddressCity = AddressCity & "7,51,长春,43.886841,125.324501;"
    AddressCity = AddressCity & "7,52,吉林,43.843578,126.553017;"
    AddressCity = AddressCity & "7,53,四平,43.170345,124.370789;"
    AddressCity = AddressCity & "7,54,辽源,42.902691,125.145348;"
    AddressCity = AddressCity & "7,55,通化,41.721176,125.936501;"
    AddressCity = AddressCity & "7,56,白山,41.942505,126.427841;"
    AddressCity = AddressCity & "7,57,松原,45.118244,124.823608;"
    AddressCity = AddressCity & "7,58,白城,45.619026,122.84111;"
    AddressCity = AddressCity & "7,59,延边朝鲜族,42.904823,129.513229;"
    AddressCity = AddressCity & "8,60,哈尔滨,45.756966,126.642464;"
    AddressCity = AddressCity & "8,61,齐齐哈尔,47.342079,123.957916;"
    AddressCity = AddressCity & "8,62,鸡西,45.300045,130.975967;"
    AddressCity = AddressCity & "8,63,鹤岗,47.332085,130.277481;"
    AddressCity = AddressCity & "8,64,双鸭山,46.64344,131.157303;"
    AddressCity = AddressCity & "8,65,大庆,46.590733,125.112717;"
    AddressCity = AddressCity & "8,66,伊春,47.724773,128.899399;"
    AddressCity = AddressCity & "8,67,佳木斯,46.809605,130.361633;"
    AddressCity = AddressCity & "8,68,七台河,45.771267,131.015579;"
    AddressCity = AddressCity & "8,69,牡丹江,44.582962,129.618607;"
    AddressCity = AddressCity & "8,70,黑河,50.249584,127.499023;"
    AddressCity = AddressCity & "8,71,绥化,46.637394,126.992928;"
    AddressCity = AddressCity & "8,72,大兴安岭,52.335262,124.711525;"
    AddressCity = AddressCity & "9,73,上海,31.231707,121.472641;"
    AddressCity = AddressCity & "10,74,南京,32.041546,118.76741;"
    AddressCity = AddressCity & "10,75,无锡,31.57473,120.301666;"
    AddressCity = AddressCity & "10,76,徐州,34.261791,117.184814;"
    AddressCity = AddressCity & "10,77,常州,31.772753,119.946976;"
    AddressCity = AddressCity & "10,78,苏州,31.299379,120.619583;"
    AddressCity = AddressCity & "10,79,南通,32.016212,120.864609;"
    AddressCity = AddressCity & "10,80,连云港,34.600018,119.178818;"
    AddressCity = AddressCity & "10,81,淮安,33.597507,119.021263;"
    AddressCity = AddressCity & "10,82,盐城,33.377632,120.139999;"
    AddressCity = AddressCity & "10,83,扬州,32.393158,119.421005;"
    AddressCity = AddressCity & "10,84,镇江,32.204403,119.452751;"
    AddressCity = AddressCity & "10,85,泰州,32.484882,119.915176;"
    AddressCity = AddressCity & "10,86,宿迁,33.963009,118.275162;"
    AddressCity = AddressCity & "11,87,杭州,30.287458,120.15358;"
    AddressCity = AddressCity & "11,88,宁波,29.868387,121.549789;"
    AddressCity = AddressCity & "11,89,温州,28.000574,120.672112;"
    AddressCity = AddressCity & "11,90,嘉兴,30.762653,120.750862;"
    AddressCity = AddressCity & "11,91,湖州,30.867199,120.102402;"
    AddressCity = AddressCity & "11,92,绍兴,29.997116,120.582115;"
    AddressCity = AddressCity & "11,93,金华,29.089523,119.649506;"
    AddressCity = AddressCity & "11,94,衢州,28.941708,118.872627;"
    AddressCity = AddressCity & "11,95,舟山,30.016027,122.106865;"
    AddressCity = AddressCity & "11,96,台州,28.661379,121.428596;"
    AddressCity = AddressCity & "11,97,丽水,28.451994,119.921783;"
    AddressCity = AddressCity & "12,98,合肥,31.861191,117.283043;"
    AddressCity = AddressCity & "12,99,芜湖,31.326319,118.37645;"
    AddressCity = AddressCity & "12,100,蚌埠,32.939667,117.363228;"
    AddressCity = AddressCity & "12,101,淮南,32.647575,117.018326;"
    AddressCity = AddressCity & "12,102,马鞍山,31.689362,118.507904;"
    AddressCity = AddressCity & "12,103,淮北,33.971706,116.794662;"
    AddressCity = AddressCity & "12,104,铜陵,30.929935,117.816574;"
    AddressCity = AddressCity & "12,105,安庆,30.508829,117.043549;"
    AddressCity = AddressCity & "12,106,黄山,29.709238,118.317322;"
    AddressCity = AddressCity & "12,107,滁州,32.303627,118.316261;"
    AddressCity = AddressCity & "12,108,阜阳,32.896969,115.819733;"
    AddressCity = AddressCity & "12,109,宿州,33.633892,116.984085;"
    AddressCity = AddressCity & "12,110,六安,31.75289,116.507675;"
    AddressCity = AddressCity & "12,111,亳州,33.869339,115.782936;"
    AddressCity = AddressCity & "12,112,池州,30.656036,117.489159;"
    AddressCity = AddressCity & "12,113,宣城,30.945667,118.757996;"
    AddressCity = AddressCity & "13,114,福州,26.075302,119.306236;"
    AddressCity = AddressCity & "13,115,厦门,24.490475,118.110222;"
    AddressCity = AddressCity & "13,116,莆田,25.431011,119.007561;"
    AddressCity = AddressCity & "13,117,三明,26.265444,117.635002;"
    AddressCity = AddressCity & "13,118,泉州,24.908854,118.589424;"
    AddressCity = AddressCity & "13,119,漳州,24.510897,117.661804;"
    AddressCity = AddressCity & "13,120,南平,26.635628,118.178459;"
    AddressCity = AddressCity & "13,121,龙岩,25.091602,117.029778;"
    AddressCity = AddressCity & "13,122,宁德,26.659241,119.527084;"
    AddressCity = AddressCity & "14,123,南昌,28.676493,115.892151;"
    AddressCity = AddressCity & "14,124,景德镇,29.292561,117.214661;"
    AddressCity = AddressCity & "14,125,萍乡,27.622946,113.852188;"
    AddressCity = AddressCity & "14,126,九江,29.712034,115.992813;"
    AddressCity = AddressCity & "14,127,新余,27.810835,114.930832;"
    AddressCity = AddressCity & "14,128,鹰潭,28.238638,117.033836;"
    AddressCity = AddressCity & "14,129,赣州,25.850969,114.940277;"
    AddressCity = AddressCity & "14,130,吉安,27.111698,114.986374;"
    AddressCity = AddressCity & "14,131,宜春,27.8043,114.391136;"
    AddressCity = AddressCity & "14,132,抚州,27.98385,116.358353;"
    AddressCity = AddressCity & "14,133,上饶,28.44442,117.971184;"
    AddressCity = AddressCity & "15,134,济南,36.675808,117.000923;"
    AddressCity = AddressCity & "15,135,青岛,36.082981,120.355171;"
    AddressCity = AddressCity & "15,136,淄博,36.814938,118.047646;"
    AddressCity = AddressCity & "15,137,枣庄,34.856422,117.557961;"
    AddressCity = AddressCity & "15,138,东营,37.434563,118.664711;"
    AddressCity = AddressCity & "15,139,烟台,37.539295,121.39138;"
    AddressCity = AddressCity & "15,140,潍坊,36.709251,119.107079;"
    AddressCity = AddressCity & "15,141,济宁,35.415394,116.587242;"
    AddressCity = AddressCity & "15,142,泰安,36.194969,117.129066;"
    AddressCity = AddressCity & "15,143,威海,37.509689,122.116394;"
    AddressCity = AddressCity & "15,144,日照,35.428589,119.461205;"
    AddressCity = AddressCity & "15,145,莱芜,36.214397,117.677734;"
    AddressCity = AddressCity & "15,146,临沂,35.065281,118.326447;"
    AddressCity = AddressCity & "15,147,德州,37.453968,116.307426;"
    AddressCity = AddressCity & "15,148,聊城,36.456013,115.98037;"
    AddressCity = AddressCity & "15,149,滨州,37.383541,118.016975;"
    AddressCity = AddressCity & "15,150,菏泽,35.246532,115.469383;"
    AddressCity = AddressCity & "16,151,郑州,34.757977,113.665413;"
    AddressCity = AddressCity & "16,152,开封,34.79705,114.341446;"
    AddressCity = AddressCity & "16,153,洛阳,34.66304,112.434471;"
    AddressCity = AddressCity & "16,154,平顶山,33.735241,113.307716;"
    AddressCity = AddressCity & "16,155,安阳,36.103443,114.352486;"
    AddressCity = AddressCity & "16,156,鹤壁,35.748238,114.295441;"
    AddressCity = AddressCity & "16,157,新乡,35.302616,113.883987;"
    AddressCity = AddressCity & "16,158,焦作,35.23904,113.238266;"
    AddressCity = AddressCity & "16,159,济源,35.090378,112.59005;"
    AddressCity = AddressCity & "16,160,濮阳,35.768234,115.041298;"
    AddressCity = AddressCity & "16,161,许昌,34.022957,113.826065;"
    AddressCity = AddressCity & "16,162,漯河,33.575855,114.026405;"
    AddressCity = AddressCity & "16,163,三门峡,34.777336,111.194099;"
    AddressCity = AddressCity & "16,164,南阳,32.999081,112.540916;"
    AddressCity = AddressCity & "16,165,商丘,34.437054,115.650497;"
    AddressCity = AddressCity & "16,166,信阳,32.123276,114.075027;"
    AddressCity = AddressCity & "16,167,周口,33.620358,114.649651;"
    AddressCity = AddressCity & "16,168,驻马店,32.980167,114.024734;"
    AddressCity = AddressCity & "17,169,武汉,30.584354,114.298569;"
    AddressCity = AddressCity & "17,170,黄石,30.220074,115.077049;"
    AddressCity = AddressCity & "17,171,十堰,32.646908,110.787918;"
    AddressCity = AddressCity & "17,172,宜昌,30.702637,111.29084;"
    AddressCity = AddressCity & "17,173,襄阳,32.042427,112.14415;"
    AddressCity = AddressCity & "17,174,鄂州,30.396536,114.890594;"
    AddressCity = AddressCity & "17,175,荆门,31.035419,112.204254;"
    AddressCity = AddressCity & "17,176,孝感,30.926422,113.926659;"
    AddressCity = AddressCity & "17,177,荆州,30.326857,112.238129;"
    AddressCity = AddressCity & "17,178,黄冈,30.447712,114.879364;"
    AddressCity = AddressCity & "17,179,咸宁,29.832798,114.328964;"
    AddressCity = AddressCity & "17,180,随州,31.717497,113.373772;"
    AddressCity = AddressCity & "17,181,恩施,30.283113,109.486992;"
    AddressCity = AddressCity & "17,182,仙桃,30.364952,113.453972;"
    AddressCity = AddressCity & "17,183,潜江,30.421215,112.896866;"
    AddressCity = AddressCity & "17,184,天门,30.653061,113.165863;"
    AddressCity = AddressCity & "17,185,神农架,30.584354,114.298569;"
    AddressCity = AddressCity & "18,186,长沙,28.19409,112.982277;"
    AddressCity = AddressCity & "18,187,株洲,27.835806,113.151733;"
    AddressCity = AddressCity & "18,188,湘潭,27.829729,112.944054;"
    AddressCity = AddressCity & "18,189,衡阳,26.900358,112.607697;"
    AddressCity = AddressCity & "18,190,邵阳,27.237843,111.469231;"
    AddressCity = AddressCity & "18,191,岳阳,29.370291,113.132858;"
    AddressCity = AddressCity & "18,192,常德,29.040224,111.691345;"
    AddressCity = AddressCity & "18,193,张家界,29.127401,110.479919;"
    AddressCity = AddressCity & "18,194,益阳,28.570066,112.355042;"
    AddressCity = AddressCity & "18,195,郴州,25.793589,113.032066;"
    AddressCity = AddressCity & "18,196,永州,26.434517,111.608017;"
    AddressCity = AddressCity & "18,197,怀化,27.550081,109.978241;"
    AddressCity = AddressCity & "18,198,娄底,27.728136,112.008499;"
    AddressCity = AddressCity & "18,199,湘西,28.314297,109.739738;"
    AddressCity = AddressCity & "19,200,广州,23.125177,113.28064;"
    AddressCity = AddressCity & "19,201,韶关,24.801323,113.591545;"
    AddressCity = AddressCity & "19,202,深圳,22.547001,114.085945;"
    AddressCity = AddressCity & "19,203,珠海,22.224979,113.553986;"
    AddressCity = AddressCity & "19,204,汕头,23.371019,116.708466;"
    AddressCity = AddressCity & "19,205,佛山,23.028763,113.122719;"
    AddressCity = AddressCity & "19,206,江门,22.590431,113.09494;"
    AddressCity = AddressCity & "19,207,湛江,21.274899,110.364975;"
    AddressCity = AddressCity & "19,208,茂名,21.659752,110.919228;"
    AddressCity = AddressCity & "19,209,肇庆,23.051546,112.472527;"
    AddressCity = AddressCity & "19,210,惠州,23.079405,114.412598;"
    AddressCity = AddressCity & "19,211,梅州,24.299112,116.117584;"
    AddressCity = AddressCity & "19,212,汕尾,22.774485,115.364235;"
    AddressCity = AddressCity & "19,213,河源,23.746265,114.6978;"
    AddressCity = AddressCity & "19,214,阳江,21.859222,111.975105;"
    AddressCity = AddressCity & "19,215,清远,23.685022,113.051224;"
    AddressCity = AddressCity & "19,216,东莞,23.046238,113.746262;"
    AddressCity = AddressCity & "19,217,中山,22.521112,113.382393;"
    AddressCity = AddressCity & "19,218,东沙,21.810463,112.552948;"
    AddressCity = AddressCity & "19,219,潮州,23.661701,116.632301;"
    AddressCity = AddressCity & "19,220,揭阳,23.543777,116.355736;"
    AddressCity = AddressCity & "19,221,云浮,22.929802,112.044441;"
    AddressCity = AddressCity & "20,222,南宁,22.82402,108.320007;"
    AddressCity = AddressCity & "20,223,柳州,24.314617,109.411705;"
    AddressCity = AddressCity & "20,224,桂林,25.274216,110.299118;"
    AddressCity = AddressCity & "20,225,梧州,23.474804,111.297607;"
    AddressCity = AddressCity & "20,226,北海,21.473343,109.119255;"
    AddressCity = AddressCity & "20,227,防城港,21.614632,108.345474;"
    AddressCity = AddressCity & "20,228,钦州,21.967127,108.624176;"
    AddressCity = AddressCity & "20,229,贵港,23.093599,109.602142;"
    AddressCity = AddressCity & "20,230,玉林,22.631359,110.154396;"
    AddressCity = AddressCity & "20,231,百色,23.897741,106.616287;"
    AddressCity = AddressCity & "20,232,贺州,24.414141,111.552055;"
    AddressCity = AddressCity & "20,233,河池,24.695898,108.062103;"
    AddressCity = AddressCity & "20,234,来宾,23.733767,109.229774;"
    AddressCity = AddressCity & "20,235,崇左,22.404108,107.353928;"
    AddressCity = AddressCity & "21,236,海口,20.031971,110.331192;"
    AddressCity = AddressCity & "21,237,三亚,18.247871,109.50827;"
    AddressCity = AddressCity & "21,238,三沙,16.831039,112.348824;"
    AddressCity = AddressCity & "21,239,五指山,18.77692,109.516663;"
    AddressCity = AddressCity & "21,240,琼海,19.246012,110.466782;"
    AddressCity = AddressCity & "21,241,儋州,19.517487,109.576782;"
    AddressCity = AddressCity & "21,242,文昌,19.612986,110.753975;"
    AddressCity = AddressCity & "21,243,万宁,18.796215,110.388794;"
    AddressCity = AddressCity & "21,244,东方,19.10198,108.653786;"
    AddressCity = AddressCity & "21,245,定安,19.684965,110.349236;"
    AddressCity = AddressCity & "21,246,屯昌,19.362917,110.102776;"
    AddressCity = AddressCity & "21,247,澄迈,19.737095,110.007149;"
    AddressCity = AddressCity & "21,248,临高,19.908293,109.687698;"
    AddressCity = AddressCity & "21,249,白沙,19.224585,109.452606;"
    AddressCity = AddressCity & "21,250,昌江,19.260967,109.053352;"
    AddressCity = AddressCity & "21,251,乐东,18.74758,109.175446;"
    AddressCity = AddressCity & "21,252,陵水,18.505007,110.037216;"
    AddressCity = AddressCity & "21,253,保亭,18.636372,109.702454;"
    AddressCity = AddressCity & "21,254,琼中,19.03557,109.839996;"
    AddressCity = AddressCity & "22,255,重庆,29.533155,106.504959;"
    AddressCity = AddressCity & "23,256,成都,30.659462,104.065735;"
    AddressCity = AddressCity & "23,257,自贡,29.352764,104.773445;"
    AddressCity = AddressCity & "23,258,攀枝花,26.580446,101.716003;"
    AddressCity = AddressCity & "23,259,泸州,28.889137,105.443352;"
    AddressCity = AddressCity & "23,260,德阳,31.127991,104.398651;"
    AddressCity = AddressCity & "23,261,绵阳,31.46402,104.741722;"
    AddressCity = AddressCity & "23,262,广元,32.433666,105.829758;"
    AddressCity = AddressCity & "23,263,遂宁,30.513311,105.571327;"
    AddressCity = AddressCity & "23,264,内江,29.58708,105.066139;"
    AddressCity = AddressCity & "23,265,乐山,29.582024,103.761261;"
    AddressCity = AddressCity & "23,266,南充,30.79528,106.082977;"
    AddressCity = AddressCity & "23,267,眉山,30.048319,103.831787;"
    AddressCity = AddressCity & "23,268,宜宾,28.760189,104.630821;"
    AddressCity = AddressCity & "23,269,广安,30.456398,106.633369;"
    AddressCity = AddressCity & "23,270,达州,31.209484,107.502258;"
    AddressCity = AddressCity & "23,271,雅安,29.987722,103.00103;"
    AddressCity = AddressCity & "23,272,巴中,31.858809,106.75367;"
    AddressCity = AddressCity & "23,273,资阳,30.122211,104.641914;"
    AddressCity = AddressCity & "23,274,阿坝,31.899792,102.221375;"
    AddressCity = AddressCity & "23,275,甘孜,30.050663,101.963814;"
    AddressCity = AddressCity & "23,276,凉山,27.886763,102.258743;"
    AddressCity = AddressCity & "24,277,贵阳,26.578342,106.713478;"
    AddressCity = AddressCity & "24,278,六盘水,26.584642,104.846741;"
    AddressCity = AddressCity & "24,279,遵义,27.706627,106.937263;"
    AddressCity = AddressCity & "24,280,安顺,26.245544,105.93219;"
    AddressCity = AddressCity & "24,281,铜仁,27.718346,109.191551;"
    AddressCity = AddressCity & "24,282,黔西南,25.08812,104.897972;"
    AddressCity = AddressCity & "24,283,毕节,27.301693,105.285011;"
    AddressCity = AddressCity & "24,284,黔东南,26.583351,107.977486;"
    AddressCity = AddressCity & "24,285,黔南,26.258219,107.517159;"
    AddressCity = AddressCity & "25,286,昆明,25.040609,102.71225;"
    AddressCity = AddressCity & "25,287,曲靖,25.501556,103.797852;"
    AddressCity = AddressCity & "25,288,玉溪,24.35046,102.543907;"
    AddressCity = AddressCity & "25,289,保山,25.111801,99.16713;"
    AddressCity = AddressCity & "25,290,昭通,27.337,103.717216;"
    AddressCity = AddressCity & "25,291,丽江,26.872108,100.233025;"
    AddressCity = AddressCity & "25,292,普洱,22.777321,100.972343;"
    AddressCity = AddressCity & "25,293,临沧,23.886566,100.086967;"
    AddressCity = AddressCity & "25,294,楚雄,25.041988,101.546043;"
    AddressCity = AddressCity & "25,295,红河,23.366776,103.384186;"
    AddressCity = AddressCity & "25,296,文山,23.369511,104.244011;"
    AddressCity = AddressCity & "25,297,西双版纳,22.001724,100.797943;"
    AddressCity = AddressCity & "25,298,大理,25.589449,100.22567;"
    AddressCity = AddressCity & "25,299,德宏,24.436693,98.578362;"
    AddressCity = AddressCity & "25,300,怒江,25.850948,98.854301;"
    AddressCity = AddressCity & "25,301,迪庆,27.826853,99.706467;"
    AddressCity = AddressCity & "26,302,拉萨,29.66036,91.13221;"
    AddressCity = AddressCity & "26,303,昌都,31.136875,97.178452;"
    AddressCity = AddressCity & "26,304,山南,29.236023,91.766525;"
    AddressCity = AddressCity & "26,305,日喀则,29.267519,88.885147;"
    AddressCity = AddressCity & "26,306,那曲,31.476004,92.060211;"
    AddressCity = AddressCity & "26,307,阿里,32.503185,80.105499;"
    AddressCity = AddressCity & "26,308,林芝,29.654694,94.36235;"
    AddressCity = AddressCity & "27,309,西安,34.263161,108.948021;"
    AddressCity = AddressCity & "27,310,铜川,34.91658,108.979607;"
    AddressCity = AddressCity & "27,311,宝鸡,34.369316,107.144867;"
    AddressCity = AddressCity & "27,312,咸阳,34.333439,108.705116;"
    AddressCity = AddressCity & "27,313,渭南,34.499382,109.502884;"
    AddressCity = AddressCity & "27,314,延安,36.596539,109.490807;"
    AddressCity = AddressCity & "27,315,汉中,33.077667,107.028618;"
    AddressCity = AddressCity & "27,316,榆林,38.290161,109.741196;"
    AddressCity = AddressCity & "27,317,安康,32.6903,109.029274;"
    AddressCity = AddressCity & "27,318,商洛,33.86832,109.939774;"
    AddressCity = AddressCity & "28,319,兰州,36.058041,103.823555;"
    AddressCity = AddressCity & "28,320,嘉峪关,39.78653,98.277306;"
    AddressCity = AddressCity & "28,321,金昌,38.514236,102.187889;"
    AddressCity = AddressCity & "28,322,白银,36.545681,104.173607;"
    AddressCity = AddressCity & "28,323,天水,34.578529,105.724998;"
    AddressCity = AddressCity & "28,324,武威,37.929996,102.634697;"
    AddressCity = AddressCity & "28,325,张掖,38.932896,100.455475;"
    AddressCity = AddressCity & "28,326,平凉,35.542789,106.684692;"
    AddressCity = AddressCity & "28,327,酒泉,39.744022,98.510796;"
    AddressCity = AddressCity & "28,328,庆阳,35.734219,107.638374;"
    AddressCity = AddressCity & "28,329,定西,35.579578,104.626297;"
    AddressCity = AddressCity & "28,330,陇南,33.388599,104.929382;"
    AddressCity = AddressCity & "28,331,临夏,35.599445,103.212006;"
    AddressCity = AddressCity & "28,332,甘南,34.986355,102.911011;"
    AddressCity = AddressCity & "29,333,西宁,36.623177,101.778915;"
    AddressCity = AddressCity & "29,334,海东,36.502914,102.103271;"
    AddressCity = AddressCity & "29,335,海北,36.959435,100.901062;"
    AddressCity = AddressCity & "29,336,黄南,35.517742,102.019989;"
    AddressCity = AddressCity & "29,337,海南藏族,36.280354,100.619545;"
    AddressCity = AddressCity & "29,338,果洛,34.473598,100.242142;"
    AddressCity = AddressCity & "29,339,玉树,33.004047,97.008522;"
    AddressCity = AddressCity & "29,340,海西,37.374664,97.370789;"
    AddressCity = AddressCity & "30,341,银川,38.46637,106.278175;"
    AddressCity = AddressCity & "30,342,石嘴山,39.013329,106.376175;"
    AddressCity = AddressCity & "30,343,吴忠,37.986164,106.199409;"
    AddressCity = AddressCity & "30,344,固原,36.004562,106.28524;"
    AddressCity = AddressCity & "30,345,中卫,37.51495,105.189568;"
    AddressCity = AddressCity & "31,346,乌鲁木齐,43.792816,87.617729;"
    AddressCity = AddressCity & "31,347,克拉玛依,45.595886,84.873947;"
    AddressCity = AddressCity & "31,348,吐鲁番,42.947613,89.184074;"
    AddressCity = AddressCity & "31,349,哈密,42.833248,93.513161;"
    AddressCity = AddressCity & "31,350,昌吉,44.014576,87.304008;"
    AddressCity = AddressCity & "31,351,博尔塔拉,44.903259,82.074776;"
    AddressCity = AddressCity & "31,352,巴音郭楞,41.768551,86.15097;"
    AddressCity = AddressCity & "31,353,阿克苏,41.170712,80.265068;"
    AddressCity = AddressCity & "31,354,克孜勒苏柯尔克孜,39.713432,76.172829;"
    AddressCity = AddressCity & "31,355,喀什,39.467663,75.989136;"
    AddressCity = AddressCity & "31,356,和田,37.110687,79.925331;"
    AddressCity = AddressCity & "31,357,伊犁,43.92186,81.317947;"
    AddressCity = AddressCity & "31,358,塔城,46.7463,82.985733;"
    AddressCity = AddressCity & "31,359,阿勒泰,47.848392,88.139633;"
    AddressCity = AddressCity & "31,360,石河子,44.305885,86.041077;"
    AddressCity = AddressCity & "31,361,阿拉尔,40.541916,81.285881;"
    AddressCity = AddressCity & "31,362,图木舒克,39.867317,79.07798;"
    AddressCity = AddressCity & "31,363,五家渠,44.1674,87.526886;"
    AddressCity = AddressCity & "32,364,台北,25.041618,121.501618;"
    AddressCity = AddressCity & "32,365,高雄,25.041618,121.501618;"
    AddressCity = AddressCity & "32,366,台南,25.041618,121.501618;"
    AddressCity = AddressCity & "32,367,台中,25.041618,121.501618;"
    AddressCity = AddressCity & "32,368,金门,25.041618,121.501618;"
    AddressCity = AddressCity & "32,369,南投,25.041618,121.501618;"
    AddressCity = AddressCity & "32,370,基隆,25.041618,121.501618;"
    AddressCity = AddressCity & "32,371,新竹,25.041618,121.501618;"
    AddressCity = AddressCity & "32,372,嘉义,25.041618,121.501618;"
    AddressCity = AddressCity & "32,373,新北,25.041618,121.501618;"
    AddressCity = AddressCity & "32,374,宜兰,25.041618,121.501618;"
    AddressCity = AddressCity & "32,375,桃园,25.041618,121.501618;"
    AddressCity = AddressCity & "32,376,苗栗,25.041618,121.501618;"
    AddressCity = AddressCity & "32,377,彰化,25.041618,121.501618;"
    AddressCity = AddressCity & "32,378,云林,25.041618,121.501618;"
    AddressCity = AddressCity & "32,379,屏东,25.041618,121.501618;"
    AddressCity = AddressCity & "32,380,台东,25.041618,121.501618;"
    AddressCity = AddressCity & "32,381,花莲,25.041618,121.501618;"
    AddressCity = AddressCity & "32,382,澎湖,25.041618,121.501618;"
    AddressCity = AddressCity & "32,383,连江,25.041618,121.501618;"
    AddressCity = AddressCity & "33,384,香港岛,22.320047,114.173355;"
    AddressCity = AddressCity & "33,385,九龙,22.320047,114.173355;"
    AddressCity = AddressCity & "33,386,新界,22.320047,114.173355;"
    AddressCity = AddressCity & "34,387,澳门半岛,22.198751,113.549133;"
    AddressCity = AddressCity & "34,398,离岛,22.198952,113.549088"

End Function



















