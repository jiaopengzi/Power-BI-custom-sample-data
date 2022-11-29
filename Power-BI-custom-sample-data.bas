Attribute VB_Name = "demo_jiaopengzi_data"
Option Compare Database
Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'1�����ߣ�������
'2�����䣺jiaopengzi@qq.com
'3�����ͣ�www.jiaopengzi.com
'4��CPU��Intel(R) Core(TM) i7-8750H CPU @ 2.20GHz   2.21 GHz
'5���ڴ棺RAM 24.0 GB
'6���������� + N1=300�����ã���Լ��Ҫ1111�룬ÿ�밴��ҵ���߼�����Լ3500�����ݣ�����388+����demo���ݣ���������ʵսѧϰ���á�
'=====================================================================================

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'�������ã���������1926�С�
Public Const N1 As Long = 30 '�ŵ�����������N1��[5,390]��
'01��NewData            ���ú��������������󣬵������к�������demodata���ݣ������һ�γ��԰��� N1=5����Լ20�롣
'02��TableNameN         ���б����������
'03��SqlCN              �������е�sql��
'04��SqlDN              ɾ�����б��sql��
'05��TableADO           ADO������
'06��DataTableD0        ���ɴ�����
'07��DataTableD1        ����ʡ�ݱ�
'08��DataTableD2        ���ɳ��б�
'09��DataTableT0        ���ɲ�Ʒ��
'10��DataTableT1        ���ɿͻ�����N1��ء�
'11��DataTableT2        ���ɿͻ���
'12��DataTableT345      ���������������������ӱ�
'13��DataTableT6        ��������Ŀ���
'14��FirstName          �����������������
'17��LastName           ���������������
'16��AddressProvince    ����ʡ�����ݣ��������ƣ�����ȡ�
'17��AddressCity        ���е������ݣ��������ƣ�����ȡ�
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
    MsgBox "��ɣ���ʱ��" & Round(Timer - t, 2) & "�룡"

End Function


Public Function TableNameN(N As Long) As String

    Select Case N
        Case 0
            TableNameN = "T00_��Ʒ��"
        Case 1
            TableNameN = "T01_�ŵ��"
        Case 2
            TableNameN = "T02_�ͻ���"
        Case 3
            TableNameN = "T03_�����Ϣ��"
        Case 4
            TableNameN = "T04_��������"
        Case 5
            TableNameN = "T05_�����ӱ�"
        Case 6
            TableNameN = "T06_����Ŀ���"
        Case 7
            TableNameN = "D00_������"
        Case 8
            TableNameN = "D01_ʡ�ݱ�"
        Case 9
            TableNameN = "D02_���б�"
    End Select

End Function

Public Function SqlCN(N As Long) As String

    Select Case N
        '��Ʒ
        Case 0
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_��Ʒ���            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_��Ʒ����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_��Ʒ����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_��Ʒ���ۼ۸�        FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_��Ʒ�ɱ��۸�        FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '�ŵ�
        Case 1
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_�ŵ���            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_�ŵ�����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_�ŵ긺����          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_��������            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_����ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_����                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_γ��                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_08_����                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_09_�ص�����            DATE          NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '�ͻ�
        Case 2
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_�ͻ����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_�ͻ�����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_�ͻ�����            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_�ͻ��Ա�            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_ע������            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_�ͻ���ҵ            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_�ͻ�ְҵ            VARCHAR(50)   NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '���
        Case 3
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_����Ʒ���        VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_����Ʒ����        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_����ŵ���        VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_�������            DATE          NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '��������
        Case 4
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_�������            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_�ŵ���            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_�µ�����            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_�ͻ�����            DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_�ͻ����            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_��������            VARCHAR(50)   NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '�����ӱ�
        Case 5
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_�������            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_��Ʒ���            VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_��Ʒ���ۼ۸�        FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_�ۿ۱���            FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_��Ʒ��������        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_��Ʒ���۽��        FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '����Ŀ��
        Case 6
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_ʡID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_ʡ���              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_�·�                DATE          NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_����Ŀ��            FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '����
        Case 7
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_����ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_����                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_����������          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_�칫�س���ID        INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_�칫�س���          VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_γ��                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_07_����                FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        'ʡ��
        Case 8
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_����ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_ʡID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_ʡȫ��              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_ʡ���              VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_γ��                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_06_����                FLOAT         NULL" & Chr(13)
            SqlCN = SqlCN & ")"
        '����
        Case 9
            SqlCN = SqlCN & "CREATE TABLE " & TableNameN(N)
            SqlCN = SqlCN & "(" & Chr(13)
            SqlCN = SqlCN & "F_00_�Զ����            INT           NOT NULL    IDENTITY(1,1) PRIMARY KEY," & Chr(13)
            SqlCN = SqlCN & "F_01_ʡID                INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_02_����ID              INT           NULL," & Chr(13)
            SqlCN = SqlCN & "F_03_����                VARCHAR(50)   NULL," & Chr(13)
            SqlCN = SqlCN & "F_04_γ��                FLOAT         NULL," & Chr(13)
            SqlCN = SqlCN & "F_05_����                FLOAT         NULL" & Chr(13)
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

    Region0 = "1,����,ŷ������,73,�Ϻ�,31.231618,121.471618;"
    Region0 = Region0 & "2,����,������,256,�ɶ�,30.651618,104.061618;"
    Region0 = Region0 & "3,����,������,200,����,23.121618,113.281618;"
    Region0 = Region0 & "4,����,������,37,����,41.791618,123.421618;"
    Region0 = Region0 & "5,����,������,1,����,39.901618,116.401618;"
    Region0 = Region0 & "6,�۰�̨,������,386,���,22.321618,114.171618"

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
        .Source = Tnn 'ʡ�ݱ�
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
        .Source = Tnn 'ʡ�ݱ�
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
        .Source = Tnn '���б�
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
        .Source = Tnn '��Ʒ��
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    For i = 1 To N00
        Rs.AddNew
        Rs(1) = "SKU_" & Format(i, "000000")
        Randomize
        Sj = Rnd()
        
        Rs(2) = Chr(Round(Sj * 9, 0) + 65) & "��"
        Rs(3) = "��Ʒ" & Chr(Round(Sj * 9, 0) + 65) & "" & Format(i, "0000")
 
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
        .Source = Tnn '�ŵ��
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    Set Dict1 = CreateObject("Scripting.Dictionary") '��������֣���֤Ψһ���ظ���
    For i = 1 To 17576 '26*26*26
        Dict1(Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & "��") = i
        If Dict1.Count = N10 Then
            Exit For
        End If
    Next
    ArrDict1 = Dict1.Keys
    Set Dict1 = Nothing

    '�����ĸ�ֱϽ��
    Rs.AddNew: Rs(1) = "SC_0001": Rs(2) = ArrDict1(0): Rs(3) = "������": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 1: Rs(6) = "����": Rs(7) = 39.901618: Rs(8) = 116.401618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0002": Rs(2) = ArrDict1(1): Rs(3) = "������": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 2: Rs(6) = "���": Rs(7) = 39.121618: Rs(8) = 117.191618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0003": Rs(2) = ArrDict1(2): Rs(3) = "������": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 73: Rs(6) = "�Ϻ�": Rs(7) = 31.231618: Rs(8) = 121.471618: Rs.Update
    Rs.AddNew: Rs(1) = "SC_0004": Rs(2) = ArrDict1(3): Rs(3) = "������": Rs(4) = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD"): Rs(5) = 255: Rs(6) = "����": Rs(7) = 29.531618: Rs(8) = 106.501618: Rs.Update

    For i = 5 To N10
        Randomize
        Sj = Rnd()
        fnUB0 = Round(UBound(ArrFN) * Sj, 0)
        fnUB1 = Round(UBound(ArrFN) * (1 - Sj), 0)
        lnUB = Round(UBound(ArrLN) * Sj, 0)
        Randomize
        addUB0 = Round(UBound(ArrAdd0) * Rnd(), 0)
        Randomize
        dateKD = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD") '+28�ݴ�Dict3N
        Randomize
        dateGD = Format(dateKD + 550 + 4320 * Rnd(), "YYYY-MM-DD") '550��ʾ����1.5����ܹص�

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
            Rs(7) = Round(ArrAdd1(addUB0)(3) + Rnd() * 0.05, 6) '��ͬ����ƫ�ƣ�����ͬһ���㡣
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
    Dim ArrHY '��ҵ
    Dim ArrZY 'ְҵ
    Dim ArrSjNL '����ֲ�
    
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

    ArrHY = Array("����ҵ", "����ҵ", "������", "ũҵ", "����", "����", "����") '��ҵ
    ArrZY = Array("���廧", "HR", "��Ӫ", "IT", "����", "����", "�з�") 'ְҵ
    ArrSjHY = Array(0.2, 0.5, 0.5, 0.8, 0.8, 1, 0.9) '��ҵ�ֲ�
    ArrSjZY = Array(0.3, 0.7, 0.6, 1, 1, 0.8, 0.1) 'ְҵ�ֲ�
    ArrSjNL = Array(0, 0.1, 0.2, 0.3, 0.3, 0.3, 0.3, 0.8, 0.8, 0.8, 0.9, 1) '����ֲ�

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
    '���յ��̹�ģע��ͻ�
    Set Rs1 = CreateObject("ADODB.Recordset")
    With Rs1
            .ActiveConnection = Conn
            .Source = TableNameN(1) '�ŵ��
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
        .Source = Tnn '�ͻ���
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
            dateSR = Format(Now - 7500 - Round((ArrSjNL(ii Mod 12) + Rnd()) * 7000, 0), "YYYY-MM-DD")  '����
            dateZC = Format(Now - 1500 + Round((ArrSjNL(ii Mod 12) + Rnd()) * 750, 0), "YYYY-MM-DD") 'ע��ʱ�䣬�ȿ��������죬����ҵ���߼����
    
            If Sj < 0.8 Then
                XingMing = ArrLN(lnUB) & ArrFN(fnUB0)
                sex = "��"
            Else
                XingMing = ArrLN(lnUB) & ArrFN(fnUB0) & ArrFN(fnUB1)
                sex = "Ů"
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
        Dim Yyts As Long 'Ӫҵ����
        Dim dateDD As Date

        Dim Arr0
        Dim Arr1
        Dim Arr2
        Dim ArrHY '��ҵ
        Dim ArrZY 'ְҵ
        Dim ArrDjjsxs '��������ϵ��
        Dim ArrDdslxsMonth '��ҵ����������
        Dim ArrDdslxsSC '����ϵ��
        Dim ArrDict3
        Dim ArrDict5
        Dim ArrZK '�ۿ�
        Dim ArrZKMonth '�ۿ��·ݷֲ�
        Dim ArrSjKF '�ͻ��ֲ�
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
            .Source = TableNameN(0) '��Ʒ��
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
            .Source = TableNameN(1) '�ŵ��
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
    ArrHY = Array("����", "������", "����", "����ҵ") '�ۿ���ҵ׼��
    ArrZY = Array("HR", "����", "����", "��Ӫ")
    Set Dict2HY = CreateObject("Scripting.Dictionary") '��ҵ
    Set Dict2ZY = CreateObject("Scripting.Dictionary") 'ְҵ
    
    For i = 0 To UBound(ArrHY)
        Dict2HY(ArrHY(i)) = ArrHY(i)
    Next

    For i = 0 To UBound(ArrZY)
        Dict2ZY(ArrZY(i)) = ArrZY(i)
    Next

    Set Rs2 = CreateObject("ADODB.Recordset")
        With Rs2
            .ActiveConnection = Conn
            .Source = TableNameN(2) '�ͻ���
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
            .Source = TableNameN(3) '����
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs4 = CreateObject("ADODB.Recordset")
    With Rs4
            .ActiveConnection = Conn
            .Source = TableNameN(4) '��������
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
    '=====================================================================================
    
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    Set Rs5 = CreateObject("ADODB.Recordset")
        With Rs5
            .ActiveConnection = Conn
            .Source = TableNameN(5) '�����ӱ�
            .LockType = 2 'adLockPessimistic
            .CursorType = 1 'adOpenKeyset
            .Open
        End With
        '=====================================================================================
        '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        UB0 = UBound(Arr0) '��Ʒ����
        UB1 = UBound(Arr1) '�ŵ�����
        UB2 = UBound(Arr2) '�ͻ�����
        OcNumber = 0
        
        
        ArrDjjsxs = Array(0.7, 0.8, 1, 1.2, 1.3) '��������ϵ����count=5
        ArrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '��ҵ���������ƣ�count=12
        ArrDdslxsSC = Array(0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6) '���򶩵�ϵ����̫�ֲ���count=34
        ArrZK = Array(1, 0.9, 0.8, 0.7, 0.6, 0.5) '�ۿ���Ϣ�ֲ���count=6
        ArrZKMonth = Array(0.95, 0.9, 1, 0.98, 0.85, 1, 0.98, 0.88, 0.8, 0.86, 0.92, 0.98) '��ҵ���������ƣ�count=12
        ArrSjKF = Array(0, 0.1, 0.5, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.8, 0.9, 1) '�ͻ��ֲ�

        
        
    For i1 = 0 To UB1

            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            'Ӫҵ����
            If IsNull(Arr1(i1, 9)) Then
                Yyts = Round(Now - Arr1(i1, 4), 0)
            Else
                Yyts = Round(Arr1(i1, 9) - Arr1(i1, 4), 0)
            End If
            '=====================================================================================
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            Set Dict3N = CreateObject("Scripting.Dictionary") '�����������
            Dict3N(1) = Yyts Mod 6 + 1
                    '������ʱ��
                    For i = 1 To Yyts
                        Dict3N(i + 1) = Dict3N.Item(i) + Round(Rnd() * 2 + 5, 0)
                        If Dict3N.Item(i + 1) > Yyts Then
                            Dict3N(i + 1) = Yyts
                            Exit For
                        End If
                    Next
            '=====================================================================================

            Set Dict3 = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
            i3 = 1

            For i4 = 1 To Yyts

                dateDD = Arr1(i1, 4) + i4 - 1
                Randomize
                
                ND = Round(Rnd() * 4 * ArrDdslxsMonth(Month(dateDD) - 1) * ArrDdslxsSC(Arr1(i1, 5) Mod (UBound(ArrDdslxsSC) + 1)), 0) 'ÿ�충��

                If ND = 0 Then GoTo Dd0 'û������

                For i = 1 To ND 'ÿ�충����
                    OcNumber = OcNumber + 1
                    Oc = "OC_" & Format(OcNumber, "0000000")
                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '��������д��
                    Randomize
                    Sj = (Rnd() + ArrSjKF(i4 * i Mod 12)) / 2

                    UB2n = Round(UB2 * Sj, 0)
                    
                'ע���빺��ֲ�
                If IsNull(Arr1(i1, 9)) Then 'δ�ص�
                    If Arr2(UB2n, 1) >= Arr1(i1, 4) And OcNumber Mod 13 > 6 Then
                        GoTo UB2nlable
                    Else
                        For i2 = UB2n To UB2 '����ǰ��
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 13 > 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 > 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                                
                        For i2 = UB2n To 0 Step -1 '����ǰ��
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 13 <= 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 <= 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                    End If
                Else '�ص�
                    If Arr2(UB2n, 1) >= Arr1(i1, 4) And Arr2(UB2n, 1) < Arr1(i1, 9) And OcNumber Mod 13 < 6 Then
                            GoTo UB2nlable
                    Else
                        For i2 = UB2n To UB2 '����ǰ��
                            If Arr2(i2, 1) >= Arr1(i1, 4) And Arr2(i2, 1) < Arr1(i1, 9) And Month(Arr2(i2, 1)) Mod 13 > 8 Then
                                UB2n = i2
                                GoTo UB2nlable
                            ElseIf Arr2(i2, 1) >= Arr1(i1, 4) And Month(Arr2(i2, 1)) Mod 3 > 1 Then
                                UB2n = i2
                                GoTo UB2nlable
                            End If
                        Next
                            
                        For i2 = UB2n To 0 Step -1 '����ǰ��
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
                        Qd = "����"
                    Else
                        Qd = "����"
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
                    '�����ӱ�

                    Randomize

                    k = Round(5 * Rnd(), 0) + 1 '�ƻ�ÿ������Ʒ��������,��ֵΪ3

                    Set Dict5 = CreateObject("Scripting.Dictionary")
                    For N = 1 To k
                        If k < 4 Then
                            UB0n = Round(UB0 * Rnd() / 5, 0)  '����ƫ��
                        Else
                            UB0n = Round(UB0 * Rnd(), 0)
                        End If
                        Dict5(UB0n) = UB0n '�ֵ�ȥ��sku
                    Next

                    ArrDict5 = Dict5.Keys

                    For N = 0 To Dict5.Count - 1

                        p = Round(5 * Rnd() * ArrDjjsxs(i1 Mod 5) * ArrDjjsxs(ArrDict5(N) Mod 5), 0) + 1 '����ϵ����Ȩ

                        'q��Ʒ�ۿ�
                        
                        If i1 Mod 40 > 30 Then '����
                            q = ArrZK(0) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf i1 Mod 40 < 10 Then
                            q = ArrZK(5) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf ArrDict5(N) Mod 8 < 1 Then '��Ʒ
                            q = ArrZK(1) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf ArrDict5(N) Mod 8 > 5 Then
                            q = ArrZK(3) * ArrZKMonth(Month(dateDD) - 1)
                            GoTo ExitIFzk
                        ElseIf Dict2HY.exists(Arr2(UB2n, 2)) Then  '�ͻ�
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

Dd0: '�����޶�����ת
                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                '���������Ϣ
                If Dict3N.Item(i3) = i4 And i4 < Yyts Then
                    i3 = i3 + 1

                    ArrDict3 = Dict3.Keys

                    For i0 = 0 To UBound(ArrDict3)
                        Rs3.AddNew
                        Rs3(1) = ArrDict3(i0)
                        Rs3(2) = Dict3.Item(ArrDict3(i0))
                        Rs3(3) = Arr1(i1, 1)
                        Rs3(4) = Arr1(i1, 4) + i4 - 14 '-14��֤�п��
                        Rs3.Update
                    Next
                    Set Dict3 = Nothing
                    Set Dict3 = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
                    GoTo Rk0
                    
                ElseIf i4 = Yyts Then  '��֤���һ������ۼƴ���0

                    i3 = i3 + 1

                    ArrDict3 = Dict3.Keys

                    For i0 = 0 To UBound(ArrDict3)
                        Rs3.AddNew
                        Rs3(1) = ArrDict3(i0)
                        Rs3(2) = Dict3.Item(ArrDict3(i0)) + Round(Rnd() * 5, 0)
                        Rs3(3) = Arr1(i1, 1)
                        Rs3(4) = Arr1(i1, 4) + i4 - 14 '-14��֤�п��
                        Rs3.Update
                    Next
                    Set Dict3 = Nothing
                    Set Dict3 = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
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
    Dim ArrDdslxsMonth '��ҵ����������
    Dim Qn As Double '��������µ�ϵ����
    Dim B As Double
    Dim UP0 As Double '����������
    
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
    'ȥ��������ȫ��&ȥ��Q4������,�¾�ȡ��
    Sqlstr = "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "TY.*,TQ.A3Q" & Chr(13)
    Sqlstr = Sqlstr & "FROM" & Chr(13)
    Sqlstr = Sqlstr & "(" & Chr(13)
    Sqlstr = Sqlstr & "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "D01_ʡ�ݱ�.F_02_ʡID AS A0ʡID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_ʡ�ݱ�.F_04_ʡ��� AS A1ʡ���" & Chr(13)
    Sqlstr = Sqlstr & ", Sum(T05_�����ӱ�.F_06_��Ʒ���۽��) AS A2Y" & Chr(13)
    Sqlstr = Sqlstr & "FROM (((T05_�����ӱ� " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T04_�������� ON T05_�����ӱ�.F_01_������� = T04_��������.F_01_�������) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T01_�ŵ�� ON T04_��������.F_02_�ŵ��� = T01_�ŵ��.F_01_�ŵ���) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D02_���б� ON T01_�ŵ��.F_05_����ID = D02_���б�.F_02_����ID) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D01_ʡ�ݱ� ON D02_���б�.F_01_ʡID = D01_ʡ�ݱ�.F_02_ʡID" & Chr(13)
    Sqlstr = Sqlstr & "WHERE T04_��������.[F_03_�µ�����]>#" & Format(Now(), "YYYY") - 2 & "-12-1# AND T04_��������.[F_03_�µ�����]<#" & Format(Now(), "YYYY") & "-1-1#" & Chr(13)
    Sqlstr = Sqlstr & "GROUP BY " & Chr(13)
    Sqlstr = Sqlstr & "D01_ʡ�ݱ�.F_02_ʡID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_ʡ�ݱ�.F_04_ʡ���" & Chr(13)
    Sqlstr = Sqlstr & ") TY" & Chr(13)
    Sqlstr = Sqlstr & "LEFT JOIN" & Chr(13)
    Sqlstr = Sqlstr & "(" & Chr(13)
    Sqlstr = Sqlstr & "SELECT" & Chr(13)
    Sqlstr = Sqlstr & "D01_ʡ�ݱ�.F_02_ʡID AS A0ʡID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_ʡ�ݱ�.F_04_ʡ��� AS A1ʡ���" & Chr(13)
    Sqlstr = Sqlstr & ", Sum(T05_�����ӱ�.F_06_��Ʒ���۽��) AS A3Q" & Chr(13)
    Sqlstr = Sqlstr & "FROM (((T05_�����ӱ� " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T04_�������� ON T05_�����ӱ�.F_01_������� = T04_��������.F_01_�������) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN T01_�ŵ�� ON T04_��������.F_02_�ŵ��� = T01_�ŵ��.F_01_�ŵ���) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D02_���б� ON T01_�ŵ��.F_05_����ID = D02_���б�.F_02_����ID) " & Chr(13)
    Sqlstr = Sqlstr & "INNER JOIN D01_ʡ�ݱ� ON D02_���б�.F_01_ʡID = D01_ʡ�ݱ�.F_02_ʡID" & Chr(13)
    Sqlstr = Sqlstr & "WHERE T04_��������.[F_03_�µ�����]>#" & Format(Now(), "YYYY") - 1 & "-9-1# AND T04_��������.[F_03_�µ�����]<#" & Format(Now(), "YYYY") & "-1-1#" & Chr(13)
    Sqlstr = Sqlstr & "GROUP BY " & Chr(13)
    Sqlstr = Sqlstr & "D01_ʡ�ݱ�.F_02_ʡID" & Chr(13)
    Sqlstr = Sqlstr & ", D01_ʡ�ݱ�.F_04_ʡ���" & Chr(13)
    Sqlstr = Sqlstr & ") TQ" & Chr(13)
    Sqlstr = Sqlstr & "ON TY.A0ʡID=TQ.A0ʡID AND TY.A1ʡ���=TQ.A1ʡ���" & Chr(13)

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
    ArrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '��ҵ����������,��һ��count=12;ͬ��DataTableT345

    Set Rs = CreateObject("ADODB.Recordset")
        
    With Rs
        .ActiveConnection = Conn
        .Source = Tnn '����Ŀ���
        .LockType = 2 'adLockPessimistic
        .CursorType = 1 'adOpenKeyset
        .Open
    End With
    
    
    For i = UBound(ArrDdslxsMonth) - 2 To UBound(ArrDdslxsMonth)
        Qn = ArrDdslxsMonth(i) + Qn
    Next
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '����ȥ��Ŀ��
    For i = 0 To UBound(ArrYQ)

        If IsNull(ArrYQ(i, 3)) Then
            B = ArrYQ(i, 2) / 12
        ElseIf ArrYQ(i, 2) / 12 > ArrYQ(i, 3) / Qn Then '�¾�ȡ��
            B = ArrYQ(i, 2) / 12
        Else
            B = ArrYQ(i, 3) / Qn
        End If
        
        'ʵ��ֵ������1���·���UP0��������ࡣ
        UP0 = 1
        
        For k = 1 To 12
            Rs.AddNew
            Rs(1) = ArrYQ(i, 0)
            Rs(2) = ArrYQ(i, 1)
            Rs(3) = Format(Now(), "YYYY") - 1 & "-" & k & "-1"
            Randomize
            Rs(4) = Round(B * (0.7 + Rnd() * 0.1 * UP0) * ArrDdslxsMonth(k - 1), 0) '�·�����Ŀ��ı�����һ�������궼����ʵ�ʸ�������
            Rs.Update
        Next
    Next
    '=====================================================================================
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '����Ŀ��
    For i = 0 To UBound(ArrYQ)

        If IsNull(ArrYQ(i, 3)) Then
            B = ArrYQ(i, 2) / 12
        ElseIf ArrYQ(i, 2) / 12 > ArrYQ(i, 3) / Qn Then '�¾�ȡ��
            B = ArrYQ(i, 2) / 12
        Else
            B = ArrYQ(i, 3) / Qn
        End If
        
        If ArrYQ(i, 0) < 5 Then '����������
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

    FirstName = "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�A,��,��,��,��,��,��,�,��,"
    FirstName = FirstName & "��,��,��,��,��,��,�,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,�,��,��,��,�,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,����,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,٤,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,ھ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ݸ,�ݾ�,��,��,��,��,��,��,�,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,ݦ��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,"
    FirstName = FirstName & "��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ݼ,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ݰ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,"
    FirstName = FirstName & "��,��,��,¦,¯,·,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ä,ð,ó,ö,÷,ý,�,ú,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,ģ,Ħ,ĩ,Į,ī,Ĳ,ĸ,Ŀ,��,��,Ļ,��,Ľ,ĺ,��,��,��,��,��,��,��,��,��,�,�,��,��,��,��,��,��,��,��,�,Ŧ,ũ,Ū,ū,Ŭ,ů,ŷ,ż,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,ا,��,��,ƪ,Ʊ,Ư,Ƹ,ƽ,��,ƺ,��,��,��,��,��,��,��,��,��,��,��,��,�,��,ٹ,��,��,��,��,��,�,��,�,��,��,Ǣ,Ǫ,ǫ,�,ǰ,Ǭ,ǿ,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,�,��,۾,ȡ,Ȣ,ȥ,Ȥ,Ȫ,��,ȴ,ȷ,Ȼ,Ƚ,Ⱦ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,ɣ,ɪ,ɳ,ɴ,ɲ,ɰ,��,ɺ,��,��,��,��,��,��,��,ۿ,��,��,��,˭,��,��,��,��,��,��,��,��,��,ʡ,ʨ,ʩ,ʯ,ʵ,ʳ,ݪ,ʷ,ʹ,ʼ,ʾ,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,�,��,�,��,��,��,��,��,��,��,��,��,����,��,��,��,��,˦,˧,ˬ,ˮ,˯,˵,˷,˶,˾,˼,��,��,��,��,��,��,��,ݿ,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,̨,̭,̳,̿,̽,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,٬,��,��,�,��,��,͡,͢,ͤ,ͥ,ͣ,"
    FirstName = FirstName & "ͦ,��,ͨ,ͬ,١,ͮ,ͩ,Ͱ,͹,ͻ,ͼ,ͽ,Ϳ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,Τ,Ψ,�,ΰ,έ,β,γ,��,ί,�,"
    FirstName = FirstName & "��,δ,λ,ζ,η,θ,ν,ξ,μ,ο,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,أ,��,��,��,��,��,��,Ϭ,Ϫ,��,��,ϰ,ϲ,��,ϵ,"
    FirstName = FirstName & "ϸ,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Ф,У,Х,Э,��,г,е,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,��,��,��,��,��,��,��,�,Ѩ,ѧ,ѫ,Ѭ,ѯ,Ѻ,ѿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Ҧ,ҡ,��,"
    FirstName = FirstName & "��,Ҫ,Ҭ,Ү,ұ,Ұ,ҵ,Ҷ,ҳ,ҹ,��,��,��,��,ڱ,��,��,��,��,��,��,߮,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ӡ,ط,"
    FirstName = FirstName & "Ӣ,��,ӭ,ӯ,Ө,ө,Ӯ,۫,�,Ӱ,ӳ,ӹ,Ӻ,��,��,ӽ,Ӿ,��,ӿ,��,��,��,��,��,��,��,��,��,ݬ,��,��,٧,��,��,��,��,��,�,��,��,��,��,�,��,��,��,��,��,��,��,"
    FirstName = FirstName & "��,��,��,��,��,Ԥ,��,��,ԣ,��,ԥ,԰,Ա,��,Ԭ,ԭ,Բ,Ԯ,Ե,Դ,Է,Ժ,�,Ը,Լ,��,��,��,��,ܿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,դ,լ,կ,"
    FirstName = FirstName & "�,ղ,չ,ո,ռ,ջ,ս,վ,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,֡,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ں,֥,֦,֪,ֱ,"
    FirstName = FirstName & "ֵ,ֲ,ֳ,ֻ,ֽ,��,ָ,ֺ,��,־,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,٪,��,��,��,��,��,��,ס,��,��,ע,��,��,ף,ר,׫,ױ,ׯ,װ,׳,׻,׼,"
    FirstName = FirstName & "��,׿,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��"

End Function

Public Function LastName() As String

    LastName = "��,��,��,��,��,��,����,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,����,��,��,��,��,��,"
    LastName = LastName & "����,�̨,��,��,��,��,����,��,��,��,����,����,����,��,��,��,�,��,��,��,��ľ,��,�θ�,��,٦,��,��,��,��,��,��,��,��,ۺ,��,��,��,��,��,��,��,��,��,��,��,"
    LastName = LastName & "��,��,��,��,��,����,����,����,����,��ұ,��,��,��,��,��,��,��,��,����,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,����,��,��,��,��,��,��,��,����,��,"
    LastName = LastName & "��,��,��,��,��,��,��,�ʸ�,��,��,��,��,��,��,��,��,��,��,��,��,��,��,�й�,��,ۣ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,"
    LastName = LastName & "��,��,��,��,��,��,��,��,����,��,��,��,��,��,��,۪,��,��,��,����,��,��,��,��,���,��,��,��,¡,¦,¬,³,½,��,»,·,��,��,��,����,��,��,��,��,ë,é,÷,��,"
    LastName = LastName & "��,��,��,�,��,ؿ,��,��,Ī,ī,Ĳ,��,��,Ľ,Ľ��,��,��,�Ϲ�,����,��,��,��,��,��,ţ,ť,ũ,ŷ,ŷ��,��,��,��,��,��,��,Ƥ,ƽ,��,�,���,��,��,���,����,��,��,"
    LastName = LastName & "Ǯ,ǿ,��,��,��,��,��,��,��,��,��,��,�,��,Ȩ,ȫ,��,Ƚ,����,��,��,��,��,��,��,��,��,��,��,ɣ,ɳ,ɽ,��,��,�Ϲ�,��,��,��,��,��,��,����,ݷ,��,��,ʢ,ʦ,ʩ,ʯ,"
    LastName = LastName & "ʱ,ʷ,��,�,��,��,˧,˫,ˮ,˾,˾��,˾��,˾��,˾ͽ,��,��,��,��,��,��,�ذ�,ۢ,̫��,̸,̷,��,��,��,��,��,ͨ,١,ͯ,Ϳ,��,��,��ٹ,��,��,Σ,΢��,Τ,��,��,ξ��,"
    LastName = LastName & "ε,κ,��,��,��,����,��,��,��,��,��,����,��,��,��,��,����,ۭ,��,ϰ,ϯ,�S,��,�ĺ�,����,��,��,��,��,��,л,��,��,��,��,��,��,��,��,��ԯ,��,Ѧ,��,۳,��,��,��,"
    LastName = LastName & "��,��,��,��,��,����,��,��,��,��,Ҧ,Ҷ,��,��,��,��,��,��,��,ӡ,Ӧ,Ӻ,��,��,��,��,��,��,�,��,��,����,��,��,��,��,Ԫ,Ԭ,��,Խ,��,��,�׸�,��,�,ղ,տ,��,��,"
    LastName = LastName & "����,��,��,��,֣,֧,��,��,����,��,����,��,��,��,���,��,ף,���,ׯ,׿,�ӳ�,��,��,����,��,��,��,����"

End Function


Public Function AddressProvince() As String

    '����ID ʡID    ʡȫ��  ʡ���  γ��    ����
    AddressProvince = "5,1,������,����,39.904987,116.405289;"
    AddressProvince = AddressProvince & "5,2,�����,���,39.125595,117.190186;"
    AddressProvince = AddressProvince & "5,3,�ӱ�ʡ,�ӱ�,38.045475,114.502464;"
    AddressProvince = AddressProvince & "5,4,ɽ��ʡ,ɽ��,37.857014,112.549248;"
    AddressProvince = AddressProvince & "4,5,���ɹ�������,���ɹ�,40.81831,111.670799;"
    AddressProvince = AddressProvince & "4,6,����ʡ,����,41.796768,123.429092;"
    AddressProvince = AddressProvince & "4,7,����ʡ,����,43.886841,125.324501;"
    AddressProvince = AddressProvince & "4,8,������ʡ,������,45.756966,126.642464;"
    AddressProvince = AddressProvince & "1,9,�Ϻ���,�Ϻ�,31.231707,121.472641;"
    AddressProvince = AddressProvince & "1,10,����ʡ,����,32.041546,118.76741;"
    AddressProvince = AddressProvince & "1,11,�㽭ʡ,�㽭,30.287458,120.15358;"
    AddressProvince = AddressProvince & "1,12,����ʡ,����,31.861191,117.283043;"
    AddressProvince = AddressProvince & "1,13,����ʡ,����,26.075302,119.306236;"
    AddressProvince = AddressProvince & "1,14,����ʡ,����,28.676493,115.892151;"
    AddressProvince = AddressProvince & "4,15,ɽ��ʡ,ɽ��,36.675808,117.000923;"
    AddressProvince = AddressProvince & "5,16,����ʡ,����,34.757977,113.665413;"
    AddressProvince = AddressProvince & "5,17,����ʡ,����,30.584354,114.298569;"
    AddressProvince = AddressProvince & "3,18,����ʡ,����,28.19409,112.982277;"
    AddressProvince = AddressProvince & "3,19,�㶫ʡ,�㶫,23.125177,113.28064;"
    AddressProvince = AddressProvince & "3,20,����׳��������,����,22.82402,108.320007;"
    AddressProvince = AddressProvince & "3,21,����ʡ,����,20.031971,110.331192;"
    AddressProvince = AddressProvince & "2,22,������,����,29.533155,106.504959;"
    AddressProvince = AddressProvince & "2,23,�Ĵ�ʡ,�Ĵ�,30.659462,104.065735;"
    AddressProvince = AddressProvince & "3,24,����ʡ,����,26.578342,106.713478;"
    AddressProvince = AddressProvince & "3,25,����ʡ,����,25.040609,102.71225;"
    AddressProvince = AddressProvince & "2,26,����������,����,29.66036,91.13221;"
    AddressProvince = AddressProvince & "5,27,����ʡ,����,34.263161,108.948021;"
    AddressProvince = AddressProvince & "2,28,����ʡ,����,36.058041,103.823555;"
    AddressProvince = AddressProvince & "2,29,�ຣʡ,�ຣ,36.623177,101.778915;"
    AddressProvince = AddressProvince & "2,30,���Ļ���������,����,38.46637,106.278175;"
    AddressProvince = AddressProvince & "2,31,�½�ά���������,�½�,43.792816,87.617729;"
    AddressProvince = AddressProvince & "6,32,̨��ʡ,̨��,25.041618,121.501618;"
    AddressProvince = AddressProvince & "6,33,����ر�������,���,22.320047,114.173355;"
    AddressProvince = AddressProvince & "6,34,�����ر�������,����,22.198952,113.549088"

End Function


Public Function AddressCity() As String

    'ʡID    ����ID  ����    γ��    ����
    AddressCity = "1,1,����,39.904987,116.405289;"
    AddressCity = AddressCity & "2,2,���,39.125595,117.190186;"
    AddressCity = AddressCity & "3,3,ʯ��ׯ,38.045475,114.502464;"
    AddressCity = AddressCity & "3,4,��ɽ,39.635113,118.175392;"
    AddressCity = AddressCity & "3,5,�ػʵ�,39.942532,119.586578;"
    AddressCity = AddressCity & "3,6,����,36.612274,114.490685;"
    AddressCity = AddressCity & "3,7,��̨,37.068199,114.50885;"
    AddressCity = AddressCity & "3,8,����,38.867657,115.48233;"
    AddressCity = AddressCity & "3,9,�żҿ�,40.811901,114.884094;"
    AddressCity = AddressCity & "3,10,�е�,40.976204,117.939156;"
    AddressCity = AddressCity & "3,11,����,38.310581,116.85746;"
    AddressCity = AddressCity & "3,12,�ȷ�,39.523926,116.704437;"
    AddressCity = AddressCity & "3,13,��ˮ,37.735096,115.665993;"
    AddressCity = AddressCity & "4,14,̫ԭ,37.857014,112.549248;"
    AddressCity = AddressCity & "4,15,��ͬ,40.090309,113.295258;"
    AddressCity = AddressCity & "4,16,��Ȫ,37.861187,113.583282;"
    AddressCity = AddressCity & "4,17,����,36.191113,113.113556;"
    AddressCity = AddressCity & "4,18,����,35.497555,112.851273;"
    AddressCity = AddressCity & "4,19,˷��,39.331261,112.433388;"
    AddressCity = AddressCity & "4,20,����,37.696495,112.736465;"
    AddressCity = AddressCity & "4,21,�˳�,35.022778,111.00396;"
    AddressCity = AddressCity & "4,22,����,38.41769,112.733536;"
    AddressCity = AddressCity & "4,23,�ٷ�,36.084148,111.517975;"
    AddressCity = AddressCity & "4,24,����,37.524364,111.134338;"
    AddressCity = AddressCity & "5,25,���ͺ���,40.81831,111.670799;"
    AddressCity = AddressCity & "5,26,��ͷ,40.658169,109.840408;"
    AddressCity = AddressCity & "5,27,�ں�,39.673733,106.825562;"
    AddressCity = AddressCity & "5,28,���,42.275318,118.956802;"
    AddressCity = AddressCity & "5,29,ͨ��,43.617428,122.263123;"
    AddressCity = AddressCity & "5,30,������˹,39.817181,109.990288;"
    AddressCity = AddressCity & "5,31,���ױ���,49.215332,119.758171;"
    AddressCity = AddressCity & "5,32,�����׶�,40.757401,107.416962;"
    AddressCity = AddressCity & "5,33,�����첼,41.034126,113.11454;"
    AddressCity = AddressCity & "5,34,�˰�,46.076267,122.07032;"
    AddressCity = AddressCity & "5,35,���ֹ���,43.944019,116.090996;"
    AddressCity = AddressCity & "5,36,������,38.844814,105.706421;"
    AddressCity = AddressCity & "6,37,����,41.796768,123.429092;"
    AddressCity = AddressCity & "6,38,����,38.914589,121.618622;"
    AddressCity = AddressCity & "6,39,��ɽ,41.110626,122.995628;"
    AddressCity = AddressCity & "6,40,��˳,41.875957,123.921112;"
    AddressCity = AddressCity & "6,41,��Ϫ,41.297909,123.770515;"
    AddressCity = AddressCity & "6,42,����,40.124294,124.383041;"
    AddressCity = AddressCity & "6,43,����,41.11927,121.135742;"
    AddressCity = AddressCity & "6,44,Ӫ��,40.667431,122.235153;"
    AddressCity = AddressCity & "6,45,����,42.011795,121.648964;"
    AddressCity = AddressCity & "6,46,����,41.269402,123.181519;"
    AddressCity = AddressCity & "6,47,�̽�,41.124485,122.069572;"
    AddressCity = AddressCity & "6,48,����,42.290585,123.844276;"
    AddressCity = AddressCity & "6,49,����,41.576759,120.45118;"
    AddressCity = AddressCity & "6,50,��«��,40.755573,120.856392;"
    AddressCity = AddressCity & "7,51,����,43.886841,125.324501;"
    AddressCity = AddressCity & "7,52,����,43.843578,126.553017;"
    AddressCity = AddressCity & "7,53,��ƽ,43.170345,124.370789;"
    AddressCity = AddressCity & "7,54,��Դ,42.902691,125.145348;"
    AddressCity = AddressCity & "7,55,ͨ��,41.721176,125.936501;"
    AddressCity = AddressCity & "7,56,��ɽ,41.942505,126.427841;"
    AddressCity = AddressCity & "7,57,��ԭ,45.118244,124.823608;"
    AddressCity = AddressCity & "7,58,�׳�,45.619026,122.84111;"
    AddressCity = AddressCity & "7,59,�ӱ߳�����,42.904823,129.513229;"
    AddressCity = AddressCity & "8,60,������,45.756966,126.642464;"
    AddressCity = AddressCity & "8,61,�������,47.342079,123.957916;"
    AddressCity = AddressCity & "8,62,����,45.300045,130.975967;"
    AddressCity = AddressCity & "8,63,�׸�,47.332085,130.277481;"
    AddressCity = AddressCity & "8,64,˫Ѽɽ,46.64344,131.157303;"
    AddressCity = AddressCity & "8,65,����,46.590733,125.112717;"
    AddressCity = AddressCity & "8,66,����,47.724773,128.899399;"
    AddressCity = AddressCity & "8,67,��ľ˹,46.809605,130.361633;"
    AddressCity = AddressCity & "8,68,��̨��,45.771267,131.015579;"
    AddressCity = AddressCity & "8,69,ĵ����,44.582962,129.618607;"
    AddressCity = AddressCity & "8,70,�ں�,50.249584,127.499023;"
    AddressCity = AddressCity & "8,71,�绯,46.637394,126.992928;"
    AddressCity = AddressCity & "8,72,���˰���,52.335262,124.711525;"
    AddressCity = AddressCity & "9,73,�Ϻ�,31.231707,121.472641;"
    AddressCity = AddressCity & "10,74,�Ͼ�,32.041546,118.76741;"
    AddressCity = AddressCity & "10,75,����,31.57473,120.301666;"
    AddressCity = AddressCity & "10,76,����,34.261791,117.184814;"
    AddressCity = AddressCity & "10,77,����,31.772753,119.946976;"
    AddressCity = AddressCity & "10,78,����,31.299379,120.619583;"
    AddressCity = AddressCity & "10,79,��ͨ,32.016212,120.864609;"
    AddressCity = AddressCity & "10,80,���Ƹ�,34.600018,119.178818;"
    AddressCity = AddressCity & "10,81,����,33.597507,119.021263;"
    AddressCity = AddressCity & "10,82,�γ�,33.377632,120.139999;"
    AddressCity = AddressCity & "10,83,����,32.393158,119.421005;"
    AddressCity = AddressCity & "10,84,��,32.204403,119.452751;"
    AddressCity = AddressCity & "10,85,̩��,32.484882,119.915176;"
    AddressCity = AddressCity & "10,86,��Ǩ,33.963009,118.275162;"
    AddressCity = AddressCity & "11,87,����,30.287458,120.15358;"
    AddressCity = AddressCity & "11,88,����,29.868387,121.549789;"
    AddressCity = AddressCity & "11,89,����,28.000574,120.672112;"
    AddressCity = AddressCity & "11,90,����,30.762653,120.750862;"
    AddressCity = AddressCity & "11,91,����,30.867199,120.102402;"
    AddressCity = AddressCity & "11,92,����,29.997116,120.582115;"
    AddressCity = AddressCity & "11,93,��,29.089523,119.649506;"
    AddressCity = AddressCity & "11,94,����,28.941708,118.872627;"
    AddressCity = AddressCity & "11,95,��ɽ,30.016027,122.106865;"
    AddressCity = AddressCity & "11,96,̨��,28.661379,121.428596;"
    AddressCity = AddressCity & "11,97,��ˮ,28.451994,119.921783;"
    AddressCity = AddressCity & "12,98,�Ϸ�,31.861191,117.283043;"
    AddressCity = AddressCity & "12,99,�ߺ�,31.326319,118.37645;"
    AddressCity = AddressCity & "12,100,����,32.939667,117.363228;"
    AddressCity = AddressCity & "12,101,����,32.647575,117.018326;"
    AddressCity = AddressCity & "12,102,��ɽ,31.689362,118.507904;"
    AddressCity = AddressCity & "12,103,����,33.971706,116.794662;"
    AddressCity = AddressCity & "12,104,ͭ��,30.929935,117.816574;"
    AddressCity = AddressCity & "12,105,����,30.508829,117.043549;"
    AddressCity = AddressCity & "12,106,��ɽ,29.709238,118.317322;"
    AddressCity = AddressCity & "12,107,����,32.303627,118.316261;"
    AddressCity = AddressCity & "12,108,����,32.896969,115.819733;"
    AddressCity = AddressCity & "12,109,����,33.633892,116.984085;"
    AddressCity = AddressCity & "12,110,����,31.75289,116.507675;"
    AddressCity = AddressCity & "12,111,����,33.869339,115.782936;"
    AddressCity = AddressCity & "12,112,����,30.656036,117.489159;"
    AddressCity = AddressCity & "12,113,����,30.945667,118.757996;"
    AddressCity = AddressCity & "13,114,����,26.075302,119.306236;"
    AddressCity = AddressCity & "13,115,����,24.490475,118.110222;"
    AddressCity = AddressCity & "13,116,����,25.431011,119.007561;"
    AddressCity = AddressCity & "13,117,����,26.265444,117.635002;"
    AddressCity = AddressCity & "13,118,Ȫ��,24.908854,118.589424;"
    AddressCity = AddressCity & "13,119,����,24.510897,117.661804;"
    AddressCity = AddressCity & "13,120,��ƽ,26.635628,118.178459;"
    AddressCity = AddressCity & "13,121,����,25.091602,117.029778;"
    AddressCity = AddressCity & "13,122,����,26.659241,119.527084;"
    AddressCity = AddressCity & "14,123,�ϲ�,28.676493,115.892151;"
    AddressCity = AddressCity & "14,124,������,29.292561,117.214661;"
    AddressCity = AddressCity & "14,125,Ƽ��,27.622946,113.852188;"
    AddressCity = AddressCity & "14,126,�Ž�,29.712034,115.992813;"
    AddressCity = AddressCity & "14,127,����,27.810835,114.930832;"
    AddressCity = AddressCity & "14,128,ӥ̶,28.238638,117.033836;"
    AddressCity = AddressCity & "14,129,����,25.850969,114.940277;"
    AddressCity = AddressCity & "14,130,����,27.111698,114.986374;"
    AddressCity = AddressCity & "14,131,�˴�,27.8043,114.391136;"
    AddressCity = AddressCity & "14,132,����,27.98385,116.358353;"
    AddressCity = AddressCity & "14,133,����,28.44442,117.971184;"
    AddressCity = AddressCity & "15,134,����,36.675808,117.000923;"
    AddressCity = AddressCity & "15,135,�ൺ,36.082981,120.355171;"
    AddressCity = AddressCity & "15,136,�Ͳ�,36.814938,118.047646;"
    AddressCity = AddressCity & "15,137,��ׯ,34.856422,117.557961;"
    AddressCity = AddressCity & "15,138,��Ӫ,37.434563,118.664711;"
    AddressCity = AddressCity & "15,139,��̨,37.539295,121.39138;"
    AddressCity = AddressCity & "15,140,Ϋ��,36.709251,119.107079;"
    AddressCity = AddressCity & "15,141,����,35.415394,116.587242;"
    AddressCity = AddressCity & "15,142,̩��,36.194969,117.129066;"
    AddressCity = AddressCity & "15,143,����,37.509689,122.116394;"
    AddressCity = AddressCity & "15,144,����,35.428589,119.461205;"
    AddressCity = AddressCity & "15,145,����,36.214397,117.677734;"
    AddressCity = AddressCity & "15,146,����,35.065281,118.326447;"
    AddressCity = AddressCity & "15,147,����,37.453968,116.307426;"
    AddressCity = AddressCity & "15,148,�ĳ�,36.456013,115.98037;"
    AddressCity = AddressCity & "15,149,����,37.383541,118.016975;"
    AddressCity = AddressCity & "15,150,����,35.246532,115.469383;"
    AddressCity = AddressCity & "16,151,֣��,34.757977,113.665413;"
    AddressCity = AddressCity & "16,152,����,34.79705,114.341446;"
    AddressCity = AddressCity & "16,153,����,34.66304,112.434471;"
    AddressCity = AddressCity & "16,154,ƽ��ɽ,33.735241,113.307716;"
    AddressCity = AddressCity & "16,155,����,36.103443,114.352486;"
    AddressCity = AddressCity & "16,156,�ױ�,35.748238,114.295441;"
    AddressCity = AddressCity & "16,157,����,35.302616,113.883987;"
    AddressCity = AddressCity & "16,158,����,35.23904,113.238266;"
    AddressCity = AddressCity & "16,159,��Դ,35.090378,112.59005;"
    AddressCity = AddressCity & "16,160,���,35.768234,115.041298;"
    AddressCity = AddressCity & "16,161,���,34.022957,113.826065;"
    AddressCity = AddressCity & "16,162,���,33.575855,114.026405;"
    AddressCity = AddressCity & "16,163,����Ͽ,34.777336,111.194099;"
    AddressCity = AddressCity & "16,164,����,32.999081,112.540916;"
    AddressCity = AddressCity & "16,165,����,34.437054,115.650497;"
    AddressCity = AddressCity & "16,166,����,32.123276,114.075027;"
    AddressCity = AddressCity & "16,167,�ܿ�,33.620358,114.649651;"
    AddressCity = AddressCity & "16,168,פ���,32.980167,114.024734;"
    AddressCity = AddressCity & "17,169,�人,30.584354,114.298569;"
    AddressCity = AddressCity & "17,170,��ʯ,30.220074,115.077049;"
    AddressCity = AddressCity & "17,171,ʮ��,32.646908,110.787918;"
    AddressCity = AddressCity & "17,172,�˲�,30.702637,111.29084;"
    AddressCity = AddressCity & "17,173,����,32.042427,112.14415;"
    AddressCity = AddressCity & "17,174,����,30.396536,114.890594;"
    AddressCity = AddressCity & "17,175,����,31.035419,112.204254;"
    AddressCity = AddressCity & "17,176,Т��,30.926422,113.926659;"
    AddressCity = AddressCity & "17,177,����,30.326857,112.238129;"
    AddressCity = AddressCity & "17,178,�Ƹ�,30.447712,114.879364;"
    AddressCity = AddressCity & "17,179,����,29.832798,114.328964;"
    AddressCity = AddressCity & "17,180,����,31.717497,113.373772;"
    AddressCity = AddressCity & "17,181,��ʩ,30.283113,109.486992;"
    AddressCity = AddressCity & "17,182,����,30.364952,113.453972;"
    AddressCity = AddressCity & "17,183,Ǳ��,30.421215,112.896866;"
    AddressCity = AddressCity & "17,184,����,30.653061,113.165863;"
    AddressCity = AddressCity & "17,185,��ũ��,30.584354,114.298569;"
    AddressCity = AddressCity & "18,186,��ɳ,28.19409,112.982277;"
    AddressCity = AddressCity & "18,187,����,27.835806,113.151733;"
    AddressCity = AddressCity & "18,188,��̶,27.829729,112.944054;"
    AddressCity = AddressCity & "18,189,����,26.900358,112.607697;"
    AddressCity = AddressCity & "18,190,����,27.237843,111.469231;"
    AddressCity = AddressCity & "18,191,����,29.370291,113.132858;"
    AddressCity = AddressCity & "18,192,����,29.040224,111.691345;"
    AddressCity = AddressCity & "18,193,�żҽ�,29.127401,110.479919;"
    AddressCity = AddressCity & "18,194,����,28.570066,112.355042;"
    AddressCity = AddressCity & "18,195,����,25.793589,113.032066;"
    AddressCity = AddressCity & "18,196,����,26.434517,111.608017;"
    AddressCity = AddressCity & "18,197,����,27.550081,109.978241;"
    AddressCity = AddressCity & "18,198,¦��,27.728136,112.008499;"
    AddressCity = AddressCity & "18,199,����,28.314297,109.739738;"
    AddressCity = AddressCity & "19,200,����,23.125177,113.28064;"
    AddressCity = AddressCity & "19,201,�ع�,24.801323,113.591545;"
    AddressCity = AddressCity & "19,202,����,22.547001,114.085945;"
    AddressCity = AddressCity & "19,203,�麣,22.224979,113.553986;"
    AddressCity = AddressCity & "19,204,��ͷ,23.371019,116.708466;"
    AddressCity = AddressCity & "19,205,��ɽ,23.028763,113.122719;"
    AddressCity = AddressCity & "19,206,����,22.590431,113.09494;"
    AddressCity = AddressCity & "19,207,տ��,21.274899,110.364975;"
    AddressCity = AddressCity & "19,208,ï��,21.659752,110.919228;"
    AddressCity = AddressCity & "19,209,����,23.051546,112.472527;"
    AddressCity = AddressCity & "19,210,����,23.079405,114.412598;"
    AddressCity = AddressCity & "19,211,÷��,24.299112,116.117584;"
    AddressCity = AddressCity & "19,212,��β,22.774485,115.364235;"
    AddressCity = AddressCity & "19,213,��Դ,23.746265,114.6978;"
    AddressCity = AddressCity & "19,214,����,21.859222,111.975105;"
    AddressCity = AddressCity & "19,215,��Զ,23.685022,113.051224;"
    AddressCity = AddressCity & "19,216,��ݸ,23.046238,113.746262;"
    AddressCity = AddressCity & "19,217,��ɽ,22.521112,113.382393;"
    AddressCity = AddressCity & "19,218,��ɳ,21.810463,112.552948;"
    AddressCity = AddressCity & "19,219,����,23.661701,116.632301;"
    AddressCity = AddressCity & "19,220,����,23.543777,116.355736;"
    AddressCity = AddressCity & "19,221,�Ƹ�,22.929802,112.044441;"
    AddressCity = AddressCity & "20,222,����,22.82402,108.320007;"
    AddressCity = AddressCity & "20,223,����,24.314617,109.411705;"
    AddressCity = AddressCity & "20,224,����,25.274216,110.299118;"
    AddressCity = AddressCity & "20,225,����,23.474804,111.297607;"
    AddressCity = AddressCity & "20,226,����,21.473343,109.119255;"
    AddressCity = AddressCity & "20,227,���Ǹ�,21.614632,108.345474;"
    AddressCity = AddressCity & "20,228,����,21.967127,108.624176;"
    AddressCity = AddressCity & "20,229,���,23.093599,109.602142;"
    AddressCity = AddressCity & "20,230,����,22.631359,110.154396;"
    AddressCity = AddressCity & "20,231,��ɫ,23.897741,106.616287;"
    AddressCity = AddressCity & "20,232,����,24.414141,111.552055;"
    AddressCity = AddressCity & "20,233,�ӳ�,24.695898,108.062103;"
    AddressCity = AddressCity & "20,234,����,23.733767,109.229774;"
    AddressCity = AddressCity & "20,235,����,22.404108,107.353928;"
    AddressCity = AddressCity & "21,236,����,20.031971,110.331192;"
    AddressCity = AddressCity & "21,237,����,18.247871,109.50827;"
    AddressCity = AddressCity & "21,238,��ɳ,16.831039,112.348824;"
    AddressCity = AddressCity & "21,239,��ָɽ,18.77692,109.516663;"
    AddressCity = AddressCity & "21,240,��,19.246012,110.466782;"
    AddressCity = AddressCity & "21,241,����,19.517487,109.576782;"
    AddressCity = AddressCity & "21,242,�Ĳ�,19.612986,110.753975;"
    AddressCity = AddressCity & "21,243,����,18.796215,110.388794;"
    AddressCity = AddressCity & "21,244,����,19.10198,108.653786;"
    AddressCity = AddressCity & "21,245,����,19.684965,110.349236;"
    AddressCity = AddressCity & "21,246,�Ͳ�,19.362917,110.102776;"
    AddressCity = AddressCity & "21,247,����,19.737095,110.007149;"
    AddressCity = AddressCity & "21,248,�ٸ�,19.908293,109.687698;"
    AddressCity = AddressCity & "21,249,��ɳ,19.224585,109.452606;"
    AddressCity = AddressCity & "21,250,����,19.260967,109.053352;"
    AddressCity = AddressCity & "21,251,�ֶ�,18.74758,109.175446;"
    AddressCity = AddressCity & "21,252,��ˮ,18.505007,110.037216;"
    AddressCity = AddressCity & "21,253,��ͤ,18.636372,109.702454;"
    AddressCity = AddressCity & "21,254,����,19.03557,109.839996;"
    AddressCity = AddressCity & "22,255,����,29.533155,106.504959;"
    AddressCity = AddressCity & "23,256,�ɶ�,30.659462,104.065735;"
    AddressCity = AddressCity & "23,257,�Թ�,29.352764,104.773445;"
    AddressCity = AddressCity & "23,258,��֦��,26.580446,101.716003;"
    AddressCity = AddressCity & "23,259,����,28.889137,105.443352;"
    AddressCity = AddressCity & "23,260,����,31.127991,104.398651;"
    AddressCity = AddressCity & "23,261,����,31.46402,104.741722;"
    AddressCity = AddressCity & "23,262,��Ԫ,32.433666,105.829758;"
    AddressCity = AddressCity & "23,263,����,30.513311,105.571327;"
    AddressCity = AddressCity & "23,264,�ڽ�,29.58708,105.066139;"
    AddressCity = AddressCity & "23,265,��ɽ,29.582024,103.761261;"
    AddressCity = AddressCity & "23,266,�ϳ�,30.79528,106.082977;"
    AddressCity = AddressCity & "23,267,üɽ,30.048319,103.831787;"
    AddressCity = AddressCity & "23,268,�˱�,28.760189,104.630821;"
    AddressCity = AddressCity & "23,269,�㰲,30.456398,106.633369;"
    AddressCity = AddressCity & "23,270,����,31.209484,107.502258;"
    AddressCity = AddressCity & "23,271,�Ű�,29.987722,103.00103;"
    AddressCity = AddressCity & "23,272,����,31.858809,106.75367;"
    AddressCity = AddressCity & "23,273,����,30.122211,104.641914;"
    AddressCity = AddressCity & "23,274,����,31.899792,102.221375;"
    AddressCity = AddressCity & "23,275,����,30.050663,101.963814;"
    AddressCity = AddressCity & "23,276,��ɽ,27.886763,102.258743;"
    AddressCity = AddressCity & "24,277,����,26.578342,106.713478;"
    AddressCity = AddressCity & "24,278,����ˮ,26.584642,104.846741;"
    AddressCity = AddressCity & "24,279,����,27.706627,106.937263;"
    AddressCity = AddressCity & "24,280,��˳,26.245544,105.93219;"
    AddressCity = AddressCity & "24,281,ͭ��,27.718346,109.191551;"
    AddressCity = AddressCity & "24,282,ǭ����,25.08812,104.897972;"
    AddressCity = AddressCity & "24,283,�Ͻ�,27.301693,105.285011;"
    AddressCity = AddressCity & "24,284,ǭ����,26.583351,107.977486;"
    AddressCity = AddressCity & "24,285,ǭ��,26.258219,107.517159;"
    AddressCity = AddressCity & "25,286,����,25.040609,102.71225;"
    AddressCity = AddressCity & "25,287,����,25.501556,103.797852;"
    AddressCity = AddressCity & "25,288,��Ϫ,24.35046,102.543907;"
    AddressCity = AddressCity & "25,289,��ɽ,25.111801,99.16713;"
    AddressCity = AddressCity & "25,290,��ͨ,27.337,103.717216;"
    AddressCity = AddressCity & "25,291,����,26.872108,100.233025;"
    AddressCity = AddressCity & "25,292,�ն�,22.777321,100.972343;"
    AddressCity = AddressCity & "25,293,�ٲ�,23.886566,100.086967;"
    AddressCity = AddressCity & "25,294,����,25.041988,101.546043;"
    AddressCity = AddressCity & "25,295,���,23.366776,103.384186;"
    AddressCity = AddressCity & "25,296,��ɽ,23.369511,104.244011;"
    AddressCity = AddressCity & "25,297,��˫����,22.001724,100.797943;"
    AddressCity = AddressCity & "25,298,����,25.589449,100.22567;"
    AddressCity = AddressCity & "25,299,�º�,24.436693,98.578362;"
    AddressCity = AddressCity & "25,300,ŭ��,25.850948,98.854301;"
    AddressCity = AddressCity & "25,301,����,27.826853,99.706467;"
    AddressCity = AddressCity & "26,302,����,29.66036,91.13221;"
    AddressCity = AddressCity & "26,303,����,31.136875,97.178452;"
    AddressCity = AddressCity & "26,304,ɽ��,29.236023,91.766525;"
    AddressCity = AddressCity & "26,305,�տ���,29.267519,88.885147;"
    AddressCity = AddressCity & "26,306,����,31.476004,92.060211;"
    AddressCity = AddressCity & "26,307,����,32.503185,80.105499;"
    AddressCity = AddressCity & "26,308,��֥,29.654694,94.36235;"
    AddressCity = AddressCity & "27,309,����,34.263161,108.948021;"
    AddressCity = AddressCity & "27,310,ͭ��,34.91658,108.979607;"
    AddressCity = AddressCity & "27,311,����,34.369316,107.144867;"
    AddressCity = AddressCity & "27,312,����,34.333439,108.705116;"
    AddressCity = AddressCity & "27,313,μ��,34.499382,109.502884;"
    AddressCity = AddressCity & "27,314,�Ӱ�,36.596539,109.490807;"
    AddressCity = AddressCity & "27,315,����,33.077667,107.028618;"
    AddressCity = AddressCity & "27,316,����,38.290161,109.741196;"
    AddressCity = AddressCity & "27,317,����,32.6903,109.029274;"
    AddressCity = AddressCity & "27,318,����,33.86832,109.939774;"
    AddressCity = AddressCity & "28,319,����,36.058041,103.823555;"
    AddressCity = AddressCity & "28,320,������,39.78653,98.277306;"
    AddressCity = AddressCity & "28,321,���,38.514236,102.187889;"
    AddressCity = AddressCity & "28,322,����,36.545681,104.173607;"
    AddressCity = AddressCity & "28,323,��ˮ,34.578529,105.724998;"
    AddressCity = AddressCity & "28,324,����,37.929996,102.634697;"
    AddressCity = AddressCity & "28,325,��Ҵ,38.932896,100.455475;"
    AddressCity = AddressCity & "28,326,ƽ��,35.542789,106.684692;"
    AddressCity = AddressCity & "28,327,��Ȫ,39.744022,98.510796;"
    AddressCity = AddressCity & "28,328,����,35.734219,107.638374;"
    AddressCity = AddressCity & "28,329,����,35.579578,104.626297;"
    AddressCity = AddressCity & "28,330,¤��,33.388599,104.929382;"
    AddressCity = AddressCity & "28,331,����,35.599445,103.212006;"
    AddressCity = AddressCity & "28,332,����,34.986355,102.911011;"
    AddressCity = AddressCity & "29,333,����,36.623177,101.778915;"
    AddressCity = AddressCity & "29,334,����,36.502914,102.103271;"
    AddressCity = AddressCity & "29,335,����,36.959435,100.901062;"
    AddressCity = AddressCity & "29,336,����,35.517742,102.019989;"
    AddressCity = AddressCity & "29,337,���ϲ���,36.280354,100.619545;"
    AddressCity = AddressCity & "29,338,����,34.473598,100.242142;"
    AddressCity = AddressCity & "29,339,����,33.004047,97.008522;"
    AddressCity = AddressCity & "29,340,����,37.374664,97.370789;"
    AddressCity = AddressCity & "30,341,����,38.46637,106.278175;"
    AddressCity = AddressCity & "30,342,ʯ��ɽ,39.013329,106.376175;"
    AddressCity = AddressCity & "30,343,����,37.986164,106.199409;"
    AddressCity = AddressCity & "30,344,��ԭ,36.004562,106.28524;"
    AddressCity = AddressCity & "30,345,����,37.51495,105.189568;"
    AddressCity = AddressCity & "31,346,��³ľ��,43.792816,87.617729;"
    AddressCity = AddressCity & "31,347,��������,45.595886,84.873947;"
    AddressCity = AddressCity & "31,348,��³��,42.947613,89.184074;"
    AddressCity = AddressCity & "31,349,����,42.833248,93.513161;"
    AddressCity = AddressCity & "31,350,����,44.014576,87.304008;"
    AddressCity = AddressCity & "31,351,��������,44.903259,82.074776;"
    AddressCity = AddressCity & "31,352,��������,41.768551,86.15097;"
    AddressCity = AddressCity & "31,353,������,41.170712,80.265068;"
    AddressCity = AddressCity & "31,354,�������տ¶�����,39.713432,76.172829;"
    AddressCity = AddressCity & "31,355,��ʲ,39.467663,75.989136;"
    AddressCity = AddressCity & "31,356,����,37.110687,79.925331;"
    AddressCity = AddressCity & "31,357,����,43.92186,81.317947;"
    AddressCity = AddressCity & "31,358,����,46.7463,82.985733;"
    AddressCity = AddressCity & "31,359,����̩,47.848392,88.139633;"
    AddressCity = AddressCity & "31,360,ʯ����,44.305885,86.041077;"
    AddressCity = AddressCity & "31,361,������,40.541916,81.285881;"
    AddressCity = AddressCity & "31,362,ͼľ���,39.867317,79.07798;"
    AddressCity = AddressCity & "31,363,�����,44.1674,87.526886;"
    AddressCity = AddressCity & "32,364,̨��,25.041618,121.501618;"
    AddressCity = AddressCity & "32,365,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,366,̨��,25.041618,121.501618;"
    AddressCity = AddressCity & "32,367,̨��,25.041618,121.501618;"
    AddressCity = AddressCity & "32,368,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,369,��Ͷ,25.041618,121.501618;"
    AddressCity = AddressCity & "32,370,��¡,25.041618,121.501618;"
    AddressCity = AddressCity & "32,371,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,372,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,373,�±�,25.041618,121.501618;"
    AddressCity = AddressCity & "32,374,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,375,��԰,25.041618,121.501618;"
    AddressCity = AddressCity & "32,376,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,377,�û�,25.041618,121.501618;"
    AddressCity = AddressCity & "32,378,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,379,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,380,̨��,25.041618,121.501618;"
    AddressCity = AddressCity & "32,381,����,25.041618,121.501618;"
    AddressCity = AddressCity & "32,382,���,25.041618,121.501618;"
    AddressCity = AddressCity & "32,383,����,25.041618,121.501618;"
    AddressCity = AddressCity & "33,384,��۵�,22.320047,114.173355;"
    AddressCity = AddressCity & "33,385,����,22.320047,114.173355;"
    AddressCity = AddressCity & "33,386,�½�,22.320047,114.173355;"
    AddressCity = AddressCity & "34,387,���Ű뵺,22.198751,113.549133;"
    AddressCity = AddressCity & "34,398,�뵺,22.198952,113.549088"

End Function



















