Attribute VB_Name = "Power-BI-custom-sample-data"
Option Compare Database
Option Explicit

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'1�����ߣ�������
'2�����䣺jiaopengzi@qq.com
'3�����ͣ�www.jiaopengzi.com
'4��CPU��12th Gen Intel(R) Core(TM) i9-12900KF   3.20 GHz
'5���ڴ棺RAM 32.0 GB
'6�����ϵ������� + ShopQuantity=300 �����ã���Լ��Ҫ 1000 �룬ÿ�밴��ҵ���߼�����Լ 1����+ ���ݣ����� 1000 ����+ demo���ݣ���������ʵսѧϰ���á�
'   ���ϵ������� + ShopQuantity=100 �����ã���Լ��Ҫ  350 �룬ÿ�밴��ҵ���߼�����Լ 1����+ ���ݣ�����  360 ����+ demo���ݣ���������ʵսѧϰ���á�
'   ���ϵ������� + ShopQuantity=10  �����ã���Լ��Ҫ   60 �룬ÿ�밴��ҵ���߼�����Լ 1����+ ���ݣ�����   60 ����+ demo���ݣ���������ʵսѧϰ���á�
'   ���ϵ������� + ShopQuantity=5   �����ã���Լ��Ҫ   20 �룬ÿ�밴��ҵ���߼�����Լ 1����+ ���ݣ�����   20 ����+ demo���ݣ���������ʵսѧϰ���á�

'=====================================================================================

Public productQuantity As Long   '��Ʒ����������ShopQuantity��[7,999]��
Public ShopQuantity As Long   '�ŵ�����������ShopQuantity��[1,390]��
Public MaxInventoryDays As Long   '����������������ShopQuantity��[5,20]��

'=====================================================================================�����ƹ���
Public Const tbNameOrg As String = "D10_��֯��"                         ' D10
Public Const tbNameRegion As String = "D20_������"                      ' D20
Public Const tbNameProvince As String = "D21_ʡ�ݱ�"                    ' D21
Public Const tbNameCity As String = "D22_���б�"                        ' D22
Public Const tbNameDistrict As String = "D23_���ر�"                    ' D23
Public Const tbNameProduct As String = "D30_��Ʒ��"                     ' D30
Public Const tbNameShop As String = "T10_�ŵ��"                        ' T10
Public Const tbNameShopRental As String = "T11_�ŵ��_����"             ' T11
Public Const tbNameShopDecoration As String = "T12_�ŵ��_װ��"         ' T12
Public Const tbNameCustomer As String = "T20_�ͻ���"                    ' T20
Public Const tbNameStorage As String = "T30_�����Ϣ��"                 ' T30
Public Const tbNameOrder As String = "T40_��������"                     ' T40
Public Const tbNameOrdersub As String = "T41_�����ӱ�"                  ' T41
Public Const tbNameSaleTarget As String = "T50_����Ŀ���"              ' T50
Public Const tbNameEmployee As String = "T60_Ա����Ϣ��"                ' T60
Public Const tbNameLaborCost As String = "T61_�˹��ɱ���"               ' T61

'=====================================================================================����ֵ����ƺ�����SQL
Public Const fLaborCostOrgID As String = "��֯ID"
Public Const fLaborCostMonth As String = "�·�"
Public Const fLaborCostAmount As String = "�˹��ɱ����_Ԫ"
Public Const createTbSqlLaborCost As String = "CREATE TABLE " & tbNameLaborCost & _
            "(                                                                  " & vbCrLf & _
            fLaborCostOrgID & "         INT,                                    " & vbCrLf & _
            fLaborCostMonth & "         DATE,                                   " & vbCrLf & _
            fLaborCostAmount & "        INT                                     " & vbCrLf & _
            ")"

Public Const fProductID As String = "��ƷID"
Public Const fProductCategory As String = "��Ʒ����"
Public Const fProductName As String = "��Ʒ����"
Public Const fProductPrice As String = "��Ʒ���ۼ۸�"
Public Const fProductCostPrice = "��Ʒ�ɱ��۸�"
Public Const createTbSqlProduct As String = "CREATE TABLE " & tbNameProduct & _
            "(                                                                  " & vbCrLf & _
            fProductID & "             VARCHAR(50) PRIMARY KEY,                 " & vbCrLf & _
            fProductCategory & "       VARCHAR(50),                             " & vbCrLf & _
            fProductName & "           VARCHAR(50),                             " & vbCrLf & _
            fProductPrice & "          INT,                                     " & vbCrLf & _
            fProductCostPrice & "      INT                                      " & vbCrLf & _
            ")"
            
'�ŵ�ID�ڴ����ǲ�����ID����,��ΪID���պ��ô���
Public Const fShopID As String = "�ŵ���֯ID"
Public Const fShopName As String = "�ŵ�����"
Public Const fShopOpenDate As String = "��ҵ����"
Public Const fShopDistrictID As String = "����ID"
Public Const fShopDistrict As String = "����"
Public Const fShopLongitude As String = "γ��"
Public Const fShopLatitude As String = "����"
Public Const fShopCloseDate As String = "�յ�����"
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
            
Public Const fShopRentalShopID As String = "�ŵ���֯ID"
Public Const fShopRentalArea As String = "�������_ƽ����"
Public Const fShopRentalPrice As String = "�������_Ԫÿ��ÿƽ����"
Public Const fShopRentalStartDate As String = "��������"
Public Const fShopRentalEndDate As String = "ֹ������"
Public Const fShopRentalIncrease As String = "�������Ƿ�"
Public Const createTbSqlShopRental As String = "CREATE TABLE " & tbNameShopRental & _
            "(                                                                  " & vbCrLf & _
            fShopRentalShopID & "      INT,                                     " & vbCrLf & _
            fShopRentalArea & "        FLOAT,                                   " & vbCrLf & _
            fShopRentalPrice & "       FLOAT,                                   " & vbCrLf & _
            fShopRentalStartDate & "   DATE,                                    " & vbCrLf & _
            fShopRentalEndDate & "     DATE,                                    " & vbCrLf & _
            fShopRentalIncrease & "    FLOAT                                    " & vbCrLf & _
            ")"
            
Public Const fShopDecorationShopID As String = "�ŵ���֯ID"
Public Const fShopDecorationStartDate As String = "װ�޿�ʼ����"
Public Const fShopDecorationEndDate As String = "װ�޽�������"
Public Const fShopDecorationAmount As String = "װ�޽��_Ԫ"
Public Const fShopDecorationYears As String = "װ���۾�����"
Public Const createTbSqlShopDecoration As String = "CREATE TABLE " & tbNameShopDecoration & _
            "(                                                                  " & vbCrLf & _
            fShopDecorationShopID & "  INT,                                     " & vbCrLf & _
            fShopDecorationStartDate & " DATE,                                  " & vbCrLf & _
            fShopDecorationEndDate & " DATE,                                    " & vbCrLf & _
            fShopDecorationAmount & "  FLOAT,                                   " & vbCrLf & _
            fShopDecorationYears & "   FLOAT                                    " & vbCrLf & _
            ")"
            
Public Const fCustomerID As String = "�ͻ�ID"
Public Const fCustomerName As String = "�ͻ�����"
Public Const fCustomerBirthday As String = "�ͻ�����"
Public Const fCustomerGender As String = "�ͻ��Ա�"
Public Const fCustomerRegister As String = "ע������"
Public Const fCustomerIndustry As String = "�ͻ���ҵ"
Public Const fCustomerOccupation As String = "�ͻ�ְҵ"
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

Public Const fStorageProductID As String = "����ƷID"
Public Const fStorageQuantity As String = "����Ʒ����"
Public Const fStorageShopID As String = "����ŵ���֯ID"
Public Const fStorageDate As String = "�������"
Public Const createTbSqlStorage As String = "CREATE TABLE " & tbNameStorage & _
            "(                                                                  " & vbCrLf & _
            fStorageProductID & "      VARCHAR(50),                             " & vbCrLf & _
            fStorageQuantity & "       INT,                                     " & vbCrLf & _
            fStorageShopID & "         INT,                                     " & vbCrLf & _
            fStorageDate & "           DATE                                     " & vbCrLf & _
            ")"
         
Public Const fOrderID As String = "����ID"
Public Const fOrderShopID As String = "�ŵ���֯ID"
Public Const fOrderDate As String = "�µ�����"
Public Const fOrderSentDate As String = "�ͻ�����"
Public Const fOrderCustomerID As String = "�ͻ�ID"
Public Const fOrderType As String = "��������"
Public Const fOrderEmployeeID As String = "����Ա��ID"
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

Public Const fOrdersubOrderID As String = "����ID"
Public Const fOrdersubProductID As String = "��ƷID"
Public Const fOrdersubPrice As String = "��Ʒ���ۼ۸�"
Public Const fOrdersubDiscount As String = "�ۿ۱���"
Public Const fOrdersubQuantity As String = "��Ʒ��������"
Public Const fOrdersubAmount As String = "��Ʒ���۽��"
Public Const createTbSqlOrdersub As String = "CREATE TABLE " & tbNameOrdersub & _
            "(                                                                  " & vbCrLf & _
            fOrdersubOrderID & "       VARCHAR(50),                             " & vbCrLf & _
            fOrdersubProductID & "     VARCHAR(50),                             " & vbCrLf & _
            fOrdersubPrice & "         INT,                                     " & vbCrLf & _
            fOrdersubDiscount & "      FLOAT,                                   " & vbCrLf & _
            fOrdersubQuantity & "      INT,                                     " & vbCrLf & _
            fOrdersubAmount & "        FLOAT,                                   " & vbCrLf & _
            "CONSTRAINT PK_" & tbNameOrdersub & " PRIMARY KEY (����ID, ��ƷID)  " & vbCrLf & _
            ")"
            
Public Const fSaleTargetProvinceID As String = "ʡID"
Public Const fSaleTargetProvinceName2 As String = "ʡ���"
Public Const fSaleTargetMonth As String = "�·�"
Public Const fSaleTargetAmount As String = "����Ŀ��"
Public Const createTbSqlSaleTarget As String = "CREATE TABLE " & tbNameSaleTarget & _
            "(                                                                  " & vbCrLf & _
            fSaleTargetProvinceID & "  INT,                                     " & vbCrLf & _
            fSaleTargetProvinceName2 & " VARCHAR(50),                           " & vbCrLf & _
            fSaleTargetMonth & "       DATE,                                    " & vbCrLf & _
            fSaleTargetAmount & "      FLOAT                                    " & vbCrLf & _
            ")"

Public Const fRegionID As String = "������֯ID"
Public Const fRegionName As String = "���"
Public Const fRegionCityID As String = "�칫�س���ID"
Public Const fRegionCity As String = "�칫�س���"
Public Const fRegionLongitude As String = "γ��"
Public Const fRegionLatitude As String = "����"
Public Const createTbSqlRegion As String = "CREATE TABLE " & tbNameRegion & _
            "(                                                                  " & vbCrLf & _
            fRegionID & "              INT PRIMARY KEY,                         " & vbCrLf & _
            fRegionName & "            VARCHAR(50),                             " & vbCrLf & _
            fRegionCityID & "          INT,                                     " & vbCrLf & _
            fRegionCity & "            VARCHAR(50),                             " & vbCrLf & _
            fRegionLongitude & "       FLOAT,                                   " & vbCrLf & _
            fRegionLatitude & "        FLOAT                                    " & vbCrLf & _
            ")"

Public Const fProvinceRegionID As String = "������֯ID"
Public Const fProvinceID As String = "ʡID"
Public Const fProvinceNameAll As String = "ʡȫ��"
Public Const fProvinceName1 As String = "ʡ���1"
Public Const fProvinceName2 As String = "ʡ���2"
Public Const fProvinceLongitude As String = "γ��"
Public Const fProvinceLatitude As String = "����"
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

Public Const fCityProvinceID As String = "ʡID"
Public Const fCityID As String = "����ID"
Public Const fCityName As String = "����"
Public Const fCityLongitude As String = "γ��"
Public Const fCityLatitude As String = "����"
Public Const createTbSqlCity As String = "CREATE TABLE " & tbNameCity & _
            "(                                                                  " & vbCrLf & _
            fCityProvinceID & "        INT,                                     " & vbCrLf & _
            fCityID & "                INT PRIMARY KEY,                         " & vbCrLf & _
            fCityName & "              VARCHAR(50),                             " & vbCrLf & _
            fCityLongitude & "         FLOAT,                                   " & vbCrLf & _
            fCityLatitude & "          FLOAT                                    " & vbCrLf & _
            ")"

Public Const fDistrictCityID As String = "����ID"
Public Const fDistrictID As String = "����ID"
Public Const fDistrictName As String = "����"
Public Const fDistrictLongitude As String = "γ��"
Public Const fDistrictLatitude As String = "����"
Public Const createTbSqlDistrict As String = "CREATE TABLE " & tbNameDistrict & _
            "(                                                                  " & vbCrLf & _
            fDistrictCityID & "        INT,                                     " & vbCrLf & _
            fDistrictID & "            INT PRIMARY KEY,                         " & vbCrLf & _
            fDistrictName & "          VARCHAR(50),                             " & vbCrLf & _
            fDistrictLongitude & "     FLOAT,                                   " & vbCrLf & _
            fDistrictLatitude & "      FLOAT                                    " & vbCrLf & _
            ")"

Public Const fOrgID As String = "��֯ID"
Public Const fOrgNameAll As String = "��֯����"
Public Const fOrgParentID As String = "�ϼ���֯ID"
Public Const fOrgName As String = "��֯���"
Public Const fOrgEmployeeID As String = "������ID"
Public Const createTbSqlOrg As String = "CREATE TABLE " & tbNameOrg & _
            "(                                                                  " & vbCrLf & _
            fOrgID & "                 INT IDENTITY(1,1) PRIMARY KEY,           " & vbCrLf & _
            fOrgNameAll & "            VARCHAR(255),                            " & vbCrLf & _
            fOrgParentID & "           INT,                                     " & vbCrLf & _
            fOrgName & "               VARCHAR(255),                            " & vbCrLf & _
            fOrgEmployeeID & "         INT                                      " & vbCrLf & _
            ")"
            
Public Const fEmployeeID As String = "Ա��ID"
Public Const fEmployeeName As String = "����"
Public Const fEmployeeGender As String = "�Ա�"
Public Const fEmployeeOrgID As String = "��֯ID"
Public Const fEmployeeJobTitle As String = "ְ��"
Public Const fEmployeeGrade As String = "ְ��"
Public Const fEmployeeEdu As String = "ѧ��"
Public Const fEmployeeBirthday As String = "��������"
Public Const fEmployeeEntryDate As String = "��ְ����"
Public Const fEmployeeResignationDate As String = "��ְ����"
Public Const fEmployeeResignationReason As String = "��ְԭ��"
Public Const createTbSqlEmployee As String = "CREATE TABLE " & tbNameEmployee & _
            "(                                                                  " & vbCrLf & _
            fEmployeeID & "            INT IDENTITY(10001,1) PRIMARY KEY,       " & vbCrLf & _
            fEmployeeName & "          VARCHAR(50),                             " & vbCrLf & _
            fEmployeeGender & "        VARCHAR(20) DEFAULT ��,                  " & vbCrLf & _
            fEmployeeOrgID & "         INT,                                     " & vbCrLf & _
            fEmployeeJobTitle & "      VARCHAR(50),                             " & vbCrLf & _
            fEmployeeGrade & "         VARCHAR(50),                             " & vbCrLf & _
            fEmployeeEdu & "           VARCHAR(50),                             " & vbCrLf & _
            fEmployeeBirthday & "      DATE,                                    " & vbCrLf & _
            fEmployeeEntryDate & "     DATE,                                    " & vbCrLf & _
            fEmployeeResignationDate & " DATE NULL,                             " & vbCrLf & _
            fEmployeeResignationReason & " VARCHAR(255) NULL                    " & vbCrLf & _
            ")"
            

'=====================================================================================ȫ�ֱ���
Public TableNameDict As Object ' �������ֵ�
Public MinDateOpen As Date ' ���翪ҵ����
Public ProvinceID2OrgIDDict As Object ' ʡ������ID��ǰ��λ����֯ID��ӳ���ֵ�
Public JobTitlesArr As Variant 'ְ��
Public GradeArr As Variant 'ְ��
Public EduArr As Variant 'ѧ����
Public EduDict As Object
Public EduSalaryDict As Object
Public GradeDict As Object
Public GradeSalaryDict As Object
Public ResignationArr As Variant '��ְԭ��


Public Function InitE()
    '��ʼ��Ա����Ϣ�������
    JobTitlesArr = Array("�ܾ���", "�ܾ�������", "��Ʒ�ܼ�", "�ɹ��ܼ�", "�����ܼ�", "�����ܼ�", "������Դ�ܼ�", "�ۺ�����ܼ�", "�����ܼ�", "��������", "ʡ������", "�ŵ꾭��", "���۹���", "�ۺ�רԱ")
    
    GradeArr = Array("�ܾ���", "�߼��ܼ�", "�ܼ�", "�߼�����", "����", "����", "רԱ")
        
    EduArr = Array("�о���", "����", "ר��", "����")
    
    ResignationArr = Array("���˷�չ", "����ԭ��", "����ǿ��", "���������뻷��", "��ͥԭ��", "����ԭ��", "Υ�������ƶ�", "Ȱ��", "����", "����ԭ��", "�������ڽ��") '�����ڷ�������10
    Set EduDict = CreateObject("Scripting.Dictionary")
    With EduDict
        .Add "PD", "��ʿ" '��ʿ��Doctorate (PD)
        .Add "PG", "˶ʿ" '�о���: Postgraduate (PG)
        .Add "UG", "����" '����: Undergraduate (UG)
        .Add "AD", "ר��" 'ר��: Associate Degree(AD)
        .Add "HS", "����" '����: High School(HS)
        .Add "MS", "����" '���У�Junior High School (JHS) �� Middle School (MS)
        .Add "PS", "Сѧ" 'Сѧ��Primary School (PS) �� Elementary School (ES)
    End With
    
    Set EduSalaryDict = CreateObject("Scripting.Dictionary")
    With EduSalaryDict
        .Add "��ʿ", 2 'н��ϵ��
        .Add "˶ʿ", 1.2
        .Add "����", 1.1
        .Add "ר��", 1
        .Add "����", 0.9
        .Add "����", 0.8
        .Add "Сѧ", 0.7
    End With
    
    Set GradeDict = CreateObject("Scripting.Dictionary")
    With GradeDict
        .Add "GM", "�ܾ���"     'General Manager (GM)
        .Add "SD", "�߼��ܼ�"   'Senior Director (SD)
        .Add "D", "�ܼ�"        'Director (D)
        .Add "SM", "�߼�����"   'Senior Manager (SM)
        .Add "M", "����"        'Manager (M)
        .Add "S", "����"        'Supervisor (S)
        .Add "SP", "רԱ"       'Specialist (SP)
    End With
    
    Set GradeSalaryDict = CreateObject("Scripting.Dictionary")
    With GradeSalaryDict
        .Add "�ܾ���", Array(50000, 100000) 'н�ʷ�Χ
        .Add "�߼��ܼ�", Array(30000, 50000)
        .Add "�ܼ�", Array(20000, 30000)
        .Add "�߼�����", Array(12000, 20000)
        .Add "����", Array(8000, 12000)
        .Add "����", Array(5000, 8000)
        .Add "רԱ", Array(3000, 5000)
    End With
End Function


Public Function InitPO()
    ' ��ʼ�� ʡ������ID��ǰ��λ����֯ID��ӳ���ֵ�
    
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
        ProvinceID2OrgIDDict.Add CInt(Left(Trim(ArrAddProvinceRow(i)(1)), 2)), i + 15 '15���� D20�����Ĭ��ֵȷ����
    Next

     
End Function


Public Function InitTables()
    ' ��ʼ���������ֵ�
    Set TableNameDict = CreateObject("Scripting.Dictionary")
    
    ' ����������Ϊ������Ӧ�ı�Ĵ���sql�����Ϊֵ��ӵ��ֵ���
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
' ���ݱ�����ɾ����
    SQLDrop = "DROP TABLE " & tableName

End Function

Public Function TableADO(tableName As String, Sql_Drop As String, Sql_Create As String)
    ' ���ɱ�
    Dim Cat As Object
    Dim cmd As Object
    
'    On Error GoTo ErrorHandler
    
    Set Cat = CreateObject("ADOX.Catalog")
    Set cmd = CreateObject("ADODB.Command")
    
    Set Cat.ActiveConnection = CurrentProject.Connection
    Set cmd.ActiveConnection = CurrentProject.Connection
    
    With cmd
        .CommandTimeout = 100
        
        ' ɾ���Ѵ��ڵı�
        If TableExists(tableName, Cat.tables) Then
            .CommandText = Sql_Drop
            .Execute
        End If
        
        ' �����±�
        .CommandText = Sql_Create
        .Execute
    End With

CleanUp:
    Set cmd = Nothing
    Set Cat = Nothing
    Exit Function

'ErrorHandler:
'    ' ��������룬���Ը�����Ҫ������Ӧ�Ĵ���
'    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
'    Resume CleanUp
End Function

Public Function TableExists(tableName As String, tables As Object) As Boolean
    ' �����Ƿ����
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
'����Ĭ����Ϣ
    Dim ArrAdd
    Dim ArrAddRow
    Dim Region As String
    Dim i As Long
    Region = "9, ����,   310000, �Ϻ�,   31.231518,  121.471518;" & _
            "10, ����,   510100, �ɶ�,   30.659518,  104.065518;" & _
            "11, ����,   440100, ����,   23.125518,  113.280518;" & _
            "12, ����,   210100, ����,   41.796518,  123.429518;" & _
            "13, ����,   110000, ����,   39.901518,  116.401518;" & _
            "14, �۰�̨, 810000, ���,   22.320518,  114.173518"
    
    ArrAdd = Split(Region, ";")

    ReDim ArrAddRow(0 To UBound(ArrAdd))

    For i = 0 To UBound(ArrAdd)
        ArrAddRow(i) = Split(ArrAdd(i), ",")
    Next
    ArrAddRegionDefault = ArrAddRow
End Function

Public Function DataTableRegion()
' ����ҵ���߼����� ������
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
' ����ҵ���߼����� ʡ�ݱ�
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
' ����ҵ���߼����� ���б�
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
' ����ҵ���߼����� ���ر�

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
    '����Ĭ����Ϣ
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
' ����ҵ���߼����� ��֯��
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
    
    Set dictGender = GenderDict() 'ǰ�� 448 �����Ը���ͷ������
    Set dictAllDate = DateStatusDict() '��������״̬���ֵ�
    
    InitE
    
    Set conn = CreateConnection
    Set RsOrg = CreateRecordset(conn, tbNameOrg)
    Set RsShop = CreateRecordset(conn, tbNameShop)
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)

'=====================================================================================
'һ������ �� ���۴���
    Const org As String = "�����ӿƼ����޹�˾,   ,      �ܲ�,       10001;" & _
                          "�ܾ���칫��,        1,      �ܾ���,     10002;" & _
                          "��Ʒ�з�����,        1,      ��Ʒ,       10003;" & _
                          "�ɹ�����,            1,      �ɹ�,       10004;" & _
                          "��������,            1,      ����,       10005;" & _
                          "������Դ����,        1,      ����,       10006;" & _
                          "�ۺ��������,        1,      �ۺ�,       10007;" & _
                          "��������,            1,      ����,       10008;" & _
                          "�������۴���,        5,      ����,       10009;" & _
                          "�������۴���,        5,      ����,       10010;" & _
                          "�ϲ����۴���,        5,      ����,       10011;" & _
                          "�������۴���,        5,      ����,       10012;" & _
                          "�в����۴���,        5,      ����,       10013;" & _
                          "�۰�̨���۴���,      5,      �۰�̨,     10014"

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
'ʡ����������
    
    ArrAddOrg = Split(AddressProvince, ";")

    ReDim ArrAddOrgRow(0 To UBound(ArrAddOrg))

    For i = 0 To UBound(ArrAddOrg)
        ArrAddOrgRow(i) = Split(ArrAddOrg(i), ",")
    Next

    rows = UBound(ArrAddOrgRow)
    
    For i = 0 To rows
        '===============================Ա����Ϣ
        myRnd = Rnd()
        employeeName = generateName(myRnd)
        If myRnd < 0.7 Then employeeGender = "Ů" Else employeeGender = "��"
        employeeJobTitle = JobTitlesArr(10)
        If myRnd < 0.8 Then
            employeeGrade = GradeArr(3)
        Else
            employeeGrade = GradeArr(4)
        End If
        employeeEdu = EduArr(Round(Rnd() * 2, 0))
        employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 7500, 0)
        employeeEntryDate = MinDateOpen - Round(Rnd() * 50, 0)
        
        AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender '��֯ID����

        '===============================��֯
        RsOrg.AddNew
            RsOrg.Fields(fOrgNameAll) = "ʡ����������" + Trim(ArrAddOrgRow(i)(4))
            RsOrg.Fields(fOrgParentID) = Trim(ArrAddOrgRow(i)(0))
            RsOrg.Fields(fOrgName) = Trim(ArrAddOrgRow(i)(4))
        RsOrg.Update
        
        '===============================����ID
        RsOrg.Fields(fOrgEmployeeID) = RsEmployee.Fields(fEmployeeID)
        RsOrg.Update
        RsEmployee.Fields(fEmployeeOrgID) = RsOrg.Fields(fOrgID)
        RsEmployee.Update
        maxOrgID = RsOrg.Fields(fOrgID)
                
    Next
'=====================================================================================
'�ŵ���֯
    InitPO '��ʼ��
    
    RsShop.MoveFirst
    Do Until RsShop.EOF
        '��ֵ�ŵ����֯ID
        maxOrgID = maxOrgID + 1
        RsShop.Fields(fShopID) = maxOrgID
        dateOpen = RsShop.Fields(fShopOpenDate)
        RsShop.Update
        
        '��֯����
        RsOrg.AddNew
            RsOrg.Fields(fOrgNameAll) = "�����ŵ�-" & RsShop.Fields(fShopName)
            RsOrg.Fields(fOrgParentID) = ProvinceID2OrgIDDict(CInt(Left(RsShop.Fields(fShopDistrictID), 2))) 'ͨ���ŵ� ����ID ��ǰ��λ��ȡ �ϼ���֯ID
            RsOrg.Fields(fOrgName) = RsShop.Fields(fShopName)
        RsOrg.Update
        
        '�ŵ긺��������
        myRnd = Rnd()
        employeeName = generateName(myRnd)
        If myRnd < 0.7 Then employeeGender = "Ů" Else employeeGender = "��"
        employeeJobTitle = JobTitlesArr(11)
        employeeGrade = GradeArr(4)
        employeeEdu = EduArr(1 + Round(Rnd() * 2, 0)) 'ѧ��Ҫ�󽵵�
        employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 6000, 0) '�����ữ
        employeeEntryDate = dateOpen - Round(Rnd() * 30, 0)
        
        AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender '��֯ID����
   
        '===============================����ID
        RsOrg.Fields(fOrgEmployeeID) = RsEmployee.Fields(fEmployeeID)
        RsOrg.Update
        
        RsEmployee.Fields(fEmployeeOrgID) = RsOrg.Fields(fOrgID)
        RsEmployee.Update
        
        RsShop.MoveNext
    Loop
    
    CloseConnRs conn, RsOrg, RsShop, RsEmployee
        
End Function

Public Function DataTableProduct()
' ����ҵ���߼����� ��Ʒ��

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
        
        RsProduct.Fields(fProductCategory) = Chr(Round(myRnd * 9, 0) + 65) & "��"
        
        RsProduct.Fields(fProductName) = "��Ʒ" & Chr(Round(myRnd * 9, 0) + 65) & "" & Format(i, "0000")
 
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
' ����ҵ���߼����� �ŵ��
    Dim i As Long
    Dim k As Long
    Dim myRnd As Double
    Dim ArrAddressDistrict
    Dim ArrAddressDistrictRow
    Dim ArrDictName
    Dim ArrDefault7 'Ĭ���ֶ����� ֱϽ��+�۰�̨��������
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

    
    Set DictName = CreateObject("Scripting.Dictionary") '��������֣��ֵ������֤Ψһ���ظ���
    For i = 1 To 17576 '26*26*26
        DictName(Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & Chr(Round(Rnd() * 25, 0) + 65) & "��") = i
        If DictName.Count = ShopQuantity Then
            Exit For
        End If
    Next
    
    ArrDictName = DictName.Keys
    Set DictName = Nothing
    
    ArrDefault7 = Array( _
                Array(110101, "������", 39.917548, 116.418758), _
                Array(120101, "��ƽ��", 39.118328, 121.490318), _
                Array(310101, "������", 31.222778, 121.471518), _
                Array(500103, "������", 29.556748, 106.562888), _
                Array(710000, "̨��", 25.044518, 121.509518), _
                Array(810001, "������", 22.28198088, 114.1543738), _
                Array(820001, "����������", 22.207878, 113.5528958) _
                )
            
    '���������ĸ�ֱϽ�� + �۰�̨
    If ShopQuantity < 8 Then
        For k = 0 To ShopQuantity - 1
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD")
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen 'ȡ��С������
            
            AddShopRecord RsShop, ArrDictName(k), dateOpen, CLng(ArrDefault7(k)(0)), CStr(ArrDefault7(k)(1)), Round(ArrDefault7(k)(2), 6), Round(ArrDefault7(k)(3), 6)
            
        Next
    End If
    
    '����ǰ���߸����к������ɴ��� 7 �����ݡ�
    If ShopQuantity > 7 Then
    
        For k = 0 To 6
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD")
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen 'ȡ��С������
            AddShopRecord RsShop, ArrDictName(k), dateOpen, CLng(ArrDefault7(k)(0)), CStr(ArrDefault7(k)(1)), Round(ArrDefault7(k)(2), 6), Round(ArrDefault7(k)(3), 6)

        Next
    
        For i = 8 To ShopQuantity
            Randomize
            myRnd = Rnd()
            
            Randomize
            addUB0 = Round(UBound(ArrAddressDistrictRow) * Rnd(), 0)
            Randomize
            dateOpen = Format(Now - Round(Rnd() * 1500 + 28, 0), "YYYY-MM-DD") '+28�ݴ�Dict3N
            If MinDateOpen > dateOpen Then MinDateOpen = dateOpen 'ȡ��С������
            Randomize
            dateClose = Format(dateOpen + 550 + 4320 * Rnd(), "YYYY-MM-DD") '550��ʾ����1.5����ܹص�
    
            If dateClose > Now Then
                AddShopRecord RsShop, ArrDictName(i - 1), dateOpen, CLng(ArrAddressDistrictRow(addUB0)(1)), CStr(ArrAddressDistrictRow(addUB0)(2)), Round(ArrAddressDistrictRow(addUB0)(3) + Rnd() * 0.05, 6), Round(ArrAddressDistrictRow(addUB0)(4) + Rnd() * 0.05, 6)
            Else '�յ�
                AddShopRecord RsShop, ArrDictName(i - 1), dateOpen, CLng(ArrAddressDistrictRow(addUB0)(1)), CStr(ArrAddressDistrictRow(addUB0)(2)), Round(ArrAddressDistrictRow(addUB0)(3) + Rnd() * 0.05, 6), Round(ArrAddressDistrictRow(addUB0)(4) + Rnd() * 0.05, 6), dateClose
            End If
        Next
    End If

    CloseConnRs conn, RsShop
        
End Function

Public Function AddShopRecord(ByRef RsShop As Object, ByVal ShopName As String, ShopOpenDate As Date, ShopDistrictID As Long, ShopDistrict As String, ShopLongitude As Double, ShopLatitude As Double, Optional ByVal ShopCloseDate As Date)
    '��ȡ�ŵ���������
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
    '�������޺�װ������
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
        dateRentalStart = RsShop.Fields(fShopOpenDate) - 30 - Round(Rnd * 15, 0) '�״����޿�ʼ����
        dateRsShopDecorationStart = dateRentalStart + Round(Rnd * 7, 0) '�״�װ�޿�ʼ����Format(Now, "YYYY-MM-DD")
        ShopRentalArea = 600 + Rnd * 600 '�������
        ShopRentalPrice = 40 + Rnd * 40 '�״����޼۸�
        decorationAmount = ShopRentalArea * 1000 * (0.8 + (Rnd() * 0.3)) 'װ�޽��
        depreciationPeriod = 3 + Round(Rnd * 2, 0) '�۾�����
        
        If IsNull(RsShop.Fields(fShopCloseDate)) Then
Rental: '����
            dateRentalEnd = dateRentalStart + 3 * 365 '���޵�������
                AddRsShopRentalRecord RsShopRental, RsShop.Fields(fShopID), Round(ShopRentalArea, 2), Round(ShopRentalPrice, 2), dateRentalStart, dateRentalEnd, Round(Rnd * 0.05, 2)
            If dateRentalEnd < Now Then
                dateRentalStart = dateRentalEnd + 1
                ShopRentalPrice = ShopRentalPrice * 0.9 + (Rnd() * 0.2)
                GoTo Rental
            End If
        Else
            AddRsShopRentalRecord RsShopRental, RsShop.Fields(fShopID), Round(ShopRentalArea, 2), Round(ShopRentalPrice, 2), dateRentalStart, RsShop.Fields(fShopCloseDate), Round(Rnd * 0.05, 2)
        End If
        

        If IsNull(RsShop.Fields(fShopCloseDate)) Then 'δ�ص�
Decoration: 'װ��
            dateRsShopDecorationEnd = Format(dateRsShopDecorationStart + Round(45 * (0.8 + (Rnd() * 0.3)), 0), "YYYY-MM-DD") 'װ�޽�������
            depreciationEndDate = Format(dateRsShopDecorationEnd + depreciationPeriod * 365)  'װ���۾ɽ�������
            AddShopDecorationRecord RsShopDecoration, RsShop.Fields(fShopID), dateRsShopDecorationStart, dateRsShopDecorationEnd, Round(decorationAmount, 2), depreciationPeriod

            If depreciationEndDate < Now Then
                dateRsShopDecorationStart = depreciationEndDate + 1
                decorationAmount = decorationAmount * 0.9 + (Rnd() * 0.2)
                GoTo Decoration
            End If
        Else
            dateRsShopDecorationEnd = Format(dateRsShopDecorationStart + Round(45 * (0.8 + (Rnd() * 0.3)), 0), "YYYY-MM-DD") 'װ�޽�������
            depreciationPeriod = Int((RsShop.Fields(fShopCloseDate) - dateRsShopDecorationEnd) / 365)
            AddShopDecorationRecord RsShopDecoration, RsShop.Fields(fShopID), dateRsShopDecorationStart, dateRsShopDecorationEnd, Round(decorationAmount, 2), depreciationPeriod
        End If
            

    RsShop.MoveNext
    Loop
End Function

Public Function AddRsShopRentalRecord(ByRef RsShopRental As Object, ByVal shopID As Long, Area As Long, price As Double, startDate As Date, endDate As Date, Increase As Double)
    '��ȡ�ŵ�������������
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
    '��ȡ�ŵ�װ����������
    RsShopDecoration.AddNew
        RsShopDecoration.Fields(fShopDecorationShopID) = shopID
        RsShopDecoration.Fields(fShopDecorationStartDate) = startDate
        RsShopDecoration.Fields(fShopDecorationEndDate) = endDate
        RsShopDecoration.Fields(fShopDecorationAmount) = amount
        RsShopDecoration.Fields(fShopDecorationYears) = Years
    RsShopDecoration.Update
End Function

Public Function DataTableCustomer()
' ����ҵ���߼����� �ͻ���
    Dim registerDays As Long
    Dim row As Long
    Dim iCount As Long
    Dim i As Long
    Dim myRnd As Double
    Dim myRndHY As Double
    Dim myRndZY As Double
    Dim customerN As Double
    
    Dim arrHY '��ҵ
    Dim arrZY 'ְҵ
    Dim arrRndNL '����ֲ�
    
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

    arrHY = Array("����ҵ", "����ҵ", "������", "ũҵ", "����", "����", "����") '��ҵ
    arrZY = Array("���廧", "HR", "��Ӫ", "IT", "����", "����", "�з�") 'ְҵ
    arrRndHY = Array(0.2, 0.5, 0.5, 0.8, 0.8, 1, 0.9) '��ҵ�ֲ�
    arrRndZY = Array(0.3, 0.7, 0.6, 1, 1, 0.8, 0.1) 'ְҵ�ֲ�
    arrRndNL = Array(0, 0.1, 0.2, 0.3, 0.3, 0.3, 0.3, 0.8, 0.8, 0.8, 0.9, 1) '����ֲ�


    Set conn = CreateConnection
    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    '���յ��̹�ģע��ͻ�
    
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
        customerN = (3 + 8 * Rnd()) '�ͻ���������ϵ��
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
            dateBirthday = Format(Now - 7500 - Round((arrRndNL(iCount Mod 12) + Rnd()) * 7000, 0), "YYYY-MM-DD")  '����
            dateRegister = Format(Now - 1500 + Round((arrRndNL(iCount Mod 12) + Rnd()) * 750, 0), "YYYY-MM-DD") 'ע��ʱ�䣬�ȿ��������죬����ҵ���߼����
    
            If myRnd < 0.8 Then
                name = generateName(myRnd)
                gender = "��"
            Else
                name = generateName(myRnd)
                gender = "Ů"
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
' ����ҵ���߼����� �����������������ӱ�
    
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
        Dim Yyts As Long 'Ӫҵ����
        Dim dateOpen As Date

        Dim arrProduct
        Dim arrShop
        Dim arrCustomer
        Dim arrHY '��ҵ
        Dim arrZY 'ְҵ
        Dim arrDjjsxs '��������ϵ��
        Dim arrDdslxsMonth '��ҵ����������
        Dim arrDdslxsSC '����ϵ��
        Dim arrDictStorage
        Dim arrProductRnd
        Dim arrZK '�ۿ�
        Dim arrZKMonth '�ۿ��·ݷֲ�
        Dim arrRndKF '�ͻ��ֲ�
        Dim Rrow As Long
        Dim Rcol As Long


        Dim conn As Object
        Dim RsProduct As Object         '��Ʒ
        Dim RsShop As Object            '�ŵ�
        Dim RsCustomer As Object        '�ͻ�
        Dim RsStorage As Object         '���
        Dim RsOrder As Object           '����
        Dim RsOrdersub As Object        '�����ӱ�
        
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
        arrHY = Array("����", "������", "����", "����ҵ") '�ۿ���ҵ׼��
        arrZY = Array("HR", "����", "����", "��Ӫ")
        Set dictCustomerHY = CreateObject("Scripting.Dictionary") '��ҵ
        Set dictCustomerZY = CreateObject("Scripting.Dictionary") 'ְҵ
        
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
        rowsProduct = UBound(arrProduct) '��Ʒ����
        rowsShop = UBound(arrShop) '�ŵ�����
        rowsCustomer = UBound(arrCustomer) '�ͻ�����
        OcNumber = 0
        
        arrDjjsxs = Array(0.7, 0.8, 1, 1.2, 1.3) '��������ϵ����count=5
        arrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '��ҵ���������ƣ�count=12
        arrDdslxsSC = Array(0.6, 0.65, 0.7, 0.75, 0.8, 0.85, 0.9, 0.95, 1, 1.05, 1.1, 1.15, 1.2, 1.25, 1.3, 1.35, 1.4, 1.4, 1.35, 1.3, 1.25, 1.2, 1.15, 1.1, 1.05, 1, 0.95, 0.9, 0.85, 0.8, 0.75, 0.7, 0.65, 0.6) '���򶩵�ϵ����̫�ֲ���count=34
        arrZK = Array(1, 0.9, 0.8, 0.7, 0.7, 0.6) '�ۿ���Ϣ�ֲ���count=6
        arrZKMonth = Array(0.95, 0.9, 1, 0.98, 0.95, 1, 0.98, 0.98, 0.9, 0.96, 0.92, 0.98) '��ҵ���������ƣ�count=12
        arrRndKF = Array(0, 0.1, 0.5, 0.6, 0.6, 0.6, 0.7, 0.7, 0.7, 0.8, 0.9, 1) '�ͻ��ֲ�

    For rowShop = 0 To rowsShop

            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            'Ӫҵ����
            If IsNull(arrShop(rowShop, 7)) Then
                Yyts = Round(Now - arrShop(rowShop, 2), 0)
            Else
                Yyts = Round(arrShop(rowShop, 7) - arrShop(rowShop, 2), 0)
            End If
            '=====================================================================================
            '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
            Set dictStorageN = CreateObject("Scripting.Dictionary") '�����������
            dictStorageN(1) = Yyts Mod MaxInventoryDays + 1
                    '������ʱ��
                    For i = 1 To Yyts
                        dictStorageN(i + 1) = dictStorageN.item(i) + Round(Rnd() * 2 + 5, 0)
                        If dictStorageN.item(i + 1) > Yyts Then
                            dictStorageN(i + 1) = Yyts
                            Exit For
                        End If
                    Next
            '=====================================================================================

            Set dictStorage = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
            iStorageN = 1
            
           arrPerson = salePersonArr(conn, arrShop(rowShop, 0)) '�ŵ�������Ա 0 ID,1��ְ����,2��ְ���ڻ�ǰ����

            For dayN = 1 To Yyts

                dateOpen = arrShop(rowShop, 2) + dayN - 1
                Randomize
                
                numOrder = Round(Rnd() * 10 * arrDdslxsMonth(Month(dateOpen) - 1) * arrDdslxsSC(arrShop(rowShop, 2) Mod (UBound(arrDdslxsSC) + 1)), 0) 'ÿ�충������

                If numOrder = 0 Then GoTo NoOrder 'û������

                For i = 1 To numOrder 'ÿ�충����
                    OcNumber = OcNumber + 1
                    orderID = "OC_" & Format(OcNumber, "0000000")
                    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                    '��������д��
                    Randomize
                    myRnd = (Rnd() + arrRndKF(dayN * i Mod 12)) / 2

                    rowsCustomerRnd = Round(rowsCustomer * myRnd, 0)
                    
                'ע���빺��ֲ�
                If IsNull(arrShop(rowShop, 7)) Then 'δ�ص�
                    If arrCustomer(rowsCustomerRnd, 1) >= arrShop(rowShop, 2) And OcNumber Mod 13 > 6 Then
                        GoTo rowsCustomerRndLable
                    Else
                        For iCustomer = rowsCustomerRnd To rowsCustomer '����ǰ��
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 13 > 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 > 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                                
                        For iCustomer = rowsCustomerRnd To 0 Step -1 '����ǰ��
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 13 <= 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 <= 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                    End If
                Else '�ص�
                    If arrCustomer(rowsCustomerRnd, 1) >= arrShop(rowShop, 2) And arrCustomer(rowsCustomerRnd, 1) < arrShop(rowShop, 7) And OcNumber Mod 13 < 6 Then
                            GoTo rowsCustomerRndLable
                    Else
                        For iCustomer = rowsCustomerRnd To rowsCustomer '����ǰ��
                            If arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And arrCustomer(iCustomer, 1) < arrShop(rowShop, 7) And Month(arrCustomer(iCustomer, 1)) Mod 13 > 8 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            ElseIf arrCustomer(iCustomer, 1) >= arrShop(rowShop, 2) And Month(arrCustomer(iCustomer, 1)) Mod 3 > 1 Then
                                rowsCustomerRnd = iCustomer
                                GoTo rowsCustomerRndLable
                            End If
                        Next
                            
                        For iCustomer = rowsCustomerRnd To 0 Step -1 '����ǰ��
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
                        Qd = "����"
                    Else
                        Qd = "����"
                    End If
                    orderCount = 0
                    personCount = UBound(arrPerson, 2)
                    orderDate = arrShop(rowShop, 2) + dayN - 1
rndSalePerson: '���������Ա
                    pc = Round(personCount * Rnd, 0)
                    If orderDate < arrPerson(1, pc) Or arrPerson(2, pc) < orderDate Then
                        
                        orderCount = orderCount + 1
                        If orderCount > 200 Then GoTo NoOrder '��֤����ѭ��
                        GoTo rndSalePerson '��Ҫ���㵱ǰ������Ա������ְ�ڼ�

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
                    '�����ӱ�

                    Randomize

                    k = Round(5 * Rnd(), 0) + 1 '�ƻ�ÿ������Ʒ��������,��ֵΪ3

                    Set dictProductRnd = CreateObject("Scripting.Dictionary")
                    For n = 1 To k
                        If k < 4 Then
                            rowsProductRnd = Round(rowsProduct * Rnd() / 5, 0)  '����ƫ��
                        Else
                            rowsProductRnd = Round(rowsProduct * Rnd(), 0)
                        End If
                        dictProductRnd(rowsProductRnd) = rowsProductRnd '�ֵ�ȥ��sku �� ID
                    Next

                    arrProductRnd = dictProductRnd.Keys

                    For n = 0 To dictProductRnd.Count - 1

                        productQuantity = Round(5 * Rnd() * arrDjjsxs(rowShop Mod 5) * arrDjjsxs(arrProductRnd(n) Mod 5), 0) + 1 '����ϵ����Ȩ

                        'q��Ʒ�ۿ�
                        
                        If rowShop Mod 40 > 30 Then '����
                            discount = Round(arrZK(0) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf rowShop Mod 40 < 10 Then
                            discount = Round(arrZK(5) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf arrProductRnd(n) Mod 8 < 1 Then '��Ʒ
                            discount = Round(arrZK(1) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf arrProductRnd(n) Mod 8 > 5 Then
                            discount = Round(arrZK(3) * arrZKMonth(Month(dateOpen) - 1), 2)
                            GoTo ExitIFZhekou
                        ElseIf dictCustomerHY.Exists(arrCustomer(rowsCustomerRnd, 2)) Then  '�ͻ�
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

NoOrder: '�����޶�����ת
                '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
                '���������Ϣ
                If dictStorageN.item(iStorageN) = dayN And dayN < Yyts Then
                    iStorageN = iStorageN + 1

                    arrDictStorage = dictStorage.Keys

                    For iStorage = 0 To UBound(arrDictStorage)
                        RsStorage.AddNew
                            RsStorage.Fields(fStorageProductID) = arrDictStorage(iStorage)
                            RsStorage.Fields(fStorageQuantity) = dictStorage.item(arrDictStorage(iStorage))
                            RsStorage.Fields(fStorageShopID) = arrShop(rowShop, 0)
                            RsStorage.Fields(fStorageDate) = arrShop(rowShop, 2) + dayN - 1 '-1 ��֤�п��
                        RsStorage.Update
                    Next
                    Set dictStorage = Nothing
                    Set dictStorage = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
                    GoTo Rk0
                    
                ElseIf dayN = Yyts Then  '��֤���һ������ۼƴ���0

                    iStorageN = iStorageN + 1

                    arrDictStorage = dictStorage.Keys

                    For iStorage = 0 To UBound(arrDictStorage)
                        RsStorage.AddNew
                            RsStorage.Fields(fStorageProductID) = arrDictStorage(iStorage)
                            RsStorage.Fields(fStorageQuantity) = dictStorage.item(arrDictStorage(iStorage)) + Round(Rnd() * 5, 0)
                            RsStorage.Fields(fStorageShopID) = arrShop(rowShop, 0)
                            RsStorage.Fields(fStorageDate) = arrShop(rowShop, 2) + dayN - 1 '-1 ��֤�п��
                        RsStorage.Update
                    Next
                    Set dictStorage = Nothing
                    Set dictStorage = CreateObject("Scripting.Dictionary") '��¼�����Ϣ
                    GoTo Rk0
                    
                End If
                '=====================================================================================
Rk0:
            Next

        Next

    CloseConnRs conn, RsStorage, RsOrder, RsOrdersub

End Function

Public Function salePersonArr(ByRef conn As Object, ByVal shopID As Long)
    '����������Ա����������Ϊ��������Ա�ֵ䣬��Ա�ֵ�ļ���ΪԱ��ID
    
    Dim EmployeeID As Long, i As Long
    Dim Arr As Variant
    Dim rows As Long, strSQL As String
    Dim RsEmployee As Object

    ' ���� SQL ��ѯ���
    strSQL = "SELECT * FROM " & tbNameEmployee & " WHERE " & fEmployeeOrgID & " = " & shopID & ";"

    ' ������¼������
    Set RsEmployee = CreateObject("ADODB.Recordset")

    ' ִ�в�ѯ����������洢�ڼ�¼��������
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
    '����ҵ���߼������˹��ɱ�

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
    
        ReDim ArrEmployee(0 To 5, 0 To rows) 'Ա��ID����֯ID��ְ��,ѧ��,��ְ����,��ְ����
        RsEmployee.MoveFirst
        i = 0
        Do Until RsEmployee.EOF
            EmployeeID = RsEmployee.Fields(fEmployeeID)
            employeeOrgID = RsEmployee.Fields(fEmployeeOrgID)
            employeeGrade = RsEmployee.Fields(fEmployeeGrade)
            employeeEdu = RsEmployee.Fields(fEmployeeEdu)
            employeeEntryDate = RsEmployee.Fields(fEmployeeEntryDate)
            
            If IsNull(RsEmployee.Fields(fEmployeeResignationDate)) Then
                employeeResignationDate = GetMonthStart(CDate(Format(Now, "YYYY-MM-DD"))) - 1 '��ǰ���ڵ����µ�
            Else
                employeeResignationDate = RsEmployee.Fields(fEmployeeResignationDate)
            End If

        daysFirst = day(employeeEntryDate)
        daysMonthFirst = GetDaysInMonth(employeeEntryDate)
        
        daysLast = day(employeeResignationDate)
        daysMonthLast = GetDaysInMonth(employeeResignationDate)
        
        numMonth = DateDiffInMonths(employeeEntryDate, employeeResignationDate)
        dateMonthStartFirst = GetMonthStart(employeeEntryDate)
        
        salaryUp = GradeSalaryDict(employeeGrade)(0) '��������
        salaryDown = GradeSalaryDict(employeeGrade)(1) '��������
        eduN = EduSalaryDict(employeeEdu) 'ѧ��ϵ��
        salaryRndBase = Round((salaryUp - salaryDown + 1) * Rnd + salaryDown, 0) '���ʻ���
        
        For i = 0 To numMonth
                    
            If i = 0 Then
                salaryRnd = Round(salaryRndBase * daysFirst / daysMonthFirst, 0) '�����жϹ�������
            ElseIf i = numMonth Then
                salaryRnd = Round(salaryRndBase * daysLast / daysMonthLast, 0) 'ĩ���жϹ�������
            Else
                salaryRnd = salaryRndBase * (0.8 + (Rnd * (1.2 - 0.8))) * 1.36  '�籣������ϵ�� 0.8 �� 1.2�ĸ���
            End If
            
            
            dateMonthStartAC = AddMonths(dateMonthStartFirst, i) '�����·�
            
            yearsService = Round(i / 12, 0) '˾��
            
            salary = salaryRnd * eduN * (1 + yearsService * 0.1) '���¹���

            key = CStr(employeeOrgID & delimiter & dateMonthStartAC) '����

            AddDictByKey laborCostDict, key, salary '������֯ID���·��ۼƳɱ�
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
' ����ҵ���߼����� ����Ŀ���
    Dim Sqlstr As String
    Dim ArrYQ
    Dim arrDdslxsMonth '��ҵ����������
    Dim Qn As Double '��������µ�ϵ����
    Dim B As Double
    Dim UP0 As Double '����������
    
    Dim i As Long
    Dim Rrow As Long
    Dim Rcol As Long
    Dim k As Long
    
    Dim conn  As Object
    Dim RsSaleTarget  As Object

    Set conn = CreateConnection

    '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
    'ȥ��������ȫ��&ȥ��Q4������,�¾�ȡ��
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
    arrDdslxsMonth = Array(1, 0.5, 0.9, 1, 1.2, 0.9, 0.9, 1, 1.3, 1.2, 1.1, 1) '��ҵ����������,��һ��count=12;ͬ��DataTableT345

    Set RsSaleTarget = CreateRecordset(conn, tbNameSaleTarget)

    For i = UBound(arrDdslxsMonth) - 2 To UBound(arrDdslxsMonth)
        Qn = arrDdslxsMonth(i) + Qn
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
            RsSaleTarget.AddNew
            RsSaleTarget.Fields(fSaleTargetProvinceID) = ArrYQ(i, 0)
            RsSaleTarget.Fields(fSaleTargetProvinceName2) = ArrYQ(i, 1)
            RsSaleTarget.Fields(fSaleTargetMonth) = Format(Now(), "YYYY") - 1 & "-" & k & "-1"
            Randomize
            RsSaleTarget.Fields(fSaleTargetAmount) = Round(B * (0.7 + Rnd() * 0.1 * UP0) * arrDdslxsMonth(k - 1), 0) '�·�����Ŀ��ı�����һ�������궼����ʵ�ʸ�������
            RsSaleTarget.Update
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
' ����ҵ���߼����� �߹�
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
    
    Set dictGender = GenderDict() 'ǰ�� 448 �����Ը���ͷ������
    Set dictAllDate = DateStatusDict() '��������״̬���ֵ�
    
    InitE '��ʼ��Ա����Ϣ�������
    Set conn = CreateConnection
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)

'=====================================================================================
'һ������ �� ���۴���

    'Ĭ����֯ �������ڡ���ְ���ڡ���ְ���ڡ���ְԭ���������
    
    Const employeeStr As String = "�ܾ���,     ��, 1,      �ܾ���,        �ܾ���       ;" & _
                                  "����,       ��, 2,      �ܾ�������,    �߼��ܼ�     ;" & _
                                  "ʩ��,       ��, 3,      ��Ʒ�ܼ�,      �߼��ܼ�     ;" & _
                                  "�ҵ���,     Ů, 4,      �ɹ��ܼ�,      �߼��ܼ�     ;" & _
                                  "���,       ��, 5,      �����ܼ�,      �߼��ܼ�     ;" & _
                                  "����,       ��, 6,      ������Դ�ܼ�,  �߼��ܼ�     ;" & _
                                  "�Ž��,     ��, 7,      �ۺ�����ܼ�,  �߼��ܼ�     ;" & _
                                  "�ɿ�,       Ů, 8,      �����ܼ�,      �߼��ܼ�     ;" & _
                                  "ŷ������,   ��, 9,      ��������,      �ܼ�         ;" & _
                                  "������,     ��, 10,     ��������,      �ܼ�         ;" & _
                                  "������,   Ů, 11,     ��������,      �ܼ�         ;" & _
                                  "������,     ��, 12,     ��������,      �ܼ�         ;" & _
                                  "������,     Ů, 13,     ��������,      �ܼ�         ;" & _
                                  "������,     ��, 14,     ��������,      �ܼ�          "

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
    '��ȡԱ����������
    Dim EmployeeID As Long, modxR As Long, mod2R As Long, dateStatus As Long, entryDate As Date

    RsEmployee.AddNew
        RsEmployee.Fields(fEmployeeName) = employeeName
        If employeeGender = "Ů" Then RsEmployee.Fields(fEmployeeGender) = "Ů"
        If employeeOrgID Then RsEmployee.Fields(fEmployeeOrgID) = employeeOrgID
        RsEmployee.Fields(fEmployeeJobTitle) = employeeJobTitle
        RsEmployee.Fields(fEmployeeGrade) = employeeGrade
        RsEmployee.Fields(fEmployeeEdu) = employeeEdu
        RsEmployee.Fields(fEmployeeBirthday) = employeeBirthday
        RsEmployee.Fields(fEmployeeEntryDate) = employeeEntryDate
    RsEmployee.Update
    
        EmployeeID = RsEmployee.Fields(fEmployeeID)
        mod2R = EmployeeID Mod 2
        
        'ͳһȷ��ǰ��448�˵��Ա����ǰ׼���õ�ͷ��ƥ��
        If dictGender.Exists(EmployeeID) Then
            RsEmployee.Fields(fEmployeeGender) = dictGender(EmployeeID)
            RsEmployee.Update
        End If
        
        
        '��ְʱ�䲻����Ա����Ϣ��
        entryDate = employeeEntryDate
        
EffectiveEntryDate: 'ȷ����Ч����ְ����

        modxR = CLng(dictAllDate(entryDate)("modx"))
        dateStatus = CLng(dictAllDate(entryDate)("status"))

        If modxR <= 4 And mod2R = 1 And dateStatus < 3 Then 'Ա�����Ϊ����������Ϊ��0,1,2,3,4 �����ա�����
            RsEmployee.Fields(fEmployeeEntryDate) = entryDate
            RsEmployee.Update
        ElseIf modxR >= 2 And mod2R = 0 And dateStatus < 3 Then 'Ա�����Ϊż��������Ϊ��2,3,4,5,6 �����ա�����
            RsEmployee.Fields(fEmployeeEntryDate) = entryDate
            RsEmployee.Update
        Else
            entryDate = entryDate + 1
            GoTo EffectiveEntryDate
        End If
                
End Function

Public Function DataTableEmployeeRegular()
' ����ҵ���߼����� �ŵ�������Ա

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
    
    Set dictGender = GenderDict() 'ǰ��448��Ա���Ա��ֵ�
    Set dictAllDate = DateStatusDict() '��������״̬���ֵ�

    InitE '��ʼ��Ա����Ϣ�������
    Set conn = CreateConnection
    Set RsEmployee = CreateRecordset(conn, tbNameEmployee)
    Set RsShop = CreateRecordset(conn, tbNameShop)
    
    RsShop.MoveFirst
    Do Until RsShop.EOF
        SetShopRandomSeed RsShop.Fields(fShopDistrictID).value
        numEmployee = Round(2 + Rnd() * 7, 0)
    
        dateOpen = RsShop.Fields(fShopOpenDate)
        
        If IsNull(RsShop.Fields(fShopCloseDate)) Then
            Yyts = Round(Now - dateOpen, 0) 'Ӫҵ����
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
            If myRnd < 0.7 Then employeeGender = "Ů" Else employeeGender = "��"
            employeeOrgID = RsShop.Fields(fShopID)
            employeeJobTitle = JobTitlesArr(12)
            employeeGrade = GradeArr(6)
            employeeEdu = EduArr(Round(1 + Rnd() * 2, 0))
            employeeBirthday = MinDateOpen - Round((Rnd() + 1) * 8000, 0)
            employeeEntryDate = dateEntry
    
            AddEmployeeRecord RsEmployee, employeeName, employeeGender, employeeJobTitle, employeeGrade, employeeEdu, employeeBirthday, employeeEntryDate, dictAllDate, dictGender, employeeOrgID
                    
'           ResignationArr = Array("���˷�չ", "����ԭ��", "����ǿ��", "���������뻷��", "��ͥԭ��", "����ԭ��", "Υ�������ƶ�", "Ȱ��", "����", "����ԭ��", "�������ڽ��") '�����ڷ�������10
            If RsEmployee.Fields(fEmployeeID) Mod 13 = 12 Then
                workDays = Round(Yyts * myRnd, 0)
                dateResignation = dateOpen + workDays
                
                days = 0
                '��֤��ְ���ڲ�������Ϣ�պͼ���
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
                
                '��ְ��һ��
                SetShopRandomSeed RsShop.Fields(fShopDistrictID).value
                Randomize
                myRnd = Rnd()
                dateEntry = dateResignation + Round(30 * myRnd - 15, 0)
                If dateEntry > Now Then dateEntry = Format(Now, "YYYY-MM-DD")
                
                employeeName = generateName(myRnd)
                If myRnd < 0.7 Then employeeGender = "Ů" Else employeeGender = "��"
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
    '�����������
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
    ' �������Ӷ���
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    Set conn = CurrentProject.Connection
    Set CreateConnection = conn
End Function

Public Function CreateRecordset(ByRef conn As Object, source As String) As Object
    '���� Recordset ����
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
    ' �ر� Recordset ����
    If Not (rs Is Nothing) Then
        rs.Close
        Set rs = Nothing
    End If
End Function

Public Function ConnClose(ByRef conn As Object)
    ' �ر� Connection ����
    If Not (conn Is Nothing) Then
        conn.Close
        Set conn = Nothing
    End If
End Function

Public Function CloseConnRs(ByRef conn As Object, ParamArray rsList() As Variant)
    Dim i As Integer
    
    ' �ر����� Recordset ����
    For i = LBound(rsList) To UBound(rsList)
        If Not (rsList(i) Is Nothing) Then
            rsList(i).Close
            Set rsList(i) = Nothing
        End If
    Next i
    
    ' �ر� Connection ����
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

    For i = 1 To Len(str) ' ������ѭ����ʼֵ����ʹ�� Len(str) ��̬���ý���ֵ
        seed = seed + CInt(Mid(str, i, 1)) ' ����Ϊ��ȷ�� Mid �����÷�
    Next

    Randomize seed
End Function
Public Function FirstName() As String
' ������������

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
' ������������
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
' ����ʡ�����ݣ��������ƣ������
    '����ID ʡID    ʡȫ��  ʡ���1  ʡ���2  γ��    ����
    AddressProvince = "13,110000,������,��,����,39.904987,116.405289;"
    AddressProvince = AddressProvince & "13,120000,�����,��,���,39.125595,117.190186;"
    AddressProvince = AddressProvince & "13,130000,�ӱ�ʡ,��,�ӱ�,38.045475,114.502464;"
    AddressProvince = AddressProvince & "13,140000,ɽ��ʡ,��,ɽ��,37.857014,112.549248;"
    AddressProvince = AddressProvince & "12,150000,���ɹ�������,��,���ɹ�,40.81831,111.670799;"
    AddressProvince = AddressProvince & "12,210000,����ʡ,��,����,41.796768,123.429092;"
    AddressProvince = AddressProvince & "12,220000,����ʡ,��,����,43.886841,125.324501;"
    AddressProvince = AddressProvince & "12,230000,������ʡ,��,������,45.756966,126.642464;"
    AddressProvince = AddressProvince & "9,310000,�Ϻ���,��,�Ϻ�,31.231707,121.472641;"
    AddressProvince = AddressProvince & "9,320000,����ʡ,��,����,32.041546,118.76741;"
    AddressProvince = AddressProvince & "9,330000,�㽭ʡ,��,�㽭,30.287458,120.15358;"
    AddressProvince = AddressProvince & "9,340000,����ʡ,��,����,31.861191,117.283043;"
    AddressProvince = AddressProvince & "9,350000,����ʡ,��,����,26.075302,119.306236;"
    AddressProvince = AddressProvince & "9,360000,����ʡ,��,����,28.676493,115.892151;"
    AddressProvince = AddressProvince & "12,370000,ɽ��ʡ,³,ɽ��,36.675808,117.000923;"
    AddressProvince = AddressProvince & "13,410000,����ʡ,ԥ,����,34.757977,113.665413;"
    AddressProvince = AddressProvince & "13,420000,����ʡ,��,����,30.584354,114.298569;"
    AddressProvince = AddressProvince & "11,430000,����ʡ,��,����,28.19409,112.982277;"
    AddressProvince = AddressProvince & "11,440000,�㶫ʡ,��,�㶫,23.125177,113.28064;"
    AddressProvince = AddressProvince & "11,450000,����׳��������,��,����,22.82402,108.320007;"
    AddressProvince = AddressProvince & "11,460000,����ʡ,��,����,20.031971,110.331192;"
    AddressProvince = AddressProvince & "10,500000,������,��,����,29.533155,106.504959;"
    AddressProvince = AddressProvince & "10,510000,�Ĵ�ʡ,��,�Ĵ�,30.659462,104.065735;"
    AddressProvince = AddressProvince & "11,520000,����ʡ,ǭ,����,26.578342,106.713478;"
    AddressProvince = AddressProvince & "11,530000,����ʡ,��,����,25.040609,102.71225;"
    AddressProvince = AddressProvince & "10,540000,����������,��,����,29.66036,91.13221;"
    AddressProvince = AddressProvince & "13,610000,����ʡ,��,����,34.263161,108.948021;"
    AddressProvince = AddressProvince & "10,620000,����ʡ,��,����,36.058041,103.823555;"
    AddressProvince = AddressProvince & "10,630000,�ຣʡ,��,�ຣ,36.623177,101.778915;"
    AddressProvince = AddressProvince & "10,640000,���Ļ���������,��,����,38.46637,106.278175;"
    AddressProvince = AddressProvince & "10,650000,�½�ά���������,��,�½�,43.792816,87.617729;"
    AddressProvince = AddressProvince & "14,710000,̨��ʡ,̨,̨��,25.041618,121.501618;"
    AddressProvince = AddressProvince & "14,810000,����ر�������,��,���,22.320047,114.173355;"
    AddressProvince = AddressProvince & "14,820000,�����ر�������,��,����,22.198952,113.549088"


End Function

Public Function AddressCity() As String
' ���е������ݣ��������ƣ�����ȡ�
    'ʡID    ����ID  ����    γ��    ����
    AddressCity = "110000,110000,����,39.904989,116.405285;"
    AddressCity = AddressCity & "120000,120000,���,39.125596,117.190182;"
    AddressCity = AddressCity & "130000,130100,ʯ��ׯ,38.045474,114.502461;"
    AddressCity = AddressCity & "130000,130200,��ɽ,39.635113,118.175393;"
    AddressCity = AddressCity & "130000,130300,�ػʵ�,39.942531,119.586579;"
    AddressCity = AddressCity & "130000,130400,����,36.612273,114.490686;"
    AddressCity = AddressCity & "130000,130500,��̨,37.0682,114.508851;"
    AddressCity = AddressCity & "130000,130600,����,38.867657,115.482331;"
    AddressCity = AddressCity & "130000,130700,�żҿ�,40.811901,114.884091;"
    AddressCity = AddressCity & "130000,130800,�е�,40.976204,117.939152;"
    AddressCity = AddressCity & "130000,130900,����,38.310582,116.857461;"
    AddressCity = AddressCity & "130000,131000,�ȷ�,39.523927,116.704441;"
    AddressCity = AddressCity & "130000,131100,��ˮ,37.735097,115.665993;"
    AddressCity = AddressCity & "140000,140100,̫ԭ,37.857014,112.549248;"
    AddressCity = AddressCity & "140000,140200,��ͬ,40.09031,113.295259;"
    AddressCity = AddressCity & "140000,140300,��Ȫ,37.861188,113.583285;"
    AddressCity = AddressCity & "140000,140400,����,36.191112,113.113556;"
    AddressCity = AddressCity & "140000,140500,����,35.497553,112.851274;"
    AddressCity = AddressCity & "140000,140600,˷��,39.331261,112.433387;"
    AddressCity = AddressCity & "140000,140700,����,37.696495,112.736465;"
    AddressCity = AddressCity & "140000,140800,�˳�,35.022778,111.003957;"
    AddressCity = AddressCity & "140000,140900,����,38.41769,112.733538;"
    AddressCity = AddressCity & "140000,141000,�ٷ�,36.08415,111.517973;"
    AddressCity = AddressCity & "140000,141100,����,37.524366,111.134335;"
    AddressCity = AddressCity & "150000,150100,���ͺ���,40.818311,111.670801;"
    AddressCity = AddressCity & "150000,150200,��ͷ,40.658168,109.840405;"
    AddressCity = AddressCity & "150000,150300,�ں�,39.673734,106.825563;"
    AddressCity = AddressCity & "150000,150400,���,42.275317,118.956806;"
    AddressCity = AddressCity & "150000,150500,ͨ��,43.617429,122.263119;"
    AddressCity = AddressCity & "150000,150600,������˹,39.817179,109.99029;"
    AddressCity = AddressCity & "150000,150700,���ױ���,49.215333,119.758168;"
    AddressCity = AddressCity & "150000,150800,�����׶�,40.757402,107.416959;"
    AddressCity = AddressCity & "150000,150900,�����첼,41.034126,113.114543;"
    AddressCity = AddressCity & "150000,152200,�˰�,46.076268,122.070317;"
    AddressCity = AddressCity & "150000,152500,���ֹ���,43.944018,116.090996;"
    AddressCity = AddressCity & "150000,152900,������,38.844814,105.706422;"
    AddressCity = AddressCity & "210000,210100,����,41.796767,123.429096;"
    AddressCity = AddressCity & "210000,210200,����,38.91459,121.618622;"
    AddressCity = AddressCity & "210000,210300,��ɽ,41.110626,122.995632;"
    AddressCity = AddressCity & "210000,210400,��˳,41.875956,123.921109;"
    AddressCity = AddressCity & "210000,210500,��Ϫ,41.297909,123.770519;"
    AddressCity = AddressCity & "210000,210600,����,40.124296,124.383044;"
    AddressCity = AddressCity & "210000,210700,����,41.119269,121.135742;"
    AddressCity = AddressCity & "210000,210800,Ӫ��,40.667432,122.235151;"
    AddressCity = AddressCity & "210000,210900,����,42.011796,121.648962;"
    AddressCity = AddressCity & "210000,211000,����,41.269402,123.18152;"
    AddressCity = AddressCity & "210000,211100,�̽�,41.124484,122.06957;"
    AddressCity = AddressCity & "210000,211200,����,42.290585,123.844279;"
    AddressCity = AddressCity & "210000,211300,����,41.576758,120.451176;"
    AddressCity = AddressCity & "210000,211400,��«��,40.755572,120.856394;"
    AddressCity = AddressCity & "220000,220100,����,43.886841,125.3245;"
    AddressCity = AddressCity & "220000,220200,����,43.843577,126.55302;"
    AddressCity = AddressCity & "220000,220300,��ƽ,43.170344,124.370785;"
    AddressCity = AddressCity & "220000,220400,��Դ,42.902692,125.145349;"
    AddressCity = AddressCity & "220000,220500,ͨ��,41.721177,125.936501;"
    AddressCity = AddressCity & "220000,220600,��ɽ,41.942505,126.427839;"
    AddressCity = AddressCity & "220000,220700,��ԭ,45.118243,124.823608;"
    AddressCity = AddressCity & "220000,220800,�׳�,45.619026,122.841114;"
    AddressCity = AddressCity & "220000,222400,�ӱ߳�����,42.904823,129.513228;"
    AddressCity = AddressCity & "230000,230100,������,45.756967,126.642464;"
    AddressCity = AddressCity & "230000,230200,�������,47.342081,123.95792;"
    AddressCity = AddressCity & "230000,230300,����,45.300046,130.975966;"
    AddressCity = AddressCity & "230000,230400,�׸�,47.332085,130.277487;"
    AddressCity = AddressCity & "230000,230500,˫Ѽɽ,46.643442,131.157304;"
    AddressCity = AddressCity & "230000,230600,����,46.590734,125.11272;"
    AddressCity = AddressCity & "230000,230700,����,47.724775,128.899396;"
    AddressCity = AddressCity & "230000,230800,��ľ˹,46.809606,130.361634;"
    AddressCity = AddressCity & "230000,230900,��̨��,45.771266,131.015584;"
    AddressCity = AddressCity & "230000,231000,ĵ����,44.582962,129.618602;"
    AddressCity = AddressCity & "230000,231100,�ں�,50.249585,127.499023;"
    AddressCity = AddressCity & "230000,231200,�绯,46.637393,126.99293;"
    AddressCity = AddressCity & "230000,232700,���˰���,52.335262,124.711526;"
    AddressCity = AddressCity & "310000,310000,�Ϻ�,31.231706,121.472644;"
    AddressCity = AddressCity & "320000,320100,�Ͼ�,32.041544,118.767413;"
    AddressCity = AddressCity & "320000,320200,����,31.574729,120.301663;"
    AddressCity = AddressCity & "320000,320300,����,34.261792,117.184811;"
    AddressCity = AddressCity & "320000,320400,����,31.772752,119.946973;"
    AddressCity = AddressCity & "320000,320500,����,31.299379,120.619585;"
    AddressCity = AddressCity & "320000,320600,��ͨ,32.016212,120.864608;"
    AddressCity = AddressCity & "320000,320700,���Ƹ�,34.600018,119.178821;"
    AddressCity = AddressCity & "320000,320800,����,33.597506,119.021265;"
    AddressCity = AddressCity & "320000,320900,�γ�,33.377631,120.139998;"
    AddressCity = AddressCity & "320000,321000,����,32.393159,119.421003;"
    AddressCity = AddressCity & "320000,321100,��,32.204402,119.452753;"
    AddressCity = AddressCity & "320000,321200,̩��,32.484882,119.915176;"
    AddressCity = AddressCity & "320000,321300,��Ǩ,33.963008,118.275162;"
    AddressCity = AddressCity & "330000,330100,����,30.287459,120.153576;"
    AddressCity = AddressCity & "330000,330200,����,29.868388,121.549792;"
    AddressCity = AddressCity & "330000,330300,����,28.000575,120.672111;"
    AddressCity = AddressCity & "330000,330400,����,30.762653,120.750865;"
    AddressCity = AddressCity & "330000,330500,����,30.867198,120.102398;"
    AddressCity = AddressCity & "330000,330600,����,29.997117,120.582112;"
    AddressCity = AddressCity & "330000,330700,��,29.089524,119.649506;"
    AddressCity = AddressCity & "330000,330800,����,28.941708,118.87263;"
    AddressCity = AddressCity & "330000,330900,��ɽ,30.016028,122.106863;"
    AddressCity = AddressCity & "330000,331000,̨��,28.661378,121.428599;"
    AddressCity = AddressCity & "330000,331100,��ˮ,28.451993,119.921786;"
    AddressCity = AddressCity & "340000,340100,�Ϸ�,31.86119,117.283042;"
    AddressCity = AddressCity & "340000,340200,�ߺ�,31.326319,118.376451;"
    AddressCity = AddressCity & "340000,340300,����,32.939667,117.363228;"
    AddressCity = AddressCity & "340000,340400,����,32.647574,117.018329;"
    AddressCity = AddressCity & "340000,340500,��ɽ,31.689362,118.507906;"
    AddressCity = AddressCity & "340000,340600,����,33.971707,116.794664;"
    AddressCity = AddressCity & "340000,340700,ͭ��,30.929935,117.816576;"
    AddressCity = AddressCity & "340000,340800,����,30.50883,117.043551;"
    AddressCity = AddressCity & "340000,341000,��ɽ,29.709239,118.317325;"
    AddressCity = AddressCity & "340000,341100,����,32.303627,118.316264;"
    AddressCity = AddressCity & "340000,341200,����,32.896969,115.819729;"
    AddressCity = AddressCity & "340000,341300,����,33.633891,116.984084;"
    AddressCity = AddressCity & "340000,341500,����,31.752889,116.507676;"
    AddressCity = AddressCity & "340000,341600,����,33.869338,115.782939;"
    AddressCity = AddressCity & "340000,341700,����,30.656037,117.489157;"
    AddressCity = AddressCity & "340000,341800,����,30.945667,118.757995;"
    AddressCity = AddressCity & "350000,350100,����,26.075302,119.306239;"
    AddressCity = AddressCity & "350000,350200,����,24.490474,118.11022;"
    AddressCity = AddressCity & "350000,350300,����,25.431011,119.007558;"
    AddressCity = AddressCity & "350000,350400,����,26.265444,117.635001;"
    AddressCity = AddressCity & "350000,350500,Ȫ��,24.908853,118.589421;"
    AddressCity = AddressCity & "350000,350600,����,24.510897,117.661801;"
    AddressCity = AddressCity & "350000,350700,��ƽ,26.635627,118.178459;"
    AddressCity = AddressCity & "350000,350800,����,25.091603,117.02978;"
    AddressCity = AddressCity & "350000,350900,����,26.65924,119.527082;"
    AddressCity = AddressCity & "360000,360100,�ϲ�,28.676493,115.892151;"
    AddressCity = AddressCity & "360000,360200,������,29.29256,117.214664;"
    AddressCity = AddressCity & "360000,360300,Ƽ��,27.622946,113.852186;"
    AddressCity = AddressCity & "360000,360400,�Ž�,29.712034,115.992811;"
    AddressCity = AddressCity & "360000,360500,����,27.810834,114.930835;"
    AddressCity = AddressCity & "360000,360600,ӥ̶,28.238638,117.033838;"
    AddressCity = AddressCity & "360000,360700,����,25.85097,114.940278;"
    AddressCity = AddressCity & "360000,360800,����,27.111699,114.986373;"
    AddressCity = AddressCity & "360000,360900,�˴�,27.8043,114.391136;"
    AddressCity = AddressCity & "360000,361000,����,27.98385,116.358351;"
    AddressCity = AddressCity & "360000,361100,����,28.44442,117.971185;"
    AddressCity = AddressCity & "370000,370100,����,36.675807,117.000923;"
    AddressCity = AddressCity & "370000,370200,�ൺ,36.082982,120.355173;"
    AddressCity = AddressCity & "370000,370300,�Ͳ�,36.814939,118.047648;"
    AddressCity = AddressCity & "370000,370400,��ׯ,34.856424,117.557964;"
    AddressCity = AddressCity & "370000,370500,��Ӫ,37.434564,118.66471;"
    AddressCity = AddressCity & "370000,370600,��̨,37.539297,121.391382;"
    AddressCity = AddressCity & "370000,370700,Ϋ��,36.70925,119.107078;"
    AddressCity = AddressCity & "370000,370800,����,35.415393,116.587245;"
    AddressCity = AddressCity & "370000,370900,̩��,36.194968,117.129063;"
    AddressCity = AddressCity & "370000,371000,����,37.509691,122.116394;"
    AddressCity = AddressCity & "370000,371100,����,35.428588,119.461208;"
    AddressCity = AddressCity & "370000,371300,����,35.065282,118.326443;"
    AddressCity = AddressCity & "370000,371400,����,37.453968,116.307428;"
    AddressCity = AddressCity & "370000,371500,�ĳ�,36.456013,115.980367;"
    AddressCity = AddressCity & "370000,371600,����,37.383542,118.016974;"
    AddressCity = AddressCity & "370000,371700,����,35.246531,115.469381;"
    AddressCity = AddressCity & "410000,410100,֣��,34.757975,113.665412;"
    AddressCity = AddressCity & "410000,410200,����,34.797049,114.341447;"
    AddressCity = AddressCity & "410000,410300,����,34.663041,112.434468;"
    AddressCity = AddressCity & "410000,410400,ƽ��ɽ,33.735241,113.307718;"
    AddressCity = AddressCity & "410000,410500,����,36.103442,114.352482;"
    AddressCity = AddressCity & "410000,410600,�ױ�,35.748236,114.295444;"
    AddressCity = AddressCity & "410000,410700,����,35.302616,113.883991;"
    AddressCity = AddressCity & "410000,410800,����,35.23904,113.238266;"
    AddressCity = AddressCity & "410000,419001,��Դ,35.090378,112.590047;"
    AddressCity = AddressCity & "410000,410900,���,35.768234,115.041299;"
    AddressCity = AddressCity & "410000,411000,���,34.022956,113.826063;"
    AddressCity = AddressCity & "410000,411100,���,33.575855,114.026405;"
    AddressCity = AddressCity & "410000,411200,����Ͽ,34.777338,111.194099;"
    AddressCity = AddressCity & "410000,411300,����,32.999082,112.540918;"
    AddressCity = AddressCity & "410000,411400,����,34.437054,115.650497;"
    AddressCity = AddressCity & "410000,411500,����,32.123274,114.075031;"
    AddressCity = AddressCity & "410000,411600,�ܿ�,33.620357,114.649653;"
    AddressCity = AddressCity & "410000,411700,פ���,32.980169,114.024736;"
    AddressCity = AddressCity & "420000,420100,�人,30.584355,114.298572;"
    AddressCity = AddressCity & "420000,420200,��ʯ,30.220074,115.077048;"
    AddressCity = AddressCity & "420000,420300,ʮ��,32.646907,110.787916;"
    AddressCity = AddressCity & "420000,420500,�˲�,30.702636,111.290843;"
    AddressCity = AddressCity & "420000,420600,����,32.042426,112.144146;"
    AddressCity = AddressCity & "420000,420700,����,30.396536,114.890593;"
    AddressCity = AddressCity & "420000,420800,����,31.03542,112.204251;"
    AddressCity = AddressCity & "420000,420900,Т��,30.926423,113.926655;"
    AddressCity = AddressCity & "420000,421000,����,30.326857,112.23813;"
    AddressCity = AddressCity & "420000,421100,�Ƹ�,30.447711,114.879365;"
    AddressCity = AddressCity & "420000,421200,����,29.832798,114.328963;"
    AddressCity = AddressCity & "420000,421300,����,31.717497,113.37377;"
    AddressCity = AddressCity & "420000,422800,��ʩ,30.283114,109.48699;"
    AddressCity = AddressCity & "420000,429004,����,30.364953,113.453974;"
    AddressCity = AddressCity & "420000,429005,Ǳ��,30.421215,112.896866;"
    AddressCity = AddressCity & "420000,429006,����,30.653061,113.165862;"
    AddressCity = AddressCity & "420000,429021,��ũ��,31.744449,110.671525;"
    AddressCity = AddressCity & "430000,430100,��ɳ,28.19409,112.982279;"
    AddressCity = AddressCity & "430000,430200,����,27.835806,113.151737;"
    AddressCity = AddressCity & "430000,430300,��̶,27.82973,112.944052;"
    AddressCity = AddressCity & "430000,430400,����,26.900358,112.607693;"
    AddressCity = AddressCity & "430000,430500,����,27.237842,111.46923;"
    AddressCity = AddressCity & "430000,430600,����,29.37029,113.132855;"
    AddressCity = AddressCity & "430000,430700,����,29.040225,111.691347;"
    AddressCity = AddressCity & "430000,430800,�żҽ�,29.127401,110.479921;"
    AddressCity = AddressCity & "430000,430900,����,28.570066,112.355042;"
    AddressCity = AddressCity & "430000,431000,����,25.793589,113.032067;"
    AddressCity = AddressCity & "430000,431100,����,26.434516,111.608019;"
    AddressCity = AddressCity & "430000,431200,����,27.550082,109.97824;"
    AddressCity = AddressCity & "430000,431300,¦��,27.728136,112.008497;"
    AddressCity = AddressCity & "430000,433100,����,28.314296,109.739735;"
    AddressCity = AddressCity & "440000,440100,����,23.125178,113.280637;"
    AddressCity = AddressCity & "440000,440200,�ع�,24.801322,113.591544;"
    AddressCity = AddressCity & "440000,440300,����,22.547,114.085947;"
    AddressCity = AddressCity & "440000,440400,�麣,22.224979,113.553986;"
    AddressCity = AddressCity & "440000,440500,��ͷ,23.37102,116.708463;"
    AddressCity = AddressCity & "440000,440600,��ɽ,23.028762,113.122717;"
    AddressCity = AddressCity & "440000,440700,����,22.590431,113.094942;"
    AddressCity = AddressCity & "440000,440800,տ��,21.274898,110.364977;"
    AddressCity = AddressCity & "440000,440900,ï��,21.659751,110.919229;"
    AddressCity = AddressCity & "440000,441200,����,23.051546,112.472529;"
    AddressCity = AddressCity & "440000,441300,����,23.079404,114.412599;"
    AddressCity = AddressCity & "440000,441400,÷��,24.299112,116.117582;"
    AddressCity = AddressCity & "440000,441500,��β,22.774485,115.364238;"
    AddressCity = AddressCity & "440000,441600,��Դ,23.746266,114.697802;"
    AddressCity = AddressCity & "440000,441700,����,21.859222,111.975107;"
    AddressCity = AddressCity & "440000,441800,��Զ,23.685022,113.051227;"
    AddressCity = AddressCity & "440000,441900,��ݸ,23.046237,113.746262;"
    AddressCity = AddressCity & "440000,442000,��ɽ,22.521113,113.382391;"
    AddressCity = AddressCity & "440000,445100,����,23.661701,116.632301;"
    AddressCity = AddressCity & "440000,445200,����,23.543778,116.355733;"
    AddressCity = AddressCity & "440000,445300,�Ƹ�,22.929801,112.044439;"
    AddressCity = AddressCity & "450000,450100,����,22.82402,108.320004;"
    AddressCity = AddressCity & "450000,450200,����,24.314617,109.411703;"
    AddressCity = AddressCity & "450000,450300,����,25.274215,110.299121;"
    AddressCity = AddressCity & "450000,450400,����,23.474803,111.297604;"
    AddressCity = AddressCity & "450000,450500,����,21.473343,109.119254;"
    AddressCity = AddressCity & "450000,450600,���Ǹ�,21.614631,108.345478;"
    AddressCity = AddressCity & "450000,450700,����,21.967127,108.624175;"
    AddressCity = AddressCity & "450000,450800,���,23.0936,109.602146;"
    AddressCity = AddressCity & "450000,450900,����,22.63136,110.154393;"
    AddressCity = AddressCity & "450000,451000,��ɫ,23.897742,106.616285;"
    AddressCity = AddressCity & "450000,451100,����,24.414141,111.552056;"
    AddressCity = AddressCity & "450000,451200,�ӳ�,24.695899,108.062105;"
    AddressCity = AddressCity & "450000,451300,����,23.733766,109.229772;"
    AddressCity = AddressCity & "450000,451400,����,22.404108,107.353926;"
    AddressCity = AddressCity & "460000,460100,����,20.031971,110.33119;"
    AddressCity = AddressCity & "460000,460200,����,18.247872,109.508268;"
    AddressCity = AddressCity & "460000,460300,��ɳ,16.831039,112.34882;"
    AddressCity = AddressCity & "460000,469001,��ָɽ,18.776921,109.516662;"
    AddressCity = AddressCity & "460000,469002,��,19.246011,110.466785;"
    AddressCity = AddressCity & "460000,460400,����,19.517486,109.576782;"
    AddressCity = AddressCity & "460000,469005,�Ĳ�,19.612986,110.753975;"
    AddressCity = AddressCity & "460000,469006,����,18.796216,110.388793;"
    AddressCity = AddressCity & "460000,469007,����,19.10198,108.653789;"
    AddressCity = AddressCity & "460000,469021,����,19.684966,110.349235;"
    AddressCity = AddressCity & "460000,469022,�Ͳ�,19.362916,110.102773;"
    AddressCity = AddressCity & "460000,469023,����,19.737095,110.007147;"
    AddressCity = AddressCity & "460000,469024,�ٸ�,19.908293,109.687697;"
    AddressCity = AddressCity & "460000,469025,��ɳ,19.224584,109.452606;"
    AddressCity = AddressCity & "460000,469026,����,19.260968,109.053351;"
    AddressCity = AddressCity & "460000,469027,�ֶ�,18.74758,109.175444;"
    AddressCity = AddressCity & "460000,469028,��ˮ,18.505006,110.037218;"
    AddressCity = AddressCity & "460000,469029,��ͤ,18.636371,109.70245;"
    AddressCity = AddressCity & "460000,469030,����,19.03557,109.839996;"
    AddressCity = AddressCity & "500000,500000,����,29.533155,106.504962;"
    AddressCity = AddressCity & "510000,510100,�ɶ�,30.659462,104.065735;"
    AddressCity = AddressCity & "510000,510300,�Թ�,29.352765,104.773447;"
    AddressCity = AddressCity & "510000,510400,��֦��,26.580446,101.716007;"
    AddressCity = AddressCity & "510000,510500,����,28.889138,105.443348;"
    AddressCity = AddressCity & "510000,510600,����,31.127991,104.398651;"
    AddressCity = AddressCity & "510000,510700,����,31.46402,104.741722;"
    AddressCity = AddressCity & "510000,510800,��Ԫ,32.433668,105.829757;"
    AddressCity = AddressCity & "510000,510900,����,30.513311,105.571331;"
    AddressCity = AddressCity & "510000,511000,�ڽ�,29.58708,105.066138;"
    AddressCity = AddressCity & "510000,511100,��ɽ,29.582024,103.761263;"
    AddressCity = AddressCity & "510000,511300,�ϳ�,30.795281,106.082974;"
    AddressCity = AddressCity & "510000,511400,üɽ,30.048318,103.831788;"
    AddressCity = AddressCity & "510000,511500,�˱�,28.760189,104.630825;"
    AddressCity = AddressCity & "510000,511600,�㰲,30.456398,106.633369;"
    AddressCity = AddressCity & "510000,511700,����,31.209484,107.502262;"
    AddressCity = AddressCity & "510000,511800,�Ű�,29.987722,103.001033;"
    AddressCity = AddressCity & "510000,511900,����,31.858809,106.753669;"
    AddressCity = AddressCity & "510000,512000,����,30.122211,104.641917;"
    AddressCity = AddressCity & "510000,513200,����,31.899792,102.221374;"
    AddressCity = AddressCity & "510000,513300,����,30.050663,101.963815;"
    AddressCity = AddressCity & "510000,513400,��ɽ,27.886762,102.258746;"
    AddressCity = AddressCity & "520000,520100,����,26.578343,106.713478;"
    AddressCity = AddressCity & "520000,520200,����ˮ,26.584643,104.846743;"
    AddressCity = AddressCity & "520000,520300,����,27.706626,106.937265;"
    AddressCity = AddressCity & "520000,520400,��˳,26.245544,105.932188;"
    AddressCity = AddressCity & "520000,520600,ͭ��,27.718346,109.191555;"
    AddressCity = AddressCity & "520000,522300,ǭ����,25.08812,104.897971;"
    AddressCity = AddressCity & "520000,520500,�Ͻ�,27.301693,105.28501;"
    AddressCity = AddressCity & "520000,522600,ǭ����,26.583352,107.977488;"
    AddressCity = AddressCity & "520000,522700,ǭ��,26.258219,107.517156;"
    AddressCity = AddressCity & "530000,530100,����,25.040609,102.712251;"
    AddressCity = AddressCity & "530000,530300,����,25.501557,103.797851;"
    AddressCity = AddressCity & "530000,530400,��Ϫ,24.350461,102.543907;"
    AddressCity = AddressCity & "530000,530500,��ɽ,25.111802,99.167133;"
    AddressCity = AddressCity & "530000,530600,��ͨ,27.336999,103.717216;"
    AddressCity = AddressCity & "530000,530700,����,26.872108,100.233026;"
    AddressCity = AddressCity & "530000,530800,�ն�,22.777321,100.972344;"
    AddressCity = AddressCity & "530000,530900,�ٲ�,23.886567,100.08697;"
    AddressCity = AddressCity & "530000,532300,����,25.041988,101.546046;"
    AddressCity = AddressCity & "530000,532500,���,23.366775,103.384182;"
    AddressCity = AddressCity & "530000,532600,��ɽ,23.36951,104.24401;"
    AddressCity = AddressCity & "530000,532800,��˫����,22.001724,100.797941;"
    AddressCity = AddressCity & "530000,532900,����,25.589449,100.225668;"
    AddressCity = AddressCity & "530000,533100,�º�,24.436694,98.578363;"
    AddressCity = AddressCity & "530000,533300,ŭ��,25.850949,98.854304;"
    AddressCity = AddressCity & "530000,533400,����,27.826853,99.706463;"
    AddressCity = AddressCity & "540000,540100,����,29.660361,91.132212;"
    AddressCity = AddressCity & "540000,540300,����,31.136875,97.178452;"
    AddressCity = AddressCity & "540000,540500,ɽ��,29.236023,91.766529;"
    AddressCity = AddressCity & "540000,540200,�տ���,29.267519,88.885148;"
    AddressCity = AddressCity & "540000,540600,����,31.476004,92.060214;"
    AddressCity = AddressCity & "540000,542500,����,32.503187,80.105498;"
    AddressCity = AddressCity & "540000,540400,��֥,29.654693,94.362348;"
    AddressCity = AddressCity & "610000,610100,����,34.263161,108.948024;"
    AddressCity = AddressCity & "610000,610200,ͭ��,34.916582,108.979608;"
    AddressCity = AddressCity & "610000,610300,����,34.369315,107.14487;"
    AddressCity = AddressCity & "610000,610400,����,34.333439,108.705117;"
    AddressCity = AddressCity & "610000,610500,μ��,34.499381,109.502882;"
    AddressCity = AddressCity & "610000,610600,�Ӱ�,36.596537,109.49081;"
    AddressCity = AddressCity & "610000,610700,����,33.077668,107.028621;"
    AddressCity = AddressCity & "610000,610800,����,38.290162,109.741193;"
    AddressCity = AddressCity & "610000,610900,����,32.6903,109.029273;"
    AddressCity = AddressCity & "610000,611000,����,33.868319,109.939776;"
    AddressCity = AddressCity & "620000,620100,����,36.058039,103.823557;"
    AddressCity = AddressCity & "620000,620200,������,39.786529,98.277304;"
    AddressCity = AddressCity & "620000,620300,���,38.514238,102.187888;"
    AddressCity = AddressCity & "620000,620400,����,36.54568,104.173606;"
    AddressCity = AddressCity & "620000,620500,��ˮ,34.578529,105.724998;"
    AddressCity = AddressCity & "620000,620600,����,37.929996,102.634697;"
    AddressCity = AddressCity & "620000,620700,��Ҵ,38.932897,100.455472;"
    AddressCity = AddressCity & "620000,620800,ƽ��,35.54279,106.684691;"
    AddressCity = AddressCity & "620000,620900,��Ȫ,39.744023,98.510795;"
    AddressCity = AddressCity & "620000,621000,����,35.734218,107.638372;"
    AddressCity = AddressCity & "620000,621100,����,35.579578,104.626294;"
    AddressCity = AddressCity & "620000,621200,¤��,33.388598,104.929379;"
    AddressCity = AddressCity & "620000,622900,����,35.599446,103.212006;"
    AddressCity = AddressCity & "620000,623000,����,34.986354,102.911008;"
    AddressCity = AddressCity & "630000,630100,����,36.623178,101.778916;"
    AddressCity = AddressCity & "630000,630200,����,36.502916,102.10327;"
    AddressCity = AddressCity & "630000,632200,����,36.959435,100.901059;"
    AddressCity = AddressCity & "630000,632300,����,35.517744,102.019988;"
    AddressCity = AddressCity & "630000,632500,���ϲ���,36.280353,100.619542;"
    AddressCity = AddressCity & "630000,632600,����,34.4736,100.242143;"
    AddressCity = AddressCity & "630000,632700,����,33.004049,97.008522;"
    AddressCity = AddressCity & "630000,632800,����,37.374663,97.370785;"
    AddressCity = AddressCity & "640000,640100,����,38.46637,106.278179;"
    AddressCity = AddressCity & "640000,640200,ʯ��ɽ,39.01333,106.376173;"
    AddressCity = AddressCity & "640000,640300,����,37.986165,106.199409;"
    AddressCity = AddressCity & "640000,640400,��ԭ,36.004561,106.285241;"
    AddressCity = AddressCity & "640000,640500,����,37.514951,105.189568;"
    AddressCity = AddressCity & "650000,650100,��³ľ��,43.792818,87.617733;"
    AddressCity = AddressCity & "650000,650200,��������,45.595886,84.873946;"
    AddressCity = AddressCity & "650000,650400,��³��,42.947613,89.184078;"
    AddressCity = AddressCity & "650000,650500,����,42.833248,93.51316;"
    AddressCity = AddressCity & "650000,652300,����,44.014577,87.304012;"
    AddressCity = AddressCity & "650000,652700,��������,44.903258,82.074778;"
    AddressCity = AddressCity & "650000,652800,��������,41.768552,86.150969;"
    AddressCity = AddressCity & "650000,652900,������,41.170712,80.265068;"
    AddressCity = AddressCity & "650000,653000,�������տ¶�����,39.713431,76.172825;"
    AddressCity = AddressCity & "650000,653100,��ʲ,39.467664,75.989138;"
    AddressCity = AddressCity & "650000,653200,����,37.110687,79.92533;"
    AddressCity = AddressCity & "650000,654000,����,43.92186,81.317946;"
    AddressCity = AddressCity & "650000,654200,����,46.746301,82.985732;"
    AddressCity = AddressCity & "650000,654300,����̩,47.848393,88.13963;"
    AddressCity = AddressCity & "650000,659001,ʯ����,44.305886,86.041075;"
    AddressCity = AddressCity & "650000,659002,������,40.541914,81.285884;"
    AddressCity = AddressCity & "650000,659003,ͼľ���,39.867316,79.077978;"
    AddressCity = AddressCity & "650000,659004,�����,44.167401,87.526884;"
    '�����ؼ��к�����Ĳ㼶����
    AddressCity = AddressCity & "650000,659005,����,47.353177,87.824932;"
    AddressCity = AddressCity & "650000,659006,���Ź�,41.827251,85.501218;"
    AddressCity = AddressCity & "650000,659007,˫��,44.840524,82.353656;"
    AddressCity = AddressCity & "650000,659008,�ɿ˴���,43.6832,80.63579;"
    AddressCity = AddressCity & "650000,659009,����,37.207994,79.287372;"
    AddressCity = AddressCity & "650000,659010,�����,44.69288853,84.8275959;"
    AddressCity = AddressCity & "710000,710000,̨��,25.044332,121.509062;"
    AddressCity = AddressCity & "810000,810000,���,22.320048,114.173355;"
    AddressCity = AddressCity & "820000,820000,����,22.198951,113.54909"

End Function

Public Function AddressDistrict() As String
' �����������ݣ��������ƣ�����ȡ�
    '����ID    ����ID  ����    γ��    ����
    AddressDistrict = "110000,110101,������,39.917544,116.418757;"
    AddressDistrict = AddressDistrict & "110000,110102,������,39.915309,116.366794;"
    AddressDistrict = AddressDistrict & "110000,110105,������,39.921489,116.486409;"
    AddressDistrict = AddressDistrict & "110000,110106,��̨��,39.863642,116.286968;"
    AddressDistrict = AddressDistrict & "110000,110107,ʯ��ɽ��,39.914601,116.195445;"
    AddressDistrict = AddressDistrict & "110000,110108,������,39.956074,116.310316;"
    AddressDistrict = AddressDistrict & "110000,110109,��ͷ����,39.937183,116.105381;"
    AddressDistrict = AddressDistrict & "110000,110111,��ɽ��,39.735535,116.139157;"
    AddressDistrict = AddressDistrict & "110000,110112,ͨ����,39.902486,116.658603;"
    AddressDistrict = AddressDistrict & "110000,110113,˳����,40.128936,116.653525;"
    AddressDistrict = AddressDistrict & "110000,110114,��ƽ��,40.218085,116.235906;"
    AddressDistrict = AddressDistrict & "110000,110115,������,39.728908,116.338033;"
    AddressDistrict = AddressDistrict & "110000,110116,������,40.324272,116.637122;"
    AddressDistrict = AddressDistrict & "110000,110117,ƽ����,40.144783,117.112335;"
    AddressDistrict = AddressDistrict & "110000,110118,������,40.377362,116.843352;"
    AddressDistrict = AddressDistrict & "110000,110119,������,40.465325,115.985006;"
    AddressDistrict = AddressDistrict & "120000,120101,��ƽ��,39.118327,117.195907;"
    AddressDistrict = AddressDistrict & "120000,120102,�Ӷ���,39.122125,117.226568;"
    AddressDistrict = AddressDistrict & "120000,120103,������,39.101897,117.217536;"
    AddressDistrict = AddressDistrict & "120000,120104,�Ͽ���,39.120474,117.164143;"
    AddressDistrict = AddressDistrict & "120000,120105,�ӱ���,39.156632,117.201569;"
    AddressDistrict = AddressDistrict & "120000,120106,������,39.175066,117.163301;"
    AddressDistrict = AddressDistrict & "120000,120110,������,39.087764,117.313967;"
    AddressDistrict = AddressDistrict & "120000,120111,������,39.139446,117.012247;"
    AddressDistrict = AddressDistrict & "120000,120112,������,38.989577,117.382549;"
    AddressDistrict = AddressDistrict & "120000,120113,������,39.225555,117.13482;"
    AddressDistrict = AddressDistrict & "120000,120114,������,39.376925,117.057959;"
    AddressDistrict = AddressDistrict & "120000,120115,������,39.716965,117.308094;"
    AddressDistrict = AddressDistrict & "120000,120116,��������,39.032846,117.654173;"
    AddressDistrict = AddressDistrict & "120000,120117,������,39.328886,117.82828;"
    AddressDistrict = AddressDistrict & "120000,120118,������,38.935671,116.925304;"
    AddressDistrict = AddressDistrict & "120000,120119,������,40.045342,117.407449;"
    AddressDistrict = AddressDistrict & "310000,310101,������,31.222771,121.490317;"
    AddressDistrict = AddressDistrict & "310000,310104,�����,31.179973,121.43752;"
    AddressDistrict = AddressDistrict & "310000,310105,������,31.218123,121.4222;"
    AddressDistrict = AddressDistrict & "310000,310106,������,31.229003,121.448224;"
    AddressDistrict = AddressDistrict & "310000,310107,������,31.241701,121.392499;"
    AddressDistrict = AddressDistrict & "310000,310109,�����,31.26097,121.491832;"
    AddressDistrict = AddressDistrict & "310000,310110,������,31.270755,121.522797;"
    AddressDistrict = AddressDistrict & "310000,310112,������,31.111658,121.375972;"
    AddressDistrict = AddressDistrict & "310000,310113,��ɽ��,31.398896,121.489934;"
    AddressDistrict = AddressDistrict & "310000,310114,�ζ���,31.383524,121.250333;"
    AddressDistrict = AddressDistrict & "310000,310115,�ֶ�����,31.245944,121.567706;"
    AddressDistrict = AddressDistrict & "310000,310116,��ɽ��,30.724697,121.330736;"
    AddressDistrict = AddressDistrict & "310000,310117,�ɽ���,31.03047,121.223543;"
    AddressDistrict = AddressDistrict & "310000,310118,������,31.151209,121.113021;"
    AddressDistrict = AddressDistrict & "310000,310120,������,30.912345,121.458472;"
    AddressDistrict = AddressDistrict & "310000,310151,������,31.626946,121.397516;"
    AddressDistrict = AddressDistrict & "500000,500101,������,30.807807,108.380246;"
    AddressDistrict = AddressDistrict & "500000,500102,������,29.703652,107.394905;"
    AddressDistrict = AddressDistrict & "500000,500103,������,29.556742,106.56288;"
    AddressDistrict = AddressDistrict & "500000,500104,��ɿ���,29.481002,106.48613;"
    AddressDistrict = AddressDistrict & "500000,500105,������,29.575352,106.532844;"
    AddressDistrict = AddressDistrict & "500000,500106,ɳƺ����,29.541224,106.4542;"
    AddressDistrict = AddressDistrict & "500000,500107,��������,29.523492,106.480989;"
    AddressDistrict = AddressDistrict & "500000,500108,�ϰ���,29.523992,106.560813;"
    AddressDistrict = AddressDistrict & "500000,500109,������,29.82543,106.437868;"
    AddressDistrict = AddressDistrict & "500000,500110,�뽭��,29.028091,106.651417;"
    AddressDistrict = AddressDistrict & "500000,500111,������,29.700498,105.715319;"
    AddressDistrict = AddressDistrict & "500000,500112,�山��,29.601451,106.512851;"
    AddressDistrict = AddressDistrict & "500000,500113,������,29.381919,106.519423;"
    AddressDistrict = AddressDistrict & "500000,500114,ǭ����,29.527548,108.782577;"
    AddressDistrict = AddressDistrict & "500000,500115,������,29.833671,107.074854;"
    AddressDistrict = AddressDistrict & "500000,500116,������,29.283387,106.253156;"
    AddressDistrict = AddressDistrict & "500000,500117,�ϴ���,29.990993,106.265554;"
    AddressDistrict = AddressDistrict & "500000,500118,������,29.348748,105.894714;"
    AddressDistrict = AddressDistrict & "500000,500119,�ϴ���,29.156646,107.098153;"
    AddressDistrict = AddressDistrict & "500000,500120,�ɽ��,29.593581,106.231126;"
    AddressDistrict = AddressDistrict & "500000,500151,ͭ����,29.839944,106.054948;"
    AddressDistrict = AddressDistrict & "500000,500152,������,30.189554,105.841818;"
    AddressDistrict = AddressDistrict & "500000,500153,�ٲ���,29.403627,105.594061;"
    AddressDistrict = AddressDistrict & "500000,500154,������,31.167735,108.413317;"
    AddressDistrict = AddressDistrict & "500000,500155,��ƽ��,30.672168,107.800034;"
    AddressDistrict = AddressDistrict & "500000,500156,��¡��,29.32376,107.75655;"
    AddressDistrict = AddressDistrict & "500000,500229,�ǿ���,31.946293,108.6649;"
    AddressDistrict = AddressDistrict & "500000,500230,�ᶼ��,29.866424,107.73248;"
    AddressDistrict = AddressDistrict & "500000,500231,�潭��,30.330012,107.348692;"
    AddressDistrict = AddressDistrict & "500000,500233,����,30.291537,108.037518;"
    AddressDistrict = AddressDistrict & "500000,500235,������,30.930529,108.697698;"
    AddressDistrict = AddressDistrict & "500000,500236,�����,31.019967,109.465774;"
    AddressDistrict = AddressDistrict & "500000,500237,��ɽ��,31.074843,109.878928;"
    AddressDistrict = AddressDistrict & "500000,500238,��Ϫ��,31.3966,109.628912;"
    AddressDistrict = AddressDistrict & "500000,500240,ʯ��������������,29.99853,108.112448;"
    AddressDistrict = AddressDistrict & "500000,500241,��ɽ����������������,28.444772,108.996043;"
    AddressDistrict = AddressDistrict & "500000,500242,��������������������,28.839828,108.767201;"
    AddressDistrict = AddressDistrict & "500000,500243,��ˮ����������������,29.293856,108.166551;"
    AddressDistrict = AddressDistrict & "810000,810001,������,22.28198083,114.1543731;"
    AddressDistrict = AddressDistrict & "810000,810002,������,22.27638889,114.1829153;"
    AddressDistrict = AddressDistrict & "810000,810003,����,22.27969306,114.2260031;"
    AddressDistrict = AddressDistrict & "810000,810004,����,22.24589667,114.1600117;"
    AddressDistrict = AddressDistrict & "810000,810005,�ͼ�����,22.31170389,114.1733317;"
    AddressDistrict = AddressDistrict & "810000,810006,��ˮ����,22.33385417,114.1632417;"
    AddressDistrict = AddressDistrict & "810000,810007,��������,22.31251,114.1928467;"
    AddressDistrict = AddressDistrict & "810000,810008,�ƴ�����,22.33632056,114.2038856;"
    AddressDistrict = AddressDistrict & "810000,810009,������,22.32083778,114.2140542;"
    AddressDistrict = AddressDistrict & "810000,810010,������,22.36830667,114.1210792;"
    AddressDistrict = AddressDistrict & "810000,810011,������,22.39384417,113.9765742;"
    AddressDistrict = AddressDistrict & "810000,810012,Ԫ����,22.44142833,114.0324381;"
    AddressDistrict = AddressDistrict & "810000,810013,����,22.49610389,114.1473639;"
    AddressDistrict = AddressDistrict & "810000,810014,������,22.44565306,114.1717431;"
    AddressDistrict = AddressDistrict & "810000,810015,������,22.31421306,114.264645;"
    AddressDistrict = AddressDistrict & "810000,810016,ɳ����,22.37953167,114.1953653;"
    AddressDistrict = AddressDistrict & "810000,810017,������,22.36387667,114.1393194;"
    AddressDistrict = AddressDistrict & "810000,810018,�뵺��,22.28640778,113.94612;"
    AddressDistrict = AddressDistrict & "820000,820001,����������,22.20787,113.5528956;"
    AddressDistrict = AddressDistrict & "820000,820002,��������,22.1992075,113.5489608;"
    AddressDistrict = AddressDistrict & "820000,820003,��������,22.19372083,113.5501828;"
    AddressDistrict = AddressDistrict & "820000,820004,������,22.18853944,113.5536475;"
    AddressDistrict = AddressDistrict & "820000,820005,��˳����,22.18736806,113.5419278;"
    AddressDistrict = AddressDistrict & "820000,820006,��ģ����,22.15375944,113.5587044;"
    AddressDistrict = AddressDistrict & "820000,820007,·�����,22.13663,113.5695992;"
    AddressDistrict = AddressDistrict & "820000,820008,ʥ���ø�����,22.12348639,113.5599542;"
    AddressDistrict = AddressDistrict & "130100,130102,������,38.047501,114.548151;"
    AddressDistrict = AddressDistrict & "130100,130104,������,38.028383,114.462931;"
    AddressDistrict = AddressDistrict & "130100,130105,�»���,38.067142,114.465974;"
    AddressDistrict = AddressDistrict & "130100,130107,�������,38.069748,114.058178;"
    AddressDistrict = AddressDistrict & "130100,130108,ԣ����,38.027696,114.533257;"
    AddressDistrict = AddressDistrict & "130100,130109,޻����,38.033767,114.849647;"
    AddressDistrict = AddressDistrict & "130100,130110,¹Ȫ��,38.093994,114.321023;"
    AddressDistrict = AddressDistrict & "130100,130111,�����,37.886911,114.654281;"
    AddressDistrict = AddressDistrict & "130100,130121,������,38.033614,114.144488;"
    AddressDistrict = AddressDistrict & "130100,130123,������,38.147835,114.569887;"
    AddressDistrict = AddressDistrict & "130100,130125,������,38.437422,114.552734;"
    AddressDistrict = AddressDistrict & "130100,130126,������,38.306546,114.37946;"
    AddressDistrict = AddressDistrict & "130100,130127,������,37.605714,114.610699;"
    AddressDistrict = AddressDistrict & "130100,130128,������,38.18454,115.200207;"
    AddressDistrict = AddressDistrict & "130100,130129,�޻���,37.660199,114.387756;"
    AddressDistrict = AddressDistrict & "130100,130130,�޼���,38.176376,114.977845;"
    AddressDistrict = AddressDistrict & "130100,130131,ƽɽ��,38.259311,114.184144;"
    AddressDistrict = AddressDistrict & "130100,130132,Ԫ����,37.762514,114.52618;"
    AddressDistrict = AddressDistrict & "130100,130133,����,37.754341,114.775362;"
    AddressDistrict = AddressDistrict & "130100,130181,������,37.92904,115.217451;"
    AddressDistrict = AddressDistrict & "130100,130183,������,38.027478,115.044886;"
    AddressDistrict = AddressDistrict & "130100,130184,������,38.344768,114.68578;"
    AddressDistrict = AddressDistrict & "130200,130202,·����,39.615162,118.210821;"
    AddressDistrict = AddressDistrict & "130200,130203,·����,39.628538,118.174736;"
    AddressDistrict = AddressDistrict & "130200,130204,��ұ��,39.715736,118.45429;"
    AddressDistrict = AddressDistrict & "130200,130205,��ƽ��,39.676171,118.264425;"
    AddressDistrict = AddressDistrict & "130200,130207,������,39.56303,118.110793;"
    AddressDistrict = AddressDistrict & "130200,130208,������,39.831363,118.155779;"
    AddressDistrict = AddressDistrict & "130200,130209,��������,39.278277,118.446585;"
    AddressDistrict = AddressDistrict & "130200,130224,������,39.506201,118.681552;"
    AddressDistrict = AddressDistrict & "130200,130225,��ͤ��,39.42813,118.905341;"
    AddressDistrict = AddressDistrict & "130200,130227,Ǩ����,40.146238,118.305139;"
    AddressDistrict = AddressDistrict & "130200,130229,������,39.887323,117.753665;"
    AddressDistrict = AddressDistrict & "130200,130281,����,40.188616,117.965875;"
    AddressDistrict = AddressDistrict & "130200,130283,Ǩ����,40.012108,118.701933;"
    AddressDistrict = AddressDistrict & "130200,130284,������,39.74485,118.699546;"
    AddressDistrict = AddressDistrict & "130300,130302,������,39.943458,119.596224;"
    AddressDistrict = AddressDistrict & "130300,130303,ɽ������,39.998023,119.753591;"
    AddressDistrict = AddressDistrict & "130300,130304,��������,39.825121,119.486286;"
    AddressDistrict = AddressDistrict & "130300,130306,������,39.887053,119.240651;"
    AddressDistrict = AddressDistrict & "130300,130321,��������������,40.406023,118.954555;"
    AddressDistrict = AddressDistrict & "130300,130322,������,39.709729,119.164541;"
    AddressDistrict = AddressDistrict & "130300,130324,¬����,39.891639,118.881809;"
    AddressDistrict = AddressDistrict & "130400,130402,��ɽ��,36.603196,114.484989;"
    AddressDistrict = AddressDistrict & "130400,130403,��̨��,36.611082,114.494703;"
    AddressDistrict = AddressDistrict & "130400,130404,������,36.615484,114.458242;"
    AddressDistrict = AddressDistrict & "130400,130406,������,36.420487,114.209936;"
    AddressDistrict = AddressDistrict & "130400,130407,������,36.555778,114.805154;"
    AddressDistrict = AddressDistrict & "130400,130408,������,36.776413,114.496162;"
    AddressDistrict = AddressDistrict & "130400,130423,������,36.337604,114.610703;"
    AddressDistrict = AddressDistrict & "130400,130424,�ɰ���,36.443832,114.680356;"
    AddressDistrict = AddressDistrict & "130400,130425,������,36.283316,115.152586;"
    AddressDistrict = AddressDistrict & "130400,130426,����,36.563143,113.673297;"
    AddressDistrict = AddressDistrict & "130400,130427,����,36.367673,114.38208;"
    AddressDistrict = AddressDistrict & "130400,130430,����,36.81325,115.168584;"
    AddressDistrict = AddressDistrict & "130400,130431,������,36.914908,114.878517;"
    AddressDistrict = AddressDistrict & "130400,130432,��ƽ��,36.483603,114.950859;"
    AddressDistrict = AddressDistrict & "130400,130433,������,36.539461,115.289057;"
    AddressDistrict = AddressDistrict & "130400,130434,κ��,36.354248,114.93411;"
    AddressDistrict = AddressDistrict & "130400,130435,������,36.773398,114.957588;"
    AddressDistrict = AddressDistrict & "130400,130481,�䰲��,36.696115,114.194581;"
    AddressDistrict = AddressDistrict & "130500,130502,�嶼��,37.064125,114.507131;"
    AddressDistrict = AddressDistrict & "130500,130503,�Ŷ���,37.068009,114.473687;"
    AddressDistrict = AddressDistrict & "130500,130505,������,37.129952,114.684469;"
    AddressDistrict = AddressDistrict & "130500,130506,�Ϻ���,37.003812,114.691377;"
    AddressDistrict = AddressDistrict & "130500,130522,�ٳ���,37.444009,114.506873;"
    AddressDistrict = AddressDistrict & "130500,130523,������,37.287663,114.511523;"
    AddressDistrict = AddressDistrict & "130500,130524,������,37.483596,114.693382;"
    AddressDistrict = AddressDistrict & "130500,130525,¡Ң��,37.350925,114.776348;"
    AddressDistrict = AddressDistrict & "130500,130528,������,37.618956,114.921027;"
    AddressDistrict = AddressDistrict & "130500,130529,��¹��,37.21768,115.038782;"
    AddressDistrict = AddressDistrict & "130500,130530,�º���,37.526216,115.247537;"
    AddressDistrict = AddressDistrict & "130500,130531,������,37.075548,115.142797;"
    AddressDistrict = AddressDistrict & "130500,130532,ƽ����,37.069404,115.029218;"
    AddressDistrict = AddressDistrict & "130500,130533,����,36.983272,115.272749;"
    AddressDistrict = AddressDistrict & "130500,130534,�����,37.059991,115.668999;"
    AddressDistrict = AddressDistrict & "130500,130535,������,36.8642,115.498684;"
    AddressDistrict = AddressDistrict & "130500,130581,�Ϲ���,37.359668,115.398102;"
    AddressDistrict = AddressDistrict & "130500,130582,ɳ����,36.861903,114.504902;"
    AddressDistrict = AddressDistrict & "130600,130602,������,38.88662,115.470659;"
    AddressDistrict = AddressDistrict & "130600,130606,������,38.865005,115.500934;"
    AddressDistrict = AddressDistrict & "130600,130607,������,38.95138,115.32442;"
    AddressDistrict = AddressDistrict & "130600,130608,��Է��,38.771012,115.492221;"
    AddressDistrict = AddressDistrict & "130600,130609,��ˮ��,39.020395,115.64941;"
    AddressDistrict = AddressDistrict & "130600,130623,�ˮ��,39.393148,115.711985;"
    AddressDistrict = AddressDistrict & "130600,130624,��ƽ��,38.847276,114.198801;"
    AddressDistrict = AddressDistrict & "130600,130626,������,39.266195,115.796895;"
    AddressDistrict = AddressDistrict & "130600,130627,����,38.748542,114.981241;"
    AddressDistrict = AddressDistrict & "130600,130628,������,38.690092,115.778878;"
    AddressDistrict = AddressDistrict & "130600,130629,�ݳ���,39.05282,115.866247;"
    AddressDistrict = AddressDistrict & "130600,130630,�Դ��,39.35755,114.692567;"
    AddressDistrict = AddressDistrict & "130600,130631,������,38.707448,115.154009;"
    AddressDistrict = AddressDistrict & "130600,130632,������,38.929912,115.931979;"
    AddressDistrict = AddressDistrict & "130600,130633,����,39.35297,115.501146;"
    AddressDistrict = AddressDistrict & "130600,130634,������,38.619992,114.704055;"
    AddressDistrict = AddressDistrict & "130600,130635,���,38.496429,115.583631;"
    AddressDistrict = AddressDistrict & "130600,130636,˳ƽ��,38.845127,115.132749;"
    AddressDistrict = AddressDistrict & "130600,130637,��Ұ��,38.458271,115.461798;"
    AddressDistrict = AddressDistrict & "130600,130638,����,38.990819,116.107474;"
    AddressDistrict = AddressDistrict & "130600,130681,������,39.485765,115.973409;"
    AddressDistrict = AddressDistrict & "130600,130682,������,38.517602,114.991389;"
    AddressDistrict = AddressDistrict & "130600,130683,������,38.421367,115.33141;"
    AddressDistrict = AddressDistrict & "130600,130684,�߱�����,39.327689,115.882704;"
    AddressDistrict = AddressDistrict & "130700,130702,�Ŷ���,40.813875,114.885658;"
    AddressDistrict = AddressDistrict & "130700,130703,������,40.824385,114.882127;"
    AddressDistrict = AddressDistrict & "130700,130705,������,40.609368,115.0632;"
    AddressDistrict = AddressDistrict & "130700,130706,�»�԰��,40.488645,115.281002;"
    AddressDistrict = AddressDistrict & "130700,130708,��ȫ��,40.765136,114.736131;"
    AddressDistrict = AddressDistrict & "130700,130709,������,40.971302,115.281652;"
    AddressDistrict = AddressDistrict & "130700,130722,�ű���,41.151713,114.715951;"
    AddressDistrict = AddressDistrict & "130700,130723,������,41.850046,114.615809;"
    AddressDistrict = AddressDistrict & "130700,130724,��Դ��,41.667419,115.684836;"
    AddressDistrict = AddressDistrict & "130700,130725,������,41.080091,113.977713;"
    AddressDistrict = AddressDistrict & "130700,130726,ε��,39.837181,114.582695;"
    AddressDistrict = AddressDistrict & "130700,130727,��ԭ��,40.113419,114.167343;"
    AddressDistrict = AddressDistrict & "130700,130728,������,40.671274,114.422364;"
    AddressDistrict = AddressDistrict & "130700,130730,������,40.405405,115.520846;"
    AddressDistrict = AddressDistrict & "130700,130731,��¹��,40.378701,115.219246;"
    AddressDistrict = AddressDistrict & "130700,130732,�����,40.912081,115.832708;"
    AddressDistrict = AddressDistrict & "130800,130802,˫����,40.976204,117.939152;"
    AddressDistrict = AddressDistrict & "130800,130803,˫����,40.959756,117.797485;"
    AddressDistrict = AddressDistrict & "130800,130804,ӥ��Ӫ�ӿ���,40.546956,117.661154;"
    AddressDistrict = AddressDistrict & "130800,130821,�е���,40.768637,118.172496;"
    AddressDistrict = AddressDistrict & "130800,130822,��¡��,40.418525,117.507098;"
    AddressDistrict = AddressDistrict & "130800,130824,��ƽ��,40.936644,117.337124;"
    AddressDistrict = AddressDistrict & "130800,130825,¡����,41.316667,117.736343;"
    AddressDistrict = AddressDistrict & "130800,130826,��������������,41.209903,116.65121;"
    AddressDistrict = AddressDistrict & "130800,130827,�������������,40.607981,118.488642;"
    AddressDistrict = AddressDistrict & "130800,130828,Χ�������ɹ���������,41.949404,117.764086;"
    AddressDistrict = AddressDistrict & "130800,130881,ƽȪ��,41.00561,118.690238;"
    AddressDistrict = AddressDistrict & "130900,130902,�»���,38.308273,116.873049;"
    AddressDistrict = AddressDistrict & "130900,130903,�˺���,38.307405,116.840063;"
    AddressDistrict = AddressDistrict & "130900,130921,����,38.219856,117.007478;"
    AddressDistrict = AddressDistrict & "130900,130922,����,38.569646,116.838384;"
    AddressDistrict = AddressDistrict & "130900,130923,������,37.88655,116.542062;"
    AddressDistrict = AddressDistrict & "130900,130924,������,38.141582,117.496606;"
    AddressDistrict = AddressDistrict & "130900,130925,��ɽ��,38.056141,117.229814;"
    AddressDistrict = AddressDistrict & "130900,130926,������,38.427102,115.835856;"
    AddressDistrict = AddressDistrict & "130900,130927,��Ƥ��,38.042439,116.709171;"
    AddressDistrict = AddressDistrict & "130900,130928,������,37.628182,116.391512;"
    AddressDistrict = AddressDistrict & "130900,130929,����,38.189661,116.123844;"
    AddressDistrict = AddressDistrict & "130900,130930,�ϴ����������,38.057953,117.105104;"
    AddressDistrict = AddressDistrict & "130900,130981,��ͷ��,38.073479,116.570163;"
    AddressDistrict = AddressDistrict & "130900,130982,������,38.706513,116.106764;"
    AddressDistrict = AddressDistrict & "130900,130983,������,38.369238,117.343803;"
    AddressDistrict = AddressDistrict & "130900,130984,�Ӽ���,38.44149,116.089452;"
    AddressDistrict = AddressDistrict & "131000,131002,������,39.502569,116.694544;"
    AddressDistrict = AddressDistrict & "131000,131003,������,39.521931,116.713708;"
    AddressDistrict = AddressDistrict & "131000,131022,�̰���,39.436468,116.299894;"
    AddressDistrict = AddressDistrict & "131000,131023,������,39.319717,116.498089;"
    AddressDistrict = AddressDistrict & "131000,131024,�����,39.757212,117.007161;"
    AddressDistrict = AddressDistrict & "131000,131025,�����,38.699215,116.640735;"
    AddressDistrict = AddressDistrict & "131000,131026,�İ���,38.866801,116.460107;"
    AddressDistrict = AddressDistrict & "131000,131028,�󳧻���������,39.889266,116.986501;"
    AddressDistrict = AddressDistrict & "131000,131081,������,39.117331,116.392021;"
    AddressDistrict = AddressDistrict & "131000,131082,������,39.982778,117.077018;"
    AddressDistrict = AddressDistrict & "131100,131102,�ҳ���,37.732237,115.694945;"
    AddressDistrict = AddressDistrict & "131100,131103,������,37.542788,115.579173;"
    AddressDistrict = AddressDistrict & "131100,131121,��ǿ��,37.511512,115.726499;"
    AddressDistrict = AddressDistrict & "131100,131122,������,37.803774,115.892415;"
    AddressDistrict = AddressDistrict & "131100,131123,��ǿ��,38.03698,115.970236;"
    AddressDistrict = AddressDistrict & "131100,131124,������,38.232671,115.726577;"
    AddressDistrict = AddressDistrict & "131100,131125,��ƽ��,38.233511,115.519627;"
    AddressDistrict = AddressDistrict & "131100,131126,�ʳ���,37.350981,115.966747;"
    AddressDistrict = AddressDistrict & "131100,131127,����,37.686622,116.258446;"
    AddressDistrict = AddressDistrict & "131100,131128,������,37.869945,116.164727;"
    AddressDistrict = AddressDistrict & "131100,131182,������,38.00347,115.554596;"
    AddressDistrict = AddressDistrict & "140100,140105,С����,37.817974,112.564273;"
    AddressDistrict = AddressDistrict & "140100,140106,ӭ����,37.855804,112.558851;"
    AddressDistrict = AddressDistrict & "140100,140107,�ӻ�����,37.879291,112.560743;"
    AddressDistrict = AddressDistrict & "140100,140108,���ƺ��,37.939893,112.487122;"
    AddressDistrict = AddressDistrict & "140100,140109,�������,37.862653,112.522258;"
    AddressDistrict = AddressDistrict & "140100,140110,��Դ��,37.715619,112.477849;"
    AddressDistrict = AddressDistrict & "140100,140121,������,37.60729,112.357961;"
    AddressDistrict = AddressDistrict & "140100,140122,������,38.058797,112.673818;"
    AddressDistrict = AddressDistrict & "140100,140123,¦����,38.066035,111.793798;"
    AddressDistrict = AddressDistrict & "140100,140181,�Ž���,37.908534,112.174353;"
    AddressDistrict = AddressDistrict & "140200,140212,������,40.258269,113.141044;"
    AddressDistrict = AddressDistrict & "140200,140213,ƽ����,40.075667,113.298027;"
    AddressDistrict = AddressDistrict & "140200,140214,�Ƹ���,40.005405,113.149693;"
    AddressDistrict = AddressDistrict & "140200,140215,������,40.040295,113.61244;"
    AddressDistrict = AddressDistrict & "140200,140221,������,40.364927,113.749871;"
    AddressDistrict = AddressDistrict & "140200,140222,������,40.421336,114.09112;"
    AddressDistrict = AddressDistrict & "140200,140223,������,39.763051,114.279252;"
    AddressDistrict = AddressDistrict & "140200,140224,������,39.438867,114.23576;"
    AddressDistrict = AddressDistrict & "140200,140225,��Դ��,39.699099,113.698091;"
    AddressDistrict = AddressDistrict & "140200,140226,������,40.012873,112.70641;"
    AddressDistrict = AddressDistrict & "140300,140302,����,37.860938,113.586513;"
    AddressDistrict = AddressDistrict & "140300,140303,����,37.870085,113.559066;"
    AddressDistrict = AddressDistrict & "140300,140311,����,37.94096,113.58664;"
    AddressDistrict = AddressDistrict & "140300,140321,ƽ����,37.800289,113.631049;"
    AddressDistrict = AddressDistrict & "140300,140322,����,38.086131,113.41223;"
    AddressDistrict = AddressDistrict & "140400,140403,º����,36.187895,113.114107;"
    AddressDistrict = AddressDistrict & "140400,140404,�ϵ���,36.052438,113.056679;"
    AddressDistrict = AddressDistrict & "140400,140405,������,36.314072,112.892741;"
    AddressDistrict = AddressDistrict & "140400,140406,º����,36.332232,113.223245;"
    AddressDistrict = AddressDistrict & "140400,140423,��ԫ��,36.532854,113.050094;"
    AddressDistrict = AddressDistrict & "140400,140425,ƽ˳��,36.200202,113.438791;"
    AddressDistrict = AddressDistrict & "140400,140426,�����,36.502971,113.387366;"
    AddressDistrict = AddressDistrict & "140400,140427,������,36.110938,113.206138;"
    AddressDistrict = AddressDistrict & "140400,140428,������,36.119484,112.884656;"
    AddressDistrict = AddressDistrict & "140400,140429,������,36.834315,112.8653;"
    AddressDistrict = AddressDistrict & "140400,140430,����,36.757123,112.70138;"
    AddressDistrict = AddressDistrict & "140400,140431,��Դ��,36.500777,112.340878;"
    AddressDistrict = AddressDistrict & "140500,140502,����,35.496641,112.853106;"
    AddressDistrict = AddressDistrict & "140500,140521,��ˮ��,35.689472,112.187213;"
    AddressDistrict = AddressDistrict & "140500,140522,������,35.482177,112.422014;"
    AddressDistrict = AddressDistrict & "140500,140524,�괨��,35.775614,113.278877;"
    AddressDistrict = AddressDistrict & "140500,140525,������,35.617221,112.899137;"
    AddressDistrict = AddressDistrict & "140500,140581,��ƽ��,35.791355,112.930691;"
    AddressDistrict = AddressDistrict & "140600,140602,˷����,39.324525,112.428676;"
    AddressDistrict = AddressDistrict & "140600,140603,ƽ³��,39.515603,112.295227;"
    AddressDistrict = AddressDistrict & "140600,140621,ɽ����,39.52677,112.816396;"
    AddressDistrict = AddressDistrict & "140600,140622,Ӧ��,39.559187,113.187505;"
    AddressDistrict = AddressDistrict & "140600,140623,������,39.988812,112.465588;"
    AddressDistrict = AddressDistrict & "140600,140681,������,39.820789,113.100511;"
    AddressDistrict = AddressDistrict & "140700,140702,�ܴ���,37.6976,112.740056;"
    AddressDistrict = AddressDistrict & "140700,140703,̫����,37.424595,112.554103;"
    AddressDistrict = AddressDistrict & "140700,140721,������,37.069019,112.973521;"
    AddressDistrict = AddressDistrict & "140700,140722,��Ȩ��,37.079672,113.377834;"
    AddressDistrict = AddressDistrict & "140700,140723,��˳��,37.327027,113.572919;"
    AddressDistrict = AddressDistrict & "140700,140724,������,37.60437,113.706166;"
    AddressDistrict = AddressDistrict & "140700,140725,������,37.891136,113.177708;"
    AddressDistrict = AddressDistrict & "140700,140727,����,37.358739,112.330532;"
    AddressDistrict = AddressDistrict & "140700,140728,ƽң��,37.195474,112.174059;"
    AddressDistrict = AddressDistrict & "140700,140729,��ʯ��,36.847469,111.772759;"
    AddressDistrict = AddressDistrict & "140700,140781,������,37.027616,111.913857;"
    AddressDistrict = AddressDistrict & "140800,140802,�κ���,35.025643,111.000627;"
    AddressDistrict = AddressDistrict & "140800,140821,�����,35.141883,110.77493;"
    AddressDistrict = AddressDistrict & "140800,140822,������,35.417042,110.843561;"
    AddressDistrict = AddressDistrict & "140800,140823,��ϲ��,35.353839,111.220306;"
    AddressDistrict = AddressDistrict & "140800,140824,�ɽ��,35.600412,110.978996;"
    AddressDistrict = AddressDistrict & "140800,140825,�����,35.613697,111.225205;"
    AddressDistrict = AddressDistrict & "140800,140826,���,35.49045,111.576182;"
    AddressDistrict = AddressDistrict & "140800,140827,ԫ����,35.298293,111.67099;"
    AddressDistrict = AddressDistrict & "140800,140828,����,35.140441,111.223174;"
    AddressDistrict = AddressDistrict & "140800,140829,ƽ½��,34.837256,111.212377;"
    AddressDistrict = AddressDistrict & "140800,140830,�ǳ���,34.694769,110.69114;"
    AddressDistrict = AddressDistrict & "140800,140881,������,34.865125,110.447984;"
    AddressDistrict = AddressDistrict & "140800,140882,�ӽ���,35.59715,110.710268;"
    AddressDistrict = AddressDistrict & "140900,140902,�ø���,38.417743,112.734112;"
    AddressDistrict = AddressDistrict & "140900,140921,������,38.484948,112.963231;"
    AddressDistrict = AddressDistrict & "140900,140922,��̨��,38.725711,113.259012;"
    AddressDistrict = AddressDistrict & "140900,140923,����,39.065138,112.962519;"
    AddressDistrict = AddressDistrict & "140900,140924,������,39.188104,113.267707;"
    AddressDistrict = AddressDistrict & "140900,140925,������,39.001718,112.307936;"
    AddressDistrict = AddressDistrict & "140900,140926,������,38.355947,111.940231;"
    AddressDistrict = AddressDistrict & "140900,140927,�����,39.088467,112.200438;"
    AddressDistrict = AddressDistrict & "140900,140928,��կ��,38.912761,111.841015;"
    AddressDistrict = AddressDistrict & "140900,140929,����,38.705625,111.56981;"
    AddressDistrict = AddressDistrict & "140900,140930,������,39.381895,111.146609;"
    AddressDistrict = AddressDistrict & "140900,140931,������,39.022576,111.085688;"
    AddressDistrict = AddressDistrict & "140900,140932,ƫ����,39.442153,111.500477;"
    AddressDistrict = AddressDistrict & "140900,140981,ԭƽ��,38.729186,112.713132;"
    AddressDistrict = AddressDistrict & "141000,141002,Ң����,36.080366,111.522945;"
    AddressDistrict = AddressDistrict & "141000,141021,������,35.641387,111.475529;"
    AddressDistrict = AddressDistrict & "141000,141022,�����,35.738621,111.713508;"
    AddressDistrict = AddressDistrict & "141000,141023,�����,35.876139,111.442932;"
    AddressDistrict = AddressDistrict & "141000,141024,�鶴��,36.255742,111.673692;"
    AddressDistrict = AddressDistrict & "141000,141025,����,36.26855,111.920207;"
    AddressDistrict = AddressDistrict & "141000,141026,������,36.146032,112.251372;"
    AddressDistrict = AddressDistrict & "141000,141027,��ɽ��,35.971359,111.850039;"
    AddressDistrict = AddressDistrict & "141000,141028,����,36.099355,110.682853;"
    AddressDistrict = AddressDistrict & "141000,141029,������,35.975402,110.857365;"
    AddressDistrict = AddressDistrict & "141000,141030,������,36.46383,110.751283;"
    AddressDistrict = AddressDistrict & "141000,141031,����,36.692675,110.935809;"
    AddressDistrict = AddressDistrict & "141000,141032,������,36.760614,110.631276;"
    AddressDistrict = AddressDistrict & "141000,141033,����,36.411682,111.09733;"
    AddressDistrict = AddressDistrict & "141000,141034,������,36.653368,111.563021;"
    AddressDistrict = AddressDistrict & "141000,141081,������,35.620302,111.371272;"
    AddressDistrict = AddressDistrict & "141000,141082,������,36.57202,111.723103;"
    AddressDistrict = AddressDistrict & "141100,141102,��ʯ��,37.524037,111.134462;"
    AddressDistrict = AddressDistrict & "141100,141121,��ˮ��,37.436314,112.032595;"
    AddressDistrict = AddressDistrict & "141100,141122,������,37.555155,112.159154;"
    AddressDistrict = AddressDistrict & "141100,141123,����,38.464136,111.124816;"
    AddressDistrict = AddressDistrict & "141100,141124,����,37.960806,110.995963;"
    AddressDistrict = AddressDistrict & "141100,141125,������,37.431664,110.89613;"
    AddressDistrict = AddressDistrict & "141100,141126,ʯ¥��,36.999426,110.837119;"
    AddressDistrict = AddressDistrict & "141100,141127,���,38.278654,111.671555;"
    AddressDistrict = AddressDistrict & "141100,141128,��ɽ��,37.892632,111.238885;"
    AddressDistrict = AddressDistrict & "141100,141129,������,37.342054,111.193319;"
    AddressDistrict = AddressDistrict & "141100,141130,������,36.983068,111.183188;"
    AddressDistrict = AddressDistrict & "141100,141181,Т����,37.144474,111.781568;"
    AddressDistrict = AddressDistrict & "141100,141182,������,37.267742,111.785273;"
    AddressDistrict = AddressDistrict & "150100,150102,�³���,40.826225,111.685964;"
    AddressDistrict = AddressDistrict & "150100,150103,������,40.815149,111.662162;"
    AddressDistrict = AddressDistrict & "150100,150104,��Ȫ��,40.799421,111.66543;"
    AddressDistrict = AddressDistrict & "150100,150105,������,40.807834,111.698463;"
    AddressDistrict = AddressDistrict & "150100,150121,��Ĭ������,40.720416,111.133615;"
    AddressDistrict = AddressDistrict & "150100,150122,�п�����,40.276729,111.197317;"
    AddressDistrict = AddressDistrict & "150100,150123,���ָ����,40.380288,111.824143;"
    AddressDistrict = AddressDistrict & "150100,150124,��ˮ����,39.912479,111.67222;"
    AddressDistrict = AddressDistrict & "150100,150125,�䴨��,41.094483,111.456563;"
    AddressDistrict = AddressDistrict & "150200,150202,������,40.587056,110.026895;"
    AddressDistrict = AddressDistrict & "150200,150203,��������,40.661345,109.822932;"
    AddressDistrict = AddressDistrict & "150200,150204,��ɽ��,40.668558,109.880049;"
    AddressDistrict = AddressDistrict & "150200,150205,ʯ����,40.672094,110.272565;"
    AddressDistrict = AddressDistrict & "150200,150206,���ƶ�������,41.769246,109.97016;"
    AddressDistrict = AddressDistrict & "150200,150207,��ԭ��,40.600581,109.968122;"
    AddressDistrict = AddressDistrict & "150200,150221,��Ĭ������,40.566434,110.526766;"
    AddressDistrict = AddressDistrict & "150200,150222,������,41.030004,110.063421;"
    AddressDistrict = AddressDistrict & "150200,150223,�����ï����������,41.702836,110.438452;"
    AddressDistrict = AddressDistrict & "150300,150302,��������,39.673527,106.817762;"
    AddressDistrict = AddressDistrict & "150300,150303,������,39.44153,106.884789;"
    AddressDistrict = AddressDistrict & "150300,150304,�ڴ���,39.502288,106.722711;"
    AddressDistrict = AddressDistrict & "150400,150402,��ɽ��,42.269732,118.961087;"
    AddressDistrict = AddressDistrict & "150400,150403,Ԫ��ɽ��,42.041168,119.289877;"
    AddressDistrict = AddressDistrict & "150400,150404,��ɽ��,42.281046,118.938958;"
    AddressDistrict = AddressDistrict & "150400,150421,��³�ƶ�����,43.87877,120.094969;"
    AddressDistrict = AddressDistrict & "150400,150422,��������,43.980715,119.391737;"
    AddressDistrict = AddressDistrict & "150400,150423,��������,43.528963,118.678347;"
    AddressDistrict = AddressDistrict & "150400,150424,������,43.605326,118.05775;"
    AddressDistrict = AddressDistrict & "150400,150425,��ʲ������,43.256233,117.542465;"
    AddressDistrict = AddressDistrict & "150400,150426,��ţ����,42.937128,119.022619;"
    AddressDistrict = AddressDistrict & "150400,150428,��������,41.92778,118.708572;"
    AddressDistrict = AddressDistrict & "150400,150429,������,41.598692,119.339242;"
    AddressDistrict = AddressDistrict & "150400,150430,������,42.287012,119.906486;"
    AddressDistrict = AddressDistrict & "150500,150502,�ƶ�����,43.617422,122.264042;"
    AddressDistrict = AddressDistrict & "150500,150521,�ƶ�����������,44.127166,123.313873;"
    AddressDistrict = AddressDistrict & "150500,150522,�ƶ����������,42.954564,122.355155;"
    AddressDistrict = AddressDistrict & "150500,150523,��³��,43.602432,121.308797;"
    AddressDistrict = AddressDistrict & "150500,150524,������,42.734692,121.774886;"
    AddressDistrict = AddressDistrict & "150500,150525,������,42.84685,120.662543;"
    AddressDistrict = AddressDistrict & "150500,150526,��³����,44.555294,120.905275;"
    AddressDistrict = AddressDistrict & "150500,150581,���ֹ�����,45.532361,119.657862;"
    AddressDistrict = AddressDistrict & "150600,150602,��ʤ��,39.81788,109.98945;"
    AddressDistrict = AddressDistrict & "150600,150603,����ʲ��,39.607472,109.790076;"
    AddressDistrict = AddressDistrict & "150600,150621,��������,40.404076,110.040281;"
    AddressDistrict = AddressDistrict & "150600,150622,׼�����,39.865221,111.238332;"
    AddressDistrict = AddressDistrict & "150600,150623,���п�ǰ��,38.183257,107.48172;"
    AddressDistrict = AddressDistrict & "150600,150624,���п���,39.095752,107.982604;"
    AddressDistrict = AddressDistrict & "150600,150625,������,39.831789,108.736324;"
    AddressDistrict = AddressDistrict & "150600,150626,������,38.596611,108.842454;"
    AddressDistrict = AddressDistrict & "150600,150627,���������,39.604312,109.787402;"
    AddressDistrict = AddressDistrict & "150700,150702,��������,49.213889,119.764923;"
    AddressDistrict = AddressDistrict & "150700,150703,����ŵ����,49.456567,117.716373;"
    AddressDistrict = AddressDistrict & "150700,150721,������,48.130503,123.464615;"
    AddressDistrict = AddressDistrict & "150700,150722,Ī�����ߴ��Ӷ���������,48.478385,124.507401;"
    AddressDistrict = AddressDistrict & "150700,150723,���״�������,50.590177,123.725684;"
    AddressDistrict = AddressDistrict & "150700,150724,���¿���������,49.143293,119.754041;"
    AddressDistrict = AddressDistrict & "150700,150725,�°Ͷ�����,49.328422,119.437609;"
    AddressDistrict = AddressDistrict & "150700,150726,�°Ͷ�������,48.216571,118.267454;"
    AddressDistrict = AddressDistrict & "150700,150727,�°Ͷ�������,48.669134,116.825991;"
    AddressDistrict = AddressDistrict & "150700,150781,��������,49.590788,117.455561;"
    AddressDistrict = AddressDistrict & "150700,150782,����ʯ��,49.287024,120.729005;"
    AddressDistrict = AddressDistrict & "150700,150783,��������,48.007412,122.744401;"
    AddressDistrict = AddressDistrict & "150700,150784,���������,50.2439,120.178636;"
    AddressDistrict = AddressDistrict & "150700,150785,������,50.780454,121.532724;"
    AddressDistrict = AddressDistrict & "150800,150802,�ٺ���,40.757092,107.417018;"
    AddressDistrict = AddressDistrict & "150800,150821,��ԭ��,41.097639,108.270658;"
    AddressDistrict = AddressDistrict & "150800,150822,�����,40.330479,107.006056;"
    AddressDistrict = AddressDistrict & "150800,150823,������ǰ��,40.725209,108.656816;"
    AddressDistrict = AddressDistrict & "150800,150824,����������,41.57254,108.515255;"
    AddressDistrict = AddressDistrict & "150800,150825,�����غ���,41.084307,107.074941;"
    AddressDistrict = AddressDistrict & "150800,150826,��������,40.888797,107.147682;"
    AddressDistrict = AddressDistrict & "150900,150902,������,41.034134,113.116453;"
    AddressDistrict = AddressDistrict & "150900,150921,׿����,40.89576,112.577702;"
    AddressDistrict = AddressDistrict & "150900,150922,������,41.899335,114.01008;"
    AddressDistrict = AddressDistrict & "150900,150923,�̶���,41.560163,113.560643;"
    AddressDistrict = AddressDistrict & "150900,150924,�˺���,40.872437,113.834009;"
    AddressDistrict = AddressDistrict & "150900,150925,������,40.531627,112.500911;"
    AddressDistrict = AddressDistrict & "150900,150926,���������ǰ��,40.786859,113.211958;"
    AddressDistrict = AddressDistrict & "150900,150927,�������������,41.274212,112.633563;"
    AddressDistrict = AddressDistrict & "150900,150928,������������,41.447213,113.1906;"
    AddressDistrict = AddressDistrict & "150900,150929,��������,41.528114,111.70123;"
    AddressDistrict = AddressDistrict & "150900,150981,������,40.437534,113.163462;"
    AddressDistrict = AddressDistrict & "152200,152201,����������,46.077238,122.068975;"
    AddressDistrict = AddressDistrict & "152200,152202,����ɽ��,47.177,119.943656;"
    AddressDistrict = AddressDistrict & "152200,152221,�ƶ�������ǰ��,46.076497,121.957544;"
    AddressDistrict = AddressDistrict & "152200,152222,�ƶ�����������,45.059645,121.472818;"
    AddressDistrict = AddressDistrict & "152200,152223,��������,46.725136,122.909332;"
    AddressDistrict = AddressDistrict & "152200,152224,ͻȪ��,45.380986,121.564856;"
    AddressDistrict = AddressDistrict & "152500,152501,����������,43.652895,111.97981;"
    AddressDistrict = AddressDistrict & "152500,152502,���ֺ�����,43.944301,116.091903;"
    AddressDistrict = AddressDistrict & "152500,152522,���͸���,44.022728,114.970618;"
    AddressDistrict = AddressDistrict & "152500,152523,����������,43.854108,113.653412;"
    AddressDistrict = AddressDistrict & "152500,152524,����������,42.746662,112.65539;"
    AddressDistrict = AddressDistrict & "152500,152525,������������,45.510307,116.980022;"
    AddressDistrict = AddressDistrict & "152500,152526,������������,44.586147,117.615249;"
    AddressDistrict = AddressDistrict & "152500,152527,̫������,41.895199,115.28728;"
    AddressDistrict = AddressDistrict & "152500,152528,�����,42.239229,113.843869;"
    AddressDistrict = AddressDistrict & "152500,152529,�������,42.286807,115.031423;"
    AddressDistrict = AddressDistrict & "152500,152530,������,42.245895,116.003311;"
    AddressDistrict = AddressDistrict & "152500,152531,������,42.197962,116.477288;"
    AddressDistrict = AddressDistrict & "152900,152921,����������,38.847241,105.70192;"
    AddressDistrict = AddressDistrict & "152900,152922,����������,39.21159,101.671984;"
    AddressDistrict = AddressDistrict & "152900,152923,�������,41.958813,101.06944;"
    AddressDistrict = AddressDistrict & "210100,210102,��ƽ��,41.788074,123.406664;"
    AddressDistrict = AddressDistrict & "210100,210103,�����,41.795591,123.445696;"
    AddressDistrict = AddressDistrict & "210100,210104,����,41.808503,123.469956;"
    AddressDistrict = AddressDistrict & "210100,210105,�ʹ���,41.822336,123.405677;"
    AddressDistrict = AddressDistrict & "210100,210106,������,41.787808,123.350664;"
    AddressDistrict = AddressDistrict & "210100,210111,�ռ�����,41.665904,123.341604;"
    AddressDistrict = AddressDistrict & "210100,210112,������,41.741946,123.458981;"
    AddressDistrict = AddressDistrict & "210100,210113,������,42.052312,123.521471;"
    AddressDistrict = AddressDistrict & "210100,210114,�ں���,41.795833,123.310829;"
    AddressDistrict = AddressDistrict & "210100,210115,������,41.512725,122.731269;"
    AddressDistrict = AddressDistrict & "210100,210123,��ƽ��,42.741533,123.352703;"
    AddressDistrict = AddressDistrict & "210100,210124,������,42.507045,123.416722;"
    AddressDistrict = AddressDistrict & "210100,210181,������,41.996508,122.828868;"
    AddressDistrict = AddressDistrict & "210200,210202,��ɽ��,38.921553,121.64376;"
    AddressDistrict = AddressDistrict & "210200,210203,������,38.914266,121.616112;"
    AddressDistrict = AddressDistrict & "210200,210204,ɳ�ӿ���,38.912859,121.593702;"
    AddressDistrict = AddressDistrict & "210200,210211,�ʾ�����,38.975148,121.582614;"
    AddressDistrict = AddressDistrict & "210200,210212,��˳����,38.812043,121.26713;"
    AddressDistrict = AddressDistrict & "210200,210213,������,39.052745,121.789413;"
    AddressDistrict = AddressDistrict & "210200,210214,��������,39.401555,121.9705;"
    AddressDistrict = AddressDistrict & "210200,210224,������,39.272399,122.587824;"
    AddressDistrict = AddressDistrict & "210200,210281,�߷�����,39.63065,122.002656;"
    AddressDistrict = AddressDistrict & "210200,210283,ׯ����,39.69829,122.970612;"
    AddressDistrict = AddressDistrict & "210300,210302,������,41.110344,122.994475;"
    AddressDistrict = AddressDistrict & "210300,210303,������,41.11069,122.971834;"
    AddressDistrict = AddressDistrict & "210300,210304,��ɽ��,41.150622,123.024806;"
    AddressDistrict = AddressDistrict & "210300,210311,ǧɽ��,41.068909,122.949298;"
    AddressDistrict = AddressDistrict & "210300,210321,̨����,41.38686,122.429736;"
    AddressDistrict = AddressDistrict & "210300,210323,�������������,40.281509,123.28833;"
    AddressDistrict = AddressDistrict & "210300,210381,������,40.852533,122.752199;"
    AddressDistrict = AddressDistrict & "210400,210402,�¸���,41.86082,123.902858;"
    AddressDistrict = AddressDistrict & "210400,210403,������,41.866829,124.047219;"
    AddressDistrict = AddressDistrict & "210400,210404,������,41.851803,123.801509;"
    AddressDistrict = AddressDistrict & "210400,210411,˳����,41.881132,123.917165;"
    AddressDistrict = AddressDistrict & "210400,210421,��˳��,41.922644,124.097979;"
    AddressDistrict = AddressDistrict & "210400,210422,�±�����������,41.732456,125.037547;"
    AddressDistrict = AddressDistrict & "210400,210423,��ԭ����������,42.10135,124.927192;"
    AddressDistrict = AddressDistrict & "210500,210502,ƽɽ��,41.291581,123.761231;"
    AddressDistrict = AddressDistrict & "210500,210503,Ϫ����,41.330056,123.765226;"
    AddressDistrict = AddressDistrict & "210500,210504,��ɽ��,41.302429,123.763288;"
    AddressDistrict = AddressDistrict & "210500,210505,�Ϸ���,41.104093,123.748381;"
    AddressDistrict = AddressDistrict & "210500,210521,��Ϫ����������,41.300344,124.126156;"
    AddressDistrict = AddressDistrict & "210500,210522,��������������,41.268997,125.359195;"
    AddressDistrict = AddressDistrict & "210600,210602,Ԫ����,40.136483,124.397814;"
    AddressDistrict = AddressDistrict & "210600,210603,������,40.102801,124.361153;"
    AddressDistrict = AddressDistrict & "210600,210604,����,40.158557,124.427709;"
    AddressDistrict = AddressDistrict & "210600,210624,�������������,40.730412,124.784867;"
    AddressDistrict = AddressDistrict & "210600,210681,������,39.883467,124.149437;"
    AddressDistrict = AddressDistrict & "210600,210682,�����,40.457567,124.071067;"
    AddressDistrict = AddressDistrict & "210700,210702,������,41.115719,121.130085;"
    AddressDistrict = AddressDistrict & "210700,210703,�����,41.114662,121.151304;"
    AddressDistrict = AddressDistrict & "210700,210711,̫����,41.105378,121.107297;"
    AddressDistrict = AddressDistrict & "210700,210726,��ɽ��,41.691804,122.117915;"
    AddressDistrict = AddressDistrict & "210700,210727,����,41.537224,121.242831;"
    AddressDistrict = AddressDistrict & "210700,210781,�躣��,41.171738,121.364236;"
    AddressDistrict = AddressDistrict & "210700,210782,������,41.598764,121.795962;"
    AddressDistrict = AddressDistrict & "210800,210802,վǰ��,40.669949,122.253235;"
    AddressDistrict = AddressDistrict & "210800,210803,������,40.663086,122.210067;"
    AddressDistrict = AddressDistrict & "210800,210804,����Ȧ��,40.263646,122.127242;"
    AddressDistrict = AddressDistrict & "210800,210811,�ϱ���,40.682723,122.382584;"
    AddressDistrict = AddressDistrict & "210800,210881,������,40.405234,122.355534;"
    AddressDistrict = AddressDistrict & "210800,210882,��ʯ����,40.633973,122.505894;"
    AddressDistrict = AddressDistrict & "210900,210902,������,42.011162,121.657639;"
    AddressDistrict = AddressDistrict & "210900,210903,������,42.086603,121.790541;"
    AddressDistrict = AddressDistrict & "210900,210904,̫ƽ��,42.011145,121.677575;"
    AddressDistrict = AddressDistrict & "210900,210905,�������,41.780477,121.42018;"
    AddressDistrict = AddressDistrict & "210900,210911,ϸ����,42.019218,121.654791;"
    AddressDistrict = AddressDistrict & "210900,210921,�����ɹ���������,42.058607,121.743125;"
    AddressDistrict = AddressDistrict & "210900,210922,������,42.384823,122.537444;"
    AddressDistrict = AddressDistrict & "211000,211002,������,41.26745,123.172611;"
    AddressDistrict = AddressDistrict & "211000,211003,��ʥ��,41.266765,123.188227;"
    AddressDistrict = AddressDistrict & "211000,211004,��ΰ��,41.205747,123.200461;"
    AddressDistrict = AddressDistrict & "211000,211005,��������,41.157831,123.431633;"
    AddressDistrict = AddressDistrict & "211000,211011,̫�Ӻ���,41.251682,123.185336;"
    AddressDistrict = AddressDistrict & "211000,211021,������,41.216479,123.079674;"
    AddressDistrict = AddressDistrict & "211000,211081,������,41.427836,123.325864;"
    AddressDistrict = AddressDistrict & "211100,211102,˫̨����,41.190365,122.055733;"
    AddressDistrict = AddressDistrict & "211100,211103,��¡̨��,41.122423,122.071624;"
    AddressDistrict = AddressDistrict & "211100,211104,������,40.994428,122.071708;"
    AddressDistrict = AddressDistrict & "211100,211122,��ɽ��,41.240701,121.98528;"
    AddressDistrict = AddressDistrict & "211200,211202,������,42.292278,123.844877;"
    AddressDistrict = AddressDistrict & "211200,211204,�����,42.542978,124.14896;"
    AddressDistrict = AddressDistrict & "211200,211221,������,42.223316,123.725669;"
    AddressDistrict = AddressDistrict & "211200,211223,������,42.738091,124.72332;"
    AddressDistrict = AddressDistrict & "211200,211224,��ͼ��,42.784441,124.11017;"
    AddressDistrict = AddressDistrict & "211200,211281,����ɽ��,42.450734,123.545366;"
    AddressDistrict = AddressDistrict & "211200,211282,��ԭ��,42.542141,124.045551;"
    AddressDistrict = AddressDistrict & "211300,211302,˫����,41.579389,120.44877;"
    AddressDistrict = AddressDistrict & "211300,211303,������,41.576749,120.413376;"
    AddressDistrict = AddressDistrict & "211300,211321,������,41.526342,120.404217;"
    AddressDistrict = AddressDistrict & "211300,211322,��ƽ��,41.402576,119.642363;"
    AddressDistrict = AddressDistrict & "211300,211324,�����������ɹ���������,41.125428,119.744883;"
    AddressDistrict = AddressDistrict & "211300,211381,��Ʊ��,41.803286,120.766951;"
    AddressDistrict = AddressDistrict & "211300,211382,��Դ��,41.243086,119.404789;"
    AddressDistrict = AddressDistrict & "211400,211402,��ɽ��,40.755143,120.85937;"
    AddressDistrict = AddressDistrict & "211400,211403,������,40.709991,120.838569;"
    AddressDistrict = AddressDistrict & "211400,211404,��Ʊ��,41.098813,120.752314;"
    AddressDistrict = AddressDistrict & "211400,211421,������,40.328407,120.342112;"
    AddressDistrict = AddressDistrict & "211400,211422,������,40.812871,119.807776;"
    AddressDistrict = AddressDistrict & "211400,211481,�˳���,40.619413,120.729365;"
    AddressDistrict = AddressDistrict & "220100,220102,�Ϲ���,43.890235,125.337237;"
    AddressDistrict = AddressDistrict & "220100,220103,�����,43.903823,125.342828;"
    AddressDistrict = AddressDistrict & "220100,220104,������,43.86491,125.318042;"
    AddressDistrict = AddressDistrict & "220100,220105,������,43.870824,125.384727;"
    AddressDistrict = AddressDistrict & "220100,220106,��԰��,43.892177,125.272467;"
    AddressDistrict = AddressDistrict & "220100,220112,˫����,43.525168,125.659018;"
    AddressDistrict = AddressDistrict & "220100,220113,��̨��,44.157155,125.844682;"
    AddressDistrict = AddressDistrict & "220100,220122,ũ����,44.431258,125.175287;"
    AddressDistrict = AddressDistrict & "220100,220182,������,44.827642,126.550107;"
    AddressDistrict = AddressDistrict & "220100,220183,�»���,44.533909,125.703327;"
    AddressDistrict = AddressDistrict & "220100,220184,��������,43.509474,124.817588;"
    AddressDistrict = AddressDistrict & "220200,220202,������,43.851118,126.570766;"
    AddressDistrict = AddressDistrict & "220200,220203,��̶��,43.909755,126.561429;"
    AddressDistrict = AddressDistrict & "220200,220204,��Ӫ��,43.843804,126.55239;"
    AddressDistrict = AddressDistrict & "220200,220211,������,43.816594,126.560759;"
    AddressDistrict = AddressDistrict & "220200,220221,������,43.667416,126.501622;"
    AddressDistrict = AddressDistrict & "220200,220281,�Ժ���,43.720579,127.342739;"
    AddressDistrict = AddressDistrict & "220200,220282,�����,42.972093,126.745445;"
    AddressDistrict = AddressDistrict & "220200,220283,������,44.410906,126.947813;"
    AddressDistrict = AddressDistrict & "220200,220284,��ʯ��,42.942476,126.059929;"
    AddressDistrict = AddressDistrict & "220300,220302,������,43.176263,124.360894;"
    AddressDistrict = AddressDistrict & "220300,220303,������,43.16726,124.388464;"
    AddressDistrict = AddressDistrict & "220300,220322,������,43.30831,124.335802;"
    AddressDistrict = AddressDistrict & "220300,220323,��ͨ����������,43.345464,125.303124;"
    AddressDistrict = AddressDistrict & "220300,220382,˫����,43.518275,123.505283;"
    AddressDistrict = AddressDistrict & "220400,220402,��ɽ��,42.902702,125.145164;"
    AddressDistrict = AddressDistrict & "220400,220403,������,42.920415,125.151424;"
    AddressDistrict = AddressDistrict & "220400,220421,������,42.675228,125.529623;"
    AddressDistrict = AddressDistrict & "220400,220422,������,42.927724,124.991995;"
    AddressDistrict = AddressDistrict & "220500,220502,������,41.721233,125.936716;"
    AddressDistrict = AddressDistrict & "220500,220503,��������,41.777564,126.045987;"
    AddressDistrict = AddressDistrict & "220500,220521,ͨ����,41.677918,125.753121;"
    AddressDistrict = AddressDistrict & "220500,220523,������,42.683459,126.042821;"
    AddressDistrict = AddressDistrict & "220500,220524,������,42.281484,125.740536;"
    AddressDistrict = AddressDistrict & "220500,220581,÷�ӿ���,42.530002,125.687336;"
    AddressDistrict = AddressDistrict & "220500,220582,������,41.126276,126.186204;"
    AddressDistrict = AddressDistrict & "220600,220602,�뽭��,41.943065,126.428035;"
    AddressDistrict = AddressDistrict & "220600,220605,��Դ��,42.048109,126.584229;"
    AddressDistrict = AddressDistrict & "220600,220621,������,42.332643,127.273796;"
    AddressDistrict = AddressDistrict & "220600,220622,������,42.389689,126.808386;"
    AddressDistrict = AddressDistrict & "220600,220623,���׳�����������,41.419361,128.203384;"
    AddressDistrict = AddressDistrict & "220600,220681,�ٽ���,41.810689,126.919296;"
    AddressDistrict = AddressDistrict & "220700,220702,������,45.176498,124.827851;"
    AddressDistrict = AddressDistrict & "220700,220721,ǰ������˹�ɹ���������,45.116288,124.826808;"
    AddressDistrict = AddressDistrict & "220700,220722,������,44.276579,123.985184;"
    AddressDistrict = AddressDistrict & "220700,220723,Ǭ����,45.006846,124.024361;"
    AddressDistrict = AddressDistrict & "220700,220781,������,44.986199,126.042758;"
    AddressDistrict = AddressDistrict & "220800,220802,䬱���,45.619253,122.842499;"
    AddressDistrict = AddressDistrict & "220800,220821,������,45.846089,123.202246;"
    AddressDistrict = AddressDistrict & "220800,220822,ͨ����,44.80915,123.088543;"
    AddressDistrict = AddressDistrict & "220800,220881,�����,45.339113,122.783779;"
    AddressDistrict = AddressDistrict & "220800,220882,����,45.507648,124.291512;"
    AddressDistrict = AddressDistrict & "222400,222401,�Ӽ���,42.906964,129.51579;"
    AddressDistrict = AddressDistrict & "222400,222402,ͼ����,42.966621,129.846701;"
    AddressDistrict = AddressDistrict & "222400,222403,�ػ���,43.366921,128.22986;"
    AddressDistrict = AddressDistrict & "222400,222404,������,42.871057,130.365787;"
    AddressDistrict = AddressDistrict & "222400,222405,������,42.771029,129.425747;"
    AddressDistrict = AddressDistrict & "222400,222406,������,42.547004,129.008748;"
    AddressDistrict = AddressDistrict & "222400,222424,������,43.315426,129.766161;"
    AddressDistrict = AddressDistrict & "222400,222426,��ͼ��,43.110994,128.901865;"
    AddressDistrict = AddressDistrict & "230100,230102,������,45.762035,126.612532;"
    AddressDistrict = AddressDistrict & "230100,230103,�ϸ���,45.755971,126.652098;"
    AddressDistrict = AddressDistrict & "230100,230104,������,45.78454,126.648838;"
    AddressDistrict = AddressDistrict & "230100,230108,ƽ����,45.605567,126.629257;"
    AddressDistrict = AddressDistrict & "230100,230109,�ɱ���,45.814656,126.563066;"
    AddressDistrict = AddressDistrict & "230100,230110,�㷻��,45.713067,126.667049;"
    AddressDistrict = AddressDistrict & "230100,230111,������,45.98423,126.603302;"
    AddressDistrict = AddressDistrict & "230100,230112,������,45.538372,126.972726;"
    AddressDistrict = AddressDistrict & "230100,230113,˫����,45.377942,126.308784;"
    AddressDistrict = AddressDistrict & "230100,230123,������,46.315105,129.565594;"
    AddressDistrict = AddressDistrict & "230100,230124,������,45.839536,128.836131;"
    AddressDistrict = AddressDistrict & "230100,230125,����,45.759369,127.48594;"
    AddressDistrict = AddressDistrict & "230100,230126,������,46.081889,127.403602;"
    AddressDistrict = AddressDistrict & "230100,230127,ľ����,45.949826,128.042675;"
    AddressDistrict = AddressDistrict & "230100,230128,ͨ����,45.977618,128.747786;"
    AddressDistrict = AddressDistrict & "230100,230129,������,45.455648,128.331886;"
    AddressDistrict = AddressDistrict & "230100,230183,��־��,45.214953,127.968539;"
    AddressDistrict = AddressDistrict & "230100,230184,�峣��,44.919418,127.15759;"
    AddressDistrict = AddressDistrict & "230200,230202,��ɳ��,47.341736,123.957338;"
    AddressDistrict = AddressDistrict & "230200,230203,������,47.354494,123.955888;"
    AddressDistrict = AddressDistrict & "230200,230204,������,47.339499,123.973555;"
    AddressDistrict = AddressDistrict & "230200,230205,����Ϫ��,47.156867,123.813181;"
    AddressDistrict = AddressDistrict & "230200,230206,����������,47.20697,123.638873;"
    AddressDistrict = AddressDistrict & "230200,230207,����ɽ��,47.51401,122.887972;"
    AddressDistrict = AddressDistrict & "230200,230208,÷��˹���Ӷ�����,47.311113,123.754599;"
    AddressDistrict = AddressDistrict & "230200,230221,������,47.336388,123.187225;"
    AddressDistrict = AddressDistrict & "230200,230223,������,47.890098,125.307561;"
    AddressDistrict = AddressDistrict & "230200,230224,̩����,46.39233,123.41953;"
    AddressDistrict = AddressDistrict & "230200,230225,������,47.917838,123.506034;"
    AddressDistrict = AddressDistrict & "230200,230227,��ԣ��,47.797172,124.469106;"
    AddressDistrict = AddressDistrict & "230200,230229,��ɽ��,48.034342,125.874355;"
    AddressDistrict = AddressDistrict & "230200,230230,�˶���,48.03732,126.249094;"
    AddressDistrict = AddressDistrict & "230200,230231,��Ȫ��,47.607363,126.091911;"
    AddressDistrict = AddressDistrict & "230200,230281,ګ����,48.481133,124.882172;"
    AddressDistrict = AddressDistrict & "230300,230302,������,45.30034,130.974374;"
    AddressDistrict = AddressDistrict & "230300,230303,��ɽ��,45.213242,130.910636;"
    AddressDistrict = AddressDistrict & "230300,230304,�ε���,45.348812,130.846823;"
    AddressDistrict = AddressDistrict & "230300,230305,������,45.092195,130.697781;"
    AddressDistrict = AddressDistrict & "230300,230306,���Ӻ���,45.338248,131.010501;"
    AddressDistrict = AddressDistrict & "230300,230307,��ɽ��,45.209607,130.481126;"
    AddressDistrict = AddressDistrict & "230300,230321,������,45.250892,131.148907;"
    AddressDistrict = AddressDistrict & "230300,230381,������,45.767985,132.973881;"
    AddressDistrict = AddressDistrict & "230300,230382,��ɽ��,45.54725,131.874137;"
    AddressDistrict = AddressDistrict & "230400,230402,������,47.345372,130.292478;"
    AddressDistrict = AddressDistrict & "230400,230403,��ũ��,47.331678,130.276652;"
    AddressDistrict = AddressDistrict & "230400,230404,��ɽ��,47.31324,130.275533;"
    AddressDistrict = AddressDistrict & "230400,230405,�˰���,47.252911,130.236169;"
    AddressDistrict = AddressDistrict & "230400,230406,��ɽ��,47.337385,130.31714;"
    AddressDistrict = AddressDistrict & "230400,230407,��ɽ��,47.35997,130.30534;"
    AddressDistrict = AddressDistrict & "230400,230421,�ܱ���,47.577577,130.829087;"
    AddressDistrict = AddressDistrict & "230400,230422,�����,47.289892,131.860526;"
    AddressDistrict = AddressDistrict & "230500,230502,��ɽ��,46.642961,131.15896;"
    AddressDistrict = AddressDistrict & "230500,230503,�붫��,46.591076,131.163675;"
    AddressDistrict = AddressDistrict & "230500,230505,�ķ�̨��,46.594347,131.333181;"
    AddressDistrict = AddressDistrict & "230500,230506,��ɽ��,46.573366,131.404294;"
    AddressDistrict = AddressDistrict & "230500,230521,������,46.72898,131.13933;"
    AddressDistrict = AddressDistrict & "230500,230522,������,46.775159,131.810622;"
    AddressDistrict = AddressDistrict & "230500,230523,������,46.328781,132.206415;"
    AddressDistrict = AddressDistrict & "230500,230524,�ĺ���,46.801288,134.021162;"
    AddressDistrict = AddressDistrict & "230600,230602,����ͼ��,46.596356,125.114643;"
    AddressDistrict = AddressDistrict & "230600,230603,������,46.573948,125.145794;"
    AddressDistrict = AddressDistrict & "230600,230604,�ú�·��,46.653254,124.868341;"
    AddressDistrict = AddressDistrict & "230600,230605,�����,46.403049,124.889528;"
    AddressDistrict = AddressDistrict & "230600,230606,��ͬ��,46.034304,124.818509;"
    AddressDistrict = AddressDistrict & "230600,230621,������,45.708685,125.273254;"
    AddressDistrict = AddressDistrict & "230600,230622,��Դ��,45.518832,125.081974;"
    AddressDistrict = AddressDistrict & "230600,230623,�ֵ���,47.186411,124.877742;"
    AddressDistrict = AddressDistrict & "230600,230624,�Ŷ������ɹ���������,46.865973,124.446259;"
    AddressDistrict = AddressDistrict & "230700,230717,������,47.728171,128.907303;"
    AddressDistrict = AddressDistrict & "230700,230718,�ڴ���,47.726728,128.669859;"
    AddressDistrict = AddressDistrict & "230700,230719,�Ѻ���,47.853778,128.84075;"
    AddressDistrict = AddressDistrict & "230700,230722,������,48.891378,130.397684;"
    AddressDistrict = AddressDistrict & "230700,230723,������,48.454651,129.571108;"
    AddressDistrict = AddressDistrict & "230700,230724,������,48.290455,129.5336;"
    AddressDistrict = AddressDistrict & "230700,230725,����ɽ��,47.028397,129.020793;"
    AddressDistrict = AddressDistrict & "230700,230726,�ϲ���,47.137314,129.28246;"
    AddressDistrict = AddressDistrict & "230700,230751,������,47.413074,129.429117;"
    AddressDistrict = AddressDistrict & "230700,230781,������,46.985772,128.030561;"
    AddressDistrict = AddressDistrict & "230800,230803,������,46.809645,130.361786;"
    AddressDistrict = AddressDistrict & "230800,230804,ǰ����,46.812345,130.377684;"
    AddressDistrict = AddressDistrict & "230800,230805,������,46.822476,130.403297;"
    AddressDistrict = AddressDistrict & "230800,230811,����,46.80712,130.351588;"
    AddressDistrict = AddressDistrict & "230800,230822,������,46.240118,130.570112;"
    AddressDistrict = AddressDistrict & "230800,230826,�봨��,47.023039,130.723713;"
    AddressDistrict = AddressDistrict & "230800,230828,��ԭ��,46.730048,129.904463;"
    AddressDistrict = AddressDistrict & "230800,230881,ͬ����,47.651131,132.510119;"
    AddressDistrict = AddressDistrict & "230800,230882,������,47.250747,132.037951;"
    AddressDistrict = AddressDistrict & "230800,230883,��Զ��,48.364707,134.294501;"
    AddressDistrict = AddressDistrict & "230900,230902,������,45.794258,130.889482;"
    AddressDistrict = AddressDistrict & "230900,230903,��ɽ��,45.771217,131.015848;"
    AddressDistrict = AddressDistrict & "230900,230904,���Ӻ���,45.776587,131.071561;"
    AddressDistrict = AddressDistrict & "230900,230921,������,45.751573,130.575025;"
    AddressDistrict = AddressDistrict & "231000,231002,������,44.582399,129.623292;"
    AddressDistrict = AddressDistrict & "231000,231003,������,44.596328,129.634645;"
    AddressDistrict = AddressDistrict & "231000,231004,������,44.595443,129.601232;"
    AddressDistrict = AddressDistrict & "231000,231005,������,44.581032,129.61311;"
    AddressDistrict = AddressDistrict & "231000,231025,�ֿ���,45.286645,130.268402;"
    AddressDistrict = AddressDistrict & "231000,231081,��Һ���,44.396864,131.164856;"
    AddressDistrict = AddressDistrict & "231000,231083,������,44.574149,129.387902;"
    AddressDistrict = AddressDistrict & "231000,231084,������,44.346836,129.470019;"
    AddressDistrict = AddressDistrict & "231000,231085,������,44.91967,130.527085;"
    AddressDistrict = AddressDistrict & "231000,231086,������,44.063578,131.125296;"
    AddressDistrict = AddressDistrict & "231100,231102,������,50.249027,127.497639;"
    AddressDistrict = AddressDistrict & "231100,231123,ѷ����,49.582974,128.476152;"
    AddressDistrict = AddressDistrict & "231100,231124,������,49.423941,127.327315;"
    AddressDistrict = AddressDistrict & "231100,231181,������,48.245437,126.508737;"
    AddressDistrict = AddressDistrict & "231100,231182,���������,48.512688,126.197694;"
    AddressDistrict = AddressDistrict & "231100,231183,�۽���,49.177461,125.229904;"
    AddressDistrict = AddressDistrict & "231200,231202,������,46.634912,126.990665;"
    AddressDistrict = AddressDistrict & "231200,231221,������,46.83352,126.484191;"
    AddressDistrict = AddressDistrict & "231200,231222,������,46.259037,126.289315;"
    AddressDistrict = AddressDistrict & "231200,231223,�����,46.686596,126.112268;"
    AddressDistrict = AddressDistrict & "231200,231224,�찲��,46.879203,127.510024;"
    AddressDistrict = AddressDistrict & "231200,231225,��ˮ��,47.183527,125.907544;"
    AddressDistrict = AddressDistrict & "231200,231226,������,47.247195,127.111121;"
    AddressDistrict = AddressDistrict & "231200,231281,������,46.410614,125.329926;"
    AddressDistrict = AddressDistrict & "231200,231282,�ض���,46.069471,125.991402;"
    AddressDistrict = AddressDistrict & "231200,231283,������,47.460428,126.969383;"
    AddressDistrict = AddressDistrict & "232700,232701,Į����,52.972074,122.536256;"
    AddressDistrict = AddressDistrict & "232700,232718,�Ӹ������,50.424654,124.126716;"
    AddressDistrict = AddressDistrict & "232700,232721,������,51.726998,126.662105;"
    AddressDistrict = AddressDistrict & "232700,232722,������,52.335229,124.710516;"
    AddressDistrict = AddressDistrict & "320100,320102,������,32.050678,118.792199;"
    AddressDistrict = AddressDistrict & "320100,320104,�ػ���,32.033818,118.786088;"
    AddressDistrict = AddressDistrict & "320100,320105,������,32.004538,118.732688;"
    AddressDistrict = AddressDistrict & "320100,320106,��¥��,32.066966,118.769739;"
    AddressDistrict = AddressDistrict & "320100,320111,�ֿ���,32.05839,118.625307;"
    AddressDistrict = AddressDistrict & "320100,320113,��ϼ��,32.102147,118.808702;"
    AddressDistrict = AddressDistrict & "320100,320114,�껨̨��,31.995946,118.77207;"
    AddressDistrict = AddressDistrict & "320100,320115,������,31.953418,118.850621;"
    AddressDistrict = AddressDistrict & "320100,320116,������,32.340655,118.85065;"
    AddressDistrict = AddressDistrict & "320100,320117,��ˮ��,31.653061,119.028732;"
    AddressDistrict = AddressDistrict & "320100,320118,�ߴ���,31.327132,118.87589;"
    AddressDistrict = AddressDistrict & "320200,320205,��ɽ��,31.585559,120.357298;"
    AddressDistrict = AddressDistrict & "320200,320206,��ɽ��,31.681019,120.303543;"
    AddressDistrict = AddressDistrict & "320200,320211,������,31.550228,120.266053;"
    AddressDistrict = AddressDistrict & "320200,320213,��Ϫ��,31.575706,120.296595;"
    AddressDistrict = AddressDistrict & "320200,320214,������,31.550966,120.352782;"
    AddressDistrict = AddressDistrict & "320200,320281,������,31.910984,120.275891;"
    AddressDistrict = AddressDistrict & "320200,320282,������,31.364384,119.820538;"
    AddressDistrict = AddressDistrict & "320300,320302,��¥��,34.269397,117.192941;"
    AddressDistrict = AddressDistrict & "320300,320303,������,34.254805,117.194589;"
    AddressDistrict = AddressDistrict & "320300,320305,������,34.441642,117.450212;"
    AddressDistrict = AddressDistrict & "320300,320311,Ȫɽ��,34.262249,117.182225;"
    AddressDistrict = AddressDistrict & "320300,320312,ͭɽ��,34.19288,117.183894;"
    AddressDistrict = AddressDistrict & "320300,320321,����,34.696946,116.592888;"
    AddressDistrict = AddressDistrict & "320300,320322,����,34.729044,116.937182;"
    AddressDistrict = AddressDistrict & "320300,320324,�����,33.899222,117.95066;"
    AddressDistrict = AddressDistrict & "320300,320381,������,34.368779,118.345828;"
    AddressDistrict = AddressDistrict & "320300,320382,������,34.314708,117.963923;"
    AddressDistrict = AddressDistrict & "320400,320402,������,31.779632,119.963783;"
    AddressDistrict = AddressDistrict & "320400,320404,��¥��,31.78096,119.948388;"
    AddressDistrict = AddressDistrict & "320400,320411,�±���,31.824664,119.974654;"
    AddressDistrict = AddressDistrict & "320400,320412,�����,31.718566,119.958773;"
    AddressDistrict = AddressDistrict & "320400,320413,��̳��,31.744399,119.573395;"
    AddressDistrict = AddressDistrict & "320400,320481,������,31.427081,119.487816;"
    AddressDistrict = AddressDistrict & "320500,320505,������,31.294845,120.566833;"
    AddressDistrict = AddressDistrict & "320500,320506,������,31.270839,120.624621;"
    AddressDistrict = AddressDistrict & "320500,320507,�����,31.396684,120.618956;"
    AddressDistrict = AddressDistrict & "320500,320508,������,31.311414,120.622249;"
    AddressDistrict = AddressDistrict & "320500,320509,�⽭��,31.160404,120.641601;"
    AddressDistrict = AddressDistrict & "320500,320581,������,31.658156,120.74852;"
    AddressDistrict = AddressDistrict & "320500,320582,�żҸ���,31.865553,120.543441;"
    AddressDistrict = AddressDistrict & "320500,320583,��ɽ��,31.381925,120.958137;"
    AddressDistrict = AddressDistrict & "320500,320585,̫����,31.452568,121.112275;"
    AddressDistrict = AddressDistrict & "320600,320602,�紨��,32.015278,120.86635;"
    AddressDistrict = AddressDistrict & "320600,320612,ͨ����,32.084287,121.073171;"
    AddressDistrict = AddressDistrict & "320600,320623,�綫��,32.311832,121.186088;"
    AddressDistrict = AddressDistrict & "320600,320681,������,31.810158,121.659724;"
    AddressDistrict = AddressDistrict & "320600,320682,�����,32.391591,120.566324;"
    AddressDistrict = AddressDistrict & "320600,320684,������,31.893528,121.176609;"
    AddressDistrict = AddressDistrict & "320600,320685,������,32.540288,120.465995;"
    AddressDistrict = AddressDistrict & "320700,320703,������,34.739529,119.366487;"
    AddressDistrict = AddressDistrict & "320700,320706,������,34.601584,119.179793;"
    AddressDistrict = AddressDistrict & "320700,320707,������,34.839154,119.128774;"
    AddressDistrict = AddressDistrict & "320700,320722,������,34.522859,118.766489;"
    AddressDistrict = AddressDistrict & "320700,320723,������,34.298436,119.255741;"
    AddressDistrict = AddressDistrict & "320700,320724,������,34.092553,119.352331;"
    AddressDistrict = AddressDistrict & "320800,320803,������,33.507499,119.14634;"
    AddressDistrict = AddressDistrict & "320800,320804,������,33.622452,119.020817;"
    AddressDistrict = AddressDistrict & "320800,320812,�彭����,33.603234,119.019454;"
    AddressDistrict = AddressDistrict & "320800,320813,������,33.294975,118.867875;"
    AddressDistrict = AddressDistrict & "320800,320826,��ˮ��,33.771308,119.266078;"
    AddressDistrict = AddressDistrict & "320800,320830,������,33.00439,118.493823;"
    AddressDistrict = AddressDistrict & "320800,320831,�����,33.018162,119.016936;"
    AddressDistrict = AddressDistrict & "320900,320902,ͤ����,33.383912,120.136078;"
    AddressDistrict = AddressDistrict & "320900,320903,�ζ���,33.341288,120.139753;"
    AddressDistrict = AddressDistrict & "320900,320904,�����,33.199531,120.470324;"
    AddressDistrict = AddressDistrict & "320900,320921,��ˮ��,34.19996,119.579573;"
    AddressDistrict = AddressDistrict & "320900,320922,������,33.989888,119.828434;"
    AddressDistrict = AddressDistrict & "320900,320923,������,33.78573,119.805338;"
    AddressDistrict = AddressDistrict & "320900,320924,������,33.773779,120.257444;"
    AddressDistrict = AddressDistrict & "320900,320925,������,33.472621,119.793105;"
    AddressDistrict = AddressDistrict & "320900,320981,��̨��,32.853174,120.314101;"
    AddressDistrict = AddressDistrict & "321000,321002,������,32.392154,119.442267;"
    AddressDistrict = AddressDistrict & "321000,321003,������,32.377899,119.397777;"
    AddressDistrict = AddressDistrict & "321000,321012,������,32.426564,119.567481;"
    AddressDistrict = AddressDistrict & "321000,321023,��Ӧ��,33.23694,119.321284;"
    AddressDistrict = AddressDistrict & "321000,321081,������,32.271965,119.182443;"
    AddressDistrict = AddressDistrict & "321000,321084,������,32.785164,119.443842;"
    AddressDistrict = AddressDistrict & "321100,321102,������,32.206191,119.454571;"
    AddressDistrict = AddressDistrict & "321100,321111,������,32.213501,119.414877;"
    AddressDistrict = AddressDistrict & "321100,321112,��ͽ��,32.128972,119.433883;"
    AddressDistrict = AddressDistrict & "321100,321181,������,31.991459,119.581911;"
    AddressDistrict = AddressDistrict & "321100,321182,������,32.237266,119.828054;"
    AddressDistrict = AddressDistrict & "321100,321183,������,31.947355,119.167135;"
    AddressDistrict = AddressDistrict & "321200,321202,������,32.488406,119.920187;"
    AddressDistrict = AddressDistrict & "321200,321203,�߸���,32.315701,119.88166;"
    AddressDistrict = AddressDistrict & "321200,321204,������,32.508483,120.148208;"
    AddressDistrict = AddressDistrict & "321200,321281,�˻���,32.938065,119.840162;"
    AddressDistrict = AddressDistrict & "321200,321282,������,32.018168,120.26825;"
    AddressDistrict = AddressDistrict & "321200,321283,̩����,32.168784,120.020228;"
    AddressDistrict = AddressDistrict & "321300,321302,�޳���,33.937726,118.278984;"
    AddressDistrict = AddressDistrict & "321300,321311,��ԥ��,33.941071,118.330012;"
    AddressDistrict = AddressDistrict & "321300,321322,������,34.129097,118.775889;"
    AddressDistrict = AddressDistrict & "321300,321323,������,33.711433,118.681284;"
    AddressDistrict = AddressDistrict & "321300,321324,������,33.456538,118.211824;"
    AddressDistrict = AddressDistrict & "330100,330102,�ϳ���,30.250236,120.171465;"
    AddressDistrict = AddressDistrict & "330100,330105,������,30.314697,120.150053;"
    AddressDistrict = AddressDistrict & "330100,330106,������,30.272934,120.147376;"
    AddressDistrict = AddressDistrict & "330100,330108,������,30.206615,120.21062;"
    AddressDistrict = AddressDistrict & "330100,330109,��ɽ��,30.162932,120.27069;"
    AddressDistrict = AddressDistrict & "330100,330110,�ຼ��,30.27365,119.978959;"
    AddressDistrict = AddressDistrict & "330100,330111,������,30.049871,119.949869;"
    AddressDistrict = AddressDistrict & "330100,330112,�ٰ���,30.231153,119.715101;"
    AddressDistrict = AddressDistrict & "330100,330114,Ǯ����,30.322904,120.493972;"
    AddressDistrict = AddressDistrict & "330100,330113,��ƽ��,30.419025,120.299376;"
    AddressDistrict = AddressDistrict & "330100,330122,ͩ®��,29.797437,119.685045;"
    AddressDistrict = AddressDistrict & "330100,330127,������,29.604177,119.044276;"
    AddressDistrict = AddressDistrict & "330100,330182,������,29.472284,119.279089;"
    AddressDistrict = AddressDistrict & "330200,330203,������,29.874452,121.539698;"
    AddressDistrict = AddressDistrict & "330200,330205,������,29.888361,121.559282;"
    AddressDistrict = AddressDistrict & "330200,330206,������,29.90944,121.831303;"
    AddressDistrict = AddressDistrict & "330200,330211,����,29.952107,121.713162;"
    AddressDistrict = AddressDistrict & "330200,330212,۴����,29.831662,121.558436;"
    AddressDistrict = AddressDistrict & "330200,330213,���,29.662348,121.41089;"
    AddressDistrict = AddressDistrict & "330200,330225,��ɽ��,29.470206,121.877091;"
    AddressDistrict = AddressDistrict & "330200,330226,������,29.299836,121.432606;"
    AddressDistrict = AddressDistrict & "330200,330281,��Ҧ��,30.045404,121.156294;"
    AddressDistrict = AddressDistrict & "330200,330282,��Ϫ��,30.177142,121.248052;"
    AddressDistrict = AddressDistrict & "330300,330302,¹����,28.003352,120.674231;"
    AddressDistrict = AddressDistrict & "330300,330303,������,27.970254,120.763469;"
    AddressDistrict = AddressDistrict & "330300,330304,걺���,28.006444,120.637145;"
    AddressDistrict = AddressDistrict & "330300,330305,��ͷ��,27.836057,121.156181;"
    AddressDistrict = AddressDistrict & "330300,330324,������,28.153886,120.690968;"
    AddressDistrict = AddressDistrict & "330300,330326,ƽ����,27.6693,120.564387;"
    AddressDistrict = AddressDistrict & "330300,330327,������,27.507743,120.406256;"
    AddressDistrict = AddressDistrict & "330300,330328,�ĳ���,27.789133,120.09245;"
    AddressDistrict = AddressDistrict & "330300,330329,̩˳��,27.557309,119.71624;"
    AddressDistrict = AddressDistrict & "330300,330381,����,27.779321,120.646171;"
    AddressDistrict = AddressDistrict & "330300,330382,������,28.116083,120.967147;"
    AddressDistrict = AddressDistrict & "330300,330383,������,27.578156,120.553039;"
    AddressDistrict = AddressDistrict & "330400,330402,�Ϻ���,30.764652,120.749953;"
    AddressDistrict = AddressDistrict & "330400,330411,������,30.763323,120.720431;"
    AddressDistrict = AddressDistrict & "330400,330421,������,30.841352,120.921871;"
    AddressDistrict = AddressDistrict & "330400,330424,������,30.522223,120.942017;"
    AddressDistrict = AddressDistrict & "330400,330481,������,30.525544,120.688821;"
    AddressDistrict = AddressDistrict & "330400,330482,ƽ����,30.698921,121.014666;"
    AddressDistrict = AddressDistrict & "330400,330483,ͩ����,30.629065,120.551085;"
    AddressDistrict = AddressDistrict & "330500,330502,������,30.867252,120.101416;"
    AddressDistrict = AddressDistrict & "330500,330503,�����,30.872742,120.417195;"
    AddressDistrict = AddressDistrict & "330500,330521,������,30.534927,119.967662;"
    AddressDistrict = AddressDistrict & "330500,330522,������,31.00475,119.910122;"
    AddressDistrict = AddressDistrict & "330500,330523,������,30.631974,119.687891;"
    AddressDistrict = AddressDistrict & "330600,330602,Խ����,29.996993,120.585315;"
    AddressDistrict = AddressDistrict & "330600,330603,������,30.078038,120.476075;"
    AddressDistrict = AddressDistrict & "330600,330604,������,30.016769,120.874185;"
    AddressDistrict = AddressDistrict & "330600,330624,�²���,29.501205,120.905665;"
    AddressDistrict = AddressDistrict & "330600,330681,������,29.713662,120.244326;"
    AddressDistrict = AddressDistrict & "330600,330683,������,29.586606,120.82888;"
    AddressDistrict = AddressDistrict & "330700,330702,�ĳ���,29.082607,119.652579;"
    AddressDistrict = AddressDistrict & "330700,330703,����,29.095835,119.681264;"
    AddressDistrict = AddressDistrict & "330700,330723,������,28.896563,119.819159;"
    AddressDistrict = AddressDistrict & "330700,330726,�ֽ���,29.451254,119.893363;"
    AddressDistrict = AddressDistrict & "330700,330727,�Ͱ���,29.052627,120.44513;"
    AddressDistrict = AddressDistrict & "330700,330781,��Ϫ��,29.210065,119.460521;"
    AddressDistrict = AddressDistrict & "330700,330782,������,29.306863,120.074911;"
    AddressDistrict = AddressDistrict & "330700,330783,������,29.262546,120.23334;"
    AddressDistrict = AddressDistrict & "330700,330784,������,28.895293,120.036328;"
    AddressDistrict = AddressDistrict & "330800,330802,�³���,28.944539,118.873041;"
    AddressDistrict = AddressDistrict & "330800,330803,�齭��,28.973195,118.957683;"
    AddressDistrict = AddressDistrict & "330800,330822,��ɽ��,28.900039,118.521654;"
    AddressDistrict = AddressDistrict & "330800,330824,������,29.136503,118.414435;"
    AddressDistrict = AddressDistrict & "330800,330825,������,29.031364,119.172525;"
    AddressDistrict = AddressDistrict & "330800,330881,��ɽ��,28.734674,118.627879;"
    AddressDistrict = AddressDistrict & "330900,330902,������,30.016423,122.108496;"
    AddressDistrict = AddressDistrict & "330900,330903,������,29.945614,122.301953;"
    AddressDistrict = AddressDistrict & "330900,330921,�ɽ��,30.242865,122.201132;"
    AddressDistrict = AddressDistrict & "330900,330922,������,30.727166,122.457809;"
    AddressDistrict = AddressDistrict & "331000,331002,������,28.67615,121.431049;"
    AddressDistrict = AddressDistrict & "331000,331003,������,28.64488,121.262138;"
    AddressDistrict = AddressDistrict & "331000,331004,·����,28.581799,121.37292;"
    AddressDistrict = AddressDistrict & "331000,331022,������,29.118955,121.376429;"
    AddressDistrict = AddressDistrict & "331000,331023,��̨��,29.141126,121.031227;"
    AddressDistrict = AddressDistrict & "331000,331024,�ɾ���,28.849213,120.735074;"
    AddressDistrict = AddressDistrict & "331000,331081,������,28.368781,121.373611;"
    AddressDistrict = AddressDistrict & "331000,331082,�ٺ���,28.845441,121.131229;"
    AddressDistrict = AddressDistrict & "331000,331083,����,28.12842,121.232337;"
    AddressDistrict = AddressDistrict & "331100,331102,������,28.451103,119.922293;"
    AddressDistrict = AddressDistrict & "331100,331121,������,28.135247,120.291939;"
    AddressDistrict = AddressDistrict & "331100,331122,������,28.654208,120.078965;"
    AddressDistrict = AddressDistrict & "331100,331123,�����,28.5924,119.27589;"
    AddressDistrict = AddressDistrict & "331100,331124,������,28.449937,119.485292;"
    AddressDistrict = AddressDistrict & "331100,331125,�ƺ���,28.111077,119.569458;"
    AddressDistrict = AddressDistrict & "331100,331126,��Ԫ��,27.618231,119.067233;"
    AddressDistrict = AddressDistrict & "331100,331127,�������������,27.977247,119.634669;"
    AddressDistrict = AddressDistrict & "331100,331181,��Ȫ��,28.069177,119.132319;"
    AddressDistrict = AddressDistrict & "340100,340102,������,31.86961,117.315358;"
    AddressDistrict = AddressDistrict & "340100,340103,®����,31.869011,117.283776;"
    AddressDistrict = AddressDistrict & "340100,340104,��ɽ��,31.855868,117.262072;"
    AddressDistrict = AddressDistrict & "340100,340111,������,31.82956,117.285751;"
    AddressDistrict = AddressDistrict & "340100,340121,������,32.478548,117.164699;"
    AddressDistrict = AddressDistrict & "340100,340122,�ʶ���,31.883992,117.463222;"
    AddressDistrict = AddressDistrict & "340100,340123,������,31.719646,117.166118;"
    AddressDistrict = AddressDistrict & "340100,340124,®����,31.251488,117.289844;"
    AddressDistrict = AddressDistrict & "340100,340181,������,31.600518,117.874155;"
    AddressDistrict = AddressDistrict & "340200,340202,������,31.32559,118.376343;"
    AddressDistrict = AddressDistrict & "340200,340207,𯽭��,31.362716,118.400174;"
    AddressDistrict = AddressDistrict & "340200,340209,߮����,31.313394,118.377476;"
    AddressDistrict = AddressDistrict & "340200,340210,��b��,31.145262,118.572301;"
    AddressDistrict = AddressDistrict & "340200,340211,������,31.080896,118.201349;"
    AddressDistrict = AddressDistrict & "340200,340223,������,30.919638,118.337104;"
    AddressDistrict = AddressDistrict & "340200,340281,��Ϊ��,31.303075,117.911432;"
    AddressDistrict = AddressDistrict & "340300,340302,���Ӻ���,32.950452,117.382312;"
    AddressDistrict = AddressDistrict & "340300,340303,��ɽ��,32.938066,117.355789;"
    AddressDistrict = AddressDistrict & "340300,340304,�����,32.931933,117.35259;"
    AddressDistrict = AddressDistrict & "340300,340311,������,32.963147,117.34709;"
    AddressDistrict = AddressDistrict & "340300,340321,��Զ��,32.956934,117.200171;"
    AddressDistrict = AddressDistrict & "340300,340322,�����,33.146202,117.888809;"
    AddressDistrict = AddressDistrict & "340300,340323,������,33.318679,117.315962;"
    AddressDistrict = AddressDistrict & "340400,340402,��ͨ��,32.632066,117.052927;"
    AddressDistrict = AddressDistrict & "340400,340403,�������,32.644342,117.018318;"
    AddressDistrict = AddressDistrict & "340400,340404,л�Ҽ���,32.598289,116.865354;"
    AddressDistrict = AddressDistrict & "340400,340405,�˹�ɽ��,32.628229,116.841111;"
    AddressDistrict = AddressDistrict & "340400,340406,�˼���,32.782117,116.816879;"
    AddressDistrict = AddressDistrict & "340400,340421,��̨��,32.705382,116.722769;"
    AddressDistrict = AddressDistrict & "340400,340422,����,32.577304,116.785349;"
    AddressDistrict = AddressDistrict & "340500,340503,��ɽ��,31.69902,118.511308;"
    AddressDistrict = AddressDistrict & "340500,340504,��ɽ��,31.685912,118.493104;"
    AddressDistrict = AddressDistrict & "340500,340506,������,31.562321,118.843742;"
    AddressDistrict = AddressDistrict & "340500,340521,��Ϳ��,31.556167,118.489873;"
    AddressDistrict = AddressDistrict & "340500,340522,��ɽ��,31.727758,118.105545;"
    AddressDistrict = AddressDistrict & "340500,340523,����,31.716634,118.362998;"
    AddressDistrict = AddressDistrict & "340600,340602,�ż���,33.991218,116.833925;"
    AddressDistrict = AddressDistrict & "340600,340603,��ɽ��,33.970916,116.790775;"
    AddressDistrict = AddressDistrict & "340600,340604,��ɽ��,33.889529,116.809465;"
    AddressDistrict = AddressDistrict & "340600,340621,�Ϫ��,33.916407,116.767435;"
    AddressDistrict = AddressDistrict & "340700,340705,ͭ����,30.927613,117.816167;"
    AddressDistrict = AddressDistrict & "340700,340706,�尲��,30.952338,117.792288;"
    AddressDistrict = AddressDistrict & "340700,340711,����,30.908927,117.80707;"
    AddressDistrict = AddressDistrict & "340700,340722,������,30.700615,117.222027;"
    AddressDistrict = AddressDistrict & "340800,340802,ӭ����,30.506375,117.044965;"
    AddressDistrict = AddressDistrict & "340800,340803,�����,30.505632,117.034512;"
    AddressDistrict = AddressDistrict & "340800,340811,������,30.541323,117.070003;"
    AddressDistrict = AddressDistrict & "340800,340822,������,30.734994,116.828664;"
    AddressDistrict = AddressDistrict & "340800,340825,̫����,30.451869,116.305225;"
    AddressDistrict = AddressDistrict & "340800,340826,������,30.158327,116.120204;"
    AddressDistrict = AddressDistrict & "340800,340827,������,30.12491,116.690927;"
    AddressDistrict = AddressDistrict & "340800,340828,������,30.848502,116.360482;"
    AddressDistrict = AddressDistrict & "340800,340881,ͩ����,31.050576,116.959656;"
    AddressDistrict = AddressDistrict & "340800,340882,Ǳɽ��,30.638222,116.573665;"
    AddressDistrict = AddressDistrict & "341000,341002,��Ϫ��,29.709186,118.317354;"
    AddressDistrict = AddressDistrict & "341000,341003,��ɽ��,30.294517,118.136639;"
    AddressDistrict = AddressDistrict & "341000,341004,������,29.825201,118.339743;"
    AddressDistrict = AddressDistrict & "341000,341021,���,29.867748,118.428025;"
    AddressDistrict = AddressDistrict & "341000,341022,������,29.788878,118.188531;"
    AddressDistrict = AddressDistrict & "341000,341023,����,29.923812,117.942911;"
    AddressDistrict = AddressDistrict & "341000,341024,������,29.853472,117.717237;"
    AddressDistrict = AddressDistrict & "341100,341102,������,32.303797,118.316475;"
    AddressDistrict = AddressDistrict & "341100,341103,������,32.329841,118.296955;"
    AddressDistrict = AddressDistrict & "341100,341122,������,32.450231,118.433293;"
    AddressDistrict = AddressDistrict & "341100,341124,ȫ����,32.09385,118.268576;"
    AddressDistrict = AddressDistrict & "341100,341125,��Զ��,32.527105,117.683713;"
    AddressDistrict = AddressDistrict & "341100,341126,������,32.867146,117.562461;"
    AddressDistrict = AddressDistrict & "341100,341181,�쳤��,32.6815,119.011212;"
    AddressDistrict = AddressDistrict & "341100,341182,������,32.781206,117.998048;"
    AddressDistrict = AddressDistrict & "341200,341202,�����,32.891238,115.813914;"
    AddressDistrict = AddressDistrict & "341200,341203,򣶫��,32.908861,115.858747;"
    AddressDistrict = AddressDistrict & "341200,341204,�Ȫ��,32.924797,115.804525;"
    AddressDistrict = AddressDistrict & "341200,341221,��Ȫ��,33.062698,115.261688;"
    AddressDistrict = AddressDistrict & "341200,341222,̫����,33.16229,115.627243;"
    AddressDistrict = AddressDistrict & "341200,341225,������,32.638102,115.590534;"
    AddressDistrict = AddressDistrict & "341200,341226,�����,32.637065,116.259122;"
    AddressDistrict = AddressDistrict & "341200,341282,������,33.26153,115.362117;"
    AddressDistrict = AddressDistrict & "341300,341302,������,33.633853,116.983309;"
    AddressDistrict = AddressDistrict & "341300,341321,�ɽ��,34.426247,116.351113;"
    AddressDistrict = AddressDistrict & "341300,341322,����,34.183266,116.945399;"
    AddressDistrict = AddressDistrict & "341300,341323,�����,33.540629,117.551493;"
    AddressDistrict = AddressDistrict & "341300,341324,����,33.47758,117.885443;"
    AddressDistrict = AddressDistrict & "341500,341502,����,31.754491,116.503288;"
    AddressDistrict = AddressDistrict & "341500,341503,ԣ����,31.750692,116.494543;"
    AddressDistrict = AddressDistrict & "341500,341504,Ҷ����,31.84768,115.913594;"
    AddressDistrict = AddressDistrict & "341500,341522,������,32.341305,116.278875;"
    AddressDistrict = AddressDistrict & "341500,341523,�����,31.462848,116.944088;"
    AddressDistrict = AddressDistrict & "341500,341524,��կ��,31.681624,115.878514;"
    AddressDistrict = AddressDistrict & "341500,341525,��ɽ��,31.402456,116.333078;"
    AddressDistrict = AddressDistrict & "341600,341602,�۳���,33.869284,115.781214;"
    AddressDistrict = AddressDistrict & "341600,341621,������,33.502831,116.211551;"
    AddressDistrict = AddressDistrict & "341600,341622,�ɳ���,33.260814,116.560337;"
    AddressDistrict = AddressDistrict & "341600,341623,������,33.143503,116.207782;"
    AddressDistrict = AddressDistrict & "341700,341702,�����,30.657378,117.488342;"
    AddressDistrict = AddressDistrict & "341700,341721,������,30.096568,117.021476;"
    AddressDistrict = AddressDistrict & "341700,341722,ʯ̨��,30.210324,117.482907;"
    AddressDistrict = AddressDistrict & "341700,341723,������,30.63818,117.857395;"
    AddressDistrict = AddressDistrict & "341800,341802,������,30.946003,118.758412;"
    AddressDistrict = AddressDistrict & "341800,341821,��Ϫ��,31.127834,119.185024;"
    AddressDistrict = AddressDistrict & "341800,341823,����,30.685975,118.412397;"
    AddressDistrict = AddressDistrict & "341800,341824,��Ϫ��,30.065267,118.594705;"
    AddressDistrict = AddressDistrict & "341800,341825,캵���,30.288057,118.543081;"
    AddressDistrict = AddressDistrict & "341800,341881,������,30.626529,118.983407;"
    AddressDistrict = AddressDistrict & "341800,341882,�����,30.893116,119.417521;"
    AddressDistrict = AddressDistrict & "350100,350102,��¥��,26.082284,119.29929;"
    AddressDistrict = AddressDistrict & "350100,350103,̨����,26.058616,119.310156;"
    AddressDistrict = AddressDistrict & "350100,350104,��ɽ��,26.038912,119.320988;"
    AddressDistrict = AddressDistrict & "350100,350105,��β��,25.991975,119.458725;"
    AddressDistrict = AddressDistrict & "350100,350111,������,26.078837,119.328597;"
    AddressDistrict = AddressDistrict & "350100,350112,������,25.960583,119.510849;"
    AddressDistrict = AddressDistrict & "350100,350121,������,26.148567,119.145117;"
    AddressDistrict = AddressDistrict & "350100,350122,������,26.202109,119.538365;"
    AddressDistrict = AddressDistrict & "350100,350123,��Դ��,26.487234,119.552645;"
    AddressDistrict = AddressDistrict & "350100,350124,������,26.223793,118.868416;"
    AddressDistrict = AddressDistrict & "350100,350125,��̩��,25.864825,118.939089;"
    AddressDistrict = AddressDistrict & "350100,350128,ƽ̶��,25.503672,119.791197;"
    AddressDistrict = AddressDistrict & "350100,350181,������,25.720402,119.376992;"
    AddressDistrict = AddressDistrict & "350200,350203,˼����,24.462059,118.087828;"
    AddressDistrict = AddressDistrict & "350200,350205,������,24.492512,118.036364;"
    AddressDistrict = AddressDistrict & "350200,350206,������,24.512764,118.10943;"
    AddressDistrict = AddressDistrict & "350200,350211,������,24.572874,118.100869;"
    AddressDistrict = AddressDistrict & "350200,350212,ͬ����,24.729333,118.150455;"
    AddressDistrict = AddressDistrict & "350200,350213,�谲��,24.637479,118.242811;"
    AddressDistrict = AddressDistrict & "350300,350302,������,25.433737,119.001028;"
    AddressDistrict = AddressDistrict & "350300,350303,������,25.459273,119.119102;"
    AddressDistrict = AddressDistrict & "350300,350304,�����,25.430047,119.020047;"
    AddressDistrict = AddressDistrict & "350300,350305,������,25.316141,119.092607;"
    AddressDistrict = AddressDistrict & "350300,350322,������,25.356529,118.694331;"
    AddressDistrict = AddressDistrict & "350400,350403,��Ԫ��,26.234191,117.607418;"
    AddressDistrict = AddressDistrict & "350400,350421,��Ϫ��,26.357375,117.201845;"
    AddressDistrict = AddressDistrict & "350400,350423,������,26.17761,116.815821;"
    AddressDistrict = AddressDistrict & "350400,350424,������,26.259932,116.659725;"
    AddressDistrict = AddressDistrict & "350400,350425,������,25.690803,117.849355;"
    AddressDistrict = AddressDistrict & "350400,350426,��Ϫ��,26.169261,118.188577;"
    AddressDistrict = AddressDistrict & "350400,350427,ɳ����,26.397361,117.789095;"
    AddressDistrict = AddressDistrict & "350400,350428,������,26.728667,117.473558;"
    AddressDistrict = AddressDistrict & "350400,350429,̩����,26.897995,117.177522;"
    AddressDistrict = AddressDistrict & "350400,350430,������,26.831398,116.845832;"
    AddressDistrict = AddressDistrict & "350400,350481,������,25.974075,117.364447;"
    AddressDistrict = AddressDistrict & "350500,350502,�����,24.907645,118.588929;"
    AddressDistrict = AddressDistrict & "350500,350503,������,24.896041,118.605147;"
    AddressDistrict = AddressDistrict & "350500,350504,�彭��,24.941153,118.670312;"
    AddressDistrict = AddressDistrict & "350500,350505,Ȫ����,25.126859,118.912285;"
    AddressDistrict = AddressDistrict & "350500,350521,�ݰ���,25.028718,118.798954;"
    AddressDistrict = AddressDistrict & "350500,350524,��Ϫ��,25.056824,118.186014;"
    AddressDistrict = AddressDistrict & "350500,350525,������,25.320721,118.29503;"
    AddressDistrict = AddressDistrict & "350500,350526,�»���,25.489004,118.242986;"
    AddressDistrict = AddressDistrict & "350500,350527,������,24.436417,118.323221;"
    AddressDistrict = AddressDistrict & "350500,350581,ʯʨ��,24.731978,118.628402;"
    AddressDistrict = AddressDistrict & "350500,350582,������,24.807322,118.577338;"
    AddressDistrict = AddressDistrict & "350500,350583,�ϰ���,24.959494,118.387031;"
    AddressDistrict = AddressDistrict & "350600,350602,ܼ����,24.509955,117.656461;"
    AddressDistrict = AddressDistrict & "350600,350603,������,24.515656,117.671387;"
    AddressDistrict = AddressDistrict & "350600,350622,������,23.950486,117.340946;"
    AddressDistrict = AddressDistrict & "350600,350623,������,24.117907,117.614023;"
    AddressDistrict = AddressDistrict & "350600,350624,گ����,23.710834,117.176083;"
    AddressDistrict = AddressDistrict & "350600,350625,��̩��,24.621475,117.755913;"
    AddressDistrict = AddressDistrict & "350600,350626,��ɽ��,23.702845,117.427679;"
    AddressDistrict = AddressDistrict & "350600,350627,�Ͼ���,24.516425,117.365462;"
    AddressDistrict = AddressDistrict & "350600,350628,ƽ����,24.366158,117.313549;"
    AddressDistrict = AddressDistrict & "350600,350629,������,25.001416,117.53631;"
    AddressDistrict = AddressDistrict & "350600,350681,������,24.445341,117.817292;"
    AddressDistrict = AddressDistrict & "350700,350702,��ƽ��,26.636079,118.178918;"
    AddressDistrict = AddressDistrict & "350700,350703,������,27.332067,118.12267;"
    AddressDistrict = AddressDistrict & "350700,350721,˳����,26.792851,117.80771;"
    AddressDistrict = AddressDistrict & "350700,350722,�ֳ���,27.920412,118.536822;"
    AddressDistrict = AddressDistrict & "350700,350723,������,27.542803,117.337897;"
    AddressDistrict = AddressDistrict & "350700,350724,��Ϫ��,27.525785,118.783491;"
    AddressDistrict = AddressDistrict & "350700,350725,������,27.365398,118.858661;"
    AddressDistrict = AddressDistrict & "350700,350781,������,27.337952,117.491544;"
    AddressDistrict = AddressDistrict & "350700,350782,����ɽ��,27.751733,118.032796;"
    AddressDistrict = AddressDistrict & "350700,350783,�����,27.03502,118.321765;"
    AddressDistrict = AddressDistrict & "350800,350802,������,25.0918,117.030721;"
    AddressDistrict = AddressDistrict & "350800,350803,������,24.720442,116.732691;"
    AddressDistrict = AddressDistrict & "350800,350821,��͡��,25.842278,116.361007;"
    AddressDistrict = AddressDistrict & "350800,350823,�Ϻ���,25.050019,116.424774;"
    AddressDistrict = AddressDistrict & "350800,350824,��ƽ��,25.08865,116.100928;"
    AddressDistrict = AddressDistrict & "350800,350825,������,25.708506,116.756687;"
    AddressDistrict = AddressDistrict & "350800,350881,��ƽ��,25.291597,117.42073;"
    AddressDistrict = AddressDistrict & "350900,350902,������,26.659253,119.527225;"
    AddressDistrict = AddressDistrict & "350900,350921,ϼ����,26.882068,120.005214;"
    AddressDistrict = AddressDistrict & "350900,350922,������,26.577491,118.743156;"
    AddressDistrict = AddressDistrict & "350900,350923,������,26.910826,118.987544;"
    AddressDistrict = AddressDistrict & "350900,350924,������,27.457798,119.506733;"
    AddressDistrict = AddressDistrict & "350900,350925,������,27.103106,119.338239;"
    AddressDistrict = AddressDistrict & "350900,350926,������,27.236163,119.898226;"
    AddressDistrict = AddressDistrict & "350900,350981,������,27.084246,119.650798;"
    AddressDistrict = AddressDistrict & "350900,350982,������,27.318884,120.219761;"
    AddressDistrict = AddressDistrict & "360100,360102,������,28.682988,115.889675;"
    AddressDistrict = AddressDistrict & "360100,360103,������,28.662901,115.91065;"
    AddressDistrict = AddressDistrict & "360100,360104,��������,28.635724,115.907292;"
    AddressDistrict = AddressDistrict & "360100,360111,��ɽ����,28.689292,115.949044;"
    AddressDistrict = AddressDistrict & "360100,360112,�½���,28.690788,115.820806;"
    AddressDistrict = AddressDistrict & "360100,360113,���̲��,28.69819928,115.8580521;"
    AddressDistrict = AddressDistrict & "360100,360121,�ϲ���,28.543781,115.942465;"
    AddressDistrict = AddressDistrict & "360100,360123,������,28.841334,115.553109;"
    AddressDistrict = AddressDistrict & "360100,360124,������,28.365681,116.267671;"
    AddressDistrict = AddressDistrict & "360200,360202,������,29.288465,117.195023;"
    AddressDistrict = AddressDistrict & "360200,360203,��ɽ��,29.292812,117.214814;"
    AddressDistrict = AddressDistrict & "360200,360222,������,29.352251,117.217611;"
    AddressDistrict = AddressDistrict & "360200,360281,��ƽ��,28.967361,117.129376;"
    AddressDistrict = AddressDistrict & "360300,360302,��Դ��,27.625826,113.855044;"
    AddressDistrict = AddressDistrict & "360300,360313,�涫��,27.639319,113.7456;"
    AddressDistrict = AddressDistrict & "360300,360321,������,27.127807,113.955582;"
    AddressDistrict = AddressDistrict & "360300,360322,������,27.877041,113.800525;"
    AddressDistrict = AddressDistrict & "360300,360323,«Ϫ��,27.633633,114.041206;"
    AddressDistrict = AddressDistrict & "360400,360402,�Ϫ��,29.676175,115.99012;"
    AddressDistrict = AddressDistrict & "360400,360403,�����,29.72465,115.995947;"
    AddressDistrict = AddressDistrict & "360400,360404,��ɣ��,29.610264,115.892977;"
    AddressDistrict = AddressDistrict & "360400,360423,������,29.260182,115.105646;"
    AddressDistrict = AddressDistrict & "360400,360424,��ˮ��,29.032729,114.573428;"
    AddressDistrict = AddressDistrict & "360400,360425,������,29.018212,115.809055;"
    AddressDistrict = AddressDistrict & "360400,360426,�°���,29.327474,115.762611;"
    AddressDistrict = AddressDistrict & "360400,360428,������,29.275105,116.205114;"
    AddressDistrict = AddressDistrict & "360400,360429,������,29.7263,116.244313;"
    AddressDistrict = AddressDistrict & "360400,360430,������,29.898865,116.55584;"
    AddressDistrict = AddressDistrict & "360400,360481,�����,29.676599,115.669081;"
    AddressDistrict = AddressDistrict & "360400,360482,�������,29.247884,115.805712;"
    AddressDistrict = AddressDistrict & "360400,360483,®ɽ��,29.456169,116.043743;"
    AddressDistrict = AddressDistrict & "360500,360502,��ˮ��,27.819171,114.923923;"
    AddressDistrict = AddressDistrict & "360500,360521,������,27.811301,114.675262;"
    AddressDistrict = AddressDistrict & "360600,360602,�º���,28.239076,117.034112;"
    AddressDistrict = AddressDistrict & "360600,360603,�཭��,28.206177,116.822763;"
    AddressDistrict = AddressDistrict & "360600,360681,��Ϫ��,28.283693,117.212103;"
    AddressDistrict = AddressDistrict & "360700,360702,�¹���,25.851367,114.93872;"
    AddressDistrict = AddressDistrict & "360700,360703,�Ͽ���,25.661721,114.756933;"
    AddressDistrict = AddressDistrict & "360700,360704,������,25.865432,115.018461;"
    AddressDistrict = AddressDistrict & "360700,360722,�ŷ���,25.38023,114.930893;"
    AddressDistrict = AddressDistrict & "360700,360723,������,25.395937,114.362243;"
    AddressDistrict = AddressDistrict & "360700,360724,������,25.794284,114.540537;"
    AddressDistrict = AddressDistrict & "360700,360725,������,25.687911,114.307348;"
    AddressDistrict = AddressDistrict & "360700,360726,��Զ��,25.134591,115.392328;"
    AddressDistrict = AddressDistrict & "360700,360728,������,24.774277,115.03267;"
    AddressDistrict = AddressDistrict & "360700,360729,ȫ����,24.742651,114.531589;"
    AddressDistrict = AddressDistrict & "360700,360730,������,26.472054,116.018782;"
    AddressDistrict = AddressDistrict & "360700,360731,�ڶ���,25.955033,115.411198;"
    AddressDistrict = AddressDistrict & "360700,360732,�˹���,26.330489,115.351896;"
    AddressDistrict = AddressDistrict & "360700,360733,�����,25.599125,115.791158;"
    AddressDistrict = AddressDistrict & "360700,360734,Ѱ����,24.954136,115.651399;"
    AddressDistrict = AddressDistrict & "360700,360735,ʯ����,26.326582,116.342249;"
    AddressDistrict = AddressDistrict & "360700,360781,�����,25.875278,116.034854;"
    AddressDistrict = AddressDistrict & "360700,360783,������,24.90476,114.792657;"
    AddressDistrict = AddressDistrict & "360800,360802,������,27.112367,114.987331;"
    AddressDistrict = AddressDistrict & "360800,360803,��ԭ��,27.105879,115.016306;"
    AddressDistrict = AddressDistrict & "360800,360821,������,27.040042,114.905117;"
    AddressDistrict = AddressDistrict & "360800,360822,��ˮ��,27.213445,115.134569;"
    AddressDistrict = AddressDistrict & "360800,360823,Ͽ����,27.580862,115.319331;"
    AddressDistrict = AddressDistrict & "360800,360824,�¸���,27.755758,115.399294;"
    AddressDistrict = AddressDistrict & "360800,360825,������,27.321087,115.435559;"
    AddressDistrict = AddressDistrict & "360800,360826,̩����,26.790164,114.901393;"
    AddressDistrict = AddressDistrict & "360800,360827,�촨��,26.323705,114.51689;"
    AddressDistrict = AddressDistrict & "360800,360828,����,26.462085,114.784694;"
    AddressDistrict = AddressDistrict & "360800,360829,������,27.382746,114.61384;"
    AddressDistrict = AddressDistrict & "360800,360830,������,26.944721,114.242534;"
    AddressDistrict = AddressDistrict & "360800,360881,����ɽ��,26.745919,114.284421;"
    AddressDistrict = AddressDistrict & "360900,360902,Ԭ����,27.800117,114.387379;"
    AddressDistrict = AddressDistrict & "360900,360921,������,28.700672,115.389899;"
    AddressDistrict = AddressDistrict & "360900,360922,������,28.104528,114.449012;"
    AddressDistrict = AddressDistrict & "360900,360923,�ϸ���,28.234789,114.932653;"
    AddressDistrict = AddressDistrict & "360900,360924,�˷���,28.388289,114.787381;"
    AddressDistrict = AddressDistrict & "360900,360925,������,28.86054,115.361744;"
    AddressDistrict = AddressDistrict & "360900,360926,ͭ����,28.520956,114.37014;"
    AddressDistrict = AddressDistrict & "360900,360981,�����,28.191584,115.786005;"
    AddressDistrict = AddressDistrict & "360900,360982,������,28.055898,115.543388;"
    AddressDistrict = AddressDistrict & "360900,360983,�߰���,28.420951,115.381527;"
    AddressDistrict = AddressDistrict & "361000,361002,�ٴ���,27.981919,116.361404;"
    AddressDistrict = AddressDistrict & "361000,361003,������,28.2325,116.605341;"
    AddressDistrict = AddressDistrict & "361000,361021,�ϳ���,27.55531,116.63945;"
    AddressDistrict = AddressDistrict & "361000,361022,�质��,27.292561,116.91457;"
    AddressDistrict = AddressDistrict & "361000,361023,�Ϸ���,27.210132,116.532994;"
    AddressDistrict = AddressDistrict & "361000,361024,������,27.760907,116.059109;"
    AddressDistrict = AddressDistrict & "361000,361025,�ְ���,27.420101,115.838432;"
    AddressDistrict = AddressDistrict & "361000,361026,�˻���,27.546512,116.223023;"
    AddressDistrict = AddressDistrict & "361000,361027,��Ϫ��,27.907387,116.778751;"
    AddressDistrict = AddressDistrict & "361000,361028,��Ϫ��,27.70653,117.066095;"
    AddressDistrict = AddressDistrict & "361000,361030,�����,26.838426,116.327291;"
    AddressDistrict = AddressDistrict & "361100,361102,������,28.445378,117.970522;"
    AddressDistrict = AddressDistrict & "361100,361103,�����,28.440285,118.189852;"
    AddressDistrict = AddressDistrict & "361100,361104,������,28.453897,117.90612;"
    AddressDistrict = AddressDistrict & "361100,361123,��ɽ��,28.673479,118.244408;"
    AddressDistrict = AddressDistrict & "361100,361124,Ǧɽ��,28.310892,117.711906;"
    AddressDistrict = AddressDistrict & "361100,361125,�����,28.415103,117.608247;"
    AddressDistrict = AddressDistrict & "361100,361126,߮����,28.402391,117.435002;"
    AddressDistrict = AddressDistrict & "361100,361127,�����,28.69173,116.691072;"
    AddressDistrict = AddressDistrict & "361100,361128,۶����,28.993374,116.673748;"
    AddressDistrict = AddressDistrict & "361100,361129,������,28.692589,117.07015;"
    AddressDistrict = AddressDistrict & "361100,361130,��Դ��,29.254015,117.86219;"
    AddressDistrict = AddressDistrict & "361100,361181,������,28.945034,117.578732;"
    AddressDistrict = AddressDistrict & "370100,370102,������,36.664169,117.03862;"
    AddressDistrict = AddressDistrict & "370100,370103,������,36.657354,116.99898;"
    AddressDistrict = AddressDistrict & "370100,370104,������,36.668205,116.947921;"
    AddressDistrict = AddressDistrict & "370100,370105,������,36.693374,116.996086;"
    AddressDistrict = AddressDistrict & "370100,370112,������,36.681744,117.063744;"
    AddressDistrict = AddressDistrict & "370100,370113,������,36.561049,116.74588;"
    AddressDistrict = AddressDistrict & "370100,370114,������,36.71209,117.54069;"
    AddressDistrict = AddressDistrict & "370100,370115,������,36.976771,117.176035;"
    AddressDistrict = AddressDistrict & "370100,370116,������,36.214395,117.675808;"
    AddressDistrict = AddressDistrict & "370100,370117,�ֳ���,36.058038,117.82033;"
    AddressDistrict = AddressDistrict & "370100,370124,ƽ����,36.286923,116.455054;"
    AddressDistrict = AddressDistrict & "370100,370126,�̺���,37.310544,117.156369;"
    AddressDistrict = AddressDistrict & "370200,370202,������,36.070892,120.395966;"
    AddressDistrict = AddressDistrict & "370200,370203,�б���,36.083819,120.355026;"
    AddressDistrict = AddressDistrict & "370200,370211,�Ƶ���,35.875138,119.995518;"
    AddressDistrict = AddressDistrict & "370200,370212,��ɽ��,36.102569,120.467393;"
    AddressDistrict = AddressDistrict & "370200,370213,�����,36.160023,120.421236;"
    AddressDistrict = AddressDistrict & "370200,370214,������,36.306833,120.389135;"
    AddressDistrict = AddressDistrict & "370200,370215,��ī��,36.390847,120.447352;"
    AddressDistrict = AddressDistrict & "370200,370281,������,36.285878,120.006202;"
    AddressDistrict = AddressDistrict & "370200,370283,ƽ����,36.788828,119.959012;"
    AddressDistrict = AddressDistrict & "370200,370285,������,36.86509,120.526226;"
    AddressDistrict = AddressDistrict & "370300,370302,�ʹ���,36.647272,117.967696;"
    AddressDistrict = AddressDistrict & "370300,370303,�ŵ���,36.807049,118.053521;"
    AddressDistrict = AddressDistrict & "370300,370304,��ɽ��,36.497567,117.85823;"
    AddressDistrict = AddressDistrict & "370300,370305,������,36.816657,118.306018;"
    AddressDistrict = AddressDistrict & "370300,370306,�ܴ���,36.803699,117.851036;"
    AddressDistrict = AddressDistrict & "370300,370321,��̨��,36.959773,118.101556;"
    AddressDistrict = AddressDistrict & "370300,370322,������,37.169581,117.829839;"
    AddressDistrict = AddressDistrict & "370300,370323,��Դ��,36.186282,118.166161;"
    AddressDistrict = AddressDistrict & "370400,370402,������,34.856651,117.557281;"
    AddressDistrict = AddressDistrict & "370400,370403,Ѧ����,34.79789,117.265293;"
    AddressDistrict = AddressDistrict & "370400,370404,ỳ���,34.767713,117.586316;"
    AddressDistrict = AddressDistrict & "370400,370405,̨��ׯ��,34.564815,117.734747;"
    AddressDistrict = AddressDistrict & "370400,370406,ɽͤ��,35.096077,117.458968;"
    AddressDistrict = AddressDistrict & "370400,370481,������,35.088498,117.162098;"
    AddressDistrict = AddressDistrict & "370500,370502,��Ӫ��,37.461567,118.507543;"
    AddressDistrict = AddressDistrict & "370500,370503,�ӿ���,37.886015,118.529613;"
    AddressDistrict = AddressDistrict & "370500,370505,������,37.588679,118.551314;"
    AddressDistrict = AddressDistrict & "370500,370522,������,37.493365,118.248854;"
    AddressDistrict = AddressDistrict & "370500,370523,������,37.05161,118.407522;"
    AddressDistrict = AddressDistrict & "370600,370602,֥���,37.540925,121.385877;"
    AddressDistrict = AddressDistrict & "370600,370611,��ɽ��,37.496875,121.264741;"
    AddressDistrict = AddressDistrict & "370600,370612,Ĳƽ��,37.388356,121.60151;"
    AddressDistrict = AddressDistrict & "370600,370613,��ɽ��,37.473549,121.448866;"
    AddressDistrict = AddressDistrict & "370600,370614,������,37.811045,120.759074;"
    AddressDistrict = AddressDistrict & "370600,370681,������,37.648446,120.528328;"
    AddressDistrict = AddressDistrict & "370600,370682,������,36.977037,120.711151;"
    AddressDistrict = AddressDistrict & "370600,370683,������,37.182725,119.942135;"
    AddressDistrict = AddressDistrict & "370600,370685,��Զ��,37.364919,120.403142;"
    AddressDistrict = AddressDistrict & "370600,370686,��ϼ��,37.305854,120.834097;"
    AddressDistrict = AddressDistrict & "370600,370687,������,36.780657,121.168392;"
    AddressDistrict = AddressDistrict & "370700,370702,Ϋ����,36.710062,119.103784;"
    AddressDistrict = AddressDistrict & "370700,370703,��ͤ��,36.772103,119.207866;"
    AddressDistrict = AddressDistrict & "370700,370704,������,36.654616,119.166326;"
    AddressDistrict = AddressDistrict & "370700,370705,������,36.709494,119.137357;"
    AddressDistrict = AddressDistrict & "370700,370724,������,36.516371,118.539876;"
    AddressDistrict = AddressDistrict & "370700,370725,������,36.703253,118.839995;"
    AddressDistrict = AddressDistrict & "370700,370781,������,36.697855,118.484693;"
    AddressDistrict = AddressDistrict & "370700,370782,�����,35.997093,119.403182;"
    AddressDistrict = AddressDistrict & "370700,370783,�ٹ���,36.874411,118.736451;"
    AddressDistrict = AddressDistrict & "370700,370784,������,36.427417,119.206886;"
    AddressDistrict = AddressDistrict & "370700,370785,������,36.37754,119.757033;"
    AddressDistrict = AddressDistrict & "370700,370786,������,36.854937,119.394502;"
    AddressDistrict = AddressDistrict & "370800,370811,�γ���,35.414828,116.595261;"
    AddressDistrict = AddressDistrict & "370800,370812,������,35.556445,116.828996;"
    AddressDistrict = AddressDistrict & "370800,370826,΢ɽ��,34.809525,117.12861;"
    AddressDistrict = AddressDistrict & "370800,370827,��̨��,34.997706,116.650023;"
    AddressDistrict = AddressDistrict & "370800,370828,������,35.06977,116.310364;"
    AddressDistrict = AddressDistrict & "370800,370829,������,35.398098,116.342885;"
    AddressDistrict = AddressDistrict & "370800,370830,������,35.721746,116.487146;"
    AddressDistrict = AddressDistrict & "370800,370831,��ˮ��,35.653216,117.273605;"
    AddressDistrict = AddressDistrict & "370800,370832,��ɽ��,35.801843,116.08963;"
    AddressDistrict = AddressDistrict & "370800,370881,������,35.592788,116.991885;"
    AddressDistrict = AddressDistrict & "370800,370883,�޳���,35.405259,116.96673;"
    AddressDistrict = AddressDistrict & "370900,370902,̩ɽ��,36.189313,117.129984;"
    AddressDistrict = AddressDistrict & "370900,370911,�����,36.1841,117.04353;"
    AddressDistrict = AddressDistrict & "370900,370921,������,35.76754,116.799297;"
    AddressDistrict = AddressDistrict & "370900,370923,��ƽ��,35.930467,116.461052;"
    AddressDistrict = AddressDistrict & "370900,370982,��̩��,35.910387,117.766092;"
    AddressDistrict = AddressDistrict & "370900,370983,�ʳ���,36.1856,116.763703;"
    AddressDistrict = AddressDistrict & "371000,371002,������,37.510754,122.116189;"
    AddressDistrict = AddressDistrict & "371000,371003,�ĵ���,37.196211,122.057139;"
    AddressDistrict = AddressDistrict & "371000,371082,�ٳ���,37.160134,122.422896;"
    AddressDistrict = AddressDistrict & "371000,371083,��ɽ��,36.919622,121.536346;"
    AddressDistrict = AddressDistrict & "371100,371102,������,35.426152,119.457703;"
    AddressDistrict = AddressDistrict & "371100,371103,�ɽ��,35.119794,119.315844;"
    AddressDistrict = AddressDistrict & "371100,371121,������,35.751936,119.206745;"
    AddressDistrict = AddressDistrict & "371100,371122,����,35.588115,118.832859;"
    AddressDistrict = AddressDistrict & "371300,371302,��ɽ��,35.061631,118.327667;"
    AddressDistrict = AddressDistrict & "371300,371311,��ׯ��,34.997204,118.284795;"
    AddressDistrict = AddressDistrict & "371300,371312,�Ӷ���,35.085004,118.398296;"
    AddressDistrict = AddressDistrict & "371300,371321,������,35.547002,118.455395;"
    AddressDistrict = AddressDistrict & "371300,371322,۰����,34.614741,118.342963;"
    AddressDistrict = AddressDistrict & "371300,371323,��ˮ��,35.787029,118.634543;"
    AddressDistrict = AddressDistrict & "371300,371324,������,34.855573,118.049968;"
    AddressDistrict = AddressDistrict & "371300,371325,����,35.269174,117.968869;"
    AddressDistrict = AddressDistrict & "371300,371326,ƽ����,35.511519,117.631884;"
    AddressDistrict = AddressDistrict & "371300,371327,������,35.175911,118.838322;"
    AddressDistrict = AddressDistrict & "371300,371328,������,35.712435,117.943271;"
    AddressDistrict = AddressDistrict & "371300,371329,������,34.917062,118.648379;"
    AddressDistrict = AddressDistrict & "371400,371402,�³���,37.453923,116.307076;"
    AddressDistrict = AddressDistrict & "371400,371403,�����,37.332848,116.574929;"
    AddressDistrict = AddressDistrict & "371400,371422,������,37.649619,116.79372;"
    AddressDistrict = AddressDistrict & "371400,371423,������,37.777724,117.390507;"
    AddressDistrict = AddressDistrict & "371400,371424,������,37.192044,116.867028;"
    AddressDistrict = AddressDistrict & "371400,371425,�����,36.795497,116.758394;"
    AddressDistrict = AddressDistrict & "371400,371426,ƽԭ��,37.164465,116.433904;"
    AddressDistrict = AddressDistrict & "371400,371427,�Ľ���,36.950501,116.003816;"
    AddressDistrict = AddressDistrict & "371400,371428,�����,37.209527,116.078627;"
    AddressDistrict = AddressDistrict & "371400,371481,������,37.729115,117.216657;"
    AddressDistrict = AddressDistrict & "371400,371482,�����,36.934485,116.642554;"
    AddressDistrict = AddressDistrict & "371500,371502,��������,36.45606,115.980023;"
    AddressDistrict = AddressDistrict & "371500,371503,��ƽ��,36.591934,116.25335;"
    AddressDistrict = AddressDistrict & "371500,371521,������,36.113708,115.784287;"
    AddressDistrict = AddressDistrict & "371500,371522,ݷ��,36.237597,115.667291;"
    AddressDistrict = AddressDistrict & "371500,371524,������,36.336004,116.248855;"
    AddressDistrict = AddressDistrict & "371500,371525,����,36.483753,115.444808;"
    AddressDistrict = AddressDistrict & "371500,371526,������,36.859755,116.229662;"
    AddressDistrict = AddressDistrict & "371500,371581,������,36.842598,115.713462;"
    AddressDistrict = AddressDistrict & "371600,371602,������,37.384842,118.020149;"
    AddressDistrict = AddressDistrict & "371600,371603,մ����,37.698456,118.129902;"
    AddressDistrict = AddressDistrict & "371600,371621,������,37.483876,117.508941;"
    AddressDistrict = AddressDistrict & "371600,371622,������,37.640492,117.581326;"
    AddressDistrict = AddressDistrict & "371600,371623,�����,37.740848,117.616325;"
    AddressDistrict = AddressDistrict & "371600,371625,������,37.147002,118.123096;"
    AddressDistrict = AddressDistrict & "371600,371681,��ƽ��,36.87803,117.736807;"
    AddressDistrict = AddressDistrict & "371700,371702,ĵ����,35.24311,115.470946;"
    AddressDistrict = AddressDistrict & "371700,371703,������,35.072701,115.569601;"
    AddressDistrict = AddressDistrict & "371700,371721,����,34.823253,115.549482;"
    AddressDistrict = AddressDistrict & "371700,371722,����,34.790851,116.08262;"
    AddressDistrict = AddressDistrict & "371700,371723,������,34.947366,115.897349;"
    AddressDistrict = AddressDistrict & "371700,371724,��Ұ��,35.390999,116.089341;"
    AddressDistrict = AddressDistrict & "371700,371725,۩����,35.594773,115.93885;"
    AddressDistrict = AddressDistrict & "371700,371726,۲����,35.560257,115.51434;"
    AddressDistrict = AddressDistrict & "371700,371728,������,35.289637,115.098412;"
    AddressDistrict = AddressDistrict & "410100,410102,��ԭ��,34.748286,113.611576;"
    AddressDistrict = AddressDistrict & "410100,410103,������,34.730936,113.645422;"
    AddressDistrict = AddressDistrict & "410100,410104,�ܳǻ�����,34.746453,113.685313;"
    AddressDistrict = AddressDistrict & "410100,410105,��ˮ��,34.775838,113.686037;"
    AddressDistrict = AddressDistrict & "410100,410106,�Ͻ���,34.808689,113.298282;"
    AddressDistrict = AddressDistrict & "410100,410108,�ݼ���,34.828591,113.61836;"
    AddressDistrict = AddressDistrict & "410100,410122,��Ĳ��,34.721976,114.022521;"
    AddressDistrict = AddressDistrict & "410100,410181,������,34.75218,112.98283;"
    AddressDistrict = AddressDistrict & "410100,410182,������,34.789077,113.391523;"
    AddressDistrict = AddressDistrict & "410100,410183,������,34.537846,113.380616;"
    AddressDistrict = AddressDistrict & "410100,410184,��֣��,34.394219,113.73967;"
    AddressDistrict = AddressDistrict & "410100,410185,�Ƿ���,34.459939,113.037768;"
    AddressDistrict = AddressDistrict & "410200,410202,��ͤ��,34.799833,114.353348;"
    AddressDistrict = AddressDistrict & "410200,410203,˳�ӻ�����,34.800459,114.364875;"
    AddressDistrict = AddressDistrict & "410200,410204,��¥��,34.792383,114.3485;"
    AddressDistrict = AddressDistrict & "410200,410205,����̨��,34.779727,114.350246;"
    AddressDistrict = AddressDistrict & "410200,410212,�����,34.756476,114.437622;"
    AddressDistrict = AddressDistrict & "410200,410221,���,34.554585,114.770472;"
    AddressDistrict = AddressDistrict & "410200,410222,ͨ����,34.477302,114.467734;"
    AddressDistrict = AddressDistrict & "410200,410223,ξ����,34.412256,114.193927;"
    AddressDistrict = AddressDistrict & "410200,410225,������,34.829899,114.820572;"
    AddressDistrict = AddressDistrict & "410300,410302,�ϳ���,34.682945,112.477298;"
    AddressDistrict = AddressDistrict & "410300,410303,������,34.667847,112.443232;"
    AddressDistrict = AddressDistrict & "410300,410304,�e�ӻ�����,34.684738,112.491625;"
    AddressDistrict = AddressDistrict & "410300,410305,������,34.654251,112.399243;"
    AddressDistrict = AddressDistrict & "410300,410306,�Ͻ���,34.826485,112.443892;"
    AddressDistrict = AddressDistrict & "410300,410311,������,34.618557,112.456634;"
    AddressDistrict = AddressDistrict & "410300,410323,�°���,34.728679,112.141403;"
    AddressDistrict = AddressDistrict & "410300,410324,�ﴨ��,33.783195,111.618386;"
    AddressDistrict = AddressDistrict & "410300,410325,����,34.131563,112.087765;"
    AddressDistrict = AddressDistrict & "410300,410326,������,34.15323,112.473789;"
    AddressDistrict = AddressDistrict & "410300,410327,������,34.516478,112.179989;"
    AddressDistrict = AddressDistrict & "410300,410328,������,34.387179,111.655399;"
    AddressDistrict = AddressDistrict & "410300,410329,������,34.423416,112.429384;"
    AddressDistrict = AddressDistrict & "410300,410381,��ʦ��,34.723042,112.787739;"
    AddressDistrict = AddressDistrict & "410400,410402,�»���,33.737579,113.299061;"
    AddressDistrict = AddressDistrict & "410400,410403,������,33.739285,113.310327;"
    AddressDistrict = AddressDistrict & "410400,410404,ʯ����,33.901538,112.889885;"
    AddressDistrict = AddressDistrict & "410400,410411,տ����,33.725681,113.320873;"
    AddressDistrict = AddressDistrict & "410400,410421,������,33.866359,113.066812;"
    AddressDistrict = AddressDistrict & "410400,410422,Ҷ��,33.621252,113.358298;"
    AddressDistrict = AddressDistrict & "410400,410423,³ɽ��,33.740325,112.906703;"
    AddressDistrict = AddressDistrict & "410400,410425,ۣ��,33.971993,113.220451;"
    AddressDistrict = AddressDistrict & "410400,410481,�����,33.302082,113.52625;"
    AddressDistrict = AddressDistrict & "410400,410482,������,34.167408,112.845336;"
    AddressDistrict = AddressDistrict & "410500,410502,�ķ���,36.098101,114.352562;"
    AddressDistrict = AddressDistrict & "410500,410503,������,36.10978,114.352646;"
    AddressDistrict = AddressDistrict & "410500,410505,����,36.108974,114.300098;"
    AddressDistrict = AddressDistrict & "410500,410506,������,36.095568,114.323522;"
    AddressDistrict = AddressDistrict & "410500,410522,������,36.130585,114.130207;"
    AddressDistrict = AddressDistrict & "410500,410523,������,35.922349,114.362357;"
    AddressDistrict = AddressDistrict & "410500,410526,����,35.574628,114.524;"
    AddressDistrict = AddressDistrict & "410500,410527,�ڻ���,35.953702,114.904582;"
    AddressDistrict = AddressDistrict & "410500,410581,������,36.063403,113.823767;"
    AddressDistrict = AddressDistrict & "410600,410602,��ɽ��,35.936128,114.166551;"
    AddressDistrict = AddressDistrict & "410600,410603,ɽ����,35.896058,114.184202;"
    AddressDistrict = AddressDistrict & "410600,410611,俱���,35.748382,114.293917;"
    AddressDistrict = AddressDistrict & "410600,410621,����,35.671282,114.550162;"
    AddressDistrict = AddressDistrict & "410600,410622,���,35.609478,114.200379;"
    AddressDistrict = AddressDistrict & "410700,410702,������,35.302684,113.878158;"
    AddressDistrict = AddressDistrict & "410700,410703,������,35.304905,113.866065;"
    AddressDistrict = AddressDistrict & "410700,410704,��Ȫ��,35.379855,113.906712;"
    AddressDistrict = AddressDistrict & "410700,410711,��Ұ��,35.312974,113.89716;"
    AddressDistrict = AddressDistrict & "410700,410721,������,35.190021,113.806186;"
    AddressDistrict = AddressDistrict & "410700,410724,�����,35.261685,113.657249;"
    AddressDistrict = AddressDistrict & "410700,410725,ԭ����,35.054001,113.965966;"
    AddressDistrict = AddressDistrict & "410700,410726,�ӽ���,35.149515,114.200982;"
    AddressDistrict = AddressDistrict & "410700,410727,������,35.04057,114.423405;"
    AddressDistrict = AddressDistrict & "410700,410781,������,35.404295,114.065855;"
    AddressDistrict = AddressDistrict & "410700,410782,������,35.461318,113.802518;"
    AddressDistrict = AddressDistrict & "410700,410783,��ԫ��,35.19615,114.673807;"
    AddressDistrict = AddressDistrict & "410800,410802,�����,35.241353,113.226126;"
    AddressDistrict = AddressDistrict & "410800,410803,��վ��,35.236145,113.175485;"
    AddressDistrict = AddressDistrict & "410800,410804,�����,35.265453,113.321703;"
    AddressDistrict = AddressDistrict & "410800,410811,ɽ����,35.21476,113.26766;"
    AddressDistrict = AddressDistrict & "410800,410821,������,35.229923,113.447465;"
    AddressDistrict = AddressDistrict & "410800,410822,������,35.170351,113.069313;"
    AddressDistrict = AddressDistrict & "410800,410823,������,35.09885,113.408334;"
    AddressDistrict = AddressDistrict & "410800,410825,����,34.941233,113.079118;"
    AddressDistrict = AddressDistrict & "410800,410882,������,35.08901,112.934538;"
    AddressDistrict = AddressDistrict & "410800,410883,������,34.90963,112.78708;"
    AddressDistrict = AddressDistrict & "410900,410902,������,35.760473,115.03184;"
    AddressDistrict = AddressDistrict & "410900,410922,�����,35.902413,115.107287;"
    AddressDistrict = AddressDistrict & "410900,410923,������,36.075204,115.204336;"
    AddressDistrict = AddressDistrict & "410900,410926,����,35.851977,115.504212;"
    AddressDistrict = AddressDistrict & "410900,410927,̨ǰ��,35.996474,115.855681;"
    AddressDistrict = AddressDistrict & "410900,410928,�����,35.710349,115.023844;"
    AddressDistrict = AddressDistrict & "411000,411002,κ����,34.02711,113.828307;"
    AddressDistrict = AddressDistrict & "411000,411003,������,34.005018,113.842898;"
    AddressDistrict = AddressDistrict & "411000,411024,۳����,34.100502,114.188507;"
    AddressDistrict = AddressDistrict & "411000,411025,�����,33.855943,113.493166;"
    AddressDistrict = AddressDistrict & "411000,411081,������,34.154403,113.471316;"
    AddressDistrict = AddressDistrict & "411000,411082,������,34.219257,113.768912;"
    AddressDistrict = AddressDistrict & "411100,411102,Դ����,33.565441,114.017948;"
    AddressDistrict = AddressDistrict & "411100,411103,۱����,33.588897,114.016813;"
    AddressDistrict = AddressDistrict & "411100,411104,������,33.567555,114.051686;"
    AddressDistrict = AddressDistrict & "411100,411121,������,33.436278,113.610565;"
    AddressDistrict = AddressDistrict & "411100,411122,�����,33.80609,113.938891;"
    AddressDistrict = AddressDistrict & "411200,411202,������,34.77812,111.19487;"
    AddressDistrict = AddressDistrict & "411200,411203,������,34.720244,111.103851;"
    AddressDistrict = AddressDistrict & "411200,411221,�ų���,34.763487,111.762992;"
    AddressDistrict = AddressDistrict & "411200,411224,¬����,34.053995,111.052649;"
    AddressDistrict = AddressDistrict & "411200,411281,������,34.746868,111.869417;"
    AddressDistrict = AddressDistrict & "411200,411282,�鱦��,34.521264,110.88577;"
    AddressDistrict = AddressDistrict & "411300,411302,�����,32.994857,112.544591;"
    AddressDistrict = AddressDistrict & "411300,411303,������,32.989877,112.528789;"
    AddressDistrict = AddressDistrict & "411300,411321,������,33.488617,112.435583;"
    AddressDistrict = AddressDistrict & "411300,411322,������,33.255138,113.010933;"
    AddressDistrict = AddressDistrict & "411300,411323,��Ͽ��,33.302981,111.485772;"
    AddressDistrict = AddressDistrict & "411300,411324,��ƽ��,33.036651,112.232722;"
    AddressDistrict = AddressDistrict & "411300,411325,������,33.046358,111.843801;"
    AddressDistrict = AddressDistrict & "411300,411326,������,33.136106,111.489026;"
    AddressDistrict = AddressDistrict & "411300,411327,������,33.056126,112.938279;"
    AddressDistrict = AddressDistrict & "411300,411328,�ƺ���,32.687892,112.838492;"
    AddressDistrict = AddressDistrict & "411300,411329,��Ұ��,32.524006,112.365624;"
    AddressDistrict = AddressDistrict & "411300,411330,ͩ����,32.367153,113.406059;"
    AddressDistrict = AddressDistrict & "411300,411381,������,32.681642,112.092716;"
    AddressDistrict = AddressDistrict & "411400,411402,��԰��,34.436553,115.65459;"
    AddressDistrict = AddressDistrict & "411400,411403,�����,34.390536,115.653813;"
    AddressDistrict = AddressDistrict & "411400,411421,��Ȩ��,34.648455,115.148146;"
    AddressDistrict = AddressDistrict & "411400,411422,���,34.428433,115.070109;"
    AddressDistrict = AddressDistrict & "411400,411423,������,34.449299,115.320055;"
    AddressDistrict = AddressDistrict & "411400,411424,�ϳ���,34.075277,115.307433;"
    AddressDistrict = AddressDistrict & "411400,411425,�ݳ���,34.399634,115.863811;"
    AddressDistrict = AddressDistrict & "411400,411426,������,34.240894,116.13989;"
    AddressDistrict = AddressDistrict & "411400,411481,������,33.931318,116.449672;"
    AddressDistrict = AddressDistrict & "411500,411502,������,32.123274,114.075031;"
    AddressDistrict = AddressDistrict & "411500,411503,ƽ����,32.098395,114.126027;"
    AddressDistrict = AddressDistrict & "411500,411521,��ɽ��,32.203206,114.533414;"
    AddressDistrict = AddressDistrict & "411500,411522,��ɽ��,32.010398,114.903577;"
    AddressDistrict = AddressDistrict & "411500,411523,����,31.63515,114.87705;"
    AddressDistrict = AddressDistrict & "411500,411524,�̳���,31.799982,115.406297;"
    AddressDistrict = AddressDistrict & "411500,411525,��ʼ��,32.183074,115.667328;"
    AddressDistrict = AddressDistrict & "411500,411526,�괨��,32.134024,115.050123;"
    AddressDistrict = AddressDistrict & "411500,411527,������,32.452639,115.415451;"
    AddressDistrict = AddressDistrict & "411500,411528,Ϣ��,32.344744,114.740713;"
    AddressDistrict = AddressDistrict & "411600,411602,������,33.614836,114.652136;"
    AddressDistrict = AddressDistrict & "411600,411603,������,33.732547,114.870166;"
    AddressDistrict = AddressDistrict & "411600,411621,������,34.054061,114.392008;"
    AddressDistrict = AddressDistrict & "411600,411622,������,33.784378,114.530067;"
    AddressDistrict = AddressDistrict & "411600,411623,��ˮ��,33.543845,114.60927;"
    AddressDistrict = AddressDistrict & "411600,411624,������,33.395514,115.078375;"
    AddressDistrict = AddressDistrict & "411600,411625,������,33.643852,115.189;"
    AddressDistrict = AddressDistrict & "411600,411627,̫����,34.065312,114.853834;"
    AddressDistrict = AddressDistrict & "411600,411628,¹����,33.861067,115.486386;"
    AddressDistrict = AddressDistrict & "411600,411681,�����,33.443085,114.899521;"
    AddressDistrict = AddressDistrict & "411700,411702,�����,32.977559,114.029149;"
    AddressDistrict = AddressDistrict & "411700,411721,��ƽ��,33.382315,114.026864;"
    AddressDistrict = AddressDistrict & "411700,411722,�ϲ���,33.264719,114.266892;"
    AddressDistrict = AddressDistrict & "411700,411723,ƽ����,32.955626,114.637105;"
    AddressDistrict = AddressDistrict & "411700,411724,������,32.601826,114.38948;"
    AddressDistrict = AddressDistrict & "411700,411725,ȷɽ��,32.801538,114.026679;"
    AddressDistrict = AddressDistrict & "411700,411726,������,32.725129,113.32605;"
    AddressDistrict = AddressDistrict & "411700,411727,������,33.004535,114.359495;"
    AddressDistrict = AddressDistrict & "411700,411728,��ƽ��,33.14698,114.00371;"
    AddressDistrict = AddressDistrict & "411700,411729,�²���,32.749948,114.975246;"
    AddressDistrict = AddressDistrict & "420100,420102,������,30.594911,114.30304;"
    AddressDistrict = AddressDistrict & "420100,420103,������,30.578771,114.283109;"
    AddressDistrict = AddressDistrict & "420100,420104,�~����,30.57061,114.264568;"
    AddressDistrict = AddressDistrict & "420100,420105,������,30.549326,114.265807;"
    AddressDistrict = AddressDistrict & "420100,420106,�����,30.546536,114.307344;"
    AddressDistrict = AddressDistrict & "420100,420107,��ɽ��,30.634215,114.39707;"
    AddressDistrict = AddressDistrict & "420100,420111,��ɽ��,30.504259,114.400718;"
    AddressDistrict = AddressDistrict & "420100,420112,��������,30.622467,114.142483;"
    AddressDistrict = AddressDistrict & "420100,420113,������,30.309637,114.08124;"
    AddressDistrict = AddressDistrict & "420100,420114,�̵���,30.582186,114.029341;"
    AddressDistrict = AddressDistrict & "420100,420115,������,30.349045,114.313961;"
    AddressDistrict = AddressDistrict & "420100,420116,������,30.874155,114.374025;"
    AddressDistrict = AddressDistrict & "420100,420117,������,30.842149,114.802108;"
    AddressDistrict = AddressDistrict & "420200,420202,��ʯ����,30.212086,115.090164;"
    AddressDistrict = AddressDistrict & "420200,420203,����ɽ��,30.205365,115.093354;"
    AddressDistrict = AddressDistrict & "420200,420204,��½��,30.177845,114.975755;"
    AddressDistrict = AddressDistrict & "420200,420205,��ɽ��,30.20601,114.901366;"
    AddressDistrict = AddressDistrict & "420200,420222,������,29.841572,115.212883;"
    AddressDistrict = AddressDistrict & "420200,420281,��ұ��,30.098804,114.974842;"
    AddressDistrict = AddressDistrict & "420300,420302,é����,32.644463,110.78621;"
    AddressDistrict = AddressDistrict & "420300,420303,������,32.652516,110.772365;"
    AddressDistrict = AddressDistrict & "420300,420304,������,32.838267,110.812099;"
    AddressDistrict = AddressDistrict & "420300,420322,������,32.991457,110.426472;"
    AddressDistrict = AddressDistrict & "420300,420323,��ɽ��,32.22586,110.2296;"
    AddressDistrict = AddressDistrict & "420300,420324,��Ϫ��,32.315342,109.717196;"
    AddressDistrict = AddressDistrict & "420300,420325,����,32.055002,110.741966;"
    AddressDistrict = AddressDistrict & "420300,420381,��������,32.538839,111.513793;"
    AddressDistrict = AddressDistrict & "420500,420502,������,30.702476,111.295468;"
    AddressDistrict = AddressDistrict & "420500,420503,��Ҹ���,30.679053,111.307215;"
    AddressDistrict = AddressDistrict & "420500,420504,�����,30.692322,111.268163;"
    AddressDistrict = AddressDistrict & "420500,420505,�Vͤ��,30.530744,111.427642;"
    AddressDistrict = AddressDistrict & "420500,420506,������,30.770199,111.326747;"
    AddressDistrict = AddressDistrict & "420500,420525,Զ����,31.059626,111.64331;"
    AddressDistrict = AddressDistrict & "420500,420526,��ɽ��,31.34795,110.754499;"
    AddressDistrict = AddressDistrict & "420500,420527,������,30.823908,110.976785;"
    AddressDistrict = AddressDistrict & "420500,420528,����������������,30.466534,111.198475;"
    AddressDistrict = AddressDistrict & "420500,420529,���������������,30.199252,110.674938;"
    AddressDistrict = AddressDistrict & "420500,420581,�˶���,30.387234,111.454367;"
    AddressDistrict = AddressDistrict & "420500,420582,������,30.824492,111.793419;"
    AddressDistrict = AddressDistrict & "420500,420583,֦����,30.425364,111.751799;"
    AddressDistrict = AddressDistrict & "420600,420602,�����,32.015088,112.150327;"
    AddressDistrict = AddressDistrict & "420600,420606,������,32.058589,112.13957;"
    AddressDistrict = AddressDistrict & "420600,420607,������,32.085517,112.197378;"
    AddressDistrict = AddressDistrict & "420600,420624,������,31.77692,111.844424;"
    AddressDistrict = AddressDistrict & "420600,420625,�ȳ���,32.262676,111.640147;"
    AddressDistrict = AddressDistrict & "420600,420626,������,31.873507,111.262235;"
    AddressDistrict = AddressDistrict & "420600,420682,�Ϻӿ���,32.385438,111.675732;"
    AddressDistrict = AddressDistrict & "420600,420683,������,32.123083,112.765268;"
    AddressDistrict = AddressDistrict & "420600,420684,�˳���,31.709203,112.261441;"
    AddressDistrict = AddressDistrict & "420700,420702,���Ӻ���,30.098191,114.681967;"
    AddressDistrict = AddressDistrict & "420700,420703,������,30.534468,114.74148;"
    AddressDistrict = AddressDistrict & "420700,420704,������,30.39669,114.890012;"
    AddressDistrict = AddressDistrict & "420800,420802,������,31.033461,112.204804;"
    AddressDistrict = AddressDistrict & "420800,420804,�޵���,30.980798,112.198413;"
    AddressDistrict = AddressDistrict & "420800,420822,ɳ����,30.70359,112.595218;"
    AddressDistrict = AddressDistrict & "420800,420881,������,31.165573,112.587267;"
    AddressDistrict = AddressDistrict & "420800,420882,��ɽ��,31.022457,113.114595;"
    AddressDistrict = AddressDistrict & "420900,420902,Т����,30.925966,113.925849;"
    AddressDistrict = AddressDistrict & "420900,420921,Т����,31.251618,113.988964;"
    AddressDistrict = AddressDistrict & "420900,420922,������,31.565483,114.126249;"
    AddressDistrict = AddressDistrict & "420900,420923,������,31.021691,113.750616;"
    AddressDistrict = AddressDistrict & "420900,420981,Ӧ����,30.939038,113.573842;"
    AddressDistrict = AddressDistrict & "420900,420982,��½��,31.26174,113.690401;"
    AddressDistrict = AddressDistrict & "420900,420984,������,30.652165,113.835301;"
    AddressDistrict = AddressDistrict & "421000,421002,ɳ����,30.315895,112.257433;"
    AddressDistrict = AddressDistrict & "421000,421003,������,30.350674,112.195354;"
    AddressDistrict = AddressDistrict & "421000,421022,������,30.059065,112.230179;"
    AddressDistrict = AddressDistrict & "421000,421023,������,29.820079,112.904344;"
    AddressDistrict = AddressDistrict & "421000,421024,������,30.033919,112.41735;"
    AddressDistrict = AddressDistrict & "421000,421081,ʯ����,29.716437,112.40887;"
    AddressDistrict = AddressDistrict & "421000,421083,�����,29.81297,113.470304;"
    AddressDistrict = AddressDistrict & "421000,421087,������,30.176037,111.77818;"
    AddressDistrict = AddressDistrict & "421100,421102,������,30.447435,114.878934;"
    AddressDistrict = AddressDistrict & "421100,421121,�ŷ���,30.63569,114.872029;"
    AddressDistrict = AddressDistrict & "421100,421122,�찲��,31.284777,114.615095;"
    AddressDistrict = AddressDistrict & "421100,421123,������,30.781679,115.398984;"
    AddressDistrict = AddressDistrict & "421100,421124,Ӣɽ��,30.735794,115.67753;"
    AddressDistrict = AddressDistrict & "421100,421125,�ˮ��,30.454837,115.26344;"
    AddressDistrict = AddressDistrict & "421100,421126,ޭ����,30.234927,115.433964;"
    AddressDistrict = AddressDistrict & "421100,421127,��÷��,30.075113,115.942548;"
    AddressDistrict = AddressDistrict & "421100,421181,�����,31.177906,115.02541;"
    AddressDistrict = AddressDistrict & "421100,421182,��Ѩ��,29.849342,115.56242;"
    AddressDistrict = AddressDistrict & "421200,421202,�̰���,29.824716,114.333894;"
    AddressDistrict = AddressDistrict & "421200,421221,������,29.973363,113.921547;"
    AddressDistrict = AddressDistrict & "421200,421222,ͨ����,29.246076,113.814131;"
    AddressDistrict = AddressDistrict & "421200,421223,������,29.54101,114.049958;"
    AddressDistrict = AddressDistrict & "421200,421224,ͨɽ��,29.604455,114.493163;"
    AddressDistrict = AddressDistrict & "421200,421281,�����,29.716879,113.88366;"
    AddressDistrict = AddressDistrict & "421300,421303,������,31.717521,113.374519;"
    AddressDistrict = AddressDistrict & "421300,421321,����,31.854246,113.301384;"
    AddressDistrict = AddressDistrict & "421300,421381,��ˮ��,31.617731,113.826601;"
    AddressDistrict = AddressDistrict & "422800,422801,��ʩ��,30.282406,109.486761;"
    AddressDistrict = AddressDistrict & "422800,422802,������,30.294247,108.943491;"
    AddressDistrict = AddressDistrict & "422800,422822,��ʼ��,30.601632,109.723822;"
    AddressDistrict = AddressDistrict & "422800,422823,�Ͷ���,31.041403,110.336665;"
    AddressDistrict = AddressDistrict & "422800,422825,������,29.98867,109.482819;"
    AddressDistrict = AddressDistrict & "422800,422826,�̷���,29.678967,109.15041;"
    AddressDistrict = AddressDistrict & "422800,422827,������,29.506945,109.408328;"
    AddressDistrict = AddressDistrict & "422800,422828,�׷���,29.887298,110.033699;"
    AddressDistrict = AddressDistrict & "430100,430102,ܽ����,28.193106,112.988094;"
    AddressDistrict = AddressDistrict & "430100,430103,������,28.192375,112.97307;"
    AddressDistrict = AddressDistrict & "430100,430104,��´��,28.213044,112.911591;"
    AddressDistrict = AddressDistrict & "430100,430105,������,28.201336,112.985525;"
    AddressDistrict = AddressDistrict & "430100,430111,�껨��,28.109937,113.016337;"
    AddressDistrict = AddressDistrict & "430100,430112,������,28.347458,112.819549;"
    AddressDistrict = AddressDistrict & "430100,430121,��ɳ��,28.237888,113.080098;"
    AddressDistrict = AddressDistrict & "430100,430181,�����,28.141112,113.633301;"
    AddressDistrict = AddressDistrict & "430100,430182,������,28.253928,112.553182;"
    AddressDistrict = AddressDistrict & "430200,430202,������,27.833036,113.162548;"
    AddressDistrict = AddressDistrict & "430200,430203,«����,27.827246,113.155169;"
    AddressDistrict = AddressDistrict & "430200,430204,ʯ����,27.871945,113.11295;"
    AddressDistrict = AddressDistrict & "430200,430211,��Ԫ��,27.826909,113.136252;"
    AddressDistrict = AddressDistrict & "430200,430212,�˿���,27.705844,113.146175;"
    AddressDistrict = AddressDistrict & "430200,430223,����,27.000071,113.345774;"
    AddressDistrict = AddressDistrict & "430200,430224,������,26.789534,113.546509;"
    AddressDistrict = AddressDistrict & "430200,430225,������,26.489459,113.776884;"
    AddressDistrict = AddressDistrict & "430200,430281,������,27.657873,113.507157;"
    AddressDistrict = AddressDistrict & "430300,430302,�����,27.86077,112.907427;"
    AddressDistrict = AddressDistrict & "430300,430304,������,27.828854,112.927707;"
    AddressDistrict = AddressDistrict & "430300,430321,��̶��,27.778601,112.952829;"
    AddressDistrict = AddressDistrict & "430300,430381,������,27.734918,112.525217;"
    AddressDistrict = AddressDistrict & "430300,430382,��ɽ��,27.922682,112.52848;"
    AddressDistrict = AddressDistrict & "430400,430405,������,26.891063,112.626324;"
    AddressDistrict = AddressDistrict & "430400,430406,�����,26.893694,112.612241;"
    AddressDistrict = AddressDistrict & "430400,430407,ʯ����,26.903908,112.607635;"
    AddressDistrict = AddressDistrict & "430400,430408,������,26.89087,112.570608;"
    AddressDistrict = AddressDistrict & "430400,430412,������,27.240536,112.734147;"
    AddressDistrict = AddressDistrict & "430400,430421,������,26.962388,112.379643;"
    AddressDistrict = AddressDistrict & "430400,430422,������,26.739973,112.677459;"
    AddressDistrict = AddressDistrict & "430400,430423,��ɽ��,27.234808,112.86971;"
    AddressDistrict = AddressDistrict & "430400,430424,�ⶫ��,27.083531,112.950412;"
    AddressDistrict = AddressDistrict & "430400,430426,���,26.787109,112.111192;"
    AddressDistrict = AddressDistrict & "430400,430481,������,26.414162,112.847215;"
    AddressDistrict = AddressDistrict & "430400,430482,������,26.406773,112.396821;"
    AddressDistrict = AddressDistrict & "430500,430502,˫����,27.240001,111.479756;"
    AddressDistrict = AddressDistrict & "430500,430503,������,27.233593,111.462968;"
    AddressDistrict = AddressDistrict & "430500,430511,������,27.245688,111.452315;"
    AddressDistrict = AddressDistrict & "430500,430522,������,27.311429,111.459762;"
    AddressDistrict = AddressDistrict & "430500,430523,������,26.989713,111.2757;"
    AddressDistrict = AddressDistrict & "430500,430524,¡����,27.116002,111.038785;"
    AddressDistrict = AddressDistrict & "430500,430525,������,27.062286,110.579212;"
    AddressDistrict = AddressDistrict & "430500,430527,������,26.580622,110.155075;"
    AddressDistrict = AddressDistrict & "430500,430528,������,26.438912,110.859115;"
    AddressDistrict = AddressDistrict & "430500,430529,�ǲ�����������,26.363575,110.313226;"
    AddressDistrict = AddressDistrict & "430500,430581,�����,26.732086,110.636804;"
    AddressDistrict = AddressDistrict & "430500,430582,�۶���,27.257273,111.743168;"
    AddressDistrict = AddressDistrict & "430600,430602,����¥��,29.366784,113.120751;"
    AddressDistrict = AddressDistrict & "430600,430603,��Ϫ��,29.473395,113.27387;"
    AddressDistrict = AddressDistrict & "430600,430611,��ɽ��,29.438062,113.004082;"
    AddressDistrict = AddressDistrict & "430600,430621,������,29.144843,113.116073;"
    AddressDistrict = AddressDistrict & "430600,430623,������,29.524107,112.559369;"
    AddressDistrict = AddressDistrict & "430600,430624,������,28.677498,112.889748;"
    AddressDistrict = AddressDistrict & "430600,430626,ƽ����,28.701523,113.593751;"
    AddressDistrict = AddressDistrict & "430600,430681,������,28.803149,113.079419;"
    AddressDistrict = AddressDistrict & "430600,430682,������,29.471594,113.450809;"
    AddressDistrict = AddressDistrict & "430700,430702,������,29.040477,111.690718;"
    AddressDistrict = AddressDistrict & "430700,430703,������,29.014426,111.685327;"
    AddressDistrict = AddressDistrict & "430700,430721,������,29.414483,112.172289;"
    AddressDistrict = AddressDistrict & "430700,430722,������,28.907319,111.968506;"
    AddressDistrict = AddressDistrict & "430700,430723,���,29.64264,111.761682;"
    AddressDistrict = AddressDistrict & "430700,430724,�����,29.443217,111.645602;"
    AddressDistrict = AddressDistrict & "430700,430725,��Դ��,28.902734,111.484503;"
    AddressDistrict = AddressDistrict & "430700,430726,ʯ����,29.584703,111.379087;"
    AddressDistrict = AddressDistrict & "430700,430781,������,29.630867,111.879609;"
    AddressDistrict = AddressDistrict & "430800,430802,������,29.125961,110.484559;"
    AddressDistrict = AddressDistrict & "430800,430811,����Դ��,29.347827,110.54758;"
    AddressDistrict = AddressDistrict & "430800,430821,������,29.423876,111.132702;"
    AddressDistrict = AddressDistrict & "430800,430822,ɣֲ��,29.399939,110.164039;"
    AddressDistrict = AddressDistrict & "430900,430902,������,28.592771,112.33084;"
    AddressDistrict = AddressDistrict & "430900,430903,��ɽ��,28.568327,112.360946;"
    AddressDistrict = AddressDistrict & "430900,430921,����,29.372181,112.410399;"
    AddressDistrict = AddressDistrict & "430900,430922,�ҽ���,28.520993,112.139732;"
    AddressDistrict = AddressDistrict & "430900,430923,������,28.377421,111.221824;"
    AddressDistrict = AddressDistrict & "430900,430981,�佭��,28.839713,112.361088;"
    AddressDistrict = AddressDistrict & "431000,431002,������,25.792628,113.032208;"
    AddressDistrict = AddressDistrict & "431000,431003,������,25.793157,113.038698;"
    AddressDistrict = AddressDistrict & "431000,431021,������,25.737447,112.734466;"
    AddressDistrict = AddressDistrict & "431000,431022,������,25.394345,112.947884;"
    AddressDistrict = AddressDistrict & "431000,431023,������,26.129392,113.114819;"
    AddressDistrict = AddressDistrict & "431000,431024,�κ���,25.587309,112.370618;"
    AddressDistrict = AddressDistrict & "431000,431025,������,25.279119,112.564589;"
    AddressDistrict = AddressDistrict & "431000,431026,�����,25.553759,113.685686;"
    AddressDistrict = AddressDistrict & "431000,431027,����,26.073917,113.945879;"
    AddressDistrict = AddressDistrict & "431000,431028,������,26.708625,113.27217;"
    AddressDistrict = AddressDistrict & "431000,431081,������,25.974152,113.23682;"
    AddressDistrict = AddressDistrict & "431100,431102,������,26.223347,111.626348;"
    AddressDistrict = AddressDistrict & "431100,431103,��ˮ̲��,26.434364,111.607156;"
    AddressDistrict = AddressDistrict & "431100,431121,������,26.585929,111.85734;"
    AddressDistrict = AddressDistrict & "431100,431122,������,26.397278,111.313035;"
    AddressDistrict = AddressDistrict & "431100,431123,˫����,25.959397,111.662146;"
    AddressDistrict = AddressDistrict & "431100,431124,����,25.518444,111.591614;"
    AddressDistrict = AddressDistrict & "431100,431125,������,25.268154,111.346803;"
    AddressDistrict = AddressDistrict & "431100,431126,��Զ��,25.584112,111.944529;"
    AddressDistrict = AddressDistrict & "431100,431127,��ɽ��,25.375255,112.194195;"
    AddressDistrict = AddressDistrict & "431100,431128,������,25.906927,112.220341;"
    AddressDistrict = AddressDistrict & "431100,431129,��������������,25.182596,111.577276;"
    AddressDistrict = AddressDistrict & "431200,431202,�׳���,27.548474,109.982242;"
    AddressDistrict = AddressDistrict & "431200,431221,�з���,27.43736,109.948061;"
    AddressDistrict = AddressDistrict & "431200,431222,������,28.455554,110.399161;"
    AddressDistrict = AddressDistrict & "431200,431223,��Ϫ��,28.005474,110.196953;"
    AddressDistrict = AddressDistrict & "431200,431224,������,27.903802,110.593373;"
    AddressDistrict = AddressDistrict & "431200,431225,��ͬ��,26.870789,109.720785;"
    AddressDistrict = AddressDistrict & "431200,431226,��������������,27.865991,109.802807;"
    AddressDistrict = AddressDistrict & "431200,431227,�»ζ���������,27.359897,109.174443;"
    AddressDistrict = AddressDistrict & "431200,431228,�ƽ�����������,27.437996,109.687777;"
    AddressDistrict = AddressDistrict & "431200,431229,�������嶱��������,26.573511,109.691159;"
    AddressDistrict = AddressDistrict & "431200,431230,ͨ������������,26.158349,109.783359;"
    AddressDistrict = AddressDistrict & "431200,431281,�齭��,27.201876,109.831765;"
    AddressDistrict = AddressDistrict & "431300,431302,¦����,27.726643,112.008486;"
    AddressDistrict = AddressDistrict & "431300,431321,˫����,27.459126,112.198245;"
    AddressDistrict = AddressDistrict & "431300,431322,�»���,27.737456,111.306747;"
    AddressDistrict = AddressDistrict & "431300,431381,��ˮ����,27.685759,111.434674;"
    AddressDistrict = AddressDistrict & "431300,431382,��Դ��,27.692301,111.670847;"
    AddressDistrict = AddressDistrict & "433100,433101,������,28.314827,109.738273;"
    AddressDistrict = AddressDistrict & "433100,433122,��Ϫ��,28.214516,110.214428;"
    AddressDistrict = AddressDistrict & "433100,433123,�����,27.948308,109.599191;"
    AddressDistrict = AddressDistrict & "433100,433124,��ԫ��,28.581352,109.479063;"
    AddressDistrict = AddressDistrict & "433100,433125,������,28.709605,109.651445;"
    AddressDistrict = AddressDistrict & "433100,433126,������,28.616973,109.949592;"
    AddressDistrict = AddressDistrict & "433100,433127,��˳��,28.998068,109.853292;"
    AddressDistrict = AddressDistrict & "433100,433130,��ɽ��,29.453438,109.441189;"
    AddressDistrict = AddressDistrict & "440100,440103,������,23.124943,113.243038;"
    AddressDistrict = AddressDistrict & "440100,440104,Խ����,23.125624,113.280714;"
    AddressDistrict = AddressDistrict & "440100,440105,������,23.103131,113.262008;"
    AddressDistrict = AddressDistrict & "440100,440106,�����,23.13559,113.335367;"
    AddressDistrict = AddressDistrict & "440100,440111,������,23.162281,113.262831;"
    AddressDistrict = AddressDistrict & "440100,440112,������,23.103239,113.450761;"
    AddressDistrict = AddressDistrict & "440100,440113,��خ��,22.938582,113.364619;"
    AddressDistrict = AddressDistrict & "440100,440114,������,23.39205,113.211184;"
    AddressDistrict = AddressDistrict & "440100,440115,��ɳ��,22.794531,113.53738;"
    AddressDistrict = AddressDistrict & "440100,440117,�ӻ���,23.545283,113.587386;"
    AddressDistrict = AddressDistrict & "440100,440118,������,23.290497,113.829579;"
    AddressDistrict = AddressDistrict & "440200,440203,�佭��,24.80016,113.588289;"
    AddressDistrict = AddressDistrict & "440200,440204,䥽���,24.803977,113.599224;"
    AddressDistrict = AddressDistrict & "440200,440205,������,24.680195,113.605582;"
    AddressDistrict = AddressDistrict & "440200,440222,ʼ����,24.948364,114.067205;"
    AddressDistrict = AddressDistrict & "440200,440224,�ʻ���,25.088226,113.748627;"
    AddressDistrict = AddressDistrict & "440200,440229,��Դ��,24.353887,114.131289;"
    AddressDistrict = AddressDistrict & "440200,440232,��Դ����������,24.776109,113.278417;"
    AddressDistrict = AddressDistrict & "440200,440233,�·���,24.055412,114.207034;"
    AddressDistrict = AddressDistrict & "440200,440281,�ֲ���,25.128445,113.352413;"
    AddressDistrict = AddressDistrict & "440200,440282,������,25.115328,114.311231;"
    AddressDistrict = AddressDistrict & "440300,440303,�޺���,22.555341,114.123885;"
    AddressDistrict = AddressDistrict & "440300,440304,������,22.541009,114.05096;"
    AddressDistrict = AddressDistrict & "440300,440305,��ɽ��,22.531221,113.92943;"
    AddressDistrict = AddressDistrict & "440300,440306,������,22.754741,113.828671;"
    AddressDistrict = AddressDistrict & "440300,440307,������,22.721511,114.251372;"
    AddressDistrict = AddressDistrict & "440300,440308,������,22.555069,114.235366;"
    AddressDistrict = AddressDistrict & "440300,440309,������,22.691963,114.044346;"
    AddressDistrict = AddressDistrict & "440300,440310,ƺɽ��,22.69423,114.338441;"
    AddressDistrict = AddressDistrict & "440300,440311,������,22.748816,113.935895;"
    AddressDistrict = AddressDistrict & "440400,440402,������,22.271249,113.55027;"
    AddressDistrict = AddressDistrict & "440400,440403,������,22.209117,113.297739;"
    AddressDistrict = AddressDistrict & "440400,440404,������,22.139122,113.345071;"
    AddressDistrict = AddressDistrict & "440500,440507,������,23.373754,116.732015;"
    AddressDistrict = AddressDistrict & "440500,440511,��ƽ��,23.367071,116.703583;"
    AddressDistrict = AddressDistrict & "440500,440512,婽���,23.279345,116.729528;"
    AddressDistrict = AddressDistrict & "440500,440513,������,23.262336,116.602602;"
    AddressDistrict = AddressDistrict & "440500,440514,������,23.249798,116.423607;"
    AddressDistrict = AddressDistrict & "440500,440515,�κ���,23.46844,116.76336;"
    AddressDistrict = AddressDistrict & "440500,440523,�ϰ���,23.419562,117.027105;"
    AddressDistrict = AddressDistrict & "440600,440604,������,23.019643,113.112414;"
    AddressDistrict = AddressDistrict & "440600,440605,�Ϻ���,23.031562,113.145577;"
    AddressDistrict = AddressDistrict & "440600,440606,˳����,22.75851,113.281826;"
    AddressDistrict = AddressDistrict & "440600,440607,��ˮ��,23.16504,112.899414;"
    AddressDistrict = AddressDistrict & "440600,440608,������,22.893855,112.882123;"
    AddressDistrict = AddressDistrict & "440700,440703,���,22.59677,113.07859;"
    AddressDistrict = AddressDistrict & "440700,440704,������,22.572211,113.120601;"
    AddressDistrict = AddressDistrict & "440700,440705,�»���,22.520247,113.038584;"
    AddressDistrict = AddressDistrict & "440700,440781,̨ɽ��,22.250713,112.793414;"
    AddressDistrict = AddressDistrict & "440700,440783,��ƽ��,22.366286,112.692262;"
    AddressDistrict = AddressDistrict & "440700,440784,��ɽ��,22.768104,112.961795;"
    AddressDistrict = AddressDistrict & "440700,440785,��ƽ��,22.182956,112.314051;"
    AddressDistrict = AddressDistrict & "440800,440802,�࿲��,21.273365,110.361634;"
    AddressDistrict = AddressDistrict & "440800,440803,ϼɽ��,21.194229,110.406382;"
    AddressDistrict = AddressDistrict & "440800,440804,��ͷ��,21.24441,110.455632;"
    AddressDistrict = AddressDistrict & "440800,440811,������,21.265997,110.329167;"
    AddressDistrict = AddressDistrict & "440800,440823,��Ϫ��,21.376915,110.255321;"
    AddressDistrict = AddressDistrict & "440800,440825,������,20.326083,110.175718;"
    AddressDistrict = AddressDistrict & "440800,440881,������,21.611281,110.284961;"
    AddressDistrict = AddressDistrict & "440800,440882,������,20.908523,110.088275;"
    AddressDistrict = AddressDistrict & "440800,440883,�⴨��,21.428453,110.780508;"
    AddressDistrict = AddressDistrict & "440900,440902,ï����,21.660425,110.920542;"
    AddressDistrict = AddressDistrict & "440900,440904,�����,21.507219,111.007264;"
    AddressDistrict = AddressDistrict & "440900,440981,������,21.915153,110.853251;"
    AddressDistrict = AddressDistrict & "440900,440982,������,21.654953,110.63839;"
    AddressDistrict = AddressDistrict & "440900,440983,������,22.352681,110.941656;"
    AddressDistrict = AddressDistrict & "441200,441202,������,23.052662,112.472329;"
    AddressDistrict = AddressDistrict & "441200,441203,������,23.155822,112.565249;"
    AddressDistrict = AddressDistrict & "441200,441204,��Ҫ��,23.027694,112.460846;"
    AddressDistrict = AddressDistrict & "441200,441223,������,23.631486,112.440419;"
    AddressDistrict = AddressDistrict & "441200,441224,������,23.913072,112.182466;"
    AddressDistrict = AddressDistrict & "441200,441225,�⿪��,23.434731,111.502973;"
    AddressDistrict = AddressDistrict & "441200,441226,������,23.141711,111.78156;"
    AddressDistrict = AddressDistrict & "441200,441284,�Ļ���,23.340324,112.695028;"
    AddressDistrict = AddressDistrict & "441300,441302,�ݳ���,23.079883,114.413978;"
    AddressDistrict = AddressDistrict & "441300,441303,������,22.78851,114.469444;"
    AddressDistrict = AddressDistrict & "441300,441322,������,23.167575,114.284254;"
    AddressDistrict = AddressDistrict & "441300,441323,�ݶ���,22.983036,114.723092;"
    AddressDistrict = AddressDistrict & "441300,441324,������,23.723894,114.259986;"
    AddressDistrict = AddressDistrict & "441400,441402,÷����,24.302593,116.12116;"
    AddressDistrict = AddressDistrict & "441400,441403,÷����,24.267825,116.083482;"
    AddressDistrict = AddressDistrict & "441400,441422,������,24.351587,116.69552;"
    AddressDistrict = AddressDistrict & "441400,441423,��˳��,23.752771,116.184419;"
    AddressDistrict = AddressDistrict & "441400,441424,�廪��,23.925424,115.775004;"
    AddressDistrict = AddressDistrict & "441400,441426,ƽԶ��,24.569651,115.891729;"
    AddressDistrict = AddressDistrict & "441400,441427,������,24.653313,116.170531;"
    AddressDistrict = AddressDistrict & "441400,441481,������,24.138077,115.731648;"
    AddressDistrict = AddressDistrict & "441500,441502,����,22.776227,115.363667;"
    AddressDistrict = AddressDistrict & "441500,441521,������,22.971042,115.337324;"
    AddressDistrict = AddressDistrict & "441500,441523,½����,23.302682,115.657565;"
    AddressDistrict = AddressDistrict & "441500,441581,½����,22.946104,115.644203;"
    AddressDistrict = AddressDistrict & "441600,441602,Դ����,23.746255,114.696828;"
    AddressDistrict = AddressDistrict & "441600,441621,�Ͻ���,23.633744,115.184383;"
    AddressDistrict = AddressDistrict & "441600,441622,������,24.101174,115.256415;"
    AddressDistrict = AddressDistrict & "441600,441623,��ƽ��,24.364227,114.495952;"
    AddressDistrict = AddressDistrict & "441600,441624,��ƽ��,24.44318,114.941473;"
    AddressDistrict = AddressDistrict & "441600,441625,��Դ��,23.789093,114.742711;"
    AddressDistrict = AddressDistrict & "441700,441702,������,21.859182,111.968909;"
    AddressDistrict = AddressDistrict & "441700,441704,������,21.864728,112.011267;"
    AddressDistrict = AddressDistrict & "441700,441721,������,21.75367,111.617556;"
    AddressDistrict = AddressDistrict & "441700,441781,������,22.169598,111.7905;"
    AddressDistrict = AddressDistrict & "441800,441802,�����,23.688976,113.048698;"
    AddressDistrict = AddressDistrict & "441800,441803,������,23.736949,113.015203;"
    AddressDistrict = AddressDistrict & "441800,441821,�����,23.866739,113.534094;"
    AddressDistrict = AddressDistrict & "441800,441823,��ɽ��,24.470286,112.634019;"
    AddressDistrict = AddressDistrict & "441800,441825,��ɽ׳������������,24.567271,112.086555;"
    AddressDistrict = AddressDistrict & "441800,441826,��������������,24.719097,112.290808;"
    AddressDistrict = AddressDistrict & "441800,441881,Ӣ����,24.18612,113.405404;"
    AddressDistrict = AddressDistrict & "441800,441882,������,24.783966,112.379271;"
    AddressDistrict = AddressDistrict & "445100,445102,������,23.664675,116.63365;"
    AddressDistrict = AddressDistrict & "445100,445103,������,23.461012,116.67931;"
    AddressDistrict = AddressDistrict & "445100,445122,��ƽ��,23.668171,117.00205;"
    AddressDistrict = AddressDistrict & "445200,445202,�ų���,23.535524,116.357045;"
    AddressDistrict = AddressDistrict & "445200,445203,�Ҷ���,23.569887,116.412947;"
    AddressDistrict = AddressDistrict & "445200,445222,������,23.4273,115.838708;"
    AddressDistrict = AddressDistrict & "445200,445224,������,23.029834,116.295832;"
    AddressDistrict = AddressDistrict & "445200,445281,������,23.29788,116.165082;"
    AddressDistrict = AddressDistrict & "445300,445302,�Ƴ���,22.930827,112.04471;"
    AddressDistrict = AddressDistrict & "445300,445303,�ư���,23.073152,112.005609;"
    AddressDistrict = AddressDistrict & "445300,445321,������,22.703204,112.23083;"
    AddressDistrict = AddressDistrict & "445300,445322,������,23.237709,111.535921;"
    AddressDistrict = AddressDistrict & "445300,445381,�޶���,22.765415,111.578201;"
    AddressDistrict = AddressDistrict & "450100,450102,������,22.819511,108.320189;"
    AddressDistrict = AddressDistrict & "450100,450103,������,22.816614,108.346113;"
    AddressDistrict = AddressDistrict & "450100,450105,������,22.799593,108.310478;"
    AddressDistrict = AddressDistrict & "450100,450107,��������,22.832779,108.306903;"
    AddressDistrict = AddressDistrict & "450100,450108,������,22.75909,108.322102;"
    AddressDistrict = AddressDistrict & "450100,450109,������,22.756598,108.484251;"
    AddressDistrict = AddressDistrict & "450100,450110,������,23.157163,108.280717;"
    AddressDistrict = AddressDistrict & "450100,450123,¡����,23.174763,107.688661;"
    AddressDistrict = AddressDistrict & "450100,450124,��ɽ��,23.711758,108.172903;"
    AddressDistrict = AddressDistrict & "450100,450125,������,23.431769,108.603937;"
    AddressDistrict = AddressDistrict & "450100,450126,������,23.216884,108.816735;"
    AddressDistrict = AddressDistrict & "450100,450127,������,22.68743,109.270987;"
    AddressDistrict = AddressDistrict & "450200,450202,������,24.312324,109.411749;"
    AddressDistrict = AddressDistrict & "450200,450203,�����,24.303848,109.415364;"
    AddressDistrict = AddressDistrict & "450200,450204,������,24.287013,109.395936;"
    AddressDistrict = AddressDistrict & "450200,450205,������,24.359145,109.406577;"
    AddressDistrict = AddressDistrict & "450200,450206,������,24.257512,109.334503;"
    AddressDistrict = AddressDistrict & "450200,450222,������,24.655121,109.245812;"
    AddressDistrict = AddressDistrict & "450200,450223,¹կ��,24.483405,109.740805;"
    AddressDistrict = AddressDistrict & "450200,450224,�ڰ���,25.214703,109.403621;"
    AddressDistrict = AddressDistrict & "450200,450225,��ˮ����������,25.068812,109.252744;"
    AddressDistrict = AddressDistrict & "450200,450226,��������������,25.78553,109.614846;"
    AddressDistrict = AddressDistrict & "450300,450302,�����,25.278544,110.292445;"
    AddressDistrict = AddressDistrict & "450300,450303,������,25.301334,110.300783;"
    AddressDistrict = AddressDistrict & "450300,450304,��ɽ��,25.261986,110.284882;"
    AddressDistrict = AddressDistrict & "450300,450305,������,25.254339,110.317577;"
    AddressDistrict = AddressDistrict & "450300,450311,��ɽ��,25.077646,110.305667;"
    AddressDistrict = AddressDistrict & "450300,450312,�ٹ���,25.246257,110.205487;"
    AddressDistrict = AddressDistrict & "450300,450321,��˷��,24.77534,110.494699;"
    AddressDistrict = AddressDistrict & "450300,450323,�鴨��,25.408541,110.325712;"
    AddressDistrict = AddressDistrict & "450300,450324,ȫ����,25.929897,111.072989;"
    AddressDistrict = AddressDistrict & "450300,450325,�˰���,25.609554,110.670783;"
    AddressDistrict = AddressDistrict & "450300,450326,������,24.986692,109.989208;"
    AddressDistrict = AddressDistrict & "450300,450327,������,25.489098,111.160248;"
    AddressDistrict = AddressDistrict & "450300,450328,��ʤ����������,25.796428,110.009423;"
    AddressDistrict = AddressDistrict & "450300,450329,��Դ��,26.0342,110.642587;"
    AddressDistrict = AddressDistrict & "450300,450330,ƽ����,24.632216,110.642821;"
    AddressDistrict = AddressDistrict & "450300,450332,��������������,24.833612,110.82952;"
    AddressDistrict = AddressDistrict & "450300,450381,������,24.497786,110.400149;"
    AddressDistrict = AddressDistrict & "450400,450403,������,23.471318,111.315817;"
    AddressDistrict = AddressDistrict & "450400,450405,������,23.4777,111.275678;"
    AddressDistrict = AddressDistrict & "450400,450406,������,23.40996,111.246035;"
    AddressDistrict = AddressDistrict & "450400,450421,������,23.845097,111.544008;"
    AddressDistrict = AddressDistrict & "450400,450422,����,23.373963,110.931826;"
    AddressDistrict = AddressDistrict & "450400,450423,��ɽ��,24.199829,110.5226;"
    AddressDistrict = AddressDistrict & "450400,450481,�Ϫ��,22.918406,110.998114;"
    AddressDistrict = AddressDistrict & "450500,450502,������,21.468443,109.107529;"
    AddressDistrict = AddressDistrict & "450500,450503,������,21.444909,109.118707;"
    AddressDistrict = AddressDistrict & "450500,450512,��ɽ����,21.5928,109.450573;"
    AddressDistrict = AddressDistrict & "450500,450521,������,21.663554,109.200695;"
    AddressDistrict = AddressDistrict & "450600,450602,�ۿ���,21.614406,108.346281;"
    AddressDistrict = AddressDistrict & "450600,450603,������,21.764758,108.358426;"
    AddressDistrict = AddressDistrict & "450600,450621,��˼��,22.151423,107.982139;"
    AddressDistrict = AddressDistrict & "450600,450681,������,21.541172,107.97017;"
    AddressDistrict = AddressDistrict & "450700,450702,������,21.966808,108.626629;"
    AddressDistrict = AddressDistrict & "450700,450703,�ձ���,22.132761,108.44911;"
    AddressDistrict = AddressDistrict & "450700,450721,��ɽ��,22.418041,109.293468;"
    AddressDistrict = AddressDistrict & "450700,450722,�ֱ���,22.268335,109.556341;"
    AddressDistrict = AddressDistrict & "450800,450802,�۱���,23.107677,109.59481;"
    AddressDistrict = AddressDistrict & "450800,450803,������,23.067516,109.604665;"
    AddressDistrict = AddressDistrict & "450800,450804,������,23.132815,109.415697;"
    AddressDistrict = AddressDistrict & "450800,450821,ƽ����,23.544546,110.397485;"
    AddressDistrict = AddressDistrict & "450800,450881,��ƽ��,23.382473,110.074668;"
    AddressDistrict = AddressDistrict & "450900,450902,������,22.632132,110.154912;"
    AddressDistrict = AddressDistrict & "450900,450903,������,22.58163,110.054155;"
    AddressDistrict = AddressDistrict & "450900,450921,����,22.856435,110.552467;"
    AddressDistrict = AddressDistrict & "450900,450922,½����,22.321054,110.264842;"
    AddressDistrict = AddressDistrict & "450900,450923,������,22.271285,109.980004;"
    AddressDistrict = AddressDistrict & "450900,450924,��ҵ��,22.74187,109.877768;"
    AddressDistrict = AddressDistrict & "450900,450981,������,22.701648,110.348052;"
    AddressDistrict = AddressDistrict & "451000,451002,�ҽ���,23.897675,106.615727;"
    AddressDistrict = AddressDistrict & "451000,451003,������,23.736079,106.904315;"
    AddressDistrict = AddressDistrict & "451000,451022,�ﶫ��,23.600444,107.12426;"
    AddressDistrict = AddressDistrict & "451000,451024,�±���,23.321464,106.618164;"
    AddressDistrict = AddressDistrict & "451000,451026,������,23.400785,105.833553;"
    AddressDistrict = AddressDistrict & "451000,451027,������,24.345643,106.56487;"
    AddressDistrict = AddressDistrict & "451000,451028,��ҵ��,24.782204,106.559638;"
    AddressDistrict = AddressDistrict & "451000,451029,������,24.290262,106.235047;"
    AddressDistrict = AddressDistrict & "451000,451030,������,24.492041,105.095025;"
    AddressDistrict = AddressDistrict & "451000,451031,¡�ָ���������,24.774318,105.342363;"
    AddressDistrict = AddressDistrict & "451000,451081,������,23.134766,106.417549;"
    AddressDistrict = AddressDistrict & "451000,451082,ƽ����,23.320479,107.580403;"
    AddressDistrict = AddressDistrict & "451100,451102,�˲���,24.412446,111.551991;"
    AddressDistrict = AddressDistrict & "451100,451103,ƽ����,24.417148,111.524014;"
    AddressDistrict = AddressDistrict & "451100,451121,��ƽ��,24.172958,110.810865;"
    AddressDistrict = AddressDistrict & "451100,451122,��ɽ��,24.528566,111.303629;"
    AddressDistrict = AddressDistrict & "451100,451123,��������������,24.81896,111.277228;"
    AddressDistrict = AddressDistrict & "451200,451202,��ǽ���,24.695625,108.062131;"
    AddressDistrict = AddressDistrict & "451200,451203,������,24.492193,108.653965;"
    AddressDistrict = AddressDistrict & "451200,451221,�ϵ���,24.983192,107.546605;"
    AddressDistrict = AddressDistrict & "451200,451222,�����,24.985964,107.174939;"
    AddressDistrict = AddressDistrict & "451200,451223,��ɽ��,24.544561,107.044592;"
    AddressDistrict = AddressDistrict & "451200,451224,������,24.509367,107.373696;"
    AddressDistrict = AddressDistrict & "451200,451225,�޳�������������,24.779327,108.902453;"
    AddressDistrict = AddressDistrict & "451200,451226,����ë����������,24.827628,108.258669;"
    AddressDistrict = AddressDistrict & "451200,451227,��������������,24.139538,107.253126;"
    AddressDistrict = AddressDistrict & "451200,451228,��������������,23.934964,108.102761;"
    AddressDistrict = AddressDistrict & "451200,451229,������������,23.739596,107.9945;"
    AddressDistrict = AddressDistrict & "451300,451302,�˱���,23.732926,109.230541;"
    AddressDistrict = AddressDistrict & "451300,451321,�ó���,24.064779,108.667361;"
    AddressDistrict = AddressDistrict & "451300,451322,������,23.959824,109.684555;"
    AddressDistrict = AddressDistrict & "451300,451323,������,23.604162,109.66287;"
    AddressDistrict = AddressDistrict & "451300,451324,��������������,24.134941,110.188556;"
    AddressDistrict = AddressDistrict & "451300,451381,��ɽ��,23.81311,108.88858;"
    AddressDistrict = AddressDistrict & "451400,451402,������,22.40469,107.354443;"
    AddressDistrict = AddressDistrict & "451400,451421,������,22.635821,107.911533;"
    AddressDistrict = AddressDistrict & "451400,451422,������,22.131353,107.067616;"
    AddressDistrict = AddressDistrict & "451400,451423,������,22.343716,106.857502;"
    AddressDistrict = AddressDistrict & "451400,451424,������,22.833369,107.200803;"
    AddressDistrict = AddressDistrict & "451400,451425,�����,23.082484,107.142441;"
    AddressDistrict = AddressDistrict & "451400,451481,ƾ����,22.108882,106.759038;"
    AddressDistrict = AddressDistrict & "460100,460105,��Ӣ��,20.008145,110.282393;"
    AddressDistrict = AddressDistrict & "460100,460106,������,20.031026,110.330373;"
    AddressDistrict = AddressDistrict & "460100,460107,��ɽ��,20.001051,110.354722;"
    AddressDistrict = AddressDistrict & "460100,460108,������,20.03074,110.356566;"
    AddressDistrict = AddressDistrict & "460200,460202,������,18.407516,109.760778;"
    AddressDistrict = AddressDistrict & "460200,460203,������,18.247436,109.512081;"
    AddressDistrict = AddressDistrict & "460200,460204,������,18.24734,109.506357;"
    AddressDistrict = AddressDistrict & "460200,460205,������,18.352192,109.174306;"
    AddressDistrict = AddressDistrict & "460300,460301,��ɳ��,16.8310066,112.3386402;"
    AddressDistrict = AddressDistrict & "460300,460302,��ɳ��,9.543575,112.891018;"
    AddressDistrict = AddressDistrict & "510100,510104,������,30.657689,104.080989;"
    AddressDistrict = AddressDistrict & "510100,510105,������,30.667648,104.055731;"
    AddressDistrict = AddressDistrict & "510100,510106,��ţ��,30.692058,104.043487;"
    AddressDistrict = AddressDistrict & "510100,510107,�����,30.630862,104.05167;"
    AddressDistrict = AddressDistrict & "510100,510108,�ɻ���,30.660275,104.103077;"
    AddressDistrict = AddressDistrict & "510100,510112,��Ȫ����,30.56065,104.269181;"
    AddressDistrict = AddressDistrict & "510100,510113,��׽���,30.883438,104.25494;"
    AddressDistrict = AddressDistrict & "510100,510114,�¶���,30.824223,104.16022;"
    AddressDistrict = AddressDistrict & "510100,510115,�½���,30.697996,103.836776;"
    AddressDistrict = AddressDistrict & "510100,510116,˫����,30.573243,103.922706;"
    AddressDistrict = AddressDistrict & "510100,510117,ۯ����,30.808752,103.887842;"
    AddressDistrict = AddressDistrict & "510100,510118,�½���,30.414284,103.812449;"
    AddressDistrict = AddressDistrict & "510100,510121,������,30.858417,104.415604;"
    AddressDistrict = AddressDistrict & "510100,510129,������,30.586602,103.522397;"
    AddressDistrict = AddressDistrict & "510100,510131,�ѽ���,30.194359,103.511541;"
    AddressDistrict = AddressDistrict & "510100,510181,��������,30.99114,103.627898;"
    AddressDistrict = AddressDistrict & "510100,510182,������,30.985161,103.941173;"
    AddressDistrict = AddressDistrict & "510100,510183,������,30.413271,103.46143;"
    AddressDistrict = AddressDistrict & "510100,510184,������,30.631478,103.671049;"
    AddressDistrict = AddressDistrict & "510100,510185,������,30.390666,104.550339;"
    AddressDistrict = AddressDistrict & "510300,510302,��������,29.343231,104.778188;"
    AddressDistrict = AddressDistrict & "510300,510303,������,29.345675,104.714372;"
    AddressDistrict = AddressDistrict & "510300,510304,����,29.367136,104.783229;"
    AddressDistrict = AddressDistrict & "510300,510311,��̲��,29.272521,104.876417;"
    AddressDistrict = AddressDistrict & "510300,510321,����,29.454851,104.423932;"
    AddressDistrict = AddressDistrict & "510300,510322,��˳��,29.181282,104.984256;"
    AddressDistrict = AddressDistrict & "510400,510402,����,26.580887,101.715134;"
    AddressDistrict = AddressDistrict & "510400,510403,����,26.596776,101.637969;"
    AddressDistrict = AddressDistrict & "510400,510411,�ʺ���,26.497185,101.737916;"
    AddressDistrict = AddressDistrict & "510400,510421,������,26.887474,102.109877;"
    AddressDistrict = AddressDistrict & "510400,510422,�α���,26.677619,101.851848;"
    AddressDistrict = AddressDistrict & "510500,510502,������,28.882889,105.445131;"
    AddressDistrict = AddressDistrict & "510500,510503,��Ϫ��,28.77631,105.37721;"
    AddressDistrict = AddressDistrict & "510500,510504,����̶��,28.897572,105.435228;"
    AddressDistrict = AddressDistrict & "510500,510521,����,29.151288,105.376335;"
    AddressDistrict = AddressDistrict & "510500,510522,�Ͻ���,28.810325,105.834098;"
    AddressDistrict = AddressDistrict & "510500,510524,������,28.167919,105.437775;"
    AddressDistrict = AddressDistrict & "510500,510525,������,28.03948,105.813359;"
    AddressDistrict = AddressDistrict & "510600,510603,�����,31.130428,104.389648;"
    AddressDistrict = AddressDistrict & "510600,510604,�޽���,31.303281,104.507126;"
    AddressDistrict = AddressDistrict & "510600,510623,�н���,31.03681,104.677831;"
    AddressDistrict = AddressDistrict & "510600,510681,�㺺��,30.97715,104.281903;"
    AddressDistrict = AddressDistrict & "510600,510682,ʲ����,31.126881,104.173653;"
    AddressDistrict = AddressDistrict & "510600,510683,������,31.343084,104.200162;"
    AddressDistrict = AddressDistrict & "510700,510703,������,31.463557,104.740971;"
    AddressDistrict = AddressDistrict & "510700,510704,������,31.484772,104.770006;"
    AddressDistrict = AddressDistrict & "510700,510705,������,31.53894,104.560341;"
    AddressDistrict = AddressDistrict & "510700,510722,��̨��,31.090909,105.090316;"
    AddressDistrict = AddressDistrict & "510700,510723,��ͤ��,31.22318,105.391991;"
    AddressDistrict = AddressDistrict & "510700,510725,������,31.635225,105.16353;"
    AddressDistrict = AddressDistrict & "510700,510726,����Ǽ��������,31.615863,104.468069;"
    AddressDistrict = AddressDistrict & "510700,510727,ƽ����,32.407588,104.530555;"
    AddressDistrict = AddressDistrict & "510700,510781,������,31.776386,104.744431;"
    AddressDistrict = AddressDistrict & "510800,510802,������,32.432276,105.826194;"
    AddressDistrict = AddressDistrict & "510800,510811,�ѻ���,32.322788,105.964121;"
    AddressDistrict = AddressDistrict & "510800,510812,������,32.642632,105.88917;"
    AddressDistrict = AddressDistrict & "510800,510821,������,32.22833,106.290426;"
    AddressDistrict = AddressDistrict & "510800,510822,�ന��,32.585655,105.238847;"
    AddressDistrict = AddressDistrict & "510800,510823,������,32.286517,105.527035;"
    AddressDistrict = AddressDistrict & "510800,510824,��Ϫ��,31.732251,105.939706;"
    AddressDistrict = AddressDistrict & "510900,510903,��ɽ��,30.502647,105.582215;"
    AddressDistrict = AddressDistrict & "510900,510904,������,30.346121,105.459383;"
    AddressDistrict = AddressDistrict & "510900,510921,��Ϫ��,30.774883,105.713699;"
    AddressDistrict = AddressDistrict & "510900,510923,��Ӣ��,30.581571,105.252187;"
    AddressDistrict = AddressDistrict & "510900,510981,�����,30.868752,105.381849;"
    AddressDistrict = AddressDistrict & "511000,511002,������,29.585265,105.065467;"
    AddressDistrict = AddressDistrict & "511000,511011,������,29.600107,105.067203;"
    AddressDistrict = AddressDistrict & "511000,511024,��Զ��,29.52686,104.668327;"
    AddressDistrict = AddressDistrict & "511000,511025,������,29.775295,104.852463;"
    AddressDistrict = AddressDistrict & "511000,511083,¡����,29.338162,105.288074;"
    AddressDistrict = AddressDistrict & "511100,511102,������,29.588327,103.75539;"
    AddressDistrict = AddressDistrict & "511100,511111,ɳ����,29.416536,103.549961;"
    AddressDistrict = AddressDistrict & "511100,511112,��ͨ����,29.406186,103.816837;"
    AddressDistrict = AddressDistrict & "511100,511113,��ں���,29.24602,103.077831;"
    AddressDistrict = AddressDistrict & "511100,511123,��Ϊ��,29.209782,103.944266;"
    AddressDistrict = AddressDistrict & "511100,511124,������,29.651645,104.06885;"
    AddressDistrict = AddressDistrict & "511100,511126,�н���,29.741019,103.578862;"
    AddressDistrict = AddressDistrict & "511100,511129,�崨��,28.956338,103.90211;"
    AddressDistrict = AddressDistrict & "511100,511132,�������������,29.230271,103.262148;"
    AddressDistrict = AddressDistrict & "511100,511133,�������������,28.838933,103.546851;"
    AddressDistrict = AddressDistrict & "511100,511181,��üɽ��,29.597478,103.492488;"
    AddressDistrict = AddressDistrict & "511300,511302,˳����,30.795572,106.084091;"
    AddressDistrict = AddressDistrict & "511300,511303,��ƺ��,30.781809,106.108996;"
    AddressDistrict = AddressDistrict & "511300,511304,������,30.762976,106.067027;"
    AddressDistrict = AddressDistrict & "511300,511321,�ϲ���,31.349407,106.061138;"
    AddressDistrict = AddressDistrict & "511300,511322,Ӫɽ��,31.075907,106.564893;"
    AddressDistrict = AddressDistrict & "511300,511323,���,31.027978,106.413488;"
    AddressDistrict = AddressDistrict & "511300,511324,��¤��,31.271261,106.297083;"
    AddressDistrict = AddressDistrict & "511300,511325,������,30.994616,105.893021;"
    AddressDistrict = AddressDistrict & "511300,511381,������,31.580466,105.975266;"
    AddressDistrict = AddressDistrict & "511400,511402,������,30.048128,103.831553;"
    AddressDistrict = AddressDistrict & "511400,511403,��ɽ��,30.192298,103.8701;"
    AddressDistrict = AddressDistrict & "511400,511421,������,29.996721,104.147646;"
    AddressDistrict = AddressDistrict & "511400,511423,������,29.904867,103.375006;"
    AddressDistrict = AddressDistrict & "511400,511424,������,30.012751,103.518333;"
    AddressDistrict = AddressDistrict & "511400,511425,������,29.831469,103.846131;"
    AddressDistrict = AddressDistrict & "511500,511502,������,28.760179,104.630231;"
    AddressDistrict = AddressDistrict & "511500,511503,��Ϫ��,28.839806,104.981133;"
    AddressDistrict = AddressDistrict & "511500,511504,������,28.695678,104.541489;"
    AddressDistrict = AddressDistrict & "511500,511523,������,28.728102,105.068697;"
    AddressDistrict = AddressDistrict & "511500,511524,������,28.577271,104.921116;"
    AddressDistrict = AddressDistrict & "511500,511525,����,28.435676,104.519187;"
    AddressDistrict = AddressDistrict & "511500,511526,����,28.449041,104.712268;"
    AddressDistrict = AddressDistrict & "511500,511527,������,28.162017,104.507848;"
    AddressDistrict = AddressDistrict & "511500,511528,������,28.302988,105.236549;"
    AddressDistrict = AddressDistrict & "511500,511529,��ɽ��,28.64237,104.162617;"
    AddressDistrict = AddressDistrict & "511600,511602,�㰲��,30.456462,106.632907;"
    AddressDistrict = AddressDistrict & "511600,511603,ǰ����,30.4963,106.893277;"
    AddressDistrict = AddressDistrict & "511600,511621,������,30.533538,106.444451;"
    AddressDistrict = AddressDistrict & "511600,511622,��ʤ��,30.344291,106.292473;"
    AddressDistrict = AddressDistrict & "511600,511623,��ˮ��,30.334323,106.934968;"
    AddressDistrict = AddressDistrict & "511600,511681,������,30.380574,106.777882;"
    AddressDistrict = AddressDistrict & "511700,511702,ͨ����,31.213522,107.501062;"
    AddressDistrict = AddressDistrict & "511700,511703,�ﴨ��,31.199062,107.507926;"
    AddressDistrict = AddressDistrict & "511700,511722,������,31.355025,107.722254;"
    AddressDistrict = AddressDistrict & "511700,511723,������,31.085537,107.864135;"
    AddressDistrict = AddressDistrict & "511700,511724,������,30.736289,107.20742;"
    AddressDistrict = AddressDistrict & "511700,511725,����,30.836348,106.970746;"
    AddressDistrict = AddressDistrict & "511700,511781,��Դ��,32.06777,108.037548;"
    AddressDistrict = AddressDistrict & "511800,511802,�����,29.981831,103.003398;"
    AddressDistrict = AddressDistrict & "511800,511803,��ɽ��,30.084718,103.112214;"
    AddressDistrict = AddressDistrict & "511800,511822,������,29.795529,102.844674;"
    AddressDistrict = AddressDistrict & "511800,511823,��Դ��,29.349915,102.677145;"
    AddressDistrict = AddressDistrict & "511800,511824,ʯ����,29.234063,102.35962;"
    AddressDistrict = AddressDistrict & "511800,511825,��ȫ��,30.059955,102.763462;"
    AddressDistrict = AddressDistrict & "511800,511826,«ɽ��,30.152907,102.924016;"
    AddressDistrict = AddressDistrict & "511800,511827,������,30.369026,102.813377;"
    AddressDistrict = AddressDistrict & "511900,511902,������,31.858366,106.753671;"
    AddressDistrict = AddressDistrict & "511900,511903,������,31.816336,106.486515;"
    AddressDistrict = AddressDistrict & "511900,511921,ͨ����,31.91212,107.247621;"
    AddressDistrict = AddressDistrict & "511900,511922,�Ͻ���,32.353164,106.843418;"
    AddressDistrict = AddressDistrict & "511900,511923,ƽ����,31.562814,107.101937;"
    AddressDistrict = AddressDistrict & "512000,512002,�㽭��,30.121686,104.642338;"
    AddressDistrict = AddressDistrict & "512000,512021,������,30.099206,105.336764;"
    AddressDistrict = AddressDistrict & "512000,512022,������,30.275619,105.031142;"
    AddressDistrict = AddressDistrict & "513200,513201,�������,31.899761,102.221187;"
    AddressDistrict = AddressDistrict & "513200,513221,�봨��,31.47463,103.580675;"
    AddressDistrict = AddressDistrict & "513200,513222,����,31.436764,103.165486;"
    AddressDistrict = AddressDistrict & "513200,513223,ï��,31.680407,103.850684;"
    AddressDistrict = AddressDistrict & "513200,513224,������,32.63838,103.599177;"
    AddressDistrict = AddressDistrict & "513200,513225,��կ����,33.262097,104.236344;"
    AddressDistrict = AddressDistrict & "513200,513226,����,31.476356,102.064647;"
    AddressDistrict = AddressDistrict & "513200,513227,С����,30.999016,102.363193;"
    AddressDistrict = AddressDistrict & "513200,513228,��ˮ��,32.061721,102.990805;"
    AddressDistrict = AddressDistrict & "513200,513230,������,32.264887,100.979136;"
    AddressDistrict = AddressDistrict & "513200,513231,������,32.904223,101.700985;"
    AddressDistrict = AddressDistrict & "513200,513232,��������,33.575934,102.963726;"
    AddressDistrict = AddressDistrict & "513200,513233,��ԭ��,32.793902,102.544906;"
    AddressDistrict = AddressDistrict & "513300,513301,������,30.050738,101.964057;"
    AddressDistrict = AddressDistrict & "513300,513322,����,29.912482,102.233225;"
    AddressDistrict = AddressDistrict & "513300,513323,������,30.877083,101.886125;"
    AddressDistrict = AddressDistrict & "513300,513324,������,29.001975,101.506942;"
    AddressDistrict = AddressDistrict & "513300,513325,�Ž���,30.03225,101.015735;"
    AddressDistrict = AddressDistrict & "513300,513326,������,30.978767,101.123327;"
    AddressDistrict = AddressDistrict & "513300,513327,¯����,31.392674,100.679495;"
    AddressDistrict = AddressDistrict & "513300,513328,������,31.61975,99.991753;"
    AddressDistrict = AddressDistrict & "513300,513329,������,30.93896,100.312094;"
    AddressDistrict = AddressDistrict & "513300,513330,�¸���,31.806729,98.57999;"
    AddressDistrict = AddressDistrict & "513300,513331,������,31.208805,98.824343;"
    AddressDistrict = AddressDistrict & "513300,513332,ʯ����,32.975302,98.100887;"
    AddressDistrict = AddressDistrict & "513300,513333,ɫ����,32.268777,100.331657;"
    AddressDistrict = AddressDistrict & "513300,513334,������,29.991807,100.269862;"
    AddressDistrict = AddressDistrict & "513300,513335,������,30.005723,99.109037;"
    AddressDistrict = AddressDistrict & "513300,513336,�����,28.930855,99.799943;"
    AddressDistrict = AddressDistrict & "513300,513337,������,29.037544,100.296689;"
    AddressDistrict = AddressDistrict & "513300,513338,������,28.71134,99.288036;"
    AddressDistrict = AddressDistrict & "513400,513401,������,27.885786,102.258758;"
    AddressDistrict = AddressDistrict & "513400,513422,ľ�����������,27.926859,101.280184;"
    AddressDistrict = AddressDistrict & "513400,513423,��Դ��,27.423415,101.508909;"
    AddressDistrict = AddressDistrict & "513400,513424,�²���,27.403827,102.178845;"
    AddressDistrict = AddressDistrict & "513400,513425,������,26.658702,102.249548;"
    AddressDistrict = AddressDistrict & "513400,513426,�ᶫ��,26.630713,102.578985;"
    AddressDistrict = AddressDistrict & "513400,513427,������,27.065205,102.757374;"
    AddressDistrict = AddressDistrict & "513400,513428,�ո���,27.376828,102.541082;"
    AddressDistrict = AddressDistrict & "513400,513429,������,27.709062,102.808801;"
    AddressDistrict = AddressDistrict & "513400,513430,������,27.695916,103.248704;"
    AddressDistrict = AddressDistrict & "513400,513431,�Ѿ���,28.010554,102.843991;"
    AddressDistrict = AddressDistrict & "513400,513432,ϲ����,28.305486,102.412342;"
    AddressDistrict = AddressDistrict & "513400,513433,������,28.550844,102.170046;"
    AddressDistrict = AddressDistrict & "513400,513434,Խ����,28.639632,102.508875;"
    AddressDistrict = AddressDistrict & "513400,513435,������,28.977094,102.775924;"
    AddressDistrict = AddressDistrict & "513400,513436,������,28.327946,103.132007;"
    AddressDistrict = AddressDistrict & "513400,513437,�ײ���,28.262946,103.571584;"
    AddressDistrict = AddressDistrict & "520100,520102,������,26.573743,106.715963;"
    AddressDistrict = AddressDistrict & "520100,520103,������,26.58301,106.713397;"
    AddressDistrict = AddressDistrict & "520100,520111,��Ϫ��,26.410464,106.670791;"
    AddressDistrict = AddressDistrict & "520100,520112,�ڵ���,26.630928,106.762123;"
    AddressDistrict = AddressDistrict & "520100,520113,������,26.676849,106.633037;"
    AddressDistrict = AddressDistrict & "520100,520115,��ɽ����,26.646358,106.626323;"
    AddressDistrict = AddressDistrict & "520100,520121,������,27.056793,106.969438;"
    AddressDistrict = AddressDistrict & "520100,520122,Ϣ����,27.092665,106.737693;"
    AddressDistrict = AddressDistrict & "520100,520123,������,26.840672,106.599218;"
    AddressDistrict = AddressDistrict & "520100,520181,������,26.551289,106.470278;"
    AddressDistrict = AddressDistrict & "520200,520201,��ɽ��,26.584805,104.846244;"
    AddressDistrict = AddressDistrict & "520200,520203,��֦����,26.210662,105.474235;"
    AddressDistrict = AddressDistrict & "520200,520221,ˮ����,26.540478,104.95685;"
    AddressDistrict = AddressDistrict & "520200,520281,������,25.706966,104.468367;"
    AddressDistrict = AddressDistrict & "520300,520302,�컨����,27.694395,106.943784;"
    AddressDistrict = AddressDistrict & "520300,520303,�㴨��,27.706626,106.937265;"
    AddressDistrict = AddressDistrict & "520300,520304,������,27.535288,106.831668;"
    AddressDistrict = AddressDistrict & "520300,520322,ͩ����,28.131559,106.826591;"
    AddressDistrict = AddressDistrict & "520300,520323,������,27.951342,107.191024;"
    AddressDistrict = AddressDistrict & "520300,520324,������,28.550337,107.441872;"
    AddressDistrict = AddressDistrict & "520300,520325,��������������������,28.880088,107.605342;"
    AddressDistrict = AddressDistrict & "520300,520326,������������������,28.521567,107.887857;"
    AddressDistrict = AddressDistrict & "520300,520327,�����,27.960858,107.722021;"
    AddressDistrict = AddressDistrict & "520300,520328,��̶��,27.765839,107.485723;"
    AddressDistrict = AddressDistrict & "520300,520329,������,27.221552,107.892566;"
    AddressDistrict = AddressDistrict & "520300,520330,ϰˮ��,28.327826,106.200954;"
    AddressDistrict = AddressDistrict & "520300,520381,��ˮ��,28.587057,105.698116;"
    AddressDistrict = AddressDistrict & "520300,520382,�ʻ���,27.803377,106.412476;"
    AddressDistrict = AddressDistrict & "520400,520402,������,26.248323,105.946169;"
    AddressDistrict = AddressDistrict & "520400,520403,ƽ����,26.40608,106.259942;"
    AddressDistrict = AddressDistrict & "520400,520422,�ն���,26.305794,105.745609;"
    AddressDistrict = AddressDistrict & "520400,520423,��������������������,26.056096,105.768656;"
    AddressDistrict = AddressDistrict & "520400,520424,���벼��������������,25.944248,105.618454;"
    AddressDistrict = AddressDistrict & "520400,520425,�������岼����������,25.751567,106.084515;"
    AddressDistrict = AddressDistrict & "520500,520502,���ǹ���,27.302085,105.284852;"
    AddressDistrict = AddressDistrict & "520500,520521,����,27.143521,105.609254;"
    AddressDistrict = AddressDistrict & "520500,520522,ǭ����,27.024923,106.038299;"
    AddressDistrict = AddressDistrict & "520500,520523,��ɳ��,27.459693,106.222103;"
    AddressDistrict = AddressDistrict & "520500,520524,֯����,26.668497,105.768997;"
    AddressDistrict = AddressDistrict & "520500,520525,��Ӻ��,26.769875,105.375322;"
    AddressDistrict = AddressDistrict & "520500,520526,���������������������,26.859099,104.286523;"
    AddressDistrict = AddressDistrict & "520500,520527,������,27.119243,104.726438;"
    AddressDistrict = AddressDistrict & "520600,520602,�̽���,27.718745,109.192117;"
    AddressDistrict = AddressDistrict & "520600,520603,��ɽ��,27.51903,109.21199;"
    AddressDistrict = AddressDistrict & "520600,520621,������,27.691904,108.848427;"
    AddressDistrict = AddressDistrict & "520600,520622,��������������,27.238024,108.917882;"
    AddressDistrict = AddressDistrict & "520600,520623,ʯ����,27.519386,108.229854;"
    AddressDistrict = AddressDistrict & "520600,520624,˼����,27.941331,108.255827;"
    AddressDistrict = AddressDistrict & "520600,520625,ӡ������������������,27.997976,108.405517;"
    AddressDistrict = AddressDistrict & "520600,520626,�½���,28.26094,108.117317;"
    AddressDistrict = AddressDistrict & "520600,520627,�غ�������������,28.560487,108.495746;"
    AddressDistrict = AddressDistrict & "520600,520628,��������������,28.165419,109.202627;"
    AddressDistrict = AddressDistrict & "522300,522301,������,25.088599,104.897982;"
    AddressDistrict = AddressDistrict & "522300,522302,������,25.431378,105.192778;"
    AddressDistrict = AddressDistrict & "522300,522323,�հ���,25.786404,104.955347;"
    AddressDistrict = AddressDistrict & "522300,522324,��¡��,25.832881,105.218773;"
    AddressDistrict = AddressDistrict & "522300,522325,�����,25.385752,105.650133;"
    AddressDistrict = AddressDistrict & "522300,522326,������,25.166667,106.091563;"
    AddressDistrict = AddressDistrict & "522300,522327,�����,24.983338,105.81241;"
    AddressDistrict = AddressDistrict & "522300,522328,������,25.108959,105.471498;"
    AddressDistrict = AddressDistrict & "522600,522601,������,26.582964,107.977541;"
    AddressDistrict = AddressDistrict & "522600,522622,��ƽ��,26.896973,107.901337;"
    AddressDistrict = AddressDistrict & "522600,522623,ʩ����,27.034657,108.12678;"
    AddressDistrict = AddressDistrict & "522600,522624,������,26.959884,108.681121;"
    AddressDistrict = AddressDistrict & "522600,522625,��Զ��,27.050233,108.423656;"
    AddressDistrict = AddressDistrict & "522600,522626,᯹���,27.173244,108.816459;"
    AddressDistrict = AddressDistrict & "522600,522627,������,26.909684,109.212798;"
    AddressDistrict = AddressDistrict & "522600,522628,������,26.680625,109.20252;"
    AddressDistrict = AddressDistrict & "522600,522629,������,26.727349,108.440499;"
    AddressDistrict = AddressDistrict & "522600,522630,̨����,26.669138,108.314637;"
    AddressDistrict = AddressDistrict & "522600,522631,��ƽ��,26.230636,109.136504;"
    AddressDistrict = AddressDistrict & "522600,522632,�Ž���,25.931085,108.521026;"
    AddressDistrict = AddressDistrict & "522600,522633,�ӽ���,25.747058,108.912648;"
    AddressDistrict = AddressDistrict & "522600,522634,��ɽ��,26.381027,108.079613;"
    AddressDistrict = AddressDistrict & "522600,522635,�齭��,26.494803,107.593172;"
    AddressDistrict = AddressDistrict & "522600,522636,��կ��,26.199497,107.794808;"
    AddressDistrict = AddressDistrict & "522700,522701,������,26.258205,107.517021;"
    AddressDistrict = AddressDistrict & "522700,522702,��Ȫ��,26.702508,107.513508;"
    AddressDistrict = AddressDistrict & "522700,522722,����,25.412239,107.8838;"
    AddressDistrict = AddressDistrict & "522700,522723,����,26.580807,107.233588;"
    AddressDistrict = AddressDistrict & "522700,522725,�Ͱ���,27.066339,107.478417;"
    AddressDistrict = AddressDistrict & "522700,522726,��ɽ��,25.826283,107.542757;"
    AddressDistrict = AddressDistrict & "522700,522727,ƽ����,25.831803,107.32405;"
    AddressDistrict = AddressDistrict & "522700,522728,�޵���,25.429894,106.750006;"
    AddressDistrict = AddressDistrict & "522700,522729,��˳��,26.022116,106.447376;"
    AddressDistrict = AddressDistrict & "522700,522730,������,26.448809,106.977733;"
    AddressDistrict = AddressDistrict & "522700,522731,��ˮ��,26.128637,106.657848;"
    AddressDistrict = AddressDistrict & "522700,522732,����ˮ��������,25.985183,107.87747;"
    AddressDistrict = AddressDistrict & "530100,530102,�廪��,25.042165,102.704412;"
    AddressDistrict = AddressDistrict & "530100,530103,������,25.070239,102.729044;"
    AddressDistrict = AddressDistrict & "530100,530111,�ٶ���,25.021211,102.723437;"
    AddressDistrict = AddressDistrict & "530100,530112,��ɽ��,25.02436,102.705904;"
    AddressDistrict = AddressDistrict & "530100,530113,������,26.08349,103.182;"
    AddressDistrict = AddressDistrict & "530100,530114,�ʹ���,24.889275,102.801382;"
    AddressDistrict = AddressDistrict & "530100,530115,������,24.666944,102.594987;"
    AddressDistrict = AddressDistrict & "530100,530124,������,25.219667,102.497888;"
    AddressDistrict = AddressDistrict & "530100,530125,������,24.918215,103.145989;"
    AddressDistrict = AddressDistrict & "530100,530126,ʯ������������,24.754545,103.271962;"
    AddressDistrict = AddressDistrict & "530100,530127,������,25.335087,103.038777;"
    AddressDistrict = AddressDistrict & "530100,530128,»Ȱ��������������,25.556533,102.46905;"
    AddressDistrict = AddressDistrict & "530100,530129,Ѱ���������������,25.559474,103.257588;"
    AddressDistrict = AddressDistrict & "530100,530181,������,24.921785,102.485544;"
    AddressDistrict = AddressDistrict & "530300,530302,������,25.501269,103.798054;"
    AddressDistrict = AddressDistrict & "530300,530303,մ����,25.600878,103.819262;"
    AddressDistrict = AddressDistrict & "530300,530304,������,25.429451,103.578755;"
    AddressDistrict = AddressDistrict & "530300,530322,½����,25.022878,103.655233;"
    AddressDistrict = AddressDistrict & "530300,530323,ʦ����,24.825681,103.993808;"
    AddressDistrict = AddressDistrict & "530300,530324,��ƽ��,24.885708,104.309263;"
    AddressDistrict = AddressDistrict & "530300,530325,��Դ��,25.67064,104.25692;"
    AddressDistrict = AddressDistrict & "530300,530326,������,26.412861,103.300041;"
    AddressDistrict = AddressDistrict & "530300,530381,������,26.227777,104.09554;"
    AddressDistrict = AddressDistrict & "530400,530402,������,24.350753,102.543468;"
    AddressDistrict = AddressDistrict & "530400,530403,������,24.291006,102.749839;"
    AddressDistrict = AddressDistrict & "530400,530423,ͨ����,24.112205,102.760039;"
    AddressDistrict = AddressDistrict & "530400,530424,������,24.189807,102.928982;"
    AddressDistrict = AddressDistrict & "530400,530425,������,24.669598,102.16211;"
    AddressDistrict = AddressDistrict & "530400,530426,��ɽ����������,24.173256,102.404358;"
    AddressDistrict = AddressDistrict & "530400,530427,��ƽ�������������,24.0664,101.990903;"
    AddressDistrict = AddressDistrict & "530400,530428,Ԫ���������������������,23.597618,101.999658;"
    AddressDistrict = AddressDistrict & "530400,530481,�ν���,24.669679,102.916652;"
    AddressDistrict = AddressDistrict & "530500,530502,¡����,25.112144,99.165825;"
    AddressDistrict = AddressDistrict & "530500,530521,ʩ����,24.730847,99.183758;"
    AddressDistrict = AddressDistrict & "530500,530523,������,24.591912,98.693567;"
    AddressDistrict = AddressDistrict & "530500,530524,������,24.823662,99.612344;"
    AddressDistrict = AddressDistrict & "530500,530581,�ڳ���,25.01757,98.497292;"
    AddressDistrict = AddressDistrict & "530600,530602,������,27.336636,103.717267;"
    AddressDistrict = AddressDistrict & "530600,530621,³����,27.191637,103.549333;"
    AddressDistrict = AddressDistrict & "530600,530622,�ɼ���,26.9117,102.929284;"
    AddressDistrict = AddressDistrict & "530600,530623,�ν���,28.106923,104.23506;"
    AddressDistrict = AddressDistrict & "530600,530624,�����,27.747114,103.891608;"
    AddressDistrict = AddressDistrict & "530600,530625,������,28.231526,103.63732;"
    AddressDistrict = AddressDistrict & "530600,530626,�罭��,28.599953,103.961095;"
    AddressDistrict = AddressDistrict & "530600,530627,������,27.436267,104.873055;"
    AddressDistrict = AddressDistrict & "530600,530628,������,27.627425,104.048492;"
    AddressDistrict = AddressDistrict & "530600,530629,������,27.843381,105.04869;"
    AddressDistrict = AddressDistrict & "530600,530681,ˮ����,28.629688,104.415376;"
    AddressDistrict = AddressDistrict & "530700,530702,�ų���,26.872229,100.234412;"
    AddressDistrict = AddressDistrict & "530700,530721,����������������,26.830593,100.238312;"
    AddressDistrict = AddressDistrict & "530700,530722,��ʤ��,26.685623,100.750901;"
    AddressDistrict = AddressDistrict & "530700,530723,��ƺ��,26.628834,101.267796;"
    AddressDistrict = AddressDistrict & "530700,530724,��������������,27.281109,100.852427;"
    AddressDistrict = AddressDistrict & "530800,530802,˼é��,22.776595,100.973227;"
    AddressDistrict = AddressDistrict & "530800,530821,��������������������,23.062507,101.04524;"
    AddressDistrict = AddressDistrict & "530800,530822,ī��������������,23.428165,101.687606;"
    AddressDistrict = AddressDistrict & "530800,530823,��������������,24.448523,100.840011;"
    AddressDistrict = AddressDistrict & "530800,530824,���ȴ�������������,23.500278,100.701425;"
    AddressDistrict = AddressDistrict & "530800,530825,�������������������������,24.005712,101.108512;"
    AddressDistrict = AddressDistrict & "530800,530826,���ǹ���������������,22.58336,101.859144;"
    AddressDistrict = AddressDistrict & "530800,530827,������������������������,22.325924,99.585406;"
    AddressDistrict = AddressDistrict & "530800,530828,����������������,22.553083,99.931201;"
    AddressDistrict = AddressDistrict & "530800,530829,��������������,22.644423,99.594372;"
    AddressDistrict = AddressDistrict & "530900,530902,������,23.886562,100.086486;"
    AddressDistrict = AddressDistrict & "530900,530921,������,24.592738,99.91871;"
    AddressDistrict = AddressDistrict & "530900,530922,����,24.439026,100.125637;"
    AddressDistrict = AddressDistrict & "530900,530923,������,24.028159,99.253679;"
    AddressDistrict = AddressDistrict & "530900,530924,����,23.761415,98.82743;"
    AddressDistrict = AddressDistrict & "530900,530925,˫�����������岼�������������,23.477476,99.824419;"
    AddressDistrict = AddressDistrict & "530900,530926,�����������������,23.534579,99.402495;"
    AddressDistrict = AddressDistrict & "530900,530927,��Դ����������,23.146887,99.2474;"
    AddressDistrict = AddressDistrict & "532300,532301,������,25.040912,101.546145;"
    AddressDistrict = AddressDistrict & "532300,532322,˫����,24.685094,101.63824;"
    AddressDistrict = AddressDistrict & "532300,532323,Ĳ����,25.312111,101.543044;"
    AddressDistrict = AddressDistrict & "532300,532324,�ϻ���,25.192408,101.274991;"
    AddressDistrict = AddressDistrict & "532300,532325,Ҧ����,25.505403,101.238399;"
    AddressDistrict = AddressDistrict & "532300,532326,��Ҧ��,25.722348,101.323602;"
    AddressDistrict = AddressDistrict & "532300,532327,������,26.056316,101.671175;"
    AddressDistrict = AddressDistrict & "532300,532328,Ԫı��,25.703313,101.870837;"
    AddressDistrict = AddressDistrict & "532300,532329,�䶨��,25.5301,102.406785;"
    AddressDistrict = AddressDistrict & "532300,532331,»����,25.14327,102.075694;"
    AddressDistrict = AddressDistrict & "532500,532501,������,23.360383,103.154752;"
    AddressDistrict = AddressDistrict & "532500,532502,��Զ��,23.713832,103.258679;"
    AddressDistrict = AddressDistrict & "532500,532503,������,23.366843,103.385005;"
    AddressDistrict = AddressDistrict & "532500,532504,������,24.40837,103.436988;"
    AddressDistrict = AddressDistrict & "532500,532523,��������������,22.987013,103.687229;"
    AddressDistrict = AddressDistrict & "532500,532524,��ˮ��,23.618387,102.820493;"
    AddressDistrict = AddressDistrict & "532500,532525,ʯ����,23.712569,102.484469;"
    AddressDistrict = AddressDistrict & "532500,532527,������,24.532368,103.759622;"
    AddressDistrict = AddressDistrict & "532500,532528,Ԫ����,23.219773,102.837056;"
    AddressDistrict = AddressDistrict & "532500,532529,�����,23.369191,102.42121;"
    AddressDistrict = AddressDistrict & "532500,532530,��ƽ�����������������,22.779982,103.228359;"
    AddressDistrict = AddressDistrict & "532500,532531,�̴���,22.99352,102.39286;"
    AddressDistrict = AddressDistrict & "532500,532532,�ӿ�����������,22.507563,103.961593;"
    AddressDistrict = AddressDistrict & "532600,532601,��ɽ��,23.369216,104.244277;"
    AddressDistrict = AddressDistrict & "532600,532622,��ɽ��,23.612301,104.343989;"
    AddressDistrict = AddressDistrict & "532600,532623,������,23.437439,104.675711;"
    AddressDistrict = AddressDistrict & "532600,532624,��������,23.124202,104.701899;"
    AddressDistrict = AddressDistrict & "532600,532625,�����,23.011723,104.398619;"
    AddressDistrict = AddressDistrict & "532600,532626,����,24.040982,104.194366;"
    AddressDistrict = AddressDistrict & "532600,532627,������,24.050272,105.056684;"
    AddressDistrict = AddressDistrict & "532600,532628,������,23.626494,105.62856;"
    AddressDistrict = AddressDistrict & "532800,532801,������,22.002087,100.797947;"
    AddressDistrict = AddressDistrict & "532800,532822,�º���,21.955866,100.448288;"
    AddressDistrict = AddressDistrict & "532800,532823,������,21.479449,101.567051;"
    AddressDistrict = AddressDistrict & "532900,532901,������,25.593067,100.241369;"
    AddressDistrict = AddressDistrict & "532900,532922,�������������,25.669543,99.95797;"
    AddressDistrict = AddressDistrict & "532900,532923,������,25.477072,100.554025;"
    AddressDistrict = AddressDistrict & "532900,532924,������,25.825904,100.578957;"
    AddressDistrict = AddressDistrict & "532900,532925,�ֶ���,25.342594,100.490669;"
    AddressDistrict = AddressDistrict & "532900,532926,�Ͻ�����������,25.041279,100.518683;"
    AddressDistrict = AddressDistrict & "532900,532927,Ρɽ�������������,25.230909,100.30793;"
    AddressDistrict = AddressDistrict & "532900,532928,��ƽ��,25.461281,99.533536;"
    AddressDistrict = AddressDistrict & "532900,532929,������,25.884955,99.369402;"
    AddressDistrict = AddressDistrict & "532900,532930,��Դ��,26.111184,99.951708;"
    AddressDistrict = AddressDistrict & "532900,532931,������,26.530066,99.905887;"
    AddressDistrict = AddressDistrict & "532900,532932,������,26.55839,100.173375;"
    AddressDistrict = AddressDistrict & "533100,533102,������,24.010734,97.855883;"
    AddressDistrict = AddressDistrict & "533100,533103,â��,24.436699,98.577608;"
    AddressDistrict = AddressDistrict & "533100,533122,������,24.80742,98.298196;"
    AddressDistrict = AddressDistrict & "533100,533123,ӯ����,24.709541,97.93393;"
    AddressDistrict = AddressDistrict & "533100,533124,¤����,24.184065,97.794441;"
    AddressDistrict = AddressDistrict & "533300,533301,��ˮ��,25.851142,98.854063;"
    AddressDistrict = AddressDistrict & "533300,533323,������,26.902738,98.867413;"
    AddressDistrict = AddressDistrict & "533300,533324,��ɽ������ŭ��������,27.738054,98.666141;"
    AddressDistrict = AddressDistrict & "533300,533325,��ƺ����������������,26.453839,99.421378;"
    AddressDistrict = AddressDistrict & "533400,533401,���������,27.825804,99.708667;"
    AddressDistrict = AddressDistrict & "533400,533422,������,28.483272,98.91506;"
    AddressDistrict = AddressDistrict & "533400,533423,ά��������������,27.180948,99.286355;"
    AddressDistrict = AddressDistrict & "540100,540102,�ǹ���,29.659472,91.132911;"
    AddressDistrict = AddressDistrict & "540100,540103,����������,29.647347,91.002823;"
    AddressDistrict = AddressDistrict & "540100,540104,������,29.670314,91.350976;"
    AddressDistrict = AddressDistrict & "540100,540121,������,29.895754,91.261842;"
    AddressDistrict = AddressDistrict & "540100,540122,������,30.474819,91.103551;"
    AddressDistrict = AddressDistrict & "540100,540123,��ľ��,29.431346,90.165545;"
    AddressDistrict = AddressDistrict & "540100,540124,��ˮ��,29.349895,90.738051;"
    AddressDistrict = AddressDistrict & "540100,540127,ī�񹤿���,29.834657,91.731158;"
    AddressDistrict = AddressDistrict & "540200,540202,ɣ������,29.267003,88.88667;"
    AddressDistrict = AddressDistrict & "540200,540221,��ľ����,29.680459,89.099434;"
    AddressDistrict = AddressDistrict & "540200,540222,������,28.908845,89.605044;"
    AddressDistrict = AddressDistrict & "540200,540223,������,28.656667,87.123887;"
    AddressDistrict = AddressDistrict & "540200,540224,������,28.901077,88.023007;"
    AddressDistrict = AddressDistrict & "540200,540225,������,29.085136,87.63743;"
    AddressDistrict = AddressDistrict & "540200,540226,������,29.294758,87.23578;"
    AddressDistrict = AddressDistrict & "540200,540227,лͨ����,29.431597,88.260517;"
    AddressDistrict = AddressDistrict & "540200,540228,������,29.106627,89.263618;"
    AddressDistrict = AddressDistrict & "540200,540229,�ʲ���,29.230299,89.843207;"
    AddressDistrict = AddressDistrict & "540200,540230,������,28.554719,89.683406;"
    AddressDistrict = AddressDistrict & "540200,540231,������,28.36409,87.767723;"
    AddressDistrict = AddressDistrict & "540200,540232,�ٰ���,29.768336,84.032826;"
    AddressDistrict = AddressDistrict & "540200,540233,�Ƕ���,27.482772,88.906806;"
    AddressDistrict = AddressDistrict & "540200,540234,��¡��,28.852416,85.298349;"
    AddressDistrict = AddressDistrict & "540200,540235,����ľ��,28.15595,85.981953;"
    AddressDistrict = AddressDistrict & "540200,540236,������,29.328194,85.234622;"
    AddressDistrict = AddressDistrict & "540200,540237,�ڰ���,28.274371,88.518903;"
    AddressDistrict = AddressDistrict & "540300,540302,������,31.137035,97.178255;"
    AddressDistrict = AddressDistrict & "540300,540321,������,31.499534,98.218351;"
    AddressDistrict = AddressDistrict & "540300,540322,������,30.859206,98.271191;"
    AddressDistrict = AddressDistrict & "540300,540323,��������,31.213048,96.601259;"
    AddressDistrict = AddressDistrict & "540300,540324,������,31.410681,95.597748;"
    AddressDistrict = AddressDistrict & "540300,540325,������,30.653038,97.565701;"
    AddressDistrict = AddressDistrict & "540300,540326,������,30.053408,96.917893;"
    AddressDistrict = AddressDistrict & "540300,540327,����,29.671335,97.840532;"
    AddressDistrict = AddressDistrict & "540300,540328,â����,29.686615,98.596444;"
    AddressDistrict = AddressDistrict & "540300,540329,��¡��,30.741947,95.823418;"
    AddressDistrict = AddressDistrict & "540300,540330,�߰���,30.933849,94.707504;"
    AddressDistrict = AddressDistrict & "540400,540402,������,29.653732,94.360987;"
    AddressDistrict = AddressDistrict & "540400,540421,����������,29.88447,93.246515;"
    AddressDistrict = AddressDistrict & "540400,540422,������,29.213811,94.213679;"
    AddressDistrict = AddressDistrict & "540400,540423,ī����,29.32573,95.332245;"
    AddressDistrict = AddressDistrict & "540400,540424,������,29.858771,95.768151;"
    AddressDistrict = AddressDistrict & "540400,540425,������,28.660244,97.465002;"
    AddressDistrict = AddressDistrict & "540400,540426,����,29.0446,93.073429;"
    AddressDistrict = AddressDistrict & "540500,540502,�˶���,29.236106,91.76525;"
    AddressDistrict = AddressDistrict & "540500,540521,������,29.246476,91.338;"
    AddressDistrict = AddressDistrict & "540500,540522,������,29.289078,90.985271;"
    AddressDistrict = AddressDistrict & "540500,540523,ɣ����,29.259774,92.015732;"
    AddressDistrict = AddressDistrict & "540500,540524,�����,29.025242,91.683753;"
    AddressDistrict = AddressDistrict & "540500,540525,������,29.063656,92.201066;"
    AddressDistrict = AddressDistrict & "540500,540526,������,28.437353,91.432347;"
    AddressDistrict = AddressDistrict & "540500,540527,������,28.385765,90.858243;"
    AddressDistrict = AddressDistrict & "540500,540528,�Ӳ���,29.140921,92.591043;"
    AddressDistrict = AddressDistrict & "540500,540529,¡����,28.408548,92.463309;"
    AddressDistrict = AddressDistrict & "540500,540530,������,27.991707,91.960132;"
    AddressDistrict = AddressDistrict & "540500,540531,�˿�����,28.96836,90.398747;"
    AddressDistrict = AddressDistrict & "540600,540602,ɫ����,31.475756,92.061862;"
    AddressDistrict = AddressDistrict & "540600,540621,������,30.640846,93.232907;"
    AddressDistrict = AddressDistrict & "540600,540622,������,31.479917,93.68044;"
    AddressDistrict = AddressDistrict & "540600,540623,������,32.107855,92.303659;"
    AddressDistrict = AddressDistrict & "540600,540624,������,32.260299,91.681879;"
    AddressDistrict = AddressDistrict & "540600,540625,������,30.929056,88.709777;"
    AddressDistrict = AddressDistrict & "540600,540626,����,31.886173,93.784964;"
    AddressDistrict = AddressDistrict & "540600,540627,�����,31.394578,90.011822;"
    AddressDistrict = AddressDistrict & "540600,540628,������,31.918691,94.054049;"
    AddressDistrict = AddressDistrict & "540600,540629,������,31.784979,87.236646;"
    AddressDistrict = AddressDistrict & "540600,540630,˫����,33.18698,88.838578;"
    AddressDistrict = AddressDistrict & "542500,542521,������,30.291896,81.177588;"
    AddressDistrict = AddressDistrict & "542500,542522,������,31.478587,79.803191;"
    AddressDistrict = AddressDistrict & "542500,542523,������,32.503373,80.105005;"
    AddressDistrict = AddressDistrict & "542500,542524,������,33.382454,79.731937;"
    AddressDistrict = AddressDistrict & "542500,542525,�Ｊ��,32.389192,81.142896;"
    AddressDistrict = AddressDistrict & "542500,542526,������,32.302076,84.062384;"
    AddressDistrict = AddressDistrict & "542500,542527,������,31.016774,85.159254;"
    AddressDistrict = AddressDistrict & "610100,610102,�³���,34.26927,108.959903;"
    AddressDistrict = AddressDistrict & "610100,610103,������,34.251061,108.946994;"
    AddressDistrict = AddressDistrict & "610100,610104,������,34.2656,108.933194;"
    AddressDistrict = AddressDistrict & "610100,610111,�����,34.267453,109.067261;"
    AddressDistrict = AddressDistrict & "610100,610112,δ����,34.30823,108.946022;"
    AddressDistrict = AddressDistrict & "610100,610113,������,34.213389,108.926593;"
    AddressDistrict = AddressDistrict & "610100,610114,������,34.662141,109.22802;"
    AddressDistrict = AddressDistrict & "610100,610115,������,34.372065,109.213986;"
    AddressDistrict = AddressDistrict & "610100,610116,������,34.157097,108.941579;"
    AddressDistrict = AddressDistrict & "610100,610117,������,34.535065,109.088896;"
    AddressDistrict = AddressDistrict & "610100,610118,������,34.108668,108.607385;"
    AddressDistrict = AddressDistrict & "610100,610122,������,34.156189,109.317634;"
    AddressDistrict = AddressDistrict & "610100,610124,������,34.161532,108.216465;"
    AddressDistrict = AddressDistrict & "610200,610202,������,35.069098,109.075862;"
    AddressDistrict = AddressDistrict & "610200,610203,ӡ̨��,35.111927,109.100814;"
    AddressDistrict = AddressDistrict & "610200,610204,ҫ����,34.910206,108.962538;"
    AddressDistrict = AddressDistrict & "610200,610222,�˾���,35.398766,109.118278;"
    AddressDistrict = AddressDistrict & "610300,610302,μ����,34.371008,107.144467;"
    AddressDistrict = AddressDistrict & "610300,610303,��̨��,34.375192,107.149943;"
    AddressDistrict = AddressDistrict & "610300,610304,�²���,34.352747,107.383645;"
    AddressDistrict = AddressDistrict & "610300,610322,������,34.521668,107.400577;"
    AddressDistrict = AddressDistrict & "610300,610323,�ɽ��,34.44296,107.624464;"
    AddressDistrict = AddressDistrict & "610300,610324,������,34.375497,107.891419;"
    AddressDistrict = AddressDistrict & "610300,610326,ü��,34.272137,107.752371;"
    AddressDistrict = AddressDistrict & "610300,610327,¤��,34.893262,106.857066;"
    AddressDistrict = AddressDistrict & "610300,610328,ǧ����,34.642584,107.132987;"
    AddressDistrict = AddressDistrict & "610300,610329,������,34.677714,107.796608;"
    AddressDistrict = AddressDistrict & "610300,610330,����,33.912464,106.525212;"
    AddressDistrict = AddressDistrict & "610300,610331,̫����,34.059215,107.316533;"
    AddressDistrict = AddressDistrict & "610400,610402,�ض���,34.329801,108.698636;"
    AddressDistrict = AddressDistrict & "610400,610403,������,34.27135,108.086348;"
    AddressDistrict = AddressDistrict & "610400,610404,μ����,34.336847,108.730957;"
    AddressDistrict = AddressDistrict & "610400,610422,��ԭ��,34.613996,108.943481;"
    AddressDistrict = AddressDistrict & "610400,610423,������,34.528493,108.83784;"
    AddressDistrict = AddressDistrict & "610400,610424,Ǭ��,34.527261,108.247406;"
    AddressDistrict = AddressDistrict & "610400,610425,��Ȫ��,34.482583,108.428317;"
    AddressDistrict = AddressDistrict & "610400,610426,������,34.692619,108.143129;"
    AddressDistrict = AddressDistrict & "610400,610428,������,35.206122,107.795835;"
    AddressDistrict = AddressDistrict & "610400,610429,Ѯ����,35.112234,108.337231;"
    AddressDistrict = AddressDistrict & "610400,610430,������,34.79797,108.581173;"
    AddressDistrict = AddressDistrict & "610400,610431,�书��,34.259732,108.212857;"
    AddressDistrict = AddressDistrict & "610400,610481,��ƽ��,34.297134,108.488493;"
    AddressDistrict = AddressDistrict & "610400,610482,������,35.034233,108.083674;"
    AddressDistrict = AddressDistrict & "610500,610502,��μ��,34.501271,109.503299;"
    AddressDistrict = AddressDistrict & "610500,610503,������,34.511958,109.76141;"
    AddressDistrict = AddressDistrict & "610500,610522,������,34.544515,110.24726;"
    AddressDistrict = AddressDistrict & "610500,610523,������,34.795011,109.943123;"
    AddressDistrict = AddressDistrict & "610500,610524,������,35.237098,110.147979;"
    AddressDistrict = AddressDistrict & "610500,610525,�γ���,35.184,109.937609;"
    AddressDistrict = AddressDistrict & "610500,610526,�ѳ���,34.956034,109.589653;"
    AddressDistrict = AddressDistrict & "610500,610527,��ˮ��,35.177291,109.594309;"
    AddressDistrict = AddressDistrict & "610500,610528,��ƽ��,34.746679,109.187174;"
    AddressDistrict = AddressDistrict & "610500,610581,������,35.475238,110.452391;"
    AddressDistrict = AddressDistrict & "610500,610582,������,34.565359,110.08952;"
    AddressDistrict = AddressDistrict & "610600,610602,������,36.596291,109.49069;"
    AddressDistrict = AddressDistrict & "610600,610603,������,36.86441,109.325341;"
    AddressDistrict = AddressDistrict & "610600,610621,�ӳ���,36.578306,110.012961;"
    AddressDistrict = AddressDistrict & "610600,610622,�Ӵ���,36.882066,110.190314;"
    AddressDistrict = AddressDistrict & "610600,610625,־����,36.823031,108.768898;"
    AddressDistrict = AddressDistrict & "610600,610626,������,36.924852,108.176976;"
    AddressDistrict = AddressDistrict & "610600,610627,��Ȫ��,36.277729,109.34961;"
    AddressDistrict = AddressDistrict & "610600,610628,����,35.996495,109.384136;"
    AddressDistrict = AddressDistrict & "610600,610629,�崨��,35.762133,109.435712;"
    AddressDistrict = AddressDistrict & "610600,610630,�˴���,36.050391,110.175537;"
    AddressDistrict = AddressDistrict & "610600,610631,������,35.583276,109.83502;"
    AddressDistrict = AddressDistrict & "610600,610632,������,35.580165,109.262469;"
    AddressDistrict = AddressDistrict & "610600,610681,�ӳ���,37.14207,109.675968;"
    AddressDistrict = AddressDistrict & "610700,610702,��̨��,33.077674,107.028233;"
    AddressDistrict = AddressDistrict & "610700,610703,��֣��,33.003341,106.942393;"
    AddressDistrict = AddressDistrict & "610700,610722,�ǹ���,33.153098,107.329887;"
    AddressDistrict = AddressDistrict & "610700,610723,����,33.223283,107.549962;"
    AddressDistrict = AddressDistrict & "610700,610724,������,32.987961,107.765858;"
    AddressDistrict = AddressDistrict & "610700,610725,����,33.155618,106.680175;"
    AddressDistrict = AddressDistrict & "610700,610726,��ǿ��,32.830806,106.25739;"
    AddressDistrict = AddressDistrict & "610700,610727,������,33.329638,106.153899;"
    AddressDistrict = AddressDistrict & "610700,610728,�����,32.535854,107.89531;"
    AddressDistrict = AddressDistrict & "610700,610729,������,33.61334,106.924377;"
    AddressDistrict = AddressDistrict & "610700,610730,��ƺ��,33.520745,107.988582;"
    AddressDistrict = AddressDistrict & "610800,610802,������,38.299267,109.74791;"
    AddressDistrict = AddressDistrict & "610800,610803,��ɽ��,37.964048,109.292596;"
    AddressDistrict = AddressDistrict & "610800,610822,������,39.029243,111.069645;"
    AddressDistrict = AddressDistrict & "610800,610824,������,37.596084,108.80567;"
    AddressDistrict = AddressDistrict & "610800,610825,������,37.59523,107.601284;"
    AddressDistrict = AddressDistrict & "610800,610826,�����,37.507701,110.265377;"
    AddressDistrict = AddressDistrict & "610800,610827,��֬��,37.759081,110.178683;"
    AddressDistrict = AddressDistrict & "610800,610828,����,38.021597,110.493367;"
    AddressDistrict = AddressDistrict & "610800,610829,�Ɽ��,37.451925,110.739315;"
    AddressDistrict = AddressDistrict & "610800,610830,�彧��,37.087702,110.12146;"
    AddressDistrict = AddressDistrict & "610800,610831,������,37.611573,110.03457;"
    AddressDistrict = AddressDistrict & "610800,610881,��ľ��,38.835641,110.497005;"
    AddressDistrict = AddressDistrict & "610900,610902,������,32.690817,109.029098;"
    AddressDistrict = AddressDistrict & "610900,610921,������,32.891121,108.510946;"
    AddressDistrict = AddressDistrict & "610900,610922,ʯȪ��,33.038512,108.250512;"
    AddressDistrict = AddressDistrict & "610900,610923,������,33.312184,108.313714;"
    AddressDistrict = AddressDistrict & "610900,610924,������,32.520176,108.537788;"
    AddressDistrict = AddressDistrict & "610900,610925,᰸���,32.31069,108.900663;"
    AddressDistrict = AddressDistrict & "610900,610926,ƽ����,32.387933,109.361865;"
    AddressDistrict = AddressDistrict & "610900,610927,��ƺ��,31.883395,109.526437;"
    AddressDistrict = AddressDistrict & "610900,610928,Ѯ����,32.833567,109.368149;"
    AddressDistrict = AddressDistrict & "610900,610929,�׺���,32.809484,110.114186;"
    AddressDistrict = AddressDistrict & "611000,611002,������,33.869208,109.937685;"
    AddressDistrict = AddressDistrict & "611000,611021,������,34.088502,110.145716;"
    AddressDistrict = AddressDistrict & "611000,611022,������,33.694711,110.33191;"
    AddressDistrict = AddressDistrict & "611000,611023,������,33.526367,110.885437;"
    AddressDistrict = AddressDistrict & "611000,611024,ɽ����,33.530411,109.880435;"
    AddressDistrict = AddressDistrict & "611000,611025,����,33.423981,109.151075;"
    AddressDistrict = AddressDistrict & "611000,611026,��ˮ��,33.682773,109.111249;"
    AddressDistrict = AddressDistrict & "620100,620102,�ǹ���,36.049115,103.841032;"
    AddressDistrict = AddressDistrict & "620100,620103,�������,36.06673,103.784326;"
    AddressDistrict = AddressDistrict & "620100,620104,������,36.100369,103.622331;"
    AddressDistrict = AddressDistrict & "620100,620105,������,36.10329,103.724038;"
    AddressDistrict = AddressDistrict & "620100,620111,�����,36.344177,102.861814;"
    AddressDistrict = AddressDistrict & "620100,620121,������,36.734428,103.262203;"
    AddressDistrict = AddressDistrict & "620100,620122,������,36.331254,103.94933;"
    AddressDistrict = AddressDistrict & "620100,620123,������,35.84443,104.114975;"
    AddressDistrict = AddressDistrict & "620300,620302,����,38.513793,102.187683;"
    AddressDistrict = AddressDistrict & "620300,620321,������,38.247354,101.971957;"
    AddressDistrict = AddressDistrict & "620400,620402,������,36.545649,104.17425;"
    AddressDistrict = AddressDistrict & "620400,620403,ƽ����,36.72921,104.819207;"
    AddressDistrict = AddressDistrict & "620400,620421,��Զ��,36.561424,104.686972;"
    AddressDistrict = AddressDistrict & "620400,620422,������,35.692486,105.054337;"
    AddressDistrict = AddressDistrict & "620400,620423,��̩��,37.193519,104.066394;"
    AddressDistrict = AddressDistrict & "620500,620502,������,34.578645,105.724477;"
    AddressDistrict = AddressDistrict & "620500,620503,�����,34.563504,105.897631;"
    AddressDistrict = AddressDistrict & "620500,620521,��ˮ��,34.75287,106.139878;"
    AddressDistrict = AddressDistrict & "620500,620522,�ذ���,34.862354,105.6733;"
    AddressDistrict = AddressDistrict & "620500,620523,�ʹ���,34.747327,105.332347;"
    AddressDistrict = AddressDistrict & "620500,620524,��ɽ��,34.721955,104.891696;"
    AddressDistrict = AddressDistrict & "620500,620525,�żҴ�����������,34.993237,106.212416;"
    AddressDistrict = AddressDistrict & "620600,620602,������,37.93025,102.634492;"
    AddressDistrict = AddressDistrict & "620600,620621,������,38.624621,103.090654;"
    AddressDistrict = AddressDistrict & "620600,620622,������,37.470571,102.898047;"
    AddressDistrict = AddressDistrict & "620600,620623,��ף����������,36.971678,103.142034;"
    AddressDistrict = AddressDistrict & "620700,620702,������,38.931774,100.454862;"
    AddressDistrict = AddressDistrict & "620700,620721,����ԣ����������,38.837269,99.617086;"
    AddressDistrict = AddressDistrict & "620700,620722,������,38.434454,100.816623;"
    AddressDistrict = AddressDistrict & "620700,620723,������,39.152151,100.166333;"
    AddressDistrict = AddressDistrict & "620700,620724,��̨��,39.376308,99.81665;"
    AddressDistrict = AddressDistrict & "620700,620725,ɽ����,38.784839,101.088442;"
    AddressDistrict = AddressDistrict & "620800,620802,�����,35.54173,106.684223;"
    AddressDistrict = AddressDistrict & "620800,620821,������,35.335283,107.365218;"
    AddressDistrict = AddressDistrict & "620800,620822,��̨��,35.064009,107.620587;"
    AddressDistrict = AddressDistrict & "620800,620823,������,35.304533,107.031253;"
    AddressDistrict = AddressDistrict & "620800,620825,ׯ����,35.203428,106.041979;"
    AddressDistrict = AddressDistrict & "620800,620826,������,35.525243,105.733489;"
    AddressDistrict = AddressDistrict & "620800,620881,��ͤ��,35.215341,106.649308;"
    AddressDistrict = AddressDistrict & "620900,620902,������,39.743858,98.511155;"
    AddressDistrict = AddressDistrict & "620900,620921,������,39.983036,98.902959;"
    AddressDistrict = AddressDistrict & "620900,620922,������,40.516525,95.780591;"
    AddressDistrict = AddressDistrict & "620900,620923,�౱�ɹ���������,39.51224,94.87728;"
    AddressDistrict = AddressDistrict & "620900,620924,��������������������,39.631642,94.337642;"
    AddressDistrict = AddressDistrict & "620900,620981,������,40.28682,97.037206;"
    AddressDistrict = AddressDistrict & "620900,620982,�ػ���,40.141119,94.664279;"
    AddressDistrict = AddressDistrict & "621000,621002,������,35.733713,107.638824;"
    AddressDistrict = AddressDistrict & "621000,621021,�����,36.013504,107.885664;"
    AddressDistrict = AddressDistrict & "621000,621022,����,36.569322,107.308754;"
    AddressDistrict = AddressDistrict & "621000,621023,������,36.457304,107.986288;"
    AddressDistrict = AddressDistrict & "621000,621024,��ˮ��,35.819005,108.019865;"
    AddressDistrict = AddressDistrict & "621000,621025,������,35.490642,108.361068;"
    AddressDistrict = AddressDistrict & "621000,621026,����,35.50201,107.921182;"
    AddressDistrict = AddressDistrict & "621000,621027,��ԭ��,35.677806,107.195706;"
    AddressDistrict = AddressDistrict & "621100,621102,������,35.579764,104.62577;"
    AddressDistrict = AddressDistrict & "621100,621121,ͨμ��,35.208922,105.250102;"
    AddressDistrict = AddressDistrict & "621100,621122,¤����,35.003409,104.637554;"
    AddressDistrict = AddressDistrict & "621100,621123,μԴ��,35.133023,104.211742;"
    AddressDistrict = AddressDistrict & "621100,621124,�����,35.376233,103.862186;"
    AddressDistrict = AddressDistrict & "621100,621125,����,34.848642,104.466756;"
    AddressDistrict = AddressDistrict & "621100,621126,���,34.439105,104.039882;"
    AddressDistrict = AddressDistrict & "621200,621202,�䶼��,33.388155,104.929866;"
    AddressDistrict = AddressDistrict & "621200,621221,����,33.739863,105.734434;"
    AddressDistrict = AddressDistrict & "621200,621222,����,32.942171,104.682448;"
    AddressDistrict = AddressDistrict & "621200,621223,崲���,34.042655,104.394475;"
    AddressDistrict = AddressDistrict & "621200,621224,����,33.328266,105.609534;"
    AddressDistrict = AddressDistrict & "621200,621225,������,34.013718,105.299737;"
    AddressDistrict = AddressDistrict & "621200,621226,����,34.189387,105.181616;"
    AddressDistrict = AddressDistrict & "621200,621227,����,33.767785,106.085632;"
    AddressDistrict = AddressDistrict & "621200,621228,������,33.910729,106.306959;"
    AddressDistrict = AddressDistrict & "622900,622901,������,35.59941,103.211634;"
    AddressDistrict = AddressDistrict & "622900,622921,������,35.49236,102.993873;"
    AddressDistrict = AddressDistrict & "622900,622922,������,35.371906,103.709852;"
    AddressDistrict = AddressDistrict & "622900,622923,������,35.938933,103.319871;"
    AddressDistrict = AddressDistrict & "622900,622924,�����,35.481688,103.576188;"
    AddressDistrict = AddressDistrict & "622900,622925,������,35.425971,103.350357;"
    AddressDistrict = AddressDistrict & "622900,622926,������������,35.66383,103.389568;"
    AddressDistrict = AddressDistrict & "622900,622927,��ʯɽ�����嶫����������������,35.712906,102.877473;"
    AddressDistrict = AddressDistrict & "623000,623001,������,34.985973,102.91149;"
    AddressDistrict = AddressDistrict & "623000,623021,��̶��,34.69164,103.353054;"
    AddressDistrict = AddressDistrict & "623000,623022,׿����,34.588165,103.508508;"
    AddressDistrict = AddressDistrict & "623000,623023,������,33.782964,104.370271;"
    AddressDistrict = AddressDistrict & "623000,623024,������,34.055348,103.221009;"
    AddressDistrict = AddressDistrict & "623000,623025,������,33.998068,102.075767;"
    AddressDistrict = AddressDistrict & "623000,623026,µ����,34.589591,102.488495;"
    AddressDistrict = AddressDistrict & "623000,623027,�ĺ���,35.200853,102.520743;"
    AddressDistrict = AddressDistrict & "630100,630102,�Ƕ���,36.616043,101.796095;"
    AddressDistrict = AddressDistrict & "630100,630103,������,36.621181,101.784554;"
    AddressDistrict = AddressDistrict & "630100,630104,������,36.628323,101.763649;"
    AddressDistrict = AddressDistrict & "630100,630105,�Ǳ���,36.648448,101.761297;"
    AddressDistrict = AddressDistrict & "630100,630106,������,36.500419,101.569475;"
    AddressDistrict = AddressDistrict & "630100,630121,��ͨ��������������,36.931343,101.684183;"
    AddressDistrict = AddressDistrict & "630100,630123,��Դ��,36.684818,101.263435;"
    AddressDistrict = AddressDistrict & "630200,630202,�ֶ���,36.480291,102.402431;"
    AddressDistrict = AddressDistrict & "630200,630203,ƽ����,36.502714,102.104295;"
    AddressDistrict = AddressDistrict & "630200,630222,��ͻ�������������,36.329451,102.804209;"
    AddressDistrict = AddressDistrict & "630200,630223,��������������,36.83994,101.956734;"
    AddressDistrict = AddressDistrict & "630200,630224,��¡����������,36.098322,102.262329;"
    AddressDistrict = AddressDistrict & "630200,630225,ѭ��������������,35.847247,102.486534;"
    AddressDistrict = AddressDistrict & "632200,632221,��Դ����������,37.376627,101.618461;"
    AddressDistrict = AddressDistrict & "632200,632222,������,38.175409,100.249778;"
    AddressDistrict = AddressDistrict & "632200,632223,������,36.959542,100.90049;"
    AddressDistrict = AddressDistrict & "632200,632224,�ղ���,37.326263,100.138417;"
    AddressDistrict = AddressDistrict & "632300,632301,ͬ����,35.516337,102.017604;"
    AddressDistrict = AddressDistrict & "632300,632322,������,35.938205,102.031953;"
    AddressDistrict = AddressDistrict & "632300,632323,�����,35.036842,101.469343;"
    AddressDistrict = AddressDistrict & "632300,632324,�����ɹ���������,34.734522,101.611877;"
    AddressDistrict = AddressDistrict & "632500,632521,������,36.280286,100.619597;"
    AddressDistrict = AddressDistrict & "632500,632522,ͬ����,35.254492,100.579465;"
    AddressDistrict = AddressDistrict & "632500,632523,�����,36.040456,101.431856;"
    AddressDistrict = AddressDistrict & "632500,632524,�˺���,35.58909,99.986963;"
    AddressDistrict = AddressDistrict & "632500,632525,������,35.587085,100.74792;"
    AddressDistrict = AddressDistrict & "632600,632621,������,34.473386,100.243531;"
    AddressDistrict = AddressDistrict & "632600,632622,������,32.931589,100.737955;"
    AddressDistrict = AddressDistrict & "632600,632623,�ʵ���,33.966987,99.902589;"
    AddressDistrict = AddressDistrict & "632600,632624,������,33.753259,99.651715;"
    AddressDistrict = AddressDistrict & "632600,632625,������,33.430217,101.484884;"
    AddressDistrict = AddressDistrict & "632600,632626,�����,34.91528,98.211343;"
    AddressDistrict = AddressDistrict & "632700,632701,������,33.00393,97.008762;"
    AddressDistrict = AddressDistrict & "632700,632722,�Ӷ���,32.891886,95.293423;"
    AddressDistrict = AddressDistrict & "632700,632723,�ƶ���,33.367884,97.110893;"
    AddressDistrict = AddressDistrict & "632700,632724,�ζ���,33.852322,95.616843;"
    AddressDistrict = AddressDistrict & "632700,632725,��ǫ��,32.203206,96.479797;"
    AddressDistrict = AddressDistrict & "632700,632726,��������,34.12654,95.800674;"
    AddressDistrict = AddressDistrict & "632800,632801,���ľ��,36.401541,94.905777;"
    AddressDistrict = AddressDistrict & "632800,632802,�������,37.374555,97.370143;"
    AddressDistrict = AddressDistrict & "632800,632803,ã����,38.247117,90.855955;"
    AddressDistrict = AddressDistrict & "632800,632821,������,36.930389,98.479852;"
    AddressDistrict = AddressDistrict & "632800,632822,������,36.298553,98.089161;"
    AddressDistrict = AddressDistrict & "632800,632823,�����,37.29906,99.02078;"
    AddressDistrict = AddressDistrict & "632800,632825,�����ɹ������������ֱϽ,37.853631,95.357233;"
    AddressDistrict = AddressDistrict & "640100,640104,������,38.46747,106.278393;"
    AddressDistrict = AddressDistrict & "640100,640105,������,38.492424,106.132116;"
    AddressDistrict = AddressDistrict & "640100,640106,�����,38.477353,106.228486;"
    AddressDistrict = AddressDistrict & "640100,640121,������,38.28043,106.253781;"
    AddressDistrict = AddressDistrict & "640100,640122,������,38.554563,106.345904;"
    AddressDistrict = AddressDistrict & "640100,640181,������,38.094058,106.334701;"
    AddressDistrict = AddressDistrict & "640200,640202,�������,39.014158,106.376651;"
    AddressDistrict = AddressDistrict & "640200,640205,��ũ��,39.230094,106.775513;"
    AddressDistrict = AddressDistrict & "640200,640221,ƽ����,38.90674,106.54489;"
    AddressDistrict = AddressDistrict & "640300,640302,��ͨ��,37.985967,106.199419;"
    AddressDistrict = AddressDistrict & "640300,640303,���±���,37.421616,106.067315;"
    AddressDistrict = AddressDistrict & "640300,640323,�γ���,37.784222,107.40541;"
    AddressDistrict = AddressDistrict & "640300,640324,ͬ����,36.9829,105.914764;"
    AddressDistrict = AddressDistrict & "640300,640381,��ͭϿ��,38.021509,106.075395;"
    AddressDistrict = AddressDistrict & "640400,640402,ԭ����,36.005337,106.28477;"
    AddressDistrict = AddressDistrict & "640400,640422,������,35.965384,105.731801;"
    AddressDistrict = AddressDistrict & "640400,640423,¡����,35.618234,106.12344;"
    AddressDistrict = AddressDistrict & "640400,640424,��Դ��,35.49344,106.338674;"
    AddressDistrict = AddressDistrict & "640400,640425,������,35.849975,106.641512;"
    AddressDistrict = AddressDistrict & "640500,640502,ɳ��ͷ��,37.514564,105.190536;"
    AddressDistrict = AddressDistrict & "640500,640521,������,37.489736,105.675784;"
    AddressDistrict = AddressDistrict & "640500,640522,��ԭ��,36.562007,105.647323;"
    AddressDistrict = AddressDistrict & "650100,650102,��ɽ��,43.796428,87.620116;"
    AddressDistrict = AddressDistrict & "650100,650103,ɳ���Ϳ���,43.788872,87.596639;"
    AddressDistrict = AddressDistrict & "650100,650104,������,43.870882,87.560653;"
    AddressDistrict = AddressDistrict & "650100,650105,ˮĥ����,43.816747,87.613093;"
    AddressDistrict = AddressDistrict & "650100,650106,ͷ�ͺ���,43.876053,87.425823;"
    AddressDistrict = AddressDistrict & "650100,650107,�������,43.36181,88.30994;"
    AddressDistrict = AddressDistrict & "650100,650109,�׶���,43.960982,87.691801;"
    AddressDistrict = AddressDistrict & "650100,650121,��³ľ����,43.982546,87.505603;"
    AddressDistrict = AddressDistrict & "650200,650202,��ɽ����,44.327207,84.882267;"
    AddressDistrict = AddressDistrict & "650200,650203,����������,45.600477,84.868918;"
    AddressDistrict = AddressDistrict & "650200,650204,�׼�̲��,45.689021,85.129882;"
    AddressDistrict = AddressDistrict & "650200,650205,�ڶ�����,46.08776,85.697767;"
    AddressDistrict = AddressDistrict & "650400,650402,�߲���,42.947627,89.182324;"
    AddressDistrict = AddressDistrict & "650400,650421,۷����,42.865503,90.212692;"
    AddressDistrict = AddressDistrict & "650400,650422,�п�ѷ��,42.793536,88.655771;"
    AddressDistrict = AddressDistrict & "650500,650502,������,42.833888,93.509174;"
    AddressDistrict = AddressDistrict & "650500,650521,������������������,43.599032,93.021795;"
    AddressDistrict = AddressDistrict & "650500,650522,������,43.252012,94.692773;"
    AddressDistrict = AddressDistrict & "652300,652301,������,44.013183,87.304112;"
    AddressDistrict = AddressDistrict & "652300,652302,������,44.152153,87.98384;"
    AddressDistrict = AddressDistrict & "652300,652323,��ͼ����,44.189342,86.888613;"
    AddressDistrict = AddressDistrict & "652300,652324,����˹��,44.305625,86.217687;"
    AddressDistrict = AddressDistrict & "652300,652325,��̨��,44.021996,89.591437;"
    AddressDistrict = AddressDistrict & "652300,652327,��ľ������,43.997162,89.181288;"
    AddressDistrict = AddressDistrict & "652300,652328,ľ�ݹ�����������,43.832442,90.282833;"
    AddressDistrict = AddressDistrict & "652700,652701,������,44.903087,82.072237;"
    AddressDistrict = AddressDistrict & "652700,652702,����ɽ����,45.16777,82.569389;"
    AddressDistrict = AddressDistrict & "652700,652722,������,44.605645,82.892938;"
    AddressDistrict = AddressDistrict & "652700,652723,��Ȫ��,44.973751,81.03099;"
    AddressDistrict = AddressDistrict & "652800,652801,�������,41.763122,86.145948;"
    AddressDistrict = AddressDistrict & "652800,652822,��̨��,41.781266,84.248542;"
    AddressDistrict = AddressDistrict & "652800,652823,ξ����,41.337428,86.263412;"
    AddressDistrict = AddressDistrict & "652800,652824,��Ǽ��,39.023807,88.168807;"
    AddressDistrict = AddressDistrict & "652800,652825,��ĩ��,38.138562,85.532629;"
    AddressDistrict = AddressDistrict & "652800,652826,���Ȼ���������,42.064349,86.5698;"
    AddressDistrict = AddressDistrict & "652800,652827,�;���,42.31716,86.391067;"
    AddressDistrict = AddressDistrict & "652800,652828,��˶��,42.268863,86.864947;"
    AddressDistrict = AddressDistrict & "652800,652829,������,41.980166,86.631576;"
    AddressDistrict = AddressDistrict & "652900,652901,��������,41.171272,80.2629;"
    AddressDistrict = AddressDistrict & "652900,652902,�⳵��,41.717141,82.96304;"
    AddressDistrict = AddressDistrict & "652900,652922,������,41.272995,80.243273;"
    AddressDistrict = AddressDistrict & "652900,652924,ɳ����,41.226268,82.78077;"
    AddressDistrict = AddressDistrict & "652900,652925,�º���,41.551176,82.610828;"
    AddressDistrict = AddressDistrict & "652900,652926,�ݳ���,41.796101,81.869881;"
    AddressDistrict = AddressDistrict & "652900,652927,��ʲ��,41.21587,79.230805;"
    AddressDistrict = AddressDistrict & "652900,652928,��������,40.638422,80.378426;"
    AddressDistrict = AddressDistrict & "652900,652929,��ƺ��,40.50624,79.04785;"
    AddressDistrict = AddressDistrict & "653000,653001,��ͼʲ��,39.712898,76.173939;"
    AddressDistrict = AddressDistrict & "653000,653022,��������,39.147079,75.945159;"
    AddressDistrict = AddressDistrict & "653000,653023,��������,40.937567,78.450164;"
    AddressDistrict = AddressDistrict & "653000,653024,��ǡ��,39.716633,75.25969;"
    AddressDistrict = AddressDistrict & "653100,653101,��ʲ��,39.467861,75.98838;"
    AddressDistrict = AddressDistrict & "653100,653121,�踽��,39.378306,75.863075;"
    AddressDistrict = AddressDistrict & "653100,653122,������,39.399461,76.053653;"
    AddressDistrict = AddressDistrict & "653100,653123,Ӣ��ɳ��,38.929839,76.174292;"
    AddressDistrict = AddressDistrict & "653100,653124,������,38.191217,77.273593;"
    AddressDistrict = AddressDistrict & "653100,653125,ɯ����,38.414499,77.248884;"
    AddressDistrict = AddressDistrict & "653100,653126,Ҷ����,37.884679,77.420353;"
    AddressDistrict = AddressDistrict & "653100,653127,�������,38.903384,77.651538;"
    AddressDistrict = AddressDistrict & "653100,653128,���պ���,39.235248,76.7724;"
    AddressDistrict = AddressDistrict & "653100,653129,٤ʦ��,39.494325,76.741982;"
    AddressDistrict = AddressDistrict & "653100,653130,�ͳ���,39.783479,78.55041;"
    AddressDistrict = AddressDistrict & "653100,653131,��ʲ�����������������,37.775437,75.228068;"
    AddressDistrict = AddressDistrict & "653200,653201,������,37.108944,79.927542;"
    AddressDistrict = AddressDistrict & "653200,653221,������,37.120031,79.81907;"
    AddressDistrict = AddressDistrict & "653200,653222,ī����,37.271511,79.736629;"
    AddressDistrict = AddressDistrict & "653200,653223,Ƥɽ��,37.616332,78.282301;"
    AddressDistrict = AddressDistrict & "653200,653224,������,37.074377,80.184038;"
    AddressDistrict = AddressDistrict & "653200,653225,������,37.001672,80.803572;"
    AddressDistrict = AddressDistrict & "653200,653226,������,36.854628,81.667845;"
    AddressDistrict = AddressDistrict & "653200,653227,�����,37.064909,82.692354;"
    AddressDistrict = AddressDistrict & "654000,654002,������,43.922209,81.316343;"
    AddressDistrict = AddressDistrict & "654000,654003,������,44.423445,84.901602;"
    AddressDistrict = AddressDistrict & "654000,654004,������˹��,44.201669,80.420759;"
    AddressDistrict = AddressDistrict & "654000,654021,������,43.977876,81.524671;"
    AddressDistrict = AddressDistrict & "654000,654022,�첼�������������,43.838883,81.150874;"
    AddressDistrict = AddressDistrict & "654000,654023,������,44.049912,80.872508;"
    AddressDistrict = AddressDistrict & "654000,654024,������,43.481618,82.227044;"
    AddressDistrict = AddressDistrict & "654000,654025,��Դ��,43.434249,83.258493;"
    AddressDistrict = AddressDistrict & "654000,654026,������,43.157765,81.126029;"
    AddressDistrict = AddressDistrict & "654000,654027,�ؿ�˹��,43.214861,81.840058;"
    AddressDistrict = AddressDistrict & "654000,654028,���տ���,43.789737,82.504119;"
    AddressDistrict = AddressDistrict & "654200,654201,������,46.746281,82.983988;"
    AddressDistrict = AddressDistrict & "654200,654202,������,44.430115,84.677624;"
    AddressDistrict = AddressDistrict & "654200,654221,������,46.522555,83.622118;"
    AddressDistrict = AddressDistrict & "654200,654223,ɳ����,44.329544,85.622508;"
    AddressDistrict = AddressDistrict & "654200,654224,������,45.935863,83.60469;"
    AddressDistrict = AddressDistrict & "654200,654225,ԣ����,46.202781,82.982157;"
    AddressDistrict = AddressDistrict & "654200,654226,�Ͳ��������ɹ�������,46.793001,85.733551;"
    AddressDistrict = AddressDistrict & "654300,654301,����̩��,47.848911,88.138743;"
    AddressDistrict = AddressDistrict & "654300,654321,��������,47.70453,86.86186;"
    AddressDistrict = AddressDistrict & "654300,654322,������,46.993106,89.524993;"
    AddressDistrict = AddressDistrict & "654300,654323,������,47.113128,87.494569;"
    AddressDistrict = AddressDistrict & "654300,654324,���ͺ���,48.059284,86.418964;"
    AddressDistrict = AddressDistrict & "654300,654325,�����,46.672446,90.381561;"
    AddressDistrict = AddressDistrict & "654300,654326,��ľ����,47.434633,85.876064;"
    '�����ؼ��к�����Ĳ㼶����
    AddressDistrict = AddressDistrict & "710000,710000,̨��,25.044332,121.509062;"
    AddressDistrict = AddressDistrict & "419001,419001,��Դ,35.090378,112.590047;"
    AddressDistrict = AddressDistrict & "429004,429004,����,30.364953,113.453974;"
    AddressDistrict = AddressDistrict & "429005,429005,Ǳ��,30.421215,112.896866;"
    AddressDistrict = AddressDistrict & "429006,429006,����,30.653061,113.165862;"
    AddressDistrict = AddressDistrict & "429021,429021,��ũ������,31.744449,110.671525;"
    AddressDistrict = AddressDistrict & "441900,441900,��ݸ,23.046237,113.746262;"
    AddressDistrict = AddressDistrict & "442000,442000,��ɽ,22.521113,113.382391;"
    AddressDistrict = AddressDistrict & "460400,460400,����,19.517486,109.576782;"
    AddressDistrict = AddressDistrict & "469001,469001,��ָɽ,18.776921,109.516662;"
    AddressDistrict = AddressDistrict & "469002,469002,��,19.246011,110.466785;"
    AddressDistrict = AddressDistrict & "469005,469005,�Ĳ�,19.612986,110.753975;"
    AddressDistrict = AddressDistrict & "469006,469006,����,18.796216,110.388793;"
    AddressDistrict = AddressDistrict & "469007,469007,����,19.10198,108.653789;"
    AddressDistrict = AddressDistrict & "469021,469021,������,19.684966,110.349235;"
    AddressDistrict = AddressDistrict & "469022,469022,�Ͳ���,19.362916,110.102773;"
    AddressDistrict = AddressDistrict & "469023,469023,������,19.737095,110.007147;"
    AddressDistrict = AddressDistrict & "469024,469024,�ٸ���,19.908293,109.687697;"
    AddressDistrict = AddressDistrict & "469025,469025,��ɳ����������,19.224584,109.452606;"
    AddressDistrict = AddressDistrict & "469026,469026,��������������,19.260968,109.053351;"
    AddressDistrict = AddressDistrict & "469027,469027,�ֶ�����������,18.74758,109.175444;"
    AddressDistrict = AddressDistrict & "469028,469028,��ˮ����������,18.505006,110.037218;"
    AddressDistrict = AddressDistrict & "469029,469029,��ͤ��������������,18.636371,109.70245;"
    AddressDistrict = AddressDistrict & "469030,469030,������������������,19.03557,109.839996;"
    AddressDistrict = AddressDistrict & "620200,620200,������,39.786529,98.277304;"
    AddressDistrict = AddressDistrict & "659001,659001,ʯ����,44.305886,86.041075;"
    AddressDistrict = AddressDistrict & "659002,659002,������,40.541914,81.285884;"
    AddressDistrict = AddressDistrict & "659003,659003,ͼľ���,39.867316,79.077978;"
    AddressDistrict = AddressDistrict & "659004,659004,�����,44.167401,87.526884;"
    AddressDistrict = AddressDistrict & "659005,659005,����,47.353177,87.824932;"
    AddressDistrict = AddressDistrict & "659006,659006,���Ź�,41.827251,85.501218;"
    AddressDistrict = AddressDistrict & "659007,659007,˫��,44.840524,82.353656;"
    AddressDistrict = AddressDistrict & "659008,659008,�ɿ˴���,43.6832,80.63579;"
    AddressDistrict = AddressDistrict & "659009,659009,����,37.207994,79.287372;"
    AddressDistrict = AddressDistrict & "659010,659010,�����,44.69288853,84.8275959"

End Function

Public Function DateStatusDict(Optional ByVal intX As Long = 7) As Object
    '����־״̬д���ֵ�
    '�ֵ�keyΪ���ڣ�ֵΪ���ֵ� ���� index��������0��ʼ�� status������״̬��modx������ �� ���� intX ȡģ���������weekday���ܼ� 1-7 ��ʾ ��-��
    'status��1=>�����գ�2=>���ࣻ3=>��ĩ��4=>���ڣ�5=>��������
    '2016-2023������״̬,ÿ�б�ʾһ��,ÿ�궨�˼��������ں���
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
    
    Set DateStatusDict = CreateObject("Scripting.Dictionary") ' ��ʼ���������ֵ�
    ArrDateStatus = Split(dateStatusList, ",")
    days = UBound(ArrDateStatus)
    
    For d = 0 To days
    
        dateAC = DateStart + d
        
        Set rowDict = CreateObject("Scripting.Dictionary")
        
        rowDict.Add "index", d '������0��ʼ
        rowDict.Add "status", CLng(ArrDateStatus(d)) '����״̬
        rowDict.Add "modx", d Mod intX '���� �� ���� intX ȡģ������� Ĭ��Ϊ 7
        rowDict.Add "weekday", Weekday(dateAC) '�ܼ� 1-7 ��ʾ ��-��

        DateStatusDict.Add dateAC, rowDict '��ӵ����ֵ���
        
        Set rowDict = Nothing
    Next

End Function

Public Function GenderDict() As Object
    'ǰ�� 448 �� ID ���Ա����ͷ���Ѿ�ȷ�ϡ�
    Const GenderList As String = "1,1,1,0,1,1,1,0,1,1,0,1,0,1,0,1,0,0,0,0,0,1,0,1,1,0,1,0,1,0,1,1,0,0,1,0,1,1,1,0,0,0,0,1,0,1,1,1,1,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,1,1,1,0,1,0,1,1,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,1,1,1,0,1,0,1,0,0,1,0,0,0,0,1,0,0,0,0,1,1,1,1,0,0,0,0,0,0,1,0,1,1,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,1,0,0,0,0,0,0,1,1,1,0,0,1,1,0,0,0,1,1,0,0,1,0,0,0,1,0,1,0,0,1,0,1,0,0,1,0,1,1,0,1,0,0,0,0,0,0,1,0,0,1,0,0,0,1,0,0,0,0,1,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,1,1,0,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,1,0,0,1,0,0,1,1,1,0,0,0,1,0,0,0,0,1,1,1,0,0,0,0,0,1,0,1,0,1,0,1,1,0,0,0,0,0,0,0,1,1,0,0,0,1,1,0,1,0,0,1,0,1,0,0,1,0,0,1,1,1,0,1,1,0,0,0,1,1,0,1,1,0,0,0,0,1,0,0,0,1,0,1,0,0,0,1,1,1,1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1,1,0,0,0,0,1,0,0,1,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,1,0,1,0,1,1,0,0,0,0,0,0,0,1,1,0,0,0,1,1,0,1,0,0,1,0,1,0,0,1,0,0,0,0,1,0,1,1,0,0,0,1,1,0,1,1,0,0,0,0,1,0,0,0,1,0,1,0,0,0,0,0,1,1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1"

    Const IDStart As Long = 10001

    Dim ID As Long, gender As String
    Dim ArrGender
    Dim n As Long, num As Long
    
    Set GenderDict = CreateObject("Scripting.Dictionary") ' ��ʼ���������ֵ�
    ArrGender = Split(GenderList, ",")
    num = UBound(ArrGender)
    
    For n = 0 To num
        ID = IDStart + n
        gender = "��"
        If ArrGender(n) = 0 Then gender = "Ů"
        GenderDict.Add ID, gender '��ӵ����ֵ���
    Next

End Function

Public Function AddMonths(d As Date, n As Long) As Date
    '���������·�
    AddMonths = DateSerial(Year(d), Month(d) + n, day(d))
End Function

Public Function GetDaysInMonth(d As Date) As Long
    '��ȡ�����ж�����
    GetDaysInMonth = day(DateSerial(Year(d), Month(d) + 1, 1) - 1)
End Function

Public Function GetMonthStart(d As Date) As Date
    '�������ڻ�ȡ�³�����
    GetMonthStart = DateSerial(Year(d), Month(d), 1)
End Function

Public Function DateDiffInMonths(DateStart As Date, DateEnd As Date) As Long
    '�������ڼ���·ݲ�������
    Dim YearsDiff As Long
    Dim MonthsDiff As Long
    
    YearsDiff = Year(DateEnd) - Year(DateStart)
    MonthsDiff = Month(DateEnd) - Month(DateStart)
    
    DateDiffInMonths = (YearsDiff * 12) + MonthsDiff
    
    ' ������Ҫ�����������촦���߼�
    If day(DateEnd) < day(DateStart) Then
        DateDiffInMonths = DateDiffInMonths - 1
    End If
End Function

Public Function AddDictByKey(ByRef dictTarget As Object, ByVal KeyTarget As String, ByVal newValue As Long) As Object
    '�����ֵ�ֵ�ۼ�
    Dim oldValue As Long
    If dictTarget.Exists(KeyTarget) Then
        oldValue = dictTarget(KeyTarget)
        dictTarget(KeyTarget) = oldValue + newValue
    Else
        dictTarget.Add KeyTarget, newValue
    End If
    
End Function

Public Function Main()
    ' ��ں��� ��������

    Dim t As Double
    t = timer
    Dim i As Long
    Dim pbRndInt  As Integer
    Dim pbLeftInt  As Integer
    Dim key As Variant
    Dim keyStr As String
    Dim valueStr As String
    
    productQuantity = 200       '��Ʒ����������ShopQuantity��[7,1688]��
    ShopQuantity = 1            '�ŵ�����������ShopQuantity��[1,390]��
    MaxInventoryDays = 14       '����������������ShopQuantity��[5,20]��
    
    InitTables
    ' �����ֵ�ļ���ֵ
    For Each key In TableNameDict.Keys
        
        keyStr = CStr(key)
        valueStr = CStr(TableNameDict(key))
'        Debug.Print valueStr
        
        ' ADO �½���
        Call TableADO(keyStr, SQLDrop(keyStr), valueStr)
        
    Next key

    DataTableRegion             ' ����
    DataTableProvince           ' ʡ��
    DataTableCity               ' ����
    DataTableDistrict           ' ����
    DataTableProduct            ' ��Ʒ
    DataTableShop               ' �ŵ�
    DataTableEmployeeExecutives ' Ա����߹�
    DataTableOrg                ' ��֯
    DataTableShopRD             ' �ŵ����޺�װ��
    DataTableEmployeeRegular    ' Ա����һ��
    DataTableCustomer           ' �ͻ�
    DataTableSOS                ' ��⡢�������������ӱ�
    DataTableSaleTarget         ' ����Ԥ��
    DataTableLaborCost          ' �˹��ɱ�


    Application.RefreshDatabaseWindow

    MsgBox "��ɣ���ʱ��" & Round(timer - t, 2) & "�룡"

End Function


