VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public StartTime As Double
Public MaxWidth As Double
Public percentWidth As Double

Function ValidateIng(value As Variant, Optional lowerLimit As Long = 1, Optional upperLimit As Long = 5) As Boolean
    'Ĭ�� 1-5
    If IsNumeric(value) Then
        Dim intValue As Long
        intValue = CLng(value)

        If intValue >= lowerLimit And intValue <= upperLimit Then
            ValidateIng = True
        Else
            ValidateIng = False
        End If

    Else
        ValidateIng = False
    End If
End Function


Private Sub CommandClose_Click()
    DoCmd.Save
    Application.Quit
End Sub

Private Sub CommandRun_Click()

With Me

    If Not ValidateIng(.TextProduct, 1, 2000) Then
        MsgBox "��Ʒ������������ 1 - 2000 ��������"
        .TextProduct.SetFocus
        GoTo NothingLabel
    End If
    If Not ValidateIng(.TextShop, 1, 400) Then
        MsgBox "�ŵ������������� 1 - 400 ��������"
        .TextShop.SetFocus
        GoTo NothingLabel
    End If
    
    If Not ValidateIng(.TextDays, 1, 40) Then
        MsgBox "������������� 5 - 20 ��������"
        .TextDays.SetFocus
        GoTo NothingLabel
    End If
    
    result = MsgBox("ģ��������Ҫ�ϳ�ʱ�䣬�Ƿ�ȷ�ϼ�����", vbYesNo + vbQuestion, "������ѡ��")
    
    If result = vbNo Then
        .CommandRun.SetFocus
        GoTo NothingLabel
    End If

    formStatus True '��ʼ��״̬
    
    productQuantity = .TextProduct.value
    ShopQuantity = .TextShop.value
    MaxInventoryDays = .TextDays.value
    
    
    StartTime = timer
    
    Dim i As Long
    Dim pbRndInt  As Long
    Dim pbLeftInt  As Long
    Dim key As Variant
    Dim keyStr As String
    Dim valueStr As String

    pbLeftInt = 100
    InitTables
    ' �����ֵ�ļ���ֵ
    For Each key In TableNameDict.Keys
        
        keyStr = CStr(key)
        valueStr = CStr(TableNameDict(key))
'        Debug.Print valueStr
        
        ' ADO �½���
        Call TableADO(keyStr, SQLDrop(keyStr), valueStr)
        
    Next key
    DoEvents
    
        .BoxUp.Width = percentWidth * 5 + .BoxUp.Width
        
    DataTableRegion             ' ����
    DataTableProvince           ' ʡ��
    DataTableCity               ' ����
    DataTableDistrict           ' ����
        refreshUI pbLeftInt, 10
    DataTableProduct            ' ��Ʒ
        refreshUI pbLeftInt, 10
    DataTableShop               ' �ŵ�
        refreshUI pbLeftInt, 10
    DataTableEmployeeExecutives ' Ա�����߹�
        refreshUI pbLeftInt, 10
    DataTableOrg                ' ��֯
    DataTableShopRD             ' �ŵ����޺�װ��
    DoEvents
        refreshUI pbLeftInt, 10
    DataTableEmployeeRegular    ' Ա����һ��
    DataTableCustomer           ' �ͻ�
        refreshUI pbLeftInt, 10
    DataTableSOS                ' ��⡢���������������ӱ�
        refreshUI pbLeftInt, 20
    DataTableSaleTarget         ' ����Ԥ��
    DataTableLaborCost          ' �˹��ɱ�
        refreshUI pbLeftInt, 0, True
    
    MsgBox "��ɣ���ʱ��" & Round(timer - StartTime, 2) & "�룡"
    
    formStatus False '�ص���ʼ��״̬
    
End With

NothingLabel: '��Ч��ǩ

End Sub

Private Sub Form_Load()

    formStatus (False)
    With Me
        .TextProduct = 200
        .TextShop = 5
        .TextDays = 14
    End With

End Sub


Function formStatus(status As Boolean)
    'TURE ����״̬ FALSE ��ʼ��״̬
    With Me

        .BoxUp.Visible = status
        .BoxDown.Visible = status

        .LabelProgress.Visible = status
        .CommandRun.Enabled = Not (status)
        .BoxUp.Width = 0

        MaxWidth = .BoxDown.Width
        percentWidth = MaxWidth / 100
    End With
End Function

Function refreshUI(pbLeftInt As Long, pbRndInt As Long, Optional isLast As Boolean = False)
    DoEvents

    With Me
        If isLast Then
            .BoxUp.Width = percentWidth * pbLeftInt + .BoxUp.Width
            .LabelProgress.Caption = "����ģ���������!" & vbCrLf & "���ȣ�100%"
        Else
        
            pbRndInt = Round(pbRndInt * Rnd)
            pbLeftInt = pbLeftInt - pbRndInt
            .BoxUp.Width = percentWidth * pbRndInt + .BoxUp.Width
            .LabelProgress.Caption = "����ģ��������..." & vbCrLf & "���ȣ�" & Round(.BoxUp.Width / MaxWidth * 100) & "%"
        End If
    End With
    
    Application.RefreshDatabaseWindow
End Function

