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
    '默认 1-5
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
        MsgBox "产品数量：请输入 1 - 2000 的整数！"
        .TextProduct.SetFocus
        GoTo NothingLabel
    End If
    If Not ValidateIng(.TextShop, 1, 400) Then
        MsgBox "门店数量：请输入 1 - 400 的整数！"
        .TextShop.SetFocus
        GoTo NothingLabel
    End If
    
    If Not ValidateIng(.TextDays, 1, 40) Then
        MsgBox "入库间隔：请输入 5 - 20 的整数！"
        .TextDays.SetFocus
        GoTo NothingLabel
    End If
    
    result = MsgBox("模拟数据需要较长时间，是否确认继续？", vbYesNo + vbQuestion, "请作出选择")
    
    If result = vbNo Then
        .CommandRun.SetFocus
        GoTo NothingLabel
    End If

    formStatus True '初始化状态
    
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
    ' 遍历字典的键和值
    For Each key In TableNameDict.Keys
        
        keyStr = CStr(key)
        valueStr = CStr(TableNameDict(key))
'        Debug.Print valueStr
        
        ' ADO 新建表
        Call TableADO(keyStr, SQLDrop(keyStr), valueStr)
        
    Next key
    DoEvents
    
        .BoxUp.Width = percentWidth * 5 + .BoxUp.Width
        
    DataTableRegion             ' 大区
    DataTableProvince           ' 省份
    DataTableCity               ' 城市
    DataTableDistrict           ' 区县
        refreshUI pbLeftInt, 10
    DataTableProduct            ' 产品
        refreshUI pbLeftInt, 10
    DataTableShop               ' 门店
        refreshUI pbLeftInt, 10
    DataTableEmployeeExecutives ' 员工表高管
        refreshUI pbLeftInt, 10
    DataTableOrg                ' 组织
    DataTableShopRD             ' 门店租赁和装修
    DoEvents
        refreshUI pbLeftInt, 10
    DataTableEmployeeRegular    ' 员工表一线
    DataTableCustomer           ' 客户
        refreshUI pbLeftInt, 10
    DataTableSOS                ' 入库、订单主表、订单子表
        refreshUI pbLeftInt, 20
    DataTableSaleTarget         ' 销售预算
    DataTableLaborCost          ' 人工成本
        refreshUI pbLeftInt, 0, True
    
    MsgBox "完成，用时：" & Round(timer - StartTime, 2) & "秒！"
    
    formStatus False '回到初始化状态
    
End With

NothingLabel: '无效标签

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
    'TURE 运行状态 FALSE 初始化状态
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
            .LabelProgress.Caption = "数据模拟生已完成!" & vbCrLf & "进度：100%"
        Else
        
            pbRndInt = Round(pbRndInt * Rnd)
            pbLeftInt = pbLeftInt - pbRndInt
            .BoxUp.Width = percentWidth * pbRndInt + .BoxUp.Width
            .LabelProgress.Caption = "数据模拟生成中..." & vbCrLf & "进度：" & Round(.BoxUp.Width / MaxWidth * 100) & "%"
        End If
    End With
    
    Application.RefreshDatabaseWindow
End Function


