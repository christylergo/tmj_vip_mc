    
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
Dim View As String
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''



Sub VIPandMCSJC_TMJ_Supply()

'    On Error GoTo Anchor
    
    Dim i&, j&, M&, Counter&, Position(1 To 2) As Long, initialLenth&
    Dim RowCounter&, ColCounter&
    Dim arrCache(), arrAtom(), arrInitial(), SHT As Worksheet
    Dim isHuoPinID As Boolean, isShangPinID As Boolean, isSKUID As Boolean, isZhouZhuan As Boolean, isShangXin As Boolean, isAddition As Boolean
    Dim ZhouZhuan&, Deduplicate As String, intBff As Integer, strBff As String, strBff1 As String, strBff2 As String
    Dim valid_TableDetail As Boolean, isAll As Boolean, Refresh_Date As String
    Dim Identifier(), arrAdd(), addPos(1 To 2) As Long
    
    View = changeVision.Caption
''''''''''''''''''''''''''''''''''
    
    If View = "唯品视角" Then
        Identifier = Array("唯品条码", "唯品后台条码", "款号", "唯品款号", "货号", "唯品货号")
    Else
        Identifier = Array("货品ID", "货品编码", "商品ID", "商品编码", "SKUID", "SKU编码")
    End If

    Set SHT = ThisWorkbook.Sheets(1)
'    RowCounter = SHT.UsedRange.Rows.Count ''''使用usedrange不能正确标识已用单元格'''
    RowCounter = 9999 '''''采用这种硬编码写入数量''''
    ColCounter = 99
    arrCache = SHT.Range(SHT.Cells(1, 1), SHT.Cells(RowCounter + 1, ColCounter + 1)).Value '''范围大一点点，避免查找skuID时数组下标越界'''
    
    '''''找到初始数据的位置,并且识别出是那种类型的数据''''
    If View = "唯品视角" Then
        strBff = "*[0-9a-zA-Z]*"
    Else
        strBff = "*[0-9]*"
    End If
    
    Dim blBff As Boolean
    For i = 1 To RowCounter
        For j = 1 To ColCounter
            blBff = Mid(CStr(arrCache(i + 1, j) & "###"), 1, 1) Like strBff
            isHuoPinID = blBff And CBool(InStr(1, CStr(arrCache(i, j)), Identifier(0), 1)) Or CBool(InStr(1, CStr(arrCache(i, j)), Identifier(1), 1)) '''''第4个参数1，表示不区分大小写
            isShangPinID = blBff And CBool(InStr(1, CStr(arrCache(i, j)), Identifier(2), 1)) Or CBool(InStr(1, CStr(arrCache(i, j)), Identifier(3), 1))
            isSKUID = CBool(InStr(1, CStr(arrCache(i, j + 1)), Identifier(4), 1)) Or CBool(InStr(1, CStr(arrCache(i, j + 1)), Identifier(5), 1))
            isAll = CBool(InStr(1, CStr(arrCache(i, j)), "全部", 1)) Or CBool(InStr(1, CStr(arrCache(i, j)), "所有", 1))
            If isHuoPinID Or isShangPinID Or isAll Then
                Position(1) = i + 1
                Position(2) = j
                Exit For
            End If
        Next
        If isHuoPinID Or isShangPinID Or isAll Then
            Exit For
        End If
    Next
    
    If i = 10000 And j = 100 Then
'        MsgBox "什么信息都没有！"
        i = 1
        j = 1
        Position(1) = 1
        Position(2) = 1
        isAll = True
    End If
    
    initialLenth = IIf(isAll, 1, 0)
    
    Do While Not isAll And Mid(CStr(arrCache(i + 1, Position(2)) & "###"), 1, 1) Like strBff
        ''''''如果skuid存在，则商品ID加上skuid才能成为主键,这里做预处理，方便接下来的操作''''''''''''
        If isSKUID Then
            ''''''''''''''''skuid值是0，会当做空字符""来处理，提兼容性'''''
            arrCache(i + 1, Position(2)) = Trim(CStr(arrCache(i + 1, Position(2)))) & IIf(Trim(CStr(arrCache(i + 1, j + 1))) = "0", "#", "#" & Trim(CStr(arrCache(i + 1, j + 1))))
        End If
        initialLenth = initialLenth + 1
        i = i + 1
        If i = UBound(arrCache, 1) Then Exit Do
    Loop
    If initialLenth = 0 Then
        isHuoPinID = False
        isShangPinID = False
        isSKUID = False
        isAll = True
        initialLenth = 1
    End If
    Counter = 1
    ReDim arrInitial(1 To initialLenth)
    

    
    '''''''''查看是否设置了周转天数以及上新''''''''''
    ZhouZhuan = 0
    For i = 1 To RowCounter
        For j = 1 To ColCounter
            isZhouZhuan = CBool(arrCache(i, j) = "周转")
            isShangXin = CBool(InStr(1, CStr(arrCache(i, j)), "上新", 1))
            If isZhouZhuan Then
                strBff = arrCache(i + 1, j) & arrCache(i, j + 1)
                intBff = InStr(1, strBff, "天", vbTextCompare)
                If intBff > 0 Then
                    arrCache(i + 1, j) = Trim(Mid(strBff, 1, intBff - 1))
                    arrCache(i, j + 1) = ""
                End If
                If Mid(CStr(arrCache(i + 1, j) & "ABC"), 1, 1) Like "*[0-9]*" Then
                    ZhouZhuan = arrCache(i + 1, j) * 1
                End If
                If Mid(CStr(arrCache(i, j + 1) & "ABC"), 1, 1) Like "*[0-9]*" Then
                    ZhouZhuan = arrCache(i, j + 1) * 1
                End If
            End If
            
            If isZhouZhuan Or isShangXin Then
                Exit For
            End If
        Next
        If isZhouZhuan Or isShangXin Then
            Exit For
        End If
    Next
    ZhouZhuan = IIf(ZhouZhuan = 0, 50, ZhouZhuan)
    
    For i = 1 To RowCounter
        For j = 1 To ColCounter
            isAddition = CBool(InStr(1, CStr(arrCache(i, j)), "加量", 1))
            If isAddition Then
                Exit For
            End If
        Next
        If isAddition Then
            addPos(1) = i + 1
            addPos(2) = j
            ReDim arrAdd(1 To initialLenth)
            For M = addPos(1) To addPos(1) + initialLenth - 1
                If Mid(CStr(arrCache(M, addPos(2)) & "ABC"), 1, 1) Like "*[0-9]*" Then
                    Exit For
                End If
            Next
            If M > addPos(1) + initialLenth - 1 Then
                isAddition = False
            End If
            Exit For
        End If
    Next
    
    
    
    '''''''''''''''''''''''去重''''''''''''''''
    Deduplicate = "Blank!"
    If Not isAll Then
        Dim initialTemp()
        Counter = 1
        For i = Position(1) To Position(1) + initialLenth - 1
            intBff = InStr(1, Deduplicate, arrCache(i, Position(2)) & "#", vbTextCompare)
            If intBff = 0 Then  ''''''剔除重复的'''''
                arrInitial(Counter) = Trim(CStr(arrCache(i, Position(2))))
                If isAddition Then
                    arrAdd(Counter) = arrCache(i, addPos(2)) '''''如果有加量列则这一列和行和主键行是对应的'''
                End If
                Counter = Counter + 1
            End If
            Deduplicate = Deduplicate & arrCache(i, Position(2)) & "#"
        Next
        ReDim initialTemp(1 To Counter - 1)
        For i = 1 To Counter - 1
            initialTemp(i) = arrInitial(i)
        Next
        Erase arrInitial
        arrInitial = initialTemp
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    Erase arrCache
    
    
    Call get_BaseTableDetail(arrCache, arrAtom, arrInitial, isHuoPinID, isShangPinID, isSKUID, isAll, arrAdd, isAddition, valid_TableDetail)

    If Not valid_TableDetail Then
        Exit Sub
    End If
    
    Call TableFormat(SHT)

    Counter = UBound(arrCache, 1)
    
    Application.ScreenUpdating = False ''''''''''''''''关闭屏幕刷新'''''
    
    With ThisWorkbook.Worksheets("越中仓单品明细")
        .Range(.Cells(2, 1), .Cells(UBound(arrAtom, 1) + 1, 20)) = arrAtom
    End With
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Dim RecentWeek()
    RecentWeek = Array(Date - 7, Date - 6, Date - 5, Date - 4, Date - 3, Date - 2, Date - 1, "成本", "重量", "供货价", "单件运费")
    For i = 0 To 6
        RecentWeek(i) = Month(RecentWeek(i)) & "/" & Day(RecentWeek(i))
    Next
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If View = "唯品视角" Then '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Call SalesExplose(SHT, arrCache)
        SHT.Range(SHT.Cells(2, 1), SHT.Cells(Counter + 1, UBound(arrCache, 2) - 1)).Value = arrCache
        SHT.Range("T1:AA1") = RecentWeek
        
    Else ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Call Calculator(arrCache, ZhouZhuan, isShangXin) '''''''计算锁定以及释放''''唯品视角不需要这个操作
        SHT.Range("Y1:AI1") = RecentWeek
        Dim intBff1&, intBff2&, BookMark&
        With SHT
            BookMark = 1
            Do While BookMark < Counter
                strBff1 = ""
                strBff2 = ""
                intBff1 = 0
                intBff2 = 0
                For i = BookMark To Counter
                    BookMark = i
                    If arrCache(i, 24) = "释放" Then
                        strBff1 = strBff1 & "A" & i + 1 & ":AI" & i + 1 & ","
                        intBff1 = intBff1 + 1
                        If intBff1 > 20 Then Exit For ''''''range一次不能处理太多个联合区域，所以加上dowhile循环来多次完成'''
                    End If
                    If arrCache(i, 24) = "特殊" Then
                        strBff2 = strBff2 & "A" & i + 1 & ":AI" & i + 1 & ","
                        intBff2 = intBff2 + 1
                        If intBff2 > 20 Then Exit For
                    End If
                Next
                If strBff1 <> "" Then
                    strBff1 = Mid(strBff1, 1, Len(strBff1) - 1)
'                    .Range(strBff1).Interior.Color = 65535
                End If
                If strBff2 <> "" Then
                    strBff2 = Mid(strBff2, 1, Len(strBff2) - 1)
                    .Range(strBff2).Interior.ThemeColor = xlThemeColorAccent1
                End If
            Loop
        End With
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Call SalesExplose(SHT, arrCache) ''''''放在后面，不会被之前的格式覆盖''''
        arrCache(1, 23) = ZhouZhuan
        SHT.Range(SHT.Cells(2, 1), SHT.Cells(Counter + 1, UBound(arrCache, 2))).Value = arrCache
        
    End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.ScreenUpdating = True '''''''''''''''''''开启屏幕刷新''''''''''''
    
    Erase arrCache
    
    Application.DisplayAlerts = False
        ThisWorkbook.Save
    Application.DisplayAlerts = True
'Anchor:    Stop

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''猫超功能模组'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function Calculator(ByRef Arr(), ByVal ZhouZhuan, Optional isSX As Boolean = False)
    Dim i As Integer, j As Integer
    Dim intBff&, strBff As String
    Dim Coefficient As Integer
    Dim Requirement()
    Dim Trigger As Boolean
    Dim toRelease As Boolean
'    On Error GoTo anchor
    
    For i = 1 To UBound(Arr, 1)
        For j = 1 To UBound(Arr, 1)
            Trigger = (Arr(i, 13) <> "") Or (Arr(i, 4) = Arr(j, 4) And Arr(j, 13) <> "")
            Trigger = Trigger And Not (Arr(i, 10) * 1 = 0 And Arr(i, 13) * 1 <> 0)
            Trigger = (Trigger Or isSX) And (Arr(i, 3) <> "")
            If Trigger Then
                Exit For
            End If
        Next
        
        If Trigger Then
            Select Case Arr(i, 16) * 1 '''''空值不能自动转换成0，和表格公式不同'''
                Case Is >= 8000
                    Requirement = Array(Int(Arr(i, 16) * 0.25), Int(Arr(i, 16) * 0.9))
                Case Is >= 5000
                    Requirement = Array(Int(Arr(i, 16) * 0.25), Int(Arr(i, 16) * 0.9))
                Case Is >= 3000
                    Requirement = Array(Int(Arr(i, 16) * 0.3), Int(Arr(i, 16) * 0.9))
                Case Is >= 1500
                    Requirement = Array(Int(Arr(i, 16) * 0.3), Int(Arr(i, 16) * 0.85))
                Case Is >= 1000
                    Requirement = Array(Int(Arr(i, 16) * 0.3), Int(Arr(i, 16) * 0.85))
                Case Is >= 600
                    Requirement = Array(Int(Arr(i, 16) * 0.4), Int(Arr(i, 16) * 0.85))
                Case Is >= 400
                    Requirement = Array(Int(Arr(i, 16) * 0.4), Int(Arr(i, 16) * 0.8))
                Case Is >= 300
                    Requirement = Array(Int(Arr(i, 16) * 0.4), Int(Arr(i, 16) * 0.8))
                Case Is >= 200
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.8))
                Case Is >= 150
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.8))
                Case Is >= 100
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.8))
                Case Is >= 70
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.75))
                Case Is >= 50
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.75))
                Case Is >= 30
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.75))
                Case Is >= 20
                    Requirement = Array(Int(Arr(i, 16) * 0.45), Int(Arr(i, 16) * 0.75))
                Case Else
                    Requirement = Array(0, 0)
            End Select
            
            Coefficient = 10 ''''默认以10作为一组补货单位''''
            
            strBff = "642200021630#641430276641#640773636304#641486140935#643155809849#642506936102#654936005492#654490172801" '''''特殊商品的id，包括浴盆，待产包，两件套文胸，SH1131文胸，SH775一次性能看,39件待产包''''
                        
            ''''''''
            ZhouZhuan = Application.Max(ZhouZhuan, 10) ''''''''周转不低于10天
            intBff = Round(((Arr(i, 10) / 28 * ZhouZhuan * 2 + Requirement(0)) / 3), 0)
            intBff = Application.Min(intBff, Requirement(1))
            intBff = Application.Max((intBff - Arr(i, 13) * 1), 0)
            If InStr(1, strBff, Arr(i, 4), vbTextCompare) > 0 And Arr(i, 9) <> "SH1002" Then '''''特殊商品的补货数量不考虑旺店通库存''''
                intBff = Application.Max(500 - Arr(i, 13), 0)
                Arr(i, 24) = "特殊"
                
            End If
            If intBff < 30 Then
                Coefficient = 5 ''''需求小于30的时候补货单位改为5''''
            End If
                 
            intBff = Round(intBff / Coefficient + 0.375, 0) * Coefficient
            
            If Arr(i, 13) * 1 > Arr(i, 16) * 1 And Arr(i, 10) * 1 > Arr(i, 16) * 1 Then
                ''''''浴盆和待产包情况特殊，旺店通库存不能真实反映可发货数量,所以排除'''''
                If InStr(1, strBff, Arr(i, 4), vbTextCompare) = 0 Then
                    Arr(i, 20) = Round((Arr(i, 13) - Arr(i, 16) * 6 / 7) / 5 + 0.25, 0) * 5
                    Arr(i, 24) = "释放"
                    intBff = Empty
                End If
            End If
            If Arr(i, 16) * 1 < 15 And Arr(i, 13) * 1 <> 0 Then
                If InStr(1, strBff, Arr(i, 4), vbTextCompare) = 0 Then
                    Arr(i, 20) = Arr(i, 13)
                    Arr(i, 24) = "释放"
                    intBff = Empty
                End If
            End If
            If Arr(i, 10) > 0 Then
                If Arr(i, 13) / Arr(i, 10) * 28 > 200 Then
                    intBff = Empty
                End If
            End If
            If Arr(i, 2) = "" Or Arr(i, 2) = "下架" Then
                intBff = Empty
            End If
            
            Arr(i, 15) = IIf(intBff = 0, Empty, intBff)

'            Erase Requirement
        End If
    Next
    
'anchor: Stop
End Function


Function SalesExplose(ByRef SHT As Worksheet, ByRef arrCache())
    Dim i&, j&, k&, M&, Counter&, strBff$
    Dim intBff&, BookMark&, DaylyPos()
    Dim Mark As Boolean, NotePos&, Days&, FWeekSales&, Baseline&
    
    k = 2 ''''标记为爆发所需的倍数'''
    M = 20
    NotePos = 18
    Days = 28
    FWeekSales = 10
    Baseline = 50
    Counter = UBound(arrCache, 1)
    DaylyPos = Array("T", "U", "V", "W", "X", "Y", "Z")
    If View = "猫超视角" Then
        k = 3
        M = 25
        NotePos = 22
        Days = 33
        FWeekSales = 11
        Baseline = 20
        DaylyPos = Array("Y", "Z", "AA", "AB", "AC", "AD", "AE")
    End If
    With SHT
        BookMark = 1
        
        Do While BookMark < Counter
            
            For i = BookMark To Counter
                BookMark = i
                If arrCache(i, Days) > 0 Then
                    For j = 0 To 6
                        If (arrCache(i, j + M) > arrCache(i, FWeekSales) / arrCache(i, Days) * k And arrCache(i, FWeekSales) / arrCache(i, Days) > Baseline) _
                        Or (arrCache(i, j + M) > arrCache(i, FWeekSales) / arrCache(i, Days) * 10 And arrCache(i, j + M) > 30) Then '''''10倍以上爆发不考虑基础销量''''
                            Mark = True
                            strBff = strBff & DaylyPos(j) & i + 1 & ","
                            intBff = intBff + 1
'                            If intBff > 20 Then Exit For ''''''range一次不能处理太多个联合区域，所以加上dowhile循环来多次完成'''
                        End If
                    Next
                    If Mark = True Then
                        arrCache(i, NotePos) = "销量爆发!(" & k & "倍以上)" & arrCache(i, NotePos)
                        strBff = Mid(strBff, 1, Len(strBff) - 1)
                        .Range(strBff).Interior.Color = vbRed
                        .Range(strBff).Font.Bold = True
                 
                        strBff = Empty
                        intBff = Empty
                    End If
                    Mark = False
                End If
            Next
'            If strBff <> "" Then
'                strBff = Mid(strBff, 1, Len(strBff) - 1)
'                .Range(strBff).Interior.Color = vbRed
'                .Range(strBff).Font.Bold = True
'            End If
            strBff = Empty
            intBff = Empty
            strCache = Empty
            Mark = False
        Loop
    End With

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function get_BaseTableDetail(ByRef Arr(), ByRef arrAtom(), ByRef arrInitial(), isHPID As Boolean, isSPID As Boolean, isSKUID As Boolean, isAll As Boolean, _
ByRef arrAdd(), ByVal isAddition As Boolean, ByRef Valid As Boolean)

    Dim keyColumn&, i&, j&, k&, M&, n&, Mark As Boolean, Counter&
    Dim ArrBase(), arrSales(), arrStock(), arrJitStock(), arrITL(), arrWDT(), arrUnit(), arrBff() '''''arritl 商品状态列表
    Dim Multi As Integer, MultiValid As Boolean '''组合中商品的种类
    Dim strBff As String, intBff& ''''处理复杂组合时使用的变量''''
    Dim Pointer&, Path$ '''长度不同数组增长填充时使用的变量'''''path唯品全表的路径'''
    
    Dim BarCodePool As String, Frequency&, Equivalent& ''''''全部的涉及到的货品旺店通编码合并放在一起''''
    Dim valid_ArrBase As Boolean, valid_arrSales As Boolean, valid_arrStock As Boolean, valid_arrJitStock As Boolean, valid_arrWDT As Boolean, valid_arrITL As Boolean
    Dim subArr() As Long, subArrBase() As Long, subArrSales() As Long, subArrStock() As Long, subArrJitStock() As Long, subArrWDT() As Long, subArrITL() As Long, intBff_sort&, Position& '''''针对多个可用库存表的情况做的调整''''
    Dim subArrUnit() As Long, valid_arrUnit As Boolean, subArrAtom() As Long, valid_arrAtom As Boolean
    Dim subArrWDTMC() As Long, arrWDTMC(), valid_arrWDTMC As Boolean, subArrWDTBULK() As Long, arrWDTBULK(), valid_arrWDTBULK As Boolean
    Dim WDTCategory(3), Cost(), subArrCost() As Long, arrCost(), valid_arrCost As Boolean, CostKeyPos&, CostValPos&
    
    Dim arrBase_Vice(), subArrBase_Vice() As Long, valid_arrBase_Vice As Boolean ''''猫超基础表''''
    Dim arrSales_Vice(), subarrSales_Vice() As Long, valid_arrSales_Vice As Boolean ''''''猫超销售日报
    Dim SalesKeyPos&, SalesValPos&, Sales_ViceKeyPos&, Sales_ViceValPos&, ITLKeyPos&, ITLValPos&, UnitKeyPos&, UnitValkeyPos&, UnitValvalPos&, UnitNamePos&, UnitArtNoPos&
    Dim StockKeyPos&, StockValPos&, JitStockKeyPos&, JitStockTimePos&, JitStockValPos&, intBffStock&, kStock&, intBffJitStock&, kJitStock&, DatePos& '''''article number货号销售表的日期所在位置''''
    Dim AtomKeyPos&, AtomValPos&, AtomValvalPos&, AtomArtNoPos&, AtomNamePos&, AtomSpecificationPos&, WDTKeyPos&, WDTValPos&, WDTNamePos&, WDTValidPos&
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SalesKeyPos = 5 '''''默认唯品模式的数值'''
    SalesValPos = 11
    Sales_ViceKeyPos = 12
    Sales_ViceValPos = 15
    ITLKeyPos = 8
    ITLValPos = 23
    StockKeyPos = 1
    StockValPos = 6
    JitStockKeyPos = 2
    JitStockValPos = 19
    JitStockTimePos = 20
    UnitKeyPos = 3
    UnitValkeyPos = 16
    UnitValvalPos = 17
    UnitNamePos = 1
    UnitArtNoPos = 15
    AtomKeyPos = 1
    AtomValPos = 0
    AtomNamePos = 3
    AtomArtNoPos = 2
    AtomSpecificationPos = 7
    WDTKeyPos = 1
    WDTValPos = 4
    WDTNamePos = 7
    WDTValidPos = 3
    CostKeyPos = 1
    CostValPos = 10
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim App As Excel.Application ''''''因为要多次打开文件，所以移到这里'''
    Set App = New Excel.Application
    App.Visible = False  '''''Visible is False by default, so this isn't necessary
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If View = "唯品视角" Then ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Call GetOROpenFile(App, ArrBase, "唯品会十月总货表", valid_ArrBase, ThisWorkbook.Path)
        If Not valid_ArrBase Then
            Path = "\\Wjgx\共享"
            Call GetOROpenFile(App, ArrBase, "唯品会十月总货表", valid_ArrBase, Path)
        End If
        If Not valid_ArrBase Then
            Path = "D:\朱敏\唯品\十月结晶" '''''  '''''此处填写唯品货品总表的路径''''
            Call GetOROpenFile(App, ArrBase, "唯品会十月总货表", valid_ArrBase, Path)
        End If
        
        Call GetOROpenFile(App, arrBase_Vice, "商家仓商品信息(基础信息)", valid_arrBase_Vice, "猫超")
        If Not valid_arrBase_Vice Then
            Call GetOROpenFile(App, arrBase_Vice, "商家仓商品信息(基础信息)", valid_arrBase_Vice, ThisWorkbook.Path)
        End If
        Call GetOROpenFile(App, arrSales, "商品明细_主站_全部人群_全国_全部标签_按日汇总", valid_arrSales)
        Call GetOROpenFile(App, arrSales_Vice, "ADS-", valid_arrSales_Vice, "猫超")
        Call GetOROpenFile(App, arrITL, "常态商品运营", valid_arrITL)
        Call GetandMergeFiles(App, arrStock, "常态可扣减及剩余可售库存", StockKeyPos, StockValPos, valid_arrStock) '''''有两张表，需要合并'''''
    Else '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ITLKeyPos = 9
        ITLValPos = 17
        StockKeyPos = 5
        StockValPos = 17
        SalesKeyPos = 12
        SalesValPos = 15
        Sales_ViceKeyPos = 5
        Sales_ViceValPos = 11
        Call GetOROpenFile(App, ArrBase, "商家仓商品信息(基础信息)", valid_ArrBase)
        If Not valid_ArrBase Then
            Call GetOROpenFile(App, ArrBase, "商家仓商品信息(基础信息)", valid_ArrBase, ThisWorkbook.Path)
        End If
        
        Call GetOROpenFile(App, arrBase_Vice, "唯品会十月总货表", valid_arrBase_Vice, ThisWorkbook.Path)
        If Not valid_ArrBase Then
            Path = "\\Wjgx\共享"
            Call GetOROpenFile(App, ArrBase, "唯品会十月总货表", valid_ArrBase, Path)
        End If
        If Not valid_ArrBase Then
            Path = "D:\朱敏\唯品\十月结晶" '''''  '''''此处填写唯品货品总表的路径''''
            Call GetOROpenFile(App, arrBase_Vice, "唯品会十月总货表", valid_arrBase_Vice, Path)
        End If
        
        Call GetOROpenFile(App, arrSales, "ADS-", valid_arrSales)
        Call GetOROpenFile(App, arrSales_Vice, "商品明细_主站_全部人群_全国_全部标签_按日汇总", valid_arrSales_Vice, "唯品")
        Call GetOROpenFile(App, arrITL, "export-", valid_arrITL)
        Call GetOROpenFile(App, arrStock, "file", valid_arrStock)
        If Not valid_arrStock Then
            StockKeyPos = 2
            StockValPos = 14
            Call GetandMergeFiles(App, arrStock, "商家仓直发一盘货库存数据", 2, 14, valid_arrStock) '''''第2列货品ID，第14列独享库存可售'''
        End If
        Call GetandMergeFiles(App, arrJitStock, "业务库存出入库流水", JitStockKeyPos, JitStockValPos, valid_arrJitStock) '''''实时库存
        Call GetandMergeFiles(App, arrWDTMC, "城东仓可用", WDTKeyPos, WDTValPos, valid_arrWDTMC) ''''''城东仓库存'''''''''''''''
        Call GetandMergeFiles(App, arrWDTBULK, "批发仓可用", WDTKeyPos, WDTValPos, valid_arrWDTBULK) ''''''批发仓库存如果有岭顶仓数据就加上一起'''''''''''''''
    
        If Not valid_arrWDTBULK Then
            Call GetOROpenFile(App, arrWDTBULK, "批发仓可用", valid_arrWDTBULK, ThisWorkbook.Path)
        End If
    End If '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
    Call GetOROpenFile(App, arrUnit, "组合装明细", valid_arrUnit, ThisWorkbook.Path)
    If Not valid_arrUnit Then
        Call GetOROpenFile(App, arrUnit, "组合装明细", valid_arrUnit)
    End If
    Call GetOROpenFile(App, arrAtom, "单品明细", valid_arrAtom, ThisWorkbook.Path)
    If Not valid_arrAtom Then
        Call GetOROpenFile(App, arrAtom, "单品明细", valid_arrAtom)
    End If
    
    Call GetOROpenFile(App, arrWDT, "越中仓可用库存", valid_arrWDT, ThisWorkbook.Path)
    If Not valid_arrWDT Then
        Call GetOROpenFile(App, arrWDT, "越中仓可用", valid_arrWDT)
    End If
    
    Call GetOROpenFile(App, arrCost, "成本", valid_arrCost, ThisWorkbook.Path)
    If Not valid_arrCost Then
        Call GetOROpenFile(App, arrCost, "成本", valid_arrCost)
    End If
    
    ''''''''''''''''''''''''''''''''''''''
    App.Quit ''''重要的步骤''''''
    Set App = Nothing ''''重要的步骤,因为要多次打开文件，所以移到这里'''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If Not valid_ArrBase Or Not valid_arrWDT Or Not valid_arrUnit Then
        Valid = False
        MsgBox "没有找到文件！"
        Exit Function
    Else
        Valid = True
        ReDim Arr(1 To UBound(ArrBase, 1) + 10, 1 To IIf(View = "唯品视角", 28, 35))
    End If
    For i = 2 To UBound(ArrBase, 1) ''''''表格中的#N/A值是直接触发报错'''''
        For j = 1 To Application.Min(78, UBound(ArrBase, 2))
            If IsError(ArrBase(i, j)) Then
                ArrBase(i, j) = ""
            End If
        Next
    Next
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If valid_arrJitStock Then ''''实时库存排序并调整主键列
        Call MgSt_main(subArrJitStock, arrJitStock, JitStockTimePos)
        Call MgSt_main(subArrJitStock, arrJitStock, JitStockKeyPos, True)
        
        '''''''''''''''''''''''''''''临时采用进销存流水作为销售数据
        For i = 2 To UBound(arrJitStock)
            If arrJitStock(i, 15) <> "TOC销售" Then
                arrJitStock(i, 18) = 0
            End If
            arrJitStock(i, 1) = arrJitStock(i, 20)
            arrJitStock(i, 12) = arrJitStock(i, 2)
            arrJitStock(i, 15) = arrJitStock(i, 18) * -1
        Next
        If View = "猫超视角" Then
            Erase arrSales
            arrSales = arrJitStock
        Else
            Erase arrSales_Vice
            arrSales_Vice = arrJitStock
        End If
        '''''''''''''''''''''''''''''临时采用进销存流水作为销售数据
        
        For i = 1 To UBound(subArrJitStock) - 1
            For j = i + 1 To UBound(subArrJitStock)
                If arrJitStock(subArrJitStock(j), JitStockKeyPos) <> arrJitStock(subArrJitStock(i), JitStockKeyPos) Then
                    Exit For
                End If
                arrJitStock(subArrJitStock(j), JitStockKeyPos) = "ABC123"
            Next
            i = j - 1
        Next
        ''''''''''''''''''''''''
        Call MgSt_main(subArrJitStock, arrJitStock, JitStockKeyPos, True)
    End If
    ''''''''''''''''''''''''''''实时库存排序并调整主键列
    
    Dim sales_Date(2) As Date
    If valid_arrSales Then
        DatePos = IIf(View = "唯品视角", 7, 1) '''''唯品销售表时间在第7列''''
        Call MgSt_main(subArrSales, arrSales, DatePos)
        '''''''用于核验导出销售日报的时间段是否合适''''
        sales_Date(0) = CDate(arrSales(subArrSales(1), DatePos))
        sales_Date(1) = CDate(arrSales(subArrSales(UBound(subArrSales)), DatePos))
        If sales_Date(0) - sales_Date(1) < 28 Then
            MsgBox "销售记录少于4周，请知悉！"
        ElseIf sales_Date(0) < Date - 1 Then
            MsgBox "销售记录不是最新的，请知悉！"
        End If
        Call MgSt_main(subArrSales, arrSales, SalesKeyPos, True)
        Call sumSales(subArrSales, arrSales, SalesKeyPos, SalesValPos, DatePos)
    End If
    If valid_arrSales_Vice Then
        DatePos = IIf(View = "唯品视角", 1, 7)  '''''''''猫超销售表日期在第1列''''''''
        Call MgSt_main(subarrSales_Vice, arrSales_Vice, DatePos)
        Call MgSt_main(subarrSales_Vice, arrSales_Vice, Sales_ViceKeyPos, True)
        Call sumSales(subarrSales_Vice, arrSales_Vice, Sales_ViceKeyPos, Sales_ViceValPos, DatePos)
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If valid_arrStock Then
        Call MgSt_main(subArrStock, arrStock, StockKeyPos)
    End If
    
    If valid_arrCost Then
        For j = 1 To UBound(arrCost, 2)
            If Trim(arrCost(1, j)) = "商家编码" Then
                CostKeyPos = j
            End If
            If Trim(arrCost(1, j)) = "会员价" Then
                CostValPos = j
            End If
        Next
        Call MgSt_main(subArrCost, arrCost, CostKeyPos)
    End If
    Cost = Array(subArrCost, arrCost, valid_arrCost, CostKeyPos, CostValPos)
    
    ''''''''''''''''''''''''''''以下初始信息是货品ID的情况''''''''''
    If isHPID Then
        keyColumn = IIf(View = "唯品视角", 1, 2)
        For i = 2 To UBound(ArrBase, 1) ''''''表格中的#N/A值是直接触发报错,如果有就替换成空值'''''
            If IsError(ArrBase(i, keyColumn)) Then
                ArrBase(i, keyColumn) = ""
            End If
        Next
        Call MgSt_main(subArrBase, ArrBase, keyColumn)
    '''''''''''''''''''''''''''''''以上是初始信息是货品ID的情况''''''''''
    '''''''''''''''''''''''''''''''以下是初始信息仅是商品ID的情况''''''''''
    ElseIf isSPID Then
        keyColumn = IIf(View = "唯品视角", 11, 3)
        Dim bl As Boolean
        bl = False
        If isSKUID Then
            bl = True
            Call MgSt_main(subArrBase, ArrBase, IIf(View = "唯品视角", 8, 4))
        End If
        Call MgSt_main(subArrBase, ArrBase, keyColumn, bl)
    ElseIf isAll Then
        '''''''''''''''''构建arrinitial数组'''''
        isHPID = True
        keyColumn = IIf(View = "唯品视角", 1, 2)
        Counter = 0
        For i = 2 To UBound(ArrBase, 1)
            If Trim(CStr(ArrBase(i, 6))) <> "" And Trim(CStr(ArrBase(i, 12))) <> "淘汰" And Trim(CStr(ArrBase(i, 12))) <> "买断" Then '''''And Trim(CStr(ArrBase(i, 12))) <> "淘汰"
                Counter = Counter + 1
            End If
        Next
        ReDim arrInitial(1 To Counter)
        Counter = 1
        For i = 2 To UBound(ArrBase, 1)
            If Trim(CStr(ArrBase(i, 6))) <> "" And Trim(CStr(ArrBase(i, 12))) <> "淘汰" And Trim(CStr(ArrBase(i, 12))) <> "买断" Then ''''And Trim(CStr(ArrBase(i, 12))) <> "淘汰"
                arrInitial(Counter) = ArrBase(i, keyColumn)
                Counter = Counter + 1
            End If
        Next
        '''''''''''''''''''''''''''''''''''''''
        Call MgSt_main(subArrBase, ArrBase, keyColumn)
        
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call MgSt_main(subArrWDT, arrWDT, WDTKeyPos)
    If valid_arrWDTMC Then
        Call MgSt_main(subArrWDTMC, arrWDTMC, WDTKeyPos)
    End If
    If valid_arrWDTBULK Then
        Call MgSt_main(subArrWDTBULK, arrWDTBULK, WDTKeyPos)
    End If
    WDTCategory(0) = Array(subArrWDT, arrWDT, valid_arrWDT)
    WDTCategory(1) = Array(subArrWDTMC, arrWDTMC, valid_arrWDTMC)
    WDTCategory(2) = Array(subArrWDTBULK, arrWDTBULK, valid_arrWDTBULK)
''''''''''''''''arrunit额外的这些位置放置重要信息''''
''''''17平台名：唯品/猫超，18合并单品编码明细，19单品涉及唯品组合数，20单品涉及猫超组合数，21组合涉及唯品折算销量，22组合涉及猫超折算销量，''''
''''''23组合唯品4周实销，24组合猫超4周实销，25组合唯品1周销量，26组合猫超1周销量，27单品唯品实际总销量，28猫超总实销，29单品1周总销量''''''
''''''30单品折算总销量，31单品越中仓实仓库存数,32~38唯品最近7天销量，39~45猫超最近7天销量,46越中仓一个sku的可用库存数，47越中仓sku可分配库存'''''''
''''''48组合预包装数量'''''49库存情况描述，50需求加量'''''51单品唯品折算销量，52单品猫超折算销量''''''53单品城东仓库存数，54单品批发仓库存数'''''''
''''''55猫超一个sku可用库存数''''56猫超一个sku可分配库存数''''''57单品成本，58组合成本''''''''59预包装情况说明''''


    Call merge_UnitAtom(arrUnit, arrAtom, UnitKeyPos, UnitValkeyPos, UnitValvalPos, UnitNamePos, UnitArtNoPos, _
                         AtomKeyPos, AtomValPos, AtomNamePos, AtomArtNoPos, AtomSpecificationPos)   ''''''把ArrAtom合并进arrunit'''
    Call MgSt_main(subArrUnit, arrUnit, UnitKeyPos)
    Call fill_ArrUnit(ArrBase, subArrSales, arrSales, subArrUnit, arrUnit, WDTCategory, SalesKeyPos, SalesValPos, UnitKeyPos, UnitValkeyPos, UnitValvalPos, WDTKeyPos, WDTValPos) ''''''将主销售信息填充进arrunit'''''
    Call fill_ArrUnit(arrBase_Vice, subarrSales_Vice, arrSales_Vice, subArrUnit, arrUnit, WDTCategory, Sales_ViceKeyPos, Sales_ViceValPos, UnitKeyPos, UnitValkeyPos, UnitValvalPos, WDTKeyPos, WDTValPos, True, arrAtom, Cost) ''''''将副销售信息填充进arrunit''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    If valid_arrITL Then ''''''对商品状态表（常态商品运营）进行排序，为读取做准备'''''
        Call MgSt_main(subArrITL, arrITL, ITLKeyPos)
    End If
    
    '''''''''''''''''''''''''开启最外层循环'''''''''''''''''''''''''''''''''''''
    Dim Margin(1 To 2) As Long
    Dim intTemp(1 To 2) As Long
    Dim intPrepack&, kPrepack&
    Pointer = 1
    For i = 1 To UBound(arrInitial) ''''arrinitial已经做了转置成水平方向数组''''
        If isSKUID Then
            strBff = Split(arrInitial(i), "#")(0)
        Else
            strBff = arrInitial(i)
        End If
        
        intBff = binarySearch(strBff, 1, UBound(subArrBase), subArrBase, ArrBase, keyColumn)
        If intBff <> -1 Then '''''''''''''''''大循环中的最外层条件判断''''
            k = subArrBase(intBff)
        If InStr(ArrBase(k, 12), "淘汰") = 0 And InStr(ArrBase(k, 12), "买断") = 0 Then '''''没有逻辑短路，所以嵌套一个同级判断是否为淘汰'''
            If isHPID Then
                Margin(1) = intBff
                Margin(2) = intBff
            ElseIf isSPID Then
                For k = intBff To 1 Step -1
                    If Trim(CStr(ArrBase(subArrBase(k), keyColumn))) <> strBff Then
                        Margin(1) = k + 1
                        Exit For
                    End If
                    Margin(1) = k
                Next
                For k = intBff To UBound(subArrBase)
                    If Trim(CStr(ArrBase(subArrBase(k), keyColumn))) <> strBff Then
                        Margin(2) = k - 1
                        Exit For
                    End If
                    Margin(2) = k
                Next
            End If

            If isSPID And isSKUID Then
                intTemp(1) = Margin(1)
                intTemp(2) = Margin(2)
                For k = intTemp(1) To intTemp(2)
                    If Trim(CStr(ArrBase(subArrBase(k), keyColumn))) & "#" & Trim(CStr(ArrBase(subArrBase(k), 8))) = arrInitial(i) Then
                        Margin(1) = k
                        Exit For
                    End If
                Next
                For k = intTemp(1) To intTemp(2)
                    If Trim(CStr(ArrBase(subArrBase(k), keyColumn))) & "#" & Trim(CStr(ArrBase(subArrBase(k), 8))) <> arrInitial(i) Then
                        Margin(2) = k - 1
                        Exit For
                    End If
                Next
            End If
            
            '''''''''''''''''''''关键的循环，在这里填充表格所需信息'''''''''''''''''''''''''''''''''''
            For k = Margin(1) To Margin(2)
                If View = "唯品视角" Then '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Arr(Pointer, 1) = Pointer '''''Arr(Pointer, 2)列放在最后填入'''''
                    Arr(Pointer, 3) = Trim(CStr(ArrBase(subArrBase(k), 1)))
                    Arr(Pointer, 4) = Trim(CStr(ArrBase(subArrBase(k), 11)))
                    Arr(Pointer, 5) = Trim(CStr(ArrBase(subArrBase(k), 8)))
                    Arr(Pointer, 6) = Trim(CStr(ArrBase(subArrBase(k), 6))) '''''Arr(Pointer, 7) ''''''此列放在最后填入'''''
                    Arr(Pointer, 8) = Trim(CStr(ArrBase(subArrBase(k), 4))) ''''''Arr(Pointer, 9)''''''此列放在最后填入'''''
                    Arr(Pointer, 19) = " "
                    
                    intTemp(1) = binarySearch(Arr(Pointer, 3), 1, UBound(subArrITL), subArrITL, arrITL, ITLKeyPos)
                    If intTemp(1) <> -1 Then
                        intTemp(2) = subArrITL(intTemp(1))
                        Arr(Pointer, 2) = arrITL(intTemp(2), ITLValPos) ''''''填充第2列''''''
                    End If
                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '''''''''''''''''''''''''以下填充旺店通编码明细''''''''''''''''''
                    strBff = Trim(CStr(ArrBase(subArrBase(k), 2)))
                    If strBff = "" Then '''''提高容错能力'''''
                        strBff = Trim(CStr(ArrBase(subArrBase(k), 1)))
                    End If
                    intTemp(1) = binarySearch(strBff, 1, UBound(subArrUnit), subArrUnit, arrUnit, UnitKeyPos)
                    If intTemp(1) <> -1 Then '''''''''''''''查找组合的判断''''''
                        intTemp(2) = subArrUnit(intTemp(1))
                        Arr(Pointer, 7) = arrUnit(intTemp(2), 18) ''''''填充第arr第7列''arrunit第18列放置的是旺店通编码明细''''
                        Arr(Pointer, 9) = IIf(arrUnit(intTemp(2), 21) = "", Empty, Int(arrUnit(intTemp(2), 21))) ''''''填充第arr第9列''arrunit第21列放置的是唯品折算4周销量''''
                        Arr(Pointer, 10) = IIf(arrUnit(intTemp(2), 23) = "", Empty, Int(arrUnit(intTemp(2), 23))) ''''''填充第arr第10列''arrunit第23列放置的是唯品折算4周实销''''
                        Arr(Pointer, 11) = IIf(arrUnit(intTemp(2), 25) = "", Empty, Int(arrUnit(intTemp(2), 25))) ''''''填充第arr第11列''arrunit第25列放置的是唯品1周实销量''''
                        Arr(Pointer, 12) = IIf(arrUnit(intTemp(2), 47) = "", Empty, Int(arrUnit(intTemp(2), 47) / (arrUnit(intTemp(2), 21) - 0.001) * 28))
                        If Arr(Pointer, 12) < 0 Then
                            Arr(Pointer, 12) = Empty
                        End If
                        Arr(Pointer, 16) = IIf(arrUnit(intTemp(2), 47) = "", Empty, Int(arrUnit(intTemp(2), 47)))
                        Arr(Pointer, 17) = IIf(arrUnit(intTemp(2), 46) = "", Empty, Int(arrUnit(intTemp(2), 46)))
                        Arr(Pointer, 18) = arrUnit(intTemp(2), 49) ''''''''库存情况描述''''''
                        Arr(Pointer, 27) = arrUnit(intTemp(2), 58) '''''''''组合成本'''''
                        Arr(Pointer, 28) = arrUnit(intTemp(2), 60) '''''''''有效销售天数'''''
                        If arrUnit(intTemp(2), 48) <> "" Then
                                Arr(Pointer, 16) = Arr(Pointer, 16) + arrUnit(intTemp(2), 48)
                                Arr(Pointer, 17) = Arr(Pointer, 17) + arrUnit(intTemp(2), 48)
                                Arr(Pointer, 18) = "预包装" & arrUnit(intTemp(2), 48) & "套。" & Arr(Pointer, 18)
                        End If
                        intBffStock = binarySearch(Arr(Pointer, 3), 1, UBound(subArrStock), subArrStock, arrStock, StockKeyPos)
                        If intBffStock <> -1 Then
                            kStock = subArrStock(intBffStock)
                            Arr(Pointer, 14) = arrStock(kStock, StockValPos)
                            If Arr(Pointer, 14) = "-" Then
                                Arr(Pointer, 14) = Empty
                            End If
                        End If
                        
                        If isAddition Then
                            Arr(Pointer, 15) = IIf(Int(arrAdd(i) / (Margin(2) - Margin(1) + 1)) = 0, "", Int(arrAdd(i) / (Margin(2) - Margin(1) + 1)))
                            arrUnit(intTemp(2), 50) = arrAdd(i) ''''''''''''''重要操作，把加量需求数据写入arrUnit，然后才能进行计算''''
                        End If
                        ''''''''''''''''填充19~25列''''''''''''''''
                        For j = 20 To 26
                            Arr(Pointer, j) = arrUnit(intTemp(2), j + 12) '''''唯品最近7天销量放在arrUnit的32~38列，猫超最近7天销量放在39~45列'''
                        Next
                        If sales_Date(0) >= Date - 5 Then
                            Arr(Pointer, 13) = Int(Arr(Pointer, 14) * 3 / (Arr(Pointer, 26 - (Date - sales_Date(0) - 1)) + Arr(Pointer, 25 - (Date - sales_Date(0) - 1)) + Arr(Pointer, 24 - (Date - sales_Date(0) - 1)) - 0.001))
                            If Arr(Pointer, 13) < 0 Or Arr(Pointer, 14) = "" Then
                                Arr(Pointer, 13) = Empty
                            End If
                        End If
                        '''''''''''''''''''''''''''''''''''''''''''
                    Else
                        Arr(Pointer, 18) = "没有找到旺店通记录！请确认是否已更新表格！" & Arr(Pointer, 18)
                    End If '''''''''''''''查找组合的判断''''''
                    Pointer = Pointer + 1
                Else '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Arr(Pointer, 1) = Pointer '''''Arr(Pointer, 2)列放在最后填入'''''
                    Arr(Pointer, 3) = Trim(CStr(ArrBase(subArrBase(k), 2)))
                    Arr(Pointer, 4) = Trim(CStr(ArrBase(subArrBase(k), 3)))
                    Arr(Pointer, 5) = Trim(CStr(ArrBase(subArrBase(k), 4)))
                    Arr(Pointer, 6) = Trim(CStr(ArrBase(subArrBase(k), 10)))
                    Arr(Pointer, 7) = Trim(CStr(ArrBase(subArrBase(k), 6))) ''''如果是空则换成使用实时合成的''''
                    Arr(Pointer, 8) = Trim(CStr(ArrBase(subArrBase(k), 9)))
                    Arr(Pointer, 9) = Trim(CStr(ArrBase(subArrBase(k), 7)))
                    Arr(Pointer, 22) = ArrBase(subArrBase(k), 12) '''''备注'''''
                    Arr(Pointer, 34) = ArrBase(subArrBase(k), 13) '''''供货价'''''
                    Arr(Pointer, 35) = ArrBase(subArrBase(k), 14) '''''单件运费'''''
                    
                    intTemp(1) = binarySearch(Arr(Pointer, 3), 1, UBound(subArrITL), subArrITL, arrITL, ITLKeyPos)
                    If intTemp(1) <> -1 Then
                        intTemp(2) = subArrITL(intTemp(1))
                        Arr(Pointer, 2) = arrITL(intTemp(2), ITLValPos) ''''''填充第2列''''''
'                        If Arr(Pointer, 3) = "650750858657" Then Stop ''''tiaoshi
                    End If
                    strBff = Trim(CStr(ArrBase(subArrBase(k), 5)))
                    intTemp(1) = binarySearch(strBff, 1, UBound(subArrUnit), subArrUnit, arrUnit, UnitKeyPos)
                    If intTemp(1) <> -1 Then '''''''''''''''查找组合的判断''''''
                        intTemp(2) = subArrUnit(intTemp(1))
                        If Arr(Pointer, 7) = "" Then
                            Arr(Pointer, 7) = arrUnit(intTemp(2), 18) ''''''填充第arr第7列''arrunit第18列放置的是旺店通编码明细''''
                        End If
                        Arr(Pointer, 10) = IIf(arrUnit(intTemp(2), 22) = "", Empty, Int(arrUnit(intTemp(2), 22))) ''''''填充第arr第10列''arrunit第22列放置的是猫超折算4周实销''''
                        Arr(Pointer, 11) = IIf(arrUnit(intTemp(2), 24) = "", Empty, Int(arrUnit(intTemp(2), 24))) ''''''填充第arr第11列''arrunit第24列放置的是猫超4周实销量''''
                        Arr(Pointer, 12) = IIf(arrUnit(intTemp(2), 26) = "", Empty, Int(arrUnit(intTemp(2), 26)))
                        intBffStock = binarySearch(Arr(Pointer, 3), 1, UBound(subArrStock), subArrStock, arrStock, StockKeyPos)
                        If intBffStock <> -1 Then
                            kStock = subArrStock(intBffStock)
                            Arr(Pointer, 13) = arrStock(kStock, StockValPos)
                        End If
'                        If valid_arrJitStock Then ''''''猫超实时库存
'                            intBffJitStock = binarySearch(Arr(Pointer, 3), 1, UBound(subArrJitStock), subArrJitStock, arrJitStock, JitStockKeyPos)
'                            If intBffJitStock <> -1 Then
'                                kJitStock = subArrJitStock(intBffJitStock)
''                                If arrJitStock(kJitStock, JitStockValPos) * 1 > Arr(Pointer, 13) * 1 Then Stop ''''''''''''''tiaoshi
'                                Arr(Pointer, 13) = arrJitStock(kJitStock, JitStockValPos)
'                            End If
'                        End If ''''''''''''''''''''''''''猫超实时库存
                        
                        Arr(Pointer, 16) = arrUnit(intTemp(2), 56) ''''''猫超sku分配库存'''''
                        If arrUnit(arrUnit(intTemp(2), 55), 53) <> "" Then
                            Arr(Pointer, 17) = Int(arrUnit(arrUnit(intTemp(2), 55), 53) / arrUnit(arrUnit(intTemp(2), 55), UnitValvalPos))
                        End If
                        If arrUnit(arrUnit(intTemp(2), 55), 31) <> "" Then
                            Arr(Pointer, 18) = Int(arrUnit(arrUnit(intTemp(2), 55), 31) / arrUnit(arrUnit(intTemp(2), 55), UnitValvalPos))
                        End If
                        If arrUnit(arrUnit(intTemp(2), 55), 54) <> "" Then
                            Arr(Pointer, 19) = Int(arrUnit(arrUnit(intTemp(2), 55), 54) / arrUnit(arrUnit(intTemp(2), 55), UnitValvalPos))
                        End If
                        If isAddition Then
                            Arr(Pointer, 21) = IIf(Int(arrAdd(i) / (Margin(2) - Margin(1) + 1)) = 0, "", Int(arrAdd(i) / (Margin(2) - Margin(1) + 1)))
                            arrUnit(intTemp(2), 50) = arrAdd(i) ''''''''''''''重要操作，把加量需求数据写入arrUnit，然后才能进行计算''''
                        End If
                        If arrUnit(intTemp(2), 49) <> "" Or Arr(Pointer, 16) < 50 Then
                            Arr(Pointer, 22) = arrUnit(intTemp(2), 49) & arrUnit(arrUnit(intTemp(2), 55), 59) & Arr(Pointer, 22) ''''找出瓶颈单品的可拆预包装信息'''
                        End If
                        Arr(Pointer, 14) = Int(Arr(Pointer, 16) * 28 / (Arr(Pointer, 10) - 0.001))
                        If Arr(Pointer, 14) < 0 Then
                            Arr(Pointer, 14) = Empty
                        End If
                        
                        ''''''''''''''''填充25~31列''''''''''''''''
                        For j = 25 To 31
                            Arr(Pointer, j) = arrUnit(intTemp(2), j + 14) '''''唯品最近7天销量放在arrUnit的32~38列，猫超最近7天销量放在39~45列'''
                        Next
                        '''''''''''''''''''''''''''''''''''''''''''
                        Arr(Pointer, 24) = " " '''''''''空格格挡'''''
                        Arr(Pointer, 32) = arrUnit(intTemp(2), 58) '''''''''组合成本'''''
                        Arr(Pointer, 33) = arrUnit(intTemp(2), 61) '''''''''组合重量'''''
                    Else
                        Arr(Pointer, 22) = "没有找到旺店通记录！请确认是否已更新表格！" & Arr(Pointer, 22)
                    End If '''''''''''''''查找组合的判断''''''
                    Pointer = Pointer + 1
                End If ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Next '''''''''''''''''''''关键的循环，在这里填充表格所需信息''''''''''''''''''''''''''''''''
''''''17平台名：唯品/猫超，18合并单品编码明细，19单品涉及唯品组合数，20单品涉及猫超组合数，21组合涉及唯品折算销量，22组合涉及猫超折算销量，''''
''''''23组合唯品4周实销，24组合猫超4周实销，25组合唯品1周销量，26组合猫超1周销量，27单品唯品实际总销量，28猫超总实销，29单品1周总销量''''''
''''''30单品折算总销量，31单品越中仓实仓库存数,32~38唯品最近7天销量，39~45猫超最近7天销量,46越中仓一个sku的可用库存数，47越中仓sku可分配库存'''''''
''''''48组合预包装数量'''''49库存情况描述，50需求加量'''''51单品唯品折算销量，52单品猫超折算销量''''''53单品城东仓库存数，54单品批发仓库存数'''''''
''''''55猫超一个sku可用库存数''''56猫超一个sku可分配库存数''''''57单品成本，58组合成本''''''''59预包装情况说明''''60有效销售天数''''


    
        End If '''''没有逻辑短路，所以嵌套一个同级判断是否为淘汰'''
        Else
            '''''没有在唯品货品全表中找到对于编码的情况'''''
            Arr(Pointer, 3) = arrInitial(i)
            If isSKUID Then
                Arr(Pointer, 4) = Split(arrInitial(i), "#")(0)
                Arr(Pointer, 5) = Split(arrInitial(i), "#")(1)
                Arr(Pointer, IIf(View = "唯品视角", 18, 22)) = "没有查询到此" & ArrBase(1, keyColumn) & ArrBase(1, IIf(View = "唯品视角", 8, 4)) & "！"
            Else
                Arr(Pointer, IIf(View = "唯品视角", 18, 22)) = "没有查询到此" & ArrBase(1, keyColumn) & "！"
            End If
            If arrInitial(i) = "" And UBound(arrInitial) = Counter - 1 Then
                Arr(Pointer, 3) = "Row" & i + 1 & "/空条码！" '''' & ArrBase(i + 1, 6)
            End If
            Pointer = Pointer + 1
        End If ''''''''''''''''''''''''''''最外层判断'''''''''''''''''
    Next
    '''''''''''''''''''''''''结束最外层循环'''''''''''''''''''''''''''''''''''''
    
    Call fill_Addition(arrAtom, subArrUnit, arrUnit, UnitKeyPos, UnitValkeyPos, UnitValvalPos, isAddition) ''''''调用此函数，如果没有需求加量则不进行任何操作，如果有，则计算单品需求。''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function fill_ArrUnit(ByRef ArrBase(), ByRef subArrSales() As Long, ByRef arrSales(), ByRef subArrUnit() As Long, ByRef arrUnit(), ByRef WDTCategory(), ByVal SalesKeyPos&, ByVal SalesValPos&, ByVal UnitKeyPos&, ByVal UnitValkeyPos&, ByVal UnitValvalPos&, ByVal WDTKeyPos&, ByVal WDTValPos&, _
Optional ByVal FullFilled As Boolean, Optional ByRef arrAtom, Optional ByRef Cost)
    
    Dim i&, j&, k&, M&, n&, p&, intBff, intCache, Counter&
    Dim strPlatform$, strBff As String, strCache$, blBff As Boolean, Top&, Bottom&
    Dim KeyPosForarrSales&, TargetValPos(), OffSet&, subArrTemp() As Long, Coefficient(1 To 200)
    Dim intPrepack&, kPrepack&, subArrWDT() As Long, arrWDT(), valid_arrWDT As Boolean
    Dim subArrWDTMC() As Long, arrWDTMC(), valid_arrWDTMC As Boolean, subArrWDTBULK() As Long, arrWDTBULK(), valid_arrWDTBULK As Boolean
    
    Dim TiaoShi
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     ''''''''''''''''''这里对设置系数，共同占用的时候旺店通数量乘以系数就是可用值''''''''''''''''''''
    Coefficient(1) = 1
    Coefficient(2) = 3 / 4
    Coefficient(3) = 7 / 10
'    Coefficient(4) = 2 / 3
'    Coefficient(5) = 5 / 8
'    Coefficient(6) = 3 / 5
    For i = 4 To 200
        Coefficient(i) = 7 / 10
    Next
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    subArrWDT = WDTCategory(0)(0)
    arrWDT = WDTCategory(0)(1)
    valid_arrWDT = WDTCategory(0)(2)
    subArrWDTMC = WDTCategory(1)(0)
    arrWDTMC = WDTCategory(1)(1)
    valid_arrWDTMC = WDTCategory(1)(2)
    subArrWDTBULK = WDTCategory(2)(0)
    arrWDTBULK = WDTCategory(2)(1)
    valid_arrWDTBULK = WDTCategory(2)(2)
''''''17平台名：唯品/猫超，18合并单品编码明细，19单品涉及唯品组合数，20单品涉及猫超组合数，21组合涉及唯品折算销量，22组合涉及猫超折算销量，''''
''''''23组合唯品4周实销，24组合猫超4周实销，25组合唯品1周销量，26组合猫超1周销量，27单品唯品实际总销量，28猫超总实销，29单品1周总销量''''''
''''''30单品折算总销量，31单品越中仓实仓库存数,32~38唯品最近7天销量，39~45猫超最近7天销量,46越中仓一个sku的可用库存数，47越中仓sku可分配库存'''''''
''''''48组合预包装数量'''''49库存情况描述，50需求加量'''''51单品唯品折算销量，52单品猫超折算销量''''''53单品城东仓库存数，54单品批发仓库存数'''''''
''''''55猫超一个sku可用库存数''''56猫超一个sku可分配库存数''''''57单品成本，58组合成本''''''59预包装情况说明''''''60唯品组合有效销售天数，61猫超组合有效销售天数''''

   
    '''''''''''''''''唯品猫超销售信息填充''''''''''''''''''''''''''''''
    If UBound(ArrBase, 2) > 26 Then
        strPlatform = "VIP"
        OffSet = 3
        KeyPosForarrUnit = 2
        KeyPosForarrSales = 1
        TargetValPos = Array(19, 21, 23, 25, 27, 60)
    Else
        strPlatform = "MC"
        OffSet = 10
        KeyPosForarrUnit = 5
        KeyPosForarrSales = 2
        TargetValPos = Array(20, 22, 24, 26, 28, 61)
    End If
    
    For i = 1 To UBound(ArrBase, 1)
        strBff = Trim(CStr(ArrBase(i, KeyPosForarrUnit)))
        If strPlatform = "VIP" And strBff = "" Then
            strBff = Trim(CStr(ArrBase(i, 1)))
        End If
        If strBff <> "" And ArrBase(i, 12) <> "淘汰" And ArrBase(i, 12) <> "买断" Then
            intBff = binarySearch(strBff, 1, UBound(subArrUnit), subArrUnit, arrUnit, UnitKeyPos)
            intPrepack = binarySearch(strBff & "-1", 1, UBound(subArrWDT), subArrWDT, arrWDT, WDTKeyPos)
            If intBff <> -1 Then ''''''''''''''''''''开始对有效值进行填充'''''''''''''''''
                k = subArrUnit(intBff)
                '''''''''''''''''''''''''填充唯品猫超相关销售信息''''''''''''''''''''''''''''''
                arrUnit(k, 17) = arrUnit(k, 17) & strPlatform '''''''''填充37列放置平台名称''''''''''''
                arrUnit(k, TargetValPos(0)) = 1
                If intPrepack <> -1 Then
                    kPrepack = subArrWDT(intPrepack)
                    If arrWDT(kPrepack, WDTValPos) > 5 Then
                        arrUnit(k, 48) = arrWDT(kPrepack, WDTValPos)
                    End If
                End If
                intCache = binarySearch(ArrBase(i, KeyPosForarrSales), 1, UBound(subArrSales), subArrSales, arrSales, SalesKeyPos)
                If intCache <> -1 Then
                    M = subArrSales(intCache)
                    arrUnit(k, TargetValPos(1)) = arrSales(M, 26) ''''''折算4周销量'''''
                    arrUnit(k, TargetValPos(2)) = arrSales(M, SalesValPos) '''''4周实际销量''''
                    arrUnit(k, TargetValPos(3)) = arrSales(M, 24)
                    arrUnit(k, TargetValPos(5)) = arrSales(M, 36)
                    For j = 29 To 35 ''''''arrsales29~35列放置最近7天的销量,放进对应唯品或猫超的位置''''
                        arrUnit(k, j + OffSet) = arrSales(M, j)
                    Next
                End If
                '''''''''''''''''''''''''''''''''''''''''''''''''''''' ''''''''''''''''''''
    
                ''''''''''''''''''''''''''''''''''''合并旺店通编码明细''18列'''''''''''''''''''''''
                For M = intBff To 1 Step -1
                    If arrUnit(subArrUnit(M), UnitKeyPos) <> strBff Then
                        Top = M + 1
                        Exit For
                    End If
                    Top = M
                Next
                For M = intBff To UBound(subArrUnit)
                    If arrUnit(subArrUnit(M), UnitKeyPos) <> strBff Then
                        Bottom = M - 1
                        Exit For
                    End If
                    Bottom = M
                Next
                n = subArrUnit(Top)
                If Top = Bottom Then ''''''填充第18列，旺店通编码合并明细'''top=bottom 则k=n'''
                    arrUnit(k, 18) = Trim(CStr(arrUnit(n, UnitValkeyPos))) & IIf(arrUnit(n, UnitValvalPos) * 1 = 1, "", "*" & arrUnit(n, UnitValvalPos) * 1)
                Else
                    For M = Top To Bottom  ''''''填充第7列''''''
                        n = subArrUnit(M)
                        arrUnit(k, 18) = arrUnit(k, 18) & "#   " & Trim(CStr(arrUnit(n, UnitValkeyPos))) & IIf(arrUnit(n, UnitValvalPos) * 1 = 1, "", "*" & arrUnit(n, UnitValvalPos) * 1)
                        arrUnit(n, TargetValPos(0)) = 1
                        arrUnit(n, TargetValPos(1)) = arrUnit(k, TargetValPos(1))
                        arrUnit(n, TargetValPos(2)) = arrUnit(k, TargetValPos(2))
                        arrUnit(n, TargetValPos(3)) = arrUnit(k, TargetValPos(3))
                        arrUnit(n, 48) = arrUnit(k, 48)
                    Next
                    arrUnit(k, 18) = Trim(Mid(arrUnit(k, 18), 2)) '''''去除多余的前缀 #''''
                End If
                '''''''''''''''''''''''''''''''完成编码明细合并''''18列''''''''''''''''''''
                    
            End If ''''''''''''''''''''''''''''''''''''''''''''结束填充'''''''''''''''''''''''
        End If
        strBff = Empty
    Next
    
    If FullFilled Then
        Dim intCache19, intCache20, intCache27, intCache28
        Dim intCache29, intCache30, intCache51, intCache52, intCacheWDT, intBffWDT&, kWDT&, intBffP, Weights
        Dim intBffWDTMC&, kWDTMC&, intCacheWDTMC, intBffWDTBULK&, kWDTBULK&, intCacheWDTBULK, intWDTall, WDTallTemp, kWDTallMin
        Dim subArrCost() As Long, arrCost(), valid_arrCost As Boolean, CostKeyPos&, CostValPos&, intCost, kCost, intCacheCost ''''''这些变量用于计算成本''''
        
        ''''''''''''''''''''''''''
        subArrCost = Cost(0)
        arrCost = Cost(1)
        valid_arrCost = Cost(2)
        CostKeyPos = Cost(3)
        CostValPos = Cost(4)
        ''''''''''''''''''''''''''''''
        Call MgSt_main(subArrTemp, arrUnit, UnitValkeyPos) '''''UnitValkeyPos就是组合装明细表中单品商家编码的位置'''
        If View = "唯品视角" Then
            TargetValPos = Array(19, 21, 23, 25, 27, 60)
        Else
            TargetValPos = Array(20, 22, 24, 26, 28, 61)
        End If
        Counter = UBound(subArrTemp)
        i = 1
        n = 1
        intBffP = 0
        intPrepack = 0
        kPrepack = 0
        Dim strPrepack$, strPrepackAll$
        ReDim arrAtom(1 To Counter, 1 To 20)
        Do
            Top = i
            For k = i To Counter
                If arrUnit(subArrTemp(k), UnitValkeyPos) <> arrUnit(subArrTemp(i), UnitValkeyPos) Then
                    Bottom = k - 1
                    Exit For
                End If
                intBffP = intBffP + arrUnit(subArrTemp(k), 19) + arrUnit(subArrTemp(k), 20)
'                If arrUnit(subArrTemp(k), 2) = "shiyue94" And arrUnit(subArrTemp(k), 12) = "6954864704113" Then Stop ''''tiaoshi
            Next
            If k > Counter Then
                Bottom = Counter
            End If
'            If Top = 19564 Then Stop ''''tiaoshi
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If intBffP > 0 Then
                intCache19 = Empty
                intCache20 = Empty
                intCache27 = Empty
                intCache28 = Empty
                intCache29 = Empty
                intCache30 = Empty
                intCache51 = Empty
                intCache52 = Empty
                intBffWDTMC = -1
                intBffWDTBULK = -1
                intBffWDT = binarySearch(arrUnit(subArrTemp(Top), UnitValkeyPos), 1, UBound(subArrWDT), subArrWDT, arrWDT, WDTKeyPos)
                If intBffWDT <> -1 Then
                    kWDT = subArrWDT(intBffWDT)
                    intCacheWDT = arrWDT(kWDT, WDTValPos)
                End If
                If View = "猫超视角" Then
                    If valid_arrWDTMC Then
                        intBffWDTMC = binarySearch(arrUnit(subArrTemp(Top), UnitValkeyPos), 1, UBound(subArrWDTMC), subArrWDTMC, arrWDTMC, WDTKeyPos)
                        If intBffWDTMC <> -1 Then
                            kWDTMC = subArrWDTMC(intBffWDTMC)
                            intCacheWDTMC = arrWDTMC(kWDTMC, WDTValPos)
                        End If
                    End If
                    If valid_arrWDTBULK Then
                        intBffWDTBULK = binarySearch(arrUnit(subArrTemp(Top), UnitValkeyPos), 1, UBound(subArrWDTBULK), subArrWDTBULK, arrWDTBULK, WDTKeyPos)
                        If intBffWDTBULK <> -1 Then
                            kWDTBULK = subArrWDTBULK(intBffWDTBULK)
                            intCacheWDTBULK = arrWDTBULK(kWDTBULK, WDTValPos)
                        End If
                    End If
                End If
                If valid_arrCost Then
                    intCost = binarySearch(arrUnit(subArrTemp(Top), UnitValkeyPos), 1, UBound(subArrCost), subArrCost, arrCost, CostKeyPos)
                    If intCost <> -1 Then
                        kCost = subArrCost(intCost)
                        intCacheCost = arrCost(kCost, CostValPos) * 1
                        Weights = arrCost(kCost, 32) * 1
                    End If
                End If
                
                For M = Top To Bottom
                    p = subArrTemp(M)
                    intBff = arrUnit(p, UnitValvalPos)
                    intCache19 = intCache19 + arrUnit(p, 19)
                    intCache20 = intCache20 + arrUnit(p, 20)
                    If arrUnit(p, 23) <> "" Then
                        intCache27 = intCache27 + arrUnit(p, 23) * intBff
                    End If
                    If arrUnit(p, 24) <> "" Then
                        intCache28 = intCache28 + arrUnit(p, 24) * intBff
                    End If
                    If arrUnit(p, 25) <> "" Or arrUnit(p, 26) <> "" Then
                        intCache29 = intCache29 + arrUnit(p, 25) * intBff + arrUnit(p, 26) * intBff
                    End If
                    If arrUnit(p, 21) <> "" Or arrUnit(p, 22) <> "" Then
                        intCache30 = intCache30 + arrUnit(p, 21) * intBff + arrUnit(p, 22) * intBff
                    End If
                    If arrUnit(p, 21) <> "" Then
                        intCache51 = intCache51 + arrUnit(p, 21) * intBff
                    End If
                    If arrUnit(p, 22) <> "" Then
                        intCache52 = intCache52 + arrUnit(p, 22) * intBff
                    End If
                    
                    If arrUnit(p, 48) <> "" Then
'                        kPrepack = subArrTemp(k)
'                        intPrepack = arrUnit(kPrepack, 48) * arrUnit(kPrepack, UnitValvalPos)
                        intPrepack = intPrepack + arrUnit(p, 48) * intBff
                        strPrepack = strPrepack & "预包组合" & arrUnit(p, UnitKeyPos) & " " & arrUnit(p, 48) & "套。"
                    End If
                    If arrUnit(p, 6) <> "" Then '''''第6列非空，表明是加进来的单品明细，包含了单品的名称规格''''
                        strCache = arrUnit(p, 6)
                    End If
                Next
                arrAtom(n, 5) = strCache '''''单品规格''''
                intCache19 = IIf(intCache19 = "", Empty, Int(intCache19))
                intCache20 = IIf(intCache20 = "", Empty, Int(intCache20))
                intCache27 = IIf(intCache27 = "", Empty, Int(intCache27))
                intCache28 = IIf(intCache28 = "", Empty, Int(intCache28))
                intCache29 = IIf(intCache29 = "", Empty, Int(intCache29))
                intCache30 = IIf(intCache30 = "", Empty, Int(intCache30))
                intCache51 = IIf(intCache51 = "", Empty, Int(intCache51))
                intCache52 = IIf(intCache52 = "", Empty, Int(intCache52))
                For M = Top To Bottom
                    p = subArrTemp(M)
                    If arrUnit(p, 19) = 1 Then
                        arrUnit(p, 19) = intCache19
                    End If
                    If arrUnit(p, 20) = 1 Then
                        arrUnit(p, 20) = intCache20
                    End If
                    arrUnit(p, 27) = intCache27
                    arrUnit(p, 28) = intCache28
                    arrUnit(p, 29) = intCache29
                    arrUnit(p, 30) = intCache30
                    If intBffWDT <> -1 Then
                        arrUnit(p, 31) = intCacheWDT
                    End If
                    If intBffWDTMC <> -1 Then
                        arrUnit(p, 53) = intCacheWDTMC
                    End If
                    If intBffWDTBULK <> -1 Then
                        arrUnit(p, 54) = intCacheWDTBULK
                    End If
                    If intCacheCost <> "" Then
                        arrUnit(p, 57) = intCacheCost
                        arrUnit(p, 60) = Weights
                    End If
                    If arrUnit(p, UnitKeyPos) = arrUnit(p, UnitValkeyPos) Then '''''两个key相同表明是加进来的单品明细，则第2列包含了单品的名称规格''''
                        strCache = arrUnit(p, 2)
                    End If
                    arrUnit(p, 6) = arrAtom(n, 5) ''''把规格写入arrunit对应的单品'''''
                    If arrUnit(p, 48) <> "" And intPrepack - arrUnit(p, 48) * arrUnit(p, UnitValvalPos) > 0 Then
                        strPrepackAll = Split(strPrepack, "预包组合" & arrUnit(p, UnitKeyPos) & " " & arrUnit(p, 48) & "套。")(0) & Split(strPrepack, "预包组合" & arrUnit(p, UnitKeyPos) & " " & arrUnit(p, 48) & "套。")(1)
                        arrUnit(p, 49) = arrUnit(p, 49) & "###" & "可拆预包单品" & arrUnit(p, UnitValkeyPos) & "共" & intPrepack - arrUnit(p, 48) * arrUnit(p, UnitValvalPos) & "件，其中:" & strPrepackAll
                    End If
                    If strPrepack <> "" Then
                        arrUnit(p, 59) = "可拆预包单品共" & intPrepack & "件，其中:" & strPrepack
                    End If
                Next
                '''-----------------------------------------------------'''
                p = subArrTemp(Top)
'                If arrUnit(p, UnitValkeyPos) = "6954864724012" Then Stop ''''''tiaoshi
                arrAtom(n, 1) = n
                arrAtom(n, 2) = IIf(intCache19 * intCache20 > 0, "共用", IIf(intCache19 > 0, "唯品", "猫超"))
                arrAtom(n, 3) = arrUnit(p, UnitValkeyPos)
                arrAtom(n, 4) = strCache
                arrAtom(n, 6) = arrUnit(p, 5)
                If intCache19 > 0 Then
                    arrAtom(n, 7) = intCache19
                End If
                If intCache20 > 0 Then
                    arrAtom(n, 8) = intCache20
                End If
                ''''''''''''''''''''''''''''''''''''
                arrAtom(n, 9) = intCache30
                arrAtom(n, 10) = IIf(intCache27 = "" And intCache28 = "", Empty, intCache27 + intCache28)
                arrAtom(n, 11) = IIf(View = "唯品视角", intCache51, intCache52)
                If InStr(View & "共用", arrAtom(n, 2)) > 0 Then
                    arrAtom(n, 12) = IIf(View = "唯品视角", intCache27, intCache28)
                End If
                arrAtom(n, 13) = intCache29
                arrAtom(n, 16) = arrUnit(p, 31)
                If strPrepack <> "" Then
                    arrAtom(n, 17) = "可拆预包单品共" & intPrepack & "件，其中:" & strPrepack
                End If
'                arrAtom(n, 17) = ""
                n = n + 1 ''''累加统计有效单品数量
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            i = k
            intBff = Empty
            intBffP = Empty
            intCacheWDT = Empty
            intCacheWDTMC = Empty
            intCacheWDTBULK = Empty
            intCacheCost = Empty
            Weights = Empty
            intPrepack = Empty
            strBff = Empty
            strCache = Empty
            strPrepack = Empty
            strPrepackAll = Empty
            If i > Counter Then Exit Do
        Loop While True
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''库存分配计算在以下代码中实现''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim Label&, intBffConvert, intTemp
        Counter = UBound(subArrUnit)
        i = 1
        intBff = 0
        strBff = ""
        Label = 0
        Do
            Top = i
            p = subArrUnit(Top)
            '''''''''''''''''''''''''''''''''''''60%按照比例系数分配，40%按照销售占比分配''''''''''''''''''''''''''''''''''''''''''''''''''
            If arrUnit(p, 19) + arrUnit(p, 20) > 0 Then
                intBffConvert = arrUnit(p, TargetValPos(1)) / (arrUnit(p, 30) - 0.001) * arrUnit(p, 31) * 4 / 10 _
                                + Coefficient(arrUnit(p, 19) + arrUnit(p, 20)) / arrUnit(p, UnitValvalPos) * arrUnit(p, 31) * 6 / 10
                intBff = arrUnit(p, 31) / arrUnit(p, UnitValvalPos)
                If arrUnit(p, UnitValkeyPos) = "Y-BH" Then ''''对蓝冰等赠品特殊处理''''
                    intBffConvert = 1000
                End If
                If arrUnit(p, 31) = "" Then
                    intBff = Empty
                End If
                ''''''''''''''''''''''''''''''
                intWDTall = intBffConvert * 0.85 + (arrUnit(p, 53) * 0.65 + arrUnit(p, 54) * 0.2) / arrUnit(p, UnitValvalPos)
                If arrUnit(p, 31) & arrUnit(p, 53) & arrUnit(p, 54) = "" Then
                    intWDTall = Empty
                End If
                kWDTallMin = p
            End If
            
            For k = i To Counter ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                p = subArrUnit(k)
                If arrUnit(p, UnitKeyPos) <> arrUnit(subArrUnit(Top), UnitKeyPos) Then
                    Bottom = k - 1
                    Exit For
                End If
                Bottom = k
                
                '''------------------------------------'''
                
                If arrUnit(p, 19) + arrUnit(p, 20) > 0 Then
                
                    '''''''''''''''''''''''''''''''''''''60%按照比例系数分配，40%按照销售占比分配''''''''''''''''''''''''''''''''''''''''''''''''''

                    intTemp = arrUnit(p, TargetValPos(1)) / (arrUnit(p, 30) - 0.001) * arrUnit(p, 31) * 4 / 10 _
                              + Coefficient(arrUnit(p, 19) + arrUnit(p, 20)) / arrUnit(p, UnitValvalPos) * arrUnit(p, 31) * 6 / 10
                    If arrUnit(p, UnitValkeyPos) = "Y-BH" Then ''''对蓝冰等赠品特殊处理''''
                        intTemp = 1000
                    End If
                    '''''''''''''''''60%按照比例系数分配，40%按照销售占比分配''''''''''''''''''''''''''''''''''''
                    WDTallTemp = intTemp * 0.85 + (arrUnit(p, 53) * 0.65 + arrUnit(p, 54) * 0.2) / arrUnit(p, UnitValvalPos)
                    '''''''''''''''''猫超可用库存从3个仓中分配按照销售占比分配''''''''''''''''''''''''''''''''''''
                    
                    If intBffConvert > intTemp Then
                        intBff = arrUnit(p, 31) / arrUnit(p, UnitValvalPos)
                        If arrUnit(p, 31) = "" Then
                            intBff = Empty
                        End If
                        intBffConvert = intTemp
                    End If
                    If intWDTall > WDTallTemp Then
                        kWDTallMin = p
                        intWDTall = WDTallTemp
                        If arrUnit(p, 31) & arrUnit(p, 53) & arrUnit(p, 54) = "" Then
                            intWDTall = Empty
                        End If
                    End If
                    If View = "猫超视角" Then
                        If WDTallTemp < 50 And intWDTall <> "" Then
                            strBff = strBff & arrUnit(p, 7) & arrUnit(p, 6) & "剩余" & arrUnit(p, 31) + arrUnit(p, 53) + arrUnit(p, 54) & "！"
                        End If
                    ElseIf intTemp < 50 And arrUnit(p, 31) <> "" Then
                        strBff = strBff & arrUnit(p, 7) & arrUnit(p, 6) & "剩余" & arrUnit(p, 31) & "！" '''''新arrunit表第7列放置的是单品名称''''
                    End If
                    If arrUnit(p, 17) <> "" Then
                        Label = p
                    End If
                    If valid_arrCost Then
                        intCacheCost = intCacheCost + arrUnit(p, 57) * arrUnit(p, UnitValvalPos) '''''''''计算组合成本''''
                        Weights = Weights + arrUnit(p, 60) * arrUnit(p, UnitValvalPos) '''''''''计算重量''''
                    End If
                End If
                '''------------------------------------'''
            Next '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            For M = Top To Bottom
                p = subArrUnit(M)
                If InStr(arrUnit(p, 49), "可拆") > 0 Then
                    strPrepackAll = strPrepackAll & Split(arrUnit(p, 49), "###")(1)
                End If
'                If p = 3461 Then Stop ''''tiaoshi
            Next
            
            If Label > 0 Then
                arrUnit(Label, 46) = IIf(intBff = "", Empty, Int(intBff)) ''''一个sku的可用库存数，由其成员中的最小值决定'''''
                arrUnit(Label, 47) = IIf(intBff = "", Empty, Int(intBffConvert)) ''''一个sku的可用库存数，由其成员中的最小值决定'''''
                arrUnit(Label, 55) = kWDTallMin
                arrUnit(Label, 56) = IIf(intWDTall = "", Empty, Int(intWDTall))
                arrUnit(Label, 58) = intCacheCost
                arrUnit(Label, 61) = Weights
                If Bottom > Top Then
                    arrUnit(Label, 49) = strBff  '''''描述库存较少的情况''''
                End If
                arrUnit(Label, 49) = arrUnit(Label, 49) & strPrepackAll
            End If
            
            i = k
            intBff = Empty
            Label = Empty
            strBff = Empty
            strPrepack = Empty
            strPrepackAll = Empty
            intBffConvert = Empty
            WDTallTemp = Empty
            intWDTall = Empty
            intCacheCost = Empty
            Weights = Empty
            kWDTallMin = Empty
            If i > Counter Then Exit Do
        Loop While True
        '''''''''''''''''''''''''''以上这段do while循环实现库存分配的计算''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        For i = 1 To UBound(arrAtom) '''''填充arratom需要计算的列'''''
            If arrAtom(i, 9) > 0 Then
                arrAtom(i, 14) = Int(arrAtom(i, 16) / arrAtom(i, 9) * 28)
            End If
            If arrAtom(i, 3) = "" Then Exit For
        Next
        
    End If
''''''17平台名：唯品/猫超，18合并单品编码明细，19单品涉及唯品组合数，20单品涉及猫超组合数，21组合涉及唯品折算销量，22组合涉及猫超折算销量，''''
''''''23组合唯品4周实销，24组合猫超4周实销，25组合唯品1周销量，26组合猫超1周销量，27单品唯品实际总销量，28猫超总实销，29单品1周总销量''''''
''''''30单品折算总销量，31单品越中仓实仓库存数,32~38唯品最近7天销量，39~45猫超最近7天销量,46越中仓一个sku的可用库存数，47越中仓sku可分配库存'''''''
''''''48组合预包装数量'''''49库存情况描述，50需求加量'''''51单品唯品折算销量，52单品猫超折算销量''''''53单品城东仓库存数，54单品批发仓库存数'''''''
''''''55猫超一个sku可用库存数''''56猫超一个sku可分配库存数''''''57单品成本，58组合成本'''''59预包装情况说明''''''60唯品组合有效销售天数，61猫超组合有效销售天数''''

End Function



'''''''''''''''''''合并组合明细、单品明细表''''''''''''''''''''''
Function merge_UnitAtom(ByRef arrUnit(), ByRef arrAtom(), ByRef UnitKeyPos&, ByRef UnitValkeyPos&, ByRef UnitValvalPos&, ByRef UnitNamePos&, _
ByRef UnitArtNoPos&, ByVal AtomKeyPos&, ByVal AtomValPos&, ByVal AtomNamePos&, ByVal AtomArtNoPos&, ByVal AtomSpecificationPos&)
    Dim i&, j&, k&, M&, Counter&
    Dim ArrTemp(), arrPos()
    
    Counter = UBound(arrUnit, 1) + UBound(arrAtom, 1) - 1
    k = UBound(arrUnit, 2)
    
    ReDim ArrTemp(1 To Counter, 1 To 61)
    
    arrPos = Array(UnitKeyPos, UnitNamePos, UnitValkeyPos, UnitValvalPos, UnitArtNoPos)
    For i = 1 To UBound(arrUnit, 1)
        For j = 0 To UBound(arrPos)
            ArrTemp(i, j + 1) = arrUnit(i, arrPos(j))
        Next
        ArrTemp(i, j + 2) = arrUnit(i, 14) ''''新表第7列放置单品名称'''
    Next
    AtomValPos = 1
    M = 2
    arrPos = Array(AtomKeyPos, AtomNamePos, AtomKeyPos, AtomValPos, AtomArtNoPos, AtomSpecificationPos)
    For i = UBound(arrUnit, 1) + 1 To Counter
        For j = 0 To UBound(arrPos)
            ArrTemp(i, j + 1) = arrAtom(M, arrPos(j))
        Next
         ArrTemp(i, 4) = 1
        M = M + 1
    Next
    
    UnitKeyPos = 1
    UnitNamePos = 2
    UnitValkeyPos = 3
    UnitValvalPos = 4
    UnitArtNoPos = 5
    
    Erase arrUnit
    arrUnit = ArrTemp
    Erase ArrTemp
    Erase arrAtom

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function fill_Addition(ByRef arrAtom(), ByRef subArrUnit() As Long, ByRef arrUnit(), ByVal UnitKeyPos&, ByVal UnitValkeyPos&, ByVal UnitValvalPos&, ByVal isAddition)
    
    Dim i&, j&, k&, M&, n&, Counter&, intBffP
    Dim intBff, Top&, Bottom&, subArrTemp() As Long
    
    If Not isAddition Then Exit Function '''''''''''''''看是否需要计算''''
    
    i = 1
    Counter = UBound(subArrUnit)
    Do
        Top = i
        For k = i To Counter
            If arrUnit(subArrUnit(k), UnitKeyPos) <> arrUnit(subArrUnit(i), UnitKeyPos) Then
                Bottom = k - 1
                Exit For
            End If
            If arrUnit(subArrUnit(k), 50) <> "" Then
                intBff = arrUnit(subArrUnit(k), 50)
            End If
            Bottom = k
        Next
        If intBff > 0 Then
            For M = Top To Bottom
                arrUnit(subArrUnit(M), 50) = intBff
            Next
        End If
        i = k
        intBff = 0
        If i > Counter Then Exit Do
    Loop While True
    
    Call MgSt_main(subArrTemp, arrUnit, UnitValkeyPos)
    
    i = 1
    n = 1
    Do
        Top = i
        For k = i To Counter
            If arrUnit(subArrTemp(k), UnitValkeyPos) <> arrUnit(subArrTemp(i), UnitValkeyPos) Then
                Bottom = k - 1
                Exit For
            End If
            intBffP = intBffP + arrUnit(subArrTemp(k), 19) + arrUnit(subArrTemp(k), 20)
            If arrUnit(subArrTemp(k), 50) > 0 Then
                intBff = intBff + arrUnit(subArrTemp(k), 50) * arrUnit(subArrTemp(k), UnitValvalPos)
            End If
            Bottom = k
        Next
        
        
        If intBffP > 0 Then
            If intBff > 0 Then
                arrAtom(n, 15) = intBff
            End If
            If arrUnit(subArrTemp(Bottom), UnitValkeyPos) <> arrAtom(n, 3) Then MsgBox "Alert! 遇到bug！" ''''Alert!''''
            n = n + 1
        End If
        i = k
        intBff = 0
        strBff = ""
        intBffP = 0
        If i > Counter Then Exit Do
    Loop While True

End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''通用销售统计函数''''''''''''''''''''''''''''''
Function sumSales(ByRef subArrSales() As Long, ByRef arrSales(), ByVal KeyPos&, ByVal ValPos, ByVal DatePos&)
    Dim i&, j&, Counter&, intBff&, floatBff!, LaunchDays
    Dim strBff$, subArrTemp() As Long, ArrTemp()
    
    Counter = 1
    For i = 2 To UBound(subArrSales)
        If arrSales(subArrSales(i), KeyPos) <> arrSales(subArrSales(i - 1), KeyPos) Then
            Counter = Counter + 1
        End If
    Next
    ReDim subArrTemp(1 To Counter)
    ReDim ArrTemp(1 To Counter + 1, 1 To 36)
    
    strBff = ""
    Counter = 1
    For i = 1 To UBound(subArrSales)
        floatBff = Application.Max(arrSales(subArrSales(i), ValPos) * 1, 0) ''''唯品销售表第valpos列是销售数量,暂存在floatbff中''''
        If i > 1 Then
            strBff = arrSales(subArrSales(i - 1), KeyPos)
        End If
        If arrSales(subArrSales(i), KeyPos) <> strBff Then
            Counter = Counter + 1
            subArrTemp(Counter - 1) = Counter
            For j = 1 To UBound(arrSales, 2)
                ArrTemp(Counter, j) = arrSales(subArrSales(i), j)
            Next
            ArrTemp(Counter, ValPos) = 0
            ''''''''''''''''''''''''''''''''''''''''''''
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 28 Then
                ArrTemp(Counter, 36) = 1 '''''第36列放置有效销售天数''''
                LaunchDays = arrSales(subArrSales(i), DatePos)
                ArrTemp(Counter, ValPos) = floatBff ''''唯品销售表第valpos列是销售数量,暂存在floatbff中''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 21 Then
                ArrTemp(Counter, 22) = floatBff  ''''新表第22列放置的是21天销量''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 14 Then
                ArrTemp(Counter, 23) = floatBff ''''新表第23列放置的是14天销量''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 7 Then
                ArrTemp(Counter, 24) = floatBff ''''新表第24列放置的是7天销量''''
                intBff = Date - CDate(arrSales(subArrSales(i), DatePos)) '''''将单日销售数量放置在对应列，29-35列''''
                If intBff = -1 Then intBff = 0
                ArrTemp(Counter, 36 - intBff) = floatBff
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 4 Then
                ArrTemp(Counter, 25) = floatBff ''''新表第25列放置的是4天销量''''
            End If
            ''''''''''''''''''''''''''''''''''''''''''''
        Else
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 28 Then
                If arrSales(subArrSales(i), DatePos) <> LaunchDays Then
                    ArrTemp(Counter, 36) = Date - CDate(arrSales(subArrSales(i), DatePos))
                    LaunchDays = arrSales(subArrSales(i), DatePos)
                End If
                ArrTemp(Counter, ValPos) = ArrTemp(Counter, ValPos) + floatBff ''''新表第valpos列放置的是28天销量''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 21 Then
                ArrTemp(Counter, 22) = ArrTemp(Counter, 22) + floatBff ''''新表第22列放置的是21天销量''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 14 Then
                ArrTemp(Counter, 23) = ArrTemp(Counter, 23) + floatBff ''''新表第23列放置的是14天销量''''
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 7 Then
                ArrTemp(Counter, 24) = ArrTemp(Counter, 24) + floatBff ''''新表第24列放置的是7天销量''''
                intBff = Date - CDate(arrSales(subArrSales(i), DatePos)) '''''将单日销售数量放置在对应列，29-35列''''
                If intBff = -1 Then intBff = 0
                ArrTemp(Counter, 36 - intBff) = ArrTemp(Counter, 36 - intBff) + floatBff
            End If
            If CDate(arrSales(subArrSales(i), DatePos)) >= Date - 4 Then
                ArrTemp(Counter, 25) = ArrTemp(Counter, 25) + floatBff ''''新表第25列放置的是4天销量''''
            End If
        End If
    Next
    For i = 2 To UBound(ArrTemp, 1) ''''''第26列放置折算4周销量
        If ArrTemp(i, 25) > ArrTemp(i, ValPos) / 2 Then
            ArrTemp(i, 26) = Int(ArrTemp(i, 25) * 7 * 0.65 + ArrTemp(i, ValPos) * 28 / ArrTemp(i, 36) * 0.35)
        ElseIf ArrTemp(i, 24) > ArrTemp(i, ValPos) * 2 / 3 Then
            ArrTemp(i, 26) = Int(ArrTemp(i, 24) * 4 * 0.65 + ArrTemp(i, ValPos) * 28 / ArrTemp(i, 36) * 0.35)
        ElseIf ArrTemp(i, 23) > ArrTemp(i, ValPos) * 4 / 5 Then
            ArrTemp(i, 26) = Int(ArrTemp(i, 23) * 2 * 0.65 + ArrTemp(i, ValPos) * 28 / ArrTemp(i, 36) * 0.35)
        ElseIf ArrTemp(i, 22) > ArrTemp(i, ValPos) * 7 / 8 Then
            ArrTemp(i, 26) = Int(ArrTemp(i, 22) * 4 / 3 * 0.65 + ArrTemp(i, ValPos) * 28 / ArrTemp(i, 36) * 0.35)
        Else
            ArrTemp(i, 26) = ArrTemp(i, ValPos)
        End If
    Next
    
    Erase subArrSales
    Erase arrSales
    subArrSales = subArrTemp
    arrSales = ArrTemp
    
    Erase subArrTemp
    Erase ArrTemp

End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetandMergeFiles(ByRef App As Excel.Application, ByRef Arr(), ByVal familyName$, ByVal KeyPos&, ByVal ValPos&, ByRef Valid As Boolean)

    Dim i&, j&, k&, M&, Counter&, intBff&, NOCAT&
    Dim strBff As String, strCache As String
    Dim arrSales_a(), arrSales_b(), arrSales_c(), arrSales_d(), arrSales_e(), arrSales_f(), arrSales_g(), arrSales_h(), arrSales_i(), arrCache()
    Dim intBff_a&, intBff_b&, intBff_c&, intBff_d&, intBff_e&, intBff_f&, intBff_g&, intBff_h&, intBff_i&
    Dim arrBff(), subArrBff() As Long
    Dim isValid As Boolean, intBffzz&, Kzz&
    Dim Position&, Directory$
    Dim arrSales_List(), arrSales(), intBff_List(20) As Long
    
''''''''''''''''''''''''''''以下是读取数据'''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Directory = Dir(ThisWorkbook.Path & "\", vbDirectory)
    Do
        
        If GetAttr(ThisWorkbook.Path & "\" & Directory) = vbDirectory And Directory <> "." And Directory <> ".." Then
            If InStr(Directory, Mid(View, 1, 2)) > 0 Then
                Exit Do
            End If
        End If
        Directory = Dir
    Loop While Directory <> "" ''''重要！空值的时候继续循环就会报错''''
    
    strBff = Dir(ThisWorkbook.Path & "\" & Directory & "\" & "*.*")
    Do
        If strBff Like "*" & familyName & "*.*" Then '''''And InStr(1, strCache, strBff, vbTextCompare) = 0 Then  ''''查找销售日报文件''''
            strCache = strCache & ";" & strBff
        End If
        strBff = Dir
    Loop While strBff <> ""
    
    ''''''''''''''''''读取导出数据文件''''
'    If familyName = "批发仓可用库存" Then Stop '''tiaoshi
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    arrSales_List = Array(arrSales_a, arrSales_b, arrSales_c, arrSales_d, arrSales_e, arrSales_f, arrSales_g, arrSales_h, arrSales_i)
    If strCache = "" Then
        Valid = False
        Exit Function
    End If
    intBff = UBound(Split(strCache, ";"))
    If intBff > 0 Then
        Call GetOROpenFile(App, arrSales, Trim(Split(strCache, ";")(1)), isValid)
        Counter = Counter + UBound(arrSales, 1)
        intBff_List(0) = 1
        arrSales_List(0) = arrSales
        Erase arrSales
    End If

    If intBff > 1 Then
        arrBff = arrSales_List(0)
        Call GetOROpenFile(App, arrSales, Trim(Split(strCache, ";")(2)), isValid)
        Counter = Counter + UBound(arrSales, 1)
        intBff_List(1) = intBff_List(0) + UBound(arrBff, 1)
        arrSales_List(1) = arrSales
        Erase arrSales
        Erase arrBff
    End If
    
    If intBff > 2 Then
        For i = 2 To intBff - 1
            arrBff = arrSales_List(i - 1)
            Call GetOROpenFile(App, arrSales, Trim(Split(strCache, ";")(i + 1)), isValid)
            Counter = Counter + UBound(arrSales, 1)
            intBff_List(i) = intBff_List(i - 1) + UBound(arrBff, 1) - 1
            arrSales_List(i) = arrSales
            'If i = 4 Then Stop
            Erase arrSales
            Erase arrBff
        Next
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    NOCAT = intBff
    Counter = Counter - NOCAT + 1
    arrSales_a = arrSales_List(0)
    k = UBound(arrSales_a, 2)
    ReDim arrCache(1 To Counter, 1 To k)
    ccc = 0
    If NOCAT > 0 Then
        For i = 1 To UBound(arrSales_a, 1)
            If Trim(arrSales_a(1, 3)) & Trim(arrSales_a(i, 3)) <> "残次品是" Then
                For j = 1 To k
                    arrCache(i, j) = arrSales_a(i, j)
                Next
            End If
        Next
        Call MgSt_main(subArrBff, arrCache, KeyPos)
    End If
   
    If NOCAT > 1 Then
        intBff = 2
        arrBff = arrCache
        arrSales_b = arrSales_List(1)
        For i = intBff_List(1) To intBff_List(1) + UBound(arrSales_b, 1) - 2
            intBffzz = binarySearch(arrSales_b(intBff, KeyPos), 1, UBound(subArrBff), subArrBff, arrBff, KeyPos)
            If intBffzz <> -1 Then
                Kzz = subArrBff(intBffzz)
                If arrCache(Kzz, ValPos) = "-" Then
                    arrCache(Kzz, ValPos) = arrSales_b(intBff, ValPos)
                ElseIf arrSales_b(intBff, ValPos) <> "-" And Trim(arrSales_b(1, 3)) & Trim(arrSales_b(intBff, 3)) <> "残次品是" Then
                    arrCache(Kzz, ValPos) = arrCache(Kzz, ValPos) * 1 + arrSales_b(intBff, ValPos) * 1
                End If
            End If
            
            If intBffzz = -1 And Trim(arrSales_b(1, 3)) & Trim(arrBff(intBff, 3)) <> "残次品是" Then
                For j = 1 To k
                    arrCache(i, j) = arrSales_b(intBff, j)
                Next
            End If
'                If familyName = "批发仓可用库存" Then Stop '''tiaoshi
            intBff = intBff + 1
        Next
        intBffzz = -1
        Erase arrBff
        Call MgSt_main(subArrBff, arrCache, KeyPos)
    End If

    If NOCAT > 2 Then
        For M = 2 To NOCAT - 1
            intBff = 2
            arrBff = arrCache
            arrSales = arrSales_List(M)
            For i = intBff_List(M) To intBff_List(M) + UBound(arrSales, 1) - 2
                intBffzz = binarySearch(arrSales(intBff, KeyPos), 1, UBound(subArrBff), subArrBff, arrBff, KeyPos)
                If intBffzz <> -1 Then
                    Kzz = subArrBff(intBffzz)
                    If arrCache(Kzz, ValPos) = "-" Then
                        arrCache(Kzz, ValPos) = arrSales(intBff, ValPos)
                    ElseIf arrSales(intBff, ValPos) <> "-" And Trim(arrSales(1, 3)) & Trim(arrSales(intBff, 3)) <> "残次品是" Then
                        arrCache(Kzz, ValPos) = arrCache(Kzz, ValPos) * 1 + arrSales(intBff, ValPos) * 1
                    End If
                End If
                
                If intBffzz = -1 And Trim(arrSales(1, 3)) & Trim(arrSales(intBff, 3)) <> "残次品是" Then
                    For j = 1 To k
                        arrCache(i, j) = arrSales(intBff, j)
                    Next
                End If
                
                intBff = intBff + 1
            Next
            intBffzz = -1
            Erase arrBff
            Erase arrSales
            Call MgSt_main(subArrBff, arrCache, KeyPos)
        Next
    End If
    
   'ThisWorkbook.Worksheets("Cache").Range(ThisWorkbook.Worksheets("Cache").Cells(1, 1), ThisWorkbook.Worksheets("Cache").Cells(Counter, k)) = arrCache
    Arr = arrCache
    Valid = True
    Erase subArrBff
    Erase arrCache
    Erase arrSales_a
    Erase arrSales_b
    Erase arrSales_c
    Erase arrSales_d
    Erase arrSales_e
    Erase arrSales_f
    Erase arrSales_g
    Erase arrSales_h
    Erase arrSales_i

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetOROpenFile(ByRef App As Excel.Application, ByRef Arr(), ByVal file_Name, ByRef Valid As Boolean, Optional Path$ = "") '''''''根据文件名读取表，减少判断''''
    Dim wbBff As Workbook, Wsh As Worksheet
    Dim strBff$, Mark As Boolean
    
    Mark = False
    For Each wbBff In Workbooks '''读取表
       If wbBff.Name Like "*" & file_Name & "*" Then
          strBff = wbBff.Name
          Set Wsh = Workbooks(strBff).Worksheets(1)
          If Wsh.AutoFilterMode = True Then
            Wsh.Rows("1:1").AutoFilter
          End If
          Arr = Wsh.UsedRange.Value
'          wbBff.Close SaveChanges:=False
          Set wbBff = Nothing
          Set Wsh = Nothing
          Mark = True
          Call simpleTableSort(Arr, strBff)
          Exit For
       End If
    Next
    
    If Mark = False Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        strBff = View '''''''稍微复杂的逻辑，为了兼容新功能避免过多修改代码.主要是为了虚拟切换视角打开对方的销售记录表
        If Len(Path) = 2 Then '''''此时path的值是“唯品”或者“猫超”，当path是完路径时，长度会大于2，不会执行这一段
            strBff = Path
            Path = ""
        End If
        Directory = Dir(ThisWorkbook.Path & "\", vbDirectory)
        Do
            If GetAttr(ThisWorkbook.Path & "\" & Directory) = vbDirectory And Directory <> "." And Directory <> ".." Then
                If InStr(Directory, Mid(strBff, 1, 2)) > 0 Then
                    Exit Do
                End If
            End If
            Directory = Dir
        Loop While Directory <> "" ''''重要！空值的时候继续循环就会报错''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          If Path = "" Then
              Path = ThisWorkbook.Path & "\" & Directory
          End If
          strBff = Dir(Path & "\" & "*.*")
          Do
             If strBff Like "*" & file_Name & "*" Then
'                Set App = New Excel.Application
'               App.Visible = False  '''''Visible is False by default, so this isn't necessary '''移到入口函数里'''

               Set wbBff = App.Workbooks.Open(Path & "\" & strBff, ReadOnly = True)
               Set Wsh = wbBff.Worksheets(1)
               If Wsh.AutoFilterMode = True Then
                  Wsh.Rows("1:1").AutoFilter
               End If
               Arr = Wsh.UsedRange.Value
               wbBff.Close SaveChanges:=False ''''重要的步骤
'               App.Quit ''''重要的步骤'''移到入口函数里'''
'               Set App = Nothing ''''重要的步骤,应为要多次打开文件，所以移到入口函数里'''
               Set wbBff = Nothing
               Set Wsh = Nothing
               Mark = True
               strBff = Path & "\" & strBff
               Call simpleTableSort(Arr, strBff)
               Exit Do
             End If
             strBff = Dir
             
          Loop While strBff <> ""
    End If
    Valid = Mark
    
End Function

'''''''''''''''''''''''''''''''''''''''''''


Function simpleTableSort(ByRef Arr(), wbName As String) '''''这个函数暂时用不到'''''
    Dim Counter&, ArrTemp(), i&, j&, Top&, Bottom& ''''这些都是为了处理一盘货库存设置的变量，以后可能不需要''''
    Dim subArr() As Long, KeyPos&, Position&, keyName As String '''''用于排序、查询'''

    
    If wbName Like "*" & "可用库存" & "*" Then  '''旺店通导出可用库存的是否是残次品进行排序''''
        KeyPos = 1
        Position = 3 ''''对第三列进行排序，方便更快筛选出商家仓相关内容''''''
        keyName = "是"
        If InStr(Arr(1, 3), "残次") = 0 Then
            MsgBox "旺店通可用库存表的残次标识不在第3列！"
        End If
    ElseIf wbName Like "*" & "唯品会十月总货表" & "*" Then
        KeyPos = 1
        Position = 12
        keyName = "淘汰"
    Else
        Exit Function
    End If
    If Arr(UBound(Arr, 1), 3) = "NA" Then
        Arr(UBound(Arr, 1), 3) = ""
    End If
    Call MgSt_main(subArr, Arr, Position)
    

    Counter = binarySearch(keyName, 1, UBound(subArr), subArr, Arr, Position)
    
    If Counter = -1 Then Exit Function
    
    For i = Counter To 1 Step -1
        If Trim(Arr(subArr(i), Position)) <> keyName Then ''''注意subArr是不包含表头的，全部是排序信息''''
            Top = i + 1
            Exit For
        End If
        Top = i
    Next
    
    For i = Counter To UBound(subArr)
        If Arr(subArr(i), Position) <> keyName Then
            Bottom = i - 1
            Exit For
        End If
        Bottom = i
    Next
    
    For i = Top To Bottom
        Arr(subArr(i), KeyPos) = Empty
    Next
'    ReDim ArrTemp(1 To UBound(Arr, 1) - (Bottom - Top + 1), 1 To UBound(Arr, 2))
'    For j = 1 To UBound(ArrTemp, 2)
'       ArrTemp(1, j) = Arr(1, j)
'    Next
'    If Top > 1 Then
'        For i = 1 To Top - 1
'            For j = 1 To UBound(ArrTemp, 2)
'                ArrTemp(i + 1, j) = Arr(subArr(i), j)
'            Next
'        Next
'    End If
'    For i = Bottom + 1 To UBound(subArr)
'        For j = 1 To UBound(ArrTemp, 2)
'            ArrTemp(i - (Bottom - Top + 1) + 1, j) = Arr(subArr(i), j)
'        Next
'    Next
'
'    Erase subArr
'    Erase Arr
'
'    Arr = ArrTemp
'    Erase ArrTemp

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function TableFormat(ByRef Wsh As Worksheet)

    Dim tableHeader(4)
    tableHeader(0) = Array("序号", "在架状态", "货品编码", "商品编码", "SKU编码", "商品名称", "旺店通编码明细", "颜色尺码", "货号", "折算4周销量", "4周实际销量", "最近1周实销", "独享库存余量", "可用库存周转", "计划锁定", "分配库存", "城东仓库存", "越中仓库存", "岭顶仓批发仓库存", "计划释放", "需求加量", "备注", "目标周转(天)", "补货类型")
    tableHeader(1) = Array("序号", "在架状态", "唯品条码", "唯品款号", "唯品货号", "商品名称", "旺店通编码明细", "分类", "折算4周销量", "4周实际销量", "最近1周实销", "可用库存周转", "页面库存周转", "页面库存余量", "需求加量", "分配库存", "越中仓库存", "备注", "补货类型")
    tableHeader(2) = Array("序号", "在售平台", "旺店通编码明细", "商品名称", "规格", "货号", "唯品SKU数", "猫超SKU数", "折算4周总销量", "实际4周总销量", "唯品折算4周销量", "唯品4周实销", "最近1周总实销", "剩余库存周转", "库存加量建议", "越中仓库存", "备注", "目标周转(天)")
    tableHeader(3) = Array("序号", "在售平台", "旺店通编码明细", "商品名称", "规格", "货号", "唯品SKU数", "猫超SKU数", "折算4周总销量", "实际4周总销量", "猫超折算4周销量", "猫超4周实销", "最近1周总实销", "剩余库存周转", "库存加量建议", "越中仓库存", "备注", "目标周转(天)")
    
    Application.ScreenUpdating = False
        
    With Wsh '''''''''''''''设置格式''''''''''''''''''''
        .Cells.Clear
        With .Range("A:AI").Font
            .Name = "微软雅黑"
            .Size = 11
            .ColorIndex = xlAutomatic
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        If View = "猫超视角" Then
            .Range("A1:X1") = tableHeader(0)
            .Range("A:AI").VerticalAlignment = xlCenter
            .Range("A:AI").HorizontalAlignment = xlCenter
            .Range("A:AI").Interior.Pattern = xlNone
            .Range("Y:AE").NumberFormatLocal = "0_ "
            .Range("C:I").NumberFormatLocal = "@"
            .Range("A1:AI1").NumberFormatLocal = "@"
            .Range("G:I").WrapText = True
            .Range("F:G,V:V").HorizontalAlignment = xlLeft
            With .Rows(1)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .WrapText = True
            End With
            .Columns(4).Font.Bold = True
            .Columns(5).Font.Bold = True
            .Columns(13).Font.Bold = True
            .Columns(15).Font.Bold = True
            .Columns(15).Font.Color = vbRed
            .Columns(16).Font.Bold = True
            .Columns(20).Font.Bold = True
            .Columns(20).Font.Color = vbRed
            .Columns(21).Font.Bold = True
            .Columns(1).ColumnWidth = 5
            .Columns(2).ColumnWidth = 5
            .Columns(3).ColumnWidth = 15
            .Columns(4).ColumnWidth = 15
            .Columns(5).ColumnWidth = 15
            .Columns(6).ColumnWidth = 45
            .Columns(7).ColumnWidth = 17
            .Columns(8).ColumnWidth = 9
            .Columns(9).ColumnWidth = 8
            .Columns(10).ColumnWidth = 6
            .Columns(11).ColumnWidth = 6
            .Columns(12).ColumnWidth = 6
            .Columns(13).ColumnWidth = 6
            .Columns(14).ColumnWidth = 6
            .Columns(15).ColumnWidth = 6
            .Columns(16).ColumnWidth = 6
            .Columns(17).ColumnWidth = 6
            .Columns(18).ColumnWidth = 6
            .Columns(19).ColumnWidth = 6
            .Columns(20).ColumnWidth = 6
            .Columns(21).ColumnWidth = 6
            .Columns(22).ColumnWidth = 11
            .Columns(23).ColumnWidth = 6
            .Columns(24).ColumnWidth = 6
            .Columns(25).ColumnWidth = 6
            .Columns(26).ColumnWidth = 6
            .Columns(27).ColumnWidth = 6
            .Columns(28).ColumnWidth = 6
            .Columns(29).ColumnWidth = 6
            .Columns(30).ColumnWidth = 6
            .Columns(31).ColumnWidth = 6
            .Columns(32).ColumnWidth = 6
            .Columns(33).ColumnWidth = 6
            .Columns(34).ColumnWidth = 6
            RUN.Left = .Range("G1").Left - 50
            
        Else ''''''''''''''''''''''''''唯品视角'''''''''''''''''''
            .Range("A1:S1") = tableHeader(1)
            .Range("A:AA").VerticalAlignment = xlCenter
            .Range("A:AA").HorizontalAlignment = xlCenter
            .Range("A:AA").Interior.Pattern = xlNone
            .Range("C:H").NumberFormatLocal = "@"
            .Range("G:H").WrapText = True
'            .Range("S1:Z1").NumberFormatLocal = "mm/dd"
            .Range("A1:Z1").NumberFormatLocal = "@"
            .Range("F:G,R:R").HorizontalAlignment = xlLeft
            .Range("H:I").WrapText = False
            .Range("B:B").HorizontalAlignment = xlRight
            With .Rows(1)
                .Font.Bold = True
                .HorizontalAlignment = xlCenter
                .WrapText = True
            End With
            .Columns(3).Font.Bold = True
'            .Columns(5).Font.Bold = True
            .Columns(10).Font.Bold = True
'            .Columns(10).Font.Color = vbRed
            .Columns(12).Font.Bold = True
            .Columns(14).Font.Bold = True
            .Columns(16).Font.Bold = True
'            .Columns(16).Font.Color = vbRed
            .Columns(1).ColumnWidth = 5
            .Columns(2).ColumnWidth = 6
            .Columns(3).ColumnWidth = 15
            .Columns(4).ColumnWidth = 13
            .Columns(5).ColumnWidth = 13
            .Columns(6).ColumnWidth = 40
            .Columns(7).ColumnWidth = 17
'            .Columns(8).ColumnWidth = 16
'            .Columns(9).ColumnWidth = 11
            .Columns(8).ColumnWidth = 6
            .Columns(9).ColumnWidth = 6
            .Columns(10).ColumnWidth = 6
            .Columns(11).ColumnWidth = 6
            .Columns(12).ColumnWidth = 6
            .Columns(13).ColumnWidth = 6
            .Columns(14).ColumnWidth = 6
            .Columns(15).ColumnWidth = 6
            .Columns(16).ColumnWidth = 6
            .Columns(17).ColumnWidth = 6
            .Columns(18).ColumnWidth = 12
            .Columns(19).ColumnWidth = 6
            .Columns(20).ColumnWidth = 6
            .Columns(21).ColumnWidth = 6
            .Columns(22).ColumnWidth = 6
            .Columns(23).ColumnWidth = 6
            .Columns(24).ColumnWidth = 6
            .Columns(25).ColumnWidth = 6
            .Columns(26).ColumnWidth = 6
            .Columns(27).ColumnWidth = 6
            RUN.Left = .Range("G1").Left - 50
        End If
        
        Application.ScreenUpdating = True
'''''''''''''''''冻结、筛选''''''''''''''''''''''''
        ActiveWindow.FreezePanes = False
        .Range("A2").Select
        If View = "唯品视角" Then
            .Range("T2").Select
        End If
        ActiveWindow.FreezePanes = True
        .Rows("1:1").AutoFilter
        If .AutoFilterMode = False Then
            .Rows("1:1").AutoFilter
        End If
    End With
''''''''''''''''完成操作界面格式设置'''''''''''''''''''''''
    With ThisWorkbook.Worksheets("越中仓单品明细")
    
        .Cells.Clear
        If View = "唯品视角" Then
            .Range("A1:Q1") = tableHeader(2)
        Else
            .Range("A1:Q1") = tableHeader(3)
        End If
        .Rows(1).Font.Bold = True
        With .Range("A:Q").Font
            .Name = "微软雅黑"
            .Size = 11
            .ColorIndex = xlAutomatic
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        .Range("A:Q").VerticalAlignment = xlCenter
        .Range("A:Q").HorizontalAlignment = xlCenter
        .Range("Q:Q").HorizontalAlignment = xlLeft
        .Range("A:Q").Interior.Pattern = xlNone
        .Range("C:F,N:N").NumberFormatLocal = "@"
        .Range("A:Q").WrapText = False
        .Rows(1).WrapText = True
        .Rows(1).Font.Bold = True
        .Rows(1).HorizontalAlignment = xlCenter
        .Columns(9).Font.Bold = True
        .Columns(14).Font.Bold = True
        .Columns(15).Font.Bold = True
        .Columns(16).Font.Bold = True
'        .Columns(12).Font.Color = vbRed
        .Columns(1).ColumnWidth = 6
        .Columns(2).ColumnWidth = 6
        .Columns(3).ColumnWidth = 18
        .Columns(4).ColumnWidth = 28
        .Columns(5).ColumnWidth = 12
        .Columns(6).ColumnWidth = 12
        .Columns(6).WrapText = True
        .Columns(7).ColumnWidth = 8
        .Columns(8).ColumnWidth = 8
        .Columns(9).ColumnWidth = 8
        .Columns(8).ColumnWidth = 8
        .Columns(9).ColumnWidth = 8
        .Columns(10).ColumnWidth = 8
        .Columns(11).ColumnWidth = 8
        .Columns(12).ColumnWidth = 8
        .Columns(13).ColumnWidth = 8
        .Columns(14).ColumnWidth = 8
        .Columns(15).ColumnWidth = 8
        .Columns(16).ColumnWidth = 8
        .Columns(17).ColumnWidth = 20
        .Columns(18).ColumnWidth = 8
        If .AutoFilterMode = False Then
            .Rows("1:1").AutoFilter
        End If
    End With

    Erase tableHeader

End Function


''''''''''''''''''''''''''''''''''''''''''''''经典二分查找''''''''''''''''''''''''''''''''''''''''''''
Public Function binarySearch(ByRef X, ByVal iBegin&, ByVal iEnd&, ByRef subArr() As Long, ByRef wbArr(), Optional ByRef Pos As Long = 1)
''''''pos形参是指定查询列的位置'''''''''
    Dim iMiddle&, intBff&
    
    iMiddle = Int((iBegin + iEnd) / 2)
    If iBegin > iEnd Or X = "" Then
      binarySearch = -1
      Exit Function
    End If
    If wbArr(subArr(iMiddle), Pos) = "" Then ''''必须单独处理空值情况'''
        intBff = binarySearch(X, iBegin, iMiddle - 1, subArr, wbArr, Pos)
    Else
        Select Case Trim(CStr(X))
        Case Is = Trim(CStr(wbArr(subArr(iMiddle), Pos)))
            intBff = iMiddle '''subArr(iMiddle)'''这里做了调整，返回的是排序后的位置'''
        Case Is > Trim(CStr(wbArr(subArr(iMiddle), Pos)))
            intBff = binarySearch(X, iBegin, iMiddle - 1, subArr, wbArr, Pos)
        Case Is < Trim(CStr(wbArr(subArr(iMiddle), Pos)))
            intBff = binarySearch(X, iMiddle + 1, iEnd, subArr, wbArr, Pos)
        End Select
    End If

    ''''''''''排序数组是降序查找'''''
    
    binarySearch = intBff

End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''经过优化的归并排序'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function MergeSort(ByVal iBegin&, ByVal iEnd&, ByRef subArr() As Long, ByRef tempArr() As Long, ByRef wbArr(), Optional ByVal Pos& = 1)
    Dim i&, j&, e&, k&
    Dim subI&, subJ&
    Dim iMiddle
    Dim forNone As Boolean
    iMiddle = Int((iBegin + iEnd) / 2)
    i = iBegin
    j = iMiddle + 1
    e = iEnd
    k = iBegin
    If iBegin < iEnd Then '''经典归并排序'''
        ''''''''''''''''''''递归''''''''''''''''''''''''''''''''''''''''
        If iEnd - iBegin >= 20 Then '''''当排序数量少的时候使用插入排序会效率更高
            Call MergeSort(i, j - 1, subArr, tempArr, wbArr, Pos)
            Call MergeSort(j, e, subArr, tempArr, wbArr, Pos)
        Else
            Call insertSt(i, j - 1, subArr, wbArr, Pos) ''''''''当排序数量少的时候使用插入排序会效率更高
            Call insertSt(j, e, subArr, wbArr, Pos)
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Do While i <= iMiddle And j <= iEnd
            If wbArr(subArr(j), Pos) = "" Then
                forNone = True
            ElseIf wbArr(subArr(i), Pos) = "" Then
                forNone = False
            Else
                forNone = (Trim(CStr(wbArr(subArr(i), Pos))) >= Trim(CStr(wbArr(subArr(j), Pos))))
            End If
            If forNone Then ''''降序
                tempArr(k) = subArr(i)
                i = i + 1
                k = k + 1
            Else
                tempArr(k) = subArr(j)
                j = j + 1
                k = k + 1
            End If
        Loop
        ''''''''''''''''''''''''''''
        If i <= iMiddle Then
            For i = i To iMiddle
                tempArr(k) = subArr(i)
                k = k + 1
            Next
        End If
        If j <= iEnd Then
            For j = j To iEnd
                tempArr(k) = subArr(j)
                k = k + 1
            Next
        End If
        For i = iBegin To iEnd
            subArr(i) = tempArr(i)
        Next
    
    End If

End Function

Function insertSt(iBegin&, iEnd&, ByRef subArr() As Long, ByRef wbArr(), Pos&)
    Dim i&, j&, temp&
    Dim forNone As Boolean
    If iEnd - iBegin < 1 Then Exit Function
    For i = iBegin + 1 To iEnd
       temp = subArr(i)
       j = i - 1
       '''Do While j >= iBegin And wbArr(subArr(j), pos) > wbArr(Temp, pos) ''''vba语言的弱点，and运算不支持逻辑短路特性导致的很容易出现下标越界的情况'''在VB.NET中使用andalso解决，orelse解决
       Do While j >= iBegin
          If wbArr(temp, Pos) = "" Then
                forNone = True
          ElseIf wbArr(subArr(j), Pos) = "" Then
                forNone = False
          Else
                forNone = (Trim(CStr(wbArr(subArr(j), Pos))) >= Trim(CStr(wbArr(temp, Pos))))
          End If
          If forNone Then Exit Do '''''降序
          subArr(j + 1) = subArr(j)
          j = j - 1
       Loop
       subArr(j + 1) = temp
    Next


End Function

Sub MgSt_main(ByRef subArr() As Long, ByRef wbArr(), Optional ByVal Pos& = 1, Optional isInitialized As Boolean = False) ''''''merge sort
    Dim i&, j&, Counter&
    Dim tempArr() As Long
    Dim iBegin&, iEnd&, k&, temp
    Counter = UBound(wbArr, 1)
    
    If Not isInitialized Then
        ReDim subArr(1 To Counter - 1)
        For i = 2 To Counter
            subArr(i - 1) = i
        Next
    End If
    
    Counter = UBound(subArr)
    ReDim tempArr(1 To Counter)
    
    iBegin = 1
    iEnd = Counter
    Call MergeSort(iBegin, iEnd, subArr, tempArr, wbArr, Pos)
End Sub



Private Sub changeVision_Click()
    If changeVision.Caption = "唯品视角" Then
        changeVision.Caption = "猫超视角"
        Range("A2").Select
    Else
        changeVision.Caption = "唯品视角"
        Range("T2").Select
    End If
    View = changeVision.Caption
    Call TableFormat(ThisWorkbook.Worksheets("操作界面"))

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub RUN_Click()
    Call VIPandMCSJC_TMJ_Supply
End Sub

