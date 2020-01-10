Attribute VB_Name = "模块1"
Sub compute_all()
    Dim wk As Worksheet
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    For i = 2 To ThisWorkbook.Worksheets.Count
       
        'MsgBox ThisWorkbook.Worksheets(i).Name
        Call gongshi(ThisWorkbook.Worksheets(i))
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    
End Sub
Sub printkaoqin()
    Dim wk As Worksheet
    Dim i As Integer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    For i = 2 To ThisWorkbook.Worksheets.Count
       
        'MsgBox ThisWorkbook.Worksheets(i).Name
        Call gongshi(ThisWorkbook.Worksheets(i))
    Next
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Sub
Sub createPrinttb(sheetname As String, num As Integer)
    Dim dy As String, dyst As Worksheet, i As Integer, activest As String
    
    activest = ActiveSheet.Name
    On Error Resume Next
    dy = sheetname
    Set dyst = ActiveWorkbook.Sheets(dy)
    
    If dyst Is Nothing Then
        ActiveWorkbook.Sheets.Add(after:=Sheets(num)).Name = dy
        Set dyst = ActiveWorkbook.Sheets(dy)
    End If
    
    
            dyst.Columns(1).ColumnWidth = 4
            dyst.Columns(2).ColumnWidth = 3
            dyst.Columns(3).ColumnWidth = 2
            dyst.Columns(4).ColumnWidth = 0.85
            dyst.Columns(5).ColumnWidth = 3
            dyst.Columns(6).ColumnWidth = 1.5
            dyst.Columns(7).ColumnWidth = 1.5
            dyst.Columns(8).ColumnWidth = 3
            dyst.Columns(9).ColumnWidth = 2.5
            dyst.Columns(10).ColumnWidth = 1.15
            dyst.Columns(11).ColumnWidth = 2
            dyst.Columns(12).ColumnWidth = 1.5
            dyst.Columns(13).ColumnWidth = 0.85
            dyst.Columns(14).ColumnWidth = 3
            dyst.Columns(15).ColumnWidth = 6
            
            dyst.Columns(16).ColumnWidth = 1
            
            dyst.Columns(17).ColumnWidth = 4
            dyst.Columns(18).ColumnWidth = 3
            dyst.Columns(19).ColumnWidth = 2
            dyst.Columns(20).ColumnWidth = 0.85
            dyst.Columns(21).ColumnWidth = 3
            dyst.Columns(22).ColumnWidth = 1.5
            dyst.Columns(23).ColumnWidth = 1.5
            dyst.Columns(24).ColumnWidth = 3
            dyst.Columns(25).ColumnWidth = 2.5
            dyst.Columns(26).ColumnWidth = 1.15
            dyst.Columns(27).ColumnWidth = 2
            dyst.Columns(28).ColumnWidth = 1.5
            dyst.Columns(29).ColumnWidth = 0.85
            dyst.Columns(30).ColumnWidth = 3
            dyst.Columns(31).ColumnWidth = 6
            
            With ActiveSheet.PageSetup
'                '按自定义纸张打印
'                '注意：需先在打印设置中自定义一个命名为“SHD”的页面尺寸（长21cm*宽14.7cm）
                .PaperSize = xlPaperA4       '设置纸张的大小为自定义的“SHD”。若为xlPaperA4则为A4纸
'
                .Orientation = xlPortrait        '该属性返回或设置页面的方向。wpsOrientPortrait 纵向；wpsOrientLandscape 横向
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.748031496062992)
                .BottomMargin = Application.InchesToPoints(0.748031496062992)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
'                .LeftMargin = Application.InchesToPoints(1.5)
'                   .RightMargin = Application.InchesToPoints(1.5)
'                   .TopMargin = Application.InchesToPoints(1.5)
'                   .BottomMargin = Application.InchesToPoints(1.5)
'                   .HeaderMargin = Application.InchesToPoints(1)
'                   .FooterMargin = Application.InchesToPoints(1)
                   .PrintGridlines = True
                   .CenterHorizontally = True        '页面的水平居中
                .CenterVertically = True        '页面垂直居中
'                '.Zoom = False        '将页面缩印在一页内
                .Zoom = 100
'                '.FitToPagesWide = 1
'
'                'If Range("A1") <> "" Then       ‘设置触发找印条件
'                   '.PrintArea = ""    '取消打印区域
'                   '.PrintArea = "$A$3:$G$18"
'                   'Range("A3:G18").PrintOut Copies:=1, Collate:=True    '打印指定区域，直接打印
'                   'Range("A3:G18").PrintOut Copies:=1, Preview:=True, Collate:=True   '打印预览。
'                'End If
'                   '上面代码即[a3:G18].PrintOut
            End With

        ActiveWorkbook.Worksheets(activest).Activate
End Sub
Function isexistssheet(bookname As String) As Boolean
    Dim buer As Boolean, dyst As Worksheet
    
    buer = False
    'Dim dy As String, dyst As Worksheet, i As Integer
    
    On Error Resume Next
    
    Set dyst = ActiveWorkbook.Sheets(bookname)
    
    If dyst Is Nothing Then
        'ActiveWorkbook.Sheets.Add(after:=Sheets(ActiveWorkbook.Worksheets.Count)).Name = dy
        'Set dyst = ActiveWorkbook.Sheets(dy)
        buer = True
    End If
    isexistssheet = buer
End Function
Sub gongshi(ws As Worksheet)
    Dim i As Integer, j As Integer, k As Integer, rq As String, xm As String
    Dim arr() As Variant
    Dim brr() As Variant    ', crr As Variant
    Dim dy As Worksheet, xmrng As Range
    Dim maxrow As Integer, curcol As Integer, nextrow As Integer, nextcol As Integer
    
    Set dicnormalday = CreateObject("scripting.dictionary")
    Set reg = CreateObject("vbscript.regexp")
    
       
    reg.Pattern = "\d+(,\d+)*"
    reg.Global = True
    
    If reg.test(ws.Name) Then
       ' MsgBox ws.Name
        ws.Activate
        
        'xm = Range("j4").Value
        rq = Range("b5").Value
        k = Val(Mid(rq, 6, 2))
        For j = 1 To Day(DateSerial(Val(Mid(rq, 1, 4)), k + 1, 1) - 1)
            If Weekday(CDate(Left(rq, 8) & j), 2) <> 7 And Weekday(CDate(Left(rq, 8) & j), 2) <> 6 Then
                dicnormalday("numOf") = dicnormalday("numOf") + 1
            End If
        Next
        
        'MsgBox (Cells(4, 15).MergeCells)
        For i = 15 To 45 Step 15
                    If Application.CountA(Range(Cells(13, i - 13), Cells(43, i - 1))) > 0 Then
                              
                                'MsgBox i
                                xm = Cells(4, i - 5).Value
                                
                                If Range(Cells(4, i), Cells(43, i)).MergeCells = True Then
                                    Range(Cells(4, i), Cells(43, i)).UnMerge
                                    
                                End If
                                Range(Cells(4, i), Cells(43, i)).NumberFormatLocal = "G/通用格式"
                                Range(Cells(4, i), Cells(43, i)).HorizontalAlignment = xlCenter
                                Range(Cells(4, i), Cells(43, i)).Font.Bold = True
                                
                                ' 设置列宽
                                Columns(i).ColumnWidth = 3
                                
                                ' 合并标题
                                If Range(numToString(i - 14) & "10").MergeArea.Count = 14 Then
                                    Range(Cells(10, i - 14), Cells(10, i - 1)).UnMerge
                                    Range(Cells(10, i - 14), Cells(10, i)).Merge
                                End If
                                
                                If Range(numToString(i - 4) & "11").MergeArea.Count = 4 Then
                                    Range(Cells(11, i - 4), Cells(11, i - 1)).UnMerge
                                    Range(Cells(11, i - 4), Cells(11, i)).Merge
                                End If
                                
                                If Cells(12, i).Value = "" Then
                                    Cells(12, i).Value = "加班工时"
                                End If
                                
                                            '格式化操作
                                    For j = 44 To 52
                                    
                                        Range(Cells(j, i - 14), Cells(j, i - 12)).Merge
                                        Range(Cells(j, i - 11), Cells(j, i - 7)).Merge
                                        Range(Cells(j, i - 6), Cells(j, i)).Merge
                                        If j > 45 Then
                                            Range(Cells(j - 1, i - 6), Cells(j, i)).Merge
                                        'Else
                                            
                                        End If
                                        
                                    Next
                                
                                '计算工时
                                    arr = Range(Cells(13, i - 14), Cells(43, i)).Value
                                    For j = 1 To UBound(arr)
'                                            If j = 11 Then
'                                                MsgBox j
'                                            End If
                                            'reg.Pattern = "六|日"
                                            reg.Pattern = "六"
                                            If arr(j, 2) <> "" And arr(j, 4) <> "" And arr(j, 7) <> "" And arr(j, 9) <> "" And reg.test(arr(j, 1)) <> True Then
                                                dicnormalday("normalday") = dicnormalday("normalday") + 1
                                           ElseIf arr(j, 2) <> "" And arr(j, 4) <> "" And reg.test(arr(j, 1)) <> True And arr(j, 7) = "" And arr(j, 9) = "" Then
                                                dicnormalday("normalday") = dicnormalday("normalday") + 0.5
                                           ElseIf arr(j, 2) <> "" And arr(j, 4) <> "" And reg.test(arr(j, 1)) <> True And arr(j, 7) <> "" And arr(j, 13) <> "" And arr(j, 11) = "" Then
                                                dicnormalday("normalday") = dicnormalday("normalday") + 1
                                            End If
                                            
                                            If arr(j, 11) <> "" Then
                                            
                                                reg.Pattern = "\+$"
                                                If reg.test(arr(j, 13)) Then
                                                    arr(j, 13) = reg.Replace(arr(j, 13), "")
                                                End If
                                                If arr(j, 13) <> "" And Hour(CDate(arr(j, 11))) = 18 Then
                                                    '判断是否为凌晨
                                                            If Hour(CDate(arr(j, 13))) > 18 And Hour(CDate(arr(j, 13))) <= 23 Then
                                                                    arr(j, 15) = Hour(CDate(arr(j, 13))) - Hour(CDate(arr(j, 11))) - 0.5
                                                                    If Minute(CDate(arr(j, 13))) >= 30 And Minute(CDate(arr(j, 13))) < 58 Then
                                                                        arr(j, 15) = arr(j, 15) + 0.5
                                                                    ElseIf Minute(CDate(arr(j, 13))) >= 59 Then
                                                                            arr(j, 15) = arr(j, 15) + 1
                                                                     'elseif Minute(CDate(arr(j, 13))) < 30
                                                                    End If
                                                            ElseIf Hour(CDate(arr(j, 13))) >= 0 And Hour(CDate(arr(j, 13))) < 5 Then
                                                                    arr(j, 15) = 24 - Hour(CDate(arr(j, 11))) - 0.5 + Hour(CDate(arr(j, 13)))
                                                                    If Minute(CDate(arr(j, 13))) >= 30 And Minute(CDate(arr(j, 13))) < 58 Then
                                                                        arr(j, 15) = arr(j, 15) + 0.5
                                                                    ElseIf Minute(CDate(arr(j, 13))) >= 59 Then
                                                                        arr(j, 15) = arr(j, 15) + 1
                                                                     'elseif Minute(CDate(arr(j, 13))) < 30
                                                                    End If
                                                            End If
                                                    
                                                Else   '加班时间为空
                                                    arr(j, 15) = 0
                                                End If
                                            ElseIf arr(j, 9) <> "" And arr(j, 13) <> "" Then
                                                arr(j, 15) = Hour(CDate(arr(j, 13))) - Hour(CDate(arr(j, 9))) - 0.5
                                                If Minute(CDate(arr(j, 11))) >= 30 And Minute(CDate(arr(j, 11))) <= 58 Then
                                                    arr(j, 15) = arr(j, 15) + 0.5
                                                ElseIf Minute(CDate(arr(j, 13))) >= 59 Then
                                                    arr(j, 15) = arr(j, 15) + 1
                                                End If
                                            ElseIf arr(j, 7) <> "" And arr(j, 13) <> "" Then
                                                arr(j, 15) = Hour(CDate(arr(j, 13))) - 17 - 0.5
                                                If Minute(CDate(arr(j, 13))) >= 30 And Minute(CDate(arr(j, 13))) <= 58 Then
                                                    'arr(j, 15) = Hour(CDate(arr(j, 13))) - Hour(CDate(arr(j, 9)))
                                                    arr(j, 15) = arr(j, 15) + 0.5
                                                ElseIf Minute(CDate(arr(j, 13))) >= 59 Then
                                                    arr(j, 15) = arr(j, 15) + 1
                                                End If
                                            Else
                                                arr(j, 15) = ""
                                            End If
                                        Next
                                     
                                     'range(
                                     Range(Cells(13, i - 14), Cells(43, i)) = arr
                                     
                                     ReDim brr(1 To 9, 1 To 15)
                                     For k = 1 To 9
                                        If k = 1 Then brr(k, 1) = "正班天数：": brr(k, 4) = dicnormalday("normalday"): brr(k, 9) = "请核对无误后签名确认"
                                        If k = 2 Then brr(k, 1) = "加班小时：": brr(k, 4) = "=sum(" & numToString(i) & "13:" & numToString(i) & "43)"
                                        If k = 3 Then brr(k, 1) = "正班工资：": brr(k, 4) = "=" & numToString(i - 11) & 44 & "*" & 100
                                        If k = 4 Then brr(k, 1) = "加班工资：": brr(k, 4) = "=" & numToString(i - 11) & 45 & "*" & 18.75
                                        If k = 5 Then
                                            brr(k, 1) = "全勤奖："
                                            If dicnormalday("normalday") = dicnormalday("numOf") Then
                                                brr(k, 4) = 100
                                            End If
                                        End If
                                        If k = 6 Then brr(k, 1) = "其他奖金："
                                        If k = 7 Then brr(k, 1) = "扣项(水电费):"
                                        If k = 8 Then brr(k, 1) = "其他扣款："
                                        If k = 9 Then brr(k, 1) = "合计：": brr(k, 4) = "=sum(" & numToString(i - 11) & 46 & "," & numToString(i - 11) & 47 & "," & numToString(i - 11) & 48 & "," & numToString(i - 11) & 49 & _
                                                                                                            ",-" & numToString(i - 11) & 50 & ",-" & numToString(i - 11) & 51 & ")"
                                        'if
                                    Next
                                    
                                    Range(Cells(44, i - 14), Cells(52, i)) = brr
                                    'set dicNormalDay
                                    dicnormalday("normalday") = 0
                                    
                                    Range(Cells(10, i - 14), Cells(52, i)).Select
                                    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                                    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                                        With Selection.Borders(xlEdgeLeft)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Borders(xlEdgeTop)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Borders(xlEdgeBottom)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Borders(xlEdgeRight)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Borders(xlInsideVertical)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Borders(xlInsideHorizontal)
                                            .LineStyle = xlContinuous
                                            .Color = -196608
                                            .TintAndShade = 0
                                            .Weight = xlThin
                                        End With
                                        With Selection.Font
                                           .Name = "宋体"
                                           .Size = "9"
'                                          .Color = -196608
                                        End With
                                    ' Range(Cells(13, i - 13), Cells(43, i - 1)).Select
                                    With Range(Cells(13, i - 13), Cells(43, i - 1))
                                        .NumberFormatLocal = "hh:mm"
                                        '.HorizontalAlignment = xlCenter
                                    End With
                                    Range(Cells(44, i - 11), Cells(52, i - 11)).Select
                                    With Selection
                                        .NumberFormatLocal = "0.0_ "
                                        .HorizontalAlignment = xlCenter
                                    End With
                                    With Range(numToString(i - 6) & 44)
                                        .HorizontalAlignment = xlCenter
                                    End With
'
                                'if activeworkbook.s
                                If isexistssheet("打印表") Then
                                    Call createPrinttb("打印表", ActiveWorkbook.Worksheets.Count)
                                End If
                                Set dy = ActiveWorkbook.Worksheets("打印表")
                                
                                If dy.Range("A1").CurrentRegion.Rows.Count = 1 Then
                                    nextrow = 1
                                    nextcol = 1
                                    maxrow = 49
                                    
                                Else
                                
                                    'crr = dy.UsedRange.Value
                                    'maxrow = UBound(crr)
                                    'maxcol = UBound(crr, 2)
                                    With dy.UsedRange
                                        Set xmrng = .Find(xm, LookIn:=xlValues)
                                        If Not xmrng Is Nothing Then
                                            maxrow = xmrng.Row + 48
                                            curcol = xmrng.Column + 5
                                            
                                            'firstAddress = xmrng.Address
'                                            Do
'                                                c.Value = 5
'                                                Set c = .FindNext(c)
'                                            Loop While Not c Is Nothing
                                        
                                        Else
                                            maxrow = .Rows.Count
                                            curcol = dy.Cells(maxrow - 40, Columns.Count).End(xlToLeft).Column
                                        End If
                                    End With
                                    If curcol = 15 Then
                                        nextrow = maxrow - 48
                                        nextcol = 17
                                    ElseIf curcol = 31 Then
                                        nextrow = ((maxrow - 49) / 54 + 1) * 54 + 1
                                        nextcol = 1
                                    End If
                                            
                                            'nextrow = (maxcol - 49) / 54 + 1
'                                            If curcol = 15 Then
'                                                nextcol = 17
'                                            ElseIf curcol = 31 Then
'                                                nextcol = 1
'                                            End If
                                        
                                    
                                 End If
                                    'maxrows = dy.Cells(Rows.Count, 1).End(xlUp).Row
                                    'maxcol = dy.cells(columns.Count
                                 Range(Cells(4, i - 14), Cells(8, i - 1)).Copy dy.Range(numToString(nextcol) & nextrow)
                                 Range(Cells(10, i - 14), Cells(52, i)).Copy dy.Range(numToString(nextcol) & (nextrow + 6))
                                 
                                 'MsgBox maxrow
                                 With dy.PageSetup
                                    .PrintArea = "$A$1:$AE$" & (nextrow + 48)
                                 End With
                                 
                                
                                
                        End If
            Next
           
        
    End If
    Set dicnormalday = Nothing
    Set reg = Nothing
End Sub
Function numToString(nm As Integer) As String
    numToString = Replace(Cells(1, nm).Address(False, False), "1", "")
End Function
Sub test_gs()
    Call gongshi(ActiveSheet)
End Sub
Sub copykaoqin(ws As Worksheet)
    Dim i As Integer, j As Integer, xm As String
    
    Dim dy As Worksheet, xmrng As Range
    Dim maxrow As Integer, curcol As Integer, nextrow As Integer, nextcol As Integer
    
    Set dicnormalday = CreateObject("scripting.dictionary")
    Set reg = CreateObject("vbscript.regexp")
    
       
    reg.Pattern = "\d+(,\d+)*"
    reg.Global = True
    
    If reg.test(ws.Name) Then
       ' MsgBox ws.Name
        ws.Activate
             
        'MsgBox (Cells(4, 15).MergeCells)
        For i = 15 To 45 Step 15
                    If Application.CountA(Range(Cells(13, i - 13), Cells(43, i - 1))) > 0 Then
                              
                                'MsgBox i
                                xm = Cells(4, i - 5).Value
'
                                'if activeworkbook.s
                                If isexistssheet("打印表") Then
                                    Call createPrinttb("打印表", ActiveWorkbook.Worksheets.Count)
                                End If
                                Set dy = ActiveWorkbook.Worksheets("打印表")
                                
                                If dy.Range("A1").CurrentRegion.Rows.Count = 1 Then
                                    nextrow = 1
                                    nextcol = 1
                                    maxrow = 49
                                    
                                Else
                                
                                    'crr = dy.UsedRange.Value
                                    'maxrow = UBound(crr)
                                    'maxcol = UBound(crr, 2)
                                    With dy.UsedRange
                                        Set xmrng = .Find(xm, LookIn:=xlValues)
                                        If Not xmrng Is Nothing Then
                                            maxrow = xmrng.Row + 48
                                            curcol = xmrng.Column + 5
                                            
                                            'firstAddress = xmrng.Address
'                                            Do
'                                                c.Value = 5
'                                                Set c = .FindNext(c)
'                                            Loop While Not c Is Nothing
                                        
                                        Else
                                            maxrow = .Rows.Count
                                            curcol = dy.Cells(maxrow - 40, Columns.Count).End(xlToLeft).Column
                                        End If
                                    End With
                                    If curcol = 15 Then
                                        nextrow = maxrow - 48
                                        nextcol = 17
                                    ElseIf curcol = 31 Then
                                        nextrow = ((maxrow - 49) / 54 + 1) * 54 + 1
                                        nextcol = 1
                                    End If
                                            
                                            'nextrow = (maxcol - 49) / 54 + 1
'                                            If curcol = 15 Then
'                                                nextcol = 17
'                                            ElseIf curcol = 31 Then
'                                                nextcol = 1
'                                            End If
                                        
                                    
                                 End If
                                    'maxrows = dy.Cells(Rows.Count, 1).End(xlUp).Row
                                    'maxcol = dy.cells(columns.Count
                                 Range(Cells(4, i - 14), Cells(8, i - 1)).Copy dy.Range(numToString(nextcol) & nextrow)
                                 Range(Cells(10, i - 14), Cells(52, i)).Copy dy.Range(numToString(nextcol) & (nextrow + 6))
                                 
                                 'MsgBox maxrow
                                 With dy.PageSetup
                                    .PrintArea = "$A$1:$AE$" & (nextrow + 48)
                                 End With
                                 
                                
                                
                        End If
            Next
           
        
    End If
    Set reg = Nothing
End Sub
