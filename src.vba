'
' 根据选中的行提取出定额数据
' 并进行数据的检查检对。
' 作者：龙仕云  2025-4-13
'
'
'

Function ABCPosion(Astr As String) As Integer
    ABCPosion = 0
    If Len(Astr) <> 1 Then Exit Function
    If UCase(Astr) = "A" Then ABCPosion = 1
    If UCase(Astr) = "B" Then ABCPosion = 2
    If UCase(Astr) = "C" Then ABCPosion = 3
    If UCase(Astr) = "D" Then ABCPosion = 4
    If UCase(Astr) = "E" Then ABCPosion = 5
    If UCase(Astr) = "F" Then ABCPosion = 6
    If UCase(Astr) = "G" Then ABCPosion = 7
    If UCase(Astr) = "H" Then ABCPosion = 8
    If UCase(Astr) = "I" Then ABCPosion = 9
    If UCase(Astr) = "J" Then ABCPosion = 10
    If UCase(Astr) = "K" Then ABCPosion = 11
    If UCase(Astr) = "L" Then ABCPosion = 12
    If UCase(Astr) = "M" Then ABCPosion = 13
    If UCase(Astr) = "N" Then ABCPosion = 14
    If UCase(Astr) = "O" Then ABCPosion = 15
End Function

Function PosionABC(APos As Long) As String
    PosionABC = ""
    If APos < 1 Then Exit Function
    If APos = 1 Then PosionABC = "A"
    If APos = 2 Then PosionABC = "B"
    If APos = 3 Then PosionABC = "C"
    If APos = 4 Then PosionABC = "D"
    If APos = 5 Then PosionABC = "E"
    If APos = 6 Then PosionABC = "F"
    If APos = 7 Then PosionABC = "G"
    If APos = 8 Then PosionABC = "H"
    If APos = 9 Then PosionABC = "I"
    If APos = 10 Then PosionABC = "J"
    If APos = 11 Then PosionABC = "K"
    If APos = 12 Then PosionABC = "L"
    If APos = 13 Then PosionABC = "M"
    If APos = 14 Then PosionABC = "N"
    If APos = 15 Then PosionABC = "O"
    
End Function

'
'获取单元格的值，考虑到单元格合并了
'
Function GetCellValue(ArowNumber As Long, AcolNumber As Long) As Variant
    Dim myStr As String
    Dim targetCell As Range
    
    myStr = PosionABC(AcolNumber) + CStr(ArowNumber)
    Set targetCell = Range(myStr) ' 要检查的单元格
    
    If targetCell.MergeCells Then
        GetCellValue = targetCell.MergeArea.Cells(1, 1).Value
    Else
        GetCellValue = ActiveSheet.Cells(ArowNumber, AcolNumber)
    End If

End Function

'字符串去空格及中间空格
Function StrTrim(ByVal Astr As String) As String
    ' 去两端空格
    StrTrim = Trim(Astr)
    ' 去中间所有空格
    StrTrim = Replace(StrTrim, " ", "")
End Function

'获取定额信息
Function GetDeInfo(ARowIndex As Long, AColIndex As Long) As Variant()
    Dim myRow As Long
    Dim mydata(1 To 10) As Variant
    Dim c As Long
    Dim rgf As Double '人工费
    Dim clf As Double '材料费
    Dim jxf As Double '机械费
    Dim glf As Double '管理费
    Dim lr As Double   '利润
    Dim zhdj As Double  '综合单价
    Dim debh As String '定额编号
    Dim demc As String '定额名称
    Dim dedw As String '定额单位
    Dim gznr As String '工作内容
    Dim myvalue As Variant
    Dim mystr2 As String
    Dim mymcs(1 To 10) As String '名称内容存在多个名称
    Dim mymcidx As Long
    Dim pos As Long
    
    Dim rowMc As Long '材料名称列
    Dim rowDw As Long '材料单位列
    
    
    
    rowMc = ABCPosion(Sheet2.Range("C3"))
    rowDw = ABCPosion(Sheet2.Range("C4"))
    
    
    
    rgf = 0
    clf = 0
    jxf = 0
    glf = 0
    lr = 0
    zhdj = 0
    c = 1
    mymcidx = 1
    
    For myRow = ARowIndex - 1 To 1 Step -1
        If c > 20 Then Exit For  '向上20行最多了后到定额退出
        myvalue = GetCellValue(myRow, AColIndex)
        mystr2 = GetCellValue(myRow, rowMc)
        If StrTrim(mystr2) = "" Then mystr2 = GetCellValue(myRow, rowDw)
        mystr2 = StrTrim(mystr2)
                
        
        If (InStr(1, mystr2, "定额编号") > 0) Then
            debh = myvalue
        End If
        
        If (InStr(1, mystr2, "项目") > 0) Then
            mymcs(mymcidx) = myvalue
            mymcidx = mymcidx + 1
        End If
        
        
        If (InStr(1, mystr2, "综合单价") > 0) Then
            zhdj = myvalue
        End If
        
        If (InStr(1, mystr2, "人工费") > 0) Then
            rgf = myvalue
        End If
        
        If (InStr(1, mystr2, "材料费") > 0) Then
            clf = myvalue
        End If
        
        If (InStr(1, mystr2, "机械费") > 0) Then
            jxf = myvalue
        End If
        
        If (InStr(1, mystr2, "管理费") > 0) Then
            glf = myvalue
        End If
        
        If (InStr(1, mystr2, "利润") > 0) Then
            lr = myvalue
        End If
        
        If (InStr(1, mystr2, "工作内容：") > 0) And (InStr(1, mystr2, "计量单位") > 0) Then
            pos = InStrRev(myvalue, "计量单位")
            dedw = StrTrim(Mid(myvalue, pos + 5))
            gznr = StrTrim(Mid(myvalue, 6, pos - 6))
            
            Exit For '退出了，已找全部数据
        End If
        
        c = c + 1
    Next myRow

    mymcidx = 10
    For mymcidx = 10 To 1 Step -1
        If demc = "" Then
            demc = mymcs(mymcidx)
        Else
            demc = demc + " " + mymcs(mymcidx)
        End If
    Next mymcidx
   
   mydata(1) = debh
   mydata(2) = demc
   mydata(3) = dedw
   mydata(4) = zhdj
   mydata(5) = rgf
   mydata(6) = clf
   mydata(7) = jxf
   mydata(8) = glf
   mydata(9) = lr
   mydata(10) = gznr
   
   GetDeInfo = mydata
End Function

'获取定额的人材机
Function GetDeRCJ(ARowIndex As Long, AColIndex As Long, ADeInfo() As Variant) As Boolean

    Dim selectedRange As Range
    Dim targetRow As Range
    Dim outputStr As String
    Dim outputStr2 As String
    Dim myRow As Long
    
    Dim rowLx As Long '类型
    Dim rowMc As Long '材料名称列
    Dim rowDw As Long '单位列
    Dim rowDj As Long '单价列
    Dim rowSl As Long '数量

    rowLx = ABCPosion(Sheet2.Range("C2"))
    rowMc = ABCPosion(Sheet2.Range("C3"))
    rowDw = ABCPosion(Sheet2.Range("C4"))
    rowDj = ABCPosion(Sheet2.Range("C5"))
    
     '遍历选区中的每一行
     Set selectedRange = Application.Selection
    For Each targetRow In selectedRange.Rows
        myRow = targetRow.Row
            
         ' 获取列的值，可能有多列
        outputStr2 = ActiveSheet.Cells(myRow, AColIndex)
    
        outputStr = outputStr & "行 " & rowNumber & " 的值: " & outputStr
    
    Next targetRow

    GetDeRCJ = True
End Function



'方法入口
Sub 获取定额()
       
    Dim rowNumber As Long
    Dim colNumber As Long
    Dim col As Range

    Dim Deinfo() As Variant
    Dim myBool As Boolean
    
    
    On Error Resume Next
    
    On Error GoTo 0
    
    If Application.Selection.Areas.Count > 1 Then
        MsgBox "只能选择一个连续的区域！"
        Exit Sub
    End If
    
    ' 遍历选区中的每一列,有合并的情况
    'For Each col In selectedRange.Columns
    '    colNumbers = colNumbers & col.Column & ", "
    'Next col
    
    
    '提取定额数据
    rowNumber = Application.Selection.Row
    colNumber = Application.Selection.Column
    If StrTrim(GetCellValue(rowNumber - 1, colNumber)) <> "数量" Then
        MsgBox "只能选择一个连续的区域时，第一个选择定额的第一个材料！"
        Exit Sub
    End If
    
    Deinfo = GetDeInfo(rowNumber - 1, colNumber)
    myBool = GetDeRCJ(rowNumber, colNumber, Deinfo)
    
    
    MsgBox "已处理" & Deinfo(1)
End Sub

