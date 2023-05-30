Attribute VB_Name = "md_GLOBALS_FUNCTIONS"
'@Folder("Mods.Globals")
' # md_GLOBALS_FUNCTIONS #

Private rowColCount()

Public Function stripLineBreaks(ByVal str As String) As String
    str = LCase(Replace(str, vbCrLf, "¦"))
    str = Replace(str, vbLf, "¦")
    str = Replace(str, Chr(9), " ")
    str = Replace(str, ";", " ")
    str = Replace(str, "¦", " ¦ ")
    str = Replace(str, ", ", ",")
    str = Replace(str, " ,", ",")
    str = Replace(str, " , ", ",")
    str = Replace(str, " as ", " ")
    str = Replace(str, " go ", " ")
    str = Replace(str, " begin ", " ")
    str = Replace(str, " end ", " ")
        
    Dim commentFound, endLine, temp
    
    commentFound = InStr(1, str, "/*")
    If commentFound > 0 Then
        Do
            temp = str
            commentFound = InStr(1, str, "/*")
            If commentFound > 0 Then
                endLine = InStr(commentFound, str, "*/")
                str = Mid(str, 1, commentFound - 1) + Mid(str, endLine + 2, Len(str) - endLine - 1)
            End If
        Loop Until (temp = str)
    End If
    
    commentFound = InStr(1, str, "--")
    If commentFound > 0 Then
        Do
            temp = str
            commentFound = InStr(1, str, "--")
            If commentFound > 0 Then
                endLine = InStr(commentFound, str, "¦")
                str = Mid(str, 1, commentFound - 1) + Mid(str, endLine + 1, Len(str) - endLine)
            End If
        Loop Until (temp = str)
    End If
    
    str = Replace(str, "with (nolock)", " ")
    Do
        temp = str
        str = Replace(str, Space(2), Space(1))
        str = Replace(str, "¦¦", "¦")
        str = Replace(str, "    ", " ")
    Loop Until (temp = str)
    
    stripLineBreaks = str
End Function

Public Function isSqlKeyWord(word_) As Boolean
    Dim keyWords_, i
    keyWords_ = getKeyWords
    For i = 0 To UBound(keyWords_)
        If keyWords_(i) = word_ Then
            isSqlKeyWord = True
            Exit Function
        End If
    Next i
    isSqlKeyWord = False
End Function

Public Function isSqlCompareWord(word_) As Boolean
    Dim keyWords_, i
    keyWords_ = getCompareOpperators
    For i = 0 To UBound(keyWords_)
        If keyWords_(i) = word_ Then
            isSqlCompareWord = True
            Exit Function
        End If
    Next i
    isSqlCompareWord = False
End Function

Public Function inSqlCompareWord(word_) As Boolean
    Dim keyWords_, i
    keyWords_ = getCompareOpperators
    For i = 0 To UBound(keyWords_)
        If InStr(1, word_, keyWords_(i)) > 0 Then
            inSqlCompareWord = True
            Exit Function
        End If
    Next i
    inSqlCompareWord = False
End Function

Public Function isSqlJoinWord(word_) As Boolean
    isSqlJoinWord = (word_ = "join") Or (word_ = "union") Or (word_ = "cross") Or (word_ = "left") Or (word_ = "right") Or (word_ = "inner") Or (word_ = "full") Or (word_ = "outer") Or InStr(1, word_, "apply")
End Function

Public Function inSqlJoinWord(word_) As Boolean
    inSqlJoinWord = InStr(1, word_, "join") > 0 Or InStr(1, word_, "union") Or InStr(1, word_, "cross") Or InStr(1, word_, "left") Or InStr(1, word_, "right") Or InStr(1, word_, "inner") Or InStr(1, word_, "full") Or InStr(1, word_, "outer") Or InStr(1, word_, "apply")
End Function

Public Function debugData(ParamArray v() As Variant)
    Dim i
    Dim outPutStr As String
    outPutStr = ""
    For i = LBound(v) To UBound(v)
        If i + 1 > UBound(v) Then
            outPutStr = outPutStr + v(i) & "¦"
        Else
            outPutStr = outPutStr + Right(Space(v(i + 1)) & v(i), v(i + 1)) & "¦"
            i = i + 1
        End If
    Next i
    outPutStr = Mid(outPutStr, 1, Len(outPutStr) - 1)
    Debug.Print outPutStr
End Function

Public Sub CodeCounter()
    On Error GoTo CodeLineCount_Err
    Dim CodeLineCount As Double
    Set CodeLineCount_Var = ThisWorkbook.VBProject
    For Each CodeLineCount_Var In CodeLineCount_Var.VBComponents
        CodeLineCount = CodeLineCount + CodeLineCount_Var.CodeModule.CountOfLines
    Next
    CodeLineCount_Total = CodeLineCount
    
CodeLineCount_Err:
    Set CodeLineCount_Var = Nothing
    Debug.Print CodeLineCount_Total & " lines of code in project!"
End Sub

Public Function setOutPutArray()
    ReDim rowColCount(0 To 0)
End Function

Public Function outputDataToSheet(ws As Worksheet, row, Optional val = "#xj-1", Optional startCol = -1, Optional cellFormats As String = "")
    Dim col, k, v
    Dim r As Range
    
    
    If row > UBound(rowColCount) Then
        ReDim Preserve rowColCount(0 To row)
    End If
    
    If startCol > -1 Then
        rowColCount(row) = startCol
        Set r = ws.Cells(row, rowColCount(row))
    Else
        rowColCount(row) = rowColCount(row) + 1
        Set r = ws.Cells(row, rowColCount(row))
    End If
    
    With r
        If val <> "#xj-1" Then
            .Value = val
        End If
        If cellFormats <> "" Then
            For Each k In Split(cellFormats, ",")
                    Select Case Split(k, ":")(0)
                        Case "back"
                            .Interior.Color = Split(k, ":")(1)
                        Case "fore"
                            .Font.Color = Split(k, ":")(1)
                        Case "bold"
                            .Font.Bold = Split(k, ":")(1)
                        Case "font"
                            .Font.name = Split(k, ":")(1)
                        Case "size"
                            .Font.Size = Split(k, ":")(1)
                        Case "align"
                            Select Case LCase(Split(k, ":")(1))
                                Case "left"
                                    .HorizontalAlignment = xlLeft
                                Case "center"
                                    .HorizontalAlignment = xlCenter
                                Case "right"
                                    .HorizontalAlignment = xlRight
                                Case Else
                                    .HorizontalAlignment = xlLeft
                            End Select
                        Case "autofit"
                            .EntireColumn.AutoFit
                    End Select
            Next k
        End If
    End With
End Function

Public Function parseAllJson()
    Dim results(), allJson, i
    allJson = Sheets("Parsed Data").UsedRange.Value
    ReDim results(0 To UBound(allJson, 1) - 1)
    Dim time_
    time_ = Timer
    For i = 0 To UBound(results)
        results(i) = Array(fromJson(allJson(i + 1, 1))(1), fromJson(allJson(i + 1, 2))(1), fromJson(allJson(i + 1, 3))(1), fromJson(allJson(i + 1, 4))(1))
    Next i
    
    parseAllJson = results
    'debug.Print Timer - time_
End Function

Public Function inArray(search, arr, Optional subIndex = Empty) As fn_inArrayResult
    Dim i
    If IsArray(arr) = True Then
        If LBound(arr) > -1 Then
            If UBound(arr) > 0 Then
                For i = LBound(arr) To UBound(arr)
                    If IsEmpty(subIndex) = True Then
                        If arr(i) = search Then
                            inArray.found = True
                            inArray.foundIndex = i
                            Exit Function
                        End If
                    Else
                        If arr(i)(subIndex) = search Then
                            inArray.found = True
                            inArray.foundIndex = i
                            Exit Function
                        End If
                    End If
                Next i
            End If
        End If
    End If
    inArray.found = False
    inArray.foundIndex = -1
End Function

Public Function isemptyLastTableSearch(v() As qr_tableMainSearchData) As Boolean
    On Error GoTo exitisemptyLastTableSearch
        Dim a
        For a = LBound(v) To UBound(v)
            If v(a).count > -1 Then
                isemptyLastTableSearch = False
                On Error GoTo 0
                Exit Function
            End If
        Next a
exitisemptyLastTableSearch:
    On Error GoTo 0
    isemptyLastTableSearch = True
End Function

Public Function searchForTableButton()
    searchJsonTables Sheets("Search For Table")
End Function


'test stuff

Public Sub affa()
    Application.EnableEvents = False
    
    Const w = 3
    Const h = 5
    Const cnt = 6
    
    
    Dim p
    p = Sheet6.Range("G1:G" + CStr(cnt)).Value
    
    Dim a, b, c, d, e, f, i
    Dim lb, ub, ci, ub1
    
    Dim arr(), cnts(), lastfound()
    ReDim arr(1 To h, 1 To w)
    ReDim cnts(1 To cnt)
    ReDim lastfound(1 To cnt)
    
    lb = CVar(CDbl(String(w * h, "1")))
    ub1 = UBound(p)
    ub = CVar(CDbl(String(w * h, CStr(ub1))))
    ci = lb
    
    Do
        For a = 1 To Len(ci)
            'Sheet6.Cells(Int((a - 1) / w) + 1, ((a - 1) Mod w) + 1).Value = p(Mid(CStr(ci), a, 1), 1)
            arr(Int((a - 1) / w) + 1, ((a - 1) Mod w) + 1) = p(Mid(CStr(ci), a, 1), 1)
        Next a
        
        For a = 1 To h
            
        Next a
        
        Sheet6.Range(Sheet6.Cells(1, 1), Sheet6.Cells(h, w)).Value = arr
        
        increaseCI ci, ub1, 1
    Loop Until (ci = ub)
    
    
    Application.EnableEvents = True
End Sub

Public Function increaseCI(ByRef ci, ub1, pos)
    Dim a, b
    Dim ln, cn, rn, l
    l = Len(ci)

    If l > 1 Then
        ln = Mid(ci, 1, (l - pos))
        cn = Mid(ci, (l - pos) + 1, 1)
        If pos > 1 Then
            rn = Mid(ci, (l - pos) + 2, l)
        Else
            rn = ""
        End If
        
        If cn + 1 > ub1 Then
            cn = 1
            ci = "" + CStr(ln) + CStr(cn) + CStr(rn)
            increaseCI ci, ub1, pos + 1
        Else
            cn = cn + 1
            ci = "" + CStr(ln) + CStr(cn) + CStr(rn)
            Exit Function
        End If
    Else
        If ci + 1 > ub1 Then
            ci = 1
        Else
            ci = ci + 1
        End If
    End If

    
End Function


Public Function sortStrReg(ByRef arr() As String)
    Dim a, b, c, d
    Do
        a = False
        For b = UBound(arr) To 1 Step -1
            If arr(b) < arr(b - 1) Then
                c = arr(b)
                arr(b) = arr(b - 1)
                arr(b - 1) = c
                a = True
            End If
        Next b
    Loop Until a = False
End Function

Public Function formatRange(rng_ As Range, params_ As String) As Boolean
    Dim a, b, c, d
    a = Split(params_, ",")
    For Each b In a
        c = Split(b, ":")
        Select Case LCase(c(0))
            Case "bg"
                rng_.Interior.Color = c(1)
            Case "fg"
                rng_.Font.Color = c(1)
            Case "fontName"
                rng_.Font.name = c(1)
            Case "fontSize"
                rng_.Font.Size = c(1)
            Case "halign"
                Select Case LCase(c(1))
                    Case "right"
                        d = xlRight
                    Case "center"
                        d = xlCenter
                    Case "left"
                        d = xlLeft
                End Select
                rng_.HorizontalAlignment = d
            Case "valign"
                Select Case LCase(c(1))
                    Case "top"
                        d = xlTop
                    Case "center"
                        d = xlCenter
                    Case "bottom"
                        d = xlBottom
                End Select
                rng_.VerticalAlignment = d
            Case "border"
                Select Case LCase(c(1))
                    Case "thin"
                        d = xlThin
                    Case "med", "medium"
                        d = xlMedium
                    Case "lrg", "large"
                        d = xlThick
                End Select
                For i = 7 To 10
                    With Selection.Borders(i)
                        .LineStyle = xlContinuous
                        .ColorIndex = 0
                        .TintAndShade = 0
                        .Weight = d
                    End With
                Next i
        End Select
    Next b
End Function
