Attribute VB_Name = "md_Query_Parser"
'@Folder("Mods.Query.Manager")
'# md_Query_Parser #

Public Function outputColSearchResults(searchWord, columnSearch() As qr_columnMainSearchData, sh_ As Worksheet)
    Dim i, j
    setOutPutArray
    sh_.Cells.ClearContents
    sh_.Cells.ClearFormats
    sh_.Cells.Font.Color = RGB(0, 0, 0)
    On Error Resume Next
    sh_.ShowAllData
    On Error GoTo 0
    
    Call outputDataToSheet(sh_, 1, "Search Column:", , "name:Consolas,back:12566463,fore:11892015,size:16,autofit:true,bold:true,align:right")
    Call outputDataToSheet(sh_, 1, searchWord, , "name:Consolas,back:12566463,size:14,autofit:true,bold:true")
    
    
    Call outputDataToSheet(sh_, 2, "Matching Columns:", , "size:12,align:right")
    Call outputDataToSheet(sh_, 2, IIf(LBound(columnSearch) > -1, UBound(columnSearch) + 1, 0), , "size:10")
    
    'sh_.Rows(c_OutputHeaderRow).AutoFilter
    
    Call outputDataToSheet(sh_, c_OutputHeaderRow, "Columns Found", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    Call outputDataToSheet(sh_, c_OutputHeaderRow, "Column Alias", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    Call outputDataToSheet(sh_, c_OutputHeaderRow, "Home Table Name", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Home Table Alias")
    Call outputDataToSheet(sh_, c_OutputHeaderRow, "Found Count", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Exact Match", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Match Streak", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Match Streak Location", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Letter Match Count", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Letter Match Location", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Word Length", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    
    If LBound(columnSearch) > -1 Then
        For j = c_OutputHeaderRow + 1 To (UBound(columnSearch) - LBound(columnSearch) + (c_OutputHeaderRow + 1))
            Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.word, "'", "`"))
            sh_.Cells(j, 1).Font.Color = RGB(100, 100, 123)
            With sh_.Cells(j, 1).Characters(Start:=columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_firstSpotOfMatch, Length:=columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_maxLenthOfMatch).Font
                .FontStyle = "bold"
                .Color = RGB(10, 105, 0)
                .Size = .Size + 2
            End With
            
            Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).aliasSearchResults.word, "'", "`"), , "align:left")
            Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.word, "'", "`"), , "align:left,autofit:true")
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasSearchResults.word, "'", "`"))
            Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).count, "'", "`"), , "align:center")
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.exactMatch, "'", "`"))
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_maxLenthOfMatch, "'", "`"))
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_firstSpotOfMatch, "'", "`"))
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_countOfMatchedLetters, "'", "`"))
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.match_firstSpotOfLetterMatch, "'", "`"))
            'Call outputDataToSheet(sh_, j, Replace(columnSearch(j - (c_OutputHeaderRow + 1)).columnSearchResults.lengthOfWord, "'", "`"))
        Next j
    End If
End Function

Public Function combineColSearchResults(ByRef searchResults() As qr_columnMainSearchData) As qr_columnMainSearchData()
    Dim final_Results() As qr_columnMainSearchData
    Dim i, j, k
    Dim found_ As Boolean
    Dim tableAliasFound_ As Boolean
    Dim columnAliasFound_ As Boolean
    Dim tempSplit As Variant
  
    ReDim final_Results(-1 To 0)
    If LBound(searchResults) > -1 Then
        For i = 0 To UBound(searchResults)
            found_ = False
            If LBound(final_Results) > -1 Then
                For j = 0 To UBound(final_Results)
                    If matchSearchResults(searchResults(i), final_Results(j)) Then
                        final_Results(j).count = final_Results(j).count + 1
                        tableAliasFound_ = False
                        columnAliasFound_ = False
                        tempSplit = Split(final_Results(j).aliasSearchResults.word, " ¦ ")
                        For k = 0 To UBound(tempSplit)
                            If tempSplit(k) = searchResults(i).aliasSearchResults.word Then
                                columnAliasFound_ = True
                                Exit For
                            End If
                        Next k
                        
                        tempSplit = Split(final_Results(j).tableData.aliasSearchResults.word, " ¦ ")
                        For k = 0 To UBound(tempSplit)
                            If tempSplit(k) = searchResults(i).tableData.aliasSearchResults.word Then
                                tableAliasFound_ = True
                                Exit For
                            End If
                        Next k
                        
                        If columnAliasFound_ = False And searchResults(i).aliasSearchResults.word <> "" Then
                            final_Results(j).aliasSearchResults.word = final_Results(j).aliasSearchResults.word & " ¦ " & searchResults(i).aliasSearchResults.word
                        End If
                        
                        If tableAliasFound_ = False And searchResults(i).tableData.aliasSearchResults.word <> "" Then
                            final_Results(j).tableData.aliasSearchResults.word = final_Results(j).tableData.aliasSearchResults.word & " ¦ " & searchResults(i).tableData.aliasSearchResults.word
                        End If
                        
                        found_ = True
                        Exit For
                    End If
                Next j
            End If
            If found_ = False Then
                If LBound(final_Results) = -1 Then
                    ReDim final_Results(0 To 0)
                    final_Results(0) = searchResults(i)
                Else
                    ReDim Preserve final_Results(0 To UBound(final_Results) + 1)
                    final_Results(UBound(final_Results)) = searchResults(i)
                End If
            End If
        Next i
    End If
    
    searchResults = final_Results
End Function

Public Function matchSearchResults(l As qr_columnMainSearchData, r As qr_columnMainSearchData) As Boolean
    matchSearchResults = l.columnSearchResults.word = r.columnSearchResults.word And _
                         l.tableData.tableSearchResults.word = r.tableData.tableSearchResults.word 'And _
                         l.columnAlias = r.columnAlias And
End Function

Public Function sortColSeachResults(ByRef searchResults() As qr_columnMainSearchData) As qr_columnMainSearchData()
    Dim i, j, tempCol As qr_columnMainSearchData
    Dim change As Boolean
    Dim sortTime
    sortTime = Timer
    change = False
       
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = True And searchResults(i - 1).columnSearchResults.exactMatch = False) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch > searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch = searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) And (searchResults(i).count > searchResults(i - 1).count) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch = searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) And (searchResults(i).count = searchResults(i - 1).count) And (searchResults(i).columnSearchResults.match_firstSpotOfMatch < searchResults(i - 1).columnSearchResults.match_firstSpotOfMatch) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch = searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) And (searchResults(i).count = searchResults(i - 1).count) And (searchResults(i).columnSearchResults.match_firstSpotOfMatch = searchResults(i - 1).columnSearchResults.match_firstSpotOfMatch) And (searchResults(i).columnSearchResults.match_countOfMatchedLetters > searchResults(i - 1).columnSearchResults.match_countOfMatchedLetters) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch = searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) And (searchResults(i).count = searchResults(i - 1).count) And (searchResults(i).columnSearchResults.match_firstSpotOfMatch = searchResults(i - 1).columnSearchResults.match_firstSpotOfMatch) And (searchResults(i).columnSearchResults.match_countOfMatchedLetters > searchResults(i - 1).columnSearchResults.match_countOfMatchedLetters) And (searchResults(i).columnSearchResults.lengthOfWord < searchResults(i - 1).columnSearchResults.lengthOfWord) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    
    Do
        change = False
        For i = UBound(searchResults) To 1 Step -1
            If (searchResults(i).columnSearchResults.exactMatch = searchResults(i - 1).columnSearchResults.exactMatch) And (searchResults(i).columnSearchResults.match_maxLenthOfMatch = searchResults(i - 1).columnSearchResults.match_maxLenthOfMatch) And (searchResults(i).count = searchResults(i - 1).count) And (searchResults(i).columnSearchResults.match_firstSpotOfMatch = searchResults(i - 1).columnSearchResults.match_firstSpotOfMatch) And (searchResults(i).columnSearchResults.match_countOfMatchedLetters > searchResults(i - 1).columnSearchResults.match_countOfMatchedLetters) And (searchResults(i).columnSearchResults.lengthOfWord = searchResults(i - 1).columnSearchResults.lengthOfWord) And (searchResults(i).columnSearchResults.word < searchResults(i - 1).columnSearchResults.word) Then
                change = True
                tempCol = searchResults(i)
                searchResults(i) = searchResults(i - 1)
                searchResults(i - 1) = tempCol
            End If
        Next i
    Loop Until (change = False)
    
    i = i
    'debug.Print "Took " & Round(Timer - sortTime, 2) & " to sort columns"
End Function

Public Function generateJson()
    Sheet8.Cells.ClearContents
    Dim i
    For i = 0 To qryTree.QueryCount - 1
        jArr = toJson(qryTree.query(i))
                
        With Sheet8
            .Cells(i + 1, 1).Value = jArr(0)
            .Cells(i + 1, 2).Value = jArr(1)
            .Cells(i + 1, 3).Value = jArr(2)
            .Cells(i + 1, 4).Value = jArr(3)
            .Cells(i + 1, 5).Value = jArr(4)
        End With
    Next i
End Function

Public Function searchForColumn(columnName, jsonData) As qr_columnMainSearchData()
    Dim scanResults() As qr_searchStringResults
    Dim columnData() As qr_columnMainSearchData
    Dim final_columnData() As qr_columnMainSearchData
    Dim i As Long, j As Integer, x As Integer, y As Integer
    'Dim tempCol As cls_Column
    'Dim temp As cls_Query
    
    Dim tempCol As Variant
    Dim temp As Variant
    
    Dim currColCount As Long
    Dim max_Count As Integer
    Dim startTime
    
    max_Count = 0
    Dim timerData(), dblArr(0 To 9) As Double
    Dim timercount, looptime, queryTime
    ReDim timerData(0 To 999999)
    ReDim columnData(0 To 999999)
    Dim dbl As Double
    timercount = 0
    looptime = Timer
    startTime = Timer
    currColCount = 0
    For i = 0 To UBound(jsonData)
        temp = jsonData(i)
        If IsArray(temp(0)) = True Then
            For j = 0 To UBound(temp(0))
                Dim jArr
                'timerData(timercount) = dblArr
                tempCol = temp(0)(j)
                
                'timerData(timercount)(0) = Round(Timer - looptime, 4)
                'looptime = Timer
                
                columnData(currColCount).homeQueryIndex = i
                
                'timerData(timercount)(1) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                columnData(currColCount).columnHomeTable = tempCol(2)
                
                'timerData(timercount)(2) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                columnData(currColCount).columnAlias = tempCol(0)
                
                'timerData(timercount)(3) = Round(Timer - looptime, 4)
                'looptime = Timer
            
                Call getScanResults(LCase(columnName), tempCol(1), columnData(currColCount).columnSearchResults)
                If columnData(currColCount).columnSearchResults.match_firstSpotOfMatch = -1 Then
                    GoTo nextJLoop
                End If
                
                'timerData(timercount)(4) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                Call getScanResults(LCase(columnName), tempCol(0), columnData(currColCount).aliasSearchResults)
                
                'timerData(timercount)(5) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                If columnData(currColCount).columnSearchResults.match_maxLenthOfMatch > max_Count Then
                    max_Count = columnData(currColCount).columnSearchResults.match_maxLenthOfMatch
                End If
                
                'timerData(timercount)(6) = Round(Timer - looptime, 4)
                'looptime = Timer
                
                'Dim tempTable As cls_TABLE
                'Set tempTable = qryTree.query(i).getTableByAlias(tempCol.table)
                tempTable = scanJsonForTable(temp(1), tempCol(2))
                
                'timerData(timercount)(7) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                If tempTable(0) <> Empty Then
                    Call getScanResults(LCase(columnName), tempTable(0), columnData(currColCount).tableData.tableSearchResults)
                    Call getScanResults(LCase(columnName), tempTable(1), columnData(currColCount).tableData.aliasSearchResults)
                Else
                    If IsArray(temp(1)) = True Then
                        If UBound(temp(1)) = 0 Then
                            columnData(currColCount).tableData.tableSearchResults.word = temp(1)(0)(0)
                        End If
                    End If
                End If
                
                With columnData(currColCount).tableData.tableSearchResults
                    .word = Replace(.word, "[", "")
                    .word = Replace(.word, "]", "")
                    .word = Replace(.word, "rpt.", "")
                    .word = Replace(.word, "dbo.", "")
                    .word = Replace(.word, "integra.", "")
                End With
                
                With columnData(currColCount).tableData.aliasSearchResults
                    .word = Replace(.word, "[", "")
                    .word = Replace(.word, "]", "")
                    .word = Replace(.word, "rpt.", "")
                    .word = Replace(.word, "dbo.", "")
                    .word = Replace(.word, "integra.", "")
                End With
                
                'timerData(timercount)(8) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                columnData(currColCount).tableData.homeQueryIndex = i
                columnData(currColCount).count = 1
                
                'timerData(timercount)(9) = Round(Timer - looptime, 4)
                'looptime = Timer
        
                'timercount = timercount + 1
                currColCount = currColCount + 1
                
nextJLoop:
            Next j
        End If
    Next i
    currColCount = currColCount - 1
    ReDim Preserve columnData(0 To currColCount)
    ReDim Preserve timerData(0 To currColCount)
    'getAvrgTimes timerData
    
    'debug.Print "Took " & Round(Timer - startTime, 2) & " to search through columns"
    startTime = Timer
    ReDim final_columnData(-1 To 0)
    For i = 0 To currColCount
        If columnData(i).columnSearchResults.match_maxLenthOfMatch >= (max_Count * c_MatchStreakPercentile) Then
            If InStr(1, columnData(i).tableData.tableSearchResults.word, "#") = 0 And InStr(1, columnData(i).tableData.tableSearchResults.word, "(") = 0 And Trim(columnData(i).tableData.tableSearchResults.word) <> "" Then
                If LBound(final_columnData) = -1 Then
                    ReDim final_columnData(0 To 0)
                Else
                    ReDim Preserve final_columnData(0 To UBound(final_columnData) + 1)
                End If
                
                final_columnData(UBound(final_columnData)) = columnData(i)
            End If
        End If
    Next i
    'debug.Print "Took " & Round(Timer - startTime, 2) & " to clean up columns"
    searchForColumn = final_columnData
End Function

Public Function getScanResults(searchWord As String, scanWord, ByRef scanResults As qr_searchStringResults)
    Dim i, j, tempc, tempd, tempe, tempf
    Dim wordLength
    Dim countOfMatchedLetters
    Dim firstSpotOfLetterMatch
    Dim firstSpotOfMatch
    Dim maxLenthOfMatch
    Dim word
    Dim currLetterScan, currLetterSearch
    Dim tempLetter
        
    Dim currentMatchSpotOfScan
    Dim currentRunOfMatch
    
    If IsEmpty(scanWord) = True Or Len(searchWord) > Len(scanWord) Then
        With scanResults
            .lengthOfWord = -1
            .match_countOfMatchedLetters = -1
            .match_firstSpotOfLetterMatch = -1
            .match_firstSpotOfMatch = -1
            .match_maxLenthOfMatch = -1
            .word = scanWord
            .exactMatch = False
        End With
        Exit Function
    End If
        
    Dim wordLen, letterSearch, matchSpot
    wordLen = Len(searchWord)
    For i = wordLen To 1 Step -1
        matchSpot = InStr(1, scanWord, Mid(searchWord, 1, i))
        If matchSpot > 0 Then
            With scanResults
                .lengthOfWord = Len(scanWord)
                .match_countOfMatchedLetters = i
                .match_firstSpotOfLetterMatch = matchSpot
                .match_firstSpotOfMatch = matchSpot
                .match_maxLenthOfMatch = i
                .word = scanWord
                .exactMatch = (wordLen = Len(scanWord))
            End With
            Exit Function
        End If
    Next i
        
    With scanResults
        .lengthOfWord = -1
        .match_countOfMatchedLetters = -1
        .match_firstSpotOfLetterMatch = -1
        .match_firstSpotOfMatch = -1
        .match_maxLenthOfMatch = -1
        .word = scanWord
        .exactMatch = False
    End With
    Exit Function
        
    word = scanWord
        
    countOfMatchedLetters = 0
    currentRunOfMatch = 0
    
    firstSpotOfLetterMatch = -1
    firstSpotOfMatch = -1
    maxLengthOfMatch = -1
    currentMatchSpotOfSearch = 1
            
    wordLength = Len(scanWord)
    
    If searchWord = scanWord Then
        With scanResults
            .lengthOfWord = wordLength
            .match_countOfMatchedLetters = wordLength
            .match_firstSpotOfLetterMatch = 1
            .match_firstSpotOfMatch = 1
            .match_maxLenthOfMatch = wordLength
            .word = word
            .exactMatch = True
            Exit Function
        End With
    End If
    
    
    For i = 1 To wordLength
        currLetterSearch = Mid(searchWord, currentMatchSpotOfSearch, 1)
        currLetterScan = Mid(word, i, 1)
        If currLetterScan = currLetterSearch Then
            countOfMatchedLetters = countOfMatchedLetters + 1
            currentRunOfMatch = currentRunOfMatch + 1
            
            If firstSpotOfLetterMatch = -1 Then
                firstSpotOfLetterMatch = i
            End If
            
            If currentRunOfMatch > maxLengthOfMatch Then
                maxLenthOfMatch = currentRunOfMatch
                firstSpotOfMatch = i - (currentRunOfMatch - 1)
            End If
            
            currentMatchSpotOfSearch = currentMatchSpotOfSearch + 1
        Else
            For j = 1 To Len(searchWord)
                If Mid(searchWord, j, 1) = currLetterScan Then
                    If firstSpotOfLetterMatch = -1 Then
                        firstSpotOfLetterMatch = i
                        firstSpotOfMatch = i
                        maxLenthOfMatch = 1
                    End If
                    countOfMatchedLetters = countOfMatchedLetters + 1
                    Exit For
                End If
            Next j
            currentRunOfMatch = 0
            currentMatchSpotOfSearch = 1
        End If
    Next i
    
    With scanResults
        .lengthOfWord = wordLength
        .match_countOfMatchedLetters = countOfMatchedLetters
        .match_firstSpotOfLetterMatch = firstSpotOfLetterMatch
        .match_firstSpotOfMatch = firstSpotOfMatch
        .match_maxLenthOfMatch = maxLenthOfMatch
        .word = word
        .exactMatch = False
    End With
    
End Function


Public Function parseColumns(selectLocation As qr_nextWord, fromLocation As qr_nextWord)
    Dim columnText
    Dim columnArray As Variant, currColumnArray As Variant
    Dim i, j, x, y
    
    columnText = Trim(Mid(qryReader.searchText, selectLocation.wordStart + 6, fromLocation.wordStart - selectLocation.wordStart - 6))
    
    If Mid(columnText, 1, 3) = "top" Then
        Dim numberStart
        Dim NumberEnd
        For i = 3 To Len(columnText)
            If Mid(columnText, i, 1) = " " Then
                numberStart = i + 1
                Exit For
            End If
        Next i
        
        For i = numberStart To Len(columnText)
            If Mid(columnText, i, 1) = " " Or Mid(columnText, i, 1) = "¦" Then
                NumberEnd = i
                Exit For
            End If
        Next i
        
        columnText = Trim(Mid(columnText, NumberEnd + 1, Len(columnText) - NumberEnd + 2))
    End If
    
    
    Dim operators_
    operators_ = getAllOpperators()
        
    For i = LBound(operators_) To UBound(operators_)
        If InStr(1, columnText, operators_(i)) > 0 Then
            If InStr(1, columnText, " " & operators_(i) & " ") = 0 Then
                If InStr(1, columnText, " " & operators_(i)) > 0 Then
                    columnText = Replace(columnText, " " & operators_(i), " " & operators_(i) & " ")
                End If
                
                If InStr(1, columnText, " " & operators_(i) & " ") = 0 Then
                    If InStr(1, columnText, operators_(i) & " ") > 0 Then
                        columnText = Replace(columnText, operators_(i) & " ", " " & operators_(i) & " ")
                    End If
                End If
                
                If InStr(1, columnText, " " & operators_(i) & " ") = 0 Then
                    If InStr(1, columnText, operators_(i)) > 0 Then
                        columnText = Replace(columnText, operators_(i), " " & operators_(i) & " ")
                    End If
                End If
            End If
        End If
    Next i
    
    Dim tempStr
    Do
        tempStr = columnText
        columnText = Replace(columnText, Space(2), Space(1))
    Loop Until (tempStr = columnText)
    
    operators_ = getJointOpperatorsCleanUp()
    For i = LBound(operators_) To UBound(operators_)
        If InStr(1, columnText, operators_(i)(0)) > 0 Then
            columnText = Replace(columnText, operators_(i)(0), operators_(i)(1))
        End If
    Next i
        
    columnArray = splitNotInBrackets(columnText, ",")
    If IsArray(columnArray) = True Then
        For i = LBound(columnArray) To UBound(columnArray)
            columnArray(i) = Replace(columnArray(i), "¦", "")
            columnArray(i) = Replace(columnArray(i), " as ", " ")
            columnArray(i) = Replace(columnArray(i), Space(2), Space(1))
            columnArray(i) = Trim(columnArray(i))
        Next i
    Else
        columnArray = Replace(columnArray, "¦", "")
        columnArray = Replace(columnArray, " as ", " ")
        columnArray = Replace(columnArray, Space(2), Space(1))
        columnArray = Trim(columnArray)
        columnArray = Array(columnArray)
    End If
    
    
    'TODO Check for operators before processing
    For i = LBound(columnArray) To UBound(columnArray)
        currColumnArray = splitNotInBrackets(columnArray(i), " ")
        If IsArray(currColumnArray) = True Then
            If UBound(currColumnArray) = 1 Then
                Dim inStringCheck
                inStringCheck = splitNotInBrackets(currColumnArray(0), ".")
                If IsArray(inStringCheck) Then
                    Call qryTree.addColumn(inStringCheck(1), currColumnArray(1), inStringCheck(0))
                Else
                    If InStr(1, inStringCheck, "(") > 0 Then
                        Call qryTree.addColumn(inStringCheck, currColumnArray(1), "FORMULA")
                    Else
                        Call qryTree.addColumn(inStringCheck, currColumnArray(1), "HOMETABLENOALIAS")
                    End If
                End If
            Else
                'TODO: Loop and process logic for column
                'Call qryTree.addColumn(currColumnArray(0), currColumnArray(1))
            End If
            
        Else
            currColumnArray = splitNotInBrackets(currColumnArray, ".")
            If IsArray(currColumnArray) = True Then
                Call qryTree.addColumn(currColumnArray(1), currColumnArray(1), currColumnArray(0))
            Else
                Call qryTree.addColumn(currColumnArray, currColumnArray, "HOMETABLENOALIAS")
            End If
        End If
    Next i
End Function

Private Function splitNotInBrackets(cText, del) As Variant
    Dim tempArr()
    Dim i, j, k
    Dim currLetter
    Dim lastColEnd
    Dim currSquareBracket
    Dim currBracket
    
    lastColEnd = 1
    currBracket = 0
    currSquareBracket = 0
    ReDim tempArr(-1 To 0)
    For i = 1 To Len(cText)
        currLetter = Mid(cText, i, 1)
        Select Case currLetter
            Case del
                If currBracket = 0 And currSquareBracket = 0 Then
                    If LBound(tempArr) = -1 Then
                        ReDim tempArr(0 To 0)
                        tempArr(0) = Trim(Mid(cText, lastColEnd, i - lastColEnd))
                    Else
                        ReDim Preserve tempArr(0 To UBound(tempArr) + 1)
                        tempArr(UBound(tempArr)) = Trim(Mid(cText, lastColEnd, i - lastColEnd))
                    End If
                    lastColEnd = i + 1
                End If
            Case "("
                currBracket = currBracket + 1
            Case ")"
                currBracket = currBracket - 1
            Case "["
                currSquareBracket = currSquareBracket + 1
            Case "]"
                currSquareBracket = currSquareBracket - 1
        End Select
    Next i
    
    If LBound(tempArr) = -1 Then
        splitNotInBrackets = cText
    Else
        ReDim Preserve tempArr(0 To UBound(tempArr) + 1)
        tempArr(UBound(tempArr)) = Mid(cText, lastColEnd, Len(cText) - lastColEnd + 1)
        splitNotInBrackets = tempArr
    End If
End Function

Public Function splitInBrackets(splitText As String, splitDelimeter As String, Optional splitStart As Integer = 1) As Variant
    Dim i, j
    Dim currLetter As String
    Dim lastSplitEnd As Integer
    
    Dim tempArray()
    Dim currentBracketBranch, currentSquareBracketBranch
        
    splitText = Replace(splitText, "¦", " ")
    splitText = Replace(splitText, Space(2), Space(1))
    splitText = Trim(splitText)
    
    If Len(splitText) = 1 Then
        splitInBrackets = splitText
        Exit Function
    End If
        
    If Mid(splitText, 1, 1) = "(" Then splitText = Mid(splitText, 2, Len(splitText) - 2)
    If Mid(splitText, Len(splitText) - 1, 1) = ")" Then splitText = Mid(splitText, 1, Len(splitText) - 2)
        
    currentBracketBranch = 0
    currentSquareBracketBranch = 0
    lastSplitEnd = 1
    ReDim tempArray(-1 To 0)
    
    For i = splitStart To Len(splitText) + 1
        currLetter = Mid(splitText, i, 1)
        Select Case currLetter
            Case splitDelimeter
                If currentBracketBranch = 0 And currentSquareBracketBranch = 0 Then
                    If LBound(tempArray) = -1 Then
                        ReDim tempArray(0 To 0)
                        tempArray(0) = Mid(splitText, lastSplitEnd, i - lastSplitEnd)
                    Else
                        ReDim Preserve tempArray(0 To UBound(tempArray) + 1)
                        tempArray(UBound(tempArray)) = Mid(splitText, lastSplitEnd, i - lastSplitEnd)
                    End If
                    lastSplitEnd = i + 1
                End If
            Case "("
                currentBracketBranch = currentBracketBranch + 1
            Case ")"
                currentBracketBranch = currentBracketBranch - 1
            Case "["
                currentSquareBracketBranch = currentSquareBracketBranch + 1
            Case "]"
                currentSquareBracketBranch = currentSquareBracketBranch - 1
        End Select
        
    Next i
        
    If LBound(tempArray) = -1 Then
        splitInBrackets = splitText
    Else
        ReDim Preserve tempArray(0 To UBound(tempArray) + 1)
        tempArray(UBound(tempArray)) = Mid(splitText, lastSplitEnd, Len(splitText) - lastSplitEnd + 1)
        splitInBrackets = tempArray
    End If
    
End Function

'TODO: FINISH THIS FUNCTION
Public Function getIntoData(intoWordLoc As qr_nextWord, columnRangeEndWord As qr_nextWord) As qr_getIntoData
    Dim i, j, x, y
    Dim scanText As String
    Dim tempResults As Variant
    
    scanText = Mid(qryReader.searchText, intoWordLoc.wordEnd + 1, columnRangeEndWord.wordStart - intoWordLoc.wordEnd - 1)
    
    Dim tempStr
    Do
        tempStr = scanText
        scanText = Replace(scanText, "¦", " ")
        scanText = Replace(scanText, "  ", " ")
        scanText = Trim(scanText)
    
    Loop Until (tempStr = scanText)
    'Temp Table
    If InStr(1, scanText, "#") > 0 Then
        If InStr(1, scanText, "(") > 0 Then
            tempResults = splitNotInBrackets(scanText, " ")
            
            Select Case UBound(tempResults)
                Case 2
                
                Case 1
                
                Case Else
                    For i = LBound(tempResults) To UBound(tempResults)
                        
                    Next i
                
            End Select
            
        Else
            tempResults = splitNotInBrackets(scanText, " ")
            If IsArray(tempResults) Then
                Select Case UBound(tempResults)
                    Case 1
                        With getIntoData
                            .intoType = "temp"
                            .isTempTable = True
                            .tableAlias = tempResults(1)
                            .tableName = tempResults(0)
                        End With
                    Case 0
                        With getIntoData
                            .intoType = "temp"
                            .isTempTable = True
                            .tableAlias = tempResults(0)
                            .tableName = tempResults(0)
                        End With
                    Case Else
                        'What happens now?!?
                        'For i = LBound(tempResults) To UBound(tempResults)
                        'Next i
                    End Select
            Else
                With getIntoData
                    .intoType = "temp"
                    .isTempTable = True
                    .tableAlias = tempResults
                    .tableName = tempResults
                End With
            End If

        End If
    Else
    
    End If
    
End Function

Public Function toJson(cls As cls_Query) As Variant
    Dim i As Integer, j
    Dim x, y
    Dim c, r
    Dim jsonText
    Dim currentJsonText
    Dim ws As Worksheet
    Dim Json(0 To 4)
    
    
    jsonText = "{"
    
    If cls.columns.count > 0 Then
        If cls.columns.count = 1 Then
            currentJsonText = "colsöcol_Aliasô" & IIf(cls.columns.column(0).alias <> "", cls.columns.column(0).alias, " ") & _
                              "æcol_Nameô" & IIf(cls.columns.column(0).column <> "", cls.columns.column(0).column, " ") & _
                              "æcol_Hostô" & IIf(cls.columns.column(0).table <> "", cls.columns.column(0).table, " ")
        Else
            currentJsonText = "colsö"
            For i = 0 To cls.columns.count - 1
                currentJsonText = currentJsonText & _
                                "col_Aliasô" & IIf(cls.columns.column(i).alias <> "", cls.columns.column(i).alias, " ") & _
                                "æcol_Nameô" & IIf(cls.columns.column(i).column <> "", cls.columns.column(i).column, " ") & _
                                "æcol_Hostô" & IIf(cls.columns.column(i).table <> "", cls.columns.column(i).table, " ") & "Æ"
            Next i
            currentJsonText = Mid(currentJsonText, 1, Len(currentJsonText) - 1)
            currentJsonText = currentJsonText
        End If
    Else
        currentJsonText = ""
    End If
    
    Json(0) = currentJsonText
    jsonText = jsonText & currentJsonText & "É"
    
    If cls.tables.count > 0 Then
        If cls.tables.count = 1 Then
            currentJsonText = "tblsötbl_Aliasô" & IIf(cls.tables.table(0).alias <> "", cls.tables.table(0).alias, " ") & _
                              "ætbl_Nameô" & IIf(cls.tables.table(0).name <> "", cls.tables.table(0).name, " ") & _
                              "ætbl_Typeô" & IIf(cls.tables.table(0).joinType <> "", cls.tables.table(0).joinType, " ") & _
                              "ætbl_HomeIndexô" & IIf(cls.tables.table(0).QueryID <> "", cls.tables.table(0).QueryID, " ")

        Else
            currentJsonText = "tblsö"
            For i = 0 To cls.tables.count - 1
                currentJsonText = currentJsonText & _
                                "tbl_Aliasô" & IIf(cls.tables.table(i).alias <> "", cls.tables.table(i).alias, " ") & _
                                "ætbl_Nameô" & IIf(cls.tables.table(i).name <> "", cls.tables.table(i).name, " ") & _
                                "ætbl_Typeô" & IIf(cls.tables.table(i).joinType <> "", cls.tables.table(i).joinType, " ") & _
                                "ætbl_HomeIndexô" & IIf(cls.tables.table(i).QueryID <> "", cls.tables.table(i).QueryID, " ") & "Æ"
            Next i
            currentJsonText = Mid(currentJsonText, 1, Len(currentJsonText) - 1)
            currentJsonText = currentJsonText
        End If
    Else
        currentJsonText = ""
    End If
    
    Json(1) = currentJsonText
    jsonText = jsonText & currentJsonText & "É"
    
    If cls.joins.count > 0 Then
        If cls.joins.count = 1 Then
            currentJsonText = "joinsöjoin_Typeô" & IIf(cls.joins.join(0).compareType <> "", cls.joins.join(0).compareType, " ") & _
                              "æjoin_Wordô" & IIf(cls.joins.join(0).compareWord <> "", cls.joins.join(0).compareWord, " ") & _
                              "æjoin_l_Aliasô" & IIf(cls.joins.join(0).left_SourceAlias <> "", cls.joins.join(0).left_SourceAlias, " ") & _
                              "æjoin_l_Clauseô" & IIf(cls.joins.join(0).left_SourceParamater <> "", cls.joins.join(0).left_SourceParamater, " ") & _
                              "æjoin_r_Aliasô" & IIf(cls.joins.join(0).right_SourceAlias <> "", cls.joins.join(0).right_SourceAlias, " ") & _
                              "æjoin_r_Clauseô" & IIf(cls.joins.join(0).right_SourceParamater <> "", cls.joins.join(0).right_SourceParamater, " ")

        Else
            currentJsonText = "joinsö"
            For i = 0 To cls.joins.count - 1
                currentJsonText = currentJsonText & _
                                "join_Typeô" & IIf(cls.joins.join(i).compareType <> "", cls.joins.join(i).compareType, " ") & _
                                "æjoin_Wordô" & IIf(cls.joins.join(i).compareWord <> "", cls.joins.join(i).compareWord, " ") & _
                                "æjoin_l_Aliasô" & IIf(cls.joins.join(i).left_SourceAlias <> "", cls.joins.join(i).left_SourceAlias, " ") & _
                                "æjoin_l_Clauseô" & IIf(cls.joins.join(i).left_SourceParamater <> "", cls.joins.join(i).left_SourceParamater, " ") & _
                                "æjoin_r_Aliasô" & IIf(cls.joins.join(i).right_SourceAlias <> "", cls.joins.join(i).right_SourceAlias, " ") & _
                                "æjoin_r_Clauseô" & IIf(cls.joins.join(i).right_SourceParamater <> "", cls.joins.join(i).right_SourceParamater, " ") & "Æ"
            Next i
            currentJsonText = Mid(currentJsonText, 1, Len(currentJsonText) - 1)
            currentJsonText = currentJsonText
        End If
    Else
        currentJsonText = ""
    End If
    
    Json(2) = currentJsonText
    jsonText = jsonText & currentJsonText & "É"
    
    If cls.wheres.count > 0 Then
        If cls.wheres.count = 1 Then
            currentJsonText = "wheresöwhere_Typeô" & IIf(cls.wheres.where(0).whereType <> "", cls.wheres.where(0).whereType, " ") & _
                              "æcompareWordô" & IIf(cls.wheres.where(0).compareWord <> "", cls.wheres.where(0).compareWord, " ") & _
                              "æcompareTypeô" & IIf(cls.wheres.where(0).compareType <> "", cls.wheres.where(0).compareType, " ") & _
                              "æwhere_l_Aliasô" & IIf(cls.wheres.where(0).Where_l_col_Alias <> "", cls.wheres.where(0).Where_l_col_Alias, " ") & _
                              "æwhere_l_Nameô" & IIf(cls.wheres.where(0).Where_l_col_Name <> "", cls.wheres.where(0).Where_l_col_Name, " ") & _
                              "æwhere_r_Aliasô" & IIf(cls.wheres.where(0).Where_r_col_Alias <> "", cls.wheres.where(0).Where_r_col_Alias, " ") & _
                              "æwhere_r_Nameô" & IIf(cls.wheres.where(0).Where_r_col_Name <> "", cls.wheres.where(0).Where_r_col_Name, " ") & _
                              "æbetween_col_Aliasô" & IIf(cls.wheres.where(0).between_col_Alias <> "", cls.wheres.where(0).between_col_Alias, " ") & _
                              "æbetween_col_Nameô" & IIf(cls.wheres.where(0).between_Col_Name <> "", cls.wheres.where(0).between_Col_Name, " ") & _
                              "æbetween_l_Clauseô" & IIf(cls.wheres.where(0).between_L_Compare <> "", cls.wheres.where(0).between_L_Compare, " ") & _
                              "æbetween_r_Clauseô" & IIf(cls.wheres.where(0).between_R_Compare <> "", cls.wheres.where(0).between_R_Compare, " ")

        Else
            currentJsonText = "wheresö"
            For i = 0 To cls.wheres.count - 1
                currentJsonText = currentJsonText & _
                              "where_Typeô" & IIf(cls.wheres.where(i).whereType <> "", cls.wheres.where(i).whereType, " ") & _
                              "æcompareWordô" & IIf(cls.wheres.where(i).compareWord <> "", cls.wheres.where(i).compareWord, " ") & _
                              "æcompareTypeô" & IIf(cls.wheres.where(i).compareType <> "", cls.wheres.where(i).compareType, " ") & _
                              "æwhere_l_Aliasô" & IIf(cls.wheres.where(i).Where_l_col_Alias <> "", cls.wheres.where(i).Where_l_col_Alias, " ") & _
                              "æwhere_l_Nameô" & IIf(cls.wheres.where(i).Where_l_col_Name <> "", cls.wheres.where(i).Where_l_col_Name, " ") & _
                              "æwhere_r_Aliasô" & IIf(cls.wheres.where(i).Where_r_col_Alias <> "", cls.wheres.where(i).Where_r_col_Alias, " ") & _
                              "æwhere_r_Nameô" & IIf(cls.wheres.where(i).Where_r_col_Name <> "", cls.wheres.where(i).Where_r_col_Name, " ") & _
                              "æbetween_col_Aliasô" & IIf(cls.wheres.where(i).between_col_Alias <> "", cls.wheres.where(i).between_col_Alias, " ") & _
                              "æbetween_col_Nameô" & IIf(cls.wheres.where(i).between_Col_Name <> "", cls.wheres.where(i).between_Col_Name, " ") & _
                              "æbetween_l_Clauseô" & IIf(cls.wheres.where(i).between_L_Compare <> "", cls.wheres.where(i).between_L_Compare, " ") & _
                              "æbetween_r_Clauseô" & IIf(cls.wheres.where(i).between_R_Compare <> "", cls.wheres.where(i).between_R_Compare, " ") & "Æ"
            Next i
            currentJsonText = Mid(currentJsonText, 1, Len(currentJsonText) - 1)
            currentJsonText = currentJsonText
        End If
    Else
        currentJsonText = ""
    End If
    
    Json(3) = currentJsonText
    jsonText = jsonText & currentJsonText & "É"
    
    toJson = Json
End Function

Public Function fromJson(jsonStr) As Variant
    Dim jS
    jS = jsonStr
    
    If jS = "" Then
        fromJson = Array("empty", " ")
        Exit Function
    End If
    
    Dim a, b, c, d(), i As Integer, j As Integer
    Dim cols As cls_Columns
    Dim tbls As cls_Tables
    Dim joins As cls_Joins
    Dim wheres As cls_Wheres
    
    a = Split(jS, "ö")(0)
    
    jS = Split(jS, "ö")(1)
    
    Select Case a
        Case "cols"
            Set cols = New cls_Columns
            b = Split(jS, "Æ")
            ReDim d(0 To UBound(b))
            
            For i = 0 To UBound(b)
                c = Split(b(i), "æ")
                'Call cols.add(Split(c(0), "ô")(1), Split(c(1), "ô")(1), Split(c(2), "ô")(1))
                d(i) = Array(Split(c(0), "ô")(1), Split(c(1), "ô")(1), Split(c(2), "ô")(1))
            Next i
            'fromJson = Array("cols", cols)
            fromJson = Array("cols", d)
            
        Case "tbls"
            Set tbls = New cls_Tables
            b = Split(jS, "Æ")
            ReDim d(0 To UBound(b))
            
            For i = 0 To UBound(b)
                c = Split(b(i), "æ")
                'Call tbls.add(Split(c(1), "ô")(1), Split(c(0), "ô")(1), Split(c(3), "ô")(1), Split(c(3), "ô")(1), Split(c(2), "ô")(1))
                d(i) = Array(Split(c(1), "ô")(1), Split(c(0), "ô")(1), Split(c(3), "ô")(1), Split(c(3), "ô")(1), Split(c(2), "ô")(1))
            Next i
            'fromJson = Array("tbls", tbls)
            fromJson = Array("tbls", d)
            
        Case "joins"
            Set joins = New cls_Joins
            b = Split(jS, "Æ")
            ReDim d(0 To UBound(b))
            
            For i = 0 To UBound(b)
                c = Split(b(i), "æ")
                'Call joins.addJoinOnData(Split(c(0), "ô")(1), Split(c(3), "ô")(1), Split(c(2), "ô")(1), Split(c(1), "ô")(1), Split(c(5), "ô")(1), Split(c(4), "ô")(1))
                d(i) = Array(Split(c(0), "ô")(1), Split(c(3), "ô")(1), Split(c(2), "ô")(1), Split(c(1), "ô")(1), Split(c(5), "ô")(1), Split(c(4), "ô")(1))
            Next i
            'fromJson = Array("joins", joins)
            fromJson = Array("joins", d)
            
        Case "wheres"
            Set wheres = New cls_Wheres
            b = Split(jS, "Æ")
            ReDim d(0 To UBound(b))
            
            For i = 0 To UBound(b)
                c = Split(b(i), "æ")
                If Split(c(2), "ô")(1) = "between" Then
                    'Call wheres.addWhereBetweenData("between", Split(c(6), "ô")(1), Split(c(7), "ô")(1), Split(c(8), "ô")(1), Split(c(8), "ô")(1), Split(c(9), "ô")(1))
                    d(i) = Array("between", Split(c(6), "ô")(1), Split(c(7), "ô")(1), Split(c(8), "ô")(1), Split(c(8), "ô")(1), Split(c(9), "ô")(1))
                Else
                    'Call wheres.addWhereData("where", Split(c(4), "ô")(1), Split(c(3), "ô")(1), Split(c(1), "ô")(1), Split(c(6), "ô")(1), Split(c(5), "ô")(1))
                    d(i) = Array("where", Split(c(4), "ô")(1), Split(c(3), "ô")(1), Split(c(1), "ô")(1), Split(c(6), "ô")(1), Split(c(5), "ô")(1))
                End If
            Next i
            'fromJson = Array("wheres", wheres)
            fromJson = Array("wheres", d)
            
    End Select
    fromJson = fromJson
End Function

Public Function scanJsonForTable(jsonData, searchWord)
    Dim i, j
    searchWord = LCase(searchWord)
    
    If IsArray(jsonData) = False Then
        If jsonData = " " Then
            scanJsonForTable = Array(Empty)
            Exit Function
        End If
        If jsonData = searchWord Then
            scanJsonForTable = Array(searchWord, searchWord, "", "")
            Exit Function
        Else
            scanJsonForTable = Array(Empty)
            Exit Function
        End If
    End If
    For i = 0 To UBound(jsonData)
        If jsonData(i)(1) = searchWord Then
            scanJsonForTable = jsonData(i)
            Exit Function
        End If
    Next i
    
    For i = 0 To UBound(jsonData)
        If jsonData(i)(0) = searchWord Then
            scanJsonForTable = jsonData(i)
            Exit Function
        End If
    Next i
    
    scanJsonForTable = Array(Empty)
End Function

Public Function getAvrgTimes(timeData)
    Dim timeT(0 To 9)
    Dim i, j
    Dim totalCount, totalTimeAvg, totalTime
    totalCount = UBound(timeData)
    For i = 0 To totalCount
        For j = 0 To 9
            timeT(j) = timeT(j) + timeData(i)(j)
        Next j
    Next i
    totalTime = 0
    For j = 0 To 9
        totalTime = totalTime + timeT(j)
        timeT(j) = Round(timeT(j) / (totalCount + 1), 8)
    Next j
    
    'debug.Print "0: " & timeT(0) & " | " & _
                "1: " & timeT(1) & " | " & _
                "2: " & timeT(2) & " | " & _
                "3: " & timeT(3) & " | " & _
                "4: " & timeT(4) & " | " & _
                "5: " & timeT(5) & " | " & _
                "6: " & timeT(6) & " | " & _
                "7: " & timeT(7) & " | " & _
                "8: " & timeT(8) & " | " & _
                "9: " & timeT(9)
    'debug.Print "0: " & timeT(0) * (totalCount + 1) & " | " & _
                "1: " & timeT(1) * (totalCount + 1) & " | " & _
                "2: " & timeT(2) * (totalCount + 1) & " | " & _
                "3: " & timeT(3) * (totalCount + 1) & " | " & _
                "4: " & timeT(4) * (totalCount + 1) & " | " & _
                "5: " & timeT(5) * (totalCount + 1) & " | " & _
                "6: " & timeT(6) * (totalCount + 1) & " | " & _
                "7: " & timeT(7) * (totalCount + 1) & " | " & _
                "8: " & timeT(8) * (totalCount + 1) & " | " & _
                "9: " & timeT(9) * (totalCount + 1)
    'debug.Print totalTime
End Function

Public Function searchForTable(tableName As String, jsonData) As qr_tableMainSearchData()
    Dim x, y, j, k
    Dim currTable
    Dim currJson
    Dim results() As qr_tableMainSearchData
    Dim finalResults() As qr_tableMainSearchData
    Dim resultsCount, finalResultsCount
    Dim searchResults
    Dim tempLetter
    Dim maxLengthFind
    Dim modAmount
    
    y = y
    ReDim results(0 To 99999)
    resultsCount = 0
    maxLengthFind = 0
    modAmount = UBound(jsonData)
    For j = 0 To modAmount
        'If j Mod (modAmount \ 10) = 0 Then
        '    If j > 0 Then
        '        debugProgress (Round(j / modAmount * 10))
        '    Else
        '        debugProgress (0)
        '    End If
        'End If
        currJson = jsonData(j)
        If IsArray(currJson) = True Then
            If IsArray(currJson(1)) = True Then
                For k = 0 To UBound(currJson(1))
                    If InStr(1, currJson(1)(k)(0), Mid(tableName, 1, 1)) > 0 Then
                        currJson(1)(k)(0) = Replace(currJson(1)(k)(0), "[", "")
                        currJson(1)(k)(0) = Replace(currJson(1)(k)(0), "]", "")
                        currJson(1)(k)(0) = Replace(currJson(1)(k)(0), "rpt.", "")
                        currJson(1)(k)(0) = Replace(currJson(1)(k)(0), "dbo.", "")
                        currJson(1)(k)(0) = Replace(currJson(1)(k)(0), "integra.", "")
                        
                        currJson(1)(k)(1) = Replace(currJson(1)(k)(1), "[", "")
                        currJson(1)(k)(1) = Replace(currJson(1)(k)(1), "]", "")
                        currJson(1)(k)(1) = Replace(currJson(1)(k)(1), "rpt.", "")
                        currJson(1)(k)(1) = Replace(currJson(1)(k)(1), "dbo.", "")
                        currJson(1)(k)(1) = Replace(currJson(1)(k)(1), "integra.", "")
                        
                        Call getScanResults(tableName, currJson(1)(k)(0), results(resultsCount).tableData.tableSearchResults)
                        Call getScanResults(tableName, currJson(1)(k)(1), results(resultsCount).tableData.aliasSearchResults)
                        If results(resultsCount).tableData.tableSearchResults.match_maxLenthOfMatch > maxLengthFind Then
                            maxLengthFind = results(resultsCount).tableData.tableSearchResults.match_maxLenthOfMatch
                        End If
                        results(resultsCount).tableData.homeQueryIndex = j
                    End If
                    resultsCount = resultsCount + 1
                Next k
            End If
        End If
    Next j
    
    ReDim finalResults(0 To resultsCount - 1)
    
    finalResultsCount = 0
    modAmount = UBound(finalResults)
    For j = 0 To modAmount
'        If j Mod (modAmount \ 10) = 0 Then
'            If j > 0 Then
'                debugProgress (10 + Round(j / modAmount * 100))
'            End If
'        End If
        If results(j).tableData.tableSearchResults.match_maxLenthOfMatch = maxLengthFind And InStr(1, results(j).tableData.tableSearchResults.word, "(") = 0 Then
            finalResults(finalResultsCount) = results(j)
            finalResults(finalResultsCount).count = 1
            finalResultsCount = finalResultsCount + 1
        End If
    Next j
    If finalResultsCount > 0 Then
        ReDim Preserve finalResults(0 To finalResultsCount - 1)
    Else
        ReDim finalResults(-1 To 0)
    End If
    searchForTable = finalResults
    x = x
End Function

Public Function combineTblSearchResults(ByRef tableData() As qr_tableMainSearchData)
    Dim x, y, i, j
    Dim init As Boolean
    Dim currentWord
    Dim resultCount
    Dim foundResult As Boolean
    Dim aliasFound As Boolean
    Dim endResult() As qr_tableMainSearchData

    init = False
    
    For i = 0 To UBound(tableData)
        If i = 25 Then
            i = i
        End If
        If init = False Then
            ReDim endResult(0 To 0)
            endResult(0) = tableData(i)
            ReDim endResult(0).tableData.aliasList(0 To 0)
            endResult(0).tableData.aliasList(0).alias = endResult(0).tableData.aliasSearchResults.word
            endResult(0).tableData.aliasList(0).count = 1
            endResult(0).tableData.aliasList(0).homeQueryIndex = tableData(i).tableData.homeQueryIndex
            init = True
            resultCount = 0
        Else
            foundResult = False
            For j = 0 To resultCount
                If endResult(j).tableData.tableSearchResults.word = tableData(i).tableData.tableSearchResults.word Then
                    aliasFound = False
                    For x = 0 To UBound(endResult(j).tableData.aliasList)
                        If endResult(j).tableData.aliasList(x).alias = tableData(i).tableData.aliasSearchResults.word Then
                            endResult(j).tableData.aliasList(x).count = endResult(j).tableData.aliasList(x).count + 1
                            endResult(j).tableData.aliasList(x).homeQueryIndex = endResult(j).tableData.aliasList(x).homeQueryIndex & "," & tableData(i).tableData.homeQueryIndex
                            aliasFound = True
                            Exit For
                            'TODO FINISH ADDING ALIASLIST - Maybe finished?
                        End If
                    Next x
                    
                    If aliasFound = False Then
                        ReDim Preserve endResult(j).tableData.aliasList(0 To UBound(endResult(j).tableData.aliasList) + 1)
                        endResult(j).tableData.aliasList(UBound(endResult(j).tableData.aliasList)).alias = tableData(i).tableData.aliasSearchResults.word
                        endResult(j).tableData.aliasList(UBound(endResult(j).tableData.aliasList)).count = endResult(j).tableData.aliasList(UBound(endResult(j).tableData.aliasList)).count + 1
                        endResult(j).tableData.aliasList(UBound(endResult(j).tableData.aliasList)).homeQueryIndex = endResult(j).tableData.aliasList(UBound(endResult(j).tableData.aliasList)).homeQueryIndex & "," & tableData(i).tableData.homeQueryIndex
                    End If
                    
                    endResult(j).count = endResult(j).count + 1
                    foundResult = True
                    Exit For
                End If
            Next j
            
            If foundResult = False Then
                resultCount = resultCount + 1
                ReDim Preserve endResult(0 To resultCount)
                endResult(resultCount) = tableData(i)
                ReDim endResult(resultCount).tableData.aliasList(0 To 0)
                endResult(resultCount).tableData.aliasList(0).alias = tableData(i).tableData.aliasSearchResults.word
                endResult(resultCount).tableData.aliasList(0).count = 1
                endResult(resultCount).tableData.aliasList(0).homeQueryIndex = tableData(i).tableData.homeQueryIndex
            End If
        End If
    Next i
    
    init = True
    tableData = endResult
End Function

Public Function sortTblSeachResults(ByRef tableData() As qr_tableMainSearchData)
    Dim i, j, x, y
    Dim tempTbl As qr_tableMainSearchData
    Dim tempAlias As qr_tableAliasData
    Dim changed_

    If UBound(tableData) = 0 Then
        Exit Function
    End If

    Do
        changed_ = False
        For i = UBound(tableData) To 1 Step -1
            If tableData(i).tableData.tableSearchResults.exactMatch = True And tableData(i - 1).tableData.tableSearchResults.exactMatch = False Then
                tempTbl = tableData(i - 1)
                tableData(i - 1) = tableData(i)
                tableData(i) = tempTbl
                changed_ = True
            End If
        Next i
    Loop Until (changed_ = False)

    Do
        changed_ = False
        For i = UBound(tableData) To 1 Step -1
            If (tableData(i).tableData.tableSearchResults.exactMatch = tableData(i - 1).tableData.tableSearchResults.exactMatch) And (tableData(i).count > tableData(i - 1).count) Then
                tempTbl = tableData(i - 1)
                tableData(i - 1) = tableData(i)
                tableData(i) = tempTbl
                changed_ = True
            End If
        Next i
    Loop Until (changed_ = False)
    
    i = i
    
    Do
        changed_ = False
        For i = 0 To UBound(tableData)
            If UBound(tableData(i).tableData.aliasList) > 0 Then
                For j = UBound(tableData(i).tableData.aliasList) To 1 Step -1
                    If tableData(i).tableData.aliasList(j).count > tableData(i).tableData.aliasList(j - 1).count Then
                        tempAlias = tableData(i).tableData.aliasList(j - 1)
                        tableData(i).tableData.aliasList(j - 1) = tableData(i).tableData.aliasList(j)
                        tableData(i).tableData.aliasList(j) = tempAlias
                        changed_ = True
                    End If
                Next j
            End If
        Next i
    Loop Until (changed_ = False)
    
    i = i
End Function

Public Function outputTblSearchResults(searchWord, tableSearch() As qr_tableMainSearchData, sh_ As Worksheet, Optional columnOutPut As Integer = -1)
    Dim i, j
    Dim al, aliasString
    
    setOutPutArray
  
    If columnOutPut = -1 Then
        sh_.Cells.ClearContents
        sh_.Cells.ClearFormats
        sh_.Cells.Font.Color = RGB(0, 0, 0)
        
        Call outputDataToSheet(sh_, 1, "Search Table:", , "name:Consolas,fore:11892015,size:16,autofit:true,bold:true,align:right")
        Call outputDataToSheet(sh_, 1, searchWord, , "name:Consolas,size:14,autofit:true,bold:true")
            
        Call outputDataToSheet(sh_, 2, "Matching Tables:", , "size:12,align:right")
        Call outputDataToSheet(sh_, 2, IIf(LBound(tableSearch) > -1, UBound(tableSearch) + 1, 0), , "size:10,align:left")
    
        sh_.Rows(c_OutputHeaderRow).AutoFilter
        
        Call outputDataToSheet(sh_, c_OutputHeaderRow, "Tables Found", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        Call outputDataToSheet(sh_, c_OutputHeaderRow, "Table Alias('s)", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Num of Alias's", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Home Table Alias", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        Call outputDataToSheet(sh_, c_OutputHeaderRow, "Found Count", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Exact Match", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Match Streak", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Match Streak Location", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Letter Match Count", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Letter Match Location", , "bold:true,size:18,back:15917529,autofit:true,align:left")
        'Call outputDataToSheet(sh_, c_OutputHeaderRow, "Word Length", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    
        
        If LBound(tableSearch) > -1 Then
            For j = c_OutputHeaderRow + 1 To (UBound(tableSearch) - LBound(tableSearch) + (c_OutputHeaderRow + 1))
                Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.word, "'", "`"), , "autofit:true,align:left")
                sh_.Cells(j, 1).Font.Color = RGB(100, 100, 123)
                With sh_.Cells(j, 1).Characters(Start:=tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_firstSpotOfMatch, Length:=tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_maxLenthOfMatch).Font
                    .FontStyle = "bold"
                    .Color = RGB(10, 105, 0)
                    .Size = .Size + 2
                End With
                aliasString = ""
                For al = 0 To UBound(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList)
                    If tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList(al).alias <> "" Or tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList(al).alias <> " " Then
                        If aliasString = "" Then
                            aliasString = tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList(al).alias
                        Else
                            aliasString = aliasString & ", " & tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList(al).alias
                        End If
                    End If
                Next al
                If aliasString <> "" And Len(aliasString) >= 3 Then
                    aliasString = Mid(aliasString, 1, Len(aliasString) - 2)
                End If
                
                Call outputDataToSheet(sh_, j, Replace(aliasString, "'", "`"), , "align:left")
                'Call outputDataToSheet(sh_, j, UBound(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasList) + 1)
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.aliasSearchResults.word, "'", "`"))
                Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).count, "'", "`"), , "align:center")
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.exactMatch, "'", "`"))
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_maxLenthOfMatch, "'", "`"))
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_firstSpotOfMatch, "'", "`"))
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_countOfMatchedLetters, "'", "`"))
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_firstSpotOfLetterMatch, "'", "`"))
                'Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.lengthOfWord, "'", "`"))
            Next j
        End If
    Else
        Dim rng As Range
        Set rng = sh_.Range(sh_.Cells(c_OutputHeaderRow, columnOutPut - 1), sh_.Cells(sh_.Cells.Rows.count, columnOutPut))
        rng.ClearContents
        rng.ClearFormats
        rng.Font.Color = RGB(0, 0, 0)
        
        Call outputDataToSheet(sh_, 1, "Search Table:", columnOutPut - 1, "name:Consolas,fore:11892015,size:16,autofit:true,bold:true,align:right")
        Call outputDataToSheet(sh_, 1, searchWord, columnOutPut, "name:Consolas,size:14,autofit:true,bold:true")
            
        Call outputDataToSheet(sh_, 2, "Matching Tables:", columnOutPut - 1, "size:12,align:right")
        Call outputDataToSheet(sh_, 2, IIf(LBound(tableSearch) > -1, UBound(tableSearch) + 1, 0), columnOutPut, "size:10,align:left")
        
        Call outputDataToSheet(sh_, c_OutputHeaderRow, "Tables Found", columnOutPut - 1, "bold:true,size:18,back:15917529,autofit:true,align:center")
        Call outputDataToSheet(sh_, c_OutputHeaderRow, "Found Count", columnOutPut, "bold:true,size:18,back:15917529,autofit:true,align:center")
        
        If LBound(tableSearch) > -1 Then
            For j = c_OutputHeaderRow + 1 To (UBound(tableSearch) - LBound(tableSearch) + (c_OutputHeaderRow + 1))
                Call outputDataToSheet(sh_, j, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.word, "'", "`"), columnOutPut - 1, "align:left")
                Call outputDataToSheet(sh_, j, tableSearch(j - (c_OutputHeaderRow + 1)).count, columnOutPut, "align:center")
                sh_.Cells(j, columnOutPut - 1).Font.Color = RGB(100, 100, 123)
                With sh_.Cells(j, columnOutPut - 1).Characters(Start:=tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_firstSpotOfMatch, Length:=tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.match_maxLenthOfMatch).Font
                    .FontStyle = "bold"
                    .Color = RGB(10, 105, 0)
                    .Size = .Size + 2
                End With
            Next j
        End If
    End If
End Function

Public Function getColumnJsonSearch(searchWord As String) As qr_columnMainSearchData()
    If IsEmpty(p_JsonData) = True Then p_JsonData = parseAllJson
    Dim columnSearch() As qr_columnMainSearchData
    columnSearch = searchForColumn(searchWord, jsond)
    Call combineColSearchResults(columnSearch)
    Call sortColSeachResults(columnSearch)
    p_LastColumnSearch = columnSearch
    getColumnJsonSearch = columnSearch
End Function

Public Function getTableJsonSearch(searchWord As String) As qr_tableMainSearchData()
    If IsEmpty(p_JsonData) = True Then p_JsonData = parseAllJson
    Dim tableSearch() As qr_tableMainSearchData
    tableSearch = searchForTable(searchWord, p_JsonData)
    Call combineTblSearchResults(tableSearch)
    Call sortTblSeachResults(tableSearch)
    p_LastTableSearch = tableSearch
    getTableJsonSearch = tableSearch
End Function

Public Function outputJoinTblSearchResults(searchWord, tableSearch() As qr_tableMainSearchData, sh_ As Worksheet)
    Dim i, j
    Dim al, aliasString
    
    setOutPutArray
  
    sh_.Cells.ClearContents
    sh_.Cells.ClearFormats
    sh_.Cells.Font.Color = RGB(0, 0, 0)
    
    Call outputDataToSheet(sh_, 1, "Search Table:", , "name:Consolas,fore:11892015,size:16,autofit:true,bold:true,align:right")
    Call outputDataToSheet(sh_, 1, searchWord, , "name:Consolas,size:14,autofit:true,bold:true,align:left")
    
    Call outputDataToSheet(sh_, 2, "Matching Tables:", , "size:12,align:right")
    Call outputDataToSheet(sh_, 2, IIf(LBound(tableSearch) > -1, UBound(tableSearch) + 1, 0), , "size:10,align:left")
    
    
    Call outputDataToSheet(sh_, c_OutputHeaderRow, "Tables Found:", , "bold:true,size:18,back:15917529,autofit:true,align:right")
    Call outputDataToSheet(sh_, c_OutputHeaderRow + 1, "Found Amount:", , "bold:true,size:18,back:15917529,autofit:true,align:right")
        
    If LBound(tableSearch) > -1 Then
        For j = c_OutputHeaderRow + 1 To UBound(tableSearch) + c_OutputHeaderRow + 1
            Call outputDataToSheet(sh_, c_OutputHeaderRow, Replace(tableSearch(j - (c_OutputHeaderRow + 1)).tableData.tableSearchResults.word, "'", "`"), , "align:center,autofit:true")
            Call outputDataToSheet(sh_, c_OutputHeaderRow + 1, tableSearch(j - (c_OutputHeaderRow + 1)).count, , "align:center")
        Next j
    End If
    'sh_.Rows(c_OutputHeaderRow).AutoFilter
End Function

Public Function getJoinsJsonSearch(searchData As qr_tableMainSearchData) As qr_JoinFinalResults
    If IsEmpty(p_JsonData) = True Then p_JsonData = parseAllJson
    Dim joinSearch As qr_JoinMainSearchData
    Dim finalSearch As qr_JoinFinalResults
    joinSearch = SearchForJoins(searchData)
    finalSearch = combineJoinSearchResults(joinSearch)
    finalSearch = sortJoinSeachResults(finalSearch)
    p_LastJoinSearch = joinSearch
    getJoinsJsonSearch = finalSearch
End Function

Public Function SearchForJoins(searchData As qr_tableMainSearchData) As qr_JoinMainSearchData
    Dim x, y, i, j
    Dim tempJoins As qr_JoinMainSearchData
    Dim joinCount As Long
    Dim tempQuery_
    Dim findAlias As fn_inArrayResult
    ReDim tempJoins.joins(0 To 999999)
    joinCount = 0
    tempJoins.tableData = searchData.tableData
        
    With searchData.tableData
        For i = 0 To UBound(.aliasList)
            If Mid(.aliasList(i).homeQueryIndex, 1, 1) = "," Then .aliasList(i).homeQueryIndex = Mid(.aliasList(i).homeQueryIndex, 2, Len(.aliasList(i).homeQueryIndex) - 1)
            For Each j In Split(.aliasList(i).homeQueryIndex, ",")
                tempQuery_ = p_JsonData(j)
                findAlias = inArray(.aliasList(i).alias, tempQuery_(1), 1)
                If findAlias.found = True Then
                    If IsArray(tempQuery_(2)) = True Then
                        For x = 0 To UBound(tempQuery_(2))
                            If tempQuery_(2)(x)(2) = .aliasList(i).alias Or tempQuery_(2)(x)(5) = .aliasList(i).alias Then
                                If tempQuery_(2)(x)(2) = .aliasList(i).alias Then
                                    tempJoins.joins(joinCount).tableName = getTableNameFromJsonByAlias(tempQuery_(2)(x)(5), j)
                                    ReDim tempJoins.joins(joinCount).onColumns(0 To 0)
                                    tempJoins.joins(joinCount).onColumns(0).columnNameL = tempQuery_(2)(x)(1)
                                    tempJoins.joins(joinCount).onColumns(0).columnNameR = tempQuery_(2)(x)(4)
                                    tempJoins.joins(joinCount).onColumns(0).joinType = tempQuery_(2)(x)(3)
                                Else
                                    tempJoins.joins(joinCount).tableName = getTableNameFromJsonByAlias(tempQuery_(2)(x)(2), j)
                                    ReDim tempJoins.joins(joinCount).onColumns(0 To 0)
                                    tempJoins.joins(joinCount).onColumns(0).columnNameL = tempQuery_(2)(x)(4)
                                    tempJoins.joins(joinCount).onColumns(0).columnNameR = tempQuery_(2)(x)(1)
                                    tempJoins.joins(joinCount).onColumns(0).joinType = tempQuery_(2)(x)(3)
                                End If
                                joinCount = joinCount + 1
                            End If
                        Next x
                    End If
                End If
            Next j
            j = j
            'TODO scan for joins that involve searchTable
        Next i
        i = i
    End With
    If joinCount = 0 Then
        ReDim tempJoins.joins(-1 To 0)
    Else
        ReDim Preserve tempJoins.joins(0 To joinCount - 1)
    End If
    SearchForJoins = tempJoins
End Function

Public Function getTableNameFromJsonByAlias(tableAlias, homeQueryIndex)
    Dim curQuery, returnResult
    Dim i, j, t
    
    If homeQueryIndex > UBound(p_JsonData) Then Exit Function
    
    curQuery = p_JsonData(homeQueryIndex)
    
    For Each t In curQuery(1)
        If t(1) = tableAlias Then
            tableAlias = t(0)
            tableAlias = Replace(tableAlias, "[", "")
            tableAlias = Replace(tableAlias, "]", "")
            tableAlias = Replace(tableAlias, "rpt.", "")
            tableAlias = Replace(tableAlias, "dbo.", "")
            tableAlias = Replace(tableAlias, "integra.", "")
            getTableNameFromJsonByAlias = tableAlias
            Exit Function
        End If
    Next t
    tableAlias = Replace(tableAlias, "[", "")
    tableAlias = Replace(tableAlias, "]", "")
    tableAlias = Replace(tableAlias, "rpt.", "")
    tableAlias = Replace(tableAlias, "dbo.", "")
    tableAlias = Replace(tableAlias, "integra.", "")
    getTableNameFromJsonByAlias = tableAlias
End Function

Public Function combineJoinSearchResults(ByRef joinData As qr_JoinMainSearchData) As qr_JoinFinalResults
    Dim i, j, x, y
    Dim tempJoin As qr_JoinPackage
    Dim finalResult As qr_JoinFinalResults
    Dim finalCount As Long
    Dim foundTable As Boolean
    Dim foundColumn As Boolean
        
    ReDim finalResult.joinedTables(0 To 999999)
    finalResult.mainTable = joinData.tableData
    finalCount = -1
    
    If LBound(joinData.joins) > -1 Then
        For i = 0 To UBound(joinData.joins)
            If joinData.joins(i).onColumns(0).columnNameR <> "" Then
                tempJoin = joinData.joins(i)
                If finalCount = -1 Then
                    finalCount = finalCount + 1
                    finalResult.joinedTables(finalCount).tableName = joinData.joins(i).tableName
                    ReDim finalResult.joinedTables(finalCount).onColumns(0 To 0)
                    finalResult.joinedTables(finalCount).onColumns(0) = joinData.joins(i).onColumns(0)
                    finalResult.joinedTables(finalCount).count = 1
                Else
                    foundTable = False
                    foundColumn = False
                    For x = 0 To finalCount
                        If finalResult.joinedTables(x).tableName = joinData.joins(i).tableName Then
                            finalResult.joinedTables(x).count = finalResult.joinedTables(x).count + 1
                            For y = 0 To UBound(finalResult.joinedTables(x).onColumns)
                                If finalResult.joinedTables(x).onColumns(y).columnNameL = joinData.joins(i).onColumns(0).columnNameL And finalResult.joinedTables(x).onColumns(y).columnNameR = joinData.joins(i).onColumns(0).columnNameR And finalResult.joinedTables(x).onColumns(y).joinType = joinData.joins(i).onColumns(0).joinType Then
                                    finalResult.joinedTables(x).onColumns(y).count = finalResult.joinedTables(x).onColumns(y).count + 1
                                    foundColumn = True
                                    Exit For
                                End If
                            Next y
                            
                            If foundColumn = False Then
                                Dim colCount: colCount = UBound(finalResult.joinedTables(x).onColumns) + 1
                                ReDim Preserve finalResult.joinedTables(x).onColumns(0 To colCount)
                                finalResult.joinedTables(x).onColumns(colCount).columnNameL = joinData.joins(i).onColumns(0).columnNameL
                                finalResult.joinedTables(x).onColumns(colCount).columnNameR = joinData.joins(i).onColumns(0).columnNameR
                                finalResult.joinedTables(x).onColumns(colCount).joinType = IIf(joinData.joins(i).onColumns(0).joinType = "", "=", joinData.joins(i).onColumns(0).joinType)
                                finalResult.joinedTables(x).onColumns(colCount).count = 1
                            End If
                            
                            foundTable = True
                            Exit For
                        End If
                    Next x
                    
                    If foundTable = False Then
                        finalCount = finalCount + 1
                        finalResult.joinedTables(finalCount).tableName = joinData.joins(i).tableName
                        ReDim Preserve finalResult.joinedTables(finalCount).onColumns(0 To 0)
                        finalResult.joinedTables(finalCount).onColumns(0).columnNameL = joinData.joins(i).onColumns(0).columnNameL
                        finalResult.joinedTables(finalCount).onColumns(0).columnNameR = joinData.joins(i).onColumns(0).columnNameR
                        finalResult.joinedTables(finalCount).onColumns(0).joinType = IIf(joinData.joins(i).onColumns(0).joinType = "", "=", joinData.joins(i).onColumns(0).joinType)
                        finalResult.joinedTables(finalCount).onColumns(0).count = 1
                        finalResult.joinedTables(finalCount).count = 1
                    End If
                End If
            End If
        Next i
        ReDim Preserve finalResult.joinedTables(0 To finalCount)
    Else
        ReDim finalResult.joinedTables(-1 To 0)
    End If
    combineJoinSearchResults = finalResult
End Function

Public Function sortJoinSeachResults(tableData As qr_JoinFinalResults) As qr_JoinFinalResults
    Dim i, j, x, y
    Dim tempV As qr_JoinPackage
    Dim tempW As jn_JoinColumnData
    Dim curV
    Dim changed_ As Boolean
    If LBound(tableData.joinedTables) > -1 Then
        Do
            changed_ = False
            For i = UBound(tableData.joinedTables) To 1 Step -1
                If tableData.joinedTables(i).count > tableData.joinedTables(i - 1).count Then
                    tempV = tableData.joinedTables(i - 1)
                    tableData.joinedTables(i - 1) = tableData.joinedTables(i)
                    tableData.joinedTables(i) = tempV
                    changed_ = True
                End If
            Next i
        Loop Until (changed_ = False)
        
        For x = 0 To UBound(tableData.joinedTables)
            With tableData.joinedTables(x)
                Do
                    changed_ = False
                    For i = UBound(.onColumns) To 1 Step -1
                        If .onColumns(i).count > .onColumns(i - 1).count Then
                            tempW = .onColumns(i - 1)
                            .onColumns(i - 1) = .onColumns(i)
                            .onColumns(i) = tempW
                            changed_ = True
                        End If
                    Next i
                Loop Until (changed_ = False)
            End With
        Next x
    End If
    sortJoinSeachResults = tableData
End Function

Public Function outputJoinJoinsSearchResults(joinSearch As qr_JoinFinalResults, sh_ As Worksheet)
    Dim i, j
    Dim al, aliasString
    
    setOutPutArray
  
    sh_.Range(sh_.Cells(c_OutputHeaderRow + 3, 1), sh_.Cells(sh_.Rows.count, sh_.columns.count)).ClearContents
    'sh_.Cells.ClearFormats
    'sh_.Cells.Font.Color = RGB(0, 0, 0)
    
    Call outputDataToSheet(sh_, c_OutputHeaderRow + 3, joinSearch.mainTable.tableSearchResults.word)
    Call outputDataToSheet(sh_, c_OutputHeaderRow + 4, "Joined On Tables:", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    Call outputDataToSheet(sh_, c_OutputHeaderRow + 4, "Joined Amount", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    Call outputDataToSheet(sh_, c_OutputHeaderRow + 4, "Joined on columns:", , "bold:true,size:18,back:15917529,autofit:true,align:left")
    
    If LBound(joinSearch.joinedTables) > -1 Then
        For j = c_OutputHeaderRow + 5 To UBound(joinSearch.joinedTables) + c_OutputHeaderRow + 5
            Call outputDataToSheet(sh_, j, joinSearch.joinedTables(j - (c_OutputHeaderRow + 5)).tableName, , "align:left")
            Call outputDataToSheet(sh_, j, joinSearch.joinedTables(j - (c_OutputHeaderRow + 5)).count, , "align:center")
            With joinSearch.joinedTables(j - (c_OutputHeaderRow + 5))
                For i = 0 To UBound(.onColumns)
                    Call outputDataToSheet(sh_, j, .onColumns(i).columnNameL & " " & .onColumns(i).joinType & " " & .onColumns(i).columnNameR & " (count: " & .onColumns(i).count & ")", , "autofit:true,align:left")
                Next i
            End With
        Next j
    End If
    'sh_.Rows(c_OutputHeaderRow).AutoFilter
End Function


'TODO: CURRENT PROJECT!!
Public Function searchJsonJoinsBranch(sh As Worksheet, l_selected, r_selected)
    Dim a, b, c  As dict_SubBranchResult, d As JoinBranchStructure, e, f, g, h
    Dim i As JoinBranchStructureSubBranch
    Dim j As SubBranchTables
    Dim x, y, xx, yy
    Dim j1() As dict_TableJoins, j2() As dict_ColumnJoins, j3, j4
    Dim output_
    Const startCol = 8
    
    debugProgress 20, , "Scanning though joins for connection"
    
    a = sh.Cells(l_selected, 1).Value
    b = sh.Cells(r_selected, 5).Value
    
    If a = "" Or b = "" Then Exit Function
    If IsEmpty(p_JsonData) = True Then p_JsonData = parseAllJson

    Set d = New JoinBranchStructure
    Call d.createBranch(a, b)
    
    c = d.getBranch
    
    With sh.Range(sh.Cells(1, startCol), sh.Cells(sh.Rows.count, sh.columns.count))
        .ClearContents
        .ClearFormats
    End With
    
    If c.finalString <> "" Then
        e = Split(c.finalString, " -> ")
        ReDim c.arr(0 To UBound(e))
        For x = 0 To UBound(e)
            c.arr(x) = d.getDictElement(e(x))
        Next x
        
        sh.Cells(c_OutputHeaderRow + 1, startCol).Value = a
        sh.columns(12).AutoFit
        
        With sh.Cells(c_OutputHeaderRow + 1, startCol + 1)
            .Value = "to"
            .HorizontalAlignment = xlCenter
        End With
        sh.columns(13).AutoFit
        
        sh.Cells(c_OutputHeaderRow + 1, startCol + 2).Value = b
        sh.columns(14).AutoFit
        Dim rowCnt
        
        For f = 0 To UBound(e) - 1
            sh.Cells(c_OutputHeaderRow + 2, startCol + (2 * f)).Value = e(f)
            sh.Cells(c_OutputHeaderRow + 2, startCol + (2 * f)).HorizontalAlignment = xlRight
            sh.Cells(c_OutputHeaderRow + 2, startCol + (2 * f) + 1).Value = e(f + 1)
            'For x = 0 To UBound(c.Arr)
                j1 = c.arr(f).joined_Tables
                For xx = 0 To UBound(j1)
                    If j1(xx).joined_Table_Name = e(f + 1) Then
                        foundTbl_ = True
                        j2 = j1(xx).joined_On
                        For y = 0 To UBound(j2)
                            With j2(y)
                                sh.Cells(c_OutputHeaderRow + 3 + y, startCol + (f * 2)).Value = .joinet_Table_Column & " " & .join_Comparison
                                sh.Cells(c_OutputHeaderRow + 3 + y, startCol + (f * 2)).HorizontalAlignment = xlRight
                                sh.Cells(c_OutputHeaderRow + 3 + y, startCol + (f * 2) + 1).Value = .joined_Table_Column
                                sh.Cells(c_OutputHeaderRow + 3 + y, startCol + (f * 2) + 1).HorizontalAlignment = xlLeft
                            End With
                        Next y
                        Exit For
                    End If
                Next xx
            'Next x
            sh.columns(startCol + (2 * f)).AutoFit
            sh.columns(startCol + (2 * f) + 1).AutoFit
'            rowCnt = -1
'            For h = c.objArrCnt To 1 Step -1
'            rowCnt = rowCnt + 1
'                Debug.Print TypeName(c.objArr(h))
'                If TypeName(c.objArr(h)) = "JoinBranchStructureSubBranch" Then
'                    Set i = c.objArr(h)
'                    If TypeName(c.objArr(h - 1)) = "JoinBranchStructureSubBranch" Then
'                        Set ii = c.objArr(h - 1)
'
'
'                    ElseIf TypeName(c.objArr(h - 1)) = "SubBranchTables" Then
'                        Set jj = c.objArr(h - 1)
'                        If i.table_count > -1 Then
'                            For x = 0 To i.table_count
'                                If i.joined_Table(x).table_name = jj.table_name Then
'                                    For xx = 0 To i.joined_Table(x).joined_UBound
'                                        output_ = i.joined_Table(x).joined_On_Data(xx).joinet_Table_Column & " " & i.joined_Table(xx).joined_On_Data(0).join_Comparison & " " & i.joined_Table(x).joined_On_Data(xx).joined_Table_Column
'                                        sh.Cells(c_OutputHeaderRow + 4 + xx, 12 + rowCnt).Value = output_
'                                    Next xx
'                                    Exit For
'                                End If
'                            Next x
'                        End If
'                    End If
'
'
'                ElseIf TypeName(c.objArr(h)) = "SubBranchTables" Then
'                    Set j = c.objArr(h)
'                    If TypeName(c.objArr(h - 1)) = "JoinBranchStructureSubBranch" Then
'                        Set ii = c.objArr(h - 1)
'                        If j.joined_UBound > -1 Then
'                            For xx = 0 To j.joined_UBound
'                                output_ = j.joined_On_Data(xx).joinet_Table_Column & " " & j.joined_On_Data(xx).join_Comparison & " " & j.joined_On_Data(xx).joined_Table_Column
'                                sh.Cells(c_OutputHeaderRow + 4 + xx, 12 + rowCnt).Value = output_
'                            Next xx
'                        End If
'                    ElseIf TypeName(c.objArr(h - 1)) = "SubBranchTables" Then
'                        Set jj = c.objArr(h - 1)
'                        If j.joined_UBound > -1 Then
'                            For xx = 0 To j.joined_UBound
'                                output_ = "1"
'                                sh.Cells(c_OutputHeaderRow + 4 + xx, 12 + rowCnt).Value = output_
'                            Next xx
'                        End If
'                    End If
'
'                End If
'            Next h
        Next f
        
        If UBound(e) = 1 Then
            sh.Range(sh.Cells(c_OutputHeaderRow + 2, startCol), sh.Cells(sh.Cells(sh.Cells.Rows.count, startCol).End(xlUp).row, startCol)).Interior.Color = RGB(123, 48, 110)
            sh.Range(sh.Cells(c_OutputHeaderRow + 2, startCol + 1), sh.Cells(sh.Cells(sh.Cells.Rows.count, startCol + 1).End(xlUp).row, startCol + 1)).Interior.Color = RGB(83, 48, 210)
        Else
        
        End If
    Else
        sh.Cells(c_OutputHeaderRow + 3, 12).Value = "No connection found..."
    End If
 
End Function

Public Sub tst()
'    Dim c As DictJoinData
'    Set c = New DictJoinData
'    c.Setup
    
    Dim a As JoinBranchStructure, b As dict_SubBranchResult, c
    
    Dim d As JoinBranchStructure
    Set d = New JoinBranchStructure
    Call d.createBranch("loan_main", "src.uvw_Rolodex_Contacts")
    
    b = d.getBranch()
    
End Sub

'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================



'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================
'========================================================================================================================================================================================================================================================================================================================================

