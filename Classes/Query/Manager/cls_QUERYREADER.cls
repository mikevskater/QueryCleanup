VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_QUERYREADER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Query.Manager")
'# cls_QueryReader #

Private p_sText As String
Private p_lastWord As qr_nextWord
Private p_currentWord As qr_nextWord

Private p_previewLastWord As qr_nextWord

Property Let searchText(v As String)
    p_sText = v
    p_lastWord = emptyNextWord
    p_currentWord.word = ""
    p_currentWord.wordStart = 0
    p_currentWord.wordEnd = 0
    p_currentWord.len = 0
    p_currentWord.EOF = False
End Property

Property Get searchText() As String
    searchText = p_sText
End Property

Property Get lastWord() As qr_nextWord
    lastWord = p_lastWord
End Property

Property Let currentWord(a As qr_nextWord)
    p_currentWord.word = a.word
    p_currentWord.wordStart = a.wordStart
    p_currentWord.wordEnd = a.wordEnd
    p_currentWord.len = a.len
    p_currentWord.EOF = a.EOF
End Property

Property Let currentWord_WordEnd(a As Integer)
    p_currentWord.wordEnd = a
End Property

Property Get currentWord() As qr_nextWord
    currentWord = p_currentWord
End Property

Property Get EOF() As Boolean
    EOF = p_currentWord.EOF
End Property

Public Function clone(c As cls_QUERYREADER)
    p_sText = c.searchText
    
    p_currentWord.word = c.currentWord.word
    p_currentWord.wordStart = c.currentWord.wordStart
    p_currentWord.wordEnd = c.currentWord.wordEnd
    p_currentWord.len = c.currentWord.len
    p_currentWord.EOF = c.currentWord.EOF
    
End Function


Property Get getNextWord(Optional includeBrackets = False) As qr_nextWord

    p_lastWord = p_currentWord
    
    If Len(p_sText) <= p_currentWord.wordEnd + 1 Then
        p_currentWord.len = -1
        p_currentWord.word = ""
        p_currentWord.wordEnd = -1
        p_currentWord.wordStart = -1
        p_currentWord.EOF = True
        Exit Property
    End If
    
    Dim startSpot, wordStart, wordEnd, currLetter, currLetter2, i, j, bracketBranch, squareBracketBranch
    Dim inQuote As Boolean
    startSpot = p_lastWord.wordEnd + 1
    wordStart = Empty
    
    For i = startSpot To Len(p_sText)
        currLetter = Mid(p_sText, i, 1)
        If (currLetter = "(" Or currLetter = ")" Or currLetter = "[" Or currLetter = "]") And includeBrackets = False Then
            p_currentWord.wordStart = i
            p_currentWord.wordEnd = i
            p_currentWord.word = currLetter
            p_currentWord.len = 1
            p_currentWord.EOF = i + 1 >= Len(p_sText)
            getNextWord = p_currentWord
            Exit Property
        End If
        
        If currLetter = "(" Then
            bracketBranch = 0
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "(" Then
                    bracketBranch = bracketBranch + 1
                End If
                
                If currLetter2 = ")" Then
                    If bracketBranch = 0 Then
                        wordStart = i
                        i = j
                        Exit For
'                        p_currentWord.wordStart = i
'                        p_currentWord.wordEnd = j
'                        p_currentWord.word = Mid(p_sText, i, j - i + 1)
'                        p_currentWord.len = j - i + 1
'                        p_currentWord.EOF = j + 1 >= Len(p_sText)
'                        getNextWord = p_currentWord
'                        Exit Property
                    Else
                        bracketBranch = bracketBranch - 1
                    End If
                End If
            Next j
        End If
        
        If currLetter = "[" Then
            squareBracketBranch = 0
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "[" Then
                    squareBracketBranch = squareBracketBranch + 1
                End If
                
                If currLetter2 = "]" Then
                    If squareBracketBranch = 0 Then
                        wordStart = i
                        i = j
                        Exit For
                    Else
                        squareBracketBranch = squareBracketBranch - 1
                    End If
                End If
            Next j
        End If
        
        If currLetter = "'" Then
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "'" Then
                    wordStart = i
                    i = j
                    Exit For
                End If
            Next j
        End If
        
        If IsEmpty(wordStart) = False Then Exit For
        
        If currLetter <> " " And currLetter <> "¦" Then
            wordStart = i
            Exit For
        End If
    Next i
    
    wordEnd = Empty
    bracketBranch = 0
    squareBracketBranch = 0
    inQuote = False
    
    For i = i + 1 To Len(p_sText)
        If i < Len(p_sText) Then
            currLetter = Mid(p_sText, i, 1)
            If currLetter = "'" Then
                inQuote = Not inQuote
            End If
            
            If includeBrackets = True Then
                Select Case currLetter
                    Case "("
                        bracketBranch = bracketBranch + 1
                    Case ")"
                        bracketBranch = bracketBranch - 1
                    Case "["
                        squareBracketBranch = squareBracketBranch + 1
                    Case "]"
                        squareBracketBranch = squareBracketBranch - 1
                    Case " ", "|", "¦"
                        If (bracketBranch = 0 And squareBracketBranch = 0 And inQuote = False) Then
                            wordEnd = i - 1
                            Exit For
                        End If
                End Select
            Else
                If (currLetter = " " Or currLetter = "¦" Or currLetter = "(" Or currLetter = "[") And inQuote = False Then
                    wordEnd = i - 1
                    Exit For
                End If
            End If
        Else
            wordEnd = Len(p_sText)
        End If
    Next i
    
    If IsEmpty(wordEnd) Then wordEnd = Len(p_sText)
    
    p_currentWord.wordStart = wordStart
    p_currentWord.wordEnd = wordEnd
    p_currentWord.word = Mid(p_sText, wordStart, wordEnd - wordStart + 1)
    p_currentWord.len = wordEnd - wordStart + 1
    p_currentWord.EOF = wordEnd + 1 >= Len(p_sText)
    
    getNextWord = p_currentWord
    p_previewLastWord = emptyNextWord
End Property

Public Function previewNextWord(Optional includeBrackets = False) As qr_nextWord

    'p_lastWord = p_currentWord
    
    Dim startSpot, wordStart, wordEnd, currLetter, currLetter2, bracketBranch, squareBracketBranch, i, j
    Dim inQuote As Boolean
    
    If p_previewLastWord.word <> "" Then
        startSpot = p_previewLastWord.wordEnd + 1
    Else
        startSpot = p_currentWord.wordEnd + 1
    End If
    
    If Len(p_sText) < startSpot Then
        previewNextWord = emptyNextWord
        Exit Function
    End If
    wordStart = Empty
    For i = startSpot To Len(p_sText)
        currLetter = Mid(p_sText, i, 1)
        If (currLetter = "(" Or currLetter = ")" Or currLetter = "[" Or currLetter = "]") And includeBrackets = False Then
            previewNextWord.wordStart = i
            previewNextWord.wordEnd = i
            previewNextWord.word = currLetter
            previewNextWord.len = 1
            previewNextWord.EOF = i + 1 >= Len(p_sText)
            p_previewLastWord = previewNextWord
            Exit Function
        End If
        
        If currLetter = "(" Then
            bracketBranch = 0
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "(" Then
                    bracketBranch = bracketBranch + 1
                End If
                
                If currLetter2 = ")" Then
                    If bracketBranch = 0 Then
                        wordStart = i
                        i = j
                        Exit For
                    Else
                        bracketBranch = bracketBranch - 1
                    End If
                End If
            Next j
        End If
        
        If currLetter = "[" Then
            squareBracketBranch = 0
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "[" Then
                    squareBracketBranch = squareBracketBranch + 1
                End If
                
                If currLetter2 = "]" Then
                    If squareBracketBranch = 0 Then
                        wordStart = i
                        i = j
                        Exit For
'                        previewNextWord.wordStart = i
'                        previewNextWord.wordEnd = j
'                        previewNextWord.word = Mid(p_sText, i, j - i + 1)
'                        previewNextWord.len = j - i + 1
'                        previewNextWord.EOF = j + 1 >= Len(p_sText)
'                        p_previewLastWord = previewNextWord
'                        Exit Function
                    Else
                        squareBracketBranch = squareBracketBranch - 1
                    End If
                End If
            Next j
        End If
        
        If currLetter = "'" Then
            For j = i + 1 To Len(p_sText)
                currLetter2 = Mid(p_sText, j, 1)
                If currLetter2 = "'" Then
                    wordStart = i
                    i = j
                    Exit For
                End If
            Next j
        End If
        
        If IsEmpty(wordStart) = False Then Exit For
        
        If currLetter <> " " And currLetter <> "¦" Then
            wordStart = i
            Exit For
        End If
    Next i
    
    If IsEmpty(wordStart) = True Then
        previewNextWord = emptyNextWord
        Exit Function
    End If
    
    wordEnd = Empty
    bracketBranch = 0
    squareBracketBranch = 0
    For i = i + 1 To Len(p_sText)
        If i < Len(p_sText) Then
            currLetter = Mid(p_sText, i, 1)
            If currLetter = "'" Then
                inQuote = Not inQuote
            End If
            
            If includeBrackets = True Then
                If currLetter = "(" Then
                    bracketBranch = bracketBranch + 1
                End If
                
                If currLetter = ")" Then
                    bracketBranch = bracketBranch - 1
                End If
                
                If currLetter = "[" Then
                    squareBracketBranch = squareBracketBranch + 1
                End If
                
                If currLetter = "]" Then
                    squareBracketBranch = squareBracketBranch - 1
                End If
                
                If (currLetter = " " Or currLetter = "¦") And (bracketBranch = 0 And squareBracketBranch = 0 And inQuote = False) Then
                    wordEnd = i - 1
                    Exit For
                End If
            Else
                If (currLetter = " " Or currLetter = "¦" Or currLetter = "(" Or currLetter = "[") And inQuote = False Then
                    wordEnd = i - 1
                    Exit For
                End If
            End If
        Else
            wordEnd = Len(p_sText)
        End If
    Next i
    
    If IsEmpty(wordEnd) Then wordEnd = Len(p_sText)
    
    previewNextWord.wordStart = wordStart
    previewNextWord.wordEnd = wordEnd
    previewNextWord.word = Mid(p_sText, wordStart, wordEnd - wordStart + 1)
    previewNextWord.len = wordEnd - wordStart + 1
    previewNextWord.EOF = wordEnd + 1 >= Len(p_sText)
    p_previewLastWord = previewNextWord
End Function

Public Function resetPreviewPlace()
    p_previewLastWord = emptyNextWord
End Function

Public Function scanForCurrentBranchNextKeyWord(ParamArray v() As Variant) As qr_nextWord
    Dim tempLastWord As qr_nextWord, tempCurWord As qr_nextWord, tempNextWord As qr_nextWord, tempStartWord As qr_nextWord
    Dim currentBranch, subBranchCount
    Dim tempScanWordBool As Boolean
    Dim bracketCount As Integer
    Dim i, ii, j, x, y
    
    Dim loop_ As Boolean
    loop_ = True
    
    If IsArray(v(0)) = True Then
        Dim tempV
        tempV = v(0)
        v = tempV
    End If
    
    Dim tQR As cls_QUERYREADER 'tempQryRdr
    Set tQR = New cls_QUERYREADER 'tempQryRdr
    Call tQR.clone(qryReader)
          
    bracketCount = 0
    currentBranch = 0
    tempStartWord = tQR.currentWord
             
    Dim searchTerms
    searchTerms = getStartStatements
    
    Do
        'Scan next word
        tempLastWord = tQR.currentWord
        tempCurWord = tQR.getNextWord(True)
        tempNextWord = tQR.previewNextWord(True)
        
        tempLastWord.word = Trim(Replace(tempLastWord.word, "¦", " "))
        tempCurWord.word = Trim(Replace(tempCurWord.word, "¦", " "))
        tempNextWord.word = Trim(Replace(tempNextWord.word, "¦", " "))
                        
        For ii = 0 To UBound(searchTerms)
            If tempCurWord.word = searchTerms(ii) Then
                tempScanWordBool = False
                For i = LBound(v) To UBound(v)
                    If v(i) = searchTerms(ii) Then
                        tempScanWordBool = True
                        Exit For
                    End If
                Next i
                If tempScanWordBool = True Then
                    If currentBranch = 0 Then
                        scanForCurrentBranchNextKeyWord = tempCurWord
                        Exit Function
                    Else
                        currentBranch = currentBranch + 1
                    End If
                Else
                    currentBranch = currentBranch + 1
                End If
                GoTo nextBranchWordScan
            End If
        Next ii
        
        If InStr(1, tempCurWord.word, "(") > 0 Then
            tempScanWordBool = False
            For i = LBound(v) To UBound(v)
                If v(i) = "(" Then
                    tempScanWordBool = True
                    Exit For
                End If
            Next i
            If tempScanWordBool = True Then
                If bracketCount = 0 Then
                    scanForCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                Else
                    bracketCount = bracketCount + 1
                End If
            Else
                bracketCount = bracketCount + 1
            End If
            GoTo nextBranchWordScan
        End If
        
        If tempCurWord.word = "from" Then
            tempScanWordBool = False
            For i = LBound(v) To UBound(v)
                If v(i) = "from" Then
                    tempScanWordBool = True
                    Exit For
                End If
            Next i
            If tempScanWordBool = True Then
                If currentBranch = 0 Then
                    scanForCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                Else
                    currentBranch = currentBranch - 1
                End If
            Else
                If currentBranch = 0 Then
                    scanForCurrentBranchNextKeyWord = emptyNextWord
                    Exit Function
                Else
                    currentBranch = currentBranch - 1
                End If
            End If
            GoTo nextBranchWordScan
        End If
        
        If InStr(1, tempCurWord.word, ")") > 0 Then
            tempScanWordBool = False
            For i = LBound(v) To UBound(v)
                If v(i) = ")" Then
                    tempScanWordBool = True
                    Exit For
                End If
            Next i
            If tempScanWordBool = True Then
                If bracketCount = 0 Then
                    scanForCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                Else
                    bracketCount = bracketCount - 1
                End If
            Else
                If bracketCount = 0 Then
                    scanForCurrentBranchNextKeyWord = emptyNextWord
                    Exit Function
                Else
                    bracketCount = bracketCount - 1
                End If
            End If
            GoTo nextBranchWordScan
        End If
        
        For i = LBound(v) To UBound(v)
            If v(i) = tempCurWord.word Then
                If currentBranch = 0 And bracketCount = 0 Then
                    scanForCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                End If
            End If
        Next i
        
nextBranchWordScan:
        If tQR.previewNextWord.EOF = True Then loop_ = False
    Loop Until (loop_ = False)
                    
    scanForCurrentBranchNextKeyWord = emptyNextWord
    
End Function

Public Function scanInCurrentBranchNextKeyWord(ParamArray v() As Variant) As qr_nextWord
    Dim tempLastWord As qr_nextWord, tempCurWord As qr_nextWord, tempNextWord As qr_nextWord, tempStartWord As qr_nextWord
    Dim currentBranch, subBranchCount
    Dim tempScanWordBool As Boolean
    Dim i, ii, j, x, y
    
    Dim loop_ As Boolean
    loop_ = True
    
    If IsArray(v(0)) = True Then
        Dim tempV
        tempV = v(0)
        v = tempV
    End If
    
    Dim tQR As cls_QUERYREADER 'tempQryRdr
    Set tQR = New cls_QUERYREADER 'tempQryRdr
    Call tQR.clone(qryReader)
          
    currentBranch = 0
    tempStartWord = tQR.currentWord
             
    Dim searchTerms
    searchTerms = getStartStatements
    Do
        'Scan next word
        tempLastWord = tQR.currentWord
        tempCurWord = tQR.getNextWord(True)
        tempNextWord = tQR.previewNextWord(True)
                
        tempLastWord.word = Replace(tempLastWord.word, "¦", " ")
        tempCurWord.word = Replace(tempCurWord.word, "¦", " ")
        tempNextWord.word = Replace(tempNextWord.word, "¦", " ")
        
        For ii = 0 To UBound(searchTerms)
            If tempCurWord.word = searchTerms(ii) Then
                tempScanWordBool = False
                For i = LBound(v) To UBound(v)
                    If v(i) = searchTerms(ii) Then
                        tempScanWordBool = True
                        Exit For
                    End If
                Next i
                If tempScanWordBool = True Then
                    If currentBranch = 0 Then
                        scanInCurrentBranchNextKeyWord = tempCurWord
                        Exit Function
                    Else
                        currentBranch = currentBranch + 1
                    End If
                Else
                    currentBranch = currentBranch + 1
                End If
                GoTo nextBranchWordScan
            End If
        Next ii
        
        'If tempCurWord.word = "from" Then
        '    Dim tempNextScanWord, tempNextEndWord As qr_nextWord
        '    tempNextEndWord = tQR.scanForCurrentBranchNextKeyWord("where")
        '    tempNextScanWord = scanForNextMajorQueryKeyWord(tQR)
        'End If
        
        If tempCurWord.word = "from" Then
            tempScanWordBool = False
            For i = LBound(v) To UBound(v)
                If v(i) = "from" Then
                    tempScanWordBool = True
                    Exit For
                End If
            Next i
            If tempScanWordBool = True Then
                If currentBranch = 0 Then
                    scanInCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                Else
                    currentBranch = currentBranch - 1
                End If
            Else
                If currentBranch = 0 Then
                    scanInCurrentBranchNextKeyWord = emptyNextWord
                    Exit Function
                Else
                    currentBranch = currentBranch - 1
                End If
            End If
            GoTo nextBranchWordScan
        End If
        
        
        For i = LBound(v) To UBound(v)
            If InStr(1, tempCurWord.word, v(i)) > 0 Then
                If currentBranch = 0 Then
                    scanInCurrentBranchNextKeyWord = tempCurWord
                    Exit Function
                End If
            End If
        Next i
        
nextBranchWordScan:
        If tempNextWord.EOF = True Then loop_ = False
    Loop Until (loop_ = False)
                    
    scanInCurrentBranchNextKeyWord = emptyNextWord
    
End Function

Private Function scanForNextMajorQueryKeyWord()

End Function

Public Function getTextBetweenWords(sWord As qr_nextWord, eWord As qr_nextWord)
    getTextBetweenWords = Mid(p_sText, sWord.wordStart, eWord.wordEnd - sWord.wordStart + 1)
End Function

Private Sub Class_Initialize()
    p_previewLastWord = emptyNextWord
    pub_EmptyWord = emptyNextWord
End Sub
