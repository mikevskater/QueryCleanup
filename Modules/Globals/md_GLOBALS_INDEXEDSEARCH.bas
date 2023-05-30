Attribute VB_Name = "md_GLOBALS_INDEXEDSEARCH"
'@Folder("Mods.Globals")
Private letteredIndex(1 To 36) As Variant
Private letteredIndexCounts(1 To 36) As Variant
Private init_ As Boolean

Public Function isTableName(t) As Boolean
    isTableName = indexedSearchMatch(t)
End Function

Public Function setup_IndexSearch()
    Dim curWord, curWords As Variant, arr(), a, b
    ReDim arr(0 To 0)
    For a = 1 To 36
        letteredIndex(a) = arr
        letteredIndexCounts(a) = -1
    Next a
    
    curWords = Sheets("Table Names").Range(Sheets("Table Names").Cells(1, 1), Sheets("Table Names").Cells(Sheets("Table Names").Cells(Sheets("Table Names").Rows.count, 1).End(xlUp).row, 1)).Value
    
    For a = 1 To UBound(curWords, 1)
        curWord = curWords(a, 1)
        b = -1
        Select Case LCase$(Left$(curWord, 1))
            Case "a"
                b = 1
            Case "b"
                b = 2
            Case "c"
                b = 3
            Case "d"
                b = 4
            Case "e"
                b = 5
            Case "f"
                b = 6
            Case "g"
                b = 7
            Case "h"
                b = 8
            Case "i"
                b = 9
            Case "j"
                b = 10
            Case "k"
                b = 11
            Case "l"
                b = 12
            Case "m"
                b = 13
            Case "n"
                b = 14
            Case "o"
                b = 15
            Case "p"
                b = 16
            Case "q"
                b = 17
            Case "r"
                b = 18
            Case "s"
                b = 19
            Case "t"
                b = 20
            Case "u"
                b = 21
            Case "v"
                b = 22
            Case "w"
                b = 23
            Case "x"
                b = 24
            Case "y"
                b = 25
            Case "z"
                b = 26
            Case "0"
                b = 27
            Case "1"
                b = 28
            Case "2"
                b = 29
            Case "3"
                b = 30
            Case "4"
                b = 31
            Case "5"
                b = 32
            Case "6"
                b = 33
            Case "7"
                b = 34
            Case "8"
                b = 35
            Case "9"
                b = 36
        End Select
        If b > -1 Then
            letteredIndexCounts(b) = letteredIndexCounts(b) + 1
            arr = letteredIndex(b)
            ReDim Preserve arr(0 To letteredIndexCounts(b))
            arr(letteredIndexCounts(b)) = LCase(curWord)
            letteredIndex(b) = arr
        End If
    Next a
   
    init_ = True
End Function

Public Function indexedSearchMatch(searchWord) As Boolean
    If init_ = False Then setup_IndexSearch
    searchWord = LCase(searchWord)
    Select Case True
        Case searchWord Like "a*"
            indexedSearchMatch = letterSearch(1, searchWord)
        Case searchWord Like "b*"
            indexedSearchMatch = letterSearch(2, searchWord)
        Case searchWord Like "c*"
            indexedSearchMatch = letterSearch(3, searchWord)
        Case searchWord Like "d*"
            indexedSearchMatch = letterSearch(4, searchWord)
        Case searchWord Like "e*"
            indexedSearchMatch = letterSearch(5, searchWord)
        Case searchWord Like "f*"
            indexedSearchMatch = letterSearch(6, searchWord)
        Case searchWord Like "g*"
            indexedSearchMatch = letterSearch(7, searchWord)
        Case searchWord Like "h*"
            indexedSearchMatch = letterSearch(8, searchWord)
        Case searchWord Like "i*"
            indexedSearchMatch = letterSearch(9, searchWord)
        Case searchWord Like "j*"
            indexedSearchMatch = letterSearch(10, searchWord)
        Case searchWord Like "k*"
            indexedSearchMatch = letterSearch(11, searchWord)
        Case searchWord Like "l*"
            indexedSearchMatch = letterSearch(12, searchWord)
        Case searchWord Like "m*"
            indexedSearchMatch = letterSearch(13, searchWord)
        Case searchWord Like "n*"
            indexedSearchMatch = letterSearch(14, searchWord)
        Case searchWord Like "o*"
            indexedSearchMatch = letterSearch(15, searchWord)
        Case searchWord Like "p*"
            indexedSearchMatch = letterSearch(16, searchWord)
        Case searchWord Like "q*"
            indexedSearchMatch = letterSearch(17, searchWord)
        Case searchWord Like "r*"
            indexedSearchMatch = letterSearch(18, searchWord)
        Case searchWord Like "s*"
            indexedSearchMatch = letterSearch(19, searchWord)
        Case searchWord Like "t*"
            indexedSearchMatch = letterSearch(20, searchWord)
        Case searchWord Like "u*"
            indexedSearchMatch = letterSearch(21, searchWord)
        Case searchWord Like "v*"
            indexedSearchMatch = letterSearch(22, searchWord)
        Case searchWord Like "w*"
            indexedSearchMatch = letterSearch(23, searchWord)
        Case searchWord Like "x*"
            indexedSearchMatch = letterSearch(24, searchWord)
        Case searchWord Like "y*"
            indexedSearchMatch = letterSearch(25, searchWord)
        Case searchWord Like "z*"
            indexedSearchMatch = letterSearch(26, searchWord)
        Case searchWord Like "0*"
            indexedSearchMatch = letterSearch(27, searchWord)
        Case searchWord Like "1*"
            indexedSearchMatch = letterSearch(28, searchWord)
        Case searchWord Like "2*"
            indexedSearchMatch = letterSearch(29, searchWord)
        Case searchWord Like "3*"
            indexedSearchMatch = letterSearch(30, searchWord)
        Case searchWord Like "4*"
            indexedSearchMatch = letterSearch(31, searchWord)
        Case searchWord Like "5*"
            indexedSearchMatch = letterSearch(32, searchWord)
        Case searchWord Like "6*"
            indexedSearchMatch = letterSearch(33, searchWord)
        Case searchWord Like "7*"
            indexedSearchMatch = letterSearch(34, searchWord)
        Case searchWord Like "8*"
            indexedSearchMatch = letterSearch(35, searchWord)
        Case searchWord Like "9*"
            indexedSearchMatch = letterSearch(36, searchWord)
        Case Else
            indexedSearchMatch = False
    End Select
    
End Function

Private Function letterSearch(letterIndex, searchWord) As Boolean
    Dim a
    For Each a In letteredIndex(letterIndex)
        If a = searchWord Then
            letterSearch = True
            Exit Function
        End If
    Next a
    letterSearch = False
End Function

Public Function inRev(wrd, wrdMtch) As Long
    inRev = InStrRev(wrd, wrdMtch)
End Function

Public Function basicSearch(wrd)
    Dim curWord, curWords As Variant, arr(), a, b
    wrd = LCase(wrd)
    curWords = Sheets("Table Names").Range(Sheets("Table Names").Cells(1, 1), Sheets("Table Names").Cells(Sheets("Table Names").Cells(Sheets("Table Names").Rows.count, 1).End(xlUp).row, 1)).Value
    For a = 1 To UBound(curWords, 1)
        If curWords(a, 1) = wrd Then
            basicSearch = True
            Exit Function
        End If
    Next a
    basicSearch = False
End Function

Sub speedTest()
    Dim a, b, c, d
    a = Timer
    indexedSearchMatch ("fds")
    Debug.Print "b " & Round(Timer - a, 6)
    a = Timer
    basicSearch ("fds")
    Debug.Print "a " & Round(Timer - a, 6)
End Sub
