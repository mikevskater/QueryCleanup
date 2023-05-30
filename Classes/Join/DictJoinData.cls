VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictJoinData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Join")
Option Explicit

'==============================================================================================================================
' Join Dictionary
'==============================================================================================================================
'Public Type dict_ColumnJoins
'    join_Type As String
'    joinet_Table_Column As String
'    join_Comparison As String
'    joined_Table_Column As String
'    count As Long
'End Type
'
'Public Type dict_JoinData
'    joined_Table_Name As String
'    joined_On() As dict_ColumnJoins
'End Type
'
'Public Type dict_TableJoinsSubBranch
'    table_Name As String
'    joined_Tables() As dict_JoinData
'End Type
'
'Public Type dict_TableJoins
'    joined_Table_Name As String
'    joined_On() As dict_ColumnJoins
'    SubBranch() As dict_TableJoinsSubBranch
'End Type
'
'Public Type dict_JoinConnections
'    table_Name As String
'    joined_Tables() As dict_TableJoins
'End Type
'
'Public Type dict_JoinBranch
'    start_Table As dict_JoinConnections
'End Type

Private joinDict() As dict_JoinConnections

Private dictCount As Long
Private init_ As Boolean

Property Get dicts() As dict_JoinConnections()
    dicts = joinDict
End Property

Public Function dict(i) As dict_JoinConnections
    dict = joinDict(i)
End Function


'==============================================================================================================================
' Join Dictionary
'==============================================================================================================================

Private Sub Class_Initialize()
    init_ = False
    ReDim joinDict(-1 To 0)
    dictCount = -1
End Sub

Public Sub Setup()
    'Parse Json
    If IsEmpty(p_JsonData) = True Then p_JsonData = parseAllJson
    dictionaryJsonJoinsIntoData
    sortDictionary
    'createBranchArray "loan_main", "custom_loanmain"
    
    init_ = True
End Sub

Private Function dict_tableExists(tableName)
    Dim a, b
    If dictCount < 0 Then
        dict_tableExists = -1
        Exit Function
    End If
    
    For a = 0 To dictCount
        If joinDict(a).table_name = tableName Then
            dict_tableExists = a
            Exit Function
        End If
    Next a
    dict_tableExists = -1
End Function

Private Function dictionaryJsonJoinsIntoData()
    Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, q, r, ii
    Dim p
    Dim p2, p1
    Dim a1, a2
    Dim d1a, d2a, d1b, d2b
    
    p = p_JsonData
    
    For c = 0 To UBound(p)
        p2 = p(c)(2)
        p1 = p(c)(1)
        If IsArray(p2) = True And IsArray(p1) = True Then
            For a = 0 To UBound(p2)
                a1 = False
                a2 = False
                For b = 0 To UBound(p1)
                    If p2(a)(2) = p1(b)(1) Then
                        p(c)(2)(a)(2) = p1(b)(0)
                        a1 = True
                    End If
                    
                    If p2(a)(5) = p1(b)(1) Then
                        p(c)(2)(a)(5) = p1(b)(0)
                        a2 = True
                    End If
                    If a1 = True And a2 = True Then Exit For
                Next b
            Next a
        End If
    Next c
    
    For c = 0 To UBound(p)
        If IsArray(p(c)(2)) = True Then
            For a = 0 To UBound(p(c)(2))
                Dim rr
                rr = Replace(p(c)(2)(a)(2), "[", "")
                rr = Replace(rr, "]", "")
                rr = Replace(rr, "rpt.", "")
                rr = Replace(rr, "dbo.", "")
                rr = Replace(rr, "integra.", "")
                p(c)(2)(a)(2) = rr
                
                rr = Replace(p(c)(2)(a)(5), "[", "")
                rr = Replace(rr, "]", "")
                rr = Replace(rr, "rpt.", "")
                rr = Replace(rr, "dbo.", "")
                rr = Replace(rr, "integra.", "")
                p(c)(2)(a)(5) = rr
            Next a
        End If
    Next c
    
    For c = 0 To UBound(p)
        d = p(c)(2)
        If IsArray(d) = True Then
            For e = 0 To UBound(d)
                'add each join to the dictionary
                If indexedSearchMatch(d(e)(2)) = True And indexedSearchMatch(d(e)(5)) = True Then
                    For ii = 1 To 2
                        If ii = 1 Then
                            d1a = d(e)(2)
                            d1b = d(e)(1)
                            d2a = d(e)(5)
                            d2b = d(e)(4)
                        Else
                            d1a = d(e)(5)
                            d1b = d(e)(4)
                            d2a = d(e)(2)
                            d2b = d(e)(1)
                        End If
                    
                        f = dict_tableExists(d1a)
                        If f = -1 Then
                            If dictCount = -1 Then
                                ReDim joinDict(0 To 0)
                            Else
                                ReDim Preserve joinDict(0 To dictCount + 1)
                            End If
                            dictCount = dictCount + 1
                            joinDict(dictCount).table_name = d1a
                            joinDict(dictCount).count = 1
                            
                            ReDim joinDict(dictCount).joined_Tables(0 To 0)
                            joinDict(dictCount).joined_Tables(0).joined_Table_Name = d2a
                            joinDict(dictCount).joined_Tables(0).count = 1
                            
                            ReDim joinDict(dictCount).joined_Tables(0).joined_On(0 To 0)
                            With joinDict(dictCount).joined_Tables(0).joined_On(0)
                                .count = 1
                                .join_Comparison = d(e)(3)
                                .joinet_Table_Column = d1b
                                .joined_Table_Column = d2b
                                .join_Type = d(e)(0)
                            End With
                        Else
                            g = False
                            For h = 0 To UBound(joinDict(f).joined_Tables)
                                If joinDict(f).joined_Tables(h).joined_Table_Name = d2a Then
                                    g = True
                                    i = h
                                    Exit For
                                End If
                            Next h
                            
                            If g = True Then
                                joinDict(f).count = joinDict(f).count + 1
                                j = False
                                joinDict(f).joined_Tables(i).count = joinDict(f).joined_Tables(i).count + 1
                                With joinDict(f).joined_Tables(i)
                                    For h = 0 To UBound(.joined_On)
                                        If .joined_On(h).joinet_Table_Column = d1b And .joined_On(h).joined_Table_Column = d2b Then
                                            j = True
                                            .joined_On(h).count = .joined_On(h).count + 1
                                            Exit For
                                        End If
                                    Next h
                                    If j = False Then
                                        ReDim Preserve .joined_On(0 To UBound(.joined_On) + 1)
                                        .joined_On(UBound(.joined_On)).count = 1
                                        .joined_On(UBound(.joined_On)).join_Comparison = d(e)(3)
                                        .joined_On(UBound(.joined_On)).joinet_Table_Column = d1b
                                        .joined_On(UBound(.joined_On)).joined_Table_Column = d2b
                                        .joined_On(UBound(.joined_On)).join_Type = d(e)(0)
                                    End If
                                End With
                            Else
                                joinDict(f).count = joinDict(f).count + 1
                                ReDim Preserve joinDict(f).joined_Tables(0 To UBound(joinDict(f).joined_Tables) + 1)
                                joinDict(f).joined_Tables(UBound(joinDict(f).joined_Tables)).count = 1
                                With joinDict(f).joined_Tables(UBound(joinDict(f).joined_Tables))
                                    .joined_Table_Name = d2a
                                    ReDim .joined_On(0 To 0)
                                    .joined_On(0).count = 1
                                    .joined_On(0).join_Comparison = d(e)(3)
                                    .joined_On(0).joinet_Table_Column = d1b
                                    .joined_On(0).joined_Table_Column = d2b
                                    .joined_On(0).join_Type = d(e)(0)
                                End With
                            End If
                            'a(b).table_Name =
                        End If
                    Next ii
                End If
            Next e
        End If
    Next c
    
    c = c
    
    
End Function

'Public Function createBranchArray(l_tbl, r_tbl)
'    Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z
'
'    For a = 0 To dictCount
'        If joinDict(a).table_Name = l_tbl Then
'            oldList(0) = l_tbl
'            joinBranch.start_Table = joinDict(a)
'            Exit For
'        End If
'    Next a
'
'    For a = 0 To UBound(joinBranch.start_Table.joined_Tables)
'        generateSubBranch joinBranch.start_Table.joined_Tables(a), r_tbl, 1
'
''        If inBranchAlready(joinBranch.start_Table.joined_Tables(a).joined_Table_Name) = False Then
''            For b = 0 To dictCount
''                If joinBranch.start_Table.joined_Tables(a).joined_Table_Name = joinDict(b).table_Name Then
''
''                    For e = 0 To UBound(joinDict(b).joined_Tables)
''
''                        If inBranchAlready(joinDict(b).joined_Tables(e).joined_Table_Name) = False Then
''                            If IsBranchEmpty(joinBranch.start_Table.joined_Tables(a).SubBranch) = True Then
''                                ReDim joinBranch.start_Table.joined_Tables(a).SubBranch(0 To 0)
''                            Else
''                                ReDim Preserve joinBranch.start_Table.joined_Tables(a).SubBranch(0 To UBound(joinBranch.start_Table.joined_Tables(a).SubBranch) + 1)
''                            End If
''                            For f = 0 To dictCount
''                                If joinDict(f).table_Name = joinDict(b).joined_Tables(e).joined_Table_Name Then
''                                    joinBranch.start_Table.joined_Tables(a).SubBranch(UBound(joinBranch.start_Table.joined_Tables(a).SubBranch)).table_Name = joinDict(f).table_Name
''                                    joinBranch.start_Table.joined_Tables(a).SubBranch(UBound(joinBranch.start_Table.joined_Tables(a).SubBranch)).joined_Tables = joinDict(f).joined_Tables
''                                    Exit For
''                                End If
''                            Next f
''                        End If
''
''                    Next e
''
''                    treeCount = treeCount + 1
''                    oldList(treeCount) = joinDict(b).table_Name
''                    Exit For
''
''                End If
''            Next b
''        End If
'    Next a
'
'    a = a
'
'End Function

'Private Function generateSubBranch(ByRef Branch As dict_TableJoins, r_tbl, branchCount)
'    Dim a, b, c, d, e, f
'    If branchCount > 10 Then Exit Function
'
'    For b = 0 To dictCount
'        If Branch.joined_Table_Name = joinDict(b).table_Name Then
'            For e = 0 To UBound(joinDict(b).joined_Tables)
'                If inBranchAlready(joinDict(b).joined_Tables(e).joined_Table_Name) = False Then
'                    If IsBranchEmpty(Branch.SubBranch) = True Then
'                        ReDim Branch.SubBranch(0 To 0)
'                    Else
'                        ReDim Preserve Branch.SubBranch(0 To UBound(joinBranch.start_Table.joined_Tables(a).SubBranch) + 1)
'                    End If
'                    For f = 0 To dictCount
'                        If joinDict(f).table_Name = joinDict(b).joined_Tables(e).joined_Table_Name Then
'                            Branch.SubBranch(UBound(Branch.SubBranch)).table_Name = joinDict(f).table_Name
'                            Branch.SubBranch(UBound(Branch.SubBranch)).joined_Tables = joinDict(f).joined_Tables
'                            Exit For
'                        End If
'                    Next f
'                End If
'            Next e
'
'            treeCount = treeCount + 1
'            oldList(treeCount) = joinDict(b).table_Name
'            Exit For
'
'        End If
'    Next b
'
'    For b = 0 To UBound(Branch.SubBranch)
'        'generateSubBranch Branch.SubBranch(b), r_tbl, branchCount + 1
'    Next b
'
'End Function


Public Function inBranchAlready(ByRef tblList() As String, tableName) As Boolean
    Dim e
    If IsArray(tblList) = False Then
        inBranchAlready = False
        Exit Function
    End If
    
    For e = 0 To UBound(tblList)
        If tblList(e) = tableName Then
            inBranchAlready = True
            Exit Function
        End If
    Next e
    inBranchAlready = False
End Function
'
'
'Public Function IsBranchEmpty(arr() As dict_TableJoinsSubBranch) As Boolean
'    On Error GoTo arremptyerrorExit
'    If LBound(arr) > -1 Then
'        IsBranchEmpty = False
'        On Error GoTo 0
'        Exit Function
'    End If
'arremptyerrorExit:
'    On Error GoTo 0
'    IsBranchEmpty = True
'End Function
'
'

Private Sub sortDictionary()
    Dim a, b, c, d, e
    Dim tmpDic As dict_JoinConnections
    Dim tmpDic_ As dict_TableJoins
    Dim cng_ As Boolean
    Dim loopCnt_ As Long
    loopCnt_ = 0
    Do
        loopCnt_ = loopCnt_ + 1
        cng_ = False
        For a = UBound(joinDict) To 1 Step -1
            If joinDict(a).count > joinDict(a - 1).count Then
                tmpDic = joinDict(a)
                joinDict(a) = joinDict(a - 1)
                joinDict(a - 1) = tmpDic
                cng_ = True
            ElseIf (joinDict(a).count = joinDict(a - 1).count) And UBound(joinDict(a).joined_Tables) > UBound(joinDict(a - 1).joined_Tables) Then
                tmpDic = joinDict(a)
                joinDict(a) = joinDict(a - 1)
                joinDict(a - 1) = tmpDic
                cng_ = True
            End If
        Next a
        If loopCnt_ = 100 Then
            loopCnt_ = 0
        End If
    Loop Until (cng_ = False)
    
    loopCnt_ = 0
    Do
        loopCnt_ = loopCnt_ + 1
        cng_ = False
        For a = 0 To UBound(joinDict)
            For b = UBound(joinDict(a).joined_Tables) To 1 Step -1
                If joinDict(a).joined_Tables(b).count > joinDict(a).joined_Tables(b - 1).count Then
                    tmpDic_ = joinDict(a).joined_Tables(b)
                    joinDict(a).joined_Tables(b) = joinDict(a).joined_Tables(b - 1)
                    joinDict(a).joined_Tables(b - 1) = tmpDic_
                    cng_ = True
                End If
            Next b
        Next a
        If loopCnt_ = 100 Then
            loopCnt_ = 0
        End If
    Loop Until (cng_ = False)
    
    b = b
    
    c = c
End Sub









