VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JoinBranchStructureSubBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Join")
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

Private p_table_name As String
Private joined_Tables() As SubBranchTables
Private joined_Tables_Count As Long

Property Get table_count()
    table_count = joined_Tables_Count
End Property

Property Get table_name()
    table_name = p_table_name
End Property

Public Function joined_Table(v) As SubBranchTables
    Set joined_Table = joined_Tables(v)
End Function

Property Let table_name(v)
    p_table_name = v
End Property

Private Sub Class_Initialize()
    joined_Tables_Count = -1
    ReDim joined_Tables(-1 To 0)
End Sub

Public Sub createBranch(l_name, dictData As DictJoinData, ByRef treeStruct() As String, currentBranch As Long, branchCount As Long)
    Dim a, b, c, d() As dict_ColumnJoins
    Dim xx As dict_TableJoins, stt
    
    For a = 0 To UBound(dictData.dicts)
        If dictData.dict(a).table_name = l_name Then
            p_table_name = l_name
                For b = 0 To UBound(dictData.dict(a).joined_Tables)
                        If dictData.inBranchAlready(treeStruct, dictData.dict(a).joined_Tables(b).joined_Table_Name) = False Then
                            If joined_Tables_Count = -1 Then
                                joined_Tables_Count = 0
                                ReDim joined_Tables(0 To joined_Tables_Count)
                                Set joined_Tables(joined_Tables_Count) = New SubBranchTables
                            Else
                                joined_Tables_Count = joined_Tables_Count + 1
                                ReDim Preserve joined_Tables(0 To joined_Tables_Count)
                                Set joined_Tables(joined_Tables_Count) = New SubBranchTables
                            End If
                            
                            xx = dictData.dict(a).joined_Tables(b)
                            
                            joined_Tables(joined_Tables_Count).table_name = xx.joined_Table_Name
                            joined_Tables(joined_Tables_Count).joined_On_Datas = xx.joined_On
                            ReDim Preserve treeStruct(0 To UBound(treeStruct) + 1)
                            treeStruct(UBound(treeStruct)) = xx.joined_Table_Name
                            
                            For st = 0 To UBound(dictData.dicts)
                                If dictData.dict(st).table_name = xx.joined_Table_Name Then
                                    For stt = 0 To UBound(dictData.dict(st).joined_Tables)
                                        joined_Tables(joined_Tables_Count).addSubTable dictData.dict(st).joined_Tables(stt).joined_Table_Name
                                    Next stt
                                    Exit For
                                End If
                            Next st
                            
                            'ReDim joined_Tables(joined_Tables_Count).subTables(0 To 10)
                            c = c
                            'Public Type dict_ColumnJoins
                            '    join_Type As String
                            '    joinet_Table_Column As String
                            '    join_Comparison As String
                            '    joined_Table_Column As String
                            '    count As Long
                            'End Type
                        Else
                        
                        End If
                Next b
            Exit For
        End If
    Next a
    
End Sub

Public Sub subBranch(dictData As DictJoinData, ByRef treeStruct() As String, currentBranch As Long, branchCount As Long)
    'Dim a As SubBranchTables, b As JoinBranchStructureSubBranch, c, d
    Dim a, b As JoinBranchStructureSubBranch, bb, c, d
    
    If joined_Tables_Count > -1 Then
        For Each a In joined_Tables
            If a.subTablesUBound > -1 Then
                For bb = 0 To a.subTablesUBound
                    Set b = a.subtable(bb)
                    If currentBranch = branchCount Then
                        Call b.createBranch(b.table_name, dictData, treeStruct, currentBranch, branchCount)
                    Else
                        Call b.subBranch(dictData, treeStruct, currentBranch + 1, branchCount)
                    End If
                Next bb
            End If
        Next a
    End If
End Sub

Public Function getSubBranch(ByRef treeStruct() As String, r_Name, Optional in_Branch = Empty) As dict_SubBranchResult
    Dim a, b As JoinBranchStructureSubBranch, c, d As dict_SubBranchResult, e, f, g
    Dim x, y, z, xx, yy, zz
    
    getSubBranch.objArrCnt = 0
    ReDim getSubBranch.objArr(0 To 0)
    
    If IsEmpty(in_Branch) = True Then
        If inBranch(treeStruct, r_Name) = False Then
            getSubBranch.finalString = ""
            Exit Function
        Else
            in_Branch = True
        End If
    Else
        in_Branch = True
    End If
    
    If p_table_name = r_Name Then
        getSubBranch.finalString = r_Name
        Set getSubBranch.objArr(getSubBranch.objArrCnt) = Me
        Exit Function
    End If
    
    If joined_Tables_Count > -1 Then
        For a = 0 To joined_Tables_Count
            If joined_Tables(a).table_name = r_Name Then
                getSubBranch.finalString = p_table_name & " -> " & r_Name
                getSubBranch.objArrCnt = 1
                ReDim Preserve getSubBranch.objArr(0 To getSubBranch.objArrCnt)
                Set getSubBranch.objArr(getSubBranch.objArrCnt - 1) = joined_Tables(a)
                Set getSubBranch.objArr(getSubBranch.objArrCnt) = Me
                Exit Function
            End If
        Next a
    End If
    
    If joined_Tables_Count > -1 Then
        For Each a In joined_Tables
            If a.subTablesUBound > -1 Then
                For c = 0 To a.subTablesUBound
                    Set b = a.subtable(c)
                    getSubBranch = b.getSubBranch(treeStruct, r_Name, in_Branch)
                    If getSubBranch.finalString <> "" Then
                        If InStr(1, getSubBranch.finalString, a.table_name) = 0 Then
                            getSubBranch.finalString = a.table_name & " -> " & getSubBranch.finalString
                            getSubBranch.objArrCnt = getSubBranch.objArrCnt + 1
                            ReDim Preserve getSubBranch.objArr(0 To getSubBranch.objArrCnt)
                            Set getSubBranch.objArr(getSubBranch.objArrCnt) = a
                        End If
                        If InStr(1, getSubBranch.finalString, p_table_name) = 0 Then
                            getSubBranch.finalString = p_table_name & " -> " & getSubBranch.finalString
                            getSubBranch.objArrCnt = getSubBranch.objArrCnt + 1
                            ReDim Preserve getSubBranch.objArr(0 To getSubBranch.objArrCnt)
                            Set getSubBranch.objArr(getSubBranch.objArrCnt) = Me
                        End If
                        Exit Function
                    End If
                Next c
            End If
        Next a
    End If
    getSubBranch.finalString = ""
End Function

Public Function inBranch(ByRef treeStruct() As String, searchName) As Boolean
    Dim a, b, c
    For a = 0 To UBound(treeStruct)
        If treeStruct(a) = searchName Then
            inBranch = True
            Exit Function
        End If
    Next a
    inBranch = False
End Function
