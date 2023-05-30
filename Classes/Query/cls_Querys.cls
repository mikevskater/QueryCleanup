VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Querys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Query")
'# cls_Querys #

Private p_querys() As cls_Query

Private p_mainQueryBlocks As Integer
Private activeQuery As Integer
Private lastQuery As Integer
Private currentlyInQuery As Boolean

Public Property Let state(v As String)
    Call p_querys(activeQuery).setPState(v)
End Property

Public Property Get state() As String
    state = p_querys(activeQuery).getPState
End Property

Public Property Get QueryCount() As Integer
    QueryCount = UBound(p_querys) + 1
End Property

Public Function query(i) As cls_Query
    Set query = p_querys(i)
End Function

Private Sub Class_Initialize()
    ReDim p_querys(-1 To 0)
    p_mainQueryBlocks = -1
    activeQuery = -1
    lastQuery = -1
End Sub

Public Function createMainQueryBlock()
    p_mainQueryBlocks = p_mainQueryBlocks + 1
End Function

Public Function createNewQuery()
    If LBound(p_querys) = -1 Then
        ReDim p_querys(0 To 0)
        Set p_querys(0) = New cls_Query
        p_querys(0).mainQuery = p_mainQueryBlocks
        p_querys(0).homeQuery = -1
        p_querys(0).ID = 0
        activeQuery = 0
    Else
        Dim ub: ub = UBound(p_querys) + 1
        ReDim Preserve p_querys(0 To ub)
        Set p_querys(ub) = New cls_Query
        p_querys(ub).mainQuery = p_mainQueryBlocks
        p_querys(ub).homeQuery = activeQuery
        p_querys(ub).ID = ub
        
        lastQuery = activeQuery
        activeQuery = ub
    End If
    currentlyInQuery = True
End Function

Public Function closeOffQuery()
    
    activeQuery = lastQuery
    If activeQuery > -1 Then
        lastQuery = p_querys(activeQuery).homeQuery
        currentlyInQuery = True
    Else
        lastQuery = -1
        currentlyInQuery = False
    End If
        
End Function

Public Function closeAllQuerys()
    activeQuery = -1
    currentlyInQuery = False
End Function

Public Function addColumn(columnName, columnAlias, homeTable)
    Call p_querys(activeQuery).addColumn(columnName, columnAlias, homeTable)
End Function

Public Function addTable(tableName, tableAlias, Optional joinType = Empty)
    Call p_querys(activeQuery).addTable(tableName, tableAlias, "table", activeQuery, joinType)
End Function

Public Function addSubQuery(queryAlias)
    Call p_querys(activeQuery).addTable(queryAlias, queryAlias, "subquery", activeQuery, Empty) 'UBound(p_querys) + 1)
End Function

Public Function addJoinOnData(compareType, l_name, l_Alias, compareWord, r_Name, r_Alias)
    Call p_querys(activeQuery).addJoinData(compareType, l_name, l_Alias, compareWord, r_Name, r_Alias)
End Function

Public Function addWhereBetween(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
    Call p_querys(activeQuery).addWhereBetweenData(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
End Function

Public Function addWhere(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
    Call p_querys(activeQuery).addWhereData(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
End Function

Public Function setCurrentQueryTempTable(tempTableName, alias)
    Call p_querys(activeQuery).setTempTable(tempTableName, alias)
End Function

Public Function debugQuery(Optional index_ = Empty)
    Dim i As Integer, j As Integer, x, y, nameMaxWidth, aliasMaxWidth
    Dim curQ As cls_Query
    
    If IsEmpty(index_) = True Then
        Set curQ = p_querys(activeQuery)
    Else
        Set curQ = p_querys(index_)
    End If
    
    If curQ.columns.count > -1 Then
        nameMaxWidth = 0
        aliasMaxWidth = 0
        For i = LBound(curQ.columns.columns) To UBound(curQ.columns.columns)
            If Len(curQ.columns.column(i).column) > nameMaxWidth Then nameMaxWidth = Len(curQ.columns.column(i).column)
            If Len(curQ.columns.column(i).alias) > aliasMaxWidth Then aliasMaxWidth = Len(curQ.columns.column(i).alias)
        Next i
        
        Debug.Print String(nameMaxWidth + aliasMaxWidth + 18, "=")
        Call debugData("Column Table", 13, "Column Name", nameMaxWidth + 1, "Column Alias", aliasMaxWidth + 1)
        Debug.Print String(nameMaxWidth + aliasMaxWidth + 18, "-")
        For i = LBound(curQ.columns.columns) To UBound(curQ.columns.columns)
            Call debugData(curQ.columns.column(i).table, 13, curQ.columns.column(i).column, nameMaxWidth + 1, curQ.columns.column(i).alias, aliasMaxWidth + 1)
        Next i
    End If
    
    'Get tempTable Data
'        If LBound(curQ.columns) > -1 Then
'            Debug.Print "Temp Table Alias", "Temp Table Name"
'            For i = LBound(curQ.columns) To UBound(curQ.columns)
'                Debug.Print curQ.columns.column(i).table, curQ.columns.column(i).column
'            Next i
'        End If
    
    'Get From/Join Data
    If curQ.tables.count > -1 Then
        Debug.Print " "
        maxWidth = 0
        For i = LBound(curQ.tables.tables) To UBound(curQ.tables.tables)
            If Len(curQ.tables.table(i).name) > maxWidth Then maxWidth = Len(curQ.tables.table(i).name)
        Next i
        
        Debug.Print String(maxWidth + 14 + 14, "=")
        Call debugData("Join Type", 12, "Table Alias", 12, "Table Name", maxWidth + 1)
        Debug.Print String(maxWidth + 14 + 14, "-")
        
        For i = LBound(curQ.tables.tables) To UBound(curQ.tables.tables)
            Call debugData(curQ.tables.table(i).joinType, 12, curQ.tables.table(i).alias, 12, curQ.tables.table(i).name, maxWidth + 1)
        Next i
    End If
    
    'Get On Data
    If curQ.joins.count > -1 Then
        Debug.Print " "
        Debug.Print String(10 + 13 + 21 + 8 + 13 + 21, "=")
        Call debugData("Join Type", 9, "Left Table", 12, "Left Paramater", 20, "Compare", 7, "Right Table", 12, "Right Paramater", 20)
        Debug.Print String(10 + 13 + 21 + 8 + 13 + 21, "-")
        'Call debugData("Column Table", 13, "Column Name", nameMaxWidth + 1, "Column Alias", aliasMaxWidth + 1)
        
        For i = LBound(curQ.joins.joins) To UBound(curQ.joins.joins)
            With curQ.joins.join(i)
                If .left_SourceAlias = .left_SourceParamater And .right_SourceAlias = .right_SourceParamater Then
                    Call debugData(.compareType, 9, _
                                 .left_SourceParamater, 33, _
                                 .compareWord, 7, _
                                 .right_SourceParamater, 33)
                ElseIf .left_SourceAlias = .left_SourceParamater Then
                    Call debugData(.compareType, 9, _
                                 .left_SourceParamater, 33, _
                                 .compareWord, 7, _
                                 .right_SourceAlias, 12, _
                                 .right_SourceParamater, 20)
                ElseIf .right_SourceAlias = .right_SourceParamater Then
                    Call debugData(.compareType, 9, _
                                 .left_SourceAlias, 12, _
                                 .left_SourceParamater, 20, _
                                 .compareWord, 7, _
                                 .right_SourceParamater, 33)
                Else
                    Call debugData(.compareType, 9, _
                                 .left_SourceAlias, 12, _
                                 .left_SourceParamater, 20, _
                                 .compareWord, 7, _
                                 .right_SourceAlias, 12, _
                                 .right_SourceParamater, 20)
                End If
            End With
        Next i
    End If
    
    'Get Where Data
    'Get On Data
    If curQ.wheres.count > -1 Then
        Debug.Print " "
        Debug.Print String(10 + 13 + 21 + 8 + 13 + 21, "=")
        Call debugData("TODO FILL IN WHERE INFO", 30)
        Debug.Print String(10 + 13 + 21 + 8 + 13 + 21, "-")
        'Call debugData("Column Table", 13, "Column Name", nameMaxWidth + 1, "Column Alias", aliasMaxWidth + 1)
        
        For i = LBound(curQ.wheres.wheres) To UBound(curQ.wheres.wheres)
            With curQ.wheres.where(i)
                If .compareType = "where" Then
                    If .Where_l_col_Alias = .Where_l_col_Name And .Where_r_col_Alias = .Where_r_col_Name Then
                        Call debugData(.compareType, 10, .Where_l_col_Name, 31, .compareWord, 10, .Where_r_col_Name, 31)
                    ElseIf .Where_l_col_Alias = .Where_l_col_Name Then
                        Call debugData(.compareType, 10, .Where_l_col_Name, 31, .compareWord, 10, .Where_r_col_Alias, 10, .Where_r_col_Name, 20)
                    ElseIf .Where_r_col_Alias = .Where_r_col_Name Then
                        Call debugData(.compareType, 10, .Where_l_col_Alias, 10, .Where_l_col_Name, 20, .compareWord, 10, .Where_r_col_Name, 31)
                    End If
                ElseIf .compareType = "between" Then
                    If .Where_l_col_Alias = .Where_l_col_Name And .Where_r_col_Alias = .Where_r_col_Name Then
                        Call debugData
                    ElseIf .Where_l_col_Alias = .Where_l_col_Name Then
                        Call debugData
                    ElseIf .Where_r_col_Alias = .Where_r_col_Name Then
                        Call debugData
                    End If
                End If
            End With
        Next i
    End If
    
    'Get Order Data
'        If LBound(curQ.columns) > -1 Then
'            For i = LBound(curQ.columns) To UBound(curQ.columns)
'                Debug.Print curQ.columns.column(i).table, curQ.columns.column(i).column
'            Next i
'        End If


End Function

Public Function debugMainQuerys()
    For i = 0 To UBound(p_querys)
        Debug.Print " "
        Debug.Print " "
        Debug.Print "Query " & i
        debugQuery i
    Next i
End Function
