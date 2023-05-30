VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Query")
'# cls_Query #

Private p_Columns As cls_Columns
Private p_Tables As cls_Tables
Private p_Joins As cls_Joins
Private p_Wheres As cls_Wheres

Private p_state As String
Private p_mainQuery As Integer
Private p_homeQuery As Integer
Private p_subQuerys() As Integer
Private p_ID As Integer

Private p_tempTableName As String
Private p_tempTableAlias As String


'=============================================================================================================================================================================================

'=============================================================================================================================================================================================

Public Property Get state() As String
    state = p_state
End Property

Public Property Let state(v As String)
    p_state = v
End Property

Public Property Get mainQuery() As Integer
    mainQuery = p_mainQuery
End Property

Public Property Let mainQuery(v As Integer)
    p_mainQuery = v
End Property

Public Property Get homeQuery() As Integer
    homeQuery = p_homeQuery
End Property

Public Property Let homeQuery(v As Integer)
    p_homeQuery = v
End Property

Public Property Get ID() As Integer
    ID_ = p_ID
End Property

Public Property Let ID(i As Integer)
    p_ID = i
End Property

Public Property Get columns() As cls_Columns
    Set columns = p_Columns
End Property

Public Property Get tables() As cls_Tables
    Set tables = p_Tables
End Property

Public Property Get joins() As cls_Joins
    Set joins = p_Joins
End Property

Public Property Get wheres() As cls_Wheres
    Set wheres = p_Wheres
End Property

Private Sub Class_Initialize()
    Set p_Columns = New cls_Columns
    Set p_Tables = New cls_Tables
    Set p_Joins = New cls_Joins
    Set p_Wheres = New cls_Wheres
    p_state = "created"
End Sub

Public Sub setPState(state)
    Select Case state
        Case "created"
            p_state = "setupquery"
        Case "addingcolumns"
            p_state = "addingcolumns"
        Case "addingtables"
            p_state = "addingtables"
        Case "wherestart"
            p_state = "addingwhere"
        Case Else
            p_state = "na"
    End Select
End Sub

Public Function getPState() As String
    getPState = p_state
End Function

Public Function getTableByAlias(tableAlias) As cls_TABLE
    Dim i As Integer
    
    If tableAlias = "FORMULA" Then
        Set getTableByAlias = Nothing
        Exit Function
    End If
    
    If tableAlias = "HOMETABLENOALIAS" Then
        Set getTableByAlias = Nothing
        Exit Function
    End If
    
    If LBound(p_Tables.tables) > -1 Then
        For i = 0 To UBound(p_Tables.tables)
            If p_Tables.table(i).alias = tableAlias Then
                Set getTableByAlias = p_Tables.table(i)
                Exit Function
            End If
        Next i
    End If
    Set getTableByAlias = Nothing
End Function

Public Function addSubQuery(v As Integer)
    If LBound(p_subQuerys) = -1 Then
        ReDim p_subQuerys(0 To 0)
        p_subQuerys(0) = v
    Else
        ReDim Preserve p_subQuerys(0 To UBound(p_subQuerys) + 1)
        p_subQuerys(UBound(p_subQuerys)) = v
    End If
End Function

Public Function setTempTable(tempTableName, alias)
    p_tempTableName = tempTableName
    p_tempTableAlias = alias
End Function

Public Function addColumn(columnName, columnAlias, homeTable)
    Call p_Columns.add(columnName, columnAlias, homeTable)
End Function

'Public Function addSubSelectColumn(columnName, queryID)
'    Call p_columns.AddSelect(columnName, queryID)
'End Function

Public Function addTable(tableName, tableAlias, tableType, SourceQueryID, joinType)
    'If tableType = "subquery" Then
    '    Call addSubQuery(sourceQueryID)
    'End If
    Call p_Tables.add(tableName, tableAlias, tableType, SourceQueryID, joinType)
End Function

Public Function addJoinData(compareType, l_Param, l_Alias, compareWord, r_Param, r_Alias)
    Call p_Joins.addJoinOnData(compareType, l_Param, l_Alias, compareWord, r_Param, r_Alias)
End Function

Public Function addWhereData(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
    Call p_Wheres.addWhereData(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
End Function

Public Function addWhereBetweenData(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
    Call p_Wheres.addWhereBetweenData(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
End Function
