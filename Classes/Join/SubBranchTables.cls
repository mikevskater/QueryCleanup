VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubBranchTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Join")

Private p_table_name
Private p_joined_On_Data() As dict_ColumnJoins
Private p_joined_On_UBound As Long

Private p_subTables() As JoinBranchStructureSubBranch
Private p_subTables_UBound As Long

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

Private Sub Class_Initialize()
    ReDim p_joined_On_Data(-1 To 0)
    p_joined_On_UBound = -1
    
    ReDim p_subTables(-1 To 0)
    p_subTables_UBound = -1
End Sub

Property Get table_name()
    table_name = p_table_name
End Property
Property Let table_name(v)
    p_table_name = v
End Property

Property Get joined_On_Datas() As dict_ColumnJoins()
    joined_On_Datas = p_joined_On_Data
End Property
Property Let joined_On_Datas(v() As dict_ColumnJoins)
    p_joined_On_Data = v
End Property

Property Get joined_On_Data(i) As dict_ColumnJoins
    joined_On_Data = p_joined_On_Data(i)
End Property

Property Let joined_On_Data(i, v As dict_ColumnJoins)
    p_joined_On_Data(i) = v
End Property

Property Get joined_UBound()
    joined_UBound = UBound(p_joined_On_Data)
End Property

Public Function subTables() As JoinBranchStructureSubBranch()
    subTables = p_subTables
End Function

Property Get subtable(v) As JoinBranchStructureSubBranch
    Set subtable = p_subTables(v)
End Property

Property Get subTablesUBound()
    subTablesUBound = p_subTables_UBound
End Property

Public Function addSubTable(tableName) ', dictData, ByRef treeStruct() As String, i As Long, ii As Long)
    If p_subTables_UBound = -1 Then
        p_subTables_UBound = 0
        ReDim p_subTables(0 To p_subTables_UBound)
    Else
        p_subTables_UBound = p_subTables_UBound + 1
        ReDim Preserve p_subTables(0 To p_subTables_UBound)
    End If
    
    Set p_subTables(p_subTables_UBound) = New JoinBranchStructureSubBranch
    p_subTables(p_subTables_UBound).table_name = tableName
    'Call p_subTables(p_subTables_UBound).createBranch(tableName, dictData, treeStruct, i, ii)
End Function
