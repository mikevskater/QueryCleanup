VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Wheres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Wheres #


Private p_Init As Boolean
Private p_WhereCount As Integer
Private p_Wheres() As cls_Where



Private Function pushWhere()
    If p_Init = False Then
        p_Init = True
        ReDim p_Wheres(0 To 0)
        Set p_Wheres(0) = New cls_Where
        p_WhereCount = 0
    Else
        p_WhereCount = p_WhereCount + 1
        ReDim Preserve p_Wheres(0 To p_WhereCount)
        Set p_Wheres(p_WhereCount) = New cls_Where
    End If
End Function

Public Function count() As Integer
    count = p_WhereCount
End Function

Public Function wheres() As cls_Where()
    wheres = p_Wheres
End Function

Public Function where(i As Integer) As cls_Where
    Set where = p_Wheres(i)
End Function

Public Function addWhereBetweenData(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
    pushWhere
    Call p_Wheres(p_WhereCount).addBetween(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
End Function

Public Function addWhereData(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
    pushWhere
    Call p_Wheres(p_WhereCount).addWhere(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
End Function

Private Sub Class_Initialize()
    p_WhereCount = -1
End Sub
