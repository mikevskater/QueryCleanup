VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Tables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Tables #

Private p_Init As Boolean
Private p_TableCount As Integer
Private p_Tables() As cls_TABLE

Private Sub Class_Initialize()
    ReDim p_Tables(-1 To 0)
    p_TableCount = -1
    p_Init = False
End Sub

Public Function count() As Integer
    If LBound(p_Tables) = -1 Then
        count = -1
    Else
        count = UBound(p_Tables) + 1
    End If
End Function

Public Function add(tableName, tableAlias, tableType, SourceQueryID, joinType)
    If p_Init = False Then
        p_Init = True
        ReDim p_Tables(0 To 0)
        Set p_Tables(0) = New cls_TABLE
        p_TableCount = 0
    Else
        ReDim Preserve p_Tables(0 To p_TableCount + 1)
        p_TableCount = p_TableCount + 1
        Set p_Tables(p_TableCount) = New cls_TABLE
    End If
    p_Tables(p_TableCount).sourceType = Trim(tableType)
    p_Tables(p_TableCount).SourceName = Trim(tableName)
    p_Tables(p_TableCount).SourceAlias = Trim(tableAlias)
    p_Tables(p_TableCount).SourceQueryID = SourceQueryID
    p_Tables(p_TableCount).joinType = joinType
End Function

Public Function setSourceType(sourceType As String)
    p_Tables(p_TableCount).sourceType = sourceType
End Function

Public Function setSourceTableInfo(tableName As String, tableAlias As String)
    p_Tables(p_TableCount).SourceName = tableName
    p_Tables(p_TableCount).SourceAlias = tableAlias
End Function

Public Function setSourceQueryID(SourceQueryID As Integer)
    p_Tables(p_TableCount).SourceQueryID = SourceQueryID
End Function

Public Function table(i As Integer) As cls_TABLE
    Set table = p_Tables(i)
End Function

Public Function tables() As cls_TABLE()
    tables = p_Tables
End Function
