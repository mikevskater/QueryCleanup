VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Column"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Column #

Private p_alias As String

Private p_homeColumn As String
Private p_homeTable As String

Public Function add(columnText, columnAlias, homeTable)
    p_homeColumn = columnText
    
    If Mid(homeTable, 1, 1) = "[" Then
        homeTable = Mid(homeTable, 2, Len(homeTable) - 1)
        homeTable = Mid(homeTable, 1, Len(homeTable) - 1)
    End If
    p_homeTable = homeTable
    
    If Mid(columnAlias, 1, 1) = "[" Then
        columnAlias = Mid(columnAlias, 2, Len(columnAlias) - 1)
        columnAlias = Mid(columnAlias, 1, Len(columnAlias) - 1)
    End If
    p_alias = columnAlias
End Function

Public Property Get alias() As String
    alias = p_alias
End Property

Public Property Get column() As String
    column = p_homeColumn
End Property

Public Property Get table() As String
    table = p_homeTable
End Property
