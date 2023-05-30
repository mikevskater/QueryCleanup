VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Columns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Columns #

Private p_Init As Boolean
Private p_Columns() As cls_Column

Private Sub Class_Initialize()
    ReDim p_Columns(-1 To 0)
    p_Init = False
End Sub

Public Function count() As Integer
    If LBound(p_Columns) = -1 Then
        count = -1
    Else
        count = UBound(p_Columns) + 1
    End If
End Function

Public Function add(columnText, columnAlias, homeTable)
    If p_Init = False Then
        p_Init = True
        ReDim p_Columns(0 To 0)
        Set p_Columns(0) = New cls_Column
    Else
        ReDim Preserve p_Columns(0 To UBound(p_Columns) + 1)
        Set p_Columns(UBound(p_Columns)) = New cls_Column
    End If
    Call p_Columns(UBound(p_Columns)).add(columnText, columnAlias, homeTable)
End Function

Public Function columns() As cls_Column()
    columns = p_Columns
End Function

Public Function column(i As Integer) As cls_Column
    Set column = p_Columns(i)
End Function
