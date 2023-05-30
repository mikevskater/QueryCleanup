VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Joins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Joins #

Private p_joinOnArray() As cls_JoinOnData
Private p_Init As Boolean
Private p_joinOnArrayCount As Integer

Property Get init_()
    init_ = p_Init
End Property

Property Get joinCount() As Integer
    joinCount = p_joinOnArrayCount
End Property

Public Function count() As Integer
    If LBound(p_joinOnArray) = -1 Then
        count = -1
    Else
        count = UBound(p_joinOnArray) + 1
    End If
End Function

Public Function joins() As cls_JoinOnData()
    joins = p_joinOnArray
End Function

Public Function join(i As Integer) As cls_JoinOnData
    Set join = p_joinOnArray(i)
End Function

Private Sub Class_Initialize()
    ReDim p_joinOnArray(-1 To 0)
    p_Init = False
    p_joinOnArrayCount = -1
End Sub

Public Function addJoinOnData(compareType, l_Param, l_Alias, compareWord, r_Param, r_Alias)
    If p_Init = False Then
        p_Init = True
        ReDim p_joinOnArray(0 To 0)
        Set p_joinOnArray(0) = New cls_JoinOnData
        p_joinOnArrayCount = 0
    Else
        ReDim Preserve p_joinOnArray(0 To p_joinOnArrayCount + 1)
        p_joinOnArrayCount = p_joinOnArrayCount + 1
        Set p_joinOnArray(p_joinOnArrayCount) = New cls_JoinOnData
    End If
    Call p_joinOnArray(p_joinOnArrayCount).addJoinSources(compareType, l_Param, l_Alias, compareWord, r_Param, r_Alias)
End Function
