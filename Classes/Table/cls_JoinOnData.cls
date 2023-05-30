VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_JoinOnData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_JoinOnData #

Private p_left_SourceAlias As String
Private p_left_SourceParamater As String
Private p_right_SourceAlias As String
Private p_right_SourceParamater As String
Private p_compareWord As String
Private p_compareType As String

Property Get left_SourceAlias() As String
    left_SourceAlias = p_left_SourceAlias
End Property

Property Let left_SourceAlias(v As String)
    p_left_SourceAlias = v
End Property

Property Get right_SourceAlias() As String
    right_SourceAlias = p_right_SourceAlias
End Property

Property Let right_SourceAlias(v As String)
    p_right_SourceAlias = v
End Property

Property Get left_SourceParamater() As String
    left_SourceParamater = p_left_SourceParamater
End Property

Property Get right_SourceParamater() As String
    right_SourceParamater = p_right_SourceParamater
End Property

Property Get compareWord() As String
    compareWord = p_compareWord
End Property

Property Get compareType() As String
    compareType = p_compareType
End Property

Public Function addJoinSources(compareType, l_Param, l_Alias, compareWord, r_Param, r_Alias)
    p_left_SourceAlias = Trim(l_Alias)
    p_left_SourceParamater = Trim(l_Param)
    p_right_SourceAlias = Trim(r_Alias)
    p_right_SourceParamater = Trim(r_Param)
    p_compareWord = Trim(compareWord)
    p_compareType = compareType
End Function
