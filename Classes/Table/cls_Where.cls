VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Where"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Where #

Private p_compareType As String
Private p_whereType As String
Private p_compareWord As String

Private p_Between_col_Name As String
Private p_Between_col_Alias As String
Private p_Between_l_Compare As String
Private p_Between_r_Compare As String

Private p_Where_l_col_Name As String
Private p_Where_l_col_Alias As String
Private p_Where_r_col_Name As String
Private p_Where_r_col_Alias As String

Public Property Get compareType() As String
    compareType = p_compareType
End Property

Public Property Get whereType() As String
    whereType = p_whereType
End Property

Public Property Get compareWord() As String
    compareWord = p_compareWord
End Property

Public Property Get between_Col_Name() As String
    between_Col_Name = p_Between_col_Name
End Property

Public Property Get between_col_Alias() As String
    between_col_Alias = p_Between_col_Alias
End Property

Public Property Get between_L_Compare() As String
    between_L_Compare = p_Between_l_Compare
End Property

Public Property Get between_R_Compare() As String
    between_R_Compare = p_Between_r_Compare
End Property

Public Property Get Where_l_col_Name() As String
    Where_l_col_Name = p_Where_l_col_Name
End Property

Public Property Get Where_l_col_Alias() As String
    Where_l_col_Alias = p_Where_l_col_Alias
End Property

Public Property Get Where_r_col_Name() As String
    Where_r_col_Name = p_Where_r_col_Name
End Property

Public Property Get Where_r_col_Alias() As String
    Where_r_col_Alias = p_Where_r_col_Alias
End Property

Public Function addBetween(where_Type, col_Name, col_Alias, compareWord, l_Compare, r_Compare)
    p_compareType = "between"
    p_whereType = where_Type
    
    p_Between_col_Name = col_Name
    p_Between_col_Alias = col_Alias
    p_compareWord = compareWord
    p_Between_l_Compare = l_Compare
    p_Between_r_Compare = r_Compare
End Function

Public Function addWhere(where_Type, l_col_Name, l_col_Alias, compareWord, r_col_Name, r_col_Alias)
    p_compareType = "where"
    p_whereType = where_Type
    
    p_Where_l_col_Name = l_col_Name
    p_Where_l_col_Alias = l_col_Alias
    p_compareWord = compareWord
    p_Where_r_col_Name = r_col_Name
    p_Where_r_col_Alias = r_col_Alias
End Function
