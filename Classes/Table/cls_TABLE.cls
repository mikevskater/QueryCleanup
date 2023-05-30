VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_TABLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Table")
'# cls_Table #

Private p_Init As Boolean

Private p_SourceType As String

Private p_SourceName As String
Private p_SourceAlias As String
Private p_SourceQueryID As Integer
Private p_JoinType As String

Private Sub Class_Initialize()
    p_Init = True
End Sub

Property Let sourceType(v As String)
    p_SourceType = v
End Property

Property Get sourceType() As String
    sourceType = p_SourceType
End Property

Property Let SourceName(v As String)
    p_SourceName = v
End Property

Property Get SourceName() As String
    SourceName = p_SourceName
End Property

Property Let SourceAlias(v As String)
    p_SourceAlias = v
End Property

Property Get SourceAlias() As String
    SourceAlias = p_SourceAlias
End Property

Property Let SourceQueryID(v As Integer)
    p_SourceQueryID = v
End Property

Property Get SourceQueryID() As Integer
    SourceQueryID = p_SourceQueryID
End Property

Property Let joinType(v As String)
    p_JoinType = v
End Property

Property Get joinType() As String
    joinType = p_JoinType
End Property

Property Get name() As String
    name = p_SourceName
End Property

Property Get alias() As String
    alias = p_SourceAlias
End Property

Property Get QueryID() As String
    QueryID = p_SourceQueryID
End Property

