VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JoinBranchStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Join")
'
Private dictData As DictJoinData
Private BranchTree As JoinBranchStructureSubBranch
Private treeStruct() As String
Private l_Table As String
Private r_Table As String
Private lastTreeCheck

Private Sub Class_Initialize()
    Set dictData = New DictJoinData
    Set BranchTree = New JoinBranchStructureSubBranch
    ReDim treeStruct(0 To 0)
    treeStruct(0) = Empty
    dictData.Setup
End Sub

Public Function getDictElement(tableName) As dict_JoinConnections
    Dim a
    For a = 0 To UBound(dictData.dicts)
        If dictData.dict(a).table_name = tableName Then
            getDictElement = dictData.dict(a)
        End If
    Next a
End Function

Public Sub createBranch(l_name, r_Name)
    Dim i As Long, a, b, c
    l_Table = l_name
    r_Table = r_Name
    treeStruct(0) = l_name
    
    Call BranchTree.createBranch(l_name, dictData, treeStruct, 1, 1)
    
    lastTreeCheck = UBound(treeStruct)
    'Call sortStrReg(treeStruct)
    For a = 0 To lastTreeCheck
        If treeStruct(a) = r_Name Then
            Exit Sub
        End If
    Next a
    
    For i = 1 To (c_MaxBranchCount - 1)
        
        Call BranchTree.subBranch(dictData, treeStruct, 1, i)
        Call sortStrReg(treeStruct)
        If lastTreeCheck + 1 <= UBound(treeStruct) Then
            For a = lastTreeCheck + 1 To UBound(treeStruct)
                If treeStruct(a) = r_Name Then
                    Exit Sub
                End If
            Next a
        End If
        lastTreeCheck = UBound(treeStruct)
        
    Next i
End Sub

Public Function getBranch() As dict_SubBranchResult
    Dim a As dict_SubBranchResult, b, c, d, e, f, g
    Dim x, y, z, xx, yy, zz
    
    a = BranchTree.getSubBranch(treeStruct, r_Table)
    
    If a.finalString = "" Then
        Debug.Print "Branch not found!"
        getBranch = a
    Else
        'a = l_Table & " -> " & a
        Debug.Print ""
        Debug.Print ""
        Debug.Print l_Table & " to " & r_Table
        Debug.Print a.finalString
        getBranch = a
    End If
    
End Function
