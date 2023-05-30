Attribute VB_Name = "md_GLOBALS_TYPES"
'@Folder("Mods.Globals")
'# md_GLOBALS_TYPES #

Public Type fn_inArrayResult
    found As Boolean
    foundIndex As Integer
End Type

Public Type qr_nextWord
    wordStart As Integer
    wordEnd As Integer
    word As String
    len As Integer
    EOF As Boolean
End Type

Public Type qr_tableAliasData
    alias As String
    homeQueryIndex As String
    count As Long
End Type

Public Type qr_getIntoData
    intoType As String
    isTempTable As Boolean
    tableName As String
    tableAlias As String
    tableColumns() As String
End Type

Public Type qr_searchStringResults
    word As String
    'Sort by maxLengthOfMatch then firstSpotOfMatch then countOfMatchedLetters then firstSpotOfLetterMatch then lengthOfWord
    match_maxLenthOfMatch As Integer
    match_firstSpotOfMatch As Integer
    match_countOfMatchedLetters As Integer
    match_firstSpotOfLetterMatch As Integer
    lengthOfWord As Integer
    exactMatch As Boolean
End Type

Public Type qr_tableSearchData
    tableSearchResults As qr_searchStringResults
    aliasSearchResults As qr_searchStringResults
    homeQueryIndex As Integer
    aliasList() As qr_tableAliasData
End Type

Public Type qr_columnMainSearchData
    columnSearchResults As qr_searchStringResults
    aliasSearchResults As qr_searchStringResults
    columnAlias As String
    columnHomeTable As String
    homeQueryIndex As Variant
    tableData As qr_tableSearchData
    count As Integer
End Type

Private Type qr_columnData
    name As String
    alias As String
    homeTable As String
    count As Long
End Type

Private Type qr_joinData
    compareType As String
    compareWord As String
    l_Alias As String
    l_Param As String
    r_Alias As String
    r_Param As String
    count As Long
End Type

Private Type qr_whereData
    type As String
    compareType As String
    compareWord As String
    where_L_Alias As String
    where_L_Param As String
    where_R_Alias As String
    where_R_Param As String
    between_col_Alias As String
    between_Col_Name As String
    between_L_Compare As String
    between_R_Compare As String
    count As Long
End Type

Public Type qr_tableMainSearchData
    columns() As qr_columnData
    joins() As qr_joinData
    wheres() As qr_whereData
    tableData As qr_tableSearchData
    count As Integer
End Type

Public Type jn_JoinColumnData
    columnNameL As String
    columnNameR As String
    joinType As String
    count As Long
End Type

Public Type qr_JoinPackage
    onColumns() As jn_JoinColumnData
    tableName As String
    count As Long
End Type

Public Type qr_JoinMainSearchData
    searchWord As String
    tableData As qr_tableSearchData
    joins() As qr_JoinPackage
End Type

Public Type qr_JoinFinalResults
    mainTable As qr_tableSearchData
    joinedTables() As qr_JoinPackage
End Type

'==============================================================================================================================
' Join Branch Types
'==============================================================================================================================
Public Type qr_JoinBranch_Main_SearchData
    table_Start As String
    joinArray() As String
    table_End As String
End Type

Public Type qr_JoinBranch_Main_TableDataJoinData
    main_columnName As String
    join_tableName As String
    join_columnName As String
    joinType As String
    count As Long
End Type

Public Type qr_JoinBranch_Main_TableData
    name As String
    joins() As qr_JoinBranch_Main_TableDataJoinData
    count As Long
    queryIndex() As Long
End Type


'==============================================================================================================================
' Join Dictionary
'==============================================================================================================================

Public Type dict_ColumnJoins
    join_Type As String
    joinet_Table_Column As String
    join_Comparison As String
    joined_Table_Column As String
    count As Long
End Type

Public Type dict_JoinData
    joined_Table_Name As String
    joined_On() As dict_ColumnJoins
End Type

Public Type dict_TableJoinsSubBranch
    table_name As String
    joined_Tables() As dict_JoinData
End Type

Public Type dict_TableJoins
    joined_Table_Name As String
    joined_On() As dict_ColumnJoins
    subBranch() As dict_TableJoinsSubBranch
    count As Long
End Type

Public Type dict_JoinConnections
    table_name As String
    joined_Tables() As dict_TableJoins
    count As Long
End Type

Public Type dict_JoinBranch
    start_Table As dict_JoinConnections
End Type

Public Type dict_SubBranchResult
    finalString As String
    arr() As dict_JoinConnections
    objArr() As Variant 'JoinBranchStructureSubBranch
    objArrCnt As Long
End Type
