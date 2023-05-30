VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_QueryManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Classes.Query.Manager")
'# cls_QueryManager #


Public Function processNextWord(word_)
    Dim i, j, x, y
    Dim loop_
    Dim currentWord As qr_nextWord, nextWord As qr_nextWord, previousWord As qr_nextWord, compareWord As qr_nextWord
    Dim endOfCurrentWord As qr_nextWord
    Dim endOfNextWord As qr_nextWord
    Dim tempWord As qr_nextWord
    Dim l_Word, r_Word
    Dim l_name, l_Alias, r_Name, r_Alias
    
    
    Select Case getNextWordStep(word_)
    
        Case "start_select"
            qryTree.createNewQuery
            
            Dim columnRangeStartWord As qr_nextWord, columnRangeEndWord As qr_nextWord, intoWordLoc As qr_nextWord
            columnRangeStartWord = qryReader.currentWord
            
            columnRangeEndWord = qryReader.scanForCurrentBranchNextKeyWord("from")
            intoWordLoc = qryReader.scanForCurrentBranchNextKeyWord("into")
            
            If intoWordLoc.wordStart > -1 Then
                Call parseColumns(columnRangeStartWord, intoWordLoc)
                
                Dim intoData As qr_getIntoData
                intoData = getIntoData(intoWordLoc, columnRangeEndWord)
                
                Call qryTree.setCurrentQueryTempTable(intoData.tableName, intoData.tableAlias)
            Else
                Call parseColumns(columnRangeStartWord, columnRangeEndWord)
            End If
            
            qryReader.currentWord_WordEnd = columnRangeEndWord.wordStart - 1
            
        Case "start_delete"
            qryTree.createNewQuery
            currentWord = qryReader.getNextWord(True)
            nextWord = qryReader.previewNextWord(True)
            
            If isSqlKeyWord(currentWord.word) = False Then
                If nextWord.EOF = False Then
                    If isSqlKeyWord(nextWord.word) = False Then
                        Call qryTree.addTable(currentWord.word, nextWord.word, "delete")
                        currentWord = qryReader.getNextWord(True)
                    Else
                        Call qryTree.addTable(currentWord.word, currentWord.word, "delete")
                    End If
                Else
                    Call qryTree.addTable(currentWord.word, currentWord.word, "delete")
                End If
            Else
                
            End If
        Case "query_from"
            loop_ = True
            
            Do
                currentWord = qryReader.getNextWord(True)
                nextWord = qryReader.previewNextWord(True)
                
                If InStr(1, currentWord.word, ")") > 0 And InStr(1, currentWord.word, "(") = 0 Then
                    Call qryTree.addTable(currentWord.word, currentWord.word)
                    Exit Do
                End If
                
                If isSqlKeyWord(currentWord.word) = False Then
                    If nextWord.EOF = False Then
                        If isSqlKeyWord(nextWord.word) = False Then
                            Call qryTree.addTable(currentWord.word, nextWord.word, "from")
                            currentWord = qryReader.getNextWord(True)
                        Else
                            Call qryTree.addTable(currentWord.word, currentWord.word, "from")
                        End If
                    Else
                        Call qryTree.addTable(currentWord.word, currentWord.word, "from")
                    End If
                Else
                    If isSqlJoinWord(currentWord.word) = True Then
                        Dim joinType As qr_nextWord
                        If currentWord.word = "outer" Then
                            currentWord = currentWord
                        End If
                        If inSqlJoinWord(nextWord.word) = True Then
                            Dim tempCurWord As qr_nextWord
                            tempCurWord = qryReader.getNextWord
                            joinType.word = currentWord.word & " " & tempCurWord.word
                        Else
                            joinType = currentWord
                        End If
                        currentWord = qryReader.getNextWord(True)
                        nextWord = qryReader.previewNextWord(True)
                        If currentWord.word = "(" Then
                            Dim endPara As qr_nextWord
                            
                            endPara = qryReader.scanInCurrentBranchNextKeyWord(")")
                            Do
                                currentWord = qryReader.getNextWord(True)
                                Call qryManager.processNextWord(currentWord.word)
                            Loop Until (qryReader.currentWord.wordEnd >= endPara.wordEnd)
                            qryTree.closeOffQuery
                            
                            currentWord = qryReader.getNextWord(True)
                            Call qryTree.addSubQuery(currentWord.word)
                            
                            
                        Else
                            If isSqlKeyWord(nextWord.word) = False Then
                                Call qryTree.addTable(currentWord.word, nextWord.word, joinType.word)
                                currentWord = qryReader.getNextWord(True)
                            Else
                                Call qryTree.addTable(currentWord.word, currentWord.word, joinType.word)
                            End If
                        End If
                    Else
                        If currentWord.word = "on" Or currentWord.word = "and" Or currentWord.word = "or" Then
                            Dim compareType As String
                            compareType = currentWord.word
                            
                            currentWord = qryReader.getNextWord(True)
                            
                            If currentWord.word = "(" Then
                                endOfCurrentWord = qryReader.scanInCurrentBranchNextKeyWord(")")
                                currentWord.wordEnd = endOfCurrentWord.wordEnd
                                currentWord.word = qryReader.getTextBetweenWords(currentWord, endOfCurrentWord)
                                qryReader.currentWord = currentWord
                            End If
                                                                                    
                            compareWord = qryReader.scanForCurrentBranchNextKeyWord(getCompareOpperators)
                            If compareWord.word = "not" Then
                                tempWord = qryReader.scanForCurrentBranchNextKeyWord(getCompareOpperators)
                                If isSqlCompareWord(tempWord.word) = True Then
                                    compareWord.word = compareWord.word & " " & tempWord.word
                                    compareWord.len = -1
                                End If
                            End If
                            
                            If compareWord.len > -1 Then
                                qryReader.currentWord = compareWord
                            End If
                            
                            nextWord = qryReader.previewNextWord(True)
                            If nextWord.word = "(" Then
                                qryReader.currentWord_WordEnd = nextWord.wordEnd
                                endOfNextWord = qryReader.scanInCurrentBranchNextKeyWord(")")
                                nextWord.wordEnd = endOfNextWord.wordEnd
                                nextWord.word = qryReader.getTextBetweenWords(nextWord, endOfNextWord)
                                qryReader.currentWord = nextWord
                            Else
                                nextWord = qryReader.getNextWord(True)
                            End If
                            
                            If inSqlCompareWord(compareWord.word) = True Then
                                
                                currentWord.word = Trim(currentWord.word)
                                If IsArray(splitInBrackets(currentWord.word, " ")) = False Then
                                    l_Word = splitInBrackets(currentWord.word, ".")
                                Else
                                    l_Word = Array(currentWord.word)
                                End If
                                
                                nextWord.word = Trim(nextWord.word)
                                If IsArray(splitInBrackets(nextWord.word, " ")) = False Then
                                    r_Word = splitInBrackets(nextWord.word, ".")
                                Else
                                    r_Word = Array(nextWord.word)
                                End If
                                
                                If IsArray(l_Word) = True Then
                                    If UBound(l_Word) > 0 Then
                                        l_Alias = l_Word(0)
                                        l_name = l_Word(1)
                                    Else
                                        l_Alias = l_Word(0)
                                        l_name = l_Word(0)
                                    End If
                                Else
                                    l_Alias = l_Word
                                    l_name = l_Word
                                End If
                                
                                If IsArray(r_Word) = True Then
                                    If UBound(r_Word) > 0 Then
                                        r_Alias = r_Word(0)
                                        r_Name = r_Word(1)
                                    Else
                                        r_Alias = r_Word(0)
                                        r_Name = r_Word(0)
                                    End If
                                Else
                                    r_Alias = r_Word
                                    r_Name = r_Word
                                End If
                                
                                Call qryTree.addJoinOnData(compareType, l_name, l_Alias, compareWord.word, r_Name, r_Alias)
                                'currentWord = qryReader.getNextWord
                                'currentWord = qryReader.getNextWord
                            Else
                                Call qryTree.addJoinOnData(compareType, currentWord.word, currentWord.word, compareWord.word, nextWord.word, nextWord.word)
                                loop_ = False
                            End If
                        Else
                            qryReader.currentWord_WordEnd = currentWord.wordStart - 1
                            
                            Dim statementDone As qr_nextWord
                            statementDone = qryReader.scanForCurrentBranchNextKeyWord("where", "group", "order")
                            
                            If statementDone.wordStart = -1 Then
                                qryTree.closeOffQuery
                            End If
                            
                            loop_ = False
                        End If
                    End If
                End If
                previousWord = currentWord
                
                If qryReader.previewNextWord.EOF = True Then loop_ = False
                
            Loop Until loop_ = False
            i = i
            
        Case "query_where"
            i = i
            loop_ = True
            Dim whereType As qr_nextWord
            whereType.word = word_
            Do
                If whereType.word = "where" Or whereType.word = "and" Or whereType.word = "or" Then
                    currentWord = qryReader.getNextWord(False)
                    If currentWord.word = "(" Then
                        endOfCurrentWord = qryReader.scanInCurrentBranchNextKeyWord(")")
                        currentWord.wordEnd = endOfCurrentWord.wordEnd
                        currentWord.word = qryReader.getTextBetweenWords(currentWord, endOfCurrentWord)
                        qryReader.currentWord = currentWord
                    End If
                    
                    If currentWord.word = "not" Then
                        tempWord = qryReader.previewNextWord(True)
                        If tempWord.word = "exists" Then
                            tempWord = qryReader.getNextWord(True)
                            currentWord.word = "not exists"
                            
                            'TODO HANDLE A SUBSELECT
                            
                        End If
                    ElseIf currentWord.word = "exists" Then
                    
                        'TODO HANDLE A SUBSELECT
                    
                    ElseIf isSqlKeyWord(currentWord.word) = True Then
                        If currentWord.word = "group" Or currentWord.word = "order" Then
                            Call qryManager.processNextWord(currentWord.word)
                        End If
                        qryTree.closeOffQuery
                        loop_ = False
                        
                    Else
                        compareWord = qryReader.scanForCurrentBranchNextKeyWord(getCompareOpperators)
                        If compareWord.word = "not" Then
                            qryReader.currentWord = compareWord
                            tempWord = qryReader.scanForCurrentBranchNextKeyWord(getCompareOpperators)
                            If isSqlCompareWord(tempWord.word) = True Then
                                compareWord.word = compareWord.word & " " & tempWord.word
                                compareWord.len = -1
                                qryReader.currentWord = tempWord
                            End If
                        End If
                        
                        If compareWord.len > -1 Then
                            qryReader.currentWord = compareWord
                        End If
                        
                        If InStr(1, compareWord.word, "between") > 0 Then
                            Dim tempThirdWord As qr_nextWord
                            nextWord = qryReader.previewNextWord(True)
                            If nextWord.word = "(" Then
                                qryReader.currentWord_WordEnd = nextWord.wordEnd
                                endOfNextWord = qryReader.scanInCurrentBranchNextKeyWord(")")
                                nextWord.wordEnd = endOfNextWord.wordEnd
                                nextWord.word = qryReader.getTextBetweenWords(nextWord, endOfNextWord)
                                qryReader.currentWord = nextWord
                            Else
                                nextWord = qryReader.getNextWord(True)
                            End If
                            tempThirdWord = qryReader.getNextWord(True)
                            tempThirdWord = qryReader.getNextWord(True)
                                              
                            currentWord.word = Trim(currentWord.word)
                            If IsArray(splitInBrackets(currentWord.word, " ")) = False Then
                                l_Word = splitInBrackets(currentWord.word, ".")
                            Else
                                l_Word = Array(currentWord.word)
                            End If
                            
                            If IsArray(l_Word) = True Then
                                If UBound(l_Word) > 0 Then
                                    l_Alias = l_Word(0)
                                    l_name = l_Word(1)
                                Else
                                    l_Alias = l_Word(0)
                                    l_name = l_Word(0)
                                End If
                            Else
                                l_Alias = l_Word
                                l_name = l_Word
                            End If
                            
                            Call qryTree.addWhereBetween(Empty, l_name, l_Alias, compareWord.word, nextWord.word, tempThirdWord.word)
                            
                        Else
                            nextWord = qryReader.previewNextWord(False)
                            If nextWord.word = "(" Then
                                qryReader.currentWord_WordEnd = nextWord.wordEnd
                                endOfNextWord = qryReader.scanInCurrentBranchNextKeyWord(")")
                                nextWord.wordEnd = endOfNextWord.wordEnd
                                nextWord.word = qryReader.getTextBetweenWords(nextWord, endOfNextWord)
                                qryReader.currentWord = nextWord
                            Else
                                nextWord = qryReader.getNextWord(True)
                            End If
                                                    
                            currentWord.word = Trim(currentWord.word)
                            If IsArray(splitInBrackets(currentWord.word, " ")) = False Then
                                l_Word = splitInBrackets(currentWord.word, ".")
                            Else
                                l_Word = Array(currentWord.word)
                            End If
                            
                            nextWord.word = Trim(nextWord.word)
                            If IsArray(splitInBrackets(nextWord.word, " ")) = False Then
                                r_Word = splitInBrackets(nextWord.word, ".")
                            Else
                                r_Word = Array(nextWord.word)
                            End If
                            
                            If IsArray(l_Word) = True Then
                                If UBound(l_Word) > 0 Then
                                    l_Alias = l_Word(0)
                                    l_name = l_Word(1)
                                Else
                                    l_Alias = l_Word(0)
                                    l_name = l_Word(0)
                                End If
                            Else
                                l_Alias = l_Word
                                l_name = l_Word
                            End If
                            
                            If IsArray(r_Word) = True Then
                                If UBound(r_Word) > 0 Then
                                    r_Alias = r_Word(0)
                                    r_Name = r_Word(1)
                                Else
                                    r_Alias = r_Word(0)
                                    r_Name = r_Word(0)
                                End If
                            Else
                                r_Alias = r_Word
                                r_Name = r_Word
                            End If
                            
                            Call qryTree.addWhere(Empty, l_name, l_Alias, compareWord.word, r_Name, r_Alias)
                        End If
                        
                    End If
                Else
                    loop_ = False
                End If
                    
                whereType = qryReader.previewNextWord(True)
                If whereType.word = "where" Or whereType.word = "and" Or whereType.word = "or" Then
                    whereType = qryReader.getNextWord(True)
                End If
                
                If whereType.EOF = True Then loop_ = False
            Loop Until (loop_ = False)
            i = i
        Case "query_group"
            i = i
            
            
        Case "query_having"
            i = i
            
            
        Case "query_order"
            i = i
            
            
        Case Else
            
    End Select
End Function

Public Function checkWordGroupType(word_)
    Select Case word_
        Case "select", "delete", "insert"
        
        Case "from", "into"
        
        Case Else
            
    End Select
End Function

Public Function getNextWordStep(word_)
    Dim r As Variant
    r = Empty
    Select Case word_
        'Case "add"

        'Case "all"
        
        'Case "alter"
        
        'Case "and"
        
        'Case "any"
        
        'Case "as"
        
        'Case "asc"
        
        'Case "authorization"
        
        'Case "backup"
        
        'Case "begin"
        
        'Case "between"
        
        'Case "break"
        
        'Case "browse"
        
        'Case "bulk"
        
        'Case "by"
        
        'Case "cascade"
        
        'Case "case"
        
        'Case "check"
        
        'Case "checkpoint"
        
        'Case "close"
        
        'Case "clustered"
        
        'Case "coalesce"
        
        'Case "collate"
        
        'Case "column"
        
        'Case "commit"
        
        'Case "compute"
        
        'Case "constraint"
        
        'Case "contains"
        
        'Case "containstable"
        
        'Case "continue"
        
        'Case "convert"
        
        'Case "create"
        
        'Case "cross"
        
        'Case "current"
        
        'Case "current_date"
        
        'Case "current_time"
        
        'Case "current_timestamp"
        
        'Case "current_user"
        
        'Case "cursor"
        
        'Case "database"
        
        'Case "dbcc"
        
        'Case "deallocate"
        
        'Case "declare"
        
        'Case "default"
        
        Case "delete"
            r = "start_delete"
            
        'Case "deny"
        
        'Case "desc"
        
        'Case "disk"
        
        'Case "distinct"
        
        'Case "distributed"
        
        'Case "double"
        
        'Case "drop"
        
        'Case "dummy"
        
        'Case "dump"
        
        'Case "else"
        
        'Case "end"
        
        'Case "errlvl"
        
        'Case "escape"
        
        'Case "except"
        
        'Case "exec"
        
        'Case "execute"
        
        'Case "exists"
        
        'Case "exit"
        
        'Case "fetch"
        
        'Case "file"
        
        'Case "fillfactor"
        
        'Case "for"
        
        'Case "foreign"
        
        'Case "freetext"
        
        'Case "freetexttable"
        
        Case "from"
            r = "query_from"
            
        'Case "full"
        
        'Case "function"
        
        'Case "goto"
        
        'Case "grant"
        
        'Case "group"
        
        'Case "having"
        
        'Case "holdlock"
        
        'Case "identity"
        
        'Case "identity_insert"
        
        'Case "identitycol"
        
        'Case "if"
        
        'Case "in"
        
        'Case "index"
        
        'Case "inner"
        
        'Case "insert"
        
        'Case "intersect"
        
        'Case "into"
        
        'Case "is"
        
        Case "join"
            Select Case qryReader.lastWord.word
                Case "left"
                    r = "query_leftjoin"
                Case "right"
                    r = "query_rightjoin"
                Case "inner"
                    r = "query_innerjoin"
                Case Else
                    r = "query_join"
            End Select
            
        'Case "key"
        
        'Case "kill"
        
        'Case "left"
        
        'Case "like"
        
        'Case "lineno"
        
        'Case "load"
        
        'Case "national"
        
        'Case "nocheck"
        
        'Case "nonclustered"
        
        'Case "not"
        
        'Case "null"
        
        'Case "nullif"
        
        'Case "of"
        
        'Case "off"
        
        'Case "offsets"
        
        Case "on"
            r = "query_on"
            
        'Case "open"
        
        'Case "opendatasource"
        
        'Case "openquery"
        
        'Case "openrowset"
        
        'Case "openxml"
        
        'Case "option"
        
        'Case "or"
        
        'Case "order"
        
        'Case "outer"
        
        'Case "over"
        
        'Case "percent"
        
        'Case "plan"
        
        'Case "precision"
        
        'Case "primary"
        
        'Case "print"
        
        'Case "proc"
        
        'Case "procedure"
        
        'Case "public"
        
        'Case "raiserror"
        
        'Case "read"
        
        'Case "readtext"
        
        'Case "reconfigure"
        
        'Case "references"
        
        'Case "replication"
        
        'Case "restore"
        
        'Case "restrict"
        
        'Case "return"
        
        'Case "revoke"
        
        'Case "right"
        
        'Case "rollback"
        
        'Case "rowcount"
        
        'Case "rowguidcol"
        
        'Case "rule"
        
        'Case "save"
        
        'Case "schema"
        
        Case "select"
            r = "start_select"
            
        'Case "session_user"
        
        'Case "set"
        
        'Case "setuser"
        
        'Case "shutdown"
        
        'Case "some"
        
        'Case "statistics"
        
        'Case "system_user"
        
        'Case "table"
        
        'Case "textsize"
        
        'Case "then"
        
        'Case "to"
        
        'Case "top"
        
        'Case "tran"
        
        'Case "transaction"
        
        'Case "trigger"
        
        'Case "truncate"
        
        'Case "tsequal"
        
        Case "union"
            Select Case qryReader.previewNextWord.word
                Case "all"
                    r = qryReader.getNextWord.word
                    r = "query_unionall"
                Case Else
                    r = "query_union"
            End Select
                        
        'Case "unique"
        
        Case "update"
            r = "start_update"
            
        'Case "updatetext"
        
        'Case "use"
        
        'Case "user"
        
        'Case "values"
        
        'Case "varying"
        
        'Case "view"
        
        'Case "waitfor"
        
        'Case "when"
        
        Case "where"
            r = "query_where"
                
        'Case "while"
        
        'Case "with"
        
        'Case "writetext"
        
        'Case "("
        
        'Case ")"
        
        'Case "["
        
        'Case "]"
        
        'Case "+"
            
        'Case "-"
        
        'Case "/"
        
        'Case "*"
        
        'Case "="
        
        Case Else
            r = "else_nonkeyword"
            
    End Select
    
    getNextWordStep = r
End Function
