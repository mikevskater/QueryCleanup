Attribute VB_Name = "md_GLOBALS_VARIABLES"
'@Folder("Mods.Globals")
'# md_GLOBALS_VARIABLES #

'==============================================================================================================================
' CONSTANT VARS
'==============================================================================================================================
Public Const c_MatchStreakPercentile = 1
Public Const c_OutputHeaderRow = 3
Public Const c_MaxBranchCount = 10


'==============================================================================================================================
' GLOBAL VARS
'==============================================================================================================================
Public p_JsonData As Variant
Public p_LastTableSearch() As qr_tableMainSearchData
Public p_LastColumnSearch() As qr_columnMainSearchData
Public p_LastJoinSearch As qr_JoinMainSearchData

'==============================================================================================================================
' DEBUG VARS
'==============================================================================================================================
Public Const debug_Enabled = True
'==============================================================================================================================
Public debug_CurrectSheet As Worksheet
Public debug_LastRunTime As Variant
Public debug_TimeOut As Variant
Public debug_objectArray() As Variant
Public debug_objectCount As Variant

'==============================================================================================================================
' CLASS OBJECTS
'==============================================================================================================================
Public qryReader As cls_QUERYREADER
Public qryManager As cls_QueryManager
Public qryTree As cls_Querys
Public pub_EmptyWord As qr_nextWord

'==============================================================================================================================
' FUNCITON VARIABLES
'==============================================================================================================================
Public Function emptyNextWord() As qr_nextWord
    emptyNextWord.EOF = True
    emptyNextWord.len = -1
    emptyNextWord.word = Empty
    emptyNextWord.wordStart = -1
    emptyNextWord.wordEnd = -1
End Function

Public Function getAllOpperators()
    getAllOpperators = Array("+", _
                             "-", _
                             "*", _
                             "/", _
                             "%", _
                             "&", _
                             "|", _
                             "^", _
                             "=", _
                             ">", _
                             "<", _
                             ">=", _
                             "<=", _
                             "<>", _
                             "!=", _
                             "+=", _
                             "-=", _
                             "*=", _
                             "/=", _
                             "%=", _
                             "&=", _
                             "^-=", _
                             "|*=")
End Function

Public Function getCompareOpperators()
    getCompareOpperators = Array("not", _
                                "exists", _
                                "=", _
                                ">", _
                                "<", _
                                ">=", _
                                "<=", _
                                "<>", _
                                "!=", _
                                "like", _
                                "in", _
                                "between", _
                                "some")
End Function

Public Function getMathOpperators()
    getMathOpperators = Array("+", _
                             "-", _
                             "*", _
                             "/", _
                             "%", _
                             "=", _
                             "!=")
End Function

Public Function getJointOpperatorsCleanUp()
    getJointOpperatorsCleanUp = Array(Array("^ - =", "^-="), _
                             Array("| * =", "|*="), _
                             Array("> =", ">="), _
                             Array("< =", "<="), _
                             Array("< >", "<>"), _
                             Array("! =", "!="), _
                             Array("+ =", "+="), _
                             Array("- =", "-="), _
                             Array("* =", "*="), _
                             Array("/ =", "/="), _
                             Array("% =", "%="), _
                             Array("& =", "&="))
End Function

Public Function getStartStatements()
    getStartStatements = Array("select", _
                               "update", _
                               "delete", _
                               "insert", _
                               "alter", _
                               "drop", _
                               "create", _
                               "add", _
                               "merge", _
                               "with", _
                               "output")
End Function

Public Function getKeyWords()
    getKeyWords = Array("absolute", "action", "ada", "add", "all", "allocate", "alter", "and", "any", "apply", "are", "as", "asc", "assertion", "at", "authorization", "avg", _
                        "begin", "between", "bit", "bit_length", "both", "by", _
                        "cascade", "cascaded", "case", "cast", "catalog", "char", "char_length", "character", "character_length", "check", "close", "coalesce", "collate", "collation", "column", "commit", "connect", "connection", "constraint", "constraints", "continue", "convert", "corresponding", "count", "create", "cross", "current", "current_date", "current_time", "current_timestamp", "current_user", "cursor", _
                        "date", "day", "deallocate", "dec", "decimal", "declare", "default", "deferrable", "deferred", "delete", "desc", "describe", "descriptor", "diagnostics", "disconnect", "distinct", "domain", "double", "drop", _
                        "else", "end", "end-exec", "escape", "except", "exception", "exec", "execute", "exists", "external", "extract", _
                        "false", "fetch", "first", "float", "for", "foreign", "fortran", "found", "from", "full", _
                        "get", "global", "go", "goto", "grant", "group", _
                        "having", "hour", _
                        "identity", "immediate", "in", "include", "index", "indicator", "initially", "inner", "input", "insensitive", "insert", "int", "integer", "intersect", "interval", "into", "is", "isolation", _
                        "join", _
                        "key", _
                        "language", "last", "leading", "left", "level", "like", "local", "lower", _
                        "match", "max", "min", "minute", "module", "month", _
                        "names", "national", "natural", "nchar", "next", "no", "none", "not", "null", "nullif", "numeric", _
                        "octet_length", "of", "on", "only", "open", "option", "or", "order", "outer", "output", "overlaps", _
                        "pad", "partial", "pascal", "position", "precision", "prepare", "preserve", "primary", "prior", "privileges", "procedure", "public", _
                        "read", "real", "references", "relative", "restrict", "revoke", "right", "rollback", "rows", _
                        "schema", "scroll", "second", "section", "select", "session", "session_user", "set", "size", "smallint", "some", "space", "sql", "sqlca", "sqlcode", "sqlerror", "sqlstate", "sqlwarning", "substring", "sum", "system_user", _
                        "table", "temporary", "then", "time", "timestamp", "timezone_hour", "timezone_minute", "to", "trailing", "transaction", "translate", "translation", "trim", "true", _
                        "union", "unique", "unknown", "update", "upper", "usage", "user", "using", _
                        "value", "values", "varchar", "varying", "view", _
                        "when", "whenever", "where", "with", "work", "write", _
                        "year", _
                        "zone")
End Function

Public Function getKeyWordsSplitByAlpha()
    getKeyWordsSplitByAlpha = Array( _
                        Array("absolute", "action", "ada", "add", "all", "allocate", "alter", "and", "any", "are", "as", "asc", "assertion", "at", "authorization", "avg"), _
                        Array("begin", "between", "bit", "bit_length", "both", "by"), _
                        Array("cascade", "cascaded", "case", "cast", "catalog", "char", "char_length", "character", "character_length", "check", "close", "coalesce", "collate", "collation", "column", "commit", "connect", "connection", "constraint", "constraints", "continue", "convert", "corresponding", "count", "create", "cross", "current", "current_date", "current_time", "current_timestamp", "current_user", "cursor"), _
                        Array("date", "day", "deallocate", "dec", "decimal", "declare", "default", "deferrable", "deferred", "delete", "desc", "describe", "descriptor", "diagnostics", "disconnect", "distinct", "domain", "double", "drop"), _
                        Array("else", "end", "end-exec", "escape", "except", "exception", "exec", "execute", "exists", "external", "extract"), _
                        Array("false", "fetch", "first", "float", "for", "foreign", "fortran", "found", "from", "full"), _
                        Array("get", "global", "go", "goto", "grant", "group"), _
                        Array("having", "hour"), _
                        Array("identity", "immediate", "in", "include", "index", "indicator", "initially", "inner", "input", "insensitive", "insert", "int", "integer", "intersect", "interval", "into", "is", "isolation"), _
                        Array("join"), _
                        Array("key"), _
                        Array("language", "last", "leading", "left", "level", "like", "local", "lower"), _
                        Array("match", "max", "min", "minute", "module", "month"), _
                        Array("names", "national", "natural", "nchar", "next", "no", "none", "not", "null", "nullif", "numeric"), _
                        Array("octet_length", "of", "on", "only", "open", "option", "or", "order", "outer", "output", "overlaps"), _
                        Array("pad", "partial", "pascal", "position", "precision", "prepare", "preserve", "primary", "prior", "privileges", "procedure", "public"), Array("q"), _
                        Array("read", "real", "references", "relative", "restrict", "revoke", "right", "rollback", "rows"), _
                        Array("schema", "scroll", "second", "section", "select", "session", "session_user", "set", "size", "smallint", "some", "space", "sql", "sqlca", "sqlcode", "sqlerror", "sqlstate", "sqlwarning", "substring", "sum", "system_user"), _
                        Array("table", "temporary", "then", "time", "timestamp", "timezone_hour", "timezone_minute", "to", "trailing", "transaction", "translate", "translation", "trim", "true"), _
                        Array("union", "unique", "unknown", "update", "upper", "usage", "user", "using"), _
                        Array("value", "values", "varchar", "varying", "view"), _
                        Array("when", "whenever", "where", "with", "work", "write"), Array("x"), _
                        Array("year"), _
                        Array("zone"))
End Function

