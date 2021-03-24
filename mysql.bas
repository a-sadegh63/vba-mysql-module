Option Compare Database
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''
' VBA MYSQL Module
' v1.0.0
''''''''''''''''''''''''''''''''''''''''''''
Public mysql_conn As ADODB.Connection
Public Enum mysqlQryType
    paramQryFromFile = 1
    paramQryFromString = 2
    qryfromFile = 3
    qryfromString = 4
End Enum

Public Sub mysqlOpenConn()
On Error GoTo errH
    
    Set mysql_conn = New ADODB.Connection
        With mysql_conn
            .ConnectionString = readConfigFile
            .CursorLocation = adUseClient
            .Mode = adModeReadWrite
            .Open
        End With
errH:
    If Err.Number <> 0 Then errHandler "mysqlOpenConn Error!", Err.Number, Err.Description
End Sub

Public Function mysql_close_conn()
    If mysql_conn.State = adStateOpen Then
        mysql_conn.Close
        Set mysql_conn = Nothing
    End If
End Function

Public Function currentPath() As String
    If Application.Name = "Microsoft Access" Then
        currentPath = Eval("application.CurrentProject.Path")
    ElseIf Application.Name = "Microsoft Excel" Then
        currentPath = Eval("Application.ActiveWorkbook.Path")
    End If
End Function

Private Function readConfigFile() As String
On Error GoTo errH
' this function read mysql server configuration from mysql.config file


Dim DataLine As String, FilePath As String, connString As String
Dim config(5) As String
Dim i As Integer, firstPos As Integer, lastPos As Integer

    FilePath = currentPath & "\mysql.config"
    
    Open FilePath For Input As #1
    i = 0
    Do Until EOF(1) Or i = 6
        Line Input #1, DataLine
        firstPos = InStr(1, DataLine, "#") + 1
        lastPos = InStrRev(DataLine, "#")
        config(i) = Mid(DataLine, firstPos, lastPos - firstPos)
        i = i + 1
    Loop
    Close #1

    readConfigFile = "DRIVER={" & config(0) & "};Server=" & _
                      config(1) & ";UID=" & config(2) & ";PWD=" & config(3) & ";DB=" & _
                      config(4) & ";PORT=" & config(5)
errH:
    If Err.Number <> 0 Then errHandler "readConfigFile Error!", Err.Number, Err.Description
End Function

Private Function sql_from_file(sqlFilePath As String) As String
On Error GoTo errH

Dim strSql As String
Dim pmNO As Integer
Dim i As Integer
Dim DataLine As String
    
    If Len(Dir(sqlFilePath)) = 0 Then
        sql_from_file = "False"
        Exit Function
    End If
    
    Open sqlFilePath For Input As #1
    Do Until EOF(1)
        Line Input #1, DataLine
        strSql = strSql & vbCrLf & DataLine
    Loop
    Close #1
    
    sql_from_file = strSql
    
errH:
    If Err.Number <> 0 Then errHandler "sql_from_file Error!", Err.Number, Err.Description
End Function

Public Function makeParameter(pmName As String, _
                              pmType As ADODB.DataTypeEnum, _
                              pmSize As LongPtr, _
                              pmValue As Variant) As String
' pmName : Parameter name who refers to a Field name on MySQL table
' pmType : Specifies the data type of a Field
' pmSize : Maximum length for the parameter value (not Field) in characters or bytes
' pmValue : A Variant that specifies the value for the Parameter object
Dim pmArray(3) As Variant
    pmArray(0) = pmName
    pmArray(1) = pmType
    pmArray(2) = pmSize
    pmArray(3) = pmValue
    makeParameter = Join(pmArray, ",")
    
End Function

Private Function customMsg(ByVal msgType As Integer) As String
    Select Case msgType
        Case 1
            customMsg = "When 'sqlType' is 'paramQryFromFile', then 'arrPm' and 'sqlPath' are required"
        Case 2
            customMsg = "When 'sqlType' is 'paramQryFromString', then 'arrPm' and 'sqlString' are required"
        Case 3
            customMsg = "When 'sqlType' is 'qryfromFile', then 'sqlPath' is required"
        Case 4
            customMsg = "When 'sqlType' is 'qryfromString', then 'sqlString' is required"
        Case 5
            customMsg = "Sql file not exist!"
        Case 6
            customMsg = "Number of parameters not match!"
        Case 7
            customMsg = "Unknown Error!"
        Case 8
            customMsg = "Please contact your system administrator"
    End Select
End Function

Public Function mysqlOpenRs(sqlType As mysqlQryType, _
                            Optional arrPm As Variant, _
                            Optional sqlString As String, _
                            Optional sqlPath As String) As ADODB.Recordset
On Error GoTo errH
' You can use this function to run your select query and store result on
' a recordset.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sqlType :
'   Specifies how function reads a query. This argument has these values:
'       paramQryFromFile :
'           "arrPm" and "sqlPath" are required.
'       paramQryFromString :
'           "arrPm" and "sqlString" are required.
'       qryfromFile :
'           "sqlPath" is required.
'       qryfromString :
'           "sqlString" is required.
' arrPm :
'   A comma separated string who holds parameters and returned from
'   "makeParameter" function,Multiple parameters passed to function
'   as a one dimensional array.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim cm As ADODB.Command
Dim i As Integer, pmCount As Integer, pmCountInSql As Integer
Dim splitPm() As String, strSql As String
Dim parameterizedQry As Boolean, readFromFile As Boolean

    If mysql_conn Is Nothing Then: Call mysqlOpenConn
    
    Select Case sqlType
        Case paramQryFromFile
            If IsNull(arrPm) Or IsNull(sqlPath) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = True
            parameterizedQry = True
        Case paramQryFromString
            If IsNull(arrPm) Or IsNull(2) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = False
            parameterizedQry = True
        Case qryfromFile
            If IsNull(sqlPath) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = True
            parameterizedQry = False
         Case qryfromString
            If IsNull(sqlString) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = False
            parameterizedQry = False
    End Select
    
    If readFromFile = True Then
        strSql = sql_from_file(sqlPath)
        If strSql = "False" Then MsgBox customMsg(5): Exit Function
    Else
        strSql = sqlString
    End If
    
    If parameterizedQry = True Then
        If IsArray(arrPm) Then
            pmCount = UBound(arrPm, 1) - LBound(arrPm, 1)
        Else
            pmCount = 0
        End If
        pmCountInSql = Len(strSql) - Len(Replace(strSql, "?", ""))
        If pmCount + 1 <> pmCountInSql Then MsgBox customMsg(6): Exit Function
    End If
    
    Set cm = New ADODB.Command
    With cm
        .ActiveConnection = mysql_conn
        .CommandText = strSql
        .CommandType = adCmdText
        If parameterizedQry = True Then
            For i = 0 To pmCount
                splitPm = Split(arrPm(i), ",")
                .Parameters.Append .CreateParameter(splitPm(0), _
                                                    splitPm(1), _
                                                    adParamInput, _
                                                    splitPm(2), _
                                                    splitPm(3) _
                                                    )
            Next
        End If
        Set mysqlOpenRs = .Execute
    End With
    
errH:
    If Err.Number <> 0 Then errHandler "mysqlOpenRs Error!", Err.Number, Err.Description
End Function

Public Function mysqlExecQry(sqlType As mysqlQryType, _
                             Optional arrPm As Variant, _
                             Optional sqlString As String, _
                             Optional sqlPath As String, _
                             Optional getLastId As Boolean = False, _
                             Optional newTrans As Boolean = True) As Double
On Error GoTo errH

' You can use this function to run your select query and store result
' on a recordset.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' sqlType :
'   Specifies how function reads a query. This argument has these values:
'       paramQryFromFile :
'           "arrPm" and "sqlPath" are required.
'       paramQryFromString :
'           "arrPm" and "sqlString" are required.
'       qryfromFile :
'           "sqlPath" is required.
'       qryfromString :
'           "sqlString" is required.
' arrPm :
'   A comma separated string who holds parameters and returned from
'   "makeParameter" function,Multiple parameters passed to function
'   as a one dimensional array.
' getLastId :
'   When you run new record query you can set it to true to get
'   new record id from Autoincrement field
' newTrans :
'   If setted to true, begins a new transaction. Beginning a new
'   transaction clears all old uncommitted transaction.
'   Note all the changes succeed when the transaction is committed so
'   no change was made until you commit changes with "mysqlCommit" sub
'   Note you can't strat new transaction when there are uncommitted changes
'   Note transactions only affected on InnoDB tables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim cm As ADODB.Command, rs As ADODB.Recordset
Dim i As Integer, pmCount As Integer, pmCountInSql As Integer
Dim splitPm() As String, strSql As String
Dim parameterizedQry As Boolean, readFromFile As Boolean

    If mysql_conn Is Nothing Then: Call mysqlOpenConn
        
    Select Case sqlType
        Case paramQryFromFile
            If IsNull(arrPm) Or IsNull(sqlPath) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = True
            parameterizedQry = True
        Case paramQryFromString
            If IsNull(arrPm) Or IsNull(2) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = False
            parameterizedQry = True
        Case qryfromFile
            If IsNull(sqlPath) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = True
            parameterizedQry = False
         Case qryfromString
            If IsNull(sqlString) Then
                MsgBox customMsg(sqlType)
                Exit Function
            End If
            readFromFile = False
            parameterizedQry = False
    End Select
    
    If readFromFile = True Then
        strSql = sql_from_file(sqlPath)
        If strSql = "False" Then MsgBox customMsg(5): Exit Function
    Else
        strSql = sqlString
    End If
    
    If parameterizedQry = True Then
        If IsArray(arrPm) Then
            pmCount = UBound(arrPm, 1) - LBound(arrPm, 1)
        Else
            pmCount = 0
        End If
        pmCountInSql = Len(strSql) - Len(Replace(strSql, "?", ""))
        If pmCount + 1 <> pmCountInSql Then MsgBox customMsg(6): Exit Function
    End If
    
    Set cm = New ADODB.Command
    With cm
        .ActiveConnection = mysql_conn
        .CommandText = strSql
        .CommandType = adCmdText
        If parameterizedQry = True Then
            For i = 0 To pmCount
                splitPm = Split(arrPm(i), ",")
                .Parameters.Append .CreateParameter(splitPm(0), _
                                                    splitPm(1), _
                                                    adParamInput, _
                                                    splitPm(2), _
                                                    splitPm(3) _
                                                    )
            Next
        End If
        If newTrans = True Then .ActiveConnection.BeginTrans
        .Execute
    End With
    
    mysqlExecQry = 1
    
    If getLastId = True Then
        cm.ActiveConnection = mysql_conn
        cm.CommandText = "SELECT LAST_INSERT_ID() AS ID;"
        Set rs = cm.Execute
        mysqlExecQry = CDbl(rs!Id)
    End If
    
errH:
    If Err.Number <> 0 Then
        errHandler "mysqlExecQry Error!", Err.Number, Err.Description
        mysqlExecQry = 0
        mysql_conn.RollbackTrans
    End If
End Function

Public Sub mysqlCommit()
' When a transaction makes multiple changes to the database, either all the changes
' succeed when the transaction is committed by this sub or undone by mysqlRollback sub
On Error GoTo errH

    If mysql_conn Is Nothing Or mysql_conn.State = 0 Then
        errHandler customMsg(7), Err.Number, "Connection to MYSQL server not established"
        Exit Sub
    End If
    
    mysql_conn.CommitTrans
    
errH:
    If Err.Number <> 0 Then errHandler "mysqlOpenRs Error!", Err.Number, Err.Description
End Sub

Public Sub mysqlRollback()
' When a transaction makes multiple changes to the database, either all the changes
' succeed when the transaction is committed by mysqlCommit sub or undone by this sub

    If mysql_conn Is Nothing Or mysql_conn.State = 0 Then
        errHandler customMsg(7), Err.Number, "Connection to MYSQL server not established"
        Exit Sub
    End If
    
    mysql_conn.RollbackTrans

End Sub

Private Function errHandler(errLable As String, errNum As LongPtr, errDesc As String)
Dim errMsg As String
Dim strPath As String

    errMsg = customMsg(8) & vbNewLine & "Error Description: " & _
              Right(errDesc, Len(errDesc) - InStrRev(errDesc, "]"))
    MsgBox errLable & vbNewLine & errMsg & vbNewLine & "Error Number = " & errNum
    
End Function
