# vba-mysql-module
VBA Mysql Module -- Simple MySQL moudule for vba with prepared statements

## Support Me
I hope my moudule will be helpful for you.
To support me please click on the link below.
[IranVBA.com](https://iranvba.com/)

### Installation
To install VBA Mysql Module follow this step:
	1 - open VBE editor on your excel or access application (ALT + F11)
	2 - click on Insert on menu bar
	3 - click on File... and select mysql.bas file

### MySQL ODBC Driver
To connect Microsoft access or Microsoft excel to MySQL server first
install MySQL ODBC Driver from [this link](https://dev.mysql.com/downloads/connector/odbc/)

### Config File Format
mysqlDriver = #your_ODBC_driver_name#
mysqlServer = #server_address#
mysqlUser = #your_database_username#
mysqlPassword = #your_database_password#
mysqlDATABASE = #your_database_name#
mysqlPort = #3306#
To find ODBC driver name follow this step:
	1- Control panel
	2- view by : large icon or small icon
	3- Administrative Tools
	4- Double click on ODBC Data Sources (32-bit) or 
	   ODBC Data Sources (64-bit) base on your OS
	5- Drivers tab
	6- You can find MySQL ODBC driver name on first column
	
### Run Parameterized Select Query
```
Sub test_1()

Dim pm(1) As String
Dim sql As String
Dim rs As ADODB.Recordset
' Run parameterized query
    sql = "SELECT * FROM `customers` WHERE `country`=? AND `state`=?"
    pm(0) = makeParameter("country", adVarChar, 50, "USA")
    pm(1) = makeParameter("state", adVarChar, 50, "CA")
    
    Set rs = mysqlOpenRs(paramQryFromString, pm, sql)
    With rs
        If .EOF = False And .BOF = False Then
            Do Until .EOF
                Debug.Print !contactFirstName & " " & !contactLastName
                .MoveNext
            Loop
        End If
    End With
    
End Sub
```

### Run Nonparameterized Select Query
```
Sub test_2()

Dim sql As String
Dim rs As ADODB.Recordset
' Run nonparameterized query
    sql = "SELECT * FROM `customers` WHERE `country`='USA' AND `state`='CA'"
    
    Set rs = mysqlOpenRs(qryfromString, , sql)
    With rs
        If .EOF = False And .BOF = False Then
            Do Until .EOF
                Debug.Print !contactFirstName & " " & !contactLastName
                .MoveNext
            Loop
        End If
    End With
    
End Sub
```

### Run Parameterized Select Query From a File
```
Sub test_3()

Dim pm(1) As String
Dim sqlPath As String
Dim rs As ADODB.Recordset
' Run parameterized query from a file

    sqlPath = currentPath & "\test.sql"
    pm(0) = makeParameter("country", adVarChar, 3, "USA")
    pm(1) = makeParameter("state", adVarChar, 2, "CA")
    
    Set rs = mysqlOpenRs(paramQryFromFile, pm, , sqlPath)
    With rs
        If .EOF = False And .BOF = False Then
            Do Until .EOF
                Debug.Print !contactFirstName & " " & !contactLastName
                .MoveNext
            Loop
        End If
    End With
    
End Sub
```

### Insert Multiple New Record With One Transaction
```
Sub test_4()
Dim sql As String
Dim newID(1) As Double

    sql = "INSERT INTO `test`(`name`, `lastname`) VALUES ('Sadegh', 'Abshenas')"
    newID(0) = mysqlExecQry(qryfromString, , sql, , True, True) ' Transaction Must Start Here
    
    
    sql = "INSERT INTO `test`(`name`, `lastname`) VALUES ('Chandler', 'Bing')"
    newID(1) = mysqlExecQry(qryfromString, , sql, , True, False) ' Note newTrans Argument Must Set to False
    
    ' New ids available before committed transaction
    Debug.Print newID(0)
    Debug.Print newID(1)
    
    ' Now you can commit transaction and make the changes on the database or rollback and undone changes.
    mysqlCommit ' or mysqlRollback

End Sub
```