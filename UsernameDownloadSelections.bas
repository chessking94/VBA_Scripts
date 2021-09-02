Attribute VB_Name = "UsernameDownloadSelections"
Sub AddNew()
Attribute AddNew.VB_ProcData.VB_Invoke_Func = " \n14"

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Dim wb As Workbook, ws As Worksheet
Dim lrow As Long
Dim LName As String, FName As String, UName As String, Src As String, EHFlg As String, DLFlg As String, DNE As String

Dim sql_insert As String, sql_cmd As String
Dim Conn As ADODB.Connection
Dim ConnString As String
Dim rs As ADODB.Recordset

Set wb = ThisWorkbook
Set ws = wb.Sheets("AddNew")
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'kill if no data present
If ws.Cells(lrow, 1).Value = "" Then
    MsgBox "No data to add!", vbCritical
    End
End If

'do validation
For i = 2 To lrow
    LName = ws.Cells(i, 1).Value
    FName = ws.Cells(i, 2).Value
    UName = ws.Cells(i, 3).Value
    Src = ws.Cells(i, 4).Value
    
    If LName = "" Or FName = "" Or UName = "" Or Src = "" Then
        MsgBox "Missing value! Row = " & i, vbCritical
        End
    End If
Next i

'create database connection
Set Conn = New Connection
ConnString = "DSN=MSSQLSERVER_ODBC;UID=eehunt;Trusted_Connection=Yes;APP=Microsoft Office;WSID=HUNT-PC1;DATABASE=ChessAnalysis;"
Conn.Open ConnString

'confirm user has not already been entered
Set rs = New Recordset
sql_chk = "SELECT PlayerID FROM UsernameXRef WHERE Username = '" & UName & "' AND Source = '" & Src & "'"
rs.Open sql_chk, ConnString
If Not (rs.BOF And rs.EOF) Then 'There are no records
    rs.Close
    Conn.Close
    Set rs = Nothing
    Set Conn = Nothing
    MsgBox "Source username already exists!", vbCritical
    End
End If

'all is good, proceed with inserting new data
sql_insert = "INSERT INTO UsernameXRef (LastName, FirstName, Username, Source, EEHFlag, DownloadFlag, UserStatus) VALUES "
sql_cmd = ""
EHFlg = "0"
DLFlg = "0"
UsrStat = "Open"
For i = 2 To lrow
    'set and format values, technically vulnerable to injection but why would I ever do that to myself?
    LName = "'" & Replace(LName, "'", "''") & "'"
    FName = "'" & Replace(FName, "'", "''") & "'"
    UName = "'" & Replace(UName, "'", "''") & "'"
    Src = "'" & Replace(Src, "'", "''") & "'"
    
    sql_cmd = sql_insert & "(" & LName & ", " & FName & ", " & UName & ", " & Src & ", " & EHFlg & ", " & DLFlg & ", '" & UsrStat & "')"
    'Debug.Print sql_cmd
    Conn.Execute sql_cmd
Next i

Conn.Close
Set Conn = Nothing

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

End Sub

Sub UpdateDownloadFlag()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Dim wb As Workbook, ws As Worksheet
Dim lrow As Long
Dim LName As String, FName As String, UName As String, Src As String, EHFlg As String, DLFlg As String

Dim sql_cmd As String
Dim Conn As ADODB.Connection
Dim ConnString As String

Set wb = ThisWorkbook
Set ws = wb.Sheets("DownloadFlag")
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Set Conn = New Connection
ConnString = "DSN=MSSQLSERVER_ODBC;UID=eehunt;Trusted_Connection=Yes;APP=Microsoft Office;WSID=HUNT-PC1;DATABASE=ChessAnalysis;"
Conn.Open ConnString

'reset DownloadFlag values to 0's to prevent any accidental requests
sql_cmd = "UPDATE UsernameXRef SET DownloadFlag = 0"
'Debug.Print sql_cmd
Conn.Execute sql_cmd

For i = 2 To lrow
    sql_cmd = ""
    'only update if new value = 1
    If ws.Cells(i, 8).Value = 1 Then
        sql_cmd = "UPDATE UsernameXRef SET DownloadFlag = 1 WHERE PlayerID = " & ws.Cells(i, 1).Value
        'Debug.Print sql_cmd
        Conn.Execute sql_cmd
    End If
Next i

Conn.Close
Set Conn = Nothing

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

End Sub
