Attribute VB_Name = "DynamicSettingsUpdater"
Sub UpdateSettings()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Dim wb As Workbook, ws As Worksheet
Dim lrow As Long
Dim SettingName As String, SettingValue As String, SettingDesc As String, Setting_Upd As String

Dim sql_cmd As String
Dim Conn As ADODB.Connection
Dim ConnString As String

Set wb = ThisWorkbook
Set ws = wb.Sheets("SettingsUpdater")
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Set Conn = New Connection
ConnString = "DSN=MSSQLSERVER_ODBC;UID=eehunt;Trusted_Connection=Yes;APP=Microsoft Office;WSID=HUNT-PC1;DATABASE=ChessAnalysis;"
Conn.Open ConnString

For i = 2 To lrow
    sql_cmd = ""
    'only update if new value = 1
    If ws.Cells(i, 5).Value = "Y" Then
        sql_cmd = "UPDATE DynamicSettings SET SettingName = '" & ws.Cells(i, 2).Value & "', SettingValue = '" & ws.Cells(i, 3).Value & "', SettingDesc = '" & ws.Cells(i, 4).Value & "' WHERE SettingID = " & ws.Cells(i, 1).Value
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
