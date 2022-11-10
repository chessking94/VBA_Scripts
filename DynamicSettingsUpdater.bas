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
ConnString = "DSN=MSSQLSERVER_ODBC;UID=eehunt;Trusted_Connection=Yes;APP=Microsoft Office;WSID=HUNT-PC1;DATABASE=ChessWarehouse;"
Conn.Open ConnString

For i = 2 To lrow
    sql_cmd = ""
    'only update if new value = 1
    If ws.Cells(i, 5).Value = "Y" Then
        nm = ws.Cells(i, 2).Value
        If nm = "" Then
            nm = "NULL"
        Else
            nm = "'" & nm & "'"
        End If
        vl = ws.Cells(i, 3).Value
        If vl = "" Then
            vl = "NULL"
        Else
            vl = "'" & vl & "'"
        End If
        dscr = ws.Cells(i, 4).Value
        If dscr = "" Then
            dscr = "NULL"
        Else
            dscr = "'" & dscr & "'"
        End If
        sql_cmd = "UPDATE Settings SET Name = " & nm & ", Value = " & vl & ", Description = " & dscr & " WHERE ID = " & ws.Cells(i, 1).Value
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
