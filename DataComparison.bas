Attribute VB_Name = "DataComparison"
Sub Update_ACPL_Control_Data(Rating As String)

Dim wb As Workbook, ws As Worksheet

Set wb = ThisWorkbook
Set ws = wb.Sheets("ACPL_Control")

With ws.ListObjects("ACPL_Data_Control").QueryTable
    .CommandText = "EXEC SelectControlACPLDataComplete_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

With ws.ListObjects("ACPL_Data_Control_Phase").QueryTable
    .CommandText = "EXEC SelectControlACPLDataPhase_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

End Sub

Sub Update_T10_Control_Data(Rating As String)
Attribute Update_T10_Control_Data.VB_ProcData.VB_Invoke_Func = " \n14"

Dim wb As Workbook, ws As Worksheet

Set wb = ThisWorkbook
Set ws = wb.Sheets("T10_Summary_Control")

With ws.ListObjects("T10_Data_Control").QueryTable
    .CommandText = "EXEC SelectControlTopXDataComplete_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

With ws.ListObjects("T10_Data_Control_Phase").QueryTable
    .CommandText = "EXEC SelectControlTopXDataPhase_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

End Sub

Sub Update_Scores_Control_Data(Rating As String)

Dim wb As Workbook, ws As Worksheet

Set wb = ThisWorkbook
Set ws = wb.Sheets("Scores_Control")

With ws.ListObjects("Score_Data_Control").QueryTable
    .CommandText = "EXEC SelectControlScoreDataComplete_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

With ws.ListObjects("Score_Data_Control_Phase").QueryTable
    .CommandText = "EXEC SelectControlScoreDataPhase_Rating" + Chr(10) + "@Rating = " + Rating
    .Refresh
    DoEvents
End With

End Sub

Sub Update_ACPL_Sample_Data(CompareWith As String, LastName As String, FirstName As String, MinRating As String, MaxRating As String, MinDate As String, MaxDate As String, ECO As String, Tmnt As String)

Dim wb As Workbook, ws As Worksheet
Dim cmdtxt As String

Set wb = ThisWorkbook
Set ws = wb.Sheets("ACPL_Sample")

With ws.ListObjects("ACPL_Data_Sample").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingACPLDataComplete_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHACPLDataComplete_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

With ws.ListObjects("ACPL_Data_Sample_Phase").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingACPLDataPhase_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHACPLDataPhase_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

End Sub

Sub Update_T10_Sample_Data(CompareWith As String, LastName As String, FirstName As String, MinRating As String, MaxRating As String, MinDate As String, MaxDate As String, ECO As String, Tmnt As String)

Dim wb As Workbook, ws As Worksheet
Dim cmdtxt As String

Set wb = ThisWorkbook
Set ws = wb.Sheets("T10_Summary_Sample")

With ws.ListObjects("T10_Data_Sample").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingTopXDataComplete_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHTopXDataComplete_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

With ws.ListObjects("T10_Data_Sample_Phase").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingTopXDataPhase_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHTopXDataPhase_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

End Sub

Sub Update_Scores_Sample_Data(CompareWith As String, LastName As String, FirstName As String, MinRating As String, MaxRating As String, MinDate As String, MaxDate As String, ECO As String, Tmnt As String)

Dim wb As Workbook, ws As Worksheet
Dim cmdtxt As String

Set wb = ThisWorkbook
Set ws = wb.Sheets("Scores_Sample")

With ws.ListObjects("Score_Data_Sample").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingScoreDataComplete_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHScoreDataComplete_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

With ws.ListObjects("Score_Data_Sample_Phase").QueryTable
    If CompareWith = "Test" Then
        cmdtxt = "EXEC SelectTestingScoreDataPhase_LastFirst"
        cmdtxt = cmdtxt + Chr(10) + "@LastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@FirstName = " + FirstName
        .CommandText = cmdtxt
    Else
        cmdtxt = "EXEC SelectEEHScoreDataPhase_Variables"
        cmdtxt = cmdtxt + Chr(10) + "@OppLastName = " + LastName + ","
        cmdtxt = cmdtxt + Chr(10) + "@OppFirstName = " + FirstName + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinOppRating = " + MinRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxOppRating = " + MaxRating + ","
        cmdtxt = cmdtxt + Chr(10) + "@MinDate = " + MinDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@MaxDate = " + MaxDate + ","
        cmdtxt = cmdtxt + Chr(10) + "@ECO = " + ECO + ","
        cmdtxt = cmdtxt + Chr(10) + "@Tmnt = " + Tmnt
        .CommandText = cmdtxt
    End If
    .Refresh
    DoEvents
End With

End Sub

Sub RefreshAllWithParameters()

Dim wb As Workbook, param As Worksheet
Dim CtrlRefresh As String, CompareWith As String
Dim Ctrl_Rating As String
Dim Test_LName As String, Test_FName As String
Dim EEH_LName As String, EEH_FName As String, EEH_MinRating As String, EEH_MaxRating As String, EEH_MinDate As String, EEH_MaxDate As String, EEH_ECO As String, EEH_Tmnt As String

Set wb = ThisWorkbook
Set param = wb.Sheets("INPUT")

'Determine refresh status of control data
CtrlRefresh = param.Cells(1, 2).Value
If CtrlRefresh = "Yes" Then
    Application.StatusBar = "Refreshing Control"
    Ctrl_Rating = param.Cells(4, 5).Value
    If Ctrl_Rating = "" Then Ctrl_Rating = "NULL"
    Call Update_ACPL_Control_Data(Ctrl_Rating)
    Call Update_T10_Control_Data(Ctrl_Rating)
    Call Update_Scores_Control_Data(Ctrl_Rating)
End If

'Determine which data to refresh and compare
CompareWith = param.Cells(2, 2).Value
If CompareWith = "Test" Then
    'set values
    Test_LName = param.Cells(4, 8).Value
    Test_FName = param.Cells(5, 8).Value
    
    'validation of inputs
    If Test_LName = "" And Test_FName = "" Then
        MsgBox "No test name entered!", vbCritical
        Application.StatusBar = False
        End
    End If
    
    'possible to switch to a for-each loop over an array of variable names? worth investigating
    If Test_LName = "" Then
        Test_LName = "NULL"
    Else
        Test_LName = "'" + Test_LName + "'"
    End If
    
    If Test_FName = "" Then
        Test_FName = "NULL"
    Else
        Test_FName = "'" + Test_FName + "'"
    End If
    
    'refresh data
    Application.StatusBar = "Refreshing Test Data"
    Call Update_ACPL_Sample_Data(CompareWith, Test_LName, Test_FName, "", "", "", "", "", "")
    Call Update_T10_Sample_Data(CompareWith, Test_LName, Test_FName, "", "", "", "", "", "")
    Call Update_Scores_Sample_Data(CompareWith, Test_LName, Test_FName, "", "", "", "", "", "")
ElseIf CompareWith = "EEH" Then
    'set values
    EEH_LName = param.Cells(4, 11).Value
    EEH_FName = param.Cells(5, 11).Value
    EEH_MinRating = param.Cells(6, 11).Value
    EEH_MaxRating = param.Cells(7, 11).Value
    EEH_MinDate = param.Cells(8, 11).Value
    EEH_MaxDate = param.Cells(9, 11).Value
    EEH_ECO = param.Cells(10, 11).Value
    EEH_Tmnt = param.Cells(11, 11).Value
    
    'validation of inputs
    'possible to switch to a for-each loop over an array of variable names? worth investigating
    If EEH_LName = "" Then
        EEH_LName = "NULL"
    Else
        EEH_LName = "'" + EEH_LName + "'"
    End If
    
    If EEH_FName = "" Then
        EEH_FName = "NULL"
    Else
        EEH_FName = "'" + EEH_FName + "'"
    End If
    
    If EEH_MinRating = "" Then
        EEH_MinRating = "NULL"
    End If
    
    If EEH_MaxRating = "" Then
        EEH_MaxRating = "NULL"
    End If
    
    If EEH_MinDate = "" Then
        EEH_MinDate = "NULL"
    Else
        EEH_MinDate = "'" + Format(EEH_MinDate, "yyyy-mm-dd") + "'"
    End If
    
    If EEH_MaxDate = "" Then
        EEH_MaxDate = "NULL"
    Else
        EEH_MaxDate = "'" + Format(EEH_MaxDate, "yyyy-mm-dd") + "'"
    End If
    
    If EEH_ECO = "" Then
        EEH_ECO = "NULL"
    Else
        EEH_ECO = "'" + EEH_ECO + "'"
    End If
    
    If EEH_Tmnt = "" Then
        EEH_Tmnt = "NULL"
    Else
        EEH_Tmnt = "'" + EEH_Tmnt + "'"
    End If
    
    'refresh data
    Application.StatusBar = "Refreshing EEH Data"
    Call Update_ACPL_Sample_Data(CompareWith, EEH_LName, EEH_FName, EEH_MinRating, EEH_MaxRating, EEH_MinDate, EEH_MaxDate, EEH_ECO, EEH_Tmnt)
    Call Update_T10_Sample_Data(CompareWith, EEH_LName, EEH_FName, EEH_MinRating, EEH_MaxRating, EEH_MinDate, EEH_MaxDate, EEH_ECO, EEH_Tmnt)
    Call Update_Scores_Sample_Data(CompareWith, EEH_LName, EEH_FName, EEH_MinRating, EEH_MaxRating, EEH_MinDate, EEH_MaxDate, EEH_ECO, EEH_Tmnt)
Else
    MsgBox "Unknown type to compare Control data to!", vbCritical
    Application.StatusBar = False
    End
End If



Application.StatusBar = False

End Sub
