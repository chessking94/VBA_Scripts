Attribute VB_Name = "MonteCarloSim"
Option Explicit
Sub Simulate()
Attribute Simulate.VB_ProcData.VB_Invoke_Func = " \n14"

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With

Dim wb As Workbook, results As Worksheet, iter As Worksheet
Dim i As Long, num_iter As Long, lrow As Long

Set wb = ThisWorkbook
Set results = wb.Sheets("Results")
Set iter = wb.Sheets("Iterations")

'clear old data
lrow = iter.Cells(Rows.Count, 1).End(xlUp).Row
iter.Range(iter.Cells(2, 1), iter.Cells(lrow, 13)).ClearContents

'create resultset
num_iter = results.Cells(2, 14).Value
For i = 2 To num_iter + 1
    Application.StatusBar = "Simulating iteration " & i - 1 & " of " & num_iter
    Application.Calculate
    results.Range("L2:L7").Copy
    iter.Range("B" & i).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    iter.Cells(i, 1).Value = i - 1
Next i

'create rankings
lrow = iter.Cells(Rows.Count, 1).End(xlUp).Row
iter.Cells(2, 8).Formula = "=RANK.EQ(B2,$B2:$G2)"
iter.Range("H2").Copy
iter.Range("H2:M" & lrow).PasteSpecial xlPasteFormulas

iter.Activate
iter.Cells(1, 1).Select

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .StatusBar = False
    .CutCopyMode = False
End With

End Sub
