Private Sub Workbook_AfterSave(ByVal Success As Boolean)

Dim wb As Workbook
Dim mPath As String

Set wb = ActiveWorkbook

mPath = "C:\Users\eehunt\Repository\VBA_Scripts"
If Right(mPath, 1) <> Application.PathSeparator Then mPath = mPath & Application.PathSeparator

For Each VBComp In wb.VBProject.VBComponents
    If VBComp.Type = 1 Then
        On Error Resume Next
        Err.Clear
        VBComp.Export mPath & VBComp.Name & ".bas"
    End If
Next

End Sub
