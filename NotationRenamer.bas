Attribute VB_Name = "NotationRenamer"
Sub RenamePDF()

Dim MyFolder As String, ToFolder As String, MyFile As String, MyOldFile As String, MyNewFile As String, filename As String
Dim i As Integer, num As Integer, count As Integer
Dim wb1 As Workbook
Dim ws1 As Worksheet
Dim response As Variant
Dim y As String
 
MyFolder = "C:\Users\eehunt\Documents\Chess\Notation Copies\To Be Formatted"
If Right(MyFolder, 1) <> Application.PathSeparator Then MyFolder = MyFolder & Application.PathSeparator
MyFile = Dir(MyFolder & "*.pdf")
y = Year(Date)
Set wb1 = ThisWorkbook
Set ws1 = wb1.Sheets("Games")

If ws1.Range("I11") = 0 Then
    MsgBox "No files to rename!", vbExclamation
    End
End If

count = ws1.Cells(Rows.count, 1).End(xlUp).Row
count = count - 10

If count <> ws1.Range("I11").Value Then
    MsgBox "Number of files does not match number of games!", vbExclamation
    End
End If

num = ws1.Range("I11").Value + 10

If ws1.Range("O2").Value = ws1.Range("J5").Value Then
    response = MsgBox("This end date was already used, continue?", vbYesNo + vbInformation)
    If response = vbNo Then
        MsgBox "Cancelled!", vbExclamation
        End
    End If
End If

ws1.Range("N2").Value = ws1.Range("J2").Value
ws1.Range("O2").Value = ws1.Range("J5").Value
MyFile = Dir(MyFolder & "*.pdf")
ToFolder = "C:\Users\eehunt\Documents\Chess\Notation Copies\" & y & Application.PathSeparator

For i = 11 To num
    If MyFile <> "" Then
        filename = ws1.Range("G" & i).Value
        MyOldFile = MyFolder & MyFile
        MyNewFile = ToFolder & filename & ".pdf"
        Name MyOldFile As MyNewFile
        MyFile = Dir
    End If
Next i

End Sub

Sub RefreshAll()

Dim wb As Workbook
Dim ws1 As Worksheet
Dim MyFolder As String, MyFile As String
Dim ctr As Integer

MyFolder = "C:\Users\eehunt\Documents\Chess\Notation Copies\To Be Formatted"
MyFile = Dir(MyFolder & "\*.pdf")
Set wb = ThisWorkbook
Set ws1 = wb.Sheets("Games")

wb.RefreshAll

ctr = 0

Do While MyFile <> ""
    ctr = ctr + 1
    MyFile = Dir
Loop

ws1.Range("I11").Value = ctr

End Sub
