We are currently working on a project to find duplicate records in a very large file. The below lines of codes are supposed to highligh the duplicate records.


Macro to highlight the duplicate counterparty names:
Private Sub CommandButton1_Click ()

Application.ScreenUpdating = False
Application.EnableEvents = False

‘Let’s find the duplicate values under column B and column A

Dim LastRow As Long
Dim i As Variant 
Dim n As Variant
LastRow = Sheets(“Report 1”).Range(“A” & Rows.Count).End(xlup).Row

For i = 2 to LastRow
For n = 1 to LastRow

If Sheets(“Report 1”).cells(i, 2).value = sheets(“Report 1”).cells(i + n, 2).value And sheets(“Report 1”).cells(i, 1).value <> sheets(“Report 1”).cells(i + n, 1).value_
	Then sheets(“Report 1”).cells(i, 2).interior.colorindex = 38: Sheets(“Report 1”).cells(i + n, 2).Interior.ColorIndex = 38: Sheets(“Report 1”).cells(i, 1).Interior.ColorIndex = 38: Sheets(“Report 1”).cells(i + n, 1).Interior.ColorIndex = 38

Next n 
Next i 

Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub
