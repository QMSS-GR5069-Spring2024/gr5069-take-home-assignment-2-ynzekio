Dim x As Variant
Dim LastRow As Long

LastRow = Sheets("Sheets").Range("c" & Rows.count).end(xlup).row
For x = 2 to LastRow

If Instr(1,(sheets("sheets1").Range("A" & x)),https://usaa.ent.box.com/file") <> 0 then 
	sheets("sheet1").cells(x,2).value = "PASS"

Elseif instr(1,(Sheets("Sheets1").range("A" & x)), https://usaa.ent.box.com/file") = 0 then 
	sheets("sheet1").cells(x,2).value ="ERROR"

