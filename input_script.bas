REM  *****  BASIC  *****

Sub Main

	Dim key, value As String
	Dim sheet As Object
	Dim cell, cell_input As Object
 	Dim col_input As Integer
 	Dim i, j As Integer
	Dim leave As Boolean
	
	sheet = ThisComponent.Sheets(0)
	col_input = 7
	
	Do
		leave = False
		values = ""
		key = InputBox("Enter search key", "Enter search key")
		If key = "" Then
			Exit Do
		End If
	
		For i = 0 To 120
			For j = 0 To 2
				cell = sheet.getCellByPosition(j, i)
				If InStr(cell.String, key) <> 0 Then
					ThisComponent.CurrentController.Select(cell)
					value = InputBox("Enter grade", "Enter grade")
					
					If value <> ""  Then
						cell_input = sheet.getCellByPosition(col_input, i)
						cell_input.Value = CDbl(value)
						leave = True
						Exit For
					End If
				End If
			Next j
			If leave Then
				Exit For
			End If
		Next i
	Loop
End Sub
