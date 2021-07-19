Sub Adding_Employee()

Dim a As String
Dim b As String
Dim c As String
Dim KID As String
Dim i As Integer
Dim total_row_outlook As Integer
Dim total_row_gas As Integer
Dim total_row_storm As Integer
Dim total_row_ilp As Integer


Sheets("Outlook List").Activate

total_row_outlook = Cells(Rows.Count, 2).End(xlUp).Row

Sheets("SQL Gas").Activate

total_row_gas = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("SQL Storm").Activate

total_row_storm = Cells(Rows.Count, 1).End(xlUp).Row


Sheets("SQL ILP").Activate

total_row_ilp = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To total_row_gas

Sheets("SQL Gas").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("Outlook List").Activate

j = 2

    Do While KID <> Cells(j, 1).Value
    
        j = j + 1
    
        If j > total_row_outlook Then
			Do While (IsEmpty(Cells(j, 1).Value) = False)
					j = j + 1
			Loop
			If (IsEmpty(Cells(j, 1).Value) = True) Then
				Cells(j, 1).Value = KID
				Cells(j, 2).Value = a
				Cells(j, 3).Value = b
				Cells(j, 4).Value = c
				Cells(j, 5).Value = "Active but not in list"
			End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
		Cells(j, 5).Value = "Active and present in list"
    
	End If

Next i

For i = 2 To total_row_storm

Sheets("SQL Storm").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("Outlook List").Activate

j = 2

    Do While KID <> Cells(j, 1).Value
    
        j = j + 1
    
        If j > Cells(Rows.Count, 2).End(xlUp).Row Then
			Do While (IsEmpty(Cells(j, 1).Value) = False)
					j = j + 1
			Loop
			If (IsEmpty(Cells(j, 1).Value) = True) Then
				Cells(j, 1).Value = KID
				Cells(j, 2).Value = a
				Cells(j, 3).Value = b
				Cells(j, 4).Value = c
				Cells(j, 5).Value = "Active but not in list"
			End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
		Cells(j, 5).Value = "Active and present in list"
    
	End If

Next i


For i = 2 To total_row_ilp

Sheets("SQL ILP").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("Outlook List").Activate

j = 2

    Do While KID <> Cells(j, 1).Value
    
        j = j + 1
    
        If j > Cells(Rows.Count, 2).End(xlUp).Row Then
			Do While (IsEmpty(Cells(j, 1).Value) = False)
					j = j + 1
			Loop
			If (IsEmpty(Cells(j, 1).Value) = True) Then
				Cells(j, 1).Value = KID
				Cells(j, 2).Value = a
				Cells(j, 3).Value = b
				Cells(j, 4).Value = c
				Cells(j, 5).Value = "Active but not in list"
			End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
		Cells(j, 5).Value = "Active and present in list"
    
	End If

Next i


End Sub
