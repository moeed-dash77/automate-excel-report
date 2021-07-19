
'Filling KID in OUTLOOK Sheet

Sub Fillig_KID()

Dim i As Integer
Dim KID AS String
Dim total_row_outlook As Integer
Dim total_row_gas As Integer
Dim total_row_storm As Integer
Dim total_row_ilp As Integer
Dim a As String
Dim b As String
Dim c As String

Sheets("OUTLOOK").Activate

total_row_outlook = Cells(Rows.Count, 2).End(xlUp).Row

Sheets("GAS").Activate

total_row_gas = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("POWER").Activate

total_row_storm = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("ILP").Activate

total_row_ilp = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To total_row_outlook 
	Sheets("OUTLOOK").Activate
	a = Cells(i, 2).Value

	b = Cells(i, 3).Value

	c = Cells(i, 4).Value
	
	j = 2
	For j = 2 To total_row_gas
		
		Sheets("GAS").Activate
		If a = Cells(j, 2).Value And b = Cells(j, 3).Value Then
			KID = Cells(j, 1).Value
			Sheets("OUTLOOK").Activate
			Cells(i, 1).Value = KID
			Exit For
		End If
	Next j
Next i


For i = 2 To total_row_outlook 
	Sheets("OUTLOOK").Activate
	a = Cells(i, 2).Value

	b = Cells(i, 3).Value

	c = Cells(i, 4).Value
	
	j = 2
	For j = 2 To total_row_storm
		
		Sheets("POWER").Activate
		If a = Cells(j, 2).Value And b = Cells(j, 3).Value Then
			KID = Cells(j, 1).Value
			Sheets("OUTLOOK").Activate
			Cells(i, 1).Value = KID
			Exit For
		End If
	Next j
Next i


For i = 2 To total_row_outlook 
	Sheets("OUTLOOK").Activate
	a = Cells(i, 2).Value

	b = Cells(i, 3).Value

	c = Cells(i, 4).Value
	
	j = 2
	For j = 2 To total_row_ilp
		
		Sheets("ILP").Activate
		If a = Cells(j, 2).Value And b = Cells(j, 3).Value Then
			KID = Cells(j, 1).Value
			Sheets("OUTLOOK").Activate
			Cells(i, 1).Value = KID
			Exit For
		End If
	Next j
Next i

End Sub

' Adding Employees, Determing Status and making Report



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


Sheets("OUTLOOK").Activate

total_row_outlook = Cells(Rows.Count, 4).End(xlUp).Row

Sheets("GAS").Activate

total_row_gas = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("POWER").Activate

total_row_storm = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("ILP").Activate

total_row_ilp = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To total_row_gas

Sheets("GAS").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("OUTLOOK").Activate

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
                Cells(j, 6).Value = "X"
            End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
        Cells(j, 6).Value = "X"
    
    End If

Next i

For i = 2 To total_row_storm

Sheets("POWER").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("OUTLOOK").Activate

j = 2

    Do While KID <> Cells(j, 1).Value
    
        j = j + 1
    
        If j > Cells(Rows.Count, 1).End(xlUp).Row Then
            Do While (IsEmpty(Cells(j, 1).Value) = False)
                    j = j + 1
            Loop
            If (IsEmpty(Cells(j, 1).Value) = True) Then
                Cells(j, 1).Value = KID
                Cells(j, 2).Value = a
                Cells(j, 3).Value = b
                Cells(j, 4).Value = c
                Cells(j, 7).Value = "X"
            End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
        Cells(j, 7).Value = "X"
    
    End If

Next i


For i = 2 To total_row_ilp

Sheets("ILP").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("OUTLOOK").Activate

j = 2

    Do While KID <> Cells(j, 1).Value
    
        j = j + 1
    
        If j > Cells(Rows.Count, 1).End(xlUp).Row Then
            Do While (IsEmpty(Cells(j, 1).Value) = False)
                    j = j + 1
            Loop
            If (IsEmpty(Cells(j, 1).Value) = True) Then
                Cells(j, 1).Value = KID
                Cells(j, 2).Value = a
                Cells(j, 3).Value = b
                Cells(j, 4).Value = c
                Cells(j, 8).Value = "X"
            End If
            Exit Do
        End If
        
    
    
    Loop


    If KID = Cells(j, 1).Value Then
        
        Cells(j, 8).Value = "X"
    
    End If

Next i


End Sub

