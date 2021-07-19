
' Macro for finding total number of columns
Sub Employee_Update()

Dim LastCol As Long

Sheets("SQL Gas").Activate

LastCol = Cells(1, Columns.Count).End(xlToLeft).Column

MsgBox LastCol

End Sub

'Macro for finding total number of rows

Sub Employee_Update()

Dim LastRow As Long

Sheets("SQL Gas").Activate

LastRow = Cells(Rows.Count, 1).End(xlUp).Row

MsgBox LastRow

End Sub

'**********Updating Employee List**********

Sub Employee_Update()

Dim a As String
Dim b As String
'Dim c As String
Dim d As String
Dim i As Integer
Dim total_row_outlook As Integer
Dim total_row_gas As Integer
Dim total_row_storm As Integer
Dim total_row_ilp As Integer
Dim j As Integer

Sheets("Outlook List").Activate

total_row_outlook = Cells(Rows.Count, 2).End(xlUp).Row

Sheets("SQL Gas").Activate

total_row_gas = Cells(Rows.Count, 1).End(xlUp).Row

Sheets("SQL Storm").Activate

total_row_storm = Cells(Rows.Count, 1).End(xlUp).Row


Sheets("SQL ILP").Activate

total_row_ilp = Cells(Rows.Count, 1).End(xlUp).Row



For i = 2 To total_row_outlook

Sheets("Outlook List").Activate

a = Cells(i, 2).Value

b = Cells(i, 3).Value

'c = Cells(i, 4).Value

Sheets("SQL Gas").Activate

j = 2

    Do While a <> Cells(j, 2).Value And b <> Cells(j, 3).Value
    
        j = j + 1
    
        If j > total_row_gas Then
            Sheets("SQL Storm").Activate
            j = 2
            Do While a <> Cells(j, 2).Value And b <> Cells(j, 3).Value
                j = j + 1
                
                If j > total_row_storm Then
                    Sheets("SQL ILP").Activate
                    j = 2
                    Do While a <> Cells(j, 2).Value And b <> Cells(j, 3).Value
                        j = j + 1
                        
                        If j > total_row_ilp Then
                            Sheets("Outlook List").Cells(i, 5).Value = "Not in database"
                            Exit Do
                        End If
                    Loop
                    Exit Do
                End If
            Loop
            Exit Do
        End If
        
    Loop
    
Next i
                            
                            

End Sub