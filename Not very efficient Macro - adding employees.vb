' VB Code for finding the employees that are not in Outlook list(new employess) or that are no more in SQL Database(employess that left)
Sub Filling_KID()

Dim a As String
Dim b As String
Dim j As Integer

For i = 2 To 263

Sheets("SQL Database").Activate

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value

Sheets("Outlook List").Activate

j = 2

    Do While a <> Cells(j, 1).Value And c <> Cells(j, 3).Value
    
        j = j + 1
    
        If j > 238 Then
            Do While (IsEmpty(Cells(j, 1).Value) = False)
                j = j + 1
            Loop
            If (IsEmpty(Cells(j, 1).Value) = True) Then
                Cells(j, 1).Value = a
                Cells(j, 2).Value = b
                Cells(j, 3).Value = c
            End If
            Exit Do
        End If
        
    
    Loop

    If a = Cells(j, 1).Value And c = Cells(j, 3).Value Then
        
        If j > 238 Then
            Cells(j, 4).Value = "Active but not in list"
        Else
            If (j <= 238) Then
                Cells(j, 4).Value = "Active and present in list"
            End If
        End If
         
    End If

Next


End Sub

' =============================================
' Code for Assigning KID to the employee names

Sub Filling_KID()

Dim a As String
Dim b As String
Dim KID As String
Dim j As Integer


For i = 2 To 263

Sheets("SQL Database").Activate

KID = Cells(i, 1).Value

a = Cells(i, 2).Value

b = Cells(i, 3).Value

c = Cells(i, 4).Value


Sheets("Outlook List").Activate

j = 2

    Do While a <> Cells(j, 1).Value And c <> Cells(j, 3).Value
    
        j = j + 1
    
        If j > 304 Then
            Exit Do
        End If
        
    
    
    Loop


    If a = Cells(j, 1).Value And c = Cells(j, 3).Value Then
        
         Cells(j, 5).Value = KID
    
    End If

Next i


End Sub

'END. There are two codes in this file