Sub Bit_Coding()

Dim i As Integer
Dim j As Integer
Dim total_row As Integer
Dim Role_Sum As Integer
Dim KID As String
Dim Role As String

Sheets("Final").Activate

total_row = Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To total_row
        Sheets("Final").Activate
        If Cells(i, 5).Value = "X" Then
            KID = Cells(i, 1).Value
            Sheets("GAS").Activate
            j = 2
            Do While KID <> Cells(j, 1).Value
                j = j + 1
            Loop
            Role_Sum = 0
            Do While KID = Cells(j, 1).Value
                Role = Cells(j, 5).Value
                Select Case Role
                    Case "Front Office"
                        Role_Sum = Role_Sum + 1
                    Case "Sales Trading"
                        Role_Sum = Role_Sum + 2
                    Case "Tec Admin"
                        Role_Sum = Role_Sum + 4
                    Case "Readonly"
                        Role_Sum = Role_Sum + 8
                    Case "Pricing"
                        Role_Sum = Role_Sum + 16
                    Case "Sales Trading Admin"
                        Role_Sum = Role_Sum + 32
                    Case "Clearing Office"
                        Role_Sum = Role_Sum + 64
                    Case "Technical"
                        Role_Sum = Role_Sum + 128
                    Case "Capacity"
                        Role_Sum = Role_Sum + 256
                    Case "Backoffice Mailbox"
                        Role_Sum = Role_Sum + 512
                End Select
                j = j + 1
                If j > Cells(Rows.Count, 1).End(xlUp).Row Then
                    Exit Do
                End If
            Loop
            Sheets("Final").Activate
            Cells(i, 6).Value = Role_Sum
        End If
Next i



For i = 2 To total_row
        Sheets("Final").Activate
        If Cells(i, 7).Value = "X" Then
            KID = Cells(i, 1).Value
            Sheets("POWER").Activate
            j = 2
            Do While KID <> Cells(j, 1).Value
                j = j + 1
            Loop
            Role_Sum = 0
            Do While KID = Cells(j, 1).Value
                Role = Cells(j, 5).Value
                Select Case Role
                    Case "FRONT_OFFICE"
                        Role_Sum = Role_Sum + 1
                    Case "SALES_TRADER"
                        Role_Sum = Role_Sum + 2
                    Case "READER"
                        Role_Sum = Role_Sum + 4
                    Case "MIDDLE_OFFICE"
                        Role_Sum = Role_Sum + 8
                    Case "ADMIN"
                        Role_Sum = Role_Sum + 16
                    Case "MAILBOX"
                        Role_Sum = Role_Sum + 32
                End Select
                j = j + 1
                If j > Cells(Rows.Count, 1).End(xlUp).Row Then
                    Exit Do
                End If
            Loop
            Sheets("Final").Activate
            Cells(i, 8).Value = Role_Sum
        End If
Next i



For i = 2 To total_row
        Sheets("Final").Activate
        If Cells(i, 9).Value = "X" Then
            KID = Cells(i, 1).Value
            Sheets("ILP").Activate
            j = 2
            Do While KID <> Cells(j, 1).Value
                j = j + 1
            Loop
            Role_Sum = 0
            Do While KID = Cells(j, 1).Value
                Role = Cells(j, 5).Value
                Select Case Role
                    Case "ROLE_FRONT_OFFICE"
                        Role_Sum = Role_Sum + 1
                    Case "ROLE_MIDDLE_OFFICE"
                        Role_Sum = Role_Sum + 2
                    Case "ROLE_SALES_TRADING"
                        Role_Sum = Role_Sum + 4
                    Case "ROLE_ADMIN"
                        Role_Sum = Role_Sum + 8
                    Case "ROLE_READ_ONLY"
                        Role_Sum = Role_Sum + 16
                    Case "ROLE_PRICING"
                        Role_Sum = Role_Sum + 32
                    Case "ROLE_WSXP_SUPPORT"
                        Role_Sum = Role_Sum + 64
                End Select
                j = j + 1
                If j > Cells(Rows.Count, 1).End(xlUp).Row Then
                    Exit Do
                End If
            Loop
            Sheets("Final").Activate
            Cells(i, 10).Value = Role_Sum
        End If
Next i

            


End Sub
