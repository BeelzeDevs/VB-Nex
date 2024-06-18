Attribute VB_Name = "M�dulo999"
Public Function Code128(SourceString As String)

    Dim Counter As Integer
    Dim CheckSum As Long
    Dim mini As Integer
    Dim dummy As Integer
    Dim UseTableB As Boolean
    Dim Code128_Barcode As String
    If Len(SourceString) > 0 Then
        'Check for valid characters
        For Counter = 1 To Len(SourceString)
            Select Case Asc(Mid(SourceString, Counter, 1))
                Case 32 To 126, 203
                Case Else
                    MsgBox "Invalid character in barcode string." & vbCrLf & vbCrLf & "Please only use standard ASCII characters", vbCritical
                    Code128 = ""
                    Exit Function
            End Select
        Next
        Code128_Barcode = ""
        UseTableB = True
        Counter = 1
        Do While Counter <= Len(SourceString)
            If UseTableB Then
                'Check if we can switch to Table C
                mini = IIf(Counter = 1 Or Counter + 3 = Len(SourceString), 4, 6)
                GoSub testnum
                If mini% < 0 Then 'Use Table C
                    If Counter = 1 Then
                        Code128_Barcode = Chr(205)
                    Else 'Switch to table C
                        Code128_Barcode = Code128_Barcode & Chr(199)
                    End If
                    UseTableB = False
                Else
                    If Counter = 1 Then Code128_Barcode = Chr(204) 'Starting with table B
                End If
            End If
            If Not UseTableB Then
                'We are using Table C, try to process 2 digits
                mini% = 2
                GoSub testnum
                If mini% < 0 Then 'OK for 2 digits, process it
                    dummy% = Val(Mid(SourceString, Counter, 2))
           dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 100)
           dummy% = IIf(dummy% = 32, 194, dummy%) ' Hace que el espacio pase a ser un �
                    Code128_Barcode = Code128_Barcode & Chr(dummy%)
                    Counter = Counter + 2
                Else 'We haven't got 2 digits, switch to Table B
                    Code128_Barcode = Code128_Barcode & Chr(200)
                    UseTableB = True
                End If
            End If
            If UseTableB Then
                'Process 1 digit with table B
                Code128_Barcode = Code128_Barcode & Mid(SourceString, Counter, 1)
                Counter = Counter + 1
            End If
        Loop
        'Calculation of the checksum
        For Counter = 1 To Len(Code128_Barcode)
            dummy% = Asc(Mid(Code128_Barcode, Counter, 1))
            dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 100)
            If Counter = 1 Then CheckSum& = dummy%
            CheckSum& = (CheckSum& + (Counter - 1) * dummy%) Mod 103
        Next
        'Calculation of the checksum ASCII code
        CheckSum& = IIf(CheckSum& < 95, CheckSum& + 32, CheckSum& + 100)
        'Add the checksum and the STOP
        Code128_Barcode = Code128_Barcode & Chr(CheckSum&) & Chr$(206)
    End If
    Code128 = Code128_Barcode
    Exit Function
testnum:
        'if the mini% characters from Counter are numeric, then mini%=0
        mini% = mini% - 1
        If Counter + mini% <= Len(SourceString) Then
            Do While mini% >= 0
                If Asc(Mid(SourceString, Counter + mini%, 1)) < 48 Or Asc(Mid(SourceString, Counter + mini%, 1)) > 57 Then Exit Do
                mini% = mini% - 1
            Loop
        End If
        Return
End Function



'dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 100)

'por

'dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 100)
'dummy% = IIf(dummy% = 32, 194, dummy%) ' Hace que el espacio pase a ser un �


