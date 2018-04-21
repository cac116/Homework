Sub SortByPledgedRanges()

Dim i As Integer
Dim counter As Integer

i = 1

Do While Cells(i, "A").Value <> ""
    If Cells(i, "B").Value = "successful" And Cells(i, "A").Value < 1000 Then
        counter = counter + 1
    End If
    i = i + 1
Loop

Range("F2").Value = counter

End Sub
Sub Successful()

Dim i As Integer
Dim Y As Integer
Dim U As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim G As Integer
Dim H As Integer
Dim Z As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer


i = 1
Do While Cells(i, "A").Value <> ""
        If Cells(i, "A").Value < 1000 And Cells(i, "B").Value = "successful" Then
                Y = Y + 1
        ElseIf Cells(i, "A").Value <= 4999 And Cells(i, "B").Value = "successful" Then
                U = U + 1
        ElseIf Cells(i, "A").Value <= 9999 And Cells(i, "B").Value = "successful" Then
                C = C + 1
        ElseIf Cells(i, "A").Value <= 14999 And Cells(i, "B").Value = "successful" Then
                D = D + 1
        ElseIf Cells(i, "A").Value <= 19999 And Cells(i, "B").Value = "successful" Then
                E = E + 1
        ElseIf Cells(i, "A").Value <= 24999 And Cells(i, "B").Value = "successful" Then
                F = F + 1
        ElseIf Cells(i, "A").Value <= 29999 And Cells(i, "B").Value = "successful" Then
                G = G + 1
        ElseIf Cells(i, "A").Value <= 34999 And Cells(i, "B").Value = "successful" Then
                H = H + 1
        ElseIf Cells(i, "A").Value <= 39999 And Cells(i, "B").Value = "successful" Then
                Z = Z + 1
        ElseIf Cells(i, "A").Value <= 44999 And Cells(i, "B").Value = "successful" Then
                J = J + 1
        ElseIf Cells(i, "A").Value <= 49999 And Cells(i, "B").Value = "successful" Then
                K = K + 1
        ElseIf Cells(i, "A").Value >= 50000 And Cells(i, "B").Value = "successful" Then
                L = L + 1
    End If

    i = i + 1
Loop

Range("F2").Value = Y
Range("F3").Value = U
Range("F4").Value = C
Range("F5").Value = D
Range("F6").Value = E
Range("F7").Value = F
Range("F8").Value = G
Range("F9").Value = H
Range("F10").Value = Z
Range("F11").Value = J
Range("F12").Value = K
Range("F13").Value = L

End Sub

Sub Failed()

Dim i As Integer
Dim Y As Integer
Dim U As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim G As Integer
Dim H As Integer
Dim Z As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer


i = 1
Do While Cells(i, "A").Value <> ""
        If Cells(i, "A").Value < 1000 And Cells(i, "B").Value = "failed" Then
                Y = Y + 1
        ElseIf Cells(i, "A").Value <= 4999 And Cells(i, "B").Value = "failed" Then
                U = U + 1
        ElseIf Cells(i, "A").Value <= 9999 And Cells(i, "B").Value = "failed" Then
                C = C + 1
        ElseIf Cells(i, "A").Value <= 14999 And Cells(i, "B").Value = "failed" Then
                D = D + 1
        ElseIf Cells(i, "A").Value <= 19999 And Cells(i, "B").Value = "failed" Then
                E = E + 1
        ElseIf Cells(i, "A").Value <= 24999 And Cells(i, "B").Value = "failed" Then
                F = F + 1
        ElseIf Cells(i, "A").Value <= 29999 And Cells(i, "B").Value = "failed" Then
                G = G + 1
        ElseIf Cells(i, "A").Value <= 34999 And Cells(i, "B").Value = "failed" Then
                H = H + 1
        ElseIf Cells(i, "A").Value <= 39999 And Cells(i, "B").Value = "failed" Then
                Z = Z + 1
        ElseIf Cells(i, "A").Value <= 44999 And Cells(i, "B").Value = "failed" Then
                J = J + 1
        ElseIf Cells(i, "A").Value <= 49999 And Cells(i, "B").Value = "failed" Then
                K = K + 1
        ElseIf Cells(i, "A").Value >= 50000 And Cells(i, "B").Value = "failed" Then
                L = L + 1
    End If

    i = i + 1
Loop

Range("G2").Value = Y
Range("G3").Value = U
Range("G4").Value = C
Range("G5").Value = D
Range("G6").Value = E
Range("G7").Value = F
Range("G8").Value = G
Range("G9").Value = H
Range("G10").Value = Z
Range("G11").Value = J
Range("G12").Value = K
Range("G13").Value = L

End Sub
Sub Canceled()
Dim i As Integer
Dim Y As Integer
Dim U As Integer
Dim C As Integer
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim G As Integer
Dim H As Integer
Dim Z As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer


i = 1
Do While Cells(i, "A").Value <> ""
        If Cells(i, "A").Value < 1000 And Cells(i, "B").Value = "canceled" Then
                Y = Y + 1
        ElseIf Cells(i, "A").Value <= 4999 And Cells(i, "B").Value = "canceled" Then
                U = U + 1
        ElseIf Cells(i, "A").Value <= 9999 And Cells(i, "B").Value = "canceled" Then
                C = C + 1
        ElseIf Cells(i, "A").Value <= 14999 And Cells(i, "B").Value = "canceled" Then
                D = D + 1
        ElseIf Cells(i, "A").Value <= 19999 And Cells(i, "B").Value = "canceled" Then
                E = E + 1
        ElseIf Cells(i, "A").Value <= 24999 And Cells(i, "B").Value = "canceled" Then
                F = F + 1
        ElseIf Cells(i, "A").Value <= 29999 And Cells(i, "B").Value = "canceled" Then
                G = G + 1
        ElseIf Cells(i, "A").Value <= 34999 And Cells(i, "B").Value = "canceled" Then
                H = H + 1
        ElseIf Cells(i, "A").Value <= 39999 And Cells(i, "B").Value = "canceled" Then
                Z = Z + 1
        ElseIf Cells(i, "A").Value <= 44999 And Cells(i, "B").Value = "canceled" Then
                J = J + 1
        ElseIf Cells(i, "A").Value <= 49999 And Cells(i, "B").Value = "canceled" Then
                K = K + 1
        ElseIf Cells(i, "A").Value >= 50000 And Cells(i, "B").Value = "canceled" Then
                L = L + 1
    End If

    i = i + 1
Loop

Range("H2").Value = Y
Range("H3").Value = U
Range("H4").Value = C
Range("H5").Value = D
Range("H6").Value = E
Range("H7").Value = F
Range("H8").Value = G
Range("H9").Value = H
Range("H10").Value = Z
Range("H11").Value = J
Range("H12").Value = K
Range("H13").Value = L
End Sub
