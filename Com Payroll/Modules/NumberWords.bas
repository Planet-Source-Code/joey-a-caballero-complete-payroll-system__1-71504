Attribute VB_Name = "NumberWords"
Public Function cNumToWord(ByVal src_num As String) As String
Dim SNUM  As Double
SNUM = Val(src_num)
If SNUM > 999999999999999# Then
    cNumToWord = "Error: To much number."
    Exit Function
End If
Dim WHOLE As String
Dim EXTRA As String
Dim WORD  As String
Dim NWHOLE As Double

If InStr(1, Str$(SNUM), ".", vbTextCompare) <> 0 Then
    WHOLE = Split(Str$(SNUM), ".")(0)
    EXTRA = Split(src_num, ".")(1)
Else
    WHOLE = SNUM
End If

If SNUM < 1 Then WORD = "Zero"

NWHOLE = Val(WHOLE)
'Check for One and Tens
If Val(Right(NWHOLE, 2)) > 0 And Val(Right(NWHOLE, 2)) < 21 Or Val(Right(NWHOLE, 2)) = 30 Or Val(Right(NWHOLE, 2)) = 40 Or Val(Right(NWHOLE, 2)) = 50 Or Val(Right(NWHOLE, 2)) = 60 Or Val(Right(NWHOLE, 2)) = 70 Or Val(Right(NWHOLE, 2)) = 80 Or Val(Right(NWHOLE, 2)) = 90 Then
    WORD = WORD & WordTens(Val(Right(NWHOLE, 2)))
ElseIf Val(Right(NWHOLE, 2)) > 20 Then
    WORD = WORD & WordTens(Left(Right(NWHOLE, 2), 1) & "0")
    WORD = WORD & WordTens(Right(NWHOLE, 1))
End If
'Check for Hundred
If NWHOLE > 99 Then
   If Left(Right(NWHOLE, 3), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 3), 1)) & " Hundred" & WORD
End If
'Check for Thousand
If NWHOLE > 999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 3))) & " Thousand" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 3), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 2, 1)) & " Thousand" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 3)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 3)) > 99 Then
            If Left(Right(NWHOLE, 6), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 6), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Million
If NWHOLE > 999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 6))) & " Million" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 6), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 2, 1)) & " Million" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 6)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 6)) > 99 Then
            If Left(Right(NWHOLE, 9), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 9), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Billion
If NWHOLE > 999999999 Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 9))) & " Billion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 9), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 2, 1)) & " Billion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 9)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 9)) > 99 Then
            If Left(Right(NWHOLE, 12), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 12), 1)) & " Hundred" & WORD
        End If
    End If
End If
'Check for Trillion
If NWHOLE > 999999999999# Then
    If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 0 And Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) < 21 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 30 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 40 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 50 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 60 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 70 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 80 Or Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) = 90 Then
        WORD = WordTens(Val(Left(NWHOLE, Len("" & NWHOLE) - 12))) & " Trillion" & WORD
    ElseIf Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 20 And Right(Left(NWHOLE, Len("" & NWHOLE) - 12), 3) <> "000" Then
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 2, 1)) & " Trillion" & WORD
        WORD = WordTens(Mid(Right(Val(Left(NWHOLE, Len("" & NWHOLE) - 12)), 2), 1, 1) & "0") & WORD
        If Val(Left(NWHOLE, Len("" & NWHOLE) - 12)) > 99 Then
            If Left(Right(NWHOLE, 15), 1) <> "0" Then WORD = WordTens(Left(Right(NWHOLE, 15), 1)) & " Hundred" & WORD
        End If
    End If
End If
If EXTRA = "" Then
    WORD = WORD & "   and   00/100"
Else
    If Val(EXTRA) < 10 Then EXTRA = "0" & EXTRA
    WORD = WORD & "   and   " & EXTRA & "/100"
End If
cNumToWord = WORD

NWHOLE = 0
WORD = ""
EXTRA = ""
WHOLE = ""
End Function
     
Private Function WordTens(ByVal SNUM As Long) As String
Select Case SNUM
    Case 1
        WordTens = " One"
    Case 2
        WordTens = " Two"
    Case 3
        WordTens = " Three"
    Case 4
        WordTens = " Four"
    Case 5
        WordTens = " Five"
    Case 6
        WordTens = " Six"
    Case 7
        WordTens = " Seven"
    Case 8
        WordTens = " Eight"
    Case 9
        WordTens = " Nine"
    Case 10
        WordTens = " Ten"
    Case 11
        WordTens = " Eleven"
    Case 12
        WordTens = " Twelve"
    Case 13
        WordTens = " Thirteen"
    Case 14
        WordTens = " Fourteen"
    Case 15
        WordTens = " Fifteen"
    Case 16
        WordTens = " Sixteen"
    Case 17
        WordTens = " Seventeen"
    Case 18
        WordTens = " Eighteen"
    Case 19
        WordTens = " Nineteen"
    Case 20
        WordTens = " Twenty"
    Case 30
        WordTens = " Thirty"
    Case 40
        WordTens = " Fourty"
    Case 50
        WordTens = " Fifty"
    Case 60
        WordTens = " Sixty"
    Case 70
        WordTens = " Seventy"
    Case 80
        WordTens = " Eighty"
    Case 90
        WordTens = " Ninty"
End Select
End Function


