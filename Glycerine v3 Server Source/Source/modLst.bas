Attribute VB_Name = "modLst"
Public Sub TimeOut(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
    DoEvents
    Loop
End Sub
Public Function EncryptA(txt As String)
    itz = ChrW(Random(25, 50)) + ReplaceString(txt, " ", "·   ") + ChrW(Random(25, 50))
    For i = 1 To Len(itz)
    Mi = Mid$(itz, i, 1)
    aa = Asc(Mi)
    ab = (aa + i) - 8
    ap = Chr$(ab)
    ch = ch & ap
    Next i
    EncryptA = ch
End Function
Public Function DecryptA(txt)
    itz = txt
    For i = 1 To Len(itz)
    Mi = Mid$(itz, i, 1)
    aa = Asc(Mi)
    ab = (aa - i) + 8
    ap = Chr$(ab)
    ch$ = ch$ & ap
    Next i
    If ch$ = Empty Then
        DecryptA = ch$
    Else
        ch$ = Right(ch$, Len(ch$) - 1)
        ch$ = Left(ch$, Len(ch$) - 1)
        ch$ = ReplaceString(ch$, "·   ", " ")
        DecryptA = ch$
    End If
End Function
Public Function Random(intFrom As Integer, intTo As Integer)
    Randomize
    Result = Int((intTo * Rnd) + intFrom)
    Random = Result
End Function

Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Sub Loadlistbox(Directory As String, TheList)
    Dim MyString As String
    On Error Resume Next
Dim a As Variant
Dim b As Variant
a = 1
Open Directory$ For Input As a
While (EOF(a) = False)
Line Input #a, b
TheList.AddItem b
Wend
Close a
End Sub

Public Sub SaveListBox(Directory As String, TheList)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
