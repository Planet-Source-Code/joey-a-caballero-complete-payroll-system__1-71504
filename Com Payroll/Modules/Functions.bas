Attribute VB_Name = "Functions"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByRef lParam As Any _
) As Long

Sub AutoTXTcomplete(LST As ListBox, TXT As TextBox)
Dim strt As Long, nIndex As Long
Dim nLen As Long, sText As String
Const LB_GETTEXTLEN As Long = &H18A
Const LB_GETTEXT As Long = &H189
Static blnBusy As Boolean

If blnBusy Then
   Exit Sub
End If
     
     bNoClick = True
     blnBusy = True
    
    'Retrieve the item's listindex
    LST.ListIndex = SendMessage(LST.hwnd, LB_FINDSTRING, -1, ByVal CStr(TXT.Text))
    
    If Not DelKey Then


    If LST.ListIndex <> -1 Then
        strt = Len(TXT.Text)
        TXT.Text = LST.List(LST.ListIndex)
        TXT.SelStart = strt
        TXT.SelLength = Len(TXT.Text) - strt
    Else
    
    End If
    End If
       DelKey = False
       blnBusy = False
       bNoClick = False
End Sub
Sub hl_Text(TXT As TextBox)
    'SendKeys "{HOME}"
    TXT.SelStart = 0
    TXT.SelLength = Len(TXT.Text)
End Sub
