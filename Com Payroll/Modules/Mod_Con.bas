Attribute VB_Name = "Mod_Con"
Public cn As ADODB.Connection
'Public xVoice As New SpeechLib.SpVoice
Public DBPathFileName As String

Public Function cWords(ByVal strTheString As String) As String
    'Description: Capitalize the first letter of each word in a string
    Dim cr As String
    Dim t As String
    Dim i
    cr = Chr$(13) + Chr$(10)
    t = strTheString  'the string
    If t <> "" Then
        Mid$(t, 1, 1) = UCase$(Mid$(t, 1, 1))
        For i = 1 To Len(t$) - 1
            If Mid$(t, i, 2) = cr Then Mid$(t$, i + 2, 1) = UCase$(Mid$(t, i + 2, 1))
            If Mid$(t, i, 1) = " " Then Mid$(t$, i + 1, 1) = UCase$(Mid$(t, i + 1, 1))
        Next
        cWords = t
    End If
End Function

Sub Main()
    On Error GoTo eh
    DBPathFileName = App.Path & "\db\dbase.mdb"
    Set cn = New ADODB.Connection
    ConnectToServer

    frmLogin.Show

    Exit Sub
eh:
    MsgBox err.Description, vbCritical, "Error"
End Sub


Public Function ConnectToServer() As Boolean
On Error GoTo err:
    ConnectToServer = False
    cn.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & App.Path & "\db\dbase.mdb;Uid=;Pwd=;"


    ConnectToServer = True
    Exit Function
    
err:
    MsgBox err.Description & vbCrLf & Asc(err.Number), vbCritical
End Function





