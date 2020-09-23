Attribute VB_Name = "Mod_ConDB"
Public cn As New ADODB.Connection
Public login_am, logout_am, login_pm, logout_pm, COMPUTE_TOTAL As Boolean
Public timein As Date, timeout As Date, TimeDifferent As Date
Public HourDiff As Integer, MinuteDiff As Integer, SecondDiff As Integer
Public Result, x1, x2, xt1, xt2 As String

Public xVoice As New SpeechLib.SpVoice

Public Sub main()
On Error GoTo err:
    xVoice.Rate = 2
    cn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cn.Open App.Path & "\db\dbase.mdb"
    FRM_DTR.Show
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Public Sub hl_text(ByRef sText)
With sText
    .SelStart = 0
    .SelLength = Len(sText.Text)
End With
End Sub
