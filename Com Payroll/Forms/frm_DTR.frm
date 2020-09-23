VERSION 5.00
Begin VB.Form FRM_DTR 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "DTR"
   ClientHeight    =   9885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12915
   Icon            =   "frm_DTR.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtEmployeeID 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   4575
   End
   Begin VB.Image imgPic 
      DataField       =   "emp_Photo"
      Height          =   2415
      Left            =   4200
      Picture         =   "frm_DTR.frx":64692
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label lblxName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      Top             =   7080
      Width           =   4815
   End
   Begin VB.Label lblxDept 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "DEPARTMENT:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblxPos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808080&
      Caption         =   "POSITION :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   5640
      Width           =   2775
   End
   Begin VB.Label lblPosition 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   4200
      TabIndex        =   7
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Label lblDept 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   4200
      TabIndex        =   6
      Top             =   5040
      Width           =   6495
   End
   Begin VB.Label lblName 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   4440
      Width           =   6495
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3120
      TabIndex        =   4
      Top             =   7200
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2175
      Left            =   960
      Top             =   4200
      Width           =   9975
   End
   Begin VB.Label lblEnterName 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "ENTER EMPLOYEE'S ID"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   4575
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   885
      TabIndex        =   1
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   945
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "frm_DTR.frx":652FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblLogID 
      Height          =   375
      Left            =   -60
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FRM_DTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
    sFalse
End Sub

Private Sub Form_Resize()
    Timer1.Enabled = True
    
    Me.Image1.Width = Me.Width
    Me.Image1.Height = Me.Height
    
    Me.Shape1.Left = Screen.Width / 2 - Me.Shape1.Width / 2
    Me.Shape1.Top = Screen.Width / 2.2
    
    Me.lblDate.Left = Screen.Width / 2 - Me.lblDate.Width / 2
    Me.lblDate.Top = Screen.Height / 2 - Screen.Height / 2.2
    
    Me.lblTime.Left = Screen.Width / 2 - Me.lblTime.Width / 2
    Me.lblTime.Top = Screen.Height / 2 - Screen.Height / 2.8
    
    Me.lblEnterName.Left = Screen.Width / 2 - Me.lblEnterName.Width / 2
    Me.lblEnterName.Top = Screen.Width / 2 - Screen.Height / 2.5
    
    Me.txtEmployeeID.Left = Screen.Width / 2 - Me.txtEmployeeID.Width / 2
    Me.txtEmployeeID.Top = lblEnterName.Top + lblEnterName.Height
    
    Me.imgPic.Top = Me.txtEmployeeID.Top + Me.txtEmployeeID.Top / 4.5
    Me.imgPic.Left = Screen.Width / 2 - Me.imgPic.Width / 2
    
    Me.lblDept.Left = Me.Shape1.Left + 3130
    Me.lblDept.Top = Me.Shape1.Top + 850
    
    Me.lblName.Left = Me.Shape1.Left + 3130
    Me.lblName.Top = Me.Shape1.Top + 180
    
    Me.lblPosition.Left = Me.Shape1.Left + 3130
    Me.lblPosition.Top = Me.Shape1.Top + 1500

    Me.lblxName.Left = Me.Shape1.Left + 200
    Me.lblxName.Top = Me.Shape1.Top + 180
    
    Me.lblxDept.Left = Me.Shape1.Left + 200
    Me.lblxDept.Top = Me.Shape1.Top + 850
        
    Me.lblxPos.Left = Me.Shape1.Left + 200
    Me.lblxPos.Top = Me.Shape1.Top + 1500
    
    Me.lblClose.Left = Screen.Width / 2 - Me.lblClose.Width / 2
    Me.lblClose.Top = Screen.Height - Screen.Height / 9.4
    
    Me.Shape4.Left = Screen.Width / 2 - Me.Shape4.Width / 2
    Me.Shape4.Top = Screen.Height - Screen.Height / 9
    
End Sub

Private Sub lblClose_Click()
    End
End Sub

Private Sub Timer1_Timer()
    Me.lblDate.Caption = FormatDateTime(Now, vbLongDate)
    Me.lblTime.Caption = FormatDateTime(Now, vbLongTime)
End Sub

Private Sub txtEmployeeID_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            frm_Log_Details.Show 1
    End Select
End Sub

Private Sub txtEmployeeID_KeyPress(KeyAscii As Integer)
On Error GoTo err:
    Select Case KeyAscii
        Case vbKeyReturn
            If Me.txtEmployeeID.Text = vbNullString Then Exit Sub
            If login_am = True And logout_am = True And login_pm = True Then
                sFalse
            End If
            
            If rs.State = adStateOpen Then rs.Close
                rs.Open "Select * from tblEmp where EM_ID = " & Me.txtEmployeeID.Text & ";", cn, adOpenKeyset, adLockPessimistic
                Me.lblName.Caption = " " & rs.Fields("NAME").Value
                Me.lblDept.Caption = " " & rs.Fields("DEPT").Value
                Me.lblPosition.Caption = " " & rs.Fields("POSITION").Value
                Set Me.imgPic.DataSource = rs
                If rs.RecordCount >= 1 Then
                    chkLog
                Else
                    MsgBox "Employee's ID Number does not Found." & vbCrLf & vbCrLf & "Please try again.", vbInformation, "Information"
                    On Error Resume Next
                    Me.imgPic.Picture = LoadPicture(App.Path & "\images\empty.gif")
                End If
                SendKeys "{Home}+{End}"
    End Select
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub chkLog()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblTimeLog WHERE ID = " & Me.txtEmployeeID.Text & " AND DATE_LOG = #" & FormatDateTime(Now, vbShortDate) & "#", cn, adOpenKeyset, adLockPessimistic
    If Not rs.Fields("IN_AM").Value = "" Then
        login_am = True
        
        With rs
            x1 = .Fields("IN_AM").Value
            lblLogID.Caption = !LOG_ID
        End With
        Call computetime
        
        If Not rs.Fields("OUT_AM").Value = "00:00:00" Then
            logout_am = True
            With rs
                lblLogID.Caption = !LOG_ID
            End With
            
            If Not rs.Fields("IN_PM").Value = "00:00:00" Then
                login_pm = True
                With rs
                    lblLogID.Caption = !LOG_ID
                    x1 = .Fields("IN_PM").Value
                End With
                Call computetime
                
                If Not rs.Fields("OUT_PM").Value = "00:00:00" Then
                    logout_pm = True
                    With rs
                        lblLogID.Caption = !LOG_ID
                    End With
                Else
                    logout_pm = False
                End If
            Else
                login_pm = False
            End If
        Else
            logout_am = False
        End If
    Else
        login_am = False
    End If
    Call Savelog
    Exit Sub
err:
    If err.Number = 3021 Then
            login_am = False
            Call Savelog
        Else
            Resume Next
        End If
End Sub


Private Sub Savelog()
On Error GoTo err:

    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblTimeLog", cn, adOpenKeyset, adLockPessimistic
    'MsgBox "LOGIN AM: " & login_am & " LOGOUT AM: " & logout_am & " LOGIN PM: " & login_pm & "LOGOUT PM:" & logout_pm
      If login_am = False And logout_am = False And login_pm = False And logout_pm = False Then
      
        With rs
       
            .AddNew
               .Fields("ID").Value = txtEmployeeID.Text
               .Fields("EM_Name").Value = Me.lblName.Caption
               .Fields("DATE_LOG").Value = FormatDateTime(Now, vbShortDate)
               .Fields("IN_AM").Value = lblTime.Caption
               .Fields("OUT_AM").Value = "00:00:00"
               .Fields("IN_PM").Value = "00:00:00"
               .Fields("OUT_PM").Value = "00:00:00"
               .Fields("TOTAL_AM").Value = "00:00:00"
               .Fields("TOTAL_PM").Value = "00:00:00"
               .Fields("GRAND_TOTAL").Value = "00:00:00"
            .Update
            
        End With
         
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG IN!", vbInformation
    
    ElseIf login_am = True And logout_am = False And login_pm = False And logout_pm = False Then
        
        rs.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With rs
                !OUT_AM = lblTime.Caption
                !TOTAL_AM = Result
                !GRAND_TOTAL = Result
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG OUT!", vbInformation
        
    ElseIf login_am = True And logout_am = True And login_pm = False And logout_pm = False Then
        
        rs.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With rs
                .Fields("IN_PM").Value = lblTime.Caption
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG IN!", vbInformation
    
    ElseIf login_am = True And logout_am = True And login_pm = True And logout_pm = False Then
        
        rs.Find "LOG_ID='" + lblLogID.Caption + "'", 0, 1
        With rs
            .Fields("OUT_PM").Value = lblTime.Caption
            .Fields("TOTAL_PM").Value = Result
            .Fields("GRAND_TOTAL").Value = Result
            .Update
        End With
        
        With rs
                xt1 = .Fields("TOTAL_AM").Value
                xt2 = .Fields("TOTAL_PM").Value
        End With
        
        COMPUTE_TOTAL = True
        Call computetime
        With rs
                !GRAND_TOTAL = Result
            .Update
        End With
        
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has been successfully LOG OUT!", vbInformation
         
    ElseIf login_am = True And logout_am = True And login_pm = True And logout_pm = True Then
    
        MsgBox "Employees ID: " & txtEmployeeID.Text & " has already logged completely  for the day!", vbInformation
        sFalse
    End If
    
    COMPUTE_TOTAL = False
    Me.txtEmployeeID.Text = ""
    
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub computetime()
        Result = ""
    If COMPUTE_TOTAL = True Then
        timein = TimeValue(xt1)
        timeout = TimeValue(xt2)
        TimeDifferent = (timein + timeout)
        HourDiff = Hour(TimeDifferent)
        MinuteDiff = Minute(TimeDifferent)
        SecondDiff = Second(TimeDifferent)
        Result = HourDiff & ":" & MinuteDiff & ":" & SecondDiff
    Else
        x2 = lblTime.Caption
        timein = TimeValue(x1)
        timeout = TimeValue(x2)
        TimeDifferent = (timein - timeout)
        HourDiff = Hour(TimeDifferent)
        MinuteDiff = Minute(TimeDifferent)
        SecondDiff = Second(TimeDifferent)
        Result = HourDiff & ":" & MinuteDiff & ":" & SecondDiff
    End If
End Sub

Private Sub sFalse()
    login_am = False
    login_pm = False
    logout_am = False
    logout_pm = False
    COMPUTE_TOTAL = False
End Sub
