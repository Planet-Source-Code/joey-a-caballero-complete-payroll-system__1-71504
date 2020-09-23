VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_New_emp 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Employee"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Picture Files (*.jpg,*.bmp,*.wmf,*.emf)|*.jpg;*.bmp;*.wmf;*.emf|All files (*.*)|*.*"
   End
   Begin Payroll.Photo Photo 
      Height          =   975
      Left            =   5160
      TabIndex        =   28
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      Picture         =   "frm_new_emp.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6915
      TabIndex        =   27
      Top             =   4800
      Width           =   6975
   End
   Begin VB.ComboBox cboEmpStat 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_new_emp.frx":4121
      Left            =   5160
      List            =   "frm_new_emp.frx":412B
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtBasicRate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtPosition 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtDept 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpHired 
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   59506689
      CurrentDate     =   38139
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1680
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.ComboBox cboStat 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_new_emp.frx":4145
      Left            =   1680
      List            =   "frm_new_emp.frx":4155
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3480
      Width           =   1935
   End
   Begin VB.ComboBox cboGender 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frm_new_emp.frx":417C
      Left            =   1680
      List            =   "frm_new_emp.frx":4186
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtLName 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtMI 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2400
      Width           =   735
   End
   Begin VB.TextBox txtFName 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox txtIDNo 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin MSForms.CommandButton cmdBrowse 
      Height          =   975
      Left            =   6480
      TabIndex        =   31
      ToolTipText     =   "Click here to Browse Picture"
      Top             =   3600
      Width           =   615
      VariousPropertyBits=   17
      PicturePosition =   131072
      Size            =   "1085;1720"
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Photo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   29
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Hired:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Rate:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   24
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   21
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Initial:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblID 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee ID No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   0
      Picture         =   "frm_new_emp.frx":4198
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "frm_New_emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cbo_empstat As New CAutoCompleteComboBox
Private cbo_gender As New CAutoCompleteComboBox
Private cbo_stat As New CAutoCompleteComboBox

Dim rs As New ADODB.Recordset

Private Sub cboEmpStat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub cboGender_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub


Private Sub cboStat_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub cmdAdd_Click()
    InitData
    sEnable
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    cmdBrowse.Enabled = True
End Sub

Private Sub cmdBrowse_Click()
    Photo.OpenPhotoFile
End Sub

Private Sub cmdCancel_Click()
    sDisable
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdBrowse.Enabled = False
    cmdAdd.Enabled = True
    cmdAdd.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub sEnable()
    Me.txtFName.BackColor = vbWhite
    Me.txtFName.Locked = False
    Me.txtFName.Text = ""
    
    Me.txtLName.BackColor = vbWhite
    Me.txtLName.Locked = False
    Me.txtLName.Text = ""
    
    Me.txtMI.BackColor = vbWhite
    Me.txtMI.Locked = False
    Me.txtMI.Text = ""
    
    Me.txtAddress.BackColor = vbWhite
    Me.txtAddress.Locked = False
    Me.txtAddress.Text = ""
    
    Me.txtBasicRate.BackColor = vbWhite
    Me.txtBasicRate.Locked = False
    Me.txtBasicRate.Text = "0.00"
    
    Me.txtDept.BackColor = vbWhite
    Me.txtDept.Locked = False
    Me.txtDept.Text = ""
    
    Me.txtPosition.BackColor = vbWhite
    Me.txtPosition.Locked = False
    Me.txtPosition.Text = ""
    
    Me.cboEmpStat.BackColor = vbWhite
    Me.cboEmpStat.Locked = False
    Me.cboEmpStat.Text = ""
    
    Me.cboGender.BackColor = vbWhite
    Me.cboGender.Locked = False
    Me.cboGender.Text = ""
    
    Me.cboStat.BackColor = vbWhite
    Me.cboStat.Locked = False
    Me.cboStat.Text = ""

    Me.txtFName.SetFocus
End Sub

Private Sub sDisable()
    Me.txtFName.BackColor = &H8000000B
    Me.txtFName.Locked = True
    Me.txtLName.BackColor = &H8000000B
    Me.txtLName.Locked = True
    Me.txtMI.BackColor = &H8000000B
    Me.txtMI.Locked = True
    Me.txtAddress.BackColor = &H8000000B
    Me.txtAddress.Locked = True
    Me.txtBasicRate.BackColor = &H8000000B
    Me.txtBasicRate.Locked = True
    Me.txtDept.BackColor = &H8000000B
    Me.txtDept.Locked = True
    Me.txtPosition.BackColor = &H8000000B
    Me.txtPosition.Locked = True
    Me.cboEmpStat.BackColor = &H8000000B
    Me.cboEmpStat.Locked = True
    Me.cboGender.BackColor = &H8000000B
    Me.cboGender.Locked = True
    Me.cboStat.BackColor = &H8000000B
    Me.cboStat.Locked = True
    Me.txtFName.SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblEmp", cn, adOpenKeyset, adLockPessimistic
    
    If Me.txtFName.Text = "" Then MsgBox "Please enter First Name.", vbInformation, "Information": Me.txtFName.SetFocus: Exit Sub
    If Me.txtMI.Text = "" Then MsgBox "Please enter Middle Initial.", vbInformation, "Information": Me.txtMI.SetFocus: Exit Sub
    If Me.txtLName.Text = "" Then MsgBox "Please enter Last Name.", vbInformation, "Information": Me.txtLName.SetFocus: Exit Sub
    If Me.cboGender.Text = "" Then MsgBox "Please Select Gender.", vbInformation, "Information": Me.cboGender.SetFocus: Exit Sub
    If Me.cboStat.Text = "" Then MsgBox "Please Select Status.", vbInformation, "Information": Me.cboStat.SetFocus: Exit Sub
    If Me.txtAddress.Text = "" Then MsgBox "Please enter Employee's Address.", vbInformation, "Information": Me.txtAddress.SetFocus: Exit Sub
    If Me.txtDept.Text = "" Then MsgBox "Please enter Department.", vbInformation, "Information": Me.txtDept.SetFocus: Exit Sub
    If Me.txtPosition.Text = "" Then MsgBox "Please enter Employee's Position.", vbInformation, "Information": Me.txtPosition.SetFocus: Exit Sub
    If Me.cboEmpStat.Text = "" Then MsgBox "Please Select Employee's Status.", vbInformation, "Information": Me.cboEmpStat.SetFocus: Exit Sub
    If Me.txtBasicRate.Text = "" Then MsgBox "Please enter Employee's Basic Salary Rate.", vbInformation, "Information": Me.txtBasicRate.SetFocus: Exit Sub
    
    With rs
        .AddNew
        .Fields("EM_ID").Value = Me.txtIDNo.Text
        .Fields("NAME").Value = Me.txtFName.Text & " " & Me.txtMI.Text & ". " & Me.txtLName.Text
        .Fields("Gender").Value = Me.cboGender.Text
        .Fields("Status").Value = Me.cboStat.Text
        .Fields("Address").Value = Me.txtAddress.Text
        .Fields("DEPT").Value = Me.txtDept.Text
        .Fields("POSITION").Value = Me.txtPosition.Text
        .Fields("Emp_Stat").Value = Me.cboEmpStat.Text
        .Fields("Basic_Rate").Value = Me.txtBasicRate.Text
        .Fields("Date_Employed").Value = Me.dtpHired.Value
        Me.Photo.SavePhoto .Fields("emp_Photo")
        .Update
    End With
    sDisable
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdBrowse.Enabled = False
    cmdAdd.Enabled = True
    cmdAdd.SetFocus
    MsgBox "Data entry successfully save.", vbInformation
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub


Private Sub dtpHired_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub Form_Load()
    cbo_stat.Init Me.cboStat
    cbo_empstat.Init Me.cboEmpStat
    cbo_gender.Init Me.cboGender
End Sub
   
Public Sub InitData()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open "Select * from tblEmp", cn, adOpenKeyset, adLockPessimistic
    If rs.EOF = False Then
        rs.MoveLast
        Me.txtIDNo.Text = CDbl(rs.Fields("EM_ID").Value) + 1
    Else
        Me.txtIDNo.Text = 280301
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub


Private Sub txtBasicRate_GotFocus()
    hl_Text txtBasicRate
End Sub

Private Sub txtBasicRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBasicRate_LostFocus()
If Me.txtBasicRate.Text = vbNullString Then Me.txtBasicRate.Text = FormatNumber(0, 2)
Me.txtBasicRate.Text = FormatNumber(Me.txtBasicRate.Text, 2)
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtFName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtLName_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtMI_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtPosition_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub
