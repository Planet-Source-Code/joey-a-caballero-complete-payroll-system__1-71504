VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Payslip 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Payslip"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtPick 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   600
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56426496
      CurrentDate     =   38139
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10395
      TabIndex        =   6
      Top             =   4440
      Width           =   10455
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
      Left            =   8040
      TabIndex        =   5
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintSel 
      Caption         =   "Print &Selected Payslip"
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
      Left            =   5880
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print &All Payslip"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvwInfo 
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp. ID Num."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Emp. Name"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Monthly Rate"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "SSS"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "PH"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Tax"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Absences"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Others"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Advances"
         Object.Width           =   2073
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Total Additional"
         Object.Width           =   2663
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Total Deduction"
         Object.Width           =   2663
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Net Pay"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSComCtl2.DTPicker dt1 
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   6120
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   56426496
      CurrentDate     =   38139
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Payslip Section"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Payslip Information for the Month of:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   650
      Width           =   3855
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   0
      Top             =   480
      Width           =   10335
   End
End
Attribute VB_Name = "frm_Payslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintAll_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblPayroll where Month(dDate)='" & Month(Me.dtPick.Value) & "' And Year(dDate)='" & Year(Me.dtPick.Value) & "'ORDER BY tblPayroll.EM_ID;", cn, adOpenKeyset, adLockPessimistic
    Set drtSel.DataSource = rs
    drtSel.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub


Private Sub cmdPrintSel_Click()
'On Error GoTo err:
    On Error Resume Next
    dt1.Value = Me.lvwInfo.ListItems(Me.lvwInfo.SelectedItem.Index).SubItems(3)
    If rs.State = adStateOpen Then rs.Close
     
    rs.Open "SELECT tblPayroll.EM_ID, tblPayroll.EM_Name, tblPayroll.Monthly_Rate, tblPayroll.dDate, tblPayroll.xBonus, tblPayroll.xOT, tblPayroll.SSS, tblPayroll.PH, tblPayroll.InTax, tblPayroll.Others, tblPayroll.absences, tblPayroll.advances, tblPayroll.NetPay " & _
    "FROM tblPayroll where EM_ID =" & Me.lvwInfo.SelectedItem.Text & " GROUP BY tblPayroll.EM_ID, tblPayroll.EM_Name, tblPayroll.Monthly_Rate, tblPayroll.dDate, tblPayroll.xBonus, tblPayroll.xOT, tblPayroll.SSS, tblPayroll.PH, tblPayroll.InTax, tblPayroll.Others, tblPayroll.absences, tblPayroll.advances, tblPayroll.NetPay having Month(dDate)='" & Month(Me.dt1.Value) & "' And Day(dDate)='" & Day(Me.dt1.Value) & "' And Year(dDate)='" & Year(Me.dt1.Value) & "' ORDER BY tblPayroll.EM_ID;", cn, adOpenKeyset, adLockPessimistic
    
    Set drtSel.DataSource = rs
    drtSel.PrintReport 'True, rptRangeAllPages
    drtSel.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub dtPick_Change()
    Init_Data
End Sub

Private Sub dtPick_Click()
    Init_Data
End Sub

Private Sub Form_Load()
    dtPick.Value = Now
    Init_Data
    If Me.lvwInfo.ListItems.Count = 0 Then
        Me.cmdPrintAll.Enabled = False
        Me.cmdPrintSel.Enabled = False
    End If
End Sub

Private Sub Init_Data()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        Me.lvwInfo.ListItems.Clear
        rs.Open "Select * from tblPayroll where Month(dDate)='" & Month(Me.dtPick.Value) & "' And Year(dDate)='" & Year(Me.dtPick.Value) & "' ORDER BY tblPayroll.EM_ID;", cn, adOpenKeyset, adLockPessimistic
        Do While rs.EOF = False
            Me.lvwInfo.ListItems.Add , , rs.Fields("EM_ID").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(1) = rs.Fields("EM_Name").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(2) = rs.Fields("Monthly_Rate").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(3) = rs.Fields("dDate").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(4) = rs.Fields("SSS").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(5) = rs.Fields("PH").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(6) = rs.Fields("InTax").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(7) = rs.Fields("Others").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(8) = rs.Fields("absences").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(9) = rs.Fields("Advances").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(10) = rs.Fields("TAD").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(11) = rs.Fields("TD").Value
            Me.lvwInfo.ListItems(Me.lvwInfo.ListItems.Count).SubItems(12) = rs.Fields("NetPay").Value
            rs.MoveNext
        Loop
    
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub
