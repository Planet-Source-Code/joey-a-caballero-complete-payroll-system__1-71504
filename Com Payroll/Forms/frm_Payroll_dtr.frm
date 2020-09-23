VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_Payroll_dtr 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Time Record..."
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvlDtr 
      Height          =   4335
      Left            =   0
      TabIndex        =   14
      Top             =   2400
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
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
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp. ID"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Emp. Name"
         Object.Width           =   4833
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Log"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Time In"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Time Out"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Time In"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Time Out"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Total Time"
         Object.Width           =   2187
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   30
      ScaleHeight     =   1035
      ScaleWidth      =   8475
      TabIndex        =   4
      Top             =   1320
      Width           =   8535
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   5040
         TabIndex        =   12
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20578305
         CurrentDate     =   39448
      End
      Begin VB.ComboBox cboName 
         Enabled         =   0   'False
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
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20578305
         CurrentDate     =   39448
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6570
         TabIndex        =   10
         Top             =   525
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   9
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   8
         Top             =   525
         Width           =   525
      End
      Begin MSForms.OptionButton OptByName 
         Height          =   375
         Left            =   720
         TabIndex        =   7
         Top             =   435
         Width           =   1095
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1931;661"
         Value           =   "0"
         Caption         =   "By Name"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton OptAll 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   435
         Width           =   615
         BackColor       =   12632256
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "1085;661"
         Value           =   "0"
         Caption         =   "All"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00800000&
         Height          =   735
         Left            =   1935
         Top             =   240
         Width           =   6495
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00800000&
         Height          =   735
         Left            =   30
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "View Option"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8520
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   3
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print &All Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrintSel 
      Caption         =   "Print &Selected Item"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   6960
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   6720
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   0
      Picture         =   "frm_Payroll_dtr.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8580
   End
End
Attribute VB_Name = "frm_Payroll_dtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim vViewAll As Boolean
Dim cbo_name As New CAutoCompleteComboBox

Private Sub cboName_Change()
    vName
End Sub

Private Sub cboName_Click()
    vName
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub vName()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        Me.lvlDtr.ListItems.Clear
       rs.Open "SELECT tblTimeLog.IN_AM, tblTimeLog.OUT_AM, tblTimeLog.IN_PM, tblTimeLog.OUT_PM, tblTimeLog.GRAND_TOTAL, tblTimeLog.Date_log, tblTimeLog.ID, tblTimeLog.EM_Name FROM tblTimeLog Where EM_Name = '" & Me.cboName.Text & "'GROUP BY tblTimeLog.IN_AM, tblTimeLog.OUT_AM, tblTimeLog.IN_PM, tblTimeLog.OUT_PM, tblTimeLog.GRAND_TOTAL, tblTimeLog.Date_log, tblTimeLog.ID, tblTimeLog.EM_Name HAVING tblTimeLog.Date_log Between #" & Me.DTPicker1.Value & "# AND #" & Me.DTPicker2.Value & "#;", cn, adOpenKeyset, adLockPessimistic
        Do While rs.EOF = False
            Me.lvlDtr.ListItems.Add , , rs.Fields("ID").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(1) = rs.Fields("EM_Name").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(2) = rs.Fields("Date_log").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(3) = rs.Fields("IN_AM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(4) = rs.Fields("OUT_AM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(5) = rs.Fields("IN_PM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(6) = rs.Fields("OUT_PM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(7) = rs.Fields("GRAND_TOTAL").Value
            rs.MoveNext
        Loop
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub cmdPrintAll_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
        If vViewAll = True Then
            rs.Open "SELECT * from tblTimeLog where Date_log Between #" & Me.DTPicker1.Value & "# AND #" & Me.DTPicker2.Value & "#;", cn, adOpenKeyset, adLockPessimistic
        Else
            rs.Open "SELECT * from tblTimeLog where ID=" & Me.lvlDtr.SelectedItem.Text & "And  Date_log Between #" & Me.DTPicker1.Value & "# AND #" & Me.DTPicker2.Value & "#;", cn, adOpenKeyset, adLockPessimistic
        End If
       Set drtSel_DTR.DataSource = rs
       drtSel_DTR.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmdPrintSel_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
        rs.Open "SELECT  * from tblTimeLog where ID=" & Me.lvlDtr.SelectedItem.Text & "And DATE_LOG = #" & Me.lvlDtr.ListItems(Me.lvlDtr.SelectedItem.Index).SubItems(2) & "#", cn, adOpenKeyset, adLockPessimistic
        Set drtSel_DTR.DataSource = rs
        drtSel_DTR.Show 1
        Me.cmdPrintSel.Enabled = False
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub DTPicker1_Change()
    If vViewAll = False Then
        vName
    Else
        vAll
    End If
End Sub

Private Sub DTPicker1_Click()
    If vViewAll = False Then
        vName
    Else
        vAll
    End If
End Sub

Private Sub DTPicker2_Change()
    If vViewAll = False Then
        vName
    Else
        vAll
    End If
End Sub

Private Sub DTPicker2_Click()
    If vViewAll = False Then
        vName
    Else
        vAll
    End If
End Sub

Private Sub Form_Load()
    Me.DTPicker1.Value = FormatDateTime(Now, vbShortDate)
    Me.DTPicker2.Value = FormatDateTime(Now, vbShortDate)
    cbo_name.Init cboName
    init_name
End Sub

Private Sub lvlDtr_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Me.lvlDtr.ListItems.Count <> 0 Then Me.cmdPrintSel.Enabled = True
End Sub

Private Sub OptAll_Click()
    If Me.OptAll.Value = True Then
        Me.cboName.Enabled = False
        vViewAll = True
    End If
End Sub

Private Sub OptByName_Click()
    If Me.OptByName.Value = True Then
        Me.cboName.Enabled = True
        vViewAll = False
    End If
End Sub

Private Sub init_name()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        rs.Open "Select tblTimeLog.EM_Name from tblTimeLog Group by tblTimeLog.EM_Name", cn, adOpenKeyset, adLockPessimistic
        Do While rs.EOF = False
            Me.cboName.AddItem rs.Fields("EM_Name").Value
            rs.MoveNext
        Loop
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub vAll()
On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        Me.lvlDtr.ListItems.Clear
       rs.Open "SELECT * from tblTimeLog where Date_log Between #" & Me.DTPicker1.Value & "# AND #" & Me.DTPicker2.Value & "#;", cn, adOpenKeyset, adLockPessimistic
        Do While rs.EOF = False
            Me.lvlDtr.ListItems.Add , , rs.Fields("ID").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(1) = rs.Fields("EM_Name").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(2) = rs.Fields("Date_log").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(3) = rs.Fields("IN_AM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(4) = rs.Fields("OUT_AM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(5) = rs.Fields("IN_PM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(6) = rs.Fields("OUT_PM").Value
            Me.lvlDtr.ListItems(Me.lvlDtr.ListItems.Count).SubItems(7) = rs.Fields("GRAND_TOTAL").Value
            rs.MoveNext
        Loop
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub
