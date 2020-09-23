VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_emp_Info 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee's Information"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   8640
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10275
      TabIndex        =   2
      Top             =   5520
      Width           =   10335
   End
   Begin MSComctlLib.ListView lvlEmpInfo 
      Height          =   5055
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8916
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Emp ID."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Emp. Name"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Gender"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   5009
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Postion"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Emp. Status"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Basic Rate"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Date Employed"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print &All "
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
      Left            =   7080
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintSel 
      Caption         =   "Print &Selected"
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
      Left            =   5520
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   3960
      TabIndex        =   7
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   2400
      TabIndex        =   6
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Employee's Information Section"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10275
   End
End
Attribute VB_Name = "frm_emp_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo eh
    
    Dim intYN
    
    intYN = MsgBox("You are about to delete a record." & vbCrLf & _
        "If you click Yes, you won't be able to undo this delete operation." & _
        vbCrLf & vbCrLf & _
        "Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm Delete")
        
    If intYN = vbNo Then Exit Sub
    
    cn.Execute "DELETE FROM tblEmp WHERE EM_ID = " & Me.lvlEmpInfo.SelectedItem.Text
    Init_Data
    MsgBox "Record deleted.", vbInformation

    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblEmp where EM_ID=" & Me.lvlEmpInfo.SelectedItem.Text & ";", cn, adOpenKeyset, adLockPessimistic
    Do While rs.EOF = False
    With frm_Edit_emp
    .txtIDNo.Text = rs.Fields("EM_ID").Value
    .txtFName = rs.Fields("NAME").Value
    .cboGender.Text = rs.Fields("Gender").Value
    .cboStat.Text = rs.Fields("Status").Value
    .txtAddress.Text = rs.Fields("Address").Value
    .txtDept.Text = rs.Fields("DEPT").Value
    .txtPosition.Text = rs.Fields("POSITION").Value
    .cboEmpStat.Text = rs.Fields("Emp_Stat").Value
    .txtBasicRate.Text = rs.Fields("Basic_Rate").Value
    .dtpHired.Value = rs.Fields("Date_Employed").Value
    .Photo.LoadPhoto rs.Fields("emp_photo")
    End With
    rs.MoveNext
    Loop
    frm_Edit_emp.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub cmdPrintAll_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblEmp", cn, adOpenKeyset, adLockPessimistic
    Set drtSel_emp_info.DataSource = rs
    drtSel_emp_info.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub cmdPrintSel_Click()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    rs.Open "Select * from tblEmp where EM_ID=" & Me.lvlEmpInfo.SelectedItem.Text & ";", cn, adOpenKeyset, adLockPessimistic
    Set drtSel_emp_info.DataSource = rs
    drtSel_emp_info.Show 1
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Init_Data
    If Me.lvlEmpInfo.ListItems.Count = 0 Then
        Me.cmdPrintAll.Enabled = False
    End If
    
End Sub

Public Sub Init_Data()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    Me.lvlEmpInfo.ListItems.Clear
    rs.Open "Select * from tblEmp", cn, adOpenKeyset, adLockPessimistic
    Do While rs.EOF = False
    Me.lvlEmpInfo.ListItems.Add , , rs.Fields("EM_ID").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(1) = rs.Fields("NAME").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(2) = rs.Fields("Gender").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(3) = rs.Fields("Status").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(4) = rs.Fields("Address").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(5) = rs.Fields("DEPT").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(6) = rs.Fields("POSITION").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(7) = rs.Fields("Emp_Stat").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(8) = rs.Fields("Basic_Rate").Value
    Me.lvlEmpInfo.ListItems(Me.lvlEmpInfo.ListItems.Count).SubItems(9) = rs.Fields("Date_Employed").Value
    rs.MoveNext
    Loop
    
    If Me.lvlEmpInfo.ListItems.Count = 0 Then
        Me.cmdDelete.Enabled = False
        Me.cmdEdit.Enabled = False
        Me.cmdPrintSel.Enabled = False
        Me.cmdPrintAll.Enabled = False
    Else
        Me.cmdPrintAll.Enabled = True
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub lvlEmpInfo_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Me.lvlEmpInfo.ListItems.Count <> 0 Then
    Me.cmdDelete.Enabled = True
    Me.cmdEdit.Enabled = True
    Me.cmdPrintSel.Enabled = True
End If
End Sub
