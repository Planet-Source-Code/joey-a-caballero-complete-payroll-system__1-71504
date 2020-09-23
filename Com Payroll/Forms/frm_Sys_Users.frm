VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Sys_Users 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System User"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Cl&ose"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      TabIndex        =   18
      Top             =   3780
      Width           =   915
   End
   Begin MSComctlLib.ListView lvwUser 
      Height          =   3015
      Left            =   3000
      TabIndex        =   1
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "User"
         Object.Width           =   4499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UserPass"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Priv"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H00C0C0C0&
      Height          =   75
      Left            =   240
      ScaleHeight     =   15
      ScaleWidth      =   5355
      TabIndex        =   16
      Top             =   3480
      Width           =   5415
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      Height          =   3015
      Left            =   240
      ScaleHeight     =   2955
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   420
      Width           =   2715
      Begin VB.TextBox txtName 
         BackColor       =   &H00E0E0E0&
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
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtPass 
         BackColor       =   &H00E0E0E0&
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
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtConfirm 
         BackColor       =   &H00E0E0E0&
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
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.ComboBox cboPriv 
         BackColor       =   &H00E0E0E0&
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
         ItemData        =   "frm_Sys_Users.frx":0000
         Left            =   360
         List            =   "frm_Sys_Users.frx":0007
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Admin"
         Top             =   2460
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         Caption         =   "User's Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   360
         TabIndex        =   14
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   360
         TabIndex        =   13
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   360
         TabIndex        =   12
         Top             =   1620
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Privilege:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   11
         Top             =   2220
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   3780
      Width           =   840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   3780
      Width           =   915
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1150
      TabIndex        =   8
      Top             =   3780
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2790
      TabIndex        =   6
      Top             =   3780
      Width           =   915
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   335
      TabIndex        =   0
      Top             =   3780
      Width           =   795
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   3315
      Left            =   45
      Top             =   360
      Width           =   5765
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "System User Section"
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
      TabIndex        =   17
      Top             =   0
      Width           =   5955
   End
End
Attribute VB_Name = "frm_Sys_Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim boolSave As Boolean
'Dim lngID
Dim rsUsers As ADODB.Recordset

Private Sub cboPriv_GotFocus()
    hl_Text cboPriv
End Sub

Private Sub cboPriv_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            cmdSave_Click
    End Select
End Sub

Private Sub cmdAdd_Click()
    txtName.Text = ""
    txtPass.Text = ""
    txtConfirm.Text = ""
    cboPriv.Text = ""
    Me.Picture2.Enabled = True
    txtName.BackColor = vbWhite
    txtPass.BackColor = vbWhite
    txtConfirm.BackColor = vbWhite
    'cboPriv.BackColor = vbWhite
    cboPriv.Text = "Admin"
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    'lstUsers.Enabled = False
    txtName.SetFocus
    boolSave = True
End Sub

Private Sub cmdCancel_Click()
    Me.Picture2.Enabled = False
    txtName.BackColor = &HE0E0E0
    txtPass.BackColor = &HE0E0E0
    txtConfirm.BackColor = &HE0E0E0
    cboPriv.BackColor = &HE0E0E0
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    cmdCancel.Enabled = False
    If Me.lvwUser.ListItems.Count = 0 Then
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
    Me.lvwUser.SetFocus
End Sub

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
    
    cn.Execute "DELETE FROM tblUser WHERE ID = " & Me.lvwUser.SelectedItem.Text

    Call GetUsers
    txtName.Text = ""
    txtPass.Text = ""
    txtConfirm.Text = ""
    cboPriv.Text = ""
    MsgBox "Record deleted.", vbInformation

    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub cmdEdit_Click()
    txtName.BackColor = vbWhite
    txtPass.BackColor = vbWhite
    txtConfirm.BackColor = vbWhite
    cboPriv.BackColor = vbWhite
    Me.Picture2.Enabled = True
    cmdSave.Enabled = True
    cmdEdit.Enabled = False
    cmdAdd.Enabled = False
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    txtName.SetFocus
    boolSave = False
End Sub

Private Sub cmdSave_Click()
   On Error GoTo eh
    
    If txtName.Text = vbNullString Then
        MsgBox "Please enter user name.", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    
    If txtPass.Text = vbNullString Then
        MsgBox "Please enter user name.", vbExclamation
        txtPass.SetFocus
        Exit Sub
    End If
    
    If txtConfirm.Text = vbNullString Then
        MsgBox "Please enter user name.", vbExclamation
        txtConfirm.SetFocus
        Exit Sub
    End If
    
    If cboPriv.Text = vbNullString Then
        MsgBox "Please enter user name.", vbExclamation
        cboPriv.SetFocus
        Exit Sub
    End If
    
    If txtPass.Text <> txtConfirm.Text Then
        MsgBox "Password does not match!", vbExclamation
        txtConfirm.SetFocus
        Exit Sub
    End If
        
    Dim rsSave As ADODB.Recordset
    
    Set rsSave = New ADODB.Recordset
            
    With rsSave
        If boolSave = True Then         'add record to database
            If .State = adStateOpen Then .Close
            .Open "SELECT * FROM tblUser WHERE ID = 0;", cn, adOpenKeyset, adLockOptimistic
            .AddNew
            .Fields("USER_NAME") = txtName.Text
            .Fields("Password") = txtPass.Text
            '.Fields("Priv") = cboPriv.Text
            .Update
            MsgBox "New record saved.", vbInformation
        ElseIf boolSave = False Then    'update current record
            .Open "SELECT * FROM tblUser WHERE ID = " & Me.lvwUser.SelectedItem.Text, cn, adOpenKeyset, adLockOptimistic
            .Fields("USER_NAME") = txtName.Text
            .Fields("Password") = txtPass.Text
            '.Fields("Priv") = cboPriv.Text
            .Update
            MsgBox "Record updated.", vbInformation
        End If
        GetUsers
    End With
    Me.Picture2.Enabled = False
    txtName.BackColor = &HE0E0E0
    txtPass.BackColor = &HE0E0E0
    txtConfirm.BackColor = &HE0E0E0
    cboPriv.BackColor = &HE0E0E0
    cmdSave.Enabled = False
    cmdEdit.Enabled = True
    cmdAdd.Enabled = True
    cmdCancel.Enabled = False
    cmdDelete.Enabled = True
    txtName.Text = ""
    txtPass.Text = ""
    txtConfirm.Text = ""
    cboPriv.Text = ""
    Exit Sub
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Form_Load()
    Me.Picture2.BackColor = &HCECECE
    Me.cboPriv.ListIndex = 0
    Call GetUsers
    If Me.lvwUser.ListItems.Count = 0 Then
        Me.lvwUser.ListItems.Clear
        cmdDelete.Enabled = False
        cmdEdit.Enabled = False
    Else
        'Me.lvwUser.SetFocus
        Me.lvwUser.ListItems.Item(1).Selected = True
        'lvwUser_ItemClick
    End If
    
End Sub

Public Sub GetUsers()
'    On Error GoTo eh
        
    Set rsUsers = New ADODB.Recordset

    With rsUsers
        If .State = adStateOpen Then .Close
        '.CursorLocation = adUseClient
        .Open "SELECT * FROM tblUser ORDER BY USER_NAME;", cn, adOpenKeyset, adLockOptimistic
        If rsUsers.EOF = True Then
            Me.lvwUser.ListItems.Clear
            Exit Sub
        End If
        
        Me.lvwUser.ListItems.Clear
        Do While .EOF = False
            Me.lvwUser.ListItems.Add , , .Fields("id")
            Me.lvwUser.ListItems(Me.lvwUser.ListItems.Count).SubItems(1) = .Fields("USER_NAME")
            Me.lvwUser.ListItems(Me.lvwUser.ListItems.Count).SubItems(2) = .Fields("Password")
           'Me.lvwUser.ListItems(Me.lvwUser.ListItems.Count).SubItems(3) = .Fields("Priv")
            .MoveNext
        Loop
    End With


    Exit Sub
    
eh:
    MsgBox err.Description, vbCritical
End Sub


Private Sub lvwUser_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo eh
    Set rsUsers = New ADODB.Recordset

    With rsUsers
        If .State = adStateOpen Then .Close
        '.CursorLocation = adUseClient
        .Open "SELECT * FROM tblUser where ID = " & Me.lvwUser.SelectedItem.Text, cn, adOpenKeyset, adLockOptimistic
        
            Me.txtName.Text = .Fields("USER_NAME")
            Me.txtPass.Text = .Fields("Password")
            Me.txtConfirm.Text = .Fields("Password")
            Me.cboPriv.Text = "Admin"
    End With
    Exit Sub
eh:
    MsgBox err.Description, vbCritical
End Sub

Private Sub txtConfirm_GotFocus()
    hl_Text txtConfirm
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub

Private Sub txtName_GotFocus()
    hl_Text txtName
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub
Private Sub txtPass_GotFocus()
    hl_Text txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            SendKeys vbTab
    End Select
End Sub
