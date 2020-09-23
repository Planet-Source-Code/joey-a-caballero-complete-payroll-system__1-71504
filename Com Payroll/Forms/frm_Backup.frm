VERSION 5.00
Begin VB.Form frm_Backup 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Backup Database"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Payroll.XP_ProgressBar Prog 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   4210816
      Scrolling       =   5
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
      Left            =   3840
      TabIndex        =   1
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Backup"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblCBK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Backup Database"
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
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1470
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H8000000D&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00EEECE8&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frm_Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents clsBKU As clsHuffman
Attribute clsBKU.VB_VarHelpID = -1

Private Sub cmdBackup_Click()
    cmdBackup.Enabled = False
    Me.cmdClose.Enabled = False
    lblCBK.Caption = "Creating Database Backup..."
    BackUpDB
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'ctrl_SkinableForm1.SkinPath = App.Path & "\skins\Wazoo"
    'Call ctrl_SkinableForm1.LoadSkin(frm_Backup)
    'Me.BackColor = &HCECECE
    'Me.Picture1.BackColor = &HCECECE
End Sub
Private Sub clsBKU_Progress(Procent As Integer)

    Prog.Value = Procent

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set clsBKU = Nothing
End Sub
Public Sub BackUpDB()
On Error GoTo err:
    Dim FSO As New FileSystemObject
    
    Dim sDBFN As String
    Dim sDBTmpFN As String
    
    If FSO.FolderExists(App.Path & "/Backup") = False Then
        FSO.CreateFolder App.Path & "/Backup"
    End If
    
    'set backup file path filename
    sDBFN = App.Path & "/Backup/" & Format$(Date, "yyyymmdd") & ".bak"
    
    'set temporary file
    sDBTmpFN = sDBFN & Now - DateValue(Now) & GetTickCount
    
    If FSO.FileExists(sDBTmpFN) = True Then
        FSO.DeleteFile sDBTmpFN
    End If
    
    'show ctl
    Prog.Visible = True
    lblCBK.Visible = True
    DoEvents
    
    'start backup
    Set frm_Backup.clsBKU = New clsHuffman
    frm_Backup.clsBKU.EncodeFile DBPathFileName, sDBTmpFN
    
    'rename file
    If FSO.FileExists(sDBFN) = True Then
        FSO.DeleteFile sDBFN
    End If
    FSO.MoveFile sDBTmpFN, sDBFN
    
    
    Set FSO = Nothing
    'Prog.Visible = False
    'lblCBK.Visible = False
    lblCBK.Caption = "Backup Complete"
    Me.cmdClose.Enabled = True
    Exit Sub
err:
    MsgBox err.Description, vbCritical
End Sub

Private Sub Prog_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
