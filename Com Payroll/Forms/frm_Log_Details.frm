VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_Log_Details 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log Details"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8490
   ForeColor       =   &H00004080&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   3440
      Left            =   120
      ScaleHeight     =   3375
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin MSFlexGridLib.MSFlexGrid msGrid 
         Height          =   3135
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   5530
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColor       =   16777088
         BackColorFixed  =   -2147483628
         BackColorSel    =   255
         ForeColorSel    =   -2147483624
         BackColorBkg    =   12640511
         GridColorFixed  =   8421504
         FocusRect       =   0
         HighLight       =   2
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   7320
         TabIndex        =   7
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   6120
         TabIndex        =   6
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Out"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   4800
         TabIndex        =   5
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   3600
         TabIndex        =   4
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   0
         Width           =   1485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. ID No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   45
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_Log_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
    init_Grid
    init_data
End Sub

Private Sub msGrid_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub init_data()
    Dim lngCurrentRow As Long
    Dim intNumberOfRows As Integer
    
'On Error GoTo err:

    If rs.State = adStateOpen Then rs.Close
        rs.Open "Select * from tblTimeLog where DATE_LOG = #" & FormatDateTime(Now, vbShortDate) & "#", cn, adOpenKeyset, adLockPessimistic
        intNumberOfRows = rs.RecordCount
        lngCurrentRow = 0
        
        'If rs.EOF Then rs.MoveFirst
        Do While rs.EOF = False
              With msGrid
                    .Rows = intNumberOfRows
                    .Row = lngCurrentRow
                    
                    .Col = 0: .Text = rs.Fields("ID").Value
                    .Col = 1: .Text = rs.Fields("EM_Name").Value
                    .Col = 2: .Text = rs.Fields("IN_AM").Value
                    .Col = 3: .Text = rs.Fields("OUT_AM").Value
                    .Col = 4: .Text = rs.Fields("IN_PM").Value
                    .Col = 5: .Text = rs.Fields("OUT_PM").Value
                    
                End With
                lngCurrentRow = lngCurrentRow + 1
        rs.MoveNext
        Loop
        Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
End Sub

Private Sub init_Grid()
    With msGrid
        .Clear
        .Cols = 6
        .Rows = 1
       
        .ColWidth(0) = 1265: .ColWidth(1) = 2180: .ColWidth(2) = 1200: .ColWidth(3) = 1200: .ColWidth(4) = 1200: .ColWidth(5) = 1200
    End With
End Sub
Private Sub Picture1_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
    Case vbKeyEscape
        Unload Me
    End Select
End Sub
