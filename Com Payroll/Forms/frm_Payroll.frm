VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Payroll 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Payroll..."
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8145
      ScaleWidth      =   10665
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   10695
      Begin VB.TextBox txtOther 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   8400
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3840
         Width           =   1695
      End
      Begin VB.TextBox txtOverTime 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   4140
         Width           =   1635
      End
      Begin VB.TextBox txtNumDays 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   4
         Text            =   "0"
         Top             =   4440
         Width           =   1635
      End
      Begin VB.TextBox txtBonus 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   3825
         Width           =   1635
      End
      Begin VB.TextBox txtAdvances 
         Alignment       =   1  'Right Justify
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
         Left            =   1800
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   3540
         Width           =   1635
      End
      Begin VB.ComboBox cboEmpID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_Payroll.frx":0000
         Left            =   1800
         List            =   "frm_Payroll.frx":0002
         TabIndex        =   0
         ToolTipText     =   "Select ID Number then press Enter"
         Top             =   1560
         Width           =   1755
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  'Center
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtph 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7080
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txttax 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7080
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox txtsss 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7080
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   3540
         Width           =   1215
      End
      Begin VB.TextBox txtAbsences 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   4440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1665
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   56688641
         CurrentDate     =   38686
      End
      Begin MSComctlLib.ListView lvwInfo 
         Height          =   2235
         Left            =   285
         TabIndex        =   40
         Top             =   5580
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3942
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
            Text            =   "Pag-ibig"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pag-ibig:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   8400
         TabIndex        =   51
         Top             =   3540
         Width           =   765
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Overtime Pay:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   50
         Top             =   4140
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8040
         TabIndex        =   49
         Top             =   2400
         Width           =   2115
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   48
         Top             =   2400
         Width           =   4275
      End
      Begin VB.Label lblPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8040
         TabIndex        =   47
         Top             =   2040
         Width           =   2115
      End
      Begin VB.Label lblName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1800
         TabIndex        =   46
         Top             =   2040
         Width           =   4275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employment Status:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6285
         TabIndex        =   45
         Top             =   2400
         Width           =   1725
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Dept:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   44
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Position:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6280
         TabIndex        =   43
         Top             =   2040
         Width           =   1620
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   42
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Payslip Information"
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
         Left            =   300
         TabIndex        =   41
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label lblDeduct 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   285
         Left            =   8400
         TabIndex        =   39
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deductions:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8580
         TabIndex        =   38
         Top             =   4140
         Width           =   1515
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Of Days Absent:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   37
         Top             =   4440
         Width           =   1155
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   36
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Advances:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   3540
         Width           =   915
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   1620
         Width           =   1275
      End
      Begin VB.Label lblRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   5040
         TabIndex        =   33
         Top             =   1560
         Width           =   1635
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emp. ID Number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   32
         Top             =   1620
         Width           =   1485
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5100
         TabIndex        =   31
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   30
         Top             =   780
         Width           =   555
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "For The Month:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   225
         TabIndex        =   29
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Generate Payroll"
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
         Left            =   180
         TabIndex        =   28
         Top             =   60
         Width           =   10275
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000D&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00EEECE8&
         FillStyle       =   0  'Solid
         Height          =   2775
         Left            =   180
         Top             =   5220
         Width           =   10275
      End
      Begin VB.Label lblNetPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   8400
         TabIndex        =   27
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Pay"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   4800
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Payroll Information"
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
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   420
         Width           =   10275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Deductions"
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
         Index           =   1
         Left            =   5520
         TabIndex        =   24
         Top             =   3180
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "PH Contribution:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   5520
         TabIndex        =   23
         Top             =   4140
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "W/Holding TAX:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   5520
         TabIndex        =   22
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "SSS Contribution:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   5520
         TabIndex        =   21
         Top             =   3540
         Width           =   1575
      End
      Begin VB.Label lblPerDay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   8040
         TabIndex        =   20
         Top             =   1560
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Absences:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   5520
         TabIndex        =   19
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Rate:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6960
         TabIndex        =   18
         Top             =   1620
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Employee Information"
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
         Index           =   2
         Left            =   165
         TabIndex        =   17
         Top             =   1200
         Width           =   10275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Data Entry"
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
         Index           =   3
         Left            =   225
         TabIndex        =   16
         Top             =   3180
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00EEECE8&
         FillStyle       =   0  'Solid
         Height          =   2595
         Left            =   180
         Top             =   420
         Width           =   10275
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000D&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00EEECE8&
         FillStyle       =   0  'Solid
         Height          =   2115
         Left            =   180
         Top             =   3060
         Width           =   10275
      End
   End
   Begin VB.CommandButton cmdCosePic 
      Caption         =   "C&lose"
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
      Left            =   9240
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8340
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintPayslip 
      Caption         =   "&Print Payslip"
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
      Left            =   8040
      TabIndex        =   53
      Top             =   8340
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      TabIndex        =   9
      Top             =   8340
      Width           =   1215
   End
   Begin VB.Label lbltemp 
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   9600
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Payroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cbo_id As New CAutoCompleteComboBox
Dim rs As New ADODB.Recordset

Private Sub cboEmpID_Change()
On Error GoTo err:
    If Me.cboEmpID.Text = vbNullString Then Exit Sub
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open "Select * from tblEmp where EM_ID = " & Me.cboEmpID.Text & ";", cn, adOpenKeyset, adLockPessimistic
    Do While rs.EOF = False
    Me.lblName.Caption = " " & rs.Fields("NAME").Value
    Me.lblDept.Caption = " " & rs.Fields("DEPT").Value
    Me.lblPosition.Caption = " " & rs.Fields("POSITION").Value
    Me.lblRate.Caption = FormatNumber(CCur(rs.Fields("Basic_Rate").Value), 2)
    Me.lblStatus.Caption = " " & rs.Fields("Emp_Stat").Value
    Me.lblPerDay.Caption = FormatNumber(CCur(Me.lblRate.Caption / 30), 2)
    Me.lblNetPay.Caption = FormatNumber(CCur(Me.lblRate.Caption), 2)
    rs.MoveNext
    Loop
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub cboEmpID_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        cboEmpID_Change
        If Me.lblName.Caption = " " Or Me.lblName.Caption = "" Then MsgBox "Invalid Employee ID Number. Please try again", vbInformation, "Information": Exit Sub
        SendKeys vbTab
    End Select
End Sub

Private Sub cmdCosePic_Click()
    Unload Me
End Sub

Private Sub cmdPrintPayslip_Click()
    frm_Payslip.Show 1
End Sub

Private Sub cmdSave_Click()
Dim strMess
On Error GoTo err:
    strMess = "Saving this data will not allow you to re-edit this transaction." & vbCrLf & vbCrLf & "Are you sure you want to save this now?"
    If MsgBox(strMess, vbQuestion + vbYesNo) = vbYes Then
        If rs.State = adStateOpen Then rs.Close
        
        rs.Open "Select * from tblPayroll", cn, adOpenKeyset, adLockPessimistic
        With rs
            .AddNew
            .Fields("EM_ID").Value = Me.cboEmpID.Text
            .Fields("EM_Name").Value = Me.lblName.Caption
            .Fields("Monthly_Rate").Value = Me.lblRate.Caption
            .Fields("dDate").Value = Me.DTPicker1.Value
            .Fields("SSS").Value = Me.txtsss.Text
            .Fields("xBonus").Value = Me.txtBonus.Text
            .Fields("xOT").Value = Me.txtOverTime.Text
            .Fields("PH").Value = Me.txtph.Text
            .Fields("InTax").Value = Me.txttax.Text
            .Fields("Others").Value = Me.txtOther.Text
            .Fields("absences").Value = Me.txtAbsences.Text
            .Fields("Advances").Value = Me.txtAdvances.Text
            .Fields("TAD").Value = Me.lbltemp.Caption
            .Fields("TD").Value = Me.lblDeduct.Caption
            .Fields("NetPay").Value = Me.lblNetPay.Caption
            .Update
            MsgBox "Data successfully save.", vbInformation, "Information"
        End With
        Init_Display_Data
        sText
        Me.cboEmpID.SetFocus
    End If
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub DTPicker1_Change()
    DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
    Me.txtMonth.Text = MonthName(Month(Me.DTPicker1.Value))
    Me.txtYear.Text = Year(Me.DTPicker1.Value)
    Init_Display_Data
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Now
    DTPicker1_Click
    cbo_id.Init cboEmpID
    
    init_ID_Data
    Init_Display_Data
End Sub

Private Sub init_ID_Data()
On Error GoTo err:
    If rs.State = adStateOpen Then rs.Close
    
    rs.Open "Select * from tblEmp", cn, adOpenKeyset, adLockPessimistic
    Do While rs.EOF = False
        cboEmpID.AddItem rs.Fields("EM_ID").Value
        rs.MoveNext
    Loop
    Exit Sub
err:
    MsgBox err.Description, vbCritical, "Error"
    Set rs = Nothing
End Sub

Private Sub lblDeduct_Change()
    If Me.lblDeduct.Caption = vbNullString Then
        Me.lblDeduct.Caption = FormatNumber(0, 2)
    End If
    On Error GoTo err
    Me.lblNetPay.Caption = FormatNumber(CCur(Me.lblRate.Caption) + CCur(Me.lbltemp.Caption) - CCur(Me.lblDeduct.Caption), 2)
    Exit Sub
err:
    Me.lblDeduct.Caption = FormatNumber(0, 2)
End Sub

Private Sub lblRate_Change()
    If Me.lblRate.Caption = vbNullString Then
        Me.lblRate.Caption = FormatNumber(0, 2)
    End If
End Sub

Private Sub lbltemp_Change()
    If Me.lbltemp.Caption = vbNullString Then
        Me.lbltemp.Caption = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblNetPay.Caption = FormatNumber(CCur(Me.lblRate.Caption) + CCur(Me.lbltemp.Caption) - CCur(Me.lblDeduct.Caption), 2)
    Exit Sub
err:
Me.lbltemp.Caption = FormatNumber(0, 2)
End Sub

Private Sub txtAbsences_Change()
    If Me.txtAbsences.Text = vbNullString Then
        Me.txtAbsences.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    'Me.lblNetPay = FormatNumber(CCur(Me.lblRate) - (CCur(Me.lblDeduct.Caption)), 2)
    Exit Sub
err:
    Me.txtAbsences.Text = FormatNumber(0, 2)
End Sub

Private Sub txtAdvances_Change()
    If Me.txtAdvances.Text = vbNullString Then
        Me.txtAdvances.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    'Me.lblNetPay = FormatNumber(CCur(Me.lblRate) - (CCur(Me.lblDeduct.Caption)), 2)
    Exit Sub
err:
    Me.txtAdvances.Text = FormatNumber(0, 2)
End Sub

Private Sub txtAdvances_GotFocus()
    hl_Text txtAdvances
End Sub

Private Sub txtAdvances_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            txtAdvances_Change
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtAdvances_LostFocus()
    On Error GoTo err:
    Me.txtAdvances.Text = FormatNumber(Me.txtAdvances.Text, 2)
    Exit Sub
err:
    Me.txtAdvances.Text = FormatNumber(0, 2)
End Sub

Private Sub txtBonus_Change()
    If Me.txtBonus.Text = vbNullString Then
        Me.txtBonus.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lbltemp.Caption = FormatNumber(CCur(Me.txtBonus.Text) + CCur(Me.txtOverTime.Text), 2)
    Exit Sub
err:
Me.txtBonus.Text = FormatNumber(0, 2)
End Sub

Private Sub txtBonus_GotFocus()
    hl_Text txtBonus
End Sub

Private Sub txtBonus_KeyPress(KeyAscii As Integer)
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

Private Sub txtBonus_LostFocus()
    On Error GoTo err:
    Me.txtBonus.Text = FormatNumber(Me.txtBonus.Text, 2)
    Exit Sub
err:
Me.txtBonus.Text = FormatNumber(0, 2)
End Sub

Private Sub txtNumDays_Change()
    If Me.txtNumDays.Text = vbNullString Then
        Me.txtNumDays.Text = FormatNumber(0, 0)
    End If
    On Error GoTo err:
    Me.txtAbsences.Text = Me.txtNumDays.Text * Me.lblPerDay.Caption
    Exit Sub
err:
    Me.txtNumDays.Text = FormatNumber(0, 0)
End Sub

Private Sub txtNumDays_GotFocus()
    hl_Text txtNumDays
End Sub

Private Sub txtNumDays_KeyPress(KeyAscii As Integer)
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

Private Sub txtOther_Change()
    If Me.txtOther.Text = vbNullString Then
        Me.txtOther.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    Exit Sub
err:
    Me.txtOther.Text = FormatNumber(0, 2)
End Sub

Private Sub txtOther_GotFocus()
    hl_Text txtOther
End Sub

Private Sub txtOther_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            txtOther_Change
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtOther_LostFocus()
    On Error GoTo err:
    Me.txtOther.Text = FormatNumber(Me.txtOther.Text, 2)
    Exit Sub
err:
    Me.txtOther.Text = FormatNumber(0, 2)
End Sub

Private Sub txtOverTime_Change()
    If Me.txtOverTime.Text = vbNullString Then
        Me.txtOverTime.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lbltemp.Caption = FormatNumber(CCur(Me.txtBonus.Text) + CCur(Me.txtOverTime.Text), 2)
    Exit Sub
err:
    Me.txtOverTime.Text = FormatNumber(0, 2)
End Sub

Private Sub txtOverTime_GotFocus()
    hl_Text txtOverTime
End Sub

Private Sub txtOverTime_KeyPress(KeyAscii As Integer)
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

Private Sub txtOverTime_LostFocus()
    On Error GoTo err:
    Me.txtOverTime.Text = FormatNumber(Me.txtOverTime.Text, 2)
    Exit Sub
err:
    Me.txtOverTime.Text = FormatNumber(0, 2)
End Sub

Private Sub txtph_Change()
    If Me.txtph.Text = vbNullString Then
        Me.txtph.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    Exit Sub
err:
    Me.txtph.Text = FormatNumber(0, 2)
End Sub

Private Sub txtph_GotFocus()
hl_Text txtph
End Sub

Private Sub txtph_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            txtph_Change
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtph_LostFocus()
  On Error GoTo err:
    Me.txtph.Text = FormatNumber(Me.txtph.Text, 2)
    Exit Sub
err:
    Me.txtph.Text = FormatNumber(0, 2)
End Sub

Private Sub txtsss_Change()
    If Me.txtsss.Text = vbNullString Then
        Me.txtsss.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    Exit Sub
err:
    Me.txtsss.Text = FormatNumber(0, 2)
End Sub

Private Sub txtsss_GotFocus()
hl_Text txtsss
End Sub

Private Sub txtsss_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            txtsss_Change
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtsss_LostFocus()
    On Error GoTo err:
    Me.txtsss.Text = FormatNumber(Me.txtsss.Text, 2)
    Exit Sub
err:
    Me.txtsss.Text = FormatNumber(0, 2)
End Sub

Private Sub txttax_Change()
    If Me.txttax.Text = vbNullString Then
        Me.txttax.Text = FormatNumber(0, 2)
    End If
    On Error GoTo err:
    Me.lblDeduct.Caption = FormatNumber(CCur(Me.txtAdvances.Text) + CCur(Me.txtAbsences.Text) _
    + CCur(Me.txtsss.Text) + CCur(Me.txttax.Text) + CCur(Me.txtph.Text) + CCur(Me.txtOther.Text), 2)
    Exit Sub
err:
 Me.txttax.Text = FormatNumber(0, 2)
End Sub


Private Sub Init_Display_Data()

On Error GoTo err:
        If rs.State = adStateOpen Then rs.Close
        Me.lvwInfo.ListItems.Clear
        rs.Open "Select * from tblPayroll where Month(dDate)='" & Month(Me.DTPicker1.Value) & "' And Year(dDate)='" & Year(Me.DTPicker1.Value) & "' ORDER BY tblPayroll.EM_ID;", cn, adOpenKeyset, adLockPessimistic
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

Private Sub txttax_GotFocus()
hl_Text txttax
End Sub

Private Sub txttax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyBack
        Case vbKeyDelete
        Case vbKeyReturn
            txttax_Change
            SendKeys vbTab
        Case 48 To 57
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txttax_LostFocus()
  On Error GoTo err:
    Me.txttax.Text = FormatNumber(Me.txttax.Text, 2)
    Exit Sub
err:
    Me.txttax.Text = FormatNumber(0, 2)
End Sub

Private Sub sText()
    Me.txtAbsences.Text = "0.00"
    Me.txtAdvances.Text = "0.00"
    Me.txtBonus.Text = "0.00"
    Me.txtNumDays.Text = "0"
    Me.txtOverTime.Text = "0.00"
    Me.txtsss.Text = "0.00"
    Me.txttax.Text = "0.00"
    Me.txtph.Text = "0.00"
    Me.txtOther.Text = "0.00"
End Sub
