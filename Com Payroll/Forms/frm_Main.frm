VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BC209F51-C609-11D9-8858-F49DACE6A63F}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm frm_Main 
   BackColor       =   &H8000000C&
   Caption         =   "Payroll System"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12675
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBak 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   800
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   12615
      TabIndex        =   2
      Top             =   570
      Width           =   12675
      Begin VB.Image image 
         Height          =   1215
         Left            =   0
         Picture         =   "frm_Main.frx":64692
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1215
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   10365
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   609
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   2
            Text            =   "User"
            TextSave        =   "User"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "Automated Payroll System"
            TextSave        =   "Automated Payroll System"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "6/1/2004"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:51 AM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin HookMenu.ctxHookMenu hkMenu 
      Left            =   0
      Top             =   480
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   13
      Bmp:1           =   "frm_Main.frx":A486C
      Mask:1          =   16777215
      Key:1           =   "#mnuNewEmp"
      Bmp:2           =   "frm_Main.frx":A4DBE
      Mask:2          =   16777215
      Key:2           =   "#mnuCPayroll"
      Bmp:3           =   "frm_Main.frx":A5310
      Mask:3          =   16777215
      Key:3           =   "#mnuCPayslip"
      Bmp:4           =   "frm_Main.frx":A5862
      Mask:4          =   16777215
      Key:4           =   "#mnuLogOut"
      Bmp:5           =   "frm_Main.frx":A5DB4
      Mask:5          =   16777215
      Key:5           =   "#mnuExit"
      Bmp:6           =   "frm_Main.frx":A6306
      Mask:6          =   16777215
      Key:6           =   "#mnuEmpInfo"
      Bmp:7           =   "frm_Main.frx":A6858
      Mask:7          =   16777215
      Key:7           =   "#mnuPayrollReport"
      Bmp:8           =   "frm_Main.frx":A6DAA
      Mask:8          =   16777215
      Key:8           =   "#mnuPayslip"
      Bmp:9           =   "frm_Main.frx":A72FC
      Mask:9          =   16777215
      Key:9           =   "#mnuUserLogin"
      Bmp:10          =   "frm_Main.frx":A784E
      Mask:10         =   6181963
      Key:10          =   "#mnuBackup"
      Bmp:11          =   "frm_Main.frx":A7DA0
      Mask:11         =   6181963
      Key:11          =   "#mnuRestore"
      Bmp:12          =   "frm_Main.frx":A82F2
      Mask:12         =   16777215
      Key:12          =   "#mnuAbout"
      Bmp:13          =   "frm_Main.frx":A8844
      Mask:13         =   6181705
      Key:13          =   "#mnuEmpDTR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DrawStyle       =   2
      MenuDrawStyle   =   2
      UserSelectedMenuBackColour=   12648384
      UserSelectedMenuBorderColour=   16777152
      UserTopMenuBackColour=   16777088
      UserTopMenuSelectedColour=   12648384
      UserTopMenuHotColour=   16761087
      UserMenuBorderColour=   8421504
      UserGradientOne =   12640511
      UserGradientTwo =   16777215
      UserUseGradient =   -1  'True
      UserUseTopMenuGradient=   -1  'True
      UserSideBarColour=   16777088
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Employee"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Create Payroll"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee Information"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Employee DTR"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Payslip"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New System User"
            ImageIndex      =   12
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Book report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Member report"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Issue report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Backup Data"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Restore Data"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "About"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Log off"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      MouseIcon       =   "frm_Main.frx":A8D96
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":A8EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":A9BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":AA8AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":AB586
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":AC260
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":ACF3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":ADC14
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":AE8EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":AF5C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B02A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B0F7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B1C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B2930
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B360A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B42E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B4FBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B5C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":B6972
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   1560
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   125
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B6C01
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B7153
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B76A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B7BF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B8149
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B869B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B8BED
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B913F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B9691
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":B9BE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BA135
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BA687
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BABD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BB12B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BB67D
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BBBCF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BC121
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BC673
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BC785
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BCCD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BD229
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BD77B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BDCCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BE21F
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BE771
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BECC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BF215
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BF767
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":BFCB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C020B
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C075D
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C0CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C1201
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C1313
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C1425
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C1977
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C1EC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C241B
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C296D
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C2EBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C3411
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C3963
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C3EB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C4407
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C4959
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C4EAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C53FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C594F
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C5EA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C63F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C6945
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C6E97
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C73E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C793B
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C7E8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C83DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C8931
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C8E83
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C93D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C9927
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":C9E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CA3CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CA91D
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CAE6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CB3C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CB913
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CBE65
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CC3B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CC909
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CCE5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CD3AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CD8FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CDE51
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CE3A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CE8F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CEE47
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CF399
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CF8EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":CFE3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D038F
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D04A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D09F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D0F45
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D1497
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D19E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D1F3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D248D
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D29DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D2F31
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D3483
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D39D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D3F27
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D4479
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D49CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D4F1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D546F
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D59C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D5F13
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D6465
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D69B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D6F09
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D745B
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D79AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D7EFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D8451
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D89A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D8EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D9447
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D9999
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":D9EEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DA43D
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DA98F
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DAEE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DB433
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DB985
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DBED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DC429
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DC97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DCECD
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DD41F
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DD971
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DDEC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DE415
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DE527
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Main.frx":DEA79
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuNewEmp 
         Caption         =   "&New Employee"
      End
      Begin VB.Menu mnuSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCPayroll 
         Caption         =   "Create &Payroll"
      End
      Begin VB.Menu mnuSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log &Out"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuEmpInfo 
         Caption         =   "&Employee Info"
      End
      Begin VB.Menu mnuSpt5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpDTR 
         Caption         =   "Employee DTR"
      End
      Begin VB.Menu mnuSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPayslip 
         Caption         =   "Pay&slip"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuUserLogin 
         Caption         =   "&User Login"
      End
      Begin VB.Menu mnuSpt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup &Data"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "&Restore Data"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xLogOff As Boolean

Private Sub MDIForm_Load()
    xLogOff = False
End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
    Me.picBak.Height = (frm_Main.Height - Me.Toolbar.Height) - 1150
    Me.image.Left = 0
    Me.image.Width = Me.picBak.Width
    Me.image.Height = Me.picBak.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim strMess
If xLogOff = False Then
    strMess = "You are about to close the Automated Payroll System." & vbCrLf & vbCrLf & "Are you sure?"
    If MsgBox(strMess, vbQuestion + vbYesNo, "Exit Confirmation") = vbYes Then
        End
    Else
        Cancel = 1
    End If
ElseIf xLogOff = True Then
    frmLogin.Show
    Unload Me
End If
End Sub

Private Sub mnuAbout_Click()
    frm_About.Show 1
End Sub

Private Sub mnuBackup_Click()
    frm_Backup.Show 1
End Sub

Private Sub mnuCPayroll_Click()
    frm_Payroll.Show 1
End Sub

Private Sub mnuEmpDTR_Click()
    frm_Payroll_dtr.Show 1
End Sub

Private Sub mnuEmpInfo_Click()
    frm_emp_Info.Show 1
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuLogOut_Click()
    xLogOff = True
    Unload Me
End Sub

Private Sub mnuNewEmp_Click()
    frm_New_emp.Show 1
End Sub

Private Sub mnuPayslip_Click()
    frm_Payslip.Show 1
End Sub

Private Sub mnuRestore_Click()
    frmRestore.Show 1
End Sub

Private Sub mnuUserLogin_Click()
    frm_Sys_Users.Show 1
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 2
        mnuNewEmp_Click
    Case 4
        mnuCPayroll_Click
    Case 6
        mnuEmpInfo_Click
    Case 7
        mnuEmpDTR_Click
    Case 8
        mnuPayslip_Click
    Case 10
        mnuUserLogin_Click
    Case 12
        mnuBackup_Click
    Case 13
        mnuRestore_Click
    Case 15
        mnuAbout_Click
    Case 17
        mnuLogOut_Click
     Case 18
        mnuExit_Click
    End Select
End Sub
