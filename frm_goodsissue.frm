VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_goodsissue 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   78
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":13162
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":13274
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":1D236
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":2E64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3DEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3E307
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3E75B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3EB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3F05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3FA80
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":3FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":4050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":40959
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":40E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":54DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":68E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":6C807
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":70FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":714FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_goodsissue.frx":78DCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   635
      ButtonWidth     =   1561
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Key             =   "ar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify"
            Key             =   "hlp"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8200
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   1
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   5355
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   3
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   16744576
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_item 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   6435
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   16744576
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtps_date 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM/yyyy"
      Format          =   67633155
      CurrentDate     =   38558
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Line Item Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   6195
      Width           =   11775
   End
End
Attribute VB_Name = "frm_goodsissue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents oSMTP As OSSMTP.SMTPSession
Attribute oSMTP.VB_VarHelpID = -1
Public Vmsr As String
Public StrCode As String
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub cmd_apply_Click()
Call flex_datadate
Call flex_itemmodi
End Sub
Private Sub Command1_Click()
Call flex_datadate
Call flex_itemmodi
End Sub
Private Sub Command2_Click()
Call flex_datadate
Call flex_itemmodi
End Sub



Private Sub dtps_date_Change()
Call flex_datadate
Call flex_itemmodi
End Sub
Private Sub dtps_date_Click()
Call flex_datadate
Call flex_itemmodi
End Sub
Private Sub flex_grid_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()

    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'RGB(202, 204, 221)
Next
flex_grid.Col = 1
End If

vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()

    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221)
Next
flex_grid.Col = 1

End If

Unload goodsissue
Unload vscrollgoodsissue
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


goodsissue.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
goodsissue.dtp_from.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
goodsissue.cbo_worklocationfrom.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
goodsissue.cbo_storagelocationfrom.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
goodsissue.cbo_issuereason.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
goodsissue.cbo_issueperson.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
'goodsissue.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
'-----------------------------------------------

'------------------------------------------------

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from gidetails where gino='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
id = 0
For id = 0 To prs.RecordCount - 1
        vscrollgoodsissue.cbo_category(id).Text = prs!itemid & "  -  " & prs!mrefcode & "  -  " & prs!material
        vscrollgoodsissue.cbo_batchno(id).Text = prs!batchno
        vscrollgoodsissue.cbo_uom(id).Text = prs!uom
        vscrollgoodsissue.txt_qty(id).Text = prs!tqty
        prs.MoveNext
Next id
prs.Close
'--------------------------------------


goodsissue.Show
SetParent goodsissue.hwnd, frm_goodsissue.hwnd
goodsissue.Height = 7335
goodsissue.Width = 10395
goodsissue.Top = 50
goodsissue.Left = 200

vprev = flex_grid.Row

End Sub
Private Sub flex_grid_SelChange()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()

    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbORANGE
Next
End If
If flex_grid.Row <> 0 Then
'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'RGB(202, 204, 221)
Next
flex_grid.Col = 1
End If


'''''''''''''''''''''''''''''''''''''''''''''''''


Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from gidetails where gino='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With flex_item
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!batchno
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!tqty

        prs.MoveNext
    Wend
End With
'''''''''''''''''''''''''''''''''''''''''''''''


vprev = flex_grid.Row
End Sub

Private Sub flex_item_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_item.Row

'Reset to previous row
If vprev > 0 Then
    flex_item.Row = vprev
    flex_item.Col = 1
    Set flex_item.CellPicture = LoadPicture()

    For i = 1 To flex_grid.Cols - 1
    flex_item.Col = i
    flex_item.CellBackColor = vbWhite
Next
End If

'Current  row
If flex_item.Row <> 0 Then
flex_item.Row = current
For i = 1 To flex_item.Cols - 1
flex_item.Col = i
flex_item.CellBackColor = RGB(202, 204, 221)
Next
flex_item.Col = 1
End If

vprev = flex_item.Row
End Sub
Private Sub flex_item_DblClick()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()

    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

If flex_grid.Row <> 0 Then
'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221)
Next
flex_grid.Col = 1
End If


Unload goodsissue
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


goodsissue.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
goodsissue.dtp_from.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
goodsissue.cbo_worklocationfrom.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
goodsissue.cbo_storagelocationfrom.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
goodsissue.cbo_issuereason.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
goodsissue.cbo_issueperson.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
'goodsissue.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 7)

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from gidetails where gd_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "'", Cn, 3, 2
id = 0
For id = 0 To prs.RecordCount - 1
        vscrollgoodsissue.cbo_category(id).Text = prs!itemid & "  -  " & prs!mrefcode & "  -  " & prs!material
        vscrollgoodsissue.cbo_batchno(id).Text = prs!batchno
        vscrollgoodsissue.cbo_uom(id).Text = prs!uom
        vscrollgoodsissue.txt_qty(id).Text = prs!tqty
        prs.MoveNext
Next id
prs.Close


goodsissue.Show
SetParent goodsissue.hwnd, frm_goodsissue.hwnd
goodsissue.Height = 7335
goodsissue.Width = 10395
goodsissue.Top = 50
goodsissue.Left = 200

'goodsissue.kl = 0
vprev = flex_grid.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
dtps_date.Value = Format(Date, "dd/MM/yyyy")
Call connect
main.lbltitle.Caption = "Goods Issue"

Call flex_title
Call flex_subtitle
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
Call flex_datadate


End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "GI No."
        .ColWidth(1) = 1700
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Date"
        .ColWidth(2) = 1000
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "From WL"
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "From SL"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0

        .TextMatrix(0, 5) = "Reason"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
       
        .TextMatrix(0, 6) = "Person"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Notes"
        .ColWidth(7) = 2000
        .ColAlignment(7) = 0

        
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload goodsissue
Unload vscrollgoodsissue
End Sub
Private Sub TabStrip1_Click()
Call flex_datadate
End Sub

Private Sub Label2_Click()
Call flex_datadate
Call flex_itemmodi
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call flex_datadate
Call flex_itemmodi
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

If Button.Caption = "New" Then

Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload goodsissue
Unload vscrollgoodsissue
goodsissue.Show
SetParent goodsissue.hwnd, frm_goodsissue.hwnd
goodsissue.Height = 7335
goodsissue.Width = 10395
goodsissue.Top = 50
goodsissue.Left = 200
' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
'validate
'20726333
'-----------
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from goodsissue", Cn, 3, 2
sv.AddNew
sv!gino = goodsissue.txt_account.Text
sv!issuedate = goodsissue.dtp_from.Value
sv!worklocationfrom = goodsissue.cbo_worklocationfrom.Text
sv!storagelocationfrom = goodsissue.cbo_storagelocationfrom.Text
sv!issuereason = goodsissue.cbo_issuereason.Text
sv!issueperson = goodsissue.cbo_issueperson.Text
'sv!Notes = goodsissue.txt_notes.Text

Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from gidetails", Cn, 3, 2

Dim j As Integer

    j = 0
        
For j = 0 To vscrollgoodsissue.cbo_category.Count - 1
If vscrollgoodsissue.cbo_category(j).Text <> "" Then
        
        pr.AddNew
        spt = Split(vscrollgoodsissue.cbo_category(j).Text, "  -  ", Len(vscrollgoodsissue.cbo_category(j).Text), vbTextCompare)
        '-----------------------------
        
If goodsissue.cbo_lookup.Text = "Item ID" Then
      pr!itemid = spt(0)
      pr!mrefcode = spt(1)
      pr!material = spt(2) & "  -  " & spt(3)
ElseIf goodsissue.cbo_lookup.Text = "Mfr PartNo." Then
      pr!itemid = spt(1)
      pr!mrefcode = spt(0)
      pr!material = spt(2) & "  -  " & spt(3)

ElseIf goodsissue.cbo_lookup.Text = "Item Description" Then
      pr!itemid = spt(2)
      pr!mrefcode = spt(3)
      pr!material = spt(0) & "  -  " & spt(1)

ElseIf goodsissue.cbo_lookup.Text = "Search" Then
      pr!itemid = spt(2)
      pr!mrefcode = spt(3)
      pr!material = spt(0) & "  -  " & spt(1)
End If
        '-----------------------------
             
             
      pr!batchno = vscrollgoodsissue.cbo_batchno(j).Text
      pr!uom = vscrollgoodsissue.cbo_uom(j).Text
      pr!tqty = vscrollgoodsissue.txt_qty(j).Text
      pr!gino = goodsissue.txt_account.Text
      'pr!giid = flex_grid.TextMatrix(flex_grid.Row, 0)
      pr.Update
        
End If
      
Next j
pr.Close
sv.Update
sv.Close
MsgBox "Goods Issue Recorded"
Unload goodsissue
Unload vscrollgoodsissue
Call flex_datadate
Call flex_title
Call flex_itemmodi
Exit Sub
assad:

       MsgBox "Duplicate Entries Not Allowed"
'to modify existing goodsissue
ElseIf Button.Caption = "Modify" Then

ElseIf Button.Caption = "Delete" Then




Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Cn.Execute "delete from goodsissue where gino='" & goodsissue.txt_account.Text & "'"

MsgBox "Selected Record Has Been Deleted"
Unload goodsissue
Unload vscrollgoodsissue
Call flex_datadate
Call flex_title
Else
Unload goodsissue
Unload vscrollgoodsissue
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload goodsissue
End If
End Sub

Public Sub flex_datadate()
'On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from goodsissue where month(issuedate)='" & Format(dtps_date.Value, "mm") & "' and year(issuedate)='" & Format(dtps_date.Value, "yyyy") & "' ", Cn, 3, 2 'and tuser = '" & main.Label2.Caption & "' order by prno,prdate

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!gino
        .TextMatrix(.Rows - 1, 2) = fldata!issuedate
        .TextMatrix(.Rows - 1, 3) = fldata!worklocationfrom
        .TextMatrix(.Rows - 1, 4) = fldata!storagelocationfrom
        .TextMatrix(.Rows - 1, 5) = fldata!issuereason
        .TextMatrix(.Rows - 1, 6) = fldata!issueperson
        '.TextMatrix(.Rows - 1, 7) = fldata!Notes
                
        fldata.MoveNext
    Wend
End With
fldata.Close

End Sub
Public Sub flex_subtitle()
On Error Resume Next

   With flex_item
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "ItemId"
        .ColWidth(1) = 1700
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Mfr.Ref Code"
        .ColWidth(2) = 1700
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Material"
        .ColWidth(3) = 5000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Batch"
        .ColWidth(4) = 800
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "UOM"
        .ColWidth(5) = 800
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Qty"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        
    End With
End Sub
Public Sub flex_itemmodi()
On Error Resume Next
current = flex_grid.Row

'Reset to previous row
'''''''''''''''''''''''''''''''''''''''''''''''''
If flex_grid.TextMatrix(flex_grid.Row, 1) = "" Then Exit Sub
If flex_grid.TextMatrix(flex_grid.Row, 1) = Null Then Exit Sub

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from gidetails where gino='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With flex_item
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!batchno
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!tqty
        
        prs.MoveNext
    Wend
End With
prs.Close
'''''''''''''''''''''''''''''''''''''''''''''''

vprev = flex_grid.Row
End Sub


