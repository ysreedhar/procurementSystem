VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_materiallevel3 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   ClientHeight    =   10095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_materialcategory 
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   5175
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   3
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   10503977
      ForeColorFixed  =   16777215
      BackColorSel    =   10503977
      BackColorBkg    =   16777215
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
   End
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
            Picture         =   "frm_materiallevel3.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":13162
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":13274
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":1D236
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":2E64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3DEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3E307
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3E75B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3EB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3F05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3FA80
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":3FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":4050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":40959
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":40E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":54DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":68E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":6C807
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":70FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":714FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_materiallevel3.frx":78DCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
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
         TabIndex        =   2
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex_item 
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   10503977
      ForeColorFixed  =   16777215
      BackColorSel    =   10503977
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00A04729&
      Caption         =   "Material Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   4
      Top             =   7200
      Width           =   11895
   End
End
Attribute VB_Name = "frm_materiallevel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cbo_materialcategory_Click()
Call flex_data
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
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture
Call flex_subtitle
Call flex_datasub

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
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture

'--END---------

Unload materiallevel3
Unload vscrollform2
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

materiallevel3.txt_categorycode.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
materiallevel3.txt_category.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
materiallevel3.cbo_uom.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
materiallevel3.cbo_type.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
materiallevel3.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
 
 
Dim rd As Integer
rd = 0
For rd = 0 To flex_item.Rows
    vscrollform2.text2(rd).Text = flex_item.TextMatrix(rd + 1, 1)
    vscrollform2.Text1(rd).Text = flex_item.TextMatrix(rd + 1, 2)
    
    vscrollform2.Text4(rd).Text = flex_item.TextMatrix(rd + 1, 3)
    vscrollform2.Text5(rd).Text = flex_item.TextMatrix(rd + 1, 4)
    vscrollform2.Combo1(rd).Text = flex_item.TextMatrix(rd + 1, 5)
    vscrollform2.Combo2(rd).Text = flex_item.TextMatrix(rd + 1, 6)
    vscrollform2.Text3(rd).Text = flex_item.TextMatrix(rd + 1, 7)
Next rd

materiallevel3.Show
materiallevel3.Top = 1900
materiallevel3.Left = 100
materiallevel3.Height = 6750
materiallevel3.Width = 11865


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
    flex_grid.CellBackColor = vbWhite
Next
End If

'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
'Set flex_grid.Row.CellPicture = ImageList1.ListImages(11).Picture
Call flex_subtitle
Call flex_datasub

vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
Call connect
main.lbltitle.Caption = "Material Master"
Call flex_title
Call flex_subtitle
Call flex_data
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
Dim mc2 As New ADODB.Recordset
If mc2.State Then mc2.Close
mc2.Open "select DISTINCT(ml2code),ml2name from ml2", Cn, 3, 2
While Not mc2.EOF
cbo_materialcategory.AddItem mc2(0) & "  -  " & mc2(1)
mc2.MoveNext
Wend
End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "Code"
        .ColWidth(1) = 2000
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Name"
        .ColWidth(2) = 4000
        .ColAlignment(2) = 0

        .TextMatrix(0, 3) = "Remarks"
        .ColWidth(3) = 6300
        .ColAlignment(3) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload materiallevel3
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then

Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload category
Unload vscrollform2
materiallevel3.Show
materiallevel3.Top = 1900
materiallevel3.Left = 100
materiallevel3.Height = 6750
materiallevel3.Width = 11865 ' to save new record
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
'validate


Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from ml3", Cn, 3, 2
sv.AddNew
sv!ml3code = materiallevel3.txt_categorycode.Text
sv!ml3name = materiallevel3.txt_category.Text
sv!ml3uom = materiallevel3.cbo_uom.Text
sv!ml3type = materiallevel3.cbo_type.Text
sv!remarks = materiallevel3.txt_notes.Text
sv.Update
sv.Close
Dim ct As Integer
ct = 0
For ct = 0 To vscrollform2.Text1.Count - 1
If Not vscrollform2.text2(ct).Text = "" Then
Dim sva As New ADODB.Recordset
If sva.State Then sva.Close
sva.Open "select * from ml4", Cn, 3, 2
sva.AddNew
sva!ml3code = materiallevel3.txt_categorycode.Text
sva!ml3name = materiallevel3.txt_category.Text
sva!ml4code = vscrollform2.text2(ct).Text
sva!ml4name = vscrollform2.Text1(ct).Text
sva!ml4itemid = vscrollform2.Text4(ct).Text
sva!ml4mrefcode = vscrollform2.Text5(ct).Text
sva!ml4uom = vscrollform2.Combo1(ct).Text
sva!ml4type = vscrollform2.Combo2(ct).Text
sva!remarks = vscrollform2.Text3(ct).Text
sva.Update
sva.Close
End If
Next ct
MsgBox "New Category Added Succesfully"
Unload materiallevel3
Call flex_data
Call flex_title
Call flex_datasub
Call flex_subtitle
Exit Sub
assad:

       MsgBox "Duplicate Entries Not Allowed"
'to modify existing medicine
ElseIf Button.Caption = "Modify" Then
On Error GoTo assad1
'validate

Toolbar1.Buttons(3).Enabled = False
Dim id1 As Double
id1 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id1 = flex_grid.TextMatrix(flex_grid.Row, 0)
Dim md As New ADODB.Recordset
If md.State Then md.Close
md.Open "select * from ml3 where ml3_id=" & id1, Cn, 3, 2
If Not md.EOF Then
'md!mcategorycode = materiallevel3.txt_categorycode.Text
md!ml3name = materiallevel3.txt_category.Text
md!ml3uom = materiallevel3.cbo_uom.Text
md!ml3type = materiallevel3.cbo_type.Text
md!remarks = materiallevel3.txt_notes.Text
md.Update
md.Close
Cn.Execute "delete from ml4 where ml3code='" & materiallevel3.txt_categorycode.Text & "'"
Dim mct As Integer
mct = 0
For mct = 0 To vscrollform2.Text1.Count - 1
If Not vscrollform2.text2(mct).Text = "" Then
Dim msva As New ADODB.Recordset
If msva.State Then msva.Close
msva.Open "select * from ml4", Cn, 3, 2
 
msva.AddNew
msva!ml3code = materiallevel3.txt_categorycode.Text
msva!ml3name = materiallevel3.txt_category.Text
msva!ml4code = vscrollform2.text2(mct).Text
msva!ml4name = vscrollform2.Text1(mct).Text
msva!ml4itemid = vscrollform2.Text4(mct).Text
msva!ml4mrefcode = vscrollform2.Text5(mct).Text
msva!ml4uom = vscrollform2.Combo1(mct).Text
msva!ml4type = vscrollform2.Combo2(mct).Text
msva!remarks = vscrollform2.Text3(mct).Text
msva.Update
msva.Close
 

End If
Next mct


MsgBox "Selected Category Modified"
End If

Unload materiallevel3
Call flex_data
Call flex_title
Call flex_datasub
Call flex_subtitle
Exit Sub
assad1:

       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete" Then
Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Dim id2 As Double
id2 = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id2 = flex_grid.TextMatrix(flex_grid.Row, 0)
Cn.Execute "delete from ml3 where ml3_id=" & id2
MsgBox "Selected Category Has Been Deleted"
Unload materiallevel3
Call flex_data
Call flex_title
Call flex_datasub
Call flex_subtitle
Else
Unload materiallevel3
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload materiallevel3
End If
End Sub
Public Sub flex_data()
'On Error Resume Next
nn = Split(cbo_materialcategory.Text, "  -  ", Len(cbo_materialcategory.Text), vbTextCompare)
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from ml3 where ml2code ='" & nn(0) & "' order by ml3code", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata("ml3code")
        .TextMatrix(.Rows - 1, 2) = fldata("ml3name")
        .TextMatrix(.Rows - 1, 3) = fldata("ml3uom")
        .TextMatrix(.Rows - 1, 4) = fldata("ml3type")
        .TextMatrix(.Rows - 1, 5) = fldata("remarks")

        fldata.MoveNext
    Wend
End With
End Sub
Public Sub flex_subtitle()
On Error Resume Next

   With flex_item
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "ML4 Code"
        .ColWidth(1) = 2000
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "ML4 Name"
        .ColWidth(2) = 4000
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "ItemId"
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Mfr.Ref Code"
        .ColWidth(4) = 4000
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "UOM"
        .ColWidth(5) = 800
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Type"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "ML4 Remarks"
        .ColWidth(7) = 4000
        .ColAlignment(7) = 0
        End With
    End Sub
Public Sub flex_datasub()
'On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from ml4 where ml3code ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "' order by ml4code", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_item
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata("ml4code")
        .TextMatrix(.Rows - 1, 2) = fldata("ml4name")
        .TextMatrix(.Rows - 1, 3) = fldata("ml4itemid")
        .TextMatrix(.Rows - 1, 4) = fldata("ml4mrefcode")
        .TextMatrix(.Rows - 1, 5) = fldata("ml4uom")
        .TextMatrix(.Rows - 1, 6) = fldata("ml4type")
        .TextMatrix(.Rows - 1, 7) = fldata("remarks")
        fldata.MoveNext
    Wend
End With
End Sub


