VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_delivery 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   10350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10350
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedCols       =   0
      RowHeightMin    =   350
      BackColor       =   12640511
      ForeColor       =   12582912
      BackColorFixed  =   16761024
      ForeColorFixed  =   4210816
      BackColorSel    =   16761024
      BackColorBkg    =   16761024
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
            Picture         =   "frm_delivery.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":13162
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":13274
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":1D236
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":2E64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3DEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3E307
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3E75B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3EB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3F05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3FA80
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":3FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":4050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":40959
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":40E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":54DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":68E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":6C807
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":70FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":714FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_delivery.frx":78DCA
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
      Width           =   12180
      _ExtentX        =   21484
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
   Begin MSFlexGridLib.MSFlexGrid flex_med 
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4048
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      RowHeightMin    =   250
      BackColor       =   12640511
      ForeColor       =   12582912
      BackColorFixed  =   16761024
      ForeColorFixed  =   4210816
      BackColorBkg    =   16761024
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Line Item Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7080
      Width           =   11775
   End
End
Attribute VB_Name = "frm_delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Private Sub cmd_exit_Click()
'On Error Resume Next
'Unload Me
'End Sub
'
'Private Sub flex_grid_Click()
'On Error Resume Next
''back color
'Toolbar1.Buttons(3).Enabled = False
'Toolbar1.Buttons(5).Enabled = True
'Toolbar1.Buttons(7).Enabled = True
'Static vprev As Integer
'
'current = flex_grid.Row
'
''Reset to previous row
'If vprev > 0 Then
'    flex_grid.Row = vprev
'    flex_grid.Col = 1
'    Set flex_grid.CellPicture = LoadPicture()
'
'    For i = 1 To flex_grid.Cols - 1
'    flex_grid.Col = i
'    flex_grid.CellBackColor = vbWhite
'Next
'End If
'
''Current  row
'flex_grid.Row = current
'For i = 1 To flex_grid.Cols - 1
'flex_grid.Col = i
'flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
'Next
'flex_grid.Col = 1
'
'
'vprev = flex_grid.Row
'End Sub
'
'Private Sub flex_grid_DblClick()
'On Error Resume Next
''back color
'Toolbar1.Buttons(3).Enabled = False
'Toolbar1.Buttons(5).Enabled = True
'Toolbar1.Buttons(7).Enabled = True
'Static vprev As Integer
'
'current = flex_grid.Row
'
''Reset to previous row
'If vprev > 0 Then
'    flex_grid.Row = vprev
'    flex_grid.Col = 1
'    Set flex_grid.CellPicture = LoadPicture()
'
'    For i = 1 To flex_grid.Cols - 1
'    flex_grid.Col = i
'    flex_grid.CellBackColor = vbWhite
'Next
'End If
'
''Current  row
'flex_grid.Row = current
'For i = 1 To flex_grid.Cols - 1
'flex_grid.Col = i
'flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
'Next
'flex_grid.Col = 1
'
'
'
'Unload delivery
'Dim id As Double
'id = 0
'If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
'id = flex_grid.TextMatrix(flex_grid.Row, 0)
'
'
'delivery.cbo_po.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
'delivery.dtp_del.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
'delivery.txt_invoice.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
'delivery.cbo_vendor.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
'delivery.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
'
'
'Dim prs As New ADODB.Recordset
'If prs.State Then prs.Close
'prs.Open "select * from deliverydetails where invoice='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
'With delivery.flex_med
'    .Rows = 1
'    While Not prs.EOF
'        .Rows = .Rows + 1
'
'        .TextMatrix(.Rows - 1, 1) = prs(1)
'        .TextMatrix(.Rows - 1, 2) = prs(2)
'        .TextMatrix(.Rows - 1, 3) = prs(3)
'        .TextMatrix(.Rows - 1, 4) = prs(4)
'        .TextMatrix(.Rows - 1, 5) = prs(5)
'
'        prs.MoveNext
'    Wend
'End With
'
'delivery.Show
'delivery.Top = 3200
'delivery.Left = 0
'delivery.Height = 6720
'delivery.Width = 11070
'
'
'
'vprev = flex_grid.Row
'
'End Sub
'
'Private Sub Form_Load()
'On Error Resume Next
'Call connect
'main.lbltitle.Caption = "Goods Received Note"
'Call flex_title
'Call flex_data
'Toolbar1.Buttons(3).Enabled = False
'Toolbar1.Buttons(5).Enabled = False
'Toolbar1.Buttons(7).Enabled = False
'Me.Top = 5
'Me.Left = 5
'
'End Sub
'Public Sub flex_title()
'On Error Resume Next
'
'   With flex_grid
'        .Row = 0:    .Col = 0
'        .ColWidth(0) = 0
'
'        .TextMatrix(0, 1) = "GRN No."
'        .ColWidth(1) = 1200
'        .ColAlignment(1) = 0
'        .TextMatrix(0, 2) = "GRN Date"
'        .ColWidth(2) = 2000
'        .ColAlignment(2) = 0
'        .TextMatrix(0, 3) = "PO No."
'        .ColWidth(3) = 2000
'        .ColAlignment(3) = 0
'        .TextMatrix(0, 4) = "Vendor"
'        .ColWidth(4) = 4000
'        .ColAlignment(4) = 0
'
'
'        .TextMatrix(0, 5) = "Notes"
'        .ColWidth(5) = 4000
'        .ColAlignment(5) = 0
'    End With
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'main.lbltitle.Caption = ""
'Unload delivery
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error Resume Next
'If Button.Caption = "New" Then
'
'Toolbar1.Buttons(3).Enabled = True
'Toolbar1.Buttons(5).Enabled = False
'Toolbar1.Buttons(7).Enabled = False
'Unload delivery
'delivery.Show
'delivery.Top = 3200
'delivery.Left = 0
'delivery.Height = 6720
'delivery.Width = 11070
'' to save new record
'ElseIf Button.Caption = "Save" Then
''On Error GoTo assad
''validate
'
'
'Dim sv As New ADODB.Recordset
'If sv.State Then sv.Close
'sv.Open "select * from delivery", Cn, 3, 2
'sv.AddNew
'sv!pono = delivery.cbo_po.Text
'sv!deldate = delivery.dtp_del.Value
'sv!invoice = delivery.txt_invoice.Text
'sv!vendor = delivery.cbo_vendor.Text
'sv!notes = delivery.txt_notes.Text
'
'Dim pr As New ADODB.Recordset
'If pr.State Then pr.Close
'pr.Open "select * from deliverydetails", Cn, 3, 2
'
'Dim j As Integer
'
'    j = 1
'        For j = 1 To delivery.flex_med.Rows - 1
'
'        pr.AddNew
'        pr!invoice = delivery.txt_invoice.Text
'
'        pr!Name = delivery.flex_med.TextMatrix(j, 1)
'        pr!uom = delivery.flex_med.TextMatrix(j, 2)
'        pr!qty = delivery.flex_med.TextMatrix(j, 3)
'        pr!reqdate = delivery.flex_med.TextMatrix(j, 4)
'        pr!batchno = delivery.flex_med.TextMatrix(j, 5)
'        pr.Update
'Dim tt As Double
'                                        tt = 0
'                                        Dim sm1 As New ADODB.Recordset
'                                        If sm1.State Then sm1.Close
'                                        sm1.Close
'                                        sm1.Open "select SUM(qty) from deliverydetails where batchno = '" & delivery.flex_med.TextMatrix(j, 5) & "'", Cn, 3, 2
'                                        If Not sm1.EOF Then
'                                        If IsNull(sm1(0)) Then
'                                        tt = 0
'                                        Else
'                                        tt = sm1(0)
'                                        End If
'                                        End If
'                                        sm1.Close
'
'
'                                        Dim mdb As New ADODB.Recordset
'                                        If mdb.State Then mdb.Close
'                                        mdb.Open "select * from medicinebatch where batchno='" & delivery.flex_med.TextMatrix(j, 5) & "' ", Cn, 3, 2
'                                        If Not mdb.EOF Then
'
'                                        mdb!total = tt
'                                        mdb.Update
'                                        End If
'
'
'
'        pr.MoveNext
'        Next j
'
'
'
'sv.Update
'sv.Close
'MsgBox "New Record Added Succesfully"
'Unload delivery
'Call flex_data
'Call flex_title
'Exit Sub
'assad:
'
'       MsgBox "Duplicate Entries Not Allowed"
''to modify existing delivery
'ElseIf Button.Caption = "Modify" Then
''On Error GoTo assad1
''validate
'
'Cn.Execute "delete from delivery where invoice='" & delivery.txt_invoice.Text & "'"
'Cn.Execute "delete from deliverydetails where invoice='" & delivery.txt_invoice.Text & "'"
'Dim svv As New ADODB.Recordset
'If svv.State Then svv.Close
'svv.Open "select * from delivery", Cn, 3, 2
'svv.AddNew
'svv!pono = delivery.cbo_po.Text
'svv!deldate = delivery.dtp_del.Value
'svv!invoice = delivery.txt_invoice.Text
'svv!vendor = delivery.cbo_vendor.Text
'svv!notes = delivery.txt_notes.Text
'
'
'Dim prr As New ADODB.Recordset
'If prr.State Then prr.Close
'prr.Open "select * from deliverydetails", Cn, 3, 2
'
'Dim k As Integer
'
'    k = 1
'        For k = 1 To delivery.flex_med.Rows - 1
'
'        prr.AddNew
'        prr!invoice = delivery.txt_invoice.Text
'
'        prr!Name = delivery.flex_med.TextMatrix(k, 1)
'        prr!uom = delivery.flex_med.TextMatrix(k, 2)
'        prr!qty = delivery.flex_med.TextMatrix(k, 3)
'        prr!reqdate = delivery.flex_med.TextMatrix(k, 4)
'        prr!batchno = delivery.flex_med.TextMatrix(k, 5)
'         prr.Update
'Cn.Execute "update medicinebatch set total ='" & delivery.s2 & "' where batchno='" & delivery.flex_med.TextMatrix(k, 5) & "'"
'
'        prr.MoveNext
'        Next k
'
'
'
'svv.Update
'svv.Close
'MsgBox "Selected Record Modify Succesfully"
'Unload delivery
'Call flex_data
'Call flex_title
'Exit Sub
'assad1:
'
'       MsgBox "Duplicate Entries Not Allowed"
''to delete
'ElseIf Button.Caption = "Delete" Then
'
'
'
'
'Toolbar1.Buttons(3).Enabled = False
'dlt = MsgBox("Do you want to Delete", vbYesNo)
'If dlt = vbYes Then
'Cn.Execute "delete from delivery where invoice='" & delivery.txt_invoice.Text & "'"
'Cn.Execute "delete from deliverydetails where invoice='" & delivery.txt_invoice.Text & "'"
'MsgBox "Selected Record Has Been Deleted"
'Unload delivery
'Call flex_data
'Call flex_title
'Else
'Unload delivery
'End If
'ElseIf Button.Caption = "Close" Then
'Unload Me
'Unload delivery
'End If
'
'
'
'
'End Sub
'
'Public Sub flex_data()
''On Error Resume Next
'Dim fldata As New ADODB.Recordset
'If fldata.State Then fldata.Close
'fldata.Open "select * from delivery order by deldate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key
'
'With flex_grid
'    .Rows = 1
'    While Not fldata.EOF
'        .Rows = .Rows + 1
'        .TextMatrix(.Rows - 1, 0) = fldata(0)
'        .TextMatrix(.Rows - 1, 1) = fldata(3)
'        .TextMatrix(.Rows - 1, 2) = fldata(2)
'        .TextMatrix(.Rows - 1, 3) = fldata(1)
'        .TextMatrix(.Rows - 1, 4) = fldata(4)
'        .TextMatrix(.Rows - 1, 5) = fldata(5)
'
'        fldata.MoveNext
'    Wend
'End With
'End Sub
'
'
'
'
'
'
'
'
'
