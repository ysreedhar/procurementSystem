VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_stocktracking 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10920
   ScaleWidth      =   12495
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   16113
      _Version        =   393216
      Rows            =   3
      Cols            =   17
      FixedCols       =   0
      RowHeightMin    =   300
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   16744576
      ForeColorFixed  =   16777215
      BackColorSel    =   16744576
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
            Picture         =   "frm_stocktracking.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":13162
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":13274
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":1D236
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":2E64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3DEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3E307
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3E75B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3EB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3F05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3FA80
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":3FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":4050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":40959
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":40E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":54DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":68E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":6C807
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":70FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":714FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_stocktracking.frx":78DCA
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
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   635
      ButtonWidth     =   1402
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
End
Attribute VB_Name = "frm_stocktracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pxin As Double

Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub flex_grid_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = True


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
 


vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = True

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


Unload stocktracking
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

stocktracking.cbo_category.Text = flex_grid.TextMatrix(flex_grid.Row, 1) & "  -  " & flex_grid.TextMatrix(flex_grid.Row, 2) & "  -  " & flex_grid.TextMatrix(flex_grid.Row, 3)
stocktracking.cbo_batchno.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
stocktracking.txt_uom_os.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
stocktracking.txt_qty_os.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
stocktracking.txt_uom_grn.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
stocktracking.txt_qty_grn.Text = flex_grid.TextMatrix(flex_grid.Row, 8)

stocktracking.txt_uom_gi.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
stocktracking.txt_qty_gi.Text = flex_grid.TextMatrix(flex_grid.Row, 10)

stocktracking.txt_uom_grj.Text = flex_grid.TextMatrix(flex_grid.Row, 11)
stocktracking.txt_qty_grj.Text = flex_grid.TextMatrix(flex_grid.Row, 12)
stocktracking.txt_uom_gr.Text = flex_grid.TextMatrix(flex_grid.Row, 13)
stocktracking.txt_qty_gr.Text = flex_grid.TextMatrix(flex_grid.Row, 14)
stocktracking.cbo_reorder.Text = flex_grid.TextMatrix(flex_grid.Row, 15)
stocktracking.txt_qty_reorder.Text = flex_grid.TextMatrix(flex_grid.Row, 16)

stocktracking.Show
SetParent stocktracking.hwnd, frm_stocktracking.hwnd

stocktracking.Top = 200
stocktracking.Left = 300
stocktracking.Height = 6150
stocktracking.Width = 7440

 
'stocktracking.txt_batchno.Enabled = False
vprev = flex_grid.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
main.lbltitle.Caption = "STOCK TRACKING"
Call flex_title
Call flex_data
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = True

Me.Top = 5
Me.Left = 5
 Me.Width = 11415
 Me.Height = 9750
End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "ItemID"
        .ColWidth(1) = 800
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "MrefCode"
        .ColWidth(2) = 800
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Material"
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Batch"
        .ColWidth(4) = 1000
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "OS UOM"
        .ColWidth(5) = 1000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "OS QTY"
        .ColWidth(6) = 1000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "GRN UOM"
        .ColWidth(7) = 1000
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "GRN QTY"
        .ColWidth(8) = 1000
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "GI UOM"
        .ColWidth(9) = 1000
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "GI QTY"
        .ColWidth(10) = 1000
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "GRJ UOM"
        .ColWidth(11) = 1000
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "GRJ QTY"
        .ColWidth(12) = 1000
        .ColAlignment(12) = 0
        .TextMatrix(0, 13) = "GB UOM"
        .ColWidth(13) = 1000
        .ColAlignment(13) = 0
        .TextMatrix(0, 14) = "GB QTY"
        .ColWidth(14) = 1000
        .ColAlignment(14) = 0
        .TextMatrix(0, 15) = "Reorder UOM"
        .ColWidth(15) = 1000
        .ColAlignment(15) = 0
        .TextMatrix(0, 16) = "Reorder QTY"
        .ColWidth(16) = 1000
        .ColAlignment(16) = 0
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload stocktracking
End Sub
Public Sub flex_data()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from materialbatch order by batchno", Cn, 3, 2

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!itemid
        .TextMatrix(.Rows - 1, 2) = fldata!mrefcode
        .TextMatrix(.Rows - 1, 3) = fldata!material
        .TextMatrix(.Rows - 1, 4) = fldata!batchno
        .TextMatrix(.Rows - 1, 5) = fldata!uom
        .TextMatrix(.Rows - 1, 6) = fldata!initialstock
        
             Dim RSgrn As New ADODB.Recordset
             If RSgrn.State Then RSgrn.Close
             RSgrn.Open "select SUM(qtyrec) from grndetails where material ='" & fldata!material & "' and batchno='" & fldata!batchno & "' ", Cn, 3, 2
             If Not RSgrn.EOF Then
             .TextMatrix(.Rows - 1, 7) = fldata!uom
             .TextMatrix(.Rows - 1, 8) = RSgrn(0)
             End If
             RSgrn.Close
             RSgrn.Open "select SUM(tqty) from gidetails where material='" & fldata!material & "' and batchno='" & fldata!batchno & "'", Cn, 3, 2
             If Not RSgrn.EOF Then
             .TextMatrix(.Rows - 1, 9) = fldata!uom
             .TextMatrix(.Rows - 1, 10) = RSgrn(0)
             End If
             RSgrn.Close
             RSgrn.Open "select SUM(qtyrej) from grndetails where material ='" & fldata!material & "' and batchno='" & fldata!batchno & "' ", Cn, 3, 2
             If Not RSgrn.EOF Then
             .TextMatrix(.Rows - 1, 11) = fldata!uom
             .TextMatrix(.Rows - 1, 12) = RSgrn(0)
             End If
             RSgrn.Close
             Dim istck As Double
             istck = 0
             Dim grstck As Double
             grstck = 0
             Dim gistck As Double
             gistck = 0
             istck = .TextMatrix(.Rows - 1, 6)
             grstck = .TextMatrix(.Rows - 1, 8)
             gistck = .TextMatrix(.Rows - 1, 10)
        .TextMatrix(.Rows - 1, 14) = (CDbl(istck) + CDbl(grstck)) - CDbl(gistck)
        .TextMatrix(.Rows - 1, 13) = fldata!uom
        
        .TextMatrix(.Rows - 1, 16) = fldata!reorderqty
        .TextMatrix(.Rows - 1, 15) = fldata!reorderuom
        fldata.MoveNext
    Wend
End With
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next
If Button.Caption = "Save" Then
On Error GoTo assad
'validate
'-------------------
                        


spl = Split(stocktracking.cbo_category.Text, "  -  ", Len(stocktracking.cbo_category.Text), vbTextCompare)

'------------------

 
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from materialbatch where batchno='" & stocktracking.cbo_batchno.Text & "'", Cn, 3, 2
If Not sv.EOF Then
sv!reorderqty = stocktracking.txt_qty_reorder.Text
sv!reorderuom = stocktracking.cbo_reorder.Text
sv.Update
End If
sv.Close
MsgBox "Re-Order Level Updated Succesfully"
Unload stockposting
Call flex_data
Call flex_title
Exit Sub
assad:
       
       MsgBox "Duplicate Entries Not Allowed"
'to modify existing record


ElseIf Button.Caption = "Close" Then
Unload Me
Unload stockposting
End If

End Sub


