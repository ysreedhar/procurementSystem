VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_vendor 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   10245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12060
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10245
   ScaleWidth      =   12060
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   16325
      _Version        =   393216
      Rows            =   3
      Cols            =   14
      FixedCols       =   0
      RowHeightMin    =   350
      BackColor       =   16777215
      ForeColor       =   10503977
      BackColorFixed  =   10503977
      ForeColorFixed  =   16777215
      BackColorSel    =   10503977
      ForeColorSel    =   10503977
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
            Picture         =   "frm_vendor.frx":0000
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":09B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":0E08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":74F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":780E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":7B28
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":80C2
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":865C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":8BF6
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":9190
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":92A2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":97E4
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":9D7E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":A318
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":ABF2
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":AD04
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":AE16
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":AF28
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":B03A
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":B14C
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":B25E
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":B7F8
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":BD92
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":C32C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":C8C6
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":C9D8
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":CAEA
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":D084
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":D196
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":D2A8
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":D842
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":D954
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":DEEE
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":E488
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":E59A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":EB34
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":F0CE
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":F668
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":F77A
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":FD14
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":FE26
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":FF38
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1004A
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1015C
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1026E
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":10808
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1091A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":10A2C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":10FC6
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":11560
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":11AFA
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":12094
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1262E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":12BC8
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":13162
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":13274
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":1D236
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":2E64E
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3DEB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3E307
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3E75B
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3EB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3F05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3F527
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3FA80
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":3FF56
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":4050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":40959
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":40E76
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":54DA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":68E03
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":6C807
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":70FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":714FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_vendor.frx":78DCA
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
      Width           =   12060
      _ExtentX        =   21273
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
End
Attribute VB_Name = "frm_vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim v As Integer
Public fl As Integer
Public ia As Integer
Public ih As Integer
Private Sub cmd_exit_Click()
On Error Resume Next
Unload Me
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
 


vprev = flex_grid.Row
End Sub

Private Sub flex_grid_DblClick()
On Error Resume Next
'back color
ih = 0
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


Unload vendor
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

vendor.txt_code.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
vendor.txt_name.Text = flex_grid.TextMatrix(flex_grid.Row, 2)
vendor.txt_regno.Text = flex_grid.TextMatrix(flex_grid.Row, 3)

vendor.txt_address.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
vendor.txt_phone.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
vendor.txt_fax.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
vendor.txt_email.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
vendor.txt_website.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
vendor.cbo_companystatus.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
vendor.txt_bm.Text = flex_grid.TextMatrix(flex_grid.Row, 10)
vendor.txt_nb.Text = flex_grid.TextMatrix(flex_grid.Row, 11)
vendor.txt_fl.Text = flex_grid.TextMatrix(flex_grid.Row, 12)
vendor.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 13)


Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select * from activityresponsible where avendor ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "' ", Cn, 3, 2
If Not vn.EOF Then
vendor.txt_rpin.Text = vn!rpin
vendor.txt_rdesig.Text = vn!rdesig
vendor.txt_ppin.Text = vn!ppin
vendor.txt_pdesig.Text = vn!pdesig
vendor.txt_epin.Text = vn!epin
vendor.txt_edesig.Text = vn!edesig
End If
vn.Close
vn.Open "select * from vendorbankdetails where bvendor ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "' ", Cn, 3, 2
If Not vn.EOF Then
vendor.txt_bank1.Text = vn!bank1
vendor.txt_bankaccount1.Text = vn!account1
vendor.txt_branch1.Text = vn!branch1
vendor.txt_bank2.Text = vn!bank2
vendor.txt_bankaccount2.Text = vn!account2
vendor.txt_branch2.Text = vn!branch2
vendor.txt_bank3.Text = vn!bank3
vendor.txt_bankaccount3.Text = vn!account3
vendor.txt_branch3.Text = vn!branch3
End If
vn.Close
vn.Open "select * from vendortermsofpayment where tvendor ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "' ", Cn, 3, 2
If Not vn.EOF Then
vendor.txt_ddays1.Text = vn!ddays1
vendor.txt_dper1.Text = vn!dper1
vendor.txt_ddays2.Text = vn!ddays2
vendor.txt_dper2.Text = vn!dper2
vendor.txt_ddays3.Text = vn!ddays3
vendor.txt_idays1.Text = vn!idays1
vendor.txt_idays2.Text = vn!idays2
vendor.txt_iper2.Text = vn!iper2
vendor.txt_idays3.Text = vn!idays3
vendor.txt_iper3.Text = vn!iper3
End If
'---------------------------------------
'---------------------------------------

Dim ib As Integer
ia = 0
ib = 0

If vn.State Then vn.Close
vn.Open "select DISTINCT(materialcode),material,expirydate from vendormaterial where vname='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
While Not vn.EOF
vendor.List1.AddItem vn(0) & "  -  " & vn(1) & "  -  " & Format(vn(2), "dd/MM/yyyy")

'-----------------------------
Dim vnc As New ADODB.Recordset
If vnc.State Then vnc.Close
vnc.Open "select DISTINCT(subcategorycode),subcategory from material where code='" & vn(0) & "'", Cn, 3, 2
While Not vnc.EOF
If ib = 0 Then
vendor.List3.AddItem vnc(0) & "  -  " & vnc(1)
End If
For ia = 0 To vendor.List3.ListCount - 1
If vendor.List3.List(ia) <> vnc(0) & "  -  " & vnc(1) Then
vendor.List3.AddItem vnc(0) & "  -  " & vnc(1)
End If
Next ia

 ib = 1
vnc.MoveNext
Wend

'----------------------------
vn.MoveNext
Wend

ia = 0
For ia = 0 To vendor.List1.ListCount - 1
vendor.List1.Selected(ia) = True
Next



ia = 0
For ia = 0 To vendor.List3.ListCount - 1
vendor.List3.Selected(ia) = True
Next

vendor.Show
vendor.Top = 1900
vendor.Left = 500
vendor.Height = 6135
vendor.Width = 10695


 
vprev = flex_grid.Row

fl = 0

End Sub

Private Sub Form_Load()
On Error Resume Next
Call connect
main.lbltitle.Caption = "Vendor"
Call flex_title
Call flex_data
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
fl = 0
ih = 0
End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Code"
        .ColWidth(1) = 1500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Name"
        .ColWidth(2) = 4000
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Reg No."
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
       
        .TextMatrix(0, 4) = "Address"
        .ColWidth(4) = 4000
        .ColAlignment(4) = 0
     
        .TextMatrix(0, 5) = "Phone"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Fax"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "E-mail"
        .ColWidth(7) = 2000
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "Website"
        .ColWidth(8) = 2000
        .ColAlignment(8) = 0
        
        .TextMatrix(0, 9) = "Cmp Status"
        .ColWidth(9) = 2000
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "BM"
        .ColWidth(10) = 2000
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "NB"
        .ColWidth(11) = 2000
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "FL"
        .ColWidth(12) = 2000
        .ColAlignment(12) = 0
        .TextMatrix(0, 13) = "Remarks"
        .ColWidth(13) = 4000
        .ColAlignment(13) = 0
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload vendor
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Caption = "New" Then
 
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload vendor
vendor.Show
vendor.Top = 1900
vendor.Left = 500
vendor.Height = 6135
vendor.Width = 10695 ' to save new record
fl = 1
ih = 1
ElseIf Button.Caption = "Save" Then
'On Error GoTo assad
'validate
 

Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from vendor", Cn, 3, 2
sv.AddNew
sv!code = vendor.txt_code.Text
sv!Name = vendor.txt_name.Text
sv!regno = vendor.txt_regno.Text

sv!Address = vendor.txt_address.Text
sv!phone = vendor.txt_phone.Text
sv!fax = vendor.txt_fax.Text
sv!Email = vendor.txt_email.Text
sv!website = vendor.txt_website.Text
sv!cmpstatus = vendor.cbo_companystatus.Text
sv!bmstatus = vendor.txt_bm.Text
sv!nbstatus = vendor.txt_nb.Text
sv!flstatus = vendor.txt_fl.Text
sv!Notes = vendor.txt_notes.Text
v = 0
For v = 0 To vendor.List1.ListCount - 1
If vendor.List1.Selected(v) = True Then
nl = Split(vendor.List1.List(v), "  -  ", Len(vendor.List1.List(v)), vbTextCompare)
Dim vm As New ADODB.Recordset
If vm.State Then vm.Close
vm.Open "select * from vendormaterial", Cn, 3, 2
vm.AddNew
vm!vname = vendor.txt_code.Text
vm!materialcode = nl(0)
vm!material = nl(1)
vm!expirydate = nl(2)
vm.Update
End If
Next


sv.Update
sv.Close
sv.Open "select * from activityresponsible", Cn, 3, 2
sv.AddNew
sv!avendor = vendor.txt_code.Text
sv!avendorname = vendor.txt_name.Text
sv!rpin = vendor.txt_rpin.Text
sv!rdesig = vendor.txt_rdesig.Text
sv!ppin = vendor.txt_ppin.Text
sv!pdesig = vendor.txt_pdesig.Text
sv!epin = vendor.txt_epin.Text
sv!edesig = vendor.txt_edesig.Text
sv.Update

sv.Close
sv.Open "select * from vendorbankdetails", Cn, 3, 2
sv.AddNew
sv!bvendor = vendor.txt_code.Text
sv!bvendorname = vendor.txt_name.Text
sv!bank1 = vendor.txt_bank1.Text
sv!account1 = vendor.txt_bankaccount1.Text
sv!branch1 = vendor.txt_branch1.Text
sv!bank2 = vendor.txt_bank2.Text
sv!account2 = vendor.txt_bankaccount2.Text
sv!branch2 = vendor.txt_branch2.Text
sv!bank3 = vendor.txt_bank3.Text
sv!account3 = vendor.txt_bankaccount3.Text
sv!branch3 = vendor.txt_branch3.Text
sv.Update

sv.Close
sv.Open "select * from vendortermsofpayment", Cn, 3, 2
sv.AddNew
sv!tvendor = vendor.txt_code.Text
sv!tvendorname = vendor.txt_name.Text
sv!ddays1 = vendor.txt_ddays1.Text
sv!dper1 = vendor.txt_dper1.Text
sv!ddays2 = vendor.txt_ddays2.Text
sv!dper2 = vendor.txt_dper2.Text
sv!ddays3 = vendor.txt_ddays3.Text
sv!idays1 = vendor.txt_idays1.Text
sv!idays2 = vendor.txt_idays2.Text
sv!iper2 = vendor.txt_iper2.Text
sv!idays3 = vendor.txt_idays3.Text
sv!iper3 = vendor.txt_iper3.Text
sv.Update

sv.Close

MsgBox "New Vendor Added Succesfully"
Unload vendor
Call flex_data
Call flex_title
Exit Sub
assad:
       
    '   MsgBox "Duplicate Entries Not Allowed"
'to modify existing vendor
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
md.Open "select * from vendor where v_id=" & id1, Cn, 3, 2
If Not md.EOF Then
md!Name = vendor.txt_name.Text
md!regno = vendor.txt_regno.Text

md!Address = vendor.txt_address.Text
md!phone = vendor.txt_phone.Text
md!fax = vendor.txt_fax.Text
md!Email = vendor.txt_email.Text
md!website = vendor.txt_website.Text
md!cmpstatus = vendor.cbo_companystatus.Text
md!bmstatus = vendor.txt_bm.Text
md!nbstatus = vendor.txt_nb.Text
md!flstatus = vendor.txt_fl.Text
md!Notes = vendor.txt_notes.Text
v = 0
Cn.Execute "delete from vendormaterial where vname='" & vendor.txt_name.Text & "'"
For v = 0 To vendor.List1.ListCount - 1
If vendor.List1.Selected(v) = True Then
nl = Split(vendor.List1.List(v), "  -  ", Len(vendor.List1.List(v)), vbTextCompare)

Dim vm1 As New ADODB.Recordset
If vm1.State Then vm1.Close
vm1.Open "select * from vendormaterial", Cn, 3, 2
vm1.AddNew
vm1!vname = vendor.txt_code.Text
vm1!materialcode = nl(0)
vm1!material = nl(1)
vm1!expirydate = nl(2)
vm1.Update
End If
Next




md.Update
md.Close

md.Open "select * from activityresponsible where avendor='" & vendor.txt_code.Text & "' ", Cn, 3, 2
If Not md.EOF Then
md!rpin = vendor.txt_rpin.Text
md!rdesig = vendor.txt_rdesig.Text
md!ppin = vendor.txt_ppin.Text
md!pdesig = vendor.txt_pdesig.Text
md!epin = vendor.txt_epin.Text
md!edesig = vendor.txt_edesig.Text
md.Update
End If
md.Close
md.Open "select * from vendorbankdetails where bvendor='" & vendor.txt_code.Text & "' ", Cn, 3, 2
If Not md.EOF Then
md!bank1 = vendor.txt_bank1.Text
md!account1 = vendor.txt_bankaccount1.Text
md!branch1 = vendor.txt_branch1.Text
md!bank2 = vendor.txt_bank2.Text
md!account2 = vendor.txt_bankaccount2.Text
md!branch2 = vendor.txt_branch2.Text
md!bank3 = vendor.txt_bank3.Text
md!account3 = vendor.txt_bankaccount3.Text
md!branch3 = vendor.txt_branch3.Text
md.Update
End If

md.Close
md.Open "select * from vendortermsofpayment where tvendor='" & vendor.txt_code.Text & "' ", Cn, 3, 2
If Not md.EOF Then
md!ddays1 = vendor.txt_ddays1.Text
md!dper1 = vendor.txt_dper1.Text
md!ddays2 = vendor.txt_ddays2.Text
md!dper2 = vendor.txt_dper2.Text
md!ddays3 = vendor.txt_ddays3.Text
md!idays1 = vendor.txt_idays1.Text
md!idays2 = vendor.txt_idays2.Text
md!iper2 = vendor.txt_iper2.Text
md!idays3 = vendor.txt_idays3.Text
md!iper3 = vendor.txt_iper3.Text
md.Update
End If
md.Close


MsgBox "Selected Vendor Modified"
End If

Unload vendor
Call flex_data
Call flex_title
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
Cn.Execute "delete from vendor where v_id=" & id2
Cn.Execute "delete from vendormaterial where vname='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'"
MsgBox "Selected Vendor Has Been Deleted"
Unload vendor
Call flex_data
Call flex_title
Else
Unload vendor
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload vendor
End If




End Sub

Public Sub flex_data()
'On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from vendor order by name", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata(1)
        .TextMatrix(.Rows - 1, 2) = fldata(2)
        .TextMatrix(.Rows - 1, 3) = fldata(3)
        .TextMatrix(.Rows - 1, 4) = fldata(4)
        .TextMatrix(.Rows - 1, 5) = fldata(5)
        .TextMatrix(.Rows - 1, 6) = fldata(6)
        .TextMatrix(.Rows - 1, 7) = fldata(7)
        .TextMatrix(.Rows - 1, 8) = fldata(8)
        .TextMatrix(.Rows - 1, 9) = fldata(9)
        
        .TextMatrix(.Rows - 1, 10) = fldata(10)
        .TextMatrix(.Rows - 1, 11) = fldata(11)
        .TextMatrix(.Rows - 1, 12) = fldata(12)
        .TextMatrix(.Rows - 1, 13) = fldata(13)
        fldata.MoveNext
    Wend
End With
End Sub





