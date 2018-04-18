VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_purchaserequisition 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "View MSR"
      Height          =   375
      Left            =   9960
      TabIndex        =   17
      Top             =   480
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   975
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1720
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "   Month/Year   "
      TabPicture(0)   =   "frm_purchaserequisition.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fr_requestor"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "   Vendor   "
      TabPicture(1)   =   "frm_purchaserequisition.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fr_job"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "     Requestor      "
      TabPicture(2)   =   "frm_purchaserequisition.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fr_date"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "      Job      "
      TabPicture(3)   =   "frm_purchaserequisition.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame1"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -75000
         TabIndex        =   9
         Top             =   300
         Width           =   11895
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox cbos_job 
            Height          =   315
            Left            =   0
            TabIndex        =   15
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame fr_requestor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   8
         Top             =   300
         Width           =   11895
         Begin MSComCtl2.DTPicker dtps_date 
            Height          =   375
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMM/yyyy"
            Format          =   16449539
            CurrentDate     =   38558
         End
      End
      Begin VB.Frame fr_job 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -75000
         TabIndex        =   7
         Top             =   300
         Width           =   11895
         Begin VB.CommandButton cmd_apply 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox cbos_vendor 
            Height          =   315
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame fr_date 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -75000
         TabIndex        =   6
         Top             =   300
         Width           =   11895
         Begin VB.ComboBox cbos_requestor 
            Height          =   315
            Left            =   0
            TabIndex        =   14
            Top             =   240
            Width           =   5775
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
      End
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
            Picture         =   "frm_purchaserequisition.frx":0070
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":0182
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":05D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":0E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":12CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":7564
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":787E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":7B98
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":8132
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":86CC
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":8C66
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":9200
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":9312
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":9854
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":9DEE
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":A388
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":AC62
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":AD74
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":AE86
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":AF98
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":B0AA
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":B1BC
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":B2CE
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":B868
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":BE02
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":C39C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":C936
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":CA48
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":CB5A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":D0F4
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":D206
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":D318
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":D8B2
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":D9C4
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":DF5E
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":E4F8
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":E60A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":EBA4
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":F13E
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":F6D8
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":F7EA
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":FD84
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":FE96
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":FFA8
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":100BA
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":101CC
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":102DE
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":10878
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":1098A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":10A9C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":11036
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":115D0
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":11B6A
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":12104
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":1269E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":12C38
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":131D2
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":132E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":1D2A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":2E6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3DF23
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3E377
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3E7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3EBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3F0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3F597
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3FAF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":3FFC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":4057B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":409C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":40EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":54E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":68E73
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":6C877
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":7102D
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":7156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_purchaserequisition.frx":78E3A
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
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   635
      ButtonWidth     =   2196
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList5"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New MSR"
            Key             =   "ar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save MSR"
            Key             =   "grd"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modify MSR"
            Key             =   "hlp"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete MSR"
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
      Top             =   1320
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9446
      _Version        =   393216
      Rows            =   3
      Cols            =   10
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
      Top             =   6915
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   14
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
      TabIndex        =   4
      Top             =   6680
      Width           =   11775
   End
End
Attribute VB_Name = "frm_purchaserequisition"
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
Call striptab
Call flex_itemmodi
End Sub
Private Sub Command1_Click()
Call striptab
Call flex_itemmodi
End Sub
Private Sub Command2_Click()
Call striptab
Call flex_itemmodi
End Sub

Private Sub Command3_Click()
Reportmsr.msrno = flex_grid.TextMatrix(flex_grid.Row, 1)
Reportmsr.loadmsrdetails
Unload Reportmsr
End Sub

Private Sub dtps_date_Change()
Call striptab
Call flex_itemmodi
End Sub
Private Sub dtps_date_Click()
Call striptab
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

Unload purchaserequisition
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


purchaserequisition.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
purchaserequisition.dtp_pr.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
purchaserequisition.cbo_project.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
purchaserequisition.txt_justification.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
purchaserequisition.cbo_recommendedvendor.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
purchaserequisition.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
purchaserequisition.cbo_expensetype.Text = flex_grid.TextMatrix(flex_grid.Row, 9)
'-----------------------------------------------

Dim msrapp As New ADODB.Recordset
If msrapp.State Then msrapp.Close
msrapp.Open "select * from purchaserequisition where status='Approved' and prno='" & purchaserequisition.txt_account.Text & "'", Cn, 3, 2
If Not msrapp.EOF Then
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
End If
'------------------------------------------------

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from prdetails where prno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With purchaserequisition.flex_med
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!qty
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!reqdate
        .TextMatrix(.Rows - 1, 10) = prs!remarks
        .TextMatrix(.Rows - 1, 7) = prs!jobcharge
        .TextMatrix(.Rows - 1, 8) = prs!Location
        .TextMatrix(.Rows - 1, 9) = prs!contactperson
        .TextMatrix(.Rows - 1, 11) = prs!mtype
        .TextMatrix(.Rows - 1, 12) = prs!fromdt
        .TextMatrix(.Rows - 1, 13) = prs!todt

        prs.MoveNext
    Wend
End With
prs.Close
'--------------------------------------
purchaserequisition.listURL.Clear
prs.Open "select fpath from fileattach where fprno ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
While Not prs.EOF
purchaserequisition.listURL.AddItem prs(0)
prs.MoveNext
Wend
prs.Close

purchaserequisition.Show
SetParent purchaserequisition.hwnd, frm_purchaserequisition.hwnd
purchaserequisition.Height = 8760
purchaserequisition.Width = 11640
purchaserequisition.Top = 50
purchaserequisition.Left = 200

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
prs.Open "select * from prdetails where prno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With flex_item
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!qty
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!reqdate
        .TextMatrix(.Rows - 1, 10) = prs!remarks
        .TextMatrix(.Rows - 1, 7) = prs!jobcharge
        .TextMatrix(.Rows - 1, 8) = prs!Location
        .TextMatrix(.Rows - 1, 9) = prs!contactperson
        .TextMatrix(.Rows - 1, 11) = prs!mtype
        .TextMatrix(.Rows - 1, 12) = prs!fromdt
        .TextMatrix(.Rows - 1, 13) = prs!todt

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


Unload purchaserequisition
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


purchaserequisition.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
purchaserequisition.dtp_pr.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
purchaserequisition.cbo_project.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
purchaserequisition.txt_justification.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
purchaserequisition.cbo_recommendedvendor.Text = flex_grid.TextMatrix(flex_grid.Row, 7)
purchaserequisition.txt_notes.Text = flex_grid.TextMatrix(flex_grid.Row, 8)
purchaserequisition.cbo_expensetype.Text = flex_grid.TextMatrix(flex_grid.Row, 9)


Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from prdetails where prno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With purchaserequisition.flex_med
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!qty
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!reqdate
        .TextMatrix(.Rows - 1, 10) = prs!remarks
        .TextMatrix(.Rows - 1, 7) = prs!jobcharge
        .TextMatrix(.Rows - 1, 8) = prs!Location
        .TextMatrix(.Rows - 1, 9) = prs!contactperson
        .TextMatrix(.Rows - 1, 11) = prs!mtype
        .TextMatrix(.Rows - 1, 12) = prs!fromdt
        .TextMatrix(.Rows - 1, 13) = prs!todt
        prs.MoveNext
    Wend
End With
prs.Close
purchaserequisition.listURL.Clear
prs.Open "select fpath from fileattach where fprno ='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
While Not prs.EOF
purchaserequisition.listURL.AddItem prs(0)
prs.MoveNext
Wend
prs.Close

purchaserequisition.lblid = flex_item.TextMatrix(flex_item.Row, 0)
Dim prd As New ADODB.Recordset
If prd.State Then prd.Close
prd.Open "select * from prdetails where pr_id='" & flex_item.TextMatrix(flex_item.Row, 0) & "'", Cn, 3, 2
If Not prd.EOF Then
Dim h As Integer
h = 1
For h = 1 To purchaserequisition.flex_med.Rows - 1
If purchaserequisition.flex_med.TextMatrix(h, 0) = prd!pr_id Then
purchaserequisition.flex_med.Row = h
purchaserequisition.cbo_category = purchaserequisition.flex_med.TextMatrix(h, 1) & "  -  " & purchaserequisition.flex_med.TextMatrix(h, 2) & "  -  " & purchaserequisition.flex_med.TextMatrix(h, 3)
purchaserequisition.txt_qty = purchaserequisition.flex_med.TextMatrix(h, 4)
purchaserequisition.cbo_uom = purchaserequisition.flex_med.TextMatrix(h, 5)
purchaserequisition.dtp_reqd = purchaserequisition.flex_med.TextMatrix(h, 6)
purchaserequisition.txt_remarks = purchaserequisition.flex_med.TextMatrix(h, 10)
purchaserequisition.cbo_jobcharge.Text = purchaserequisition.flex_med.TextMatrix(h, 7)
purchaserequisition.cbo_worklocation.Text = purchaserequisition.flex_med.TextMatrix(h, 8)
purchaserequisition.cbo_location.Text = purchaserequisition.flex_med.TextMatrix(h, 9)

purchaserequisition.cbo_materialtype.Text = purchaserequisition.flex_med.TextMatrix(h, 11)
purchaserequisition.dtp_from.Value = purchaserequisition.flex_med.TextMatrix(h, 12)
purchaserequisition.dtp_to.Value = purchaserequisition.flex_med.TextMatrix(h, 13)

End If
Next
End If
purchaserequisition.Show
SetParent purchaserequisition.hwnd, frm_purchaserequisition.hwnd
purchaserequisition.Height = 8760
purchaserequisition.Width = 11640
purchaserequisition.Top = 50
purchaserequisition.Left = 200

purchaserequisition.kl = 0
vprev = flex_grid.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
dtps_date.Value = Format(Date, "dd/MM/yyyy")
Call connect
main.lbltitle.Caption = "Material Service Requisition"

Call flex_title
Call flex_subtitle
Toolbar1.Buttons(3).Enabled = False
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Me.Top = 5
Me.Left = 5
Call striptab


cbos_vendor.Clear
Dim rvn As New ADODB.Recordset
If rvn.State Then rvn.Close
rvn.Open "select DISTINCT(name) from vendor order by name", Cn, 3, 2
While Not rvn.EOF
cbos_vendor.AddItem rvn(0)
rvn.MoveNext
Wend
rvn.Close

  
cbos_requestor.Clear
Dim rrq As New ADODB.Recordset
If rrq.State Then rrq.Close
rrq.Open "select DISTINCT(requestor) from purchaserequisition order by requestor", Cn, 3, 2
While Not rrq.EOF
cbos_requestor.AddItem rrq(0)
rrq.MoveNext
Wend
rrq.Close


  
cbos_job.Clear
Dim jb As New ADODB.Recordset
If jb.State Then jb.Close
jb.Open "select DISTINCT(jobcharge) from prdetails order by jobcharge", Cn, 3, 2
While Not jb.EOF
cbos_job.AddItem jb(0)
jb.MoveNext
Wend
jb.Close
End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "MSR No."
        .ColWidth(1) = 1700
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Date"
        .ColWidth(2) = 1000
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Department"
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Requestor"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0

        .TextMatrix(0, 5) = "Project"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0

        
        .TextMatrix(0, 6) = "Justification"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Recommended Vendor"
        .ColWidth(7) = 2000
        .ColAlignment(7) = 0

        .TextMatrix(0, 8) = "Remarks"
        .ColWidth(8) = 4000
        .ColAlignment(8) = 0
        
        .TextMatrix(0, 9) = "Expense Type"
        .ColWidth(9) = 2000
        .ColAlignment(9) = 0
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload purchaserequisition
End Sub
Private Sub TabStrip1_Click()
Call striptab
End Sub

Private Sub Label2_Click()
Call striptab
Call flex_itemmodi
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Call striptab
Call flex_itemmodi
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

If Button.Caption = "New MSR" Then
purchaserequisition.Command1.Enabled = True
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Unload purchaserequisition
purchaserequisition.Show
SetParent purchaserequisition.hwnd, frm_purchaserequisition.hwnd
purchaserequisition.Height = 8760
purchaserequisition.Width = 11640
purchaserequisition.Top = 50
purchaserequisition.Left = 200
' to save new record
ElseIf Button.Caption = "Save MSR" Then
On Error GoTo assad
'validate
If purchaserequisition.cbo_project.Text = "" Then
MsgBox "Select Project"
purchaserequisition.cbo_project.SetFocus
Exit Sub
End If

If purchaserequisition.cbo_expensetype.Text = "" Then
MsgBox "Select ExpenseType"
purchaserequisition.cbo_expensetype.SetFocus
Exit Sub
End If

If purchaserequisition.flex_med.Rows <= 1 Then
MsgBox "Enter Items Required"
purchaserequisition.cbo_materialtype.SetFocus
Exit Sub
End If
'-----------


spl = Split(purchaserequisition.txt_dept.Text, "\", Len(purchaserequisition.txt_dept.Text), vbTextCompare)
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from purchaserequisition", Cn, 3, 2
sv.AddNew
sv!prno = purchaserequisition.txt_account.Text
sv!prdate = purchaserequisition.dtp_pr.Value
sv!department = spl(0)
sv!requestor = spl(1)
sv!project = purchaserequisition.cbo_project.Text
sv!justification = purchaserequisition.txt_justification.Text
sv!recommendedvendor = purchaserequisition.cbo_recommendedvendor.Text
sv!Notes = purchaserequisition.txt_notes.Text
sv!tdate = Now
sv!tuser = main.Label2.Caption
sv!expensetype = purchaserequisition.cbo_expensetype.Text
Vmsr = purchaserequisition.txt_account.Text
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from prdetails", Cn, 3, 2

Dim j As Integer

    j = 1
        For j = 1 To purchaserequisition.flex_med.Rows - 1

        pr.AddNew
        pr!prno = purchaserequisition.txt_account.Text
        sp1 = Split(purchaserequisition.flex_med.TextMatrix(j, 1), "   -   ", Len(purchaserequisition.flex_med.TextMatrix(j, 1)), vbTextCompare)
        If purchaserequisition.flex_med.TextMatrix(j, 1) = "" Then
        pr!itemid = ""
        Else
        pr!itemid = sp1(0)
        End If
        If purchaserequisition.flex_med.TextMatrix(j, 2) = "" Then
        pr!mrefcode = ""
        Else
        sp2 = Split(purchaserequisition.flex_med.TextMatrix(j, 2), "   -   ", Len(purchaserequisition.flex_med.TextMatrix(j, 2)), vbTextCompare)
        pr!mrefcode = sp2(0)
        End If
        pr!material = purchaserequisition.flex_med.TextMatrix(j, 3)
        pr!qty = purchaserequisition.flex_med.TextMatrix(j, 4)
        pr!uom = purchaserequisition.flex_med.TextMatrix(j, 5)
        pr!reqdate = purchaserequisition.flex_med.TextMatrix(j, 6)
        pr!remarks = purchaserequisition.flex_med.TextMatrix(j, 10)
        pr!jobcharge = purchaserequisition.flex_med.TextMatrix(j, 7)
        pr!Location = purchaserequisition.flex_med.TextMatrix(j, 8)
        pr!contactperson = purchaserequisition.flex_med.TextMatrix(j, 9)
        pr!mtype = purchaserequisition.flex_med.TextMatrix(j, 11)
        pr!fromdt = purchaserequisition.flex_med.TextMatrix(j, 12)
        pr!todt = purchaserequisition.flex_med.TextMatrix(j, 13)
        pr.Update
        pr.MoveNext
        Next j
pr.Close
sv.Update
sv.Close

Unload purchaserequisition
Call flex_itemmodi
Call striptab
Call flex_title
ms = MsgBox("Do you want to forwed MSR for Approval", vbYesNo)
If ms = vbYes Then

sv.Open "select DISTINCT(rs.rs_code) from releasedetails rd, releasestrategy rs where rs.rs_code=rd.rs_code and rs.flg='Yes' and rs.prpo='MSR' ", Cn, 3, 2
If Not sv.EOF Then
   Dim rls As New ADODB.Recordset
   If rls.State Then rls.Close
   rls.Open "select * from prauth", Cn, 3, 2
   rls.AddNew
   rls!at_msrno = Vmsr
   rls!at_user = sv(0)
   StrCode = sv(0)
   rls.Update
   rls.Close
   
End If
 If main.VstrEmail = 0 Then
 MsgBox "New MSR Added Succesfully, System is sending mail to the Authorized Members for Approval: Kindly wait for confirmation message"
 Call assademailservice
 Else
 MsgBox "New MSR added successfully"
 End If
End If
Exit Sub
assad:

       MsgBox "Duplicate Entries Not Allowed"
'to modify existing purchaserequisition
ElseIf Button.Caption = "Modify MSR" Then
On Error GoTo assad1
'validate
If purchaserequisition.cbo_project.Text = "" Then
MsgBox "Select Project"
purchaserequisition.cbo_project.SetFocus
Exit Sub
End If

If purchaserequisition.cbo_expensetype.Text = "" Then
MsgBox "Select ExpenseType"
purchaserequisition.cbo_expensetype.SetFocus
Exit Sub
End If

If purchaserequisition.flex_med.Rows <= 1 Then
MsgBox "Enter Items Required"
purchaserequisition.cbo_materialtype.SetFocus
Exit Sub
End If


spl1 = Split(purchaserequisition.txt_dept.Text, "\", Len(purchaserequisition.txt_dept.Text), vbTextCompare)

Cn.Execute "delete from purchaserequisition where prno='" & purchaserequisition.txt_account.Text & "'"
Cn.Execute "delete from prdetails where prno='" & purchaserequisition.txt_account.Text & "'"
Dim svv As New ADODB.Recordset
If svv.State Then svv.Close
svv.Open "select * from purchaserequisition", Cn, 3, 2
svv.AddNew
svv!prno = purchaserequisition.txt_account.Text
svv!prdate = purchaserequisition.dtp_pr.Value
svv!department = spl1(0)
svv!requestor = spl1(1)
svv!project = purchaserequisition.cbo_project.Text
svv!justification = purchaserequisition.txt_justification.Text
svv!recommendedvendor = purchaserequisition.cbo_recommendedvendor.Text
svv!Notes = purchaserequisition.txt_notes.Text
svv!tdate = Now
svv!tuser = main.Label2.Caption
svv!expensetype = purchaserequisition.cbo_expensetype.Text
Dim prr As New ADODB.Recordset
If prr.State Then prr.Close
prr.Open "select * from prdetails", Cn, 3, 2

Dim k As Integer

    k = 1
        For k = 1 To purchaserequisition.flex_med.Rows - 1

        prr.AddNew
        prr!prno = purchaserequisition.txt_account.Text
        sp3 = Split(purchaserequisition.flex_med.TextMatrix(k, 1), "   -   ", Len(purchaserequisition.flex_med.TextMatrix(k, 1)), vbTextCompare)
        prr!itemid = sp3(0)
        sp4 = Split(purchaserequisition.flex_med.TextMatrix(k, 2), "   -   ", Len(purchaserequisition.flex_med.TextMatrix(k, 2)), vbTextCompare)
        If purchaserequisition.flex_med.TextMatrix(k, 2) = "" Then
        prr!mrefcode = ""
        Else
        sp2 = Split(purchaserequisition.flex_med.TextMatrix(k, 2), "   -   ", Len(purchaserequisition.flex_med.TextMatrix(k, 2)), vbTextCompare)
        prr!mrefcode = sp2(0)
        End If
        prr!material = purchaserequisition.flex_med.TextMatrix(k, 3)
        prr!qty = purchaserequisition.flex_med.TextMatrix(k, 4)
        prr!uom = purchaserequisition.flex_med.TextMatrix(k, 5)
        prr!reqdate = purchaserequisition.flex_med.TextMatrix(k, 6)
        prr!remarks = purchaserequisition.flex_med.TextMatrix(k, 10)
        prr!jobcharge = purchaserequisition.flex_med.TextMatrix(k, 7)
        prr!Location = purchaserequisition.flex_med.TextMatrix(k, 8)
        prr!contactperson = purchaserequisition.flex_med.TextMatrix(k, 9)
        prr!mtype = purchaserequisition.flex_med.TextMatrix(k, 11)
        prr!fromdt = purchaserequisition.flex_med.TextMatrix(k, 12)
        prr!todt = purchaserequisition.flex_med.TextMatrix(k, 13)
        prr.Update
        prr.MoveNext
        Next k
prr.Close
   

svv.Update
svv.Close
MsgBox "Selected Record Modify Succesfully"
Unload purchaserequisition
Call flex_itemmodi
Call striptab
Call flex_title

Exit Sub
assad1:

       MsgBox "Duplicate Entries Not Allowed"
'to delete
ElseIf Button.Caption = "Delete MSR" Then




Toolbar1.Buttons(3).Enabled = False
dlt = MsgBox("Do you want to Delete", vbYesNo)
If dlt = vbYes Then
Cn.Execute "delete from purchaserequisition where prno='" & purchaserequisition.txt_account.Text & "'"
Cn.Execute "delete from prdetails where prno='" & purchaserequisition.txt_account.Text & "'"
MsgBox "Selected Record Has Been Deleted"
Unload purchaserequisition
Call striptab
Call flex_title
Else
Unload purchaserequisition
End If
ElseIf Button.Caption = "Close" Then
Unload Me
Unload purchaserequisition
End If
End Sub

Public Sub flex_datadate()
'On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from purchaserequisition where month(prdate)='" & Format(dtps_date.Value, "mm") & "' and year(prdate)='" & Format(dtps_date.Value, "yyyy") & "' and tuser = '" & main.Label2.Caption & "' order by prno,prdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!prno
        .TextMatrix(.Rows - 1, 2) = fldata!prdate
        .TextMatrix(.Rows - 1, 3) = fldata!department
        .TextMatrix(.Rows - 1, 4) = fldata!requestor
        .TextMatrix(.Rows - 1, 5) = fldata!project
        .TextMatrix(.Rows - 1, 6) = fldata!justification
        .TextMatrix(.Rows - 1, 7) = fldata!recommendedvendor
        .TextMatrix(.Rows - 1, 8) = fldata!Notes
        .TextMatrix(.Rows - 1, 9) = fldata!expensetype
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
        .TextMatrix(0, 4) = "Qty"
        .ColWidth(4) = 800
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "UOM"
        .ColWidth(5) = 800
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "ReqDate"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        .TextMatrix(0, 10) = "Remarks"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        .TextMatrix(0, 7) = "JobCharge"
        .ColWidth(7) = 1200
        .ColAlignment(7) = 0

        .TextMatrix(0, 8) = "Work Location"
        .ColWidth(8) = 1200
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "Stor Loc."
        .ColWidth(9) = 1200
        .ColAlignment(9) = 0
        
        .TextMatrix(0, 11) = "Mat.Type"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "From"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0
        .TextMatrix(0, 13) = "To"
        .ColWidth(13) = 1200
        .ColAlignment(13) = 0


    End With
End Sub
Public Sub flex_itemmodi()
current = flex_grid.Row

'Reset to previous row
'''''''''''''''''''''''''''''''''''''''''''''''''


Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from prdetails where prno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
With flex_item
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = prs(0)
        .TextMatrix(.Rows - 1, 1) = prs!itemid
        .TextMatrix(.Rows - 1, 2) = prs!mrefcode
        .TextMatrix(.Rows - 1, 3) = prs!material
        .TextMatrix(.Rows - 1, 4) = prs!qty
        .TextMatrix(.Rows - 1, 5) = prs!uom
        .TextMatrix(.Rows - 1, 6) = prs!reqdate
        .TextMatrix(.Rows - 1, 10) = prs!remarks
        .TextMatrix(.Rows - 1, 7) = prs!jobcharge
        .TextMatrix(.Rows - 1, 8) = prs!Location
        .TextMatrix(.Rows - 1, 9) = prs!contactperson
        .TextMatrix(.Rows - 1, 11) = prs!mtype
        .TextMatrix(.Rows - 1, 12) = prs!fromdt
        .TextMatrix(.Rows - 1, 13) = prs!todt
        prs.MoveNext
    Wend
End With
prs.Close
'''''''''''''''''''''''''''''''''''''''''''''''

vprev = flex_grid.Row
End Sub

Public Sub flex_datarequestor()
If cbos_requestor.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from purchaserequisition where requestor='" & cbos_requestor.Text & "' and tuser = '" & main.Label2.Caption & "' order by prno,prdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!prno
        .TextMatrix(.Rows - 1, 2) = fldata!prdate
        .TextMatrix(.Rows - 1, 3) = fldata!department
        .TextMatrix(.Rows - 1, 4) = fldata!requestor
        .TextMatrix(.Rows - 1, 5) = fldata!project
        .TextMatrix(.Rows - 1, 6) = fldata!justification
        .TextMatrix(.Rows - 1, 7) = fldata!recommendedvendor
        .TextMatrix(.Rows - 1, 8) = fldata!Notes
        .TextMatrix(.Rows - 1, 9) = fldata!expensetype
        fldata.MoveNext
    Wend
End With
fldata.Close
End Sub

Public Sub flex_datajob()
If cbos_job.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select DISTINCT(p.pr_id),p.prno,p.prdate,p.department,p.requestor,p.project,p.justification,p.recommendedvendor,p.notes,p.expensetype from purchaserequisition p,prdetails pr where p.prno=pr.prno and pr.jobcharge='" & cbos_job.Text & "' and p.tuser = '" & main.Label2.Caption & "' order by p.prno,p.prdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key
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
        fldata.MoveNext
    Wend
End With
fldata.Close
End Sub
Public Sub flex_vendor()
If cbos_vendor.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from purchaserequisition where recommendedvendor='" & cbos_vendor.Text & "' and tuser = '" & main.Label2.Caption & "' order by prno,prdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & main.Label2.Caption & "' order by p.proj_key

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!prno
        .TextMatrix(.Rows - 1, 2) = fldata!prdate
        .TextMatrix(.Rows - 1, 3) = fldata!department
        .TextMatrix(.Rows - 1, 4) = fldata!requestor
        .TextMatrix(.Rows - 1, 5) = fldata!project
        .TextMatrix(.Rows - 1, 6) = fldata!justification
        .TextMatrix(.Rows - 1, 7) = fldata!recommendedvendor
        .TextMatrix(.Rows - 1, 8) = fldata!Notes
        .TextMatrix(.Rows - 1, 9) = fldata!expensetype
        fldata.MoveNext
    Wend
End With
fldata.Close
End Sub
Public Sub striptab()
If SSTab1.Caption = "   Vendor   " Then
Call flex_vendor
ElseIf SSTab1.Caption = "     Requestor      " Then

Call flex_datarequestor
ElseIf SSTab1.Caption = "      Job      " Then

Call flex_datajob
ElseIf SSTab1.Caption = "   Month/Year   " Then

Call flex_datadate
End If
End Sub
 
 
