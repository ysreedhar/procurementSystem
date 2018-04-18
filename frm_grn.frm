VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_grn 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   600
      Width           =   5055
      Begin VB.OptionButton opt_a 
         BackColor       =   &H00FF8080&
         Caption         =   "Received"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3480
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_na 
         BackColor       =   &H00FF8080&
         Caption         =   "Not Received"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton opt_all 
         BackColor       =   &H00FF8080&
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
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
      TabPicture(0)   =   "frm_grn.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "   Vendor   "
      TabPicture(1)   =   "frm_grn.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "fr_requestor"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "   Requestor  "
      TabPicture(2)   =   "frm_grn.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fr_job"
      Tab(2).Control(1)=   "Label6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "       Job        "
      TabPicture(3)   =   "frm_grn.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fr_date"
      Tab(3).Control(1)=   "Label7"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   14
         Top             =   300
         Width           =   9495
         Begin MSComCtl2.DTPicker dtps_date 
            Height          =   375
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "MMM/yyyy"
            Format          =   67043331
            CurrentDate     =   38558
         End
      End
      Begin VB.Frame fr_requestor 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -75000
         TabIndex        =   11
         Top             =   300
         Width           =   12375
         Begin VB.ComboBox cbos_vendor 
            Height          =   315
            Left            =   0
            TabIndex        =   13
            Top             =   120
            Width           =   5775
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame fr_job 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -75000
         TabIndex        =   8
         Top             =   300
         Width           =   12375
         Begin VB.ComboBox cbos_requestor 
            Height          =   315
            Left            =   0
            TabIndex        =   10
            Top             =   120
            Width           =   5775
         End
         Begin VB.CommandButton cmd_apply 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame fr_date 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   750
         Left            =   -75000
         TabIndex        =   5
         Top             =   300
         Width           =   12375
         Begin VB.ComboBox cbos_job 
            Height          =   315
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   5775
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FFC0C0&
            Caption         =   "OK"
            Height          =   375
            Left            =   6000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Pending   "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   19
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Pending   "
         Height          =   255
         Left            =   -69000
         TabIndex        =   18
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Pending   "
         Height          =   255
         Left            =   -69000
         TabIndex        =   17
         Top             =   0
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Pending   "
         Height          =   255
         Left            =   -69000
         TabIndex        =   16
         Top             =   0
         Width           =   915
      End
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   0
      Top             =   120
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
            Picture         =   "frm_grn.frx":0070
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":0182
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":05D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":0E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":12CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":7564
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":787E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":7B98
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":8132
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":86CC
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":8C66
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9200
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9312
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9854
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9DEE
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A388
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":AC62
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":AD74
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":AE86
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":AF98
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":B0AA
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":B1BC
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":B2CE
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":B868
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":BE02
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":C39C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":C936
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CA48
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CB5A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D0F4
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D206
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D318
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D8B2
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D9C4
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":DF5E
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":E4F8
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":E60A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":EBA4
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":F13E
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":F6D8
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":F7EA
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":FD84
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":FE96
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":FFA8
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":100BA
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":101CC
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":102DE
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":10878
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":1098A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":10A9C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":11036
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":115D0
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":11B6A
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":12104
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":1269E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":12C38
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":131D2
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":132E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":1D2A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":2E6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3DF23
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3E377
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3E7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3EBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3F0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3F597
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3FAF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":3FFC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":4057B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":409C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":40EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":54E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":68E73
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":6C877
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":7102D
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":7156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":78E3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   5535
      Left            =   0
      TabIndex        =   20
      Top             =   1320
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   15
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
      TabIndex        =   21
      Top             =   7080
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   20
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
   Begin MSComctlLib.ImageList ImageList1 
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
            Picture         =   "frm_grn.frx":8F46D
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":8F57F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":8F9D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":8FE23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":90275
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":906C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":96961
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":96C7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":96F95
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9752F
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":97AC9
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":98063
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":985FD
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9870F
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":98C51
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":991EB
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":99785
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A05F
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A171
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A283
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A395
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A4A7
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A5B9
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9A6CB
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9AC65
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9B1FF
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9B799
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9BD33
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9BE45
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9BF57
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9C4F1
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9C603
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9C715
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9CCAF
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9CDC1
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9D35B
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9D8F5
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9DA07
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9DFA1
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9E53B
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9EAD5
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9EBE7
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F181
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F293
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F3A5
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F4B7
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F5C9
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9F6DB
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9FC75
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9FD87
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":9FE99
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A0433
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A09CD
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A0F67
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A1501
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A1A9B
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A2035
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A25CF
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":A26E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":AC6A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":BDABB
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CD320
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CD774
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CDBC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CDFF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CE4CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CE994
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CEEED
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CF3C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CF978
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":CFDC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":D02E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":E4211
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":F8270
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":FBC74
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":10042A
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":10096B
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_grn.frx":108237
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12045
      _ExtentX        =   21246
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
         TabIndex        =   26
         Top             =   0
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "Line Item Details"
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
      TabIndex        =   24
      Top             =   6840
      Width           =   12015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Confirmed   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   23
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Approved and Not Confirmed  or  Partially Approved   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   360
      Width           =   4080
   End
End
Attribute VB_Name = "frm_grn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents oSMTP As OSSMTP.SMTPSession
Attribute oSMTP.VB_VarHelpID = -1
Public Vmsr As String
Public StrCode As String
Dim Vgdid(50) As Integer
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
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221)
Next
flex_grid.Col = 1

End If

Unload GRN
Unload vscrollGRN
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

GRN.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
GRN.dtp_grn.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
GRN.cbo_po.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
GRN.cbo_vendor.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
GRN.cbo_worklocation.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
GRN.cbo_storagelocation.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
'------------------------------------------------

Vgdid(50) = 0
Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from grndetails where grno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
id = 0
For id = 0 To prs.RecordCount - 1
        vscrollGRN.cbo_category(id).Text = prs!itemid & "  -  " & prs!mrefcode & "  -  " & prs!material
        vscrollGRN.cbo_batchno(id).Text = prs!batchno
        vscrollGRN.cbo_uom(id).Text = prs!uom
        vscrollGRN.txt_qty(id).Text = prs!qty
        vscrollGRN.txt_qtyrec(id).Text = prs!qtyrec
        vscrollGRN.txt_qtyrej(id).Text = prs!qtyrej
        vscrollGRN.cbo_rejection(id).Text = prs!rejdesc
        
        If prs!chkitem = "Yes" Then
        vscrollGRN.chk_item(id).Value = 1
        Else
        vscrollGRN.chk_item(id).Value = 0
        End If
        If prs!chkqty = "Yes" Then
        vscrollGRN.chk_qty(id).Value = 1
        Else
        vscrollGRN.chk_qty(id).Value = 0
        End If
        
        vscrollGRN.cbo_category(id).Enabled = False
        vscrollGRN.cbo_uom(id).Enabled = False
        vscrollGRN.txt_qty(id).Enabled = False
        
        vscrollGRN.txt_qtyrec(id).Enabled = False
        vscrollGRN.txt_qtyrej(id).Enabled = False
        vscrollGRN.cbo_rejection(id).Enabled = False
                
        Vgdid(id) = prs!grn_id
        prs.MoveNext
Next id
prs.Close
'--------------------------------------


GRN.Show
SetParent GRN.hwnd, frm_grn.hwnd
GRN.Height = 7155
GRN.Width = 11895
GRN.Top = 50
GRN.Left = 200

vprev = flex_grid.Row

End Sub
Private Sub flex_grid_SelChange()
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
prs.Open "select * from grndetails where grno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
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
        .TextMatrix(.Rows - 1, 6) = prs!qty
        .TextMatrix(.Rows - 1, 7) = prs!qtyrec
        .TextMatrix(.Rows - 1, 8) = prs!qtyrej
        .TextMatrix(.Rows - 1, 9) = prs!rejdesc

        prs.MoveNext
    Wend
End With
'''''''''''''''''''''''''''''''''''''''''''''''
vprev = flex_grid.Row
End Sub

Private Sub flex_item_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = True

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

If flex_grid.Row <> 0 Then
'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221)
Next
flex_grid.Col = 1
End If


Unload GRN
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


GRN.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
GRN.dtp_grn.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
GRN.cbo_po.Text = flex_grid.TextMatrix(flex_grid.Row, 3)
GRN.cbo_vendor.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
GRN.cbo_worklocation.Text = flex_grid.TextMatrix(flex_grid.Row, 5)
GRN.cbo_storagelocation.Text = flex_grid.TextMatrix(flex_grid.Row, 6)

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from grndetails where gd_id='" & flex_item.TextMatrix(flex_item.Row, 0) & "'", Cn, 3, 2
id = 0
For id = 0 To prs.RecordCount - 1
        vscrollGRN.cbo_category(id).Text = prs!itemid & "  -  " & prs!mrefcode & "  -  " & prs!material
        vscrollGRN.cbo_batchno(id).Text = prs!batchno
        vscrollGRN.cbo_uom(id).Text = prs!uom
        vscrollGRN.txt_qty(id).Text = prs!qty
        vscrollGRN.txt_qtyrec(id).Text = prs!qtyrec
        vscrollGRN.txt_qtyrej(id).Text = prs!qtyrej
        vscrollGRN.cbo_rejection(id).Text = prs!rejdesc
        If prs!chkitem = "Yes" Then
        vscrollGRN.chk_item(id).Value = 1
        Else
        vscrollGRN.chk_item(id).Value = 0
        End If
        If prs!chkqty = "Yes" Then
        vscrollGRN.chk_qty(id).Value = 1
        Else
        vscrollGRN.chk_qty(id).Value = 0
        End If
        
        vscrollGRN.cbo_category(id).Enabled = False
        vscrollGRN.cbo_uom(id).Enabled = False
        vscrollGRN.txt_qty(id).Enabled = False
        vscrollGRN.txt_qtyrec(id).Enabled = False
        vscrollGRN.txt_qtyrej(id).Enabled = False
        vscrollGRN.cbo_rejection(id).Enabled = False
        Vgdid(id) = prs!grn_id
        prs.MoveNext
Next id
prs.Close


GRN.Show
SetParent GRN.hwnd, frm_grn.hwnd
GRN.Height = 7155
GRN.Width = 11895
GRN.Top = 50
GRN.Left = 200

'GRN.kl = 0
vprev = flex_grid.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
dtps_date.Value = Format(Date, "dd/MM/yyyy")
Call connect
main.lbltitle.Caption = "Goods Received Note"

Call flex_title
Call flex_subtitle
Toolbar1.Buttons(3).Enabled = True
Toolbar1.Buttons(5).Enabled = True

Me.Top = 5
Me.Left = 5
Call flex_datadate


End Sub
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "GRN No."
        .ColWidth(1) = 1700
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Date"
        .ColWidth(2) = 1000
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "PO No."
        .ColWidth(3) = 2000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Vendor"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0

        .TextMatrix(0, 5) = "Work Loc."
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Stor. Loc"
        .ColWidth(6) = 2000
        .ColAlignment(6) = 0
        
        
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload GRN
Unload vscrollGRN
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
Unload GRN
GRN.Show
SetParent GRN.hwnd, frm_grn.hwnd
GRN.Height = 7155
GRN.Width = 11895
GRN.Top = 50
GRN.Left = 200
ElseIf Button.Caption = "Save" Then
On Error GoTo assad
'validate

'-----------
Dim sv As New ADODB.Recordset
If sv.State Then sv.Close
sv.Open "select * from grn ", Cn, 3, 2
sv.AddNew

sv!grno = GRN.txt_account.Text
sv!grndate = GRN.dtp_grn.Value
sv!pono = GRN.cbo_po.Text
sv!vendor = GRN.cbo_vendor.Text
sv!worklocation = GRN.cbo_worklocation.Text
sv!storagelocation = GRN.cbo_storagelocation.Text

'--------------
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from grndetails", Cn, 3, 2

Dim j As Integer

    j = 0
        
For j = 0 To vscrollGRN.cbo_category.Count - 1
If vscrollGRN.cbo_category(j).Text <> "" Then
        
        pr.AddNew
        spt = Split(vscrollGRN.cbo_category(j).Text, "  -  ", Len(vscrollGRN.cbo_category(j).Text), vbTextCompare)
        '-----------------------------
        
If GRN.cbo_lookup.Text = "Item ID" Then
      pr!itemid = spt(0)
      pr!mrefcode = spt(1)
      pr!material = spt(2) & "  -  " & spt(3)
ElseIf GRN.cbo_lookup.Text = "Mfr PartNo." Then
      pr!itemid = spt(1)
      pr!mrefcode = spt(0)
      pr!material = spt(2) & "  -  " & spt(3)

ElseIf GRN.cbo_lookup.Text = "Item Description" Then
      pr!itemid = spt(2)
      pr!mrefcode = spt(3)
      pr!material = spt(0) & "  -  " & spt(1)

ElseIf GRN.cbo_lookup.Text = "Search" Then
      pr!itemid = spt(2)
      pr!mrefcode = spt(3)
      pr!material = spt(0) & "  -  " & spt(1)
      Else
      pr!itemid = spt(0)
      pr!mrefcode = spt(1)
      pr!material = spt(2) & "  -  " & spt(3)
End If
        '-----------------------------
       If vscrollGRN.chk_item(j).Value = 1 Then
       pr!chkitem = "Yes"
       End If
       If vscrollGRN.chk_qty(j).Value = 1 Then
       pr!chkqty = "Yes"
       End If
             
      pr!batchno = vscrollGRN.cbo_batchno(j).Text
      pr!uom = vscrollGRN.cbo_uom(j).Text
      pr!qty = vscrollGRN.txt_qty(j).Text
      pr!qtyrec = vscrollGRN.txt_qtyrec(j).Text
      pr!qtyrej = vscrollGRN.txt_qtyrej(j).Text
      pr!rejdesc = vscrollGRN.cbo_rejection(j).Text
      
      pr!grno = GRN.txt_account.Text
      'pr!pono = GRN.cbo_po.Text
      
      pr.Update
        
End If
      
Next j
pr.Close
sv.Update
sv.Close
'End If
MsgBox "GRN Recorded Successfully"
Unload GRN
Unload vscrollGRN
Call flex_datadate
Call flex_title
'Call flex_itemmodi
Exit Sub
assad:



'------------



       MsgBox "Duplicate Entries Not Allowed"
'to modify existing GRN


ElseIf Button.Caption = "Close" Then
Unload Me
Unload GRN
End If
End Sub

Public Sub flex_datadate()
'On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
fldata.Open "select * from GRN where month(grndate)='" & Format(dtps_date.Value, "mm") & "' and year(grndate)='" & Format(dtps_date.Value, "yyyy") & "' ", Cn, 3, 2 'and tuser = '" & main.Label2.Caption & "' order by prno,prdate

With flex_grid
    .Rows = 1
    While Not fldata.EOF
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = fldata(0)
        .TextMatrix(.Rows - 1, 1) = fldata!grno
        .TextMatrix(.Rows - 1, 2) = fldata!grndate
        .TextMatrix(.Rows - 1, 3) = fldata!pono
        .TextMatrix(.Rows - 1, 4) = fldata!vendor
        .TextMatrix(.Rows - 1, 5) = fldata!worklocation
        .TextMatrix(.Rows - 1, 6) = fldata!storagelocation
        
        
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
        .ColWidth(6) = 800
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Rec Qty"
        .ColWidth(7) = 800
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "Rej Qty"
        .ColWidth(8) = 800
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "Rej Desc"
        .ColWidth(9) = 2000
        .ColAlignment(9) = 0
        
        
    End With
End Sub
Public Sub flex_itemmodi()
On Error Resume Next
current = flex_grid.Row

'Reset to previous row
'''''''''''''''''''''''''''''''''''''''''''''''''


Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from grndetails where grno='" & flex_grid.TextMatrix(flex_grid.Row, 1) & "'", Cn, 3, 2
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
        .TextMatrix(.Rows - 1, 6) = prs!qty
        .TextMatrix(.Rows - 1, 7) = prs!qtyrec
        .TextMatrix(.Rows - 1, 8) = prs!qtyrej
        .TextMatrix(.Rows - 1, 9) = prs!rejdesc
                
        prs.MoveNext
    Wend
End With
prs.Close
'''''''''''''''''''''''''''''''''''''''''''''''

vprev = flex_grid.Row
End Sub



