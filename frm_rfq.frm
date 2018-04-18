VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rfq 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   12015
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vendor Group"
      Height          =   615
      Left            =   6960
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.OptionButton opt_a 
         BackColor       =   &H00FF8080&
         Caption         =   "Assigned"
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
         Caption         =   "Not Assigned"
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
      Top             =   0
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
      TabPicture(0)   =   "frm_rfq.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "   Vendor   "
      TabPicture(1)   =   "frm_rfq.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fr_requestor"
      Tab(1).Control(1)=   "Label5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "   Requestor  "
      TabPicture(2)   =   "frm_rfq.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fr_job"
      Tab(2).Control(1)=   "Label6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "       Job        "
      TabPicture(3)   =   "frm_rfq.frx":0054
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
            Format          =   67239939
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
         Height          =   255
         Left            =   6000
         TabIndex        =   19
         Top             =   0
         Width           =   915
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
      Top             =   -360
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
            Picture         =   "frm_rfq.frx":0070
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":0182
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":05D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":0A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":0E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":12CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":7564
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":787E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":7B98
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":8132
            Key             =   "close"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":86CC
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":8C66
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":9200
            Key             =   "properties"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":9312
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":9854
            Key             =   "pagesetup"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":9DEE
            Key             =   "print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":A388
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":AC62
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":AD74
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":AE86
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":AF98
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":B0AA
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":B1BC
            Key             =   "find"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":B2CE
            Key             =   "replace"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":B868
            Key             =   "findinfiles"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":BE02
            Key             =   "findsymbol"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":C39C
            Key             =   "replaceinfiles"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":C936
            Key             =   "left"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":CA48
            Key             =   "justify"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":CB5A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":D0F4
            Key             =   "right"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":D206
            Key             =   "center"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":D318
            Key             =   "arrange"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":D8B2
            Key             =   "viewdetails"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":D9C4
            Key             =   "source"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":DF5E
            Key             =   "designer"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":E4F8
            Key             =   "spell"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":E60A
            Key             =   "immediate"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":EBA4
            Key             =   "quickwatch"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":F13E
            Key             =   "breakpoints"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":F6D8
            Key             =   "viewlist"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":F7EA
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":FD84
            Key             =   "bold"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":FE96
            Key             =   "italic"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":FFA8
            Key             =   "underline"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":100BA
            Key             =   "viewlrgicons"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":101CC
            Key             =   "viewsmlicons"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":102DE
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":10878
            Key             =   "sortasc"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":1098A
            Key             =   "sortdesc"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":10A9C
            Key             =   "cascade"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":11036
            Key             =   "split"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":115D0
            Key             =   "tilehoriz"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":11B6A
            Key             =   "tilevert"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":12104
            Key             =   "dynamic"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":1269E
            Key             =   "index"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":12C38
            Key             =   "helpsearch"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":131D2
            Key             =   "help"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":132E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":1D2A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":2E6BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3DF23
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3E377
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3E7CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3EBF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3F0CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3F597
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3FAF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":3FFC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":4057B
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":409C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":40EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":54E14
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":68E73
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":6C877
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":7102D
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":7156E
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_rfq.frx":78E3A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid flex_grid 
      Height          =   5535
      Left            =   0
      TabIndex        =   20
      Top             =   960
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   9763
      _Version        =   393216
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
      TabIndex        =   21
      Top             =   6720
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   17
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
      TabIndex        =   24
      Top             =   6480
      Width           =   12015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Confirmed   "
      Height          =   255
      Left            =   11040
      TabIndex        =   23
      Top             =   0
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Approved and Not Confirmed  or  Partially Approved   "
      Height          =   255
      Left            =   6960
      TabIndex        =   22
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frm_rfq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gg As Integer
Public flg As Integer
Public strprices1 As String
Public byr As String
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
Private Sub Command3_Click()
Call striptab
Call flex_itemmodi
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
 
Static vprev As Integer
current = flex_grid.Row
'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
If flex_grid.TextMatrix(flex_grid.Row, 9) = "YES" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbGreen
        flex_grid.CellForeColor = vbGrayed
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 8) = "Pending" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 9) <> "YES" Then
        
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbWhite
        flex_grid.CellForeColor = vbBlue
        Next
End If
End If

'Current  row
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
End If
 
'----------------------
 
 '---------------------------
vprev = flex_grid.Row
End Sub
Private Sub flex_grid_DblClick()
On Error Resume Next
'back color
 
 
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
If flex_grid.TextMatrix(flex_grid.Row, 9) = "YES" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbGreen
        flex_grid.CellForeColor = vbGrayed
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 8) = "Pending" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 9) <> "YES" Then
        
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbWhite
        flex_grid.CellForeColor = vbBlue
        Next
End If
End If
 

'Current  row
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
End If


Unload buyerassigning
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)


requestquotation.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
requestquotation.dtp_rfq.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
requestquotation.cbo_vendorgroup.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
requestquotation.dtp_closingdate.Value = flex_grid.TextMatrix(flex_grid.Row, 5)
requestquotation.cbo_buyer.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
requestquotation.txt_remarks.Text = flex_grid.TextMatrix(flex_grid.Row, 7)


Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from prdetails pr,rfq r , purchaserequisition p where pr.rfqno = r.rfqno and pr.prno=p.prno and rfq_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "' and pr.buyer='" & byr & "' ", Cn, 3, 2
If Not prs.EOF Then
requestquotation.txt_project.Text = prs!project
requestquotation.txt_department.Text = prs!department
requestquotation.txt_requestor.Text = prs!requestor
With requestquotation.flex_med
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        
                    .TextMatrix(.Rows - 1, 0) = prs!pr_id
                    .TextMatrix(.Rows - 1, 2) = prs!Buyer
                    .TextMatrix(.Rows - 1, 3) = prs!Status
                    .TextMatrix(.Rows - 1, 4) = prs!itemid
                    .TextMatrix(.Rows - 1, 5) = prs!mrefcode
                    .TextMatrix(.Rows - 1, 6) = prs!material
                    .TextMatrix(.Rows - 1, 7) = prs!qty
                    .TextMatrix(.Rows - 1, 8) = prs!uom
                    .TextMatrix(.Rows - 1, 9) = prs!reqdate
                    .TextMatrix(.Rows - 1, 10) = prs!remarks
                    .TextMatrix(.Rows - 1, 11) = prs!jobcharge
                    .TextMatrix(.Rows - 1, 12) = prs!Location
                    .TextMatrix(.Rows - 1, 13) = prs!contactperson
                    .TextMatrix(.Rows - 1, 14) = prs!mtype
                    .TextMatrix(.Rows - 1, 15) = prs!fromdt
                    .TextMatrix(.Rows - 1, 16) = prs!todt
         
                     
        prs.MoveNext
    Wend
    prs.Close
End With
End If
Dim cnt As Integer
cnt = 0
flg = 0
Dim rfqterms As New ADODB.Recordset
If rfqterms.State Then rfqterms.Close
rfqterms.Open "select * from rfqterms where rfqno='" & requestquotation.txt_account.Text & "' order by rfqt_id ", Cn, 3, 2

If Not rfqterms.EOF Then
flg = 1
For cnt = 0 To rfqterms.RecordCount - 1
        If cnt <= 13 Then
        requestquotation.Check1(cnt).Caption = rfqterms!terms
                                    If rfqterms!chq = "Yes" Then
                                    requestquotation.Check1(cnt).Value = 1
                                    requestquotation.txt_terms(cnt).Text = rfqterms!termsdesc
                                    Else
                                    requestquotation.Check1(cnt).Value = 0
                                    End If
        rfqterms.MoveNext
        ElseIf cnt > 13 Then
                If rfqterms!terms <> "" Then
                requestquotation.txt_others(cnt - 14) = rfqterms!terms
                                    If rfqterms!chq = "Yes" Then
                                    requestquotation.Check1(cnt).Value = 1
                                    requestquotation.txt_terms(cnt).Text = rfqterms!termsdesc
                                    Else
                                    requestquotation.Check1(cnt).Value = 0
                                    End If
                End If
        rfqterms.MoveNext
        End If

Next cnt
End If
requestquotation.Show
SetParent requestquotation.hwnd, frm_rfq.hwnd
requestquotation.Height = 8280
requestquotation.Width = 11715
requestquotation.Top = 50
requestquotation.Left = 200



vprev = flex_grid.Row

End Sub
Private Sub flex_grid_SelChange()
On Error Resume Next
'back color
 
 
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
If flex_grid.TextMatrix(flex_grid.Row, 9) = "YES" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbGreen
        flex_grid.CellForeColor = vbGrayed
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 8) = "Pending" Then
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
        ElseIf flex_grid.TextMatrix(flex_grid.Row, 9) <> "YES" Then
        
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbWhite
        flex_grid.CellForeColor = vbBlue
        Next
End If


End If

 
'Current  row
If flex_grid.Row <> 0 Then
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
End If

Call flex_itemmodi

vprev = flex_grid.Row
'-------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Private Sub flex_item_Click()
On Error Resume Next
'back color
Toolbar1.Buttons(1).Enabled = True
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
flex_item.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_item.Col = 1

End If
vprev = flex_item.Row
End Sub
Private Sub flex_item_DblClick()
On Error Resume Next
'back color
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
'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
End If


Unload buyerassigning
Dim id As Double
id = 0
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
id = flex_grid.TextMatrix(flex_grid.Row, 0)

requestquotation.txt_account.Text = flex_grid.TextMatrix(flex_grid.Row, 1)
requestquotation.dtp_rfq.Value = flex_grid.TextMatrix(flex_grid.Row, 2)
requestquotation.cbo_vendorgroup.Text = flex_grid.TextMatrix(flex_grid.Row, 4)
requestquotation.dtp_closingdate.Value = flex_grid.TextMatrix(flex_grid.Row, 5)
requestquotation.cbo_buyer.Text = flex_grid.TextMatrix(flex_grid.Row, 6)
requestquotation.txt_remarks.Text = flex_grid.TextMatrix(flex_grid.Row, 7)

Dim prsd As New ADODB.Recordset
If prsd.State Then prsd.Close
prsd.Open "select * from prdetails pr,rfq r where pr.rfqno = r.rfqno and r.rfq_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "' and pr.buyer='" & byr & "'", Cn, 3, 2
With requestquotation.flex_med
    .Rows = 1
    While Not prsd.EOF
        .Rows = .Rows + 1
        
                    .TextMatrix(.Rows - 1, 0) = prs!pr_id
                    .TextMatrix(.Rows - 1, 2) = prs!Buyer
                    .TextMatrix(.Rows - 1, 3) = prs!Status
                    .TextMatrix(.Rows - 1, 4) = prs!itemid
                    .TextMatrix(.Rows - 1, 5) = prs!mrefcode
                    .TextMatrix(.Rows - 1, 6) = prs!material
                    .TextMatrix(.Rows - 1, 7) = prs!qty
                    .TextMatrix(.Rows - 1, 8) = prs!uom
                    .TextMatrix(.Rows - 1, 9) = prs!reqdate
                    .TextMatrix(.Rows - 1, 10) = prs!remarks
                    .TextMatrix(.Rows - 1, 11) = prs!jobcharge
                    .TextMatrix(.Rows - 1, 12) = prs!Location
                    .TextMatrix(.Rows - 1, 13) = prs!contactperson
                    .TextMatrix(.Rows - 1, 14) = prs!mtype
                    .TextMatrix(.Rows - 1, 15) = prs!fromdt
                    .TextMatrix(.Rows - 1, 16) = prs!todt
         
        prsd.MoveNext
    Wend
End With
 


requestquotation.Show
SetParent requestquotation.hwnd, frm_rfq.hwnd
requestquotation.Height = 8280
requestquotation.Width = 11715
requestquotation.Top = 50
requestquotation.Left = 200

 

vprev = flex_grid.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
dtps_date.Value = Format(Date, "MMM/yyyy")
Call connect
main.lbltitle.Caption = "RFQ"

Call flex_title
Call flex_subtitle
 
 
Me.Top = 5
Me.Left = 5


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

  
cbos_job.Clear
Dim jb As New ADODB.Recordset
If jb.State Then jb.Close
jb.Open "select DISTINCT(jobcharge) from prdetails order by jobcharge", Cn, 3, 2
While Not jb.EOF
cbos_job.AddItem jb(0)
jb.MoveNext
Wend
jb.Close

jb.Open "select DISTINCT(a_name) from userid where a_userid='" & main.Label2.Caption & "' ", Cn, 3, 2
If Not jb.EOF Then
byr = jb(0)
End If




Call striptab
 
'    AnimateForm buyerassigning, -1, -1, aload, Val(25), Val(20), Val(1), Val(2)
    End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
Unload buyerassigning
'    AnimateForm buyerassigning, -1, -1, aUnload, Val(25), Val(20), Val(1), Val(2)
End Sub
 
Public Sub flex_title()
On Error Resume Next

   With flex_grid
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "RFQ No."
        .ColWidth(1) = 1500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "RFQ Date"
        .ColWidth(2) = 1500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "MSR No."
        .ColWidth(3) = 1500
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Vendor Group"
        .ColWidth(4) = 1500
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Closing Date"
        .ColWidth(5) = 1500
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Buyer"
        .ColWidth(6) = 4000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Remarks"
        .ColWidth(7) = 4500
        .ColAlignment(7) = 0
        .ColWidth(8) = 0
        .ColWidth(9) = 0
    End With
End Sub
Private Sub opt_a_Click()
Call striptab
Call flex_itemmodi
End Sub
Public Sub flex_datadate()
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
If opt_na.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.rfqdate,r.prno,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and month(r.rfqdate)='" & Format(dtps_date.Value, "mm") & "' and year(r.rfqdate)='" & Format(dtps_date.Value, "yyyy") & "' and pr.buyer ='" & byr & "' and r.vendorgroup ='Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_all.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.rfqdate,r.prno,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and month(r.rfqdate)='" & Format(dtps_date.Value, "mm") & "' and year(r.rfqdate)='" & Format(dtps_date.Value, "yyyy") & "' and pr.buyer ='" & byr & "'  order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_a.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.rfqdate,r.prno,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and month(r.rfqdate)='" & Format(dtps_date.Value, "mm") & "' and year(r.rfqdate)='" & Format(dtps_date.Value, "yyyy") & "' and pr.buyer ='" & byr & "' and r.vendorgroup <> 'Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
End If
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
If fldata(8) = "Pending" Then
flex_grid.Row = .Rows - 1
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
flex_grid.Col = 1

End If
If fldata(9) = "YES" Then

flex_grid.Row = .Rows - 1
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbGreen
flex_grid.CellForeColor = vbGrayed
Next
flex_grid.Col = 1

End If

        fldata.MoveNext
    Wend
End With

End Sub
Public Sub flex_subtitle()
On Error Resume Next

   
      With flex_item
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Buyer"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        
        .TextMatrix(0, 3) = "Status"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        
        .TextMatrix(0, 4) = "ItemId"
        .ColWidth(4) = 1500
        .ColAlignment(4) = 0
        
        .TextMatrix(0, 5) = "Mfr.Ref Code"
        .ColWidth(5) = 1200
        .ColAlignment(5) = 0
        
        .TextMatrix(0, 6) = "Material"
        .ColWidth(6) = 5000
        .ColAlignment(6) = 0
        
        .TextMatrix(0, 7) = "Qty"
        .ColWidth(7) = 600
        .ColAlignment(7) = 0
        
        .TextMatrix(0, 8) = "UOM"
        .ColWidth(8) = 600
        .ColAlignment(8) = 0
        
        .TextMatrix(0, 9) = "ReqDate"
        .ColWidth(9) = 800
        .ColAlignment(9) = 0
        
        .TextMatrix(0, 10) = "Remarks"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0

        .TextMatrix(0, 11) = "Jobcharge"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0
        
        .TextMatrix(0, 12) = "Work Location"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0

        .TextMatrix(0, 13) = "Stor Loc"
        .ColWidth(13) = 1200
        .ColAlignment(13) = 0
        
               
        .TextMatrix(0, 14) = "Mat Type"
        .ColWidth(14) = 1200
        .ColAlignment(14) = 0
        
        .TextMatrix(0, 15) = "FromDate"
        .ColWidth(15) = 1200
        .ColAlignment(15) = 0

        .TextMatrix(0, 16) = "ToDate"
        .ColWidth(16) = 1200
        .ColAlignment(16) = 0
         
    End With
End Sub
Public Sub flex_itemmodi()
On Error Resume Next
current = flex_grid.Row
'Reset to previous row
flex_item.Clear
Call flex_subtitle
If flex_grid.TextMatrix(flex_grid.Row, 0) = "" Then Exit Sub
'''''''''''''''''''''''''''''''''''''''''''''''''
Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
If opt_na.Value = True Then
prs.Open "select * from prdetails pr,rfq r where pr.rfqno = r.rfqno and rfq_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "' and pr.buyer ='" & byr & "' and r.vendorgroup ='Not Assigned'", Cn, 3, 2
ElseIf opt_all.Value = True Then
prs.Open "select * from prdetails pr,rfq r where pr.rfqno = r.rfqno and rfq_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "' and pr.buyer ='" & byr & "'", Cn, 3, 2
ElseIf opt_a.Value = True Then
prs.Open "select * from prdetails pr,rfq r where pr.rfqno = r.rfqno and rfq_id='" & flex_grid.TextMatrix(flex_grid.Row, 0) & "' and pr.buyer ='" & byr & "' and r.vendorgroup <> 'Not Assigned'", Cn, 3, 2
End If

With flex_item
    .Rows = 1
    While Not prs.EOF
        .Rows = .Rows + 1
        
                    .TextMatrix(.Rows - 1, 0) = prs!pr_id
                    .TextMatrix(.Rows - 1, 2) = prs!Buyer
                    .TextMatrix(.Rows - 1, 3) = prs!Status
                    .TextMatrix(.Rows - 1, 4) = prs!itemid
                    .TextMatrix(.Rows - 1, 5) = prs!mrefcode
                    .TextMatrix(.Rows - 1, 6) = prs!material
                    .TextMatrix(.Rows - 1, 7) = prs!qty
                    .TextMatrix(.Rows - 1, 8) = prs!uom
                    .TextMatrix(.Rows - 1, 9) = prs!reqdate
                    .TextMatrix(.Rows - 1, 10) = prs!remarks
                    .TextMatrix(.Rows - 1, 11) = prs!jobcharge
                    .TextMatrix(.Rows - 1, 12) = prs!Location
                    strprices1 = ""
                    strprices1 = prs!Location
                    .TextMatrix(.Rows - 1, 13) = prs!contactperson
                    .TextMatrix(.Rows - 1, 14) = prs!mtype
                    .TextMatrix(.Rows - 1, 15) = prs!fromdt
                    .TextMatrix(.Rows - 1, 16) = prs!todt

        prs.MoveNext
    Wend
End With

''''''''''''''''''''''''''''''''''''''''''''
vprev = flex_grid.Row
End Sub

Public Sub flex_datarequestor()
If cbos_requestor.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
If opt_na.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.requestor='" & cbos_requestor.Text & "' and pr.buyer ='" & byr & "' and  r.vendorgroup ='Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_all.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.requestor='" & cbos_requestor.Text & "' and pr.buyer ='" & byr & "'  order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_a.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.requestor='" & cbos_requestor.Text & "' and pr.buyer ='" & byr & "' and  r.vendorgroup <> 'Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
End If
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
        
If fldata(8) = "Pending" Then
flex_grid.Row = .Rows - 1
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
flex_grid.Col = 1

End If
If fldata(9) = "YES" Then

flex_grid.Row = .Rows - 1
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbGreen
flex_grid.CellForeColor = vbGrayed
Next
flex_grid.Col = 1

End If

        fldata.MoveNext
    Wend
End With

End Sub

Public Sub flex_datajob()
If cbos_job.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If
On Error Resume Next
Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
If opt_na.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and pr.jobcharge='" & cbos_job.Text & "' and pr.buyer ='" & byr & "' and  r.vendorgroup ='Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_all.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and pr.jobcharge='" & cbos_job.Text & "' and pr.buyer ='" & byr & "'  order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_a.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and pr.jobcharge='" & cbos_job.Text & "' and pr.buyer ='" & byr & "' and  r.vendorgroup <Not Assigned> '' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
End If
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
        
If fldata(8) = "Pending" Then
flex_grid.Row = .Rows - 1
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
flex_grid.Col = 1

End If
If fldata(9) = "YES" Then

flex_grid.Row = .Rows - 1
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbGreen
flex_grid.CellForeColor = vbGrayed
Next
flex_grid.Col = 1

End If

        fldata.MoveNext
    Wend
End With

End Sub
Public Sub flex_vendor()
On Error Resume Next
If cbos_vendor.Text = "" Then
flex_grid.Clear
flex_grid.Rows = 2
Call flex_title
Exit Sub
End If

Dim fldata As New ADODB.Recordset
If fldata.State Then fldata.Close
If opt_na.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.recommendedvendor='" & cbos_vendor.Text & "' and pr.buyer ='" & byr & "' and r.vendorgroup ='Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_all.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.recommendedvendor='" & cbos_vendor.Text & "' and pr.buyer ='" & byr & "'  order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
ElseIf opt_a.Value = True Then
fldata.Open "select DISTINCT(r.rfq_id),r.rfqno,r.prno,r.rfqdate,r.vendorgroup,r.closingdate,r.buyer,r.remarks,p.status,p.confirmation from rfq r ,prdetails pr ,purchaserequisition p where r.rfqno=pr.rfqno and pr.prno=p.prno and p.recommendedvendor='" & cbos_vendor.Text & "' and pr.buyer ='" & byr & "' and r.vendorgroup <> 'Not Assigned' order by r.rfqno,r.rfqdate", Cn, 3, 2 'p,userproject u where p.proj_key=u.project and u.username='" & byr & "' order by p.proj_key
End If
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
        
If fldata(8) = "Pending" Then
flex_grid.Row = .Rows - 1
        For i = 1 To flex_grid.Cols - 1
        flex_grid.Col = i
        flex_grid.CellBackColor = vbRed
        flex_grid.CellForeColor = vbBlue
        Next
flex_grid.Col = 1

End If
If fldata(9) = "YES" Then

flex_grid.Row = .Rows - 1
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = vbGreen
flex_grid.CellForeColor = vbGrayed
Next
flex_grid.Col = 1

End If

        fldata.MoveNext
    Wend
End With

End Sub


Public Sub striptab()

If SSTab1.Caption = "   Vendor   " Then
Call flex_vendor
ElseIf SSTab1.Caption = "   Requestor  " Then
 
Call flex_datarequestor
ElseIf SSTab1.Caption = "       Job        " Then
 
Call flex_datajob
ElseIf SSTab1.Caption = "   Month/Year   " Then
 
Call flex_datadate
End If
 
End Sub
 
Public Sub buyassign1()

End Sub

Private Sub opt_all_Click()
Call striptab
Call flex_itemmodi

End Sub

Private Sub opt_na_Click()
Call striptab
Call flex_itemmodi
End Sub
