VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form quotation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quotation"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "quotation.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Quotation Details"
      TabPicture(0)   =   "quotation.frx":11409
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Terms and Conditions"
      TabPicture(1)   =   "quotation.frx":11425
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other Details/ Remarks"
      TabPicture(2)   =   "quotation.frx":11441
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   -75000
         TabIndex        =   54
         Top             =   300
         Width           =   11055
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   16
            Left            =   360
            TabIndex        =   113
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   15
            Left            =   360
            TabIndex        =   112
            Top             =   5640
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   14
            Left            =   360
            TabIndex        =   111
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   13
            Left            =   360
            TabIndex        =   110
            Top             =   4920
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   12
            Left            =   360
            TabIndex        =   109
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   11
            Left            =   360
            TabIndex        =   108
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   10
            Left            =   360
            TabIndex        =   107
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   9
            Left            =   360
            TabIndex        =   106
            Top             =   3480
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   360
            TabIndex        =   105
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   7
            Left            =   360
            TabIndex        =   104
            Top             =   2760
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   6
            Left            =   360
            TabIndex        =   103
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   5
            Left            =   360
            TabIndex        =   102
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   4
            Left            =   360
            TabIndex        =   101
            Text            =   "Price"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   3
            Left            =   360
            TabIndex        =   100
            Text            =   "Quotation Validity"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   1
            Left            =   1920
            TabIndex        =   91
            Top             =   600
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   90
            Top             =   600
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   2
            Left            =   1920
            TabIndex        =   89
            Top             =   960
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   3
            Left            =   1920
            TabIndex        =   87
            Top             =   1320
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   86
            Top             =   1320
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   0
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   240
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   4
            Left            =   1920
            TabIndex        =   83
            Top             =   1680
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   82
            Top             =   1680
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   5
            Left            =   1920
            TabIndex        =   81
            Top             =   2040
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   80
            Top             =   2040
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   6
            Left            =   1920
            TabIndex        =   79
            Top             =   2400
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   7
            Left            =   1920
            TabIndex        =   77
            Top             =   2760
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   76
            Top             =   2760
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   8
            Left            =   1920
            TabIndex        =   75
            Top             =   3120
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   74
            Top             =   3120
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   9
            Left            =   1920
            TabIndex        =   73
            Top             =   3480
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   72
            Top             =   3480
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   10
            Left            =   1920
            TabIndex        =   71
            Top             =   3840
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   70
            Top             =   3840
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   11
            Left            =   1920
            TabIndex        =   69
            Top             =   4200
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   68
            Top             =   4200
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   67
            Top             =   4560
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   12
            Left            =   1920
            TabIndex        =   66
            Top             =   4560
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   65
            Top             =   4920
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   13
            Left            =   1920
            TabIndex        =   64
            Top             =   4920
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   120
            TabIndex        =   63
            Top             =   5280
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   14
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   5280
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   0
            Left            =   360
            TabIndex        =   61
            Text            =   "Delivery Terms"
            Top             =   240
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   60
            Top             =   5640
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   15
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   5640
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   1
            Left            =   360
            TabIndex        =   58
            Text            =   "Delivery Period"
            Top             =   600
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   57
            Top             =   6000
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            BackColor       =   &H00C0FFFF&
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
            Index           =   16
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   56
            Top             =   6000
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
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
            Index           =   2
            Left            =   360
            TabIndex        =   55
            Text            =   "Payment Terms"
            Top             =   960
            Width           =   1455
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   -75000
         TabIndex        =   36
         Top             =   300
         Width           =   11055
         Begin VB.TextBox txt_remarks 
            Height          =   3615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   52
            Top             =   3000
            Width           =   10695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Other Details"
            ForeColor       =   &H00800000&
            Height          =   2415
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   10695
            Begin VB.TextBox txt_desig 
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
               Left            =   120
               TabIndex        =   44
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox cbo_toperson 
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
               Left            =   120
               TabIndex        =   43
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox txt_dept 
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
               Left            =   120
               TabIndex        =   42
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox txt_oref 
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
               Left            =   5880
               TabIndex        =   41
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox cbo_mode 
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
               Left            =   5880
               TabIndex        =   40
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txt_yref 
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
               Left            =   5880
               TabIndex        =   39
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox txt_refno 
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
               Left            =   8040
               TabIndex        =   38
               Top             =   480
               Width           =   2295
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Addressed-To-Person"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   240
               Width           =   1845
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   1005
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   49
               Top             =   1680
               Width           =   1020
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Mode-Of-QuotReceived"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   5880
               TabIndex        =   48
               Top             =   240
               Width           =   1980
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Our Reference"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   5880
               TabIndex        =   47
               Top             =   960
               Width           =   1245
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Your Reference"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   5880
               TabIndex        =   46
               Top             =   1680
               Width           =   1320
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Ref No./ Details"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   8040
               TabIndex        =   45
               Top             =   240
               Width           =   1590
            End
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   2760
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7455
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   11055
         Begin VB.Frame fr_auth 
            BackColor       =   &H00FF8080&
            Caption         =   "Availability Status"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1095
            Left            =   120
            TabIndex        =   30
            Top             =   6240
            Width           =   10815
            Begin VB.OptionButton opt_all 
               BackColor       =   &H00FF8080&
               Caption         =   "Apply All"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   3360
               TabIndex        =   34
               ToolTipText     =   "Click to Authorize all Line Items"
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton opt_ind 
               BackColor       =   &H00FF8080&
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   4920
               TabIndex        =   33
               ToolTipText     =   "Click to Uncheck LineItems"
               Top             =   480
               Width           =   975
            End
            Begin VB.ComboBox cbo_astatus 
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
               ItemData        =   "quotation.frx":1145D
               Left            =   120
               List            =   "quotation.frx":11473
               TabIndex        =   32
               Text            =   "Comply"
               ToolTipText     =   "Select Authorization Status"
               Top             =   480
               Width           =   2715
            End
            Begin VB.CommandButton cmd_authorize 
               BackColor       =   &H00FFC0C0&
               Caption         =   " << Save Quotation >>"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   735
               Left            =   8400
               Picture         =   "quotation.frx":114AE
               Style           =   1  'Graphical
               TabIndex        =   31
               ToolTipText     =   "Click to Authorize MSR"
               Top             =   195
               Width           =   2175
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   29
            Top             =   1680
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "EDIT LINE ITEM"
            Height          =   2655
            Left            =   120
            TabIndex        =   9
            Top             =   3600
            Width           =   10695
            Begin VB.ComboBox cbo_materialtype 
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
               Left            =   1440
               TabIndex        =   118
               Top             =   720
               Width           =   3075
            End
            Begin VB.Frame frame_ms 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Material Service Duration"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   735
               Left            =   4680
               TabIndex        =   114
               ToolTipText     =   $"quotation.frx":117B8
               Top             =   600
               Width           =   3735
               Begin MSComCtl2.DTPicker dtp_from 
                  Height          =   285
                  Left            =   120
                  TabIndex        =   115
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   90701825
                  CurrentDate     =   38455
               End
               Begin MSComCtl2.DTPicker dtp_to 
                  Height          =   285
                  Left            =   2280
                  TabIndex        =   116
                  Top             =   360
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   503
                  _Version        =   393216
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Format          =   90701825
                  CurrentDate     =   38455
               End
               Begin VB.Label Label5 
                  BackColor       =   &H80000009&
                  Caption         =   "To"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00C00000&
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   117
                  Top             =   360
                  Width           =   375
               End
            End
            Begin VB.Frame fr_tit 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               Height          =   300
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   10575
               Begin VB.CheckBox chk_qty 
                  BackColor       =   &H00FF8080&
                  Caption         =   "Change Quantity"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   4320
                  TabIndex        =   99
                  ToolTipText     =   "If requested quantity is not available check to quote partially"
                  Top             =   0
                  Width           =   1935
               End
               Begin VB.CheckBox chk_item 
                  BackColor       =   &H00FF8080&
                  Caption         =   "Substitute Item"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   255
                  Left            =   2520
                  TabIndex        =   98
                  ToolTipText     =   "If requested item is not available check to quote Substitute Item"
                  Top             =   0
                  Width           =   1695
               End
               Begin VB.CommandButton Command1 
                  Caption         =   "Save Item"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   97
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.CommandButton cmd_delete 
                  Caption         =   "Clear"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   1320
                  Style           =   1  'Graphical
                  TabIndex        =   96
                  Top             =   0
                  Width           =   975
               End
            End
            Begin VB.ComboBox cbo_lookup 
               Enabled         =   0   'False
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
               Left            =   1440
               TabIndex        =   93
               Top             =   1320
               Width           =   3075
            End
            Begin VB.ComboBox cbo_category 
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   92
               Top             =   1680
               Width           =   6795
            End
            Begin VB.TextBox txt_xchg 
               Enabled         =   0   'False
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
               Left            =   1290
               TabIndex        =   13
               Text            =   "1.00"
               Top             =   2280
               Width           =   840
            End
            Begin VB.ComboBox cbo_curr 
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
               Left            =   135
               TabIndex        =   12
               Text            =   "RM"
               Top             =   2280
               Width           =   1095
            End
            Begin VB.TextBox txt_amount 
               Enabled         =   0   'False
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
               Left            =   2190
               TabIndex        =   14
               Top             =   2280
               Width           =   1800
            End
            Begin VB.TextBox txt_unitrate 
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
               Left            =   9540
               TabIndex        =   11
               Top             =   1680
               Width           =   1080
            End
            Begin MSComCtl2.DTPicker dtp_reqd 
               Height          =   285
               Left            =   4050
               TabIndex        =   15
               Top             =   2280
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   90701825
               CurrentDate     =   38455
            End
            Begin VB.TextBox txt_qty 
               Enabled         =   0   'False
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
               Left            =   6960
               TabIndex        =   17
               Top             =   1680
               Width           =   1080
            End
            Begin VB.ComboBox cbo_uom 
               Enabled         =   0   'False
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
               Left            =   8115
               TabIndex        =   10
               Top             =   1680
               Width           =   1335
            End
            Begin VB.TextBox txt_notes 
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
               Left            =   6840
               TabIndex        =   18
               Top             =   2280
               Width           =   3735
            End
            Begin MSComCtl2.DTPicker dtp_pdate 
               Height          =   285
               Left            =   5445
               TabIndex        =   16
               Top             =   2280
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   90701825
               CurrentDate     =   38455
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Material Type"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00A04729&
               Height          =   195
               Left            =   120
               TabIndex        =   119
               Top             =   840
               Width           =   1155
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Select Item By"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   1440
               Width           =   1275
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Exch Rate"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   1305
               TabIndex        =   28
               Top             =   2040
               Width           =   855
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Currency"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   135
               TabIndex        =   27
               Top             =   2040
               Width           =   795
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Amount(RM)"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   2220
               TabIndex        =   26
               Top             =   2040
               Width           =   1065
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Unit Rate"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   9540
               TabIndex        =   25
               Top             =   1440
               Width           =   780
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Promised Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   0
               Left            =   5400
               TabIndex        =   23
               Top             =   2040
               Width           =   1260
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   6960
               TabIndex        =   22
               Top             =   1440
               Width           =   720
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   8115
               TabIndex        =   21
               Top             =   1440
               Width           =   390
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   6840
               TabIndex        =   20
               Top             =   2040
               Width           =   765
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Reqd. Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   1
               Left            =   4080
               TabIndex        =   19
               Top             =   2040
               Width           =   960
            End
         End
         Begin VB.TextBox txt_account 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cbo_vendor 
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
            Left            =   3360
            TabIndex        =   2
            Top             =   360
            Width           =   5295
         End
         Begin MSComCtl2.DTPicker dtp_qt 
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   90701825
            CurrentDate     =   38455
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   2775
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   1
            Cols            =   21
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            BackColorBkg    =   16777215
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   5640
            TabIndex        =   24
            Top             =   840
            Width           =   5175
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Quotation No."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Quotation  Date"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            Top             =   120
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3360
            TabIndex        =   5
            Top             =   120
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "quotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kl As Integer

Private Sub cmd_new_Click()
kl = 1
cbo_name.Clear
End Sub

Private Sub cbo_category_Change()
If cbo_lookup.Text = "Search" Then
'cbo_category.Clear
Dim mat As New ADODB.Recordset
mat.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml3name like'" & cbo_category.Text & "%' or ml4name like'" & cbo_category.Text & "%' order by ml3name,ml4name", Cn, 3, 2
'While Not mat.EOF
Dim i As Integer
i = 0
For i = 0 To mat.RecordCount - 1
'cbo_category.Clear
If cbo_category.List(i) = mat(2) & "  -  " & mat(3) & "  -  " & mat(0) & "  -  " & mat(1) Then
Else
cbo_category.AddItem mat(2) & "  -  " & mat(3) & "  -  " & mat(0) & "  -  " & mat(1)
End If
mat.MoveNext
Next i
'mat.MoveNext
'Wend
mat.Close
End If

End Sub

Private Sub cbo_category_Click()
sc = Split(cbo_category.Text, "  -  ", Len(cbo_category.Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close

If cbo_lookup.Text = "Item ID" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(2) & "' and ml4name='" & sc(3) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(2) & "' and ml4name='" & sc(3) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Item Description" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Enter Item Manually" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Search" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2

End If
If Not um.EOF Then
cbo_uom.Text = um(0)
End If
  
Command1.Enabled = True

End Sub

Private Sub cbo_category_KeyPress(KeyAscii As Integer)
If cbo_lookup.Text = "Item ID" Then
KeyAscii = 0
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
KeyAscii = 0
ElseIf cbo_lookup.Text = "Item Description" Then
KeyAscii = 0
End If
End Sub

Private Sub cbo_curr_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_lookup_Click()
cbo_category.Clear
 Dim med As New ADODB.Recordset
If med.State Then med.Close

 If cbo_lookup.Text = "Item ID" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(0) & "  -  " & med(1) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(1) & "  -  " & med(0) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close


ElseIf cbo_lookup.Text = "Item Description" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close

ElseIf cbo_lookup.Text = "Enter Item Manually" Then


ElseIf cbo_lookup.Text = "Search" Then
cbo_category.Clear
End If


End Sub

Private Sub cbo_lookup_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_materialtype_Click()
Call mrenproc
End Sub

Private Sub cbo_materialtype_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub
Public Sub mrenproc()
On Error Resume Next
spmt = Split(cbo_materialtype.Text, "  -  ", Len(cbo_materialtype.Text), vbTextCompare)
Dim mren As New ADODB.Recordset
If mren.State Then mren.Close
mren.Open "select DISTINCT(rental) from materialtype where mtcode='" & spmt(0) & "'", Cn, 3, 2
If Not mren.EOF Then
If mren(0) = "Yes" Then
frame_ms.Visible = True
Else
frame_ms.Visible = False
End If
End If
End Sub

Private Sub cbo_uom_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub chk_item_Click()
Call itemvalues
End Sub

Private Sub chk_qty_Click()
Call itemvalues
End Sub

Private Sub cmd_authorize_Click()

Dim rcnt As Integer
rcnt = 0
For rcnt = 1 To flex_med.Rows - 1
If flex_med.TextMatrix(rcnt, 8) = "" Then
MsgBox " Kindly update Quotation before saving"
Exit Sub
End If
Next rcnt
ms = MsgBox("Do you want to change Terms & Conditions", vbYesNo)
If ms = vbYes Then
Exit Sub
Else

  '--------------------------------------------------------------------------------------
'If frm_quotation.flg = 1 Then
Cn.Execute "delete from quotationterms where qno='" & txt_account.Text & "'"
Dim rft As New ADODB.Recordset
If rft.State Then rft.Close
rft.Open "select * from quotationterms", Cn, 3, 2
Dim inc As Integer
inc = 0
For inc = 0 To 16

            If txt_others(inc).Text <> "" Then
            
            rft.AddNew
            rft!qno = txt_account.Text
            rft!terms = txt_others(inc).Text
            rft!termsdesc = txt_terms(inc).Text
                            If Check1(inc).Value = 1 Then
                            rft!chq = "Yes"
                            Else
                            rft!chq = "No"
                            End If
            rft.Update
            rft.MoveNext
            End If


Next inc
rft.Close
'End If
            '--------------------------------------------------------------------------------------
          
Call quotationupdate
frm_quotation.striptab
frm_quotation.flex_itemmodi
If main.VstrEmail = 0 Then
MsgBox " Quotation Updated Successfully, Mail is been sent to Next level Authorization: Kindly wait for confirmation message"
Call assadmailservicequote
Else
MsgBox "Quotation Updated Successfully"
End If
'Call generatepo
End If
End Sub

Private Sub cmd_delete_Click()
Call clearitems

End Sub

Private Sub Command1_Click()
If cbo_category.Text = "" Then
MsgBox "Select Material"
cbo_category.SetFocus
Exit Sub
End If
If txt_qty.Text = "" Then
MsgBox " Enter Quantity"
txt_qty.SetFocus
Exit Sub
End If
If cbo_uom.Text = "" Then
MsgBox "Select UOM"
cbo_uom.SetFocus
Exit Sub
End If

If txt_unitrate.Text = "" Then
MsgBox " Enter UnitRate"
txt_unitrate.SetFocus
Exit Sub
End If
If cbo_curr.Text = "" Then
MsgBox "Select Currency"
cbo_curr.SetFocus
Exit Sub
End If
If txt_amount.Text = "" Then
MsgBox "Enter Amount"
txt_amount.SetFocus
Exit Sub
End If
If txt_xchg.Text = "" Then
MsgBox "Enter Exchange Rate"
txt_xchg.SetFocus
Exit Sub
End If

If dtp_reqd.Value < Date Then
MsgBox "Required Date cannot be Earlier then Todays Date"

Exit Sub
End If
If dtp_pdate.Value < Date Then
MsgBox "Promised Date cannot be Earlier then Todays Date"

Exit Sub
End If

Call itemvalues
Dim jj As Integer
jj = 0
 If Not cbo_category.Text = "" Then
    If Not txt_qty.Text = "" Then
          spl = Split(cbo_category.Text, "  -  ", Len(cbo_category.Text), vbTextCompare)
        If kl = 1 Then

                    With flex_med
                        
                        .Rows = .Rows + 1
                        

                        flex_med.TextMatrix(.Rows - 1, 2) = cbo_astatus.Text
                        
                        If cbo_lookup.Text = "Item ID" Then
                        .TextMatrix(.Rows - 1, 3) = spl(0)
                        .TextMatrix(.Rows - 1, 4) = spl(1)
                        .TextMatrix(.Rows - 1, 5) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Mfr PartNo." Then
                        .TextMatrix(.Rows - 1, 3) = spl(1)
                        .TextMatrix(.Rows - 1, 4) = spl(0)
                        .TextMatrix(.Rows - 1, 5) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Item Description" Then
                        .TextMatrix(.Rows - 1, 5) = spl(0) & "  -  " & spl(1)
                        .TextMatrix(.Rows - 1, 3) = spl(2)
                        .TextMatrix(.Rows - 1, 4) = spl(3)
                        
                        ElseIf cbo_lookup.Text = "Enter Item Manually" Then
                        .TextMatrix(.Rows - 1, 5) = cbo_category.Text
                        .TextMatrix(.Rows - 1, 3) = ""
                        .TextMatrix(.Rows - 1, 4) = ""
                        
                        ElseIf cbo_lookup.Text = "Search" Then
                        .TextMatrix(.Rows - 1, 5) = spl(0) & "  -  " & spl(1)
                        .TextMatrix(.Rows - 1, 3) = spl(2)
                        .TextMatrix(.Rows - 1, 4) = spl(3)
                        
                        End If
                        flex_med.TextMatrix(.Rows - 1, 6) = txt_qty.Text
                        flex_med.TextMatrix(.Rows - 1, 7) = cbo_uom.Text
                        flex_med.TextMatrix(.Rows - 1, 8) = txt_unitrate.Text
                        flex_med.TextMatrix(.Rows - 1, 9) = cbo_curr.Text
                        flex_med.TextMatrix(.Rows - 1, 10) = txt_xchg.Text
                        flex_med.TextMatrix(.Rows - 1, 11) = txt_amount.Text
                        flex_med.TextMatrix(.Rows - 1, 12) = dtp_reqd.Value
                        flex_med.TextMatrix(.Rows - 1, 13) = dtp_pdate.Value
                        flex_med.TextMatrix(.Rows - 1, 17) = txt_notes.Text
                        flex_med.TextMatrix(.Rows - 1, 14) = cbo_materialtype
                        flex_med.TextMatrix(.Rows - 1, 15) = dtp_from.Value
                        flex_med.TextMatrix(.Rows - 1, 16) = dtp_to.Value
                   
                        
                    End With
        Else
                      jj = flex_med.Row
                      
                        flex_med.TextMatrix(jj, 2) = cbo_astatus.Text
                        
                        If cbo_lookup.Text = "Item ID" Then
                        flex_med.TextMatrix(jj, 3) = spl(0)
                        flex_med.TextMatrix(jj, 4) = spl(1)
                        flex_med.TextMatrix(jj, 5) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Mfr PartNo." Then
                        flex_med.TextMatrix(jj, 3) = spl(1)
                        flex_med.TextMatrix(jj, 4) = spl(0)
                        flex_med.TextMatrix(jj, 5) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Item Description" Then
                        flex_med.TextMatrix(jj, 5) = spl(0) & "  -  " & spl(1)
                        flex_med.TextMatrix(jj, 3) = spl(2)
                        flex_med.TextMatrix(jj, 4) = spl(3)
                        
                         ElseIf cbo_lookup.Text = "Enter Item Manually" Then
                        flex_med.TextMatrix(jj, 5) = cbo_category.Text
                        flex_med.TextMatrix(jj, 3) = ""
                        flex_med.TextMatrix(jj, 4) = ""
                        
                        ElseIf cbo_lookup.Text = "Search" Then
                        flex_med.TextMatrix(jj, 5) = spl(0) & "  -  " & spl(1)
                        flex_med.TextMatrix(jj, 3) = spl(2)
                        flex_med.TextMatrix(jj, 4) = spl(3)
                        End If
                        
                        flex_med.TextMatrix(jj, 6) = txt_qty.Text
                        flex_med.TextMatrix(jj, 7) = cbo_uom.Text
                        flex_med.TextMatrix(jj, 8) = txt_unitrate.Text
                        flex_med.TextMatrix(jj, 9) = cbo_curr.Text
                        flex_med.TextMatrix(jj, 10) = txt_xchg.Text
                        flex_med.TextMatrix(jj, 11) = txt_amount.Text
                        flex_med.TextMatrix(jj, 12) = dtp_reqd.Value
                        flex_med.TextMatrix(jj, 13) = dtp_pdate.Value
                        flex_med.TextMatrix(jj, 17) = txt_notes.Text
                        flex_med.TextMatrix(jj, 14) = cbo_materialtype
                        flex_med.TextMatrix(jj, 15) = dtp_from.Value
                        flex_med.TextMatrix(jj, 16) = dtp_to.Value
        
        End If
    End If
    End If
Call clearitems
fr_tit.Enabled = False

End Sub

Private Sub flex_med_Click()
On Error Resume Next
'back color
 
Static vprev As Integer

current = flex_med.Row

'Reset to previous row
If vprev > 0 Then
    flex_med.Row = vprev
    flex_med.Col = 1
    Set flex_med.CellPicture = LoadPicture()
    
    For i = 1 To flex_med.Cols - 1
    flex_med.Col = i
    flex_med.CellBackColor = vbWhite
Next
End If

'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1


vprev = flex_med.Row
End Sub

Private Sub flex_med_DblClick()
On Error Resume Next
'back color
 fr_tit.Enabled = True
 Command1.Enabled = False
Static vprev As Integer

current = flex_med.Row

'Reset to previous row
If vprev > 0 Then
    flex_med.Row = vprev
    flex_med.Col = 1
    Set flex_med.CellPicture = LoadPicture()
    
    For i = 1 To flex_med.Cols - 1
    flex_med.Col = i
    flex_med.CellBackColor = vbWhite
Next
End If

'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1

Dim idd As Double
idd = 0
 
cbo_astatus.Text = flex_med.TextMatrix(flex_med.Row, 2)
cbo_category.Text = flex_med.TextMatrix(flex_med.Row, 3) & "  -  " & flex_med.TextMatrix(flex_med.Row, 4) & "  -  " & flex_med.TextMatrix(flex_med.Row, 5)
txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 6)
cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 7)
txt_unitrate.Text = flex_med.TextMatrix(flex_med.Row, 8)
If flex_med.TextMatrix(flex_med.Row, 9) = "" Then
cbo_curr.Text = "RM"
Else
cbo_curr.Text = flex_med.TextMatrix(flex_med.Row, 9)
End If
If flex_med.TextMatrix(flex_med.Row, 10) = "" Then
txt_xchg.Text = 1
Else
txt_xchg.Text = flex_med.TextMatrix(flex_med.Row, 10)
End If
txt_amount.Text = flex_med.TextMatrix(flex_med.Row, 11)
dtp_reqd.Value = flex_med.TextMatrix(flex_med.Row, 12)
dtp_pdate.Value = flex_med.TextMatrix(flex_med.Row, 13)
txt_notes.Text = flex_med.TextMatrix(flex_med.Row, 17)


cbo_materialtype.Text = flex_med.TextMatrix(flex_med.Row, 14)
dtp_from.Value = flex_med.TextMatrix(flex_med.Row, 15)
dtp_to.Value = flex_med.TextMatrix(flex_med.Row, 16)

Call mrenproc
vprev = flex_med.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
Call flex_titleqt
kl = 0

fr_tit.Enabled = False
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(code) from vendor order by code", Cn, 3, 2
While Not vn.EOF
cbo_vendor.AddItem vn(0)
vn.MoveNext
Wend
 cbo_lookup.AddItem "Item ID"
 cbo_lookup.AddItem "Mfr PartNo."
 cbo_lookup.AddItem "Item Description"
 cbo_lookup.AddItem "Enter Item Manually"
 cbo_lookup.AddItem "Search"


cbo_location.Clear
 Dim sh As New ADODB.Recordset
 If sh.State Then sh.Close
 sh.Open "select DISTINCT(location) from shipping order by location", Cn, 3, 2
 While Not sh.EOF
 cbo_location.AddItem sh(0)
 sh.MoveNext
 Wend
 sh.Close

cbo_project.Clear
Dim prj As New ADODB.Recordset
If prj.State Then prj.Close
prj.Open "select DISTINCT(proj_key),proj_title from projectmaster  ", Cn, 3, 2
While Not prj.EOF
cbo_project.AddItem prj(0) & "  -  " & prj(1)
prj.MoveNext
Wend
prj.Close


prj.Open "select DISTINCT(mtcode),mtdesc from materialtype order by mtcode", Cn, 3, 2
While Not prj.EOF
cbo_materialtype.AddItem prj(0) & "  -  " & prj(1)
prj.MoveNext
Wend
prj.Close


'Call itemvalues
Call clearitems
End Sub
Public Sub flex_titleqt()
On Error Resume Next

flex_med.Rows = 1
   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 300
        .TextMatrix(0, 2) = "Status"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "ItemId"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Mfr.Ref Code"
        .ColWidth(4) = 1200
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Material"
        .ColWidth(5) = 5000
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Qty"
        .ColWidth(6) = 600
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "UOM"
        .ColWidth(7) = 600
        .ColAlignment(7) = 0
        
        .TextMatrix(0, 8) = "Unit Rate"
        .ColWidth(8) = 800
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "Curr"
        .ColWidth(9) = 600
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Xchg Rate"
        .ColWidth(10) = 600
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "Amount"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0


        .TextMatrix(0, 12) = "ReqDate"
        .ColWidth(12) = 800
        .ColAlignment(12) = 0
        .TextMatrix(0, 13) = "Promised Date"
        .ColWidth(13) = 1200
        .ColAlignment(13) = 0
        .TextMatrix(0, 17) = "Remarks"
        .ColWidth(17) = 1200
        .ColAlignment(17) = 0
        
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


Private Sub txt_unitrate_Change()
On Error Resume Next
txt_amount.Text = CDbl(txt_qty.Text) * CDbl(txt_unitrate.Text)
End Sub

Private Sub txt_unitrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_amount.Text = CDbl(txt_qty.Text) * CDbl(txt_unitrate.Text)
Command1.Enabled = True
End Sub
Private Sub opt_all_Click()
Dim g As Integer
g = 0

With flex_med
For g = 1 To flex_med.Rows - 1
chk_app(g).Value = 1
flex_med.TextMatrix(g, 2) = cbo_astatus.Text
Next
End With
opt_all.Value = False
End Sub

Private Sub opt_ind_Click()
Dim w As Integer
w = 0

With flex_med
For w = 1 To flex_med.Rows - 1
chk_app(w).Value = 0
flex_med.TextMatrix(w, 2) = cbo_astatus.Text
Next
End With
End Sub
Private Sub chk_app_Click(Index As Integer)

If opt_all.Value = False Then
If frm_quotation.gg <> 0 Then
If chk_app(Index) = 1 Then
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = cbo_astatus.Text
Else
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = ""
 
End If
End If


frm_quotation.gg = 1
End If
 
End Sub

Public Sub quotationupdate()
Dim qtid As Double
qtid = 0
'            Cn.Execute "delete from quotation where qno='" & txt_account.Text & "'"
'            Cn.Execute "delete from quotationdetails where qno='" & txt_account.Text & "'"
            Dim upqt As New ADODB.Recordset
            If upqt.State Then upqt.Close
            upqt.Open "select * from quotation where q_id =" & flex_med.TextMatrix(flex_med.Row, 19), Cn, 3, 2
            upqt!rfqno = flex_med.TextMatrix(flex_med.Row, 18)
            upqt!qno = txt_account.Text
            upqt!vendor = cbo_vendor.Text
            upqt!contactperson = "-"
            upqt!qdate = dtp_qt.Value
            upqt!toperson = cbo_toperson.Text
            upqt!desig = txt_desig.Text
            upqt!dept = txt_dept.Text
            upqt!Mode = cbo_mode.Text
            upqt!refno = txt_refno.Text
            upqt!oref = txt_oref.Text
            upqt!yref = txt_yref.Text
            upqt!Notes = txt_notes.Text
            upqt!Status = cbo_astatus.Text
            upqt.Update
            
            upqt.Close
            
            
         Dim qd As New ADODB.Recordset
         If qd.State Then qd.Close
         qd.Open "select * from quotationdetails where qtid =" & flex_med.TextMatrix(flex_med.Row, 19), Cn, 3, 2
                      
        Dim j As Integer

        j = 1
        For j = 1 To flex_med.Rows - 1
                        
                        
                        qd!qno = txt_account.Text
                        qd!rfqno = flex_med.TextMatrix(j, 18)
                        qd!Status = flex_med.TextMatrix(j, 2)
                        qd!itemid = flex_med.TextMatrix(j, 3)
                        qd!mrefcode = flex_med.TextMatrix(j, 4)
                        qd!material = flex_med.TextMatrix(j, 5)
                        qd!qty = flex_med.TextMatrix(j, 6)
                        qd!uom = flex_med.TextMatrix(j, 7)
                        qd!unitrate = flex_med.TextMatrix(j, 8)
                        qd!Currency = flex_med.TextMatrix(j, 9)
                        qd!xchg = flex_med.TextMatrix(j, 10)
                        qd!amount = flex_med.TextMatrix(j, 11)
                        qd!reqdate = flex_med.TextMatrix(j, 12)
                        qd!promisedate = flex_med.TextMatrix(j, 13)
                        qd!remarks = flex_med.TextMatrix(j, 17)
                        qd!mtype = flex_med.TextMatrix(j, 14)
                        qd!fromdt = flex_med.TextMatrix(j, 15)
                        qd!todt = flex_med.TextMatrix(j, 16)
                        qd!prid = flex_med.TextMatrix(j, 20)
                        qd.Update
                        qd.MoveNext
        Next j
                    
            
         
End Sub



Public Sub clearitems()
cbo_astatus.Text = ""
cbo_lookup.Text = ""
cbo_category.Text = ""
txt_qty.Text = ""
cbo_uom.Text = ""
txt_unitrate.Text = ""
cbo_curr.Text = ""
txt_amount.Text = ""
txt_xchg.Text = ""
txt_notes.Text = ""
cbo_materialtype.Text = ""
End Sub

Public Sub itemvalues()
'--------------------
On Error Resume Next

'--------------------
If chk_qty.Value = 1 And chk_item.Value = 0 Then
cbo_astatus.Text = "NC, Qty"
cbo_lookup.Enabled = False
cbo_materialtype.Enabled = False
frame_ms.Visible = False
cbo_category.Enabled = False
txt_qty.Enabled = True
cbo_uom.Enabled = False
txt_unitrate.Enabled = True
cbo_curr.Enabled = True
dtp_pdate.Enabled = True
ElseIf chk_qty.Value = 1 And chk_item.Value = 1 Then
cbo_astatus.Text = "NC, Item/Qty"
cbo_lookup.Enabled = True
cbo_materialtype.Enabled = True
frame_ms.Visible = False
cbo_category.Enabled = True
txt_qty.Enabled = True
cbo_uom.Enabled = True
txt_unitrate.Enabled = True
cbo_curr.Enabled = True
dtp_pdate.Enabled = True
ElseIf chk_qty.Value = 0 And chk_item.Value = 1 Then
cbo_astatus.Text = "NC, Item"
cbo_lookup.Enabled = True
cbo_materialtype.Enabled = True
frame_ms.Visible = False
cbo_category.Enabled = True
txt_qty.Enabled = False
cbo_uom.Enabled = True
txt_unitrate.Enabled = True
cbo_curr.Enabled = True
dtp_pdate.Enabled = True
ElseIf chk_qty.Value = 0 And chk_item.Value = 0 Then
cbo_astatus.Text = "Comply"
cbo_lookup.Enabled = False
cbo_materialtype.Enabled = False
frame_ms.Visible = False
cbo_category.Enabled = False
txt_qty.Enabled = False
cbo_uom.Enabled = False
txt_unitrate.Enabled = True
cbo_curr.Enabled = True
dtp_pdate.Enabled = True

End If

End Sub


Public Function assadmailservicequote()
On Error Resume Next
Set oSMTP = New OSSMTP.SMTPSession
'-----------------------------
Dim ct As Double
ct = 0
Dim usid As New ADODB.Recordset
If usid.State Then usid.Close
usid.Open "select * from userid where a_userid ='" & main.Label2.Caption & "' ", Cn, 3, 2
If Not usid.EOF Then

Dim mset As New ADODB.Recordset
If mset.State Then mset.Close
mset.Open "select * from mailsettings", Cn, 3, 2
If Not mset.EOF Then
With oSMTP
    'connection
    .Server = "cpxeon.crest.com.my"
    'message
    .MailFrom = usid!a_email
 Dim usd1 As New ADODB.Recordset
 If usd1.State Then usd1.Close
 usd1.Open "select u.a_email,u.a_name,ur.tforms from userrights ur , userid u where u.a_name=ur.u_name ", Cn, 3, 2
While Not usd1.EOF
If Mid(usd1(2), 8, 1) = 1 Then
.SendTo = usd1(0)
        StrDesc = ""
        .MessageSubject = "Quotation No: " & txt_account.Text & "    " & "For RFQ: " & flex_med.TextMatrix(1, 18)
        StrDesc = "You have received Quotation : " & txt_account.Text & "   " & " For RFQ: " & flex_med.TextMatrix(1, 18)
        StrDesc = StrDesc & vbNewLine & ", Quotation Date is: " & Format(dtp_qt.Value, "dd/MMM/yyyy") & "   " & "Submitted by Vendor  :" & cbo_vendor.Text & ""
        StrItem = ""
        ct = 0
        For ct = 1 To flex_med.Rows - 1
        StrItem = StrItem & vbNewLine & "  Item " & ct & ":" & flex_med.TextMatrix(ct, 5) & "   " & "Qty: " & flex_med.TextMatrix(ct, 6) & "   " & " UOM: " & flex_med.TextMatrix(ct, 7) & "    " & "UnitRate: " & flex_med.TextMatrix(ct, 8)
        StrItem = StrItem & vbNewLine & "  Currency " & ":" & flex_med.TextMatrix(ct, 9) & "   " & "Xchg: " & flex_med.TextMatrix(ct, 10) & "   " & " Amount: " & flex_med.TextMatrix(ct, 11) & "    " & "Reqd Date: " & flex_med.TextMatrix(ct, 12) & "    " & "Promised Delivery Date: " & flex_med.TextMatrix(ct, 13) & vbNewLine & vbNewLine
        Next ct
        .MessageText = StrDesc & vbNewLine & StrItem
        .SendEmail
End If
usd1.MoveNext
Wend

End With
End If
End If

MsgBox "Mail sent Successfully"


End Function

Private Sub txt_xchg_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
