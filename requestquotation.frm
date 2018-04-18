VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form requestquotation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RFQ"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "requestquotation.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "RFQ Details"
      TabPicture(0)   =   "requestquotation.frx":10D5F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "RFQ Terms / Conditions"
      TabPicture(1)   =   "requestquotation.frx":10D7B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other Details(MSR)"
      TabPicture(2)   =   "requestquotation.frx":10D97
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   29
         Top             =   300
         Width           =   11055
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   120
            TabIndex        =   92
            Top             =   5280
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   14
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   91
            Top             =   5280
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   0
            Left            =   360
            TabIndex        =   90
            Top             =   5280
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   15
            Left            =   120
            TabIndex        =   89
            Top             =   5640
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   15
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   88
            Top             =   5640
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   1
            Left            =   360
            TabIndex        =   87
            Top             =   5640
            Width           =   1455
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   86
            Top             =   6000
            Width           =   255
         End
         Begin VB.TextBox txt_terms 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   16
            Left            =   1920
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   6000
            Width           =   8775
         End
         Begin VB.TextBox txt_others 
            BackColor       =   &H00C0FFFF&
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   84
            Top             =   6000
            Width           =   1455
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   57
            Top             =   4920
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Closing Date"
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
            Height          =   285
            Index           =   13
            Left            =   120
            TabIndex        =   56
            Top             =   4920
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   55
            Top             =   4560
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Acceptance"
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
            Height          =   285
            Index           =   12
            Left            =   120
            TabIndex        =   54
            Top             =   4560
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delivery Point"
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
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   4200
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   52
            Top             =   4200
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Delay In Delivery"
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
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   51
            Top             =   3840
            Value           =   1  'Checked
            Width           =   1810
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   50
            Top             =   3840
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Validity"
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
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   49
            Top             =   3480
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   48
            Top             =   3480
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Packing"
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
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   3120
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   46
            Top             =   3120
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(iii)"
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
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Top             =   2760
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   44
            Top             =   2760
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(ii)"
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
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   43
            Top             =   2400
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   42
            Top             =   2400
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "(i)"
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
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   41
            Top             =   2040
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   40
            Top             =   2040
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Payment Terms"
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
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   1680
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   38
            Top             =   1680
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Prices"
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   36
            Top             =   240
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Purchase Order"
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
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   34
            Top             =   1320
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Quality"
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
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   960
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   32
            Top             =   960
            Width           =   8775
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Specification"
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
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txt_terms 
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
            TabIndex        =   30
            Top             =   600
            Width           =   8775
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   21
         Top             =   300
         Width           =   10815
         Begin VB.TextBox txt_notes 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   27
            Top             =   3960
            Width           =   10575
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3375
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   10695
            Begin VB.ComboBox cbo_recommendedvendor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   2880
               Width           =   10395
            End
            Begin VB.TextBox txt_justification 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   2085
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Top             =   360
               Width           =   10455
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00800000&
               BackStyle       =   0  'Transparent
               Caption         =   "Justification / Purpose"
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
               TabIndex        =   26
               Top             =   120
               Width           =   1875
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00800000&
               BackStyle       =   0  'Transparent
               Caption         =   "Recommended Vendor"
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
               TabIndex        =   25
               Top             =   2640
               Width           =   1950
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00A04729&
            Height          =   2175
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Visible         =   0   'False
            Width           =   10695
            Begin VB.TextBox txt_jobcharge 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   67
               Top             =   1800
               Width           =   2235
            End
            Begin VB.TextBox txt_location 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   66
               Top             =   1800
               Width           =   2235
            End
            Begin VB.TextBox txt_contactperson 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   65
               Top             =   1800
               Width           =   5955
            End
            Begin VB.TextBox txt_uom 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1330
               TabIndex        =   64
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_material 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   63
               Top             =   480
               Width           =   5955
            End
            Begin VB.TextBox txt_subcategory 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   62
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_category 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   61
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_remarks 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4260
               TabIndex        =   60
               Top             =   1200
               Width           =   6375
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   59
               Top             =   1200
               Width           =   1080
            End
            Begin MSComCtl2.DTPicker dtp_reqd 
               Height          =   285
               Left            =   2795
               TabIndex        =   68
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   67371009
               CurrentDate     =   38455
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   78
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Jobcharge"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   77
               Top             =   1560
               Width           =   750
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   76
               Top             =   1560
               Width           =   570
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Reqd. Date"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   195
               Index           =   0
               Left            =   2810
               TabIndex        =   75
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   74
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemId"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Mfr.Ref Code"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   72
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4260
               TabIndex        =   71
               Top             =   960
               Width           =   6375
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1330
               TabIndex        =   70
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   69
               Top             =   960
               Width           =   1065
            End
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            TabIndex        =   28
            Top             =   3720
            Width           =   765
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7215
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   10935
         Begin VB.Frame Frame5 
            BackColor       =   &H8000000E&
            Caption         =   "VIEW HEADER"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   79
            Top             =   0
            Width           =   10695
            Begin VB.TextBox txt_account 
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
               Left            =   120
               TabIndex        =   80
               Top             =   480
               Width           =   2295
            End
            Begin MSComCtl2.DTPicker dtp_rfq 
               Height          =   285
               Left            =   2520
               TabIndex        =   81
               Top             =   480
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
               Format          =   67371009
               CurrentDate     =   38455
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "RFQ No."
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
               TabIndex        =   83
               Top             =   240
               Width           =   690
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "RFQ Date"
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
               Left            =   2520
               TabIndex        =   82
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   1200
            TabIndex        =   6
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txt_project 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Visible         =   0   'False
            Width           =   4215
         End
         Begin VB.TextBox txt_requestor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   6960
            TabIndex        =   10
            Top             =   960
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox txt_department 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.Frame fr_auth 
            BackColor       =   &H00FF8080&
            Caption         =   "VENDOR ASSIGNING"
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
            TabIndex        =   7
            Top             =   5640
            Width           =   10695
            Begin VB.ComboBox cbo_vendorgroup 
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
               TabIndex        =   18
               Top             =   480
               Width           =   4095
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "<<  Send E-Mail  >>"
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
               Left            =   8280
               Picture         =   "requestquotation.frx":10DB3
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Click to  Send EMails to Vendors"
               Top             =   240
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.CommandButton cmd_confirmation 
               BackColor       =   &H00FFC0C0&
               Caption         =   "<< Save RFQ >>"
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
               Left            =   6240
               Picture         =   "requestquotation.frx":111F5
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Click to  Save Vendor Group Assignment    "
               Top             =   240
               Width           =   1935
            End
            Begin MSComCtl2.DTPicker dtp_closingdate 
               Height          =   315
               Left            =   4320
               TabIndex        =   19
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
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
               Format          =   67371009
               CurrentDate     =   38455
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Width           =   45
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Closing Date"
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
               Index           =   0
               Left            =   4320
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Vendor Group"
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
               TabIndex        =   8
               Top             =   240
               Width           =   1200
            End
         End
         Begin VB.ComboBox cbo_buyer 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Text            =   "ALL"
            Top             =   240
            Visible         =   0   'False
            Width           =   6735
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   4455
            Left            =   120
            TabIndex        =   2
            Top             =   1200
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   7858
            _Version        =   393216
            Cols            =   18
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
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
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Project "
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Requestor"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   6960
            TabIndex        =   13
            Top             =   720
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Department"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   4440
            TabIndex        =   12
            Top             =   720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Buyer"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.Label lblid 
            Height          =   255
            Left            =   2040
            TabIndex        =   3
            Top             =   1560
            Visible         =   0   'False
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "requestquotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

Private Sub cmd_confirmation_Click()
'Cn.Execute "update rfq set  vendorgroup='" & cbo_vendorgroup.Text & "' , closingdate = " & Format(dtp_closingdate.Value, "dd/MM/yyyy") & " where rfqno='" & txt_account.Text & "'"
ms = MsgBox("Do you want to change Terms & Conditions", vbYesNo)
If ms = vbYes Then
Exit Sub
Else

Dim rf As New ADODB.Recordset
If rf.State Then rf.Close
rf.Open "select * from rfq where rfqno='" & txt_account.Text & "'", Cn, 3, 2
If Not rf.EOF Then
rf!vendorgroup = cbo_vendorgroup.Text
rf!closingdate = Format(dtp_closingdate.Value, "dd/MM/yyyy")
rf!tdate = Now
rf!tuser = main.Label2.Caption
rf.Update
rf.Close
End If
'If frm_rfq.flg = 1 Then
Cn.Execute "delete from rfqterms where rfqno='" & txt_account.Text & "'"
Dim rft As New ADODB.Recordset
If rft.State Then rft.Close
rft.Open "select * from rfqterms", Cn, 3, 2
Dim inc As Integer
inc = 0
For inc = 0 To 16
If inc > 13 Then
If txt_others(inc - 14).Text <> "" Then

rft.AddNew
rft!rfqno = txt_account.Text
rft!terms = txt_others(inc - 14).Text
rft!termsdesc = txt_terms(inc).Text
If Check1(inc).Value = 1 Then
rft!chq = "Yes"
Else
rft!chq = "No"
End If
rft.Update
rft.MoveNext
End If
Else
rft.AddNew
rft!rfqno = txt_account.Text
rft!terms = Check1(inc).Caption
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
Call quotationupdate
frm_rfq.striptab
frm_rfq.flex_itemmodi
End If
If main.VstrEmail = 0 Then
MsgBox "RFQ Generated Successfully, system is sending mail to  the selected Vendors: Kindly wait for confirmation message"
Command1_Click
Else
MsgBox "RFQ Generated Successfully"
End If
End Sub

Private Sub Command1_Click()
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
usd1.Open "select v.email,v.name from vendor v, vendorgroup vg where v.name=vg.vendor and vg.vgroup='" & cbo_vendorgroup.Text & "' ", Cn, 3, 2
While Not usd1.EOF
.SendTo = usd1(0)
        .MessageSubject = "RFQ No: " & txt_account.Text
        StrDesc = "You have received an RFQ No : " & txt_account.Text & "   "
        StrDesc = StrDesc & " " & ", Closing Date is: " & Format(dtp_closingdate.Value, "dd/MMM/yyyy") & " "
        ct = 0
          Dim ct1 As Integer
              ct1 = 0
        For ct = 1 To flex_med.Rows - 1
        ct1 = 0
        ct1 = ct + 1
        StrItem = StrItem & vbNewLine & "  ITEM" & ct1 & ":" & flex_med.TextMatrix(ct, 6) & "   " & "Qty: " & flex_med.TextMatrix(ct, 7) & "   " & " UOM: " & flex_med.TextMatrix(ct, 8) & "    " & "Req Date: " & flex_med.TextMatrix(ct, 9)
        Next ct
        .MessageText = StrDesc & vbNewLine & StrItem
        .SendEmail
 
usd1.MoveNext
Wend

End With
End If
End If

MsgBox " Email Sent Successfully"




'-----------------------------
End Sub

Private Sub flex_med_Click()
Call flxcolor
End Sub
Private Sub flex_med_SelChange()
Call flxcolor
End Sub
Private Sub Form_Load()
On Error Resume Next
dtp_rfq.Value = Format(Date, "dd/MM/yyyy")
dtp_closingdate.Value = Format(Date, "dd/MM/yyyy")
'-------------
Set poSendMail = New clsSendMail
'-------------
flex_titlerfq
 Me.Top = 2000
 Me.Left = 0
 Me.Height = 6990
 Me.Width = 10980
 Call termscond
 End Sub
Public Sub flex_titlerfq()
On Error Resume Next
flex_med.Rows = 1
   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .TextMatrix(0, 2) = "Buyer"
        .ColWidth(2) = 0
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "MSR Status"
        .ColWidth(3) = 0
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
Public Sub flxcolor()
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

If flex_med.Row <> 0 Then
'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1
End If
 
cbo_buyer.Text = flex_med.TextMatrix(flex_med.Row, 2)

vprev = flex_med.Row

End Sub

Private Sub opt_all_Click()

End Sub

Private Sub Text12_Change()

End Sub

Public Sub termscond()
On Error Resume Next
Dim strprices As String
strprices = ""
strprices = "Shall be in Ringgit Malaysia (RM), remain firm with no escalation, inclusive of the government taxes and duties, clearances,  with holding taxes, transportation, storage, packing, handling, insurance and other related charges, DDP TLO Sdn Bhd at "
strprices = strprices & frm_rfq.strprices1 & ". Delivery is considered as door to door basis. "
If Check1(0) Then
txt_terms(0).Text = strprices
End If
If Check1(1) Then
txt_terms(1).Text = "Items shall be as specified in Attachment 1, free from all defects in manufacturer's design, workmanship and materials. Any alternative and equivalent offer shall be printed clearly in the proposal/quotation with the cataloques/brochures attached."
End If
If Check1(2) Then
txt_terms(2).Text = "All items shall be new, current or recent production and be merchantability quality. All items shall also be of high quality suitable for Oil & Gas environment."
End If
If Check1(3) Then
txt_terms(3).Text = "Purchase Order shall be issued for either the whole of the above items or part of it, where prices will remain unchanged."
End If
If Check1(4) Then
txt_terms(4).Text = "Shall be  45 days   upon  received  of  undisputed invoices by TLO's Account Department subject  to submission  of  the following documents :"
End If
If Check1(5) Then
txt_terms(5).Text = "The Delivery Order (DO) signed and stamped by TLO authorized personnel at the delivery point. The DO shall be clearly marked with the MSR and the Purchase Order (PO) number."
End If
If Check1(6) Then
txt_terms(6).Text = "The copy of Good Receive Note (GRN)."
End If
If Check1(7) Then
txt_terms(7).Text = "Other relevant documents or certification of the product/services (if applicable/requested by TLO), for e.g Mill Certificates, Notice to Invoice, Work Completion Notice, etc."
End If
If Check1(8) Then
txt_terms(8).Text = "All items shall be properly packed and clearly marked with the MSR No. Loose items are not accepted. All delivery packing list shall be clearly marked with the DO and PO number."
End If
If Check1(9) Then
txt_terms(9).Text = "The proposal/quotation shall be valid for 30 days from the closing date of the RFQ."
End If
If Check1(10) Then
txt_terms(10).Text = "Should Vendor failed to deliver the Goods within agreed delivery date and place, TLO has the right to recover the following Liquidated Damages from Vendor. Liquidated Damage : 2% of the price per day beyond agreed date with a limitation of maximum 20% of the price."
End If

Dim ws As New ADODB.Recordset
If ws.State Then ws.Close
ws.Open "select * from shipping where location='" & frm_rfq.strprices1 & "' ", Cn, 3, 2
If Not ws.EOF Then
If Check1(11) Then
txt_terms(11).Text = ws!Address
End If
End If

If Check1(12) Then
txt_terms(12).Text = "Acceptance of all items subject to the final acceptance by the end-user. All non-conformance items rejected shall be returned and replaced at Vendor's expenses."
End If

If Check1(13) Then
txt_terms(13).Text = CStr(dtp_closingdate.Value)
End If

End Sub

Private Sub SSTab1_DblClick()
Call termscond
End Sub
Public Sub quotationupdate()
Cn.Execute "delete from quotation where rfqno='" & txt_account.Text & "'"
Cn.Execute "delete from quotationdetails where rfqno='" & txt_account.Text & "'"
Dim vd As New ADODB.Recordset
If vd.State Then vd.Close
vd.Open "select DISTINCT(vendor) from vendorgroup where vgroup='" & cbo_vendorgroup.Text & "' ", Cn, 3, 2
While Not vd.EOF

Dim upq As New ADODB.Recordset
If upq.State Then upq.Close
upq.Open "select * from rfq where rfqno='" & txt_account.Text & "' ", Cn, 3, 2
If Not upq.EOF Then

            Dim qtid As Double
            qtid = 0
            Dim upqt As New ADODB.Recordset
            If upqt.State Then upqt.Close
            upqt.Open "select * from quotation", Cn, 3, 2
            upqt.AddNew
            upqt!vendor = vd(0)
            upqt!contactperson = "-"
            upqt!rfqno = upq!rfqno
            upqt!qdate = Format(Date, "dd/MM/yyyy")
            upqt.Update
            upqt.Close
            upqt.Open "select MAX(q_id) from quotation where rfqno='" & txt_account.Text & "' ", Cn, 3, 2
            If Not upqt.EOF Then
             qtid = upqt(0)
            End If
            
End If

Dim prs As New ADODB.Recordset
If prs.State Then prs.Close
prs.Open "select * from prdetails where rfqno='" & txt_account.Text & "' order by pr_id", Cn, 3, 2
    While Not prs.EOF
         Dim qd As New ADODB.Recordset
         If qd.State Then qd.Close
         qd.Open "select * from quotationdetails", Cn, 3, 2
            qd.AddNew
            qd!itemid = prs!itemid
            qd!mrefcode = prs!mrefcode
            qd!material = prs!material
            qd!qty = prs!qty
            qd!uom = prs!uom
            qd!reqdate = prs!reqdate
            qd!remarks = prs!remarks
            qd!rfqno = txt_account.Text
            qd!promisedate = prs!reqdate
            qd!qtid = qtid
            qd!vendor = vd(0)
            qd!Location = prs!Location
            qd!prno = prs!prno
            qd!mtype = prs!mtype
            qd!fromdt = prs!fromdt
            qd!todt = prs!todt
            qd!prid = prs!pr_id
            qd.Update
            
         prs.MoveNext
    Wend



vd.MoveNext
Wend




End Sub

Private Sub txt_account_Change()
cbo_vendorgroup.Clear
Dim vg As New ADODB.Recordset
 If vg.State Then vg.Close
 vg.Open "select DISTINCT(vgroup) from vendorgroup v , rfq r where v.vprno=r.prno and r.rfqno='" & txt_account.Text & "' order by vgroup", Cn, 3, 2
 While Not vg.EOF
 cbo_vendorgroup.AddItem vg(0)
 vg.MoveNext
 Wend
vg.Close
End Sub
