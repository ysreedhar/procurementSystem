VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchaseorder 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Purchase Order"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "PO Line Items"
      TabPicture(0)   =   "purchaseorder.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Details / Terms and Conditions / Remarks"
      TabPicture(1)   =   "purchaseorder.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   0
         TabIndex        =   32
         Top             =   300
         Width           =   11055
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   960
            Width           =   5295
         End
         Begin VB.TextBox txt_account 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txt_contactperson 
            Height          =   285
            Left            =   5640
            TabIndex        =   39
            Top             =   960
            Width           =   5295
         End
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   38
            Top             =   2400
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame fr_auth 
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization Status"
            ForeColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   120
            TabIndex        =   33
            Top             =   5760
            Width           =   10815
            Begin VB.CommandButton cmd_confirmation 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Verify Approval "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   8880
               Picture         =   "purchaseorder.frx":0038
               Style           =   1  'Graphical
               TabIndex        =   75
               ToolTipText     =   "Click to  Verify Approval"
               Top             =   120
               Width           =   1815
            End
            Begin VB.CommandButton cmd_authorize 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Approve Line Items"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   7080
               Picture         =   "purchaseorder.frx":063C
               Style           =   1  'Graphical
               TabIndex        =   74
               ToolTipText     =   "Click to Authorize MSR"
               Top             =   120
               Width           =   1695
            End
            Begin VB.CommandButton cmd_recommend 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Review"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   5280
               Picture         =   "purchaseorder.frx":0946
               Style           =   1  'Graphical
               TabIndex        =   73
               ToolTipText     =   "Click to Authorize MSR"
               Top             =   120
               Width           =   1695
            End
            Begin VB.ComboBox cbo_astatus 
               Height          =   315
               ItemData        =   "purchaseorder.frx":0C50
               Left            =   120
               List            =   "purchaseorder.frx":0C5D
               TabIndex        =   36
               Text            =   "Pending"
               ToolTipText     =   "Select Authorization Status"
               Top             =   480
               Width           =   2715
            End
            Begin VB.OptionButton opt_ind 
               BackColor       =   &H00FF8080&
               Caption         =   "Clear"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3240
               TabIndex        =   35
               ToolTipText     =   "Click to Uncheck LineItems"
               Top             =   600
               Width           =   975
            End
            Begin VB.OptionButton opt_all 
               BackColor       =   &H00FF8080&
               Caption         =   "Apply All"
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3240
               TabIndex        =   34
               ToolTipText     =   "Click to Authorize all Line Items"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   450
            End
         End
         Begin MSComCtl2.DTPicker dtp_qt 
            Height          =   285
            Left            =   1800
            TabIndex        =   67
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   67108865
            CurrentDate     =   38455
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   4095
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   1
            Cols            =   18
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   12582912
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Height          =   2175
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Visible         =   0   'False
            Width           =   10815
            Begin VB.ComboBox txt_material 
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
               Left            =   4680
               TabIndex        =   51
               Top             =   480
               Width           =   5955
            End
            Begin VB.ComboBox txt_category 
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
               TabIndex        =   50
               Top             =   480
               Width           =   2235
            End
            Begin VB.ComboBox txt_subcategory 
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
               Left            =   2400
               TabIndex        =   49
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox Text2 
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
               TabIndex        =   48
               Top             =   1800
               Width           =   10455
            End
            Begin VB.ComboBox cbo_uom 
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
               Left            =   1279
               TabIndex        =   47
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_qty 
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
               TabIndex        =   46
               Top             =   1200
               Width           =   1080
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
               Left            =   2693
               TabIndex        =   44
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox txt_amount 
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
               Left            =   5945
               TabIndex        =   43
               Top             =   1200
               Width           =   1800
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
               Left            =   3852
               TabIndex        =   42
               Text            =   "RM"
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox txt_xchg 
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
               Left            =   5026
               TabIndex        =   41
               Text            =   "1.00"
               Top             =   1200
               Width           =   840
            End
            Begin MSComCtl2.DTPicker dtp_reqd 
               Height          =   285
               Left            =   7824
               TabIndex        =   45
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   67108865
               CurrentDate     =   38455
            End
            Begin MSComCtl2.DTPicker dtp_pdate 
               Height          =   285
               Left            =   9240
               TabIndex        =   52
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   67108865
               CurrentDate     =   38455
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   7845
               TabIndex        =   64
               Top             =   960
               Width           =   960
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Material Code/ Desc"
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
               Left            =   4695
               TabIndex        =   63
               Top             =   240
               Width           =   1740
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Item Id"
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
               TabIndex        =   62
               Top             =   240
               Width           =   645
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Mfr Ref Code"
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
               Left            =   2400
               TabIndex        =   61
               Top             =   240
               Width           =   1125
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   60
               Top             =   1560
               Width           =   765
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   1275
               TabIndex        =   59
               Top             =   960
               Width           =   390
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   120
               TabIndex        =   58
               Top             =   960
               Width           =   720
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   9240
               TabIndex        =   57
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   2700
               TabIndex        =   56
               Top             =   960
               Width           =   780
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   5940
               TabIndex        =   55
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   3855
               TabIndex        =   54
               Top             =   960
               Width           =   795
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
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
               Left            =   5025
               TabIndex        =   53
               Top             =   960
               Width           =   855
            End
         End
         Begin VB.Label lamt 
            Height          =   255
            Left            =   4200
            TabIndex        =   76
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
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
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "PO  Date"
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
            TabIndex        =   71
            Top             =   120
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "PO No."
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
            TabIndex        =   70
            Top             =   120
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Person"
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
            Left            =   5640
            TabIndex        =   69
            Top             =   720
            Width           =   1305
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   11055
         Begin VB.TextBox txt_remarks 
            Height          =   1335
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   30
            Top             =   5400
            Width           =   10695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Other Details"
            ForeColor       =   &H00C00000&
            Height          =   2415
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   10695
            Begin VB.TextBox txt_desig 
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox cbo_toperson 
               Height          =   315
               Left            =   120
               TabIndex        =   21
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox txt_dept 
               Height          =   285
               Left            =   120
               TabIndex        =   20
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox txt_oref 
               Height          =   285
               Left            =   5880
               TabIndex        =   19
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox cbo_mode 
               Height          =   315
               Left            =   5880
               TabIndex        =   18
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txt_yref 
               Height          =   285
               Left            =   5880
               TabIndex        =   17
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox txt_refno 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7680
               TabIndex        =   16
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Addressed-To-Person"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   4455
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   28
               Top             =   960
               Width           =   4455
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   27
               Top             =   1680
               Width           =   4455
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Mode-of-PO-Sent"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   5880
               TabIndex        =   26
               Top             =   240
               Width           =   1230
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Our Reference"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   5880
               TabIndex        =   25
               Top             =   960
               Width           =   4455
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Your Reference"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   5880
               TabIndex        =   24
               Top             =   1680
               Width           =   4455
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Ref No./ Details"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   7680
               TabIndex        =   23
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Terms / Conditions"
            ForeColor       =   &H00C00000&
            Height          =   2295
            Left            =   120
            TabIndex        =   2
            Top             =   2760
            Width           =   10695
            Begin VB.TextBox Text14 
               Height          =   285
               Left            =   120
               TabIndex        =   8
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox Text15 
               Height          =   285
               Left            =   120
               TabIndex        =   7
               Top             =   1200
               Width           =   10215
            End
            Begin VB.TextBox Text11 
               Height          =   285
               Left            =   120
               TabIndex        =   6
               Top             =   480
               Width           =   10215
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   1320
               TabIndex        =   5
               Top             =   1920
               Width           =   1455
            End
            Begin VB.TextBox Text12 
               Height          =   285
               Left            =   3240
               TabIndex        =   4
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   4440
               TabIndex        =   3
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Validity"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   14
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Delivery"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   960
               Width           =   10215
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Price"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   10215
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Duration"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   1320
               TabIndex        =   11
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Terms"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   3240
               TabIndex        =   10
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00C00000&
               BackStyle       =   0  'Transparent
               Caption         =   "Duration"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   4440
               TabIndex        =   9
               Top             =   1680
               Width           =   1455
            End
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H00C00000&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   5160
            Width           =   10575
         End
      End
   End
End
Attribute VB_Name = "purchaseorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kl As Integer
Public hmt As Double
Private Sub cmd_new_Click()
kl = 1
cbo_name.Clear
End Sub

Private Sub cmd_authorize_Click()
hmt = 0
Dim rls As New ADODB.Recordset
If rls.State Then rls.Close
rls.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='APP' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype='Project Expenses'", Cn, 3, 2
If Not rls.EOF Then

Dim ap As Integer
ap = 0
For ap = 1 To flex_med.Rows - 1
Cn.Execute "update podetails set postatus='" & flex_med.TextMatrix(ap, 2) & "'  where pono='" & txt_account.Text & "' and po_id =" & flex_med.TextMatrix(ap, 0)
Next ap

Cn.Execute "update po set status='" & cbo_astatus.Text & "', personincharge='" & main.Label2.Caption & "'    where pono='" & txt_account.Text & "'"

MsgBox "" & txt_account.Text & "  has been " & cbo_astatus.Text & ""

hmt = 1
Else
hmt = 1
MsgBox "You are not authorised to Approve PO"
End If
Dim rlsa As New ADODB.Recordset
If rlsa.State Then rlsa.Close
rlsa.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='APP' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype='Capital Expenses'", Cn, 3, 2
If Not rlsa.EOF Then


ap = 0
For ap = 1 To flex_med.Rows - 1
Cn.Execute "update podetails set postatus='" & flex_med.TextMatrix(ap, 2) & "'  where pono='" & txt_account.Text & "' and po_id =" & flex_med.TextMatrix(ap, 0)
Next ap

Cn.Execute "update po set status='" & cbo_astatus.Text & "', personincharge='" & main.Label2.Caption & "'    where pono='" & txt_account.Text & "'"

MsgBox "" & txt_account.Text & "  has been " & cbo_astatus.Text & ""

Else
If hmt = 0 Then
MsgBox "You are not authorised to Approve PO"
End If
End If
 

End Sub

Private Sub cmd_confirmation_Click()
hmt = 0
Dim jy As Integer
Dim rls1 As New ADODB.Recordset
If rls1.State Then rls1.Close
rls1.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='VER' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype='Project Expenses'", Cn, 3, 2
If Not rls1.EOF Then

jy = 0

With flex_med
For jy = 1 To flex_med.Rows - 1
chk_app(jy).Enabled = False
Next
End With

Cn.Execute "update po set confirmation ='YES' , approver =' " & main.Label2.Caption & "',adate = '" & Format(Date, "MM/dd/yyyy") & "'  where pono='" & txt_account.Text & "'"
cmd_confirmation.Enabled = False

MsgBox "PO Approval Confirmed"
Call frm_purchaseorder.striptab
Call frm_purchaseorder.flex_itemmodi
hmt = 1
Else
hmt = 1
MsgBox "You are not authorised to Verify PO Approval"
End If
Dim rls1a As New ADODB.Recordset
If rls1a.State Then rls1a.Close
rls1a.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='VER' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype='Capital Expenses'", Cn, 3, 2
If Not rls1a.EOF Then

jy = 0

With flex_med
For jy = 1 To flex_med.Rows - 1
chk_app(jy).Enabled = False
Next
End With

Cn.Execute "update po set confirmation ='YES' , approver =' " & main.Label2.Caption & "',adate = '" & Format(Date, "MM/dd/yyyy") & "'  where pono='" & txt_account.Text & "'"
cmd_confirmation.Enabled = False
MsgBox "PO Approval Confirmed"
Call frm_purchaseorder.striptab
Call frm_purchaseorder.flex_itemmodi
Else
If hmt = 0 Then
MsgBox "You are not authorised to Verify PO Approval"
End If
End If
End Sub

Private Sub cmd_recommend_Click()

hmt = 0
Dim rcm As New ADODB.Recordset
If rcm.State Then rcm.Close
rcm.Open "select DISTINCT(RS.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='REV' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype= 'Project Expenses' ", Cn, 3, 2
If Not rcm.EOF Then

Cn.Execute "update po set reviewer='" & main.Label2.Caption & "' , rdate = '" & Format(Date, "MM/dd/yyyy") & "'  where pono='" & txt_account.Text & "'"
MsgBox "PO Reviewed"
hmt = 1
Else
hmt = 1
MsgBox "You are not Authorised to Review this PO"
End If
Dim rcm1 As New ADODB.Recordset
If rcm1.State Then rcm1.Close
rcm1.Open "select DISTINCT(RS.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='REV' and u.a_userid='" & main.Label2.Caption & "' and rs.amount <= '" & CDbl(lamt.Caption) & "' and rs.amountmax >= '" & CDbl(lamt.Caption) & "' and rs.prpo='PO' and rs.expensetype= 'Capital Expenses' ", Cn, 3, 2
If Not rcm1.EOF Then

Cn.Execute "update po set reviewer='" & main.Label2.Caption & "' , rdate = '" & Format(Date, "MM/dd/yyyy") & "'  where pono='" & txt_account.Text & "'"
MsgBox "PO Reviewed"
Else
If hmt = 0 Then
MsgBox "You are not Authorised to Review this PO"
End If
End If

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
txt_category.Text = flex_med.TextMatrix(flex_med.Row, 4)
txt_subcategory.Text = flex_med.TextMatrix(flex_med.Row, 5)
txt_material.Text = flex_med.TextMatrix(flex_med.Row, 6)
txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 7)
cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 8)
txt_unitrate.Text = flex_med.TextMatrix(flex_med.Row, 9)
cbo_curr.Text = flex_med.TextMatrix(flex_med.Row, 10)
txt_xchg.Text = flex_med.TextMatrix(flex_med.Row, 11)
txt_amount.Text = flex_med.TextMatrix(flex_med.Row, 12)
dtp_reqd.Value = flex_med.TextMatrix(flex_med.Row, 13)
dtp_pdate.Value = flex_med.TextMatrix(flex_med.Row, 14)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 15)




vprev = flex_med.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
Call flex_titleqt
kl = 0



Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(name) from vendor order by name", Cn, 3, 2
While Not vn.EOF
cbo_vendor.AddItem vn(0)
vn.MoveNext
Wend


End Sub
Public Sub flex_titleqt()
On Error Resume Next

flex_med.Rows = 1
   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 300
        .TextMatrix(0, 2) = "PO Status"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "MSE Status"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "ItemId"
        .ColWidth(4) = 1200
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
        
        .TextMatrix(0, 9) = "Unit Rate"
        .ColWidth(9) = 800
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Curr"
        .ColWidth(10) = 600
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "Xchg Rate"
        .ColWidth(11) = 600
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "Amount"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0


        .TextMatrix(0, 13) = "ReqDate"
        .ColWidth(13) = 800
        .ColAlignment(13) = 0
        .TextMatrix(0, 14) = "Promised Date"
        .ColWidth(14) = 1200
        .ColAlignment(14) = 0

        .TextMatrix(0, 15) = "Remarks"
        .ColWidth(15) = 1200
        .ColAlignment(15) = 0
        
        .TextMatrix(0, 16) = "Quotation No"
        .ColWidth(16) = 1200
        .ColAlignment(16) = 0

        .TextMatrix(0, 17) = "MSR No"
        .ColWidth(17) = 1200
        .ColAlignment(17) = 0
        
        
         
    End With
End Sub

Private Sub txt_unitrate_Change()
On Error Resume Next
txt_amount.Text = CDbl(txt_qty.Text) * CDbl(txt_unitrate.Text)
End Sub

Private Sub txt_unitrate_KeyPress(KeyAscii As Integer)
On Error Resume Next
txt_amount.Text = CDbl(txt_qty.Text) * CDbl(txt_unitrate.Text)
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
If frm_purchaseorder.gg <> 0 Then
If chk_app(Index) = 1 Then
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = cbo_astatus.Text
Dim ik As Integer
ik = 0
Dim kk As Integer
kk = 0
If lsid = 0 Then 'list box click
For ik = 1 To flex_med.Rows - 1
     If flex_med.TextMatrix(flex_med.Row, 2) = flex_med.TextMatrix(ik, 2) Then
     If kk = 0 Then
     cbo_astatus.Text = "Approved"
     End If
     Else
     cbo_astatus.Text = "Approved Partially"
     kk = 1
     End If

Next ik
End If
Else
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = ""
 
End If
End If


frm_purchaseorder.gg = 1
End If
 
End Sub
Public Sub poupdate()
Dim ap As Integer
ap = 0
For ap = 1 To flex_med.Rows - 1
Cn.Execute "update podetails set postatus='" & flex_med.TextMatrix(ap, 2) & "'  where pono='" & txt_account.Text & "' and po_id =" & flex_med.TextMatrix(ap, 0)
Next ap

Cn.Execute "update po set status='" & cbo_astatus.Text & "' where pono='" & txt_account.Text & "'"

'''''-----------------------------------------

MsgBox "" & txt_account.Text & "  Has Been " & cbo_astatus.Text & ""
Call frm_purchaseorder.striptab
Call frm_purchaseorder.flex_itemmodi
End Sub
