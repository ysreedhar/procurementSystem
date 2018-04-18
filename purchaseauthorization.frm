VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchaseauthorization 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MSR Authorization"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "purchaseauthorization.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7215
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12726
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
      TabCaption(0)   =   "MSR Authorization"
      TabPicture(0)   =   "purchaseauthorization.frx":10D5F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Justification / Remarks"
      TabPicture(1)   =   "purchaseauthorization.frx":10D7B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8655
         Left            =   0
         TabIndex        =   2
         Top             =   300
         Width           =   11055
         Begin VB.Frame Frame4 
            BackColor       =   &H80000009&
            Height          =   1575
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   10935
            Begin VB.TextBox txt_expensetype 
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
               Left            =   8520
               TabIndex        =   54
               Top             =   480
               Width           =   2295
            End
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
               TabIndex        =   46
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txt_project 
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
               Left            =   3240
               TabIndex        =   45
               Top             =   480
               Width           =   5175
            End
            Begin VB.TextBox txt_requestor 
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
               Left            =   3240
               TabIndex        =   44
               Top             =   1080
               Width           =   5175
            End
            Begin VB.TextBox txt_department 
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
               TabIndex        =   43
               Top             =   1080
               Width           =   3015
            End
            Begin MSComCtl2.DTPicker dtp_pa 
               Height          =   285
               Left            =   1800
               TabIndex        =   47
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   67043329
               CurrentDate     =   38455
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Expense Type"
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
               Left            =   8520
               TabIndex        =   55
               Top             =   240
               Width           =   1200
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "MSR No."
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
               TabIndex        =   52
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Auth.  Date"
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
               TabIndex        =   51
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Project "
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
               Left            =   3240
               TabIndex        =   50
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Requestor"
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
               Left            =   3240
               TabIndex        =   49
               Top             =   840
               Width           =   870
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
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
               TabIndex        =   48
               Top             =   840
               Width           =   1020
            End
         End
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame fr_auth 
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization Section"
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
            Left            =   0
            TabIndex        =   11
            Top             =   5760
            Width           =   10935
            Begin VB.CommandButton cmd_recommend 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Recommend MSR"
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
               Left            =   5400
               Picture         =   "purchaseauthorization.frx":10D97
               Style           =   1  'Graphical
               TabIndex        =   53
               ToolTipText     =   "Click to Authorize MSR"
               Top             =   195
               Width           =   1695
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
               Left            =   7200
               Picture         =   "purchaseauthorization.frx":110A1
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Click to Authorize MSR"
               Top             =   195
               Width           =   1695
            End
            Begin VB.CommandButton cmd_confirmation 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Confirm MSR Approval"
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
               Left            =   9000
               Picture         =   "purchaseauthorization.frx":113AB
               Style           =   1  'Graphical
               TabIndex        =   16
               ToolTipText     =   "Click to  Confirm Approval"
               Top             =   195
               Width           =   1815
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
               ItemData        =   "purchaseauthorization.frx":119AF
               Left            =   120
               List            =   "purchaseauthorization.frx":119BF
               TabIndex        =   14
               Text            =   "Pending"
               ToolTipText     =   "Select Authorization Status"
               Top             =   480
               Width           =   2715
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
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3120
               TabIndex        =   13
               ToolTipText     =   "Click to Uncheck LineItems"
               Top             =   600
               Width           =   975
            End
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
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   3120
               TabIndex        =   12
               ToolTipText     =   "Click to Authorize all Line Items"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Authorization Status"
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
               TabIndex        =   15
               Top             =   240
               Width           =   1725
            End
         End
         Begin VB.ComboBox cbo_personincharge 
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
            Left            =   240
            TabIndex        =   3
            Top             =   5040
            Visible         =   0   'False
            Width           =   8295
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   4215
            Left            =   0
            TabIndex        =   17
            Top             =   1560
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7435
            _Version        =   393216
            Rows            =   1
            Cols            =   13
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Person Incharge"
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
            Left            =   240
            TabIndex        =   4
            Top             =   4800
            Visible         =   0   'False
            Width           =   1410
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   -75000
         TabIndex        =   1
         Top             =   420
         Width           =   11055
         Begin VB.TextBox txt_notes 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   9
            Top             =   3960
            Width           =   10575
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   3375
            Left            =   0
            TabIndex        =   5
            Top             =   120
            Width           =   10695
            Begin VB.ComboBox cbo_recommendedvendor 
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               TabIndex        =   7
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
               TabIndex        =   6
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
               TabIndex        =   19
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
               TabIndex        =   8
               Top             =   2640
               Width           =   1950
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00A04729&
            Height          =   2175
            Left            =   0
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   10695
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   30
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox txt_remarks 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4260
               TabIndex        =   29
               Top             =   1200
               Width           =   6375
            End
            Begin VB.TextBox txt_category 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   28
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_subcategory 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   27
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_material 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   26
               Top             =   480
               Width           =   5955
            End
            Begin VB.TextBox txt_uom 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1330
               TabIndex        =   25
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_contactperson 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   24
               Top             =   1800
               Width           =   5955
            End
            Begin VB.TextBox txt_location 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   23
               Top             =   1800
               Width           =   2235
            End
            Begin VB.TextBox txt_jobcharge 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   1800
               Width           =   2235
            End
            Begin MSComCtl2.DTPicker dtp_reqd 
               Height          =   285
               Left            =   2795
               TabIndex        =   31
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   67043329
               CurrentDate     =   38455
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1330
               TabIndex        =   40
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4260
               TabIndex        =   39
               Top             =   960
               Width           =   6375
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Mfr.Ref Code"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   38
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemId"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   240
               Width           =   435
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   36
               Top             =   240
               Width           =   5955
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
               TabIndex        =   35
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   34
               Top             =   1560
               Width           =   570
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Jobcharge"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   33
               Top             =   1560
               Width           =   750
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   32
               Top             =   1560
               Width           =   615
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
            TabIndex        =   10
            Top             =   3720
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "purchaseauthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kk As Integer
Public lsid As Integer
Public StrRcm As String
Public StrApp As String
Public StrMsr As String
Private Sub chk_app_Click(Index As Integer)

If opt_all.Value = False Then
If frm_purchaseauthorization.gg <> 0 Then
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
     cbo_astatus.Text = "Partially Approved"
     kk = 1
     End If

Next ik
End If
Else
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = ""
 
End If
End If


frm_purchaseauthorization.gg = 1
End If
' Dim disb As New ADODB.Recordset
' If disb.State Then disb.Close
' disb.Open "select * from purchaserequisition where prno='" & txt_account.Text & "' ", Cn, 3, 2
' If Not disb.EOF Then
' chk_app(Index).Enabled = False
' End If
End Sub

Private Sub cmd_authorize_Click()
Dim rls As New ADODB.Recordset
If rls.State Then rls.Close
rls.Open "select DISTINCT(RS.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='APP' and u.a_userid='" & main.Label2.Caption & "' and rs.prpo='MSR' and rs.expensetype='" & txt_expensetype.Text & "'", Cn, 3, 2
If Not rls.EOF Then

Dim ap As Integer
ap = 0
For ap = 1 To flex_med.Rows - 1
Cn.Execute "update prdetails set status='" & flex_med.TextMatrix(ap, 2) & "'  where prno='" & txt_account.Text & "' and pr_id =" & flex_med.TextMatrix(ap, 0)
Next ap

Cn.Execute "update purchaserequisition set status='" & cbo_astatus.Text & "', personincharge='" & cbo_personincharge.Text & "'    where prno='" & txt_account.Text & "'"

MsgBox "" & txt_account.Text & "  has been " & cbo_astatus.Text & ""
Call frm_purchaseauthorization.striptab
Call frm_purchaseauthorization.flex_itemmodi
Else
MsgBox "You are not authorised to Approve MSR"
End If

End Sub

Private Sub cmd_confirmation_Click()
Dim rls1 As New ADODB.Recordset
If rls1.State Then rls1.Close
rls1.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='APP' and u.a_userid='" & main.Label2.Caption & "' and rs.prpo='MSR' and rs.expensetype='" & txt_expensetype.Text & "'", Cn, 3, 2
If Not rls1.EOF Then
Dim jy As Integer
jy = 0

With flex_med
For jy = 1 To flex_med.Rows - 1
chk_app(jy).Enabled = False
Next
End With

Cn.Execute "update purchaserequisition set confirmation ='YES' , approver ='" & main.Label2.Caption & "',adate = '" & Format(dtp_pa.Value, "MM/dd/yyyy") & "'  where prno='" & txt_account.Text & "'"
cmd_confirmation.Enabled = False
StrRcm = ""
StrMsr = ""
StrRcm = txt_account.Text & " " & "Approved by " & " " & main.Label2.Caption & " " & "on" & " " & Format(dtp_pa.Value, "MM/dd/yyyy") & ""
StrMsr = txt_account.Text
If main.VstrEmail = 0 Then
MsgBox "MSR Approval Confirmed, System is sending mail to update MSR status to the Requestor: Kindly wait for confirmation message"
Call assademailserviceauthorization
Else
MsgBox "MSR Approval Confirmed"
End If
Else
MsgBox "You are not authorised to Confirm Approval MSR"
End If
 
End Sub

Private Sub cmd_recommend_Click()
Dim rcm As New ADODB.Recordset
If rcm.State Then rcm.Close
rcm.Open "select DISTINCT(rs.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='RCM' and u.a_userid='" & main.Label2.Caption & "' and rs.prpo='MSR' and rs.expensetype='" & txt_expensetype.Text & "'", Cn, 3, 2
If Not rcm.EOF Then
Cn.Execute "update purchaserequisition set recommendor='" & main.Label2.Caption & "',rdate = '" & Format(dtp_pa.Value, "MM/dd/yyyy") & "' where prno='" & txt_account.Text & "'"
StrRcm = txt_account.Text & " " & "Recommended by " & " " & main.Label2.Caption & " " & "on" & " " & Format(dtp_pa.Value, "MM/dd/yyyy") & ""
StrMsr = txt_account.Text
If main.VstrEmail = 0 Then
MsgBox "MSR Recommended, System is sending mail to update MSR status to the Requestor: Kindly wait for confirmation message"

Call assademailserviceauthorization
Else
MsgBox " MSR Recommended"
End If
Else
MsgBox "You are not Authorized to Recommend this MSR"
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

Private Sub flex_med_SelChange()
On Error Resume Next
'back color
 kl = 0
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


vprev = flex_med.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
dtp_pa.Value = Format(Date, "dd/MM/yyyy")
 Me.Top = 2000
 Me.Left = 0
 Me.Height = 6990
 Me.Width = 10980
  
    Call flex_titlepa
    kl = 1
    lsid = 0
    Dim lst As New ADODB.Recordset
lst.Open "select DISTINCT(a_name) from userid where a_userid='" & main.Label2.Caption & "' ", Cn, 3, 2
If Not lst.EOF Then
cbo_personincharge.Text = lst(0)
    
  End If
End Sub
Public Sub flex_titlepa()
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
        .TextMatrix(0, 8) = "ReqDate"
        .ColWidth(8) = 800
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "Remarks"
        .ColWidth(9) = 1200
        .ColAlignment(9) = 0

        .TextMatrix(0, 10) = "Jobcharge"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        
        .TextMatrix(0, 11) = "Work Location"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0

        .TextMatrix(0, 12) = "Stor Loc"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0
         
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rls2 As New ADODB.Recordset
If rls2.State Then rls2.Close
rls2.Open "select DISTINCT(RS.rs_code) from releasestrategy rs, releasedetails rd ,userid u where rs.rs_code=rd.rs_code and rd.rs_desig=u.a_designation and rd.rs_rlcode='APP' and u.a_userid='" & main.Label2.Caption & "'", Cn, 3, 2
If Not rls2.EOF Then
Dim ul As New ADODB.Recordset
If ul.State Then ul.Close
ul.Open "select * from purchaserequisition where confirmation='NO' and prno='" & txt_account.Text & "' ", Cn, 3, 2
If Not ul.EOF Then
ms = MsgBox("You have not Confirmed Approval, Click yes to continue Confirmation", vbYesNo)
If ms = vbYes Then
Dim jy As Integer
jy = 0

With flex_med
For jy = 1 To flex_med.Rows - 1
chk_app(jy).Enabled = False
Next
End With

Cn.Execute "update purchaserequisition set confirmation ='YES' where prno='" & txt_account.Text & "'"
cmd_confirmation.Enabled = False

MsgBox "MSR Confirmed"
End If
End If
Else
End If
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


