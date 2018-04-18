VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form quotationauthorization 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quotation Authorization"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   11055
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
      BackColor       =   16744576
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Quot Authorization"
      TabPicture(0)   =   "quotationauthorization.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Other Details / Terms and Conditions / Remarks"
      TabPicture(1)   =   "quotationauthorization.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   7215
         Left            =   -75000
         TabIndex        =   11
         Top             =   300
         Width           =   11055
         Begin VB.TextBox txt_notes 
            Height          =   1335
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   75
            Top             =   5280
            Width           =   10695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Other Details"
            Height          =   2415
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Width           =   10695
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   67
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox Combo10 
               Height          =   315
               Left            =   120
               TabIndex        =   66
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   65
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5880
               TabIndex        =   64
               Top             =   1200
               Width           =   4455
            End
            Begin VB.ComboBox Combo9 
               Height          =   315
               Left            =   5880
               TabIndex        =   63
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox Text9 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5880
               TabIndex        =   62
               Top             =   1920
               Width           =   4455
            End
            Begin VB.TextBox Text10 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7680
               TabIndex        =   61
               Top             =   480
               Width           =   2655
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Addressed-To-Person"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   4455
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Designation"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   73
               Top             =   960
               Width           =   4455
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Department"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   1680
               Width           =   4455
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Mode-Of-QuotReceived"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   5880
               TabIndex        =   71
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Our Reference"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   5880
               TabIndex        =   70
               Top             =   960
               Width           =   4455
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Your Reference"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   5880
               TabIndex        =   69
               Top             =   1680
               Width           =   4455
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Ref No./ Details"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   7680
               TabIndex        =   68
               Top             =   240
               Width           =   2655
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Terms / Conditions"
            Height          =   2295
            Left            =   120
            TabIndex        =   47
            Top             =   2640
            Width           =   10695
            Begin VB.TextBox Text14 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   53
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox Text15 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   52
               Top             =   1200
               Width           =   10215
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   51
               Top             =   480
               Width           =   10215
            End
            Begin VB.ComboBox Combo8 
               Height          =   315
               Left            =   1320
               TabIndex        =   50
               Top             =   1920
               Width           =   1455
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3240
               TabIndex        =   49
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   4440
               TabIndex        =   48
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Validity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   59
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Delivery"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   58
               Top             =   960
               Width           =   10215
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Price"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Width           =   10215
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Duration"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1320
               TabIndex        =   56
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Terms"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   3240
               TabIndex        =   55
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Duration"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4440
               TabIndex        =   54
               Top             =   1680
               Width           =   1455
            End
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Remarks"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   5040
            Width           =   10695
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   7455
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   11055
         Begin VB.ComboBox Combo5 
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   6735
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   6735
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FF8080&
            Caption         =   "Delete"
            Height          =   275
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   2400
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF8080&
            Caption         =   "Save"
            Height          =   275
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   2400
            Width           =   855
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Height          =   2775
            Left            =   120
            TabIndex        =   12
            Top             =   2640
            Width           =   10815
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   1800
               TabIndex        =   40
               Top             =   2400
               Width           =   5355
            End
            Begin VB.ComboBox Combo4 
               Height          =   315
               Left            =   120
               TabIndex        =   39
               Top             =   2400
               Width           =   1635
            End
            Begin VB.ComboBox cbo_name 
               Height          =   315
               Left            =   4680
               TabIndex        =   23
               Top             =   480
               Width           =   5955
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   2235
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   2400
               TabIndex        =   21
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   20
               Top             =   1800
               Width           =   10455
            End
            Begin VB.ComboBox cbo_uom 
               Height          =   315
               Left            =   1290
               TabIndex        =   19
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   18
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2715
               TabIndex        =   16
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6000
               TabIndex        =   15
               Top             =   1200
               Width           =   1800
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   3885
               TabIndex        =   14
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5070
               TabIndex        =   13
               Top             =   1200
               Width           =   840
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   285
               Left            =   7830
               TabIndex        =   17
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   20578305
               CurrentDate     =   38455
            End
            Begin MSComCtl2.DTPicker dtp_expiration 
               Height          =   285
               Left            =   9240
               TabIndex        =   24
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   20578305
               CurrentDate     =   38455
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Authorized Buyer"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   1800
               TabIndex        =   42
               Top             =   2160
               Width           =   5355
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Authorization Status"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   2160
               Width           =   1635
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
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
               Index           =   1
               Left            =   7845
               TabIndex        =   36
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   35
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Sub Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   33
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   1560
               Width           =   10455
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1278
               TabIndex        =   31
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Promised Date"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
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
               Left            =   9240
               TabIndex        =   29
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Unit Rate"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   2706
               TabIndex        =   28
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Amount(RM)"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   6000
               TabIndex        =   27
               Top             =   960
               Width           =   1800
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Currency"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   3885
               TabIndex        =   26
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Exch Rate"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   5070
               TabIndex        =   25
               Top             =   960
               Width           =   840
            End
         End
         Begin VB.TextBox txt_account 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   3360
            TabIndex        =   4
            Top             =   360
            Width           =   4935
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select  Quot No."
            ForeColor       =   &H00000000&
            Height          =   1335
            Left            =   8520
            TabIndex        =   2
            Top             =   0
            Width           =   2415
            Begin VB.ListBox List1 
               Height          =   960
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   3
               Top             =   240
               Width           =   2175
            End
         End
         Begin MSComCtl2.DTPicker dtp_po 
            Height          =   285
            Left            =   1800
            TabIndex        =   6
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   20578305
            CurrentDate     =   38455
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   1575
            Left            =   120
            TabIndex        =   7
            Top             =   5400
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2778
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Vendor"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   840
            Width           =   6735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Contact Person"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   1560
            Width           =   6735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization ID"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization Date"
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
            Index           =   1
            Left            =   1800
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Person Incharge"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   3360
            TabIndex        =   8
            Top             =   120
            Width           =   4935
         End
      End
   End
End
Attribute VB_Name = "quotationauthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

