VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vendor 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vendor"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Vendor Sub-Range"
      Top             =   0
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Basic Data"
      TabPicture(0)   =   "vendor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Licensing"
      TabPicture(1)   =   "vendor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Activity Responsible"
      TabPicture(2)   =   "vendor.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bank Details"
      TabPicture(3)   =   "vendor.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Terms of Payment"
      TabPicture(4)   =   "vendor.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame14"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame14 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   -75000
         TabIndex        =   74
         Top             =   300
         Width           =   10815
         Begin VB.Frame Frame17 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Interest"
            Height          =   3495
            Left            =   4680
            TabIndex        =   87
            Top             =   360
            Width           =   4455
            Begin VB.TextBox txt_iper3 
               Height          =   285
               Left            =   1200
               TabIndex        =   92
               Text            =   "0"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox txt_idays1 
               Height          =   285
               Left            =   120
               TabIndex        =   91
               Text            =   "0"
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txt_iper2 
               Height          =   285
               Left            =   1200
               TabIndex        =   90
               Text            =   "0"
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_idays2 
               Height          =   285
               Left            =   120
               TabIndex        =   89
               Text            =   "0"
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_idays3 
               Height          =   285
               Left            =   120
               TabIndex        =   88
               Text            =   "0"
               Top             =   1920
               Width           =   855
            End
            Begin VB.Label Label47 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Interest(%)"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   1200
               TabIndex        =   98
               Top             =   1680
               Width           =   735
            End
            Begin VB.Label Label46 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   97
               Top             =   480
               Width           =   360
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Interest(%)"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   1200
               TabIndex        =   96
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   95
               Top             =   1080
               Width           =   360
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   94
               Top             =   1680
               Width           =   360
            End
            Begin VB.Label Label40 
               BackColor       =   &H00C0E0FF&
               Caption         =   $"vendor.frx":008C
               Height          =   2895
               Left            =   2280
               TabIndex        =   93
               Top             =   360
               Width           =   2055
            End
         End
         Begin VB.Frame Frame16 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Discount"
            Height          =   3495
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   4455
            Begin VB.TextBox txt_ddays3 
               Height          =   285
               Left            =   120
               TabIndex        =   84
               Text            =   "0"
               Top             =   1920
               Width           =   855
            End
            Begin VB.TextBox txt_ddays2 
               Height          =   285
               Left            =   120
               TabIndex        =   81
               Text            =   "0"
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_dper2 
               Height          =   285
               Left            =   1200
               TabIndex        =   80
               Text            =   "0"
               Top             =   1320
               Width           =   855
            End
            Begin VB.TextBox txt_ddays1 
               Height          =   285
               Left            =   120
               TabIndex        =   77
               Text            =   "0"
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txt_dper1 
               Height          =   285
               Left            =   1200
               TabIndex        =   76
               Text            =   "0"
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label42 
               BackColor       =   &H00C0E0FF&
               Caption         =   $"vendor.frx":0175
               Height          =   2895
               Left            =   2280
               TabIndex        =   86
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   85
               Top             =   1680
               Width           =   360
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   1080
               Width           =   360
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount(%)"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   1200
               TabIndex        =   82
               Top             =   1080
               Width           =   840
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Days"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   79
               Top             =   480
               Width           =   360
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Discount(%)"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   1200
               TabIndex        =   78
               Top             =   480
               Width           =   840
            End
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   -75000
         TabIndex        =   52
         Top             =   300
         Width           =   10815
         Begin VB.Frame Frame13 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alternative 2"
            Height          =   1095
            Left            =   120
            TabIndex        =   67
            Top             =   2640
            Width           =   10455
            Begin VB.TextBox txt_bankaccount3 
               Height          =   285
               Left            =   3240
               TabIndex        =   70
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txt_bank3 
               Height          =   285
               Left            =   120
               TabIndex        =   69
               Top             =   720
               Width           =   3015
            End
            Begin VB.TextBox txt_branch3 
               Height          =   285
               Left            =   5160
               TabIndex        =   68
               Top             =   720
               Width           =   5175
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Account #"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   3240
               TabIndex        =   73
               Top             =   480
               Width           =   750
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   72
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Place/Branch"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   5160
               TabIndex        =   71
               Top             =   480
               Width           =   990
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Alternative 1"
            Height          =   1095
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   10455
            Begin VB.TextBox txt_bankaccount2 
               Height          =   285
               Left            =   3240
               TabIndex        =   63
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox txt_bank2 
               Height          =   285
               Left            =   120
               TabIndex        =   62
               Top             =   720
               Width           =   3015
            End
            Begin VB.TextBox txt_branch2 
               Height          =   285
               Left            =   5160
               TabIndex        =   61
               Top             =   720
               Width           =   5175
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Account #"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   3240
               TabIndex        =   66
               Top             =   480
               Width           =   750
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   65
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Place/Branch"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   5160
               TabIndex        =   64
               Top             =   480
               Width           =   990
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Primary"
            Height          =   1095
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   10455
            Begin VB.TextBox txt_branch1 
               Height          =   285
               Left            =   5160
               TabIndex        =   58
               Top             =   720
               Width           =   5175
            End
            Begin VB.TextBox txt_bank1 
               Height          =   285
               Left            =   120
               TabIndex        =   55
               Top             =   720
               Width           =   3015
            End
            Begin VB.TextBox txt_bankaccount1 
               Height          =   285
               Left            =   3240
               TabIndex        =   54
               Top             =   720
               Width           =   1815
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Place/Branch"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   5160
               TabIndex        =   59
               Top             =   480
               Width           =   990
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Bank"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   375
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Account #"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   3240
               TabIndex        =   56
               Top             =   480
               Width           =   750
            End
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   -75000
         TabIndex        =   31
         Top             =   300
         Width           =   10815
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "RFQ"
            Height          =   1815
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txt_rpin 
               Height          =   285
               Left            =   120
               TabIndex        =   49
               Top             =   720
               Width           =   4695
            End
            Begin VB.TextBox txt_rdesig 
               Height          =   285
               Left            =   120
               TabIndex        =   48
               Top             =   1320
               Width           =   4695
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Person Incharge"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   51
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   1080
               Width           =   840
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PO"
            Height          =   1815
            Left            =   5160
            TabIndex        =   42
            Top             =   240
            Width           =   4935
            Begin VB.TextBox txt_pdesig 
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   1320
               Width           =   4695
            End
            Begin VB.TextBox txt_ppin 
               Height          =   285
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   4695
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   1080
               Width           =   840
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Person Incharge"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   45
               Top             =   480
               Width           =   1815
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Contract"
            Height          =   1815
            Left            =   120
            TabIndex        =   37
            Top             =   2160
            Width           =   4935
            Begin VB.TextBox txt_cdesig 
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   1320
               Width           =   4695
            End
            Begin VB.TextBox txt_cpin 
               Height          =   285
               Left            =   120
               TabIndex        =   38
               Top             =   720
               Width           =   4695
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   41
               Top             =   1080
               Width           =   840
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Person Incharge"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   40
               Top             =   480
               Width           =   1815
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enquiries"
            Height          =   1815
            Left            =   5160
            TabIndex        =   32
            Top             =   2160
            Width           =   4935
            Begin VB.TextBox txt_edesig 
               Height          =   285
               Left            =   120
               TabIndex        =   34
               Top             =   1320
               Width           =   4695
            End
            Begin VB.TextBox txt_epin 
               Height          =   285
               Left            =   120
               TabIndex        =   33
               Top             =   720
               Width           =   4695
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Designation"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   36
               Top             =   1080
               Width           =   840
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Name of Person Incharge"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   480
               Width           =   1815
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   -75000
         TabIndex        =   6
         Top             =   360
         Width           =   10695
         Begin VB.TextBox txt_notes 
            Height          =   5085
            Left            =   6480
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   15
            Top             =   480
            Width           =   4095
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "PLWC"
            ForeColor       =   &H80000007&
            Height          =   5535
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "Check to Add , Specialized Materials of this Vendor"
            Top             =   0
            Width           =   6255
            Begin MSComCtl2.DTPicker dtp_expiry 
               Height          =   255
               Left            =   4800
               TabIndex        =   27
               Top             =   2640
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               Format          =   67305473
               CurrentDate     =   38579
            End
            Begin VB.ListBox List3 
               Height          =   960
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   10
               Top             =   1560
               Width           =   6015
            End
            Begin VB.ListBox List2 
               Height          =   735
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   9
               Top             =   480
               Width           =   6015
            End
            Begin VB.ListBox List1 
               Height          =   2535
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   8
               Top             =   2880
               Width           =   6015
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Expiry Date"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   3840
               TabIndex        =   28
               Top             =   2640
               Width           =   810
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Level 2"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   1320
               Width           =   525
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Level 1"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   525
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Level 3"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   2640
               Width           =   525
            End
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H80000007&
            Height          =   195
            Left            =   6480
            TabIndex        =   17
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   5415
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   10815
         Begin VB.Frame Frame18 
            BackColor       =   &H8000000E&
            Caption         =   "Contact Information"
            Height          =   3375
            Left            =   240
            TabIndex        =   99
            Top             =   2040
            Width           =   5655
            Begin VB.TextBox txt_website 
               Height          =   285
               Left            =   120
               TabIndex        =   104
               Top             =   2955
               Width           =   5415
            End
            Begin VB.TextBox txt_email 
               Height          =   285
               Left            =   120
               TabIndex        =   103
               Top             =   2355
               Width           =   5415
            End
            Begin VB.TextBox txt_fax 
               Height          =   285
               Left            =   2880
               TabIndex        =   102
               Top             =   1755
               Width           =   2655
            End
            Begin VB.TextBox txt_phone 
               Height          =   285
               Left            =   120
               TabIndex        =   101
               Top             =   1755
               Width           =   2655
            End
            Begin VB.TextBox txt_address 
               Height          =   765
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   100
               Top             =   600
               Width           =   5415
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Website"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   109
               Top             =   2760
               Width           =   2655
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Email"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   108
               Top             =   2160
               Width           =   2655
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Fax"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   2880
               TabIndex        =   107
               Top             =   1560
               Width           =   2655
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   106
               Top             =   1560
               Width           =   2655
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H80000006&
               Height          =   195
               Left            =   120
               TabIndex        =   105
               Top             =   360
               Width           =   5415
            End
         End
         Begin VB.TextBox txt_code 
            Height          =   285
            Left            =   240
            TabIndex        =   29
            Top             =   315
            Width           =   2535
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Status"
            ForeColor       =   &H80000006&
            Height          =   3375
            Left            =   6000
            TabIndex        =   14
            Top             =   2040
            Width           =   4455
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Equity Status( % )"
               ForeColor       =   &H80000007&
               Height          =   1815
               Left            =   240
               TabIndex        =   20
               Top             =   1440
               Width           =   3975
               Begin VB.TextBox txt_fl 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   26
                  Text            =   "0"
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.TextBox txt_nb 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   25
                  Text            =   "0"
                  Top             =   840
                  Width           =   1095
               End
               Begin VB.TextBox txt_bm 
                  Height          =   285
                  Left            =   1200
                  TabIndex        =   24
                  Text            =   "0"
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Label Label17 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "FL"
                  ForeColor       =   &H80000007&
                  Height          =   255
                  Left            =   720
                  TabIndex        =   23
                  Top             =   1320
                  Width           =   495
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "NB"
                  ForeColor       =   &H80000007&
                  Height          =   255
                  Left            =   720
                  TabIndex        =   22
                  Top             =   900
                  Width           =   495
               End
               Begin VB.Label Label15 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "BM"
                  ForeColor       =   &H80000007&
                  Height          =   255
                  Left            =   720
                  TabIndex        =   21
                  Top             =   480
                  Width           =   495
               End
            End
            Begin VB.ComboBox cbo_companystatus 
               Height          =   315
               ItemData        =   "vendor.frx":0262
               Left            =   240
               List            =   "vendor.frx":0275
               TabIndex        =   18
               Top             =   600
               Width           =   3975
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Company Status"
               ForeColor       =   &H80000007&
               Height          =   195
               Left            =   240
               TabIndex        =   19
               Top             =   360
               Width           =   1155
            End
         End
         Begin VB.TextBox txt_regno 
            Height          =   285
            Left            =   2880
            TabIndex        =   3
            Top             =   315
            Width           =   2775
         End
         Begin VB.TextBox txt_name 
            Height          =   285
            Left            =   240
            TabIndex        =   2
            Top             =   960
            Width           =   5415
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor No."
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   120
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   240
            TabIndex        =   5
            Top             =   720
            Width           =   5415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Company Reg No"
            ForeColor       =   &H80000006&
            Height          =   195
            Left            =   2880
            TabIndex        =   4
            Top             =   120
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "vendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dtp_expiry.Value = Format(Date, "dd/MM/yyyy")
Dim ct As New ADODB.Recordset
If ct.State Then ct.Close
ct.Open "select Distinct(categorycode),category from material order by category", Cn, 3, 2
While Not ct.EOF
List2.AddItem ct(0) & "  -  " & ct(1)
ct.MoveNext
Wend
End Sub

Private Sub List1_ItemCheck(Item As Integer)
If frm_vendor.ih = 1 Then
MsgBox "Expiry Date is:" & dtp_expiry.Value & " To change Expiry date ,Select the Date"
dtp_expiry.SetFocus
If List1.Selected(Item) = True Then
List1.List(Item) = List1.List(Item) & "  -  " & Format(dtp_expiry.Value, "dd/MM/yyyy")
End If
End If
End Sub
Private Sub List2_Click()
List3.Clear
Dim i As Integer
i = 0
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
nm = Split(List2.List(i), "  -  ", Len(List2.List(i)), vbTextCompare)
Dim sct As New ADODB.Recordset
If sct.State Then sct.Close
sct.Open "select Distinct(subcategorycode),subcategory from subcategory where categorycode ='" & nm(0) & "' order by subcategory", Cn, 3, 2
While Not sct.EOF
List3.AddItem sct(0) & "  -  " & sct(1)
sct.MoveNext
Wend


End If
Next
End Sub

Private Sub List3_Click()
If frm_vendor.fl = 1 Then
List1.Clear
End If
If frm_vendor.ih = 1 Then
Dim j As Integer
j = 0
Dim p As Integer
p = 0
For j = 0 To List3.ListCount - 1
If List3.Selected(j) = True Then
nm1 = Split(List3.List(j), "  -  ", Len(List3.List(j)), vbTextCompare)
Dim mt As New ADODB.Recordset
If mt.State Then mt.Close
mt.Open "select Distinct(code),name from material where subcategorycode='" & nm1(0) & "' order by name", Cn, 3, 2

While Not mt.EOF
p = 0
For p = 0 To List1.ListCount - 1
If List1.List(p) = mt(0) & "  -  " & mt(1) Then Exit Sub

Next

List1.AddItem mt(0) & "  -  " & mt(1)



mt.MoveNext
Wend
End If
Next
End If
End Sub

