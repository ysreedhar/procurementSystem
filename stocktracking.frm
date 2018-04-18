VERSION 5.00
Begin VB.Form stocktracking 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stock Tracking"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Details"
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goods Remaining"
         Height          =   1815
         Left            =   2160
         TabIndex        =   26
         Top             =   3600
         Width           =   1935
         Begin VB.TextBox txt_uom_gr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_qty_gr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   240
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Re Order Level"
         Height          =   1815
         Left            =   4200
         TabIndex        =   21
         Top             =   3600
         Width           =   1935
         Begin VB.TextBox txt_qty_reorder 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cbo_reorder 
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
            TabIndex        =   22
            Top             =   600
            Width           =   1395
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goods Rejected"
         Height          =   1815
         Left            =   120
         TabIndex        =   17
         Top             =   3600
         Width           =   1935
         Begin VB.TextBox txt_uom_grj 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_qty_grj 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   18
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   240
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goods Issued"
         Height          =   1815
         Left            =   4200
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
         Begin VB.TextBox txt_uom_gi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   32
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_qty_gi 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Goods Received"
         Height          =   1815
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   1935
         Begin VB.TextBox txt_uom_grn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_qty_grn 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   240
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opening Stock"
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1935
         Begin VB.TextBox txt_uom_os 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txt_qty_os 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00A04729&
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   390
         End
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
         TabIndex        =   2
         Top             =   600
         Width           =   6915
      End
      Begin VB.ComboBox cbo_batchno 
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
         TabIndex        =   1
         Top             =   1320
         Width           =   2835
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
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
         TabIndex        =   4
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Material batch"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1200
      End
   End
End
Attribute VB_Name = "stocktracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
