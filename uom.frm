VERSION 5.00
Begin VB.Form uom 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "UOM"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   8865
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00A04729&
      Height          =   1575
      Left            =   5520
      TabIndex        =   10
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txt_midesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txt_miuom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Minor UOM Description"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Minor UOM"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3255
      Begin VB.TextBox txt_mjuomdesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt_mjuom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Major UOM Description"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1650
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Major UOM"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   810
      End
   End
   Begin VB.TextBox txt_mformula 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4080
      TabIndex        =   3
      Text            =   "0"
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txt_remarks 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1920
      Width           =   8655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Formula (Digits)"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00A04729&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   3480
      TabIndex        =   2
      Top             =   480
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   630
   End
End
Attribute VB_Name = "uom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
