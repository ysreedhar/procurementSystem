VERSION 5.00
Begin VB.Form materiallevel2 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Level2"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Caption         =   "Material Master/Material Category Level-3"
      ForeColor       =   &H8000000E&
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   1800
      Width           =   11775
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00A04729&
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   120
         ScaleHeight     =   4215
         ScaleWidth      =   11535
         TabIndex        =   7
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.TextBox txt_categorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   7815
   End
   Begin VB.TextBox txt_notes 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   840
      Width           =   9375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   3
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "materiallevel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ct1 = 2
vscrollform1.Show
vscrollform1.Left = 0
vscrollform1.Top = 0
 
SetParent vscrollform1.hWnd, materiallevel2.Picture1.hWnd

End Sub

