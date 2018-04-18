VERSION 5.00
Begin VB.Form materiallevel3 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Level3"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Caption         =   "Material Master"
      ForeColor       =   &H8000000E&
      Height          =   4695
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   11775
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00A04729&
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   120
         ScaleHeight     =   4335
         ScaleWidth      =   11535
         TabIndex        =   11
         Top             =   240
         Width           =   11535
      End
   End
   Begin VB.ComboBox cbo_type 
      Height          =   315
      Left            =   9000
      TabIndex        =   7
      Top             =   240
      Width           =   2655
   End
   Begin VB.ComboBox cbo_uom 
      Height          =   315
      Left            =   7560
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox txt_notes 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   840
      Width           =   9375
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   5775
   End
   Begin VB.TextBox txt_categorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   9000
      TabIndex        =   9
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "UOM"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   7560
      TabIndex        =   8
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1680
      TabIndex        =   5
      Top             =   0
      Width           =   420
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   630
   End
End
Attribute VB_Name = "materiallevel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ct1 = 3
vscrollform2.Show
vscrollform2.Left = 0
vscrollform2.Top = 0
SetParent vscrollform2.hWnd, materiallevel3.Picture1.hWnd

End Sub


