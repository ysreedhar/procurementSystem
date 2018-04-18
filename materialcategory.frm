VERSION 5.00
Begin VB.Form materialcategory 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Category Level-1"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00A04729&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   0
      ScaleHeight     =   4095
      ScaleWidth      =   11085
      TabIndex        =   6
      Top             =   1920
      Width           =   11085
   End
   Begin VB.TextBox txt_notes 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   960
      Width           =   9375
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   7815
   End
   Begin VB.TextBox txt_categorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      TabIndex        =   5
      Top             =   120
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
      Top             =   120
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
      Top             =   720
      Width           =   630
   End
End
Attribute VB_Name = "materialcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ct1 = 1
vscrollform.Show
SetParent vscrollform.hWnd, materialcategory.Picture1.hWnd

End Sub
