VERSION 5.00
Begin VB.Form materialtype 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Type"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton opt_rental 
      BackColor       =   &H8000000E&
      Caption         =   "Rental"
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox txt_notes 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1800
      Width           =   7695
   End
   Begin VB.TextBox txt_materialtype 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
   End
   Begin VB.TextBox txt_materialtypecode 
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
      Caption         =   "Type Desc"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Type Code"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   630
   End
End
Attribute VB_Name = "materialtype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
opt_rental.Value = False
End Sub
