VERSION 5.00
Begin VB.Form releasecodes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Release Codes"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_name 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txt_desc 
      Height          =   1005
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
   End
   Begin VB.TextBox txt_code 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code "
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1035
   End
End
Attribute VB_Name = "releasecodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cbo_name.AddItem "RECOMMENDED"
cbo_name.AddItem "APPROVED"
cbo_name.AddItem "REVIEWED"
cbo_name.AddItem "VERIFIED"


End Sub
