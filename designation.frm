VERSION 5.00
Begin VB.Form designation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Designation"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_dept 
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txt_dname 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5895
   End
   Begin VB.TextBox txt_dcode 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code "
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "designation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim dp As New ADODB.Recordset
If dp.State Then dp.Close
dp.Open "select DISTINCT(dcode) from department order by dcode", Cn, 3, 2
While Not dp.EOF
cbo_dept.AddItem dp(0)
dp.MoveNext
Wend
dp.Close
End Sub
