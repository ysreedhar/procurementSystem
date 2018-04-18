VERSION 5.00
Begin VB.Form jobcharge 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "JobCharge"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txt_projectkey 
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   5415
   End
   Begin VB.TextBox txt_costcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox txt_jobno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   1035
      Width           =   5415
   End
   Begin VB.TextBox txt_jobdesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1725
      Width           =   5415
   End
   Begin VB.TextBox txt_notes 
      Height          =   1215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3120
      Width           =   5535
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Key"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JobCharge No."
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   795
   End
End
Attribute VB_Name = "jobcharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim pk As New ADODB.Recordset
If pk.State Then pk.Close
pk.Open "select DISTINCT(proj_key) from projectmaster order by proj_key", Cn, 3, 2
While Not pk.EOF
txt_projectkey.AddItem pk(0)
pk.MoveNext
Wend
End Sub
