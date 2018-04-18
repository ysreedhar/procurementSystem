VERSION 5.00
Begin VB.Form subcategory 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LEVEL-2(PLWC)"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7860
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_subcategory 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox txt_subcategorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   5895
   End
   Begin VB.ComboBox cbo_categorycode 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txt_notes 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1920
      Width           =   7695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   6615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1920
      TabIndex        =   8
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Level-2  Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Level-1  Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "subcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_categorycode_Click()
Dim cd As New ADODB.Recordset
If cd.State Then cd.Close
cd.Open "select category from category where categorycode ='" & cbo_categorycode.Text & "' ", Cn, 3, 2
If Not cd.EOF Then
txt_category.Text = cd(0)
End If
End Sub

Private Sub Form_Load()
Dim ct As New ADODB.Recordset
If ct.State Then ct.Close
ct.Open "select DISTINCT(categorycode) from category order by categorycode", Cn, 3, 2
While Not ct.EOF
cbo_categorycode.AddItem ct(0)
ct.MoveNext
Wend
End Sub

