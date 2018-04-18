VERSION 5.00
Begin VB.Form material 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LEVEL-3 (PLWC)"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_description 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2520
      Width           =   8775
   End
   Begin VB.TextBox txt_name 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   6735
   End
   Begin VB.ComboBox cbo_categorycode 
      Height          =   315
      ItemData        =   "material.frx":0000
      Left            =   120
      List            =   "material.frx":0002
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Width           =   6735
   End
   Begin VB.ComboBox cbo_subcategorycode 
      Height          =   315
      ItemData        =   "material.frx":0004
      Left            =   120
      List            =   "material.frx":0006
      TabIndex        =   3
      Top             =   1020
      Width           =   1695
   End
   Begin VB.TextBox txt_subcategory 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   1020
      Width           =   6735
   End
   Begin VB.TextBox txt_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txt_notes 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3720
      Width           =   8775
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Material Name"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   14
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Material Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   6735
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Level-1 Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Level-2 Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Level-3 Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   945
   End
End
Attribute VB_Name = "material"
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
Dim scd As New ADODB.Recordset
If scd.State Then scd.Close
scd.Open "select subcategorycode from subcategory where categorycode ='" & cbo_categorycode.Text & "' ", Cn, 3, 2
While Not scd.EOF
cbo_subcategorycode.AddItem scd(0)
scd.MoveNext
Wend
End Sub

Private Sub cbo_subcategorycode_Click()
Dim scd As New ADODB.Recordset
If scd.State Then scd.Close
scd.Open "select subcategory from subcategory where subcategorycode ='" & cbo_subcategorycode.Text & "' ", Cn, 3, 2
If Not scd.EOF Then
txt_subcategory.Text = scd(0)
End If
End Sub

Private Sub cbo_type_Change()

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

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
