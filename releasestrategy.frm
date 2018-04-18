VERSION 5.00
Begin VB.Form releasestrategy 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Release Strategy"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   10620
   Begin VB.ComboBox cbo_expensetype 
      Height          =   315
      Left            =   3840
      TabIndex        =   55
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txt_amtmax 
      Height          =   285
      Left            =   8160
      TabIndex        =   53
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txt_amt 
      Height          =   285
      Left            =   6000
      TabIndex        =   51
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Default"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Frame fr_prpo 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1920
      TabIndex        =   47
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton opt_po 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO"
         Height          =   375
         Left            =   2400
         TabIndex        =   49
         Top             =   120
         Width           =   975
      End
      Begin VB.OptionButton opt_msr 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MSR"
         Height          =   375
         Left            =   600
         TabIndex        =   48
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   10455
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   46
         Top             =   3360
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   8
         Left            =   2040
         TabIndex        =   45
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   8
         Left            =   4560
         TabIndex        =   44
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   8
         Left            =   5640
         TabIndex        =   43
         Top             =   3360
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   7
         Left            =   2040
         TabIndex        =   41
         Top             =   3000
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   7
         Left            =   4560
         TabIndex        =   40
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   7
         Left            =   5640
         TabIndex        =   39
         Top             =   3000
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   33
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   0
         ItemData        =   "releasestrategy.frx":0000
         Left            =   4560
         List            =   "releasestrategy.frx":0002
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   0
         Left            =   5640
         TabIndex        =   31
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   29
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   1
         Left            =   5640
         TabIndex        =   27
         Top             =   840
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   2
         Left            =   2040
         TabIndex        =   25
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   2
         Left            =   4560
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   2
         Left            =   5640
         TabIndex        =   23
         Top             =   1200
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   21
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   3
         Left            =   4560
         TabIndex        =   20
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   3
         Left            =   5640
         TabIndex        =   19
         Top             =   1560
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   17
         Top             =   1920
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   4
         Left            =   4560
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   4
         Left            =   5640
         TabIndex        =   15
         Top             =   1920
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   5
         Left            =   2040
         TabIndex        =   13
         Top             =   2280
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   5
         Left            =   4560
         TabIndex        =   12
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   5
         Left            =   5640
         TabIndex        =   11
         Top             =   2280
         Width           =   4695
      End
      Begin VB.ComboBox cbo_dept 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox cbo_desig 
         Height          =   315
         Index           =   6
         Left            =   2040
         TabIndex        =   9
         Top             =   2640
         Width           =   2535
      End
      Begin VB.ComboBox cbo_rcode 
         Height          =   315
         Index           =   6
         Left            =   4560
         TabIndex        =   8
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_rname 
         Height          =   285
         Index           =   6
         Left            =   5640
         TabIndex        =   7
         Top             =   2640
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Designation"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   2040
         TabIndex        =   37
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Rls Code"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   36
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Rls Name"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   5640
         TabIndex        =   35
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.TextBox txt_name 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txt_code 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txt_desc 
      Height          =   645
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   5775
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Type"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3840
      TabIndex        =   56
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label lblamtmax 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Max Amount(RM)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   8160
      TabIndex        =   54
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label lblamt 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Min Amount (RM)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   6000
      TabIndex        =   52
      Top             =   1680
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   4
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "releasestrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub cbo_dept_Click(Index As Integer)
Dim rs1 As New ADODB.Recordset
If rs1.State Then rs1.Close
rs1.Open "select DISTINCT(dcode) from designation where ddept= '" & cbo_dept(Index).Text & "' order by dcode", Cn, 3, 2
While Not rs1.EOF
cbo_desig(Index).AddItem rs1(0)
rs1.MoveNext
Wend
rs1.Close
End Sub

Private Sub cbo_rcode_Click(Index As Integer)
Dim rs2 As New ADODB.Recordset
If rs2.State Then rs2.Close
rs2.Open "select DISTINCT(r_name) from releasecodes where r_code= '" & cbo_rcode(Index).Text & "' order by r_name", Cn, 3, 2
If Not rs2.EOF Then
txt_rname(Index).Text = rs2(0)
End If
rs2.Close
End Sub

Private Sub Form_Load()

i = 0
For i = 0 To 8
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select DISTINCT(dcode) from department order by dcode", Cn, 3, 2
While Not rs.EOF
cbo_dept(i).AddItem rs(0)
rs.MoveNext
Wend
Next
rs.Close
i = 0
For i = 0 To 8
rs.Open "select DISTINCT(r_code) from releasecodes order by r_code ", Cn, 3, 2
While Not rs.EOF
cbo_rcode(i).AddItem rs(0)
rs.MoveNext
Wend
rs.Close
Next
txt_amt.Visible = True
lblamt.Visible = True
txt_amtmax.Visible = True
lblamtmax.Visible = True
opt_msr.Value = True

'--------------
cbo_expensetype.Text = "Project Expenses"
cbo_expensetype.AddItem "Project Expenses"
cbo_expensetype.AddItem "Capital Expenses"

End Sub

Private Sub opt_msr_Click()
lblamt.Visible = False
txt_amt.Visible = False
txt_amtmax.Visible = False
lblamtmax.Visible = False
End Sub

Private Sub opt_po_Click()
lblamt.Visible = True
txt_amt.Visible = True
txt_amtmax.Visible = True
lblamtmax.Visible = True
End Sub
