VERSION 5.00
Begin VB.Form shipping 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Storage Location"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10005
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_worklocation 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   31
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact 2"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   9855
      Begin VB.TextBox txt_remarks3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   25
         Top             =   960
         Width           =   8775
      End
      Begin VB.TextBox txt_personincharge3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txt_phone3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   20
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_email3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   26
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   24
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   23
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   22
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact 2"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   9855
      Begin VB.TextBox txt_remarks2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   27
         Top             =   960
         Width           =   8775
      End
      Begin VB.TextBox txt_personincharge2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txt_phone2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_email2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   12
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   28
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   17
         Top             =   120
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   16
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   15
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contact 1"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   9855
      Begin VB.TextBox txt_remarks1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   29
         Top             =   960
         Width           =   8775
      End
      Begin VB.TextBox txt_email1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6720
         TabIndex        =   9
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txt_phone1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txt_personincharge1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   30
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6720
         TabIndex        =   10
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4560
         TabIndex        =   8
         Top             =   120
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.TextBox txt_address 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   360
      Width           =   6135
   End
   Begin VB.TextBox txt_location 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Work Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Storage Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "shipping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
cbo_worklocation.Clear
Dim prj As New ADODB.Recordset
If prj.State Then prj.Close
prj.Open "select DISTINCT(workloc) from worklocation  ", Cn, 3, 2
While Not prj.EOF
cbo_worklocation.AddItem prj(0)
prj.MoveNext
Wend
prj.Close
End Sub
