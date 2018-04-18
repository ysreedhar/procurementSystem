VERSION 5.00
Begin VB.Form shippinglocation 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shipping Location"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox cbo_contactperson 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   6135
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   1320
         Width           =   6135
      End
      Begin VB.TextBox txt_telno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   3
         Top             =   3360
         Width           =   2655
      End
      Begin VB.ComboBox cbo_location 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   6135
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Tel No"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Contact person"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   2400
         Width           =   4455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Location"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
   End
End
Attribute VB_Name = "shippinglocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ap As Integer
Private Sub cbo_contactperson_Click()
Dim cp As New ADODB.Recordset
If cp.State Then cp.Close
cp.Open "select * from shipping ", Cn, 3, 2
While Not cp.EOF
If cp!personincharge1 = cbo_contactperson.Text Then
txt_telno.Text = cp!phone1
ElseIf cp!personincharge2 = cbo_contactperson.Text Then
txt_telno.Text = cp!phone2
ElseIf cp!personincharge3 = cbo_contactperson.Text Then
txt_telno.Text = cp!phone3
Else
MsgBox "Enter Phone No"
End If

cp.MoveNext
Wend
End Sub

Private Sub cbo_location_Click()
Dim sh1 As New ADODB.Recordset
If sh1.State Then sh1.Close
sh1.Open "select * from shipping where location='" & cbo_location.Text & "' ", Cn, 3, 2
If Not sh1.EOF Then

     cbo_contactperson.AddItem sh1!personincharge1
     cbo_contactperson.AddItem sh1!personincharge2
     cbo_contactperson.AddItem sh1!personincharge3
     txt_address.Text = sh1!address
    
End If
End Sub


Private Sub Form_Load()
 Dim sh As New ADODB.Recordset
 If sh.State Then sh.Close
 sh.Open "select DISTINCT(location) from shipping order by location", Cn, 3, 2
 While Not sh.EOF
 cbo_location.AddItem sh(0)
 sh.MoveNext
 Wend
 ap = 0
End Sub

