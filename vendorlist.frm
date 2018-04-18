VERSION 5.00
Begin VB.Form vendorlist 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Vendors"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Vendor List"
      ForeColor       =   &H8000000E&
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Send Email"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   4440
         TabIndex        =   2
         Top             =   120
         Width           =   1190
      End
      Begin VB.ListBox List1 
         Height          =   3435
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   360
         Width           =   5415
      End
   End
End
Attribute VB_Name = "vendorlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
List1.Clear
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(name) from vendor order by name", Cn, 3, 2
While Not vn.EOF
List1.AddItem vn(0)
vn.MoveNext
Wend
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Dim str As String
'str = MsgBox("Do you want to cancel selected Vendor/s", vbOKCancel)
'If str = vbOK Then
'Unload Me
'Else
'Exit Sub
'End If
End Sub
