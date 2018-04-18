VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form goodstransfer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goods Transfer"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbo_status 
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
      TabIndex        =   19
      Top             =   1560
      Width           =   1755
   End
   Begin VB.TextBox txt_account 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Details"
      Height          =   4575
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   10095
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   9855
         TabIndex        =   15
         Top             =   1320
         Width           =   9855
      End
      Begin VB.ComboBox cbo_lookup 
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
         Left            =   1440
         TabIndex        =   12
         Top             =   960
         Width           =   3075
      End
      Begin VB.ComboBox cbo_materialtype 
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
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   3075
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Select Item By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Item Type"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "To"
      Height          =   1935
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cbo_storagelocationto 
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
         TabIndex        =   7
         Top             =   1440
         Width           =   3795
      End
      Begin VB.ComboBox cbo_worklocationto 
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
         TabIndex        =   6
         Top             =   720
         Width           =   3795
      End
      Begin MSComCtl2.DTPicker dtp_to 
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67239937
         CurrentDate     =   38873
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Storage Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Transfered From"
      Height          =   1935
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin MSComCtl2.DTPicker dtp_from 
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67239937
         CurrentDate     =   38873
      End
      Begin VB.ComboBox cbo_worklocationfrom 
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
         TabIndex        =   2
         Top             =   720
         Width           =   3795
      End
      Begin VB.ComboBox cbo_storagelocationfrom 
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
         TabIndex        =   1
         Top             =   1440
         Width           =   3795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Work Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Storage Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1440
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer Status"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A04729&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "GT No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A04729&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   585
   End
End
Attribute VB_Name = "goodstransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_lookup_Click()
If cbo_materialtype.Text = "" Then
MsgBox "Select Item Type"
cbo_lookup.Text = ""
cbo_materialtype.SetFocus
Exit Sub
End If
Call lookupdetails
End Sub

Private Sub cbo_worklocationfrom_Click()
cbo_storagelocationfrom.Clear
Dim str1 As New ADODB.Recordset
If str1.State Then str1.Close
str1.Open "select DISTINCT(location) from shipping where worklocation='" & cbo_worklocationfrom.Text & "'  ", Cn, 3, 2
While Not str1.EOF
cbo_storagelocationfrom.AddItem str1(0)
str1.MoveNext
Wend
str1.Close

End Sub
Private Sub cbo_worklocationto_Click()
cbo_storagelocationto.Clear
Dim str2 As New ADODB.Recordset
If str2.State Then str2.Close
str2.Open "select DISTINCT(location) from shipping where worklocation='" & cbo_worklocationto.Text & "'  ", Cn, 3, 2
While Not str2.EOF
cbo_storagelocationto.AddItem str2(0)
str2.MoveNext
Wend
str2.Close
End Sub

Private Sub Form_Load()
On Error Resume Next
Unload vscrollgoodstransfer
vscrollgoodstransfer.Show
vscrollgoodstransfer.Left = 0
vscrollgoodstransfer.Top = 0
 
SetParent vscrollgoodstransfer.hwnd, goodstransfer.Picture1.hwnd


'item
cbo_lookup.AddItem "Item ID"
cbo_lookup.AddItem "Mfr PartNo."
cbo_lookup.AddItem "Item Description"
cbo_lookup.AddItem "Search"

cbo_status.AddItem "InTransit"
cbo_status.AddItem "Received"

dtp_from.Value = Format(Date, "dd/MM/yyyy hh:mm:ss")
dtp_to.Value = Format(Date, "dd/MM/yyyy hh:mm:ss")
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select DISTINCT(mtcode),mtdesc from materialtype where rental='No' order by mtcode", Cn, 3, 2
While Not mb.EOF
cbo_materialtype.AddItem mb(0) & "  -  " & mb(1)
mb.MoveNext
Wend
mb.Close

'-----------
cbo_worklocationfrom.Clear
cbo_worklocationto.Clear
Dim prj As New ADODB.Recordset
If prj.State Then prj.Close
prj.Open "select DISTINCT(workloc) from worklocation", Cn, 3, 2
While Not prj.EOF
cbo_worklocationfrom.AddItem prj(0)
cbo_worklocationto.AddItem prj(0)
prj.MoveNext
Wend
prj.Close



mb.Open "select * from goodstransfer", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "GT-000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from goodstransfer where gtno='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_account.Text = "GT-000" & i
   Else
   i = i + 1
 
GoTo assad
  End If
  mb.Close


End Sub
Public Sub lookupdetails()

mty = Split(cbo_materialtype.Text, "  -  ", Len(cbo_materialtype.Text), vbTextCompare)
Dim itn As Integer
itn = 0
For itn = 0 To vscrollgoodstransfer.cbo_category.Count - 1
 
 Dim med As New ADODB.Recordset
If med.State Then med.Close

 If cbo_lookup.Text = "Item ID" Then
vscrollgoodstransfer.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollgoodstransfer.cbo_category(itn).AddItem med(0) & "  -  " & med(1) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
vscrollgoodstransfer.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollgoodstransfer.cbo_category(itn).AddItem med(1) & "  -  " & med(0) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close


ElseIf cbo_lookup.Text = "Item Description" Then
vscrollgoodstransfer.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollgoodstransfer.cbo_category(itn).AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close

ElseIf cbo_lookup.Text = "Search" Then
vscrollgoodstransfer.cbo_category(itn).Clear
End If

Next itn

End Sub


