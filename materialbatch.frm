VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form materialbatch 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Material Batch"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Details"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   6975
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
         TabIndex        =   9
         Top             =   960
         Width           =   3075
      End
      Begin VB.ComboBox cbo_category 
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
         TabIndex        =   8
         Top             =   1320
         Width           =   6795
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
         TabIndex        =   7
         Top             =   360
         Width           =   3075
      End
      Begin VB.Label Label2 
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label Label4 
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
         TabIndex        =   10
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.TextBox txt_batchno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txt_notes 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin MSComCtl2.DTPicker DTP_bdate 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67698689
      CurrentDate     =   38733
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Batch Description"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material Batch"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "MB Creation Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "materialbatch"
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

Private Sub Form_Load()
DTP_bdate.Value = Format(Date, "dd/MM/yyyy")
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from materialbatch", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "MB-000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from materialbatch where batchno='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_batchno.Text = "MB-000" & i
   Else
   i = i + 1
 
GoTo assad
  End If
  mb.Close
  
mb.Open "select DISTINCT(mtcode),mtdesc from materialtype where rental='No' order by mtcode", Cn, 3, 2
While Not mb.EOF
cbo_materialtype.AddItem mb(0) & "  -  " & mb(1)
mb.MoveNext
Wend
mb.Close
cbo_lookup.AddItem "Item ID"
cbo_lookup.AddItem "Mfr PartNo."
cbo_lookup.AddItem "Item Description"
cbo_lookup.AddItem "Search"
End Sub
Public Sub lookupdetails()
cbo_category.Clear
mty = Split(cbo_materialtype.Text, "  -  ", Len(cbo_materialtype.Text), vbTextCompare)
 Dim med As New ADODB.Recordset
If med.State Then med.Close

 If cbo_lookup.Text = "Item ID" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(0) & "  -  " & med(1) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(1) & "  -  " & med(0) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close


ElseIf cbo_lookup.Text = "Item Description" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close

ElseIf cbo_lookup.Text = "Enter Item Manually" Then


ElseIf cbo_lookup.Text = "Search" Then
cbo_category.Clear
End If
End Sub

