VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form GRN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goods Received Note"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PO Details"
      Height          =   1935
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cbo_vendor 
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
         TabIndex        =   14
         Top             =   1440
         Width           =   3795
      End
      Begin VB.ComboBox cbo_po 
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
         TabIndex        =   13
         Top             =   720
         Width           =   3795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
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
         TabIndex        =   16
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "PO No"
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
         TabIndex        =   15
         Top             =   480
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delivery Details"
      Height          =   1935
      Left            =   6840
      TabIndex        =   7
      Top             =   120
      Width           =   4095
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
         TabIndex        =   9
         Top             =   720
         Width           =   3795
      End
      Begin VB.ComboBox cbo_storagelocation 
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
         Top             =   1440
         Width           =   3795
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
         TabIndex        =   11
         Top             =   480
         Width           =   1215
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
         TabIndex        =   10
         Top             =   1200
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Details"
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Width           =   11895
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
         TabIndex        =   4
         Top             =   360
         Width           =   3075
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
         TabIndex        =   3
         Top             =   960
         Width           =   3075
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   11655
         TabIndex        =   2
         Top             =   1320
         Width           =   11655
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
         TabIndex        =   6
         Top             =   480
         Width           =   885
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1275
      End
   End
   Begin VB.TextBox txt_account 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtp_grn 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   67633153
      CurrentDate     =   38873
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "GRN Date"
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
      TabIndex        =   19
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "GRN No."
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
      TabIndex        =   17
      Top             =   600
      Width           =   720
   End
End
Attribute VB_Name = "GRN"
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

Private Sub cbo_po_Click()
Call procpodetails
End Sub

Private Sub cbo_worklocation_Click()
cbo_storagelocation.Clear
Dim str1 As New ADODB.Recordset
If str1.State Then str1.Close
str1.Open "select DISTINCT(location) from shipping where worklocation='" & cbo_worklocation.Text & "'  ", Cn, 3, 2
While Not str1.EOF
cbo_storagelocation.AddItem str1(0)
str1.MoveNext
Wend
str1.Close

End Sub


Private Sub Form_Load()
On Error Resume Next
Unload vscrollGRN
vscrollGRN.Show
vscrollGRN.Left = 0
vscrollGRN.Top = 0
 
SetParent vscrollGRN.hwnd, GRN.Picture1.hwnd


'item
cbo_lookup.AddItem "Item ID"
cbo_lookup.AddItem "Mfr PartNo."
cbo_lookup.AddItem "Item Description"
cbo_lookup.AddItem "Search"



dtp_grn.Value = Format(Date, "dd/MM/yyyy hh:mm:ss")

Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select DISTINCT(mtcode),mtdesc from materialtype where rental='No' order by mtcode", Cn, 3, 2
While Not mb.EOF
cbo_materialtype.AddItem mb(0) & "  -  " & mb(1)
mb.MoveNext
Wend
mb.Close

'-----------
cbo_worklocation.Clear

Dim prj As New ADODB.Recordset
If prj.State Then prj.Close
prj.Open "select DISTINCT(workloc) from worklocation", Cn, 3, 2
While Not prj.EOF
cbo_worklocation.AddItem prj(0)
prj.MoveNext
Wend
prj.Close
cbo_po.Clear
prj.Open "select DISTINCT(pono) from po order by pono", Cn, 3, 2
While Not prj.EOF
cbo_po.AddItem prj(0)
prj.MoveNext
Wend
prj.Close

mb.Open "select * from GRN", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "GRN-000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from GRN where grno='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_account.Text = "GRN-000" & i
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
For itn = 0 To vscrollGRN.cbo_category.Count - 1
 
 Dim med As New ADODB.Recordset
If med.State Then med.Close

 If cbo_lookup.Text = "Item ID" Then
vscrollGRN.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollGRN.cbo_category(itn).AddItem med(0) & "  -  " & med(1) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
vscrollGRN.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollGRN.cbo_category(itn).AddItem med(1) & "  -  " & med(0) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close


ElseIf cbo_lookup.Text = "Item Description" Then
vscrollGRN.cbo_category(itn).Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
vscrollGRN.cbo_category(itn).AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close

ElseIf cbo_lookup.Text = "Search" Then
vscrollGRN.cbo_category(itn).Clear
End If

Next itn

End Sub
Public Sub procpodetails()
Dim cpo As Integer
cpo = 0
Dim pod As New ADODB.Recordset
If pod.State Then pod.Close
pod.Open "select * from podetails where pono='" & cbo_po.Text & "' ", Cn, 3, 2
For cpo = 0 To pod.RecordCount - 1
             vscrollGRN.cbo_category(cpo).Text = pod!itemid & "  -  " & pod!mrefcode & "  -  " & pod!material
             vscrollGRN.cbo_uom(cpo).Text = pod!uom
             vscrollGRN.txt_qty(cpo).Text = pod!qty
             vscrollGRN.txt_qtyrec(cpo).Text = pod!qty
             vscrollGRN.txt_qtyrej(cpo).Text = vscrollGRN.txt_qty(cpo).Text - vscrollGRN.txt_qtyrec(cpo).Text
             vscrollGRN.cbo_category(cpo).Enabled = False
             vscrollGRN.cbo_uom(cpo).Enabled = False
             vscrollGRN.txt_qty(cpo).Enabled = False
             vscrollGRN.txt_qtyrec(cpo).Enabled = False
             vscrollGRN.txt_qtyrej(cpo).Enabled = False
             vscrollGRN.cbo_rejection(cpo).Enabled = False

pod.MoveNext
Next cpo
pod.Close
pod.Open "select vendor from po where pono='" & cbo_po.Text & "' ", Cn, 3, 2
If Not pod.EOF Then
cbo_vendor.Text = pod(0)
End If

End Sub
