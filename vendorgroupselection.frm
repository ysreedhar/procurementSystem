VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form vendorgroupselection 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_remarks 
      Enabled         =   0   'False
      Height          =   885
      Left            =   120
      TabIndex        =   13
      Top             =   8040
      Width           =   11415
   End
   Begin VB.CommandButton cmd_supplyorder 
      Caption         =   "Save Vendor Group"
      Height          =   375
      Left            =   6840
      TabIndex        =   12
      Top             =   240
      Width           =   2175
   End
   Begin VB.ComboBox cbo_msr 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txt_account 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.Frame OtherVendors 
      BackColor       =   &H8000000E&
      Caption         =   "Other Vendors"
      Height          =   6615
      Left            =   7800
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   6105
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame LicensedVendors 
      BackColor       =   &H8000000E&
      Caption         =   "Licensed Vendors"
      Height          =   6615
      Left            =   3960
      TabIndex        =   1
      Top             =   1080
      Width           =   3735
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   6105
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame BestVendors 
      BackColor       =   &H8000000E&
      Caption         =   "Best Vendors (Evaluated by System)"
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   6105
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComCtl2.DTPicker dtp_rfq 
      Height          =   300
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   28246017
      CurrentDate     =   38455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H00A04729&
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   7800
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "MSR No."
      ForeColor       =   &H00A04729&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   660
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "VG Date"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00A04729&
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor Group"
      ForeColor       =   &H00A04729&
      Height          =   195
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "vendorgroupselection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_msr_Click()
Call vgroupno
Call LicensedVendor

End Sub

Private Sub cmd_supplyorder_Click()
Cn.Execute "delete from vendorgroup where vgroup='" & txt_account.Text & "'"
Dim i As Integer
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
i = 0
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
vn.Open "select * from vendorgroup", Cn, 3, 2
vn.AddNew
vn!vgroup = txt_account.Text
vn!vdate = Format(dtp_rfq.Value, "dd/MM/yyyy")
vn!vprno = cbo_msr.Text
vn!vendor = List1.List(i)
vn!vcategory = List1.Name
vn!remarks = txt_remarks.Text
vn.Update
vn.Close
End If
Next i

i = 0
For i = 0 To List2.ListCount - 1
If List2.Selected(i) = True Then
vn.Open "select * from vendorgroup", Cn, 3, 2
vn.AddNew
vn!vgroup = txt_account.Text
vn!vdate = Format(dtp_rfq.Value, "dd/MM/yyyy")
vn!vprno = cbo_msr.Text
Dim vd As New ADODB.Recordset
If vd.State Then vd.Close
vd.Open "select DISTINCT(code) from vendor where name='" & List1.List(i) & "' ", Cn, 3, 2
If Not vd.EOF Then
vn!vendor = vd(0)

vn!vcategory = List2.Name
vn!remarks = txt_remarks.Text
vn.Update
vn.Close
End If

End If
Next i

i = 0
For i = 0 To List3.ListCount - 1
If List3.Selected(i) = True Then
vn.Open "select * from vendorgroup", Cn, 3, 2
vn.AddNew
vn!vgroup = txt_account.Text
vn!vdate = Format(dtp_rfq.Value, "dd/MM/yyyy")
vn!vprno = cbo_msr.Text
vn!vendor = List3.List(i)
vn!vcategory = List3.Name
vn!remarks = txt_remarks.Text
vn.Update
vn.Close
End If
Next i

MsgBox "Vendor selection saved successfully"
End Sub
Private Sub Form_Load()
dtp_rfq.Value = Format(Date, "dd/MM/yyyy")

Dim msr As New ADODB.Recordset
If msr.State Then msr.Close
msr.Open "select DISTINCT(prno) from purchaserequisition order by prno desc", Cn, 3, 2
While Not msr.EOF
cbo_msr.AddItem msr(0)
msr.MoveNext
Wend
Call LicensedVendor
End Sub
Public Sub LicensedVendor()
List1.Clear
List2.Clear
List3.Clear

Dim l As New ADODB.Recordset
If l.State Then l.Close
l.Open "select DISTINCT(material) from prdetails where prno='" & cbo_msr.Text & "' ", Cn, 3, 2
While Not l.EOF
sp = Split(l(0), "  -  ", Len(l(0)), vbTextCompare)
Dim lw1 As New ADODB.Recordset
If lw1.State Then lw1.Close
lw1.Open "select DISTINCT(vname) from vendormaterial v, prdetails p ,ml3 m where v.materialcode = p.itemid and v.material=m.ml3name and v.expirydate <= " & Format(Date, "dd/MM/yyyy") & " and p.prno= '" & txt_account.Text & "' order by vname", Cn, 3, 2
While Not lw1.EOF
List2.AddItem lw1(0)
lw1.MoveNext
Wend
lw1.Close

l.MoveNext
Wend
l.Close
'lw.Close
Dim lw As New ADODB.Recordset
lw.Open "select DISTINCT(name) from vendor order by name", Cn, 3, 2
While Not lw.EOF
List3.AddItem lw(0)
lw.MoveNext
Wend
lw.Close

End Sub

Public Sub vgroupno()
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from vendorgroup", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "VG000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from vendorgroup where vgroup='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_account.Text = "VG000" & i
   Else
   i = i + 1
 
GoTo assad
  End If

End Sub
