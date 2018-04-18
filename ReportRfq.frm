VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReportRfq 
   BackColor       =   &H8000000E&
   Caption         =   "REQUEST FOR QUOTATION"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Frame f1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month / Year"
      Height          =   2655
      Left            =   120
      TabIndex        =   31
      Top             =   6600
      Width           =   11415
      Begin VB.ListBox vlist 
         Height          =   2085
         Left            =   5760
         Style           =   1  'Checkbox
         TabIndex        =   43
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox cbo_rfq 
         Height          =   1410
         Left            =   1800
         Style           =   1  'Checkbox
         TabIndex        =   32
         Top             =   480
         Width           =   3855
      End
      Begin VB.CommandButton cmd_rfq 
         Caption         =   "View RFQ"
         Height          =   375
         Left            =   9840
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp_rfqf 
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_rfqt 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Vendor/s list for the selected RFQ"
         Height          =   195
         Left            =   5760
         TabIndex        =   44
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RFQ No."
         Height          =   195
         Left            =   1800
         TabIndex        =   37
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   195
      End
   End
   Begin VB.Frame f2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Requestor"
      Height          =   2895
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   11415
      Begin VB.ListBox vlist1 
         Height          =   2085
         Left            =   5760
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd_rfq1 
         Caption         =   "View RFQ"
         Height          =   375
         Left            =   9840
         TabIndex        =   24
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbo_requestor 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox cbo_rfq1 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   1200
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtp_rfqf1 
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_rfqt1 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Vendor/s list for the selected RFQ"
         Height          =   195
         Left            =   5760
         TabIndex        =   42
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RFQ No."
         Height          =   195
         Left            =   1680
         TabIndex        =   30
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Requestor"
         Height          =   195
         Left            =   1680
         TabIndex        =   29
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame f3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JobCharge"
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   11415
      Begin VB.ListBox vlist2 
         Height          =   2085
         Left            =   5760
         Style           =   1  'Checkbox
         TabIndex        =   39
         Top             =   480
         Width           =   3975
      End
      Begin VB.ComboBox cbo_jobcharge 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd_rfq2 
         Caption         =   "View RFQ"
         Height          =   375
         Left            =   9840
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox cbo_rfq2 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   1200
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtp_rfqf2 
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_rfqt2 
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Vendor/s list for the selected RFQ"
         Height          =   195
         Left            =   5760
         TabIndex        =   40
         Top             =   240
         Width           =   2400
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "JobCharge"
         Height          =   195
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RFQ No."
         Height          =   195
         Left            =   1680
         TabIndex        =   19
         Top             =   960
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000009&
      Caption         =   "View By"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      Begin VB.OptionButton opt_rfq 
         BackColor       =   &H80000009&
         Caption         =   "RFQ"
         Height          =   255
         Left            =   2160
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_requestor 
         BackColor       =   &H80000009&
         Caption         =   "Requestor"
         Height          =   255
         Left            =   4500
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_jobcharge 
         BackColor       =   &H80000009&
         Caption         =   "Jobcharge"
         Height          =   255
         Left            =   7320
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   5040
      Width           =   11415
      Begin VB.ComboBox cbo_rfqmsr 
         Height          =   315
         Left            =   5640
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "View RFQ"
         Height          =   375
         Left            =   9360
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtp_msr 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   28311555
         CurrentDate     =   38662
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "RFQ No."
         Height          =   195
         Left            =   5640
         TabIndex        =   6
         Top             =   120
         Width           =   630
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MSR No."
         Height          =   195
         Left            =   1560
         TabIndex        =   4
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Month/Year"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "ReportRfq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rfqno As String
Public ii As Integer
Public VendNo As String
Private Sub cbo_jobcharge_Click()
Call loadrfq2
End Sub

Private Sub cbo_msr_Click()
Call loadrfqmsr
End Sub

Private Sub cbo_requestor_Click()
Call loadrfq1
End Sub



Private Sub cbo_rfq_ItemCheck(Item As Integer)
If cbo_rfq.SelCount > 1 Then
MsgBox " Multiple RFQ's cannot be viewed"
cbo_rfq.Selected(Item) = False
Else
vlist.Clear
Dim v As New ADODB.Recordset
If v.State Then v.Close
v.Open "select * from rfq r, vendorgroup v  where r.prno=v.vprno and  r.rfqno='" & cbo_rfq.List(Item) & "'", Cn, 3, 2
While Not v.EOF
vlist.AddItem v!vendor
v.MoveNext
Wend
v.Close
End If
End Sub


Private Sub cbo_rfq1_ItemCheck(Item As Integer)
If cbo_rfq1.SelCount > 1 Then
MsgBox " Multiple RFQ's cannot be viewed"
cbo_rfq1.Selected(Item) = False
Else
vlist1.Clear
Dim v1 As New ADODB.Recordset
If v1.State Then v1.Close
v1.Open "select * from rfq r, vendorgroup v  where r.prno=v.vprno and  r.rfqno='" & cbo_rfq1.List(Item) & "'", Cn, 3, 2
While Not v1.EOF
vlist1.AddItem v1!vendor
v1.MoveNext
Wend
v1.Close
End If
End Sub

Private Sub cbo_rfq2_ItemCheck(Item As Integer)
If cbo_rfq2.SelCount > 1 Then
MsgBox " Multiple RFQ's cannot be viewed"
cbo_rfq2.Selected(Item) = False
Else
vlist2.Clear
Dim v2 As New ADODB.Recordset
If v2.State Then v2.Close
v2.Open "select * from rfq r, vendorgroup v  where r.prno=v.vprno and  r.rfqno='" & cbo_rfq2.List(Item) & "'", Cn, 3, 2
While Not v2.EOF
vlist2.AddItem v2!vendor
v2.MoveNext
Wend
v2.Close
End If
End Sub

Private Sub cmd_rfq_Click()
rfqno = ""
ii = 0
For ii = 0 To cbo_rfq.ListCount - 1
If cbo_rfq.Selected(ii) = True Then
rfqno = cbo_rfq.List(ii)
    Call loadrfqdetails
End If
Next ii
VendNo = ""
ii = 0
For ii = 0 To vlist.ListCount - 1
If vlist.Selected(ii) = True Then
VendNo = vlist.List(ii)
   
End If
Next ii
Call loadrfqdetails
End Sub

Private Sub cmd_rfq1_Click()
rfqno = ""
ii = 0
For ii = 0 To cbo_rfq1.ListCount - 1
If cbo_rfq1.Selected(ii) = True Then
rfqno = cbo_rfq1.List(ii)
    
End If
Next ii
VendNo = ""
ii = 0
For ii = 0 To vlist1.ListCount - 1
If vlist1.Selected(ii) = True Then
VendNo = vlist1.List(ii)
   
End If
Next ii
Call loadrfqdetails
End Sub

Private Sub cmd_rfq2_Click()
rfqno = ""

ii = 0
For ii = 0 To cbo_rfq2.ListCount - 1
If cbo_rfq2.Selected(ii) = True Then
rfqno = cbo_rfq2.List(ii)
   
End If
Next ii
VendNo = ""
ii = 0
For ii = 0 To vlist2.ListCount - 1
If vlist2.Selected(ii) = True Then
VendNo = vlist2.List(ii)
  
End If
Next ii
Call loadrfqdetails
End Sub
Private Sub Command1_Click()
rfqno = ""
rfqno = cbo_rfqmsr.Text
Call loadrfqdetails
End Sub
Private Sub dtp_msr_Click()
Call loadmsr
End Sub
Private Sub dtp_rfqf_Change()
Call loadrfq
End Sub
Private Sub dtp_rfqf_Click()
Call loadrfq
End Sub
Private Sub dtp_rfqf1_Change()
Call loadrequestor
End Sub
Private Sub dtp_rfqf1_Click()
Call loadrequestor
End Sub
Private Sub dtp_rfqf2_Change()
Call loadjobcharge
End Sub
Private Sub dtp_rfqf2_Click()
Call loadjobcharge
End Sub
Private Sub dtp_rfqt_Change()
Call loadrfq
End Sub
Private Sub dtp_rfqt_Click()
Call loadrfq
End Sub

Private Sub dtp_rfqt1_Change()
Call loadrequestor
End Sub

Private Sub dtp_rfqt1_Click()
Call loadrequestor

End Sub

Private Sub dtp_rfqt2_Change()
Call loadjobcharge
End Sub

Private Sub dtp_rfqt2_Click()
Call loadjobcharge
End Sub

Private Sub Form_Load()
dtp_rfqf.Value = Format(Date, "dd/MM/yyyy")
dtp_rfqt.Value = Format(Date, "dd/MM/yyyy")
dtp_rfqf1.Value = Format(Date, "dd/MM/yyyy")
dtp_rfqt1.Value = Format(Date, "dd/MM/yyyy")
dtp_rfqf2.Value = Format(Date, "dd/MM/yyyy")
dtp_rfqt2.Value = Format(Date, "dd/MM/yyyy")
Call loadrequestor
Call loadjobcharge
Call loadmsr
Call loadrfq
Call loadrfq1
Call loadrfq2
Call loadrfqmsr
End Sub
Public Sub loadrfq()
Dim msr As New ADODB.Recordset
If msr.State Then msr.Close
msr.Open "select DISTINCT(rfqno) from rfq where (rfqdate) >='" & Format(dtp_rfqf.Value, "mm/dd/yyyy") & "' and (rfqdate) <= '" & Format(dtp_rfqt.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
While Not msr.EOF
cbo_rfq.AddItem msr(0)
msr.MoveNext
Wend
msr.Close
End Sub
Public Sub loadrfq1()
Dim msr1 As New ADODB.Recordset
If msr1.State Then msr1.Close
msr1.Open "select DISTINCT(r.rfqno) from rfq r, purchaserequisition p where p.prno=r.prno and (r.rfqdate) >= '" & Format(dtp_rfqf1.Value, "mm/dd/yyyy") & "' and r.rfqdate <= '" & Format(dtp_rfqt1.Value, "mm/dd/yyyy") & "' and  p.requestor ='" & cbo_requestor.Text & "'", Cn, 3, 2
While Not msr1.EOF
cbo_rfq1.AddItem msr1(0)
msr1.MoveNext
Wend
msr1.Close
End Sub
Public Sub loadrfq2()
Dim msr2 As New ADODB.Recordset
If msr2.State Then msr2.Close
msr2.Open "select DISTINCT(r.rfqno) from rfq r, prdetails p where p.prno=r.prno and (r.rfqdate) >= '" & Format(dtp_rfqf2.Value, "mm/dd/yyyy") & "' and r.rfqdate <= '" & Format(dtp_rfqt2.Value, "mm/dd/yyyy") & "' and  p.jobcharge ='" & cbo_jobcharge.Text & "'", Cn, 3, 2
While Not msr2.EOF
cbo_rfq2.AddItem msr2(0)
msr2.MoveNext
Wend
msr2.Close
End Sub

Public Sub loadrfqdetails()
On Error Resume Next
Dim ef1 As Double
Dim ef2 As Double
Dim ef3 As Double
Dim ef4 As Double
ef1 = 0: ef2 = 0: ef3 = 0: ef4 = 0
Dim ADVA As Double
   ADVA = 0
Dim rstemp As New Recordset


rstemp.Fields.Append "lblname", adVarChar, 200
rstemp.Fields.Append "lbladdress", adVarChar, 500
rstemp.Fields.Append "lblrfq", adVarChar, 100
rstemp.Fields.Append "lbldate", adDate, 100

rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Fields.Append "txtitem", adVarChar, 200
rstemp.Fields.Append "txtqty", adVarChar, 100
rstemp.Fields.Append "txtuom", adVarChar, 100
rstemp.Fields.Append "txtdate", adDate, 100
rstemp.Fields.Append "txtremarks", adVarChar, 500

rstemp.Fields.Append "lbljobcharge", adVarChar, 100
rstemp.Fields.Append "lbllocations", adVarChar, 100
rstemp.Open
Dim sn As Integer
sn = 0
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from rfq r, vendorgroup v  where r.prno=v.vprno and  r.rfqno='" & rfqno & "' and vendor ='" & VendNo & "'", Cn, 3, 2
If Not pr.EOF Then
                 Dim vn As New ADODB.Recordset
                 If vn.State Then vn.Close
                 vn.Open "select * from activityresponsible a, vendor v  where a.avendor=v.code and  a.avendorname='" & VendNo & "' ", Cn, 3, 2
                 If Not vn.EOF Then
    RptRfq.Sections("section4").Controls("lblname").Caption = "Attn: Mr." & vn!rpin
    RptRfq.Sections("section4").Controls("lbladdress").Caption = vn!Name & vn!Address
                 End If
    RptRfq.Sections("section4").Controls("lblrfq").Caption = pr("rfqno")
    RptRfq.Sections("section4").Controls("lbldate").Caption = Format(pr("rfqdate"), "dd/MMM/yy")
End If
pr.Close
pr.Open "select * from prdetails p ,rfq r where p.rfqno=r.rfqno and r.rfqno='" & rfqno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtitem") = pr("material")
    rstemp("txtqty") = pr("qty")
    rstemp("txtuom") = pr("uom")
    rstemp("txtremarks") = pr("remarks")
    rstemp("txtdate") = pr("reqdate")
    jc(sn) = pr("jobcharge")
    lc(sn) = pr("location")
pr.MoveNext
Wend
'location
Dim j As Integer
j = 0
cnt = ""
Dim jbc As New ADODB.Recordset
If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(location) from prdetails where rfqno= '" & rfqno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptRfq.Sections("section5").Controls("lbllocation").Caption = "All Items to be Delivered @ " & jbc(0)
Else
While Not jbc.EOF
j = j + 1
If cnt = "" Then
cnt = j & ". " & jbc(0)
Else
cnt = cnt + ",  " & j & ". " & jbc(0)
End If
jbc.MoveNext
Wend
jbc.Close
RptRfq.Sections("section5").Controls("lbllocation").Caption = "Delivery for " & cnt
End If
rstemp.Update
pr.Close

Set RptRfq.DataSource = rstemp
RptRfq.Show
SetParent RptRfq.hwnd, ReportRfq.hwnd

End Sub

Public Sub loadrequestor()
cbo_requestor.Clear
Dim rrq As New ADODB.Recordset
If rrq.State Then rrq.Close
rrq.Open "select DISTINCT(p.requestor) from purchaserequisition p, rfq r where p.prno=r.prno and (r.rfqdate)>= '" & Format(dtp_rfqf1.Value, "mm/dd/yyyy") & "' and r.rfqdate <= '" & Format(dtp_rfqt1.Value, "mm/dd/yyyy") & "'  order by p.requestor", Cn, 3, 2
While Not rrq.EOF
cbo_requestor.AddItem rrq(0)
rrq.MoveNext
Wend
rrq.Close
End Sub
Public Sub loadjobcharge()
cbo_jobcharge.Clear
Dim jch As New ADODB.Recordset
If jch.State Then jch.Close
jch.Open "select DISTINCT(p.jobcharge) from rfq r,prdetails p where p.prno=r.prno and (r.rfqdate)>= '" & Format(dtp_rfqf2.Value, "mm/dd/yyyy") & "' and r.rfqdate <='" & Format(dtp_rfqt2.Value, "mm/dd/yyyy") & "'  order by p.jobcharge", Cn, 3, 2
While Not jch.EOF
cbo_jobcharge.AddItem jch(0)
jch.MoveNext
Wend
jch.Close
End Sub
Public Sub loadmsr()
'cbo_rfq.Clear
'Dim rmsr As New ADODB.Recordset
'If rmsr.State Then rmsr.Close
'rmsr.Open "select DISTINCT(rfqno) from rfq where (rfqdate) between '" & Format(dtp_rfqf.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_rfqt.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
'While Not rmsr.EOF
'cbo_msr.AddItem rmsr(0)
'rmsr.MoveNext
'Wend
'rmsr.Close
End Sub
Public Sub loadrfqmsr()
'cbo_rfqmsr.Clear
'Dim rfm As New ADODB.Recordset
'If rfm.State Then rfm.Close
'rfm.Open "select DISTINCT(rfqno) from rfq where Month(rfqdate)='" & Format(dtp_msr.Value, "MM") & "' and Year(rfqdate)='" & Format(dtp_msr.Value, "yyyy") & "' and prno='" & cbo_msr.Text & "'", Cn, 3, 2
'While Not rfm.EOF
'cbo_rfqmsr.AddItem rfm(0)
'rfm.MoveNext
'Wend
'rfm.Close

End Sub

Public Sub rfq_job_rfq()
On Error Resume Next
Dim ef1 As Double
Dim ef2 As Double
Dim ef3 As Double
Dim ef4 As Double
ef1 = 0: ef2 = 0: ef3 = 0: ef4 = 0
Dim ADVA As Double
   ADVA = 0
Dim rstemp As New Recordset


rstemp.Fields.Append "lblname", adVarChar, 200
rstemp.Fields.Append "lbladdress", adVarChar, 500
rstemp.Fields.Append "lblrfq", adVarChar, 100
rstemp.Fields.Append "lbldate", adDate, 100

rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Fields.Append "txtitem", adVarChar, 200
rstemp.Fields.Append "txtqty", adVarChar, 100
rstemp.Fields.Append "txtuom", adVarChar, 100
rstemp.Fields.Append "txtdate", adDate, 100
rstemp.Fields.Append "txtremarks", adVarChar, 500

rstemp.Fields.Append "lbljobcharge", adVarChar, 100
rstemp.Fields.Append "lbllocations", adVarChar, 100
rstemp.Open
Dim sn As Integer
sn = 0
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from rfq r, vendorgroup v  where r.prno=v.vprno and  r.rfqno='" & rfqno & "'", Cn, 3, 2
If Not pr.EOF Then
                 Dim vn As New ADODB.Recordset
                 If vn.State Then vn.Close
                 vn.Open "select * from vendor where name='" & pr!vendor & "' ", Cn, 3, 2
                 If Not vn.EOF Then
    RptRfq.Sections("section4").Controls("lblname").Caption = "Attn: Mr." & vn!personincharge
    RptRfq.Sections("section4").Controls("lbladdress").Caption = vn!Name & vn!Address
                 End If
    RptRfq.Sections("section4").Controls("lblrfq").Caption = pr("rfqno")
    RptRfq.Sections("section4").Controls("lbldate").Caption = Format(pr("rfqdate"), "dd/MMM/yy")
End If
pr.Close
pr.Open "select * from prdetails p ,rfq r where p.rfqno=r.rfqno and r.rfqno='" & rfqno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtitem") = pr("material")
    rstemp("txtqty") = pr("qty")
    rstemp("txtuom") = pr("uom")
    rstemp("txtremarks") = pr("remarks")
    rstemp("txtdate") = pr("reqdate")
    jc(sn) = pr("jobcharge")
    lc(sn) = pr("location")
pr.MoveNext
Wend
'location
Dim j As Integer
j = 0
cnt = ""
Dim jbc As New ADODB.Recordset
If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(location) from prdetails where rfqno= '" & rfqno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptRfq.Sections("section5").Controls("lbllocation").Caption = "All Items to be Delivered @ " & jbc(0)
Else
While Not jbc.EOF
j = j + 1
If cnt = "" Then
cnt = j & ". " & jbc(0)
Else
cnt = cnt + ",  " & j & ". " & jbc(0)
End If
jbc.MoveNext
Wend
jbc.Close
RptRfq.Sections("section5").Controls("lbllocation").Caption = "Delivery for " & cnt
End If
rstemp.Update
pr.Close

Set RptRfq.DataSource = rstemp
RptRfq.Show
SetParent RptRfq.hwnd, ReportRfq.hwnd


End Sub

Private Sub vlist_ItemCheck(Item As Integer)
If vlist.SelCount > 1 Then
MsgBox " Multiple Vendor's cannot be viewed"
vlist.Selected(Item) = False
End If
End Sub

Private Sub vlist1_ItemCheck(Item As Integer)
If vlist1.SelCount > 1 Then
MsgBox " Multiple Vendor's cannot be viewed"
vlist1.Selected(Item) = False
End If
End Sub

Private Sub vlist2_ItemCheck(Item As Integer)
If vlist2.SelCount > 1 Then
MsgBox " Multiple Vendor's cannot be viewed"
vlist2.Selected(Item) = False
End If
End Sub
