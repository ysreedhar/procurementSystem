VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReportQuot 
   BackColor       =   &H8000000E&
   Caption         =   "QUOTATION RECEIVED"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10080
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.Frame fm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MSR"
      Height          =   3015
      Left            =   120
      TabIndex        =   38
      Top             =   1320
      Width           =   7575
      Begin VB.CommandButton cmd_quot4 
         Appearance      =   0  'Flat
         Caption         =   "View Quotation"
         Height          =   375
         Left            =   5880
         TabIndex        =   41
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbo_msr 
         Height          =   315
         Left            =   1680
         TabIndex        =   40
         Top             =   480
         Width           =   3975
      End
      Begin VB.ListBox cbo_quot4 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   39
         Top             =   1200
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtp_msrf 
         Height          =   315
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_msrt 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Quotation No."
         Height          =   195
         Left            =   1680
         TabIndex        =   47
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MSR"
         Height          =   195
         Left            =   1680
         TabIndex        =   46
         Top             =   240
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   345
      End
   End
   Begin VB.Frame fj 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JobCharge"
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   7575
      Begin VB.ListBox cbo_quot2 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton cmd_quot2 
         Caption         =   "View Quotation"
         Height          =   375
         Left            =   5880
         TabIndex        =   20
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbo_jobcharge 
         Height          =   315
         Left            =   1680
         TabIndex        =   19
         Top             =   480
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtp_jobf 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_jobt 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Quotation No."
         Height          =   195
         Left            =   1680
         TabIndex        =   25
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "JobCharge"
         Height          =   195
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Requestor"
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   7575
      Begin VB.ListBox cbo_quot1 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox cbo_requestor 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd_quot1 
         Caption         =   "View Quotation"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp_reqf 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_reqt 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Requestor"
         Height          =   195
         Left            =   1680
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Quotation No."
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   990
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Caption         =   "View By"
      Height          =   735
      Left            =   120
      TabIndex        =   48
      Top             =   240
      Width           =   11415
      Begin VB.OptionButton opt_vendor 
         BackColor       =   &H80000009&
         Caption         =   "Vendor"
         Height          =   255
         Left            =   3210
         TabIndex        =   53
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_quot 
         BackColor       =   &H80000009&
         Caption         =   "Quotation"
         Height          =   255
         Left            =   1680
         TabIndex        =   52
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton opt_msr 
         BackColor       =   &H80000009&
         Caption         =   "MSR"
         Height          =   255
         Left            =   8040
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opt_requestor 
         BackColor       =   &H80000009&
         Caption         =   "Requestor"
         Height          =   255
         Left            =   4500
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_jobcharge 
         BackColor       =   &H80000009&
         Caption         =   "Jobcharge"
         Height          =   255
         Left            =   6270
         TabIndex        =   49
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fv 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vendor"
      Height          =   2895
      Left            =   120
      TabIndex        =   28
      Top             =   1320
      Width           =   7575
      Begin VB.ListBox cbo_quot3 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   31
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox cbo_vendor 
         Height          =   315
         Left            =   1680
         TabIndex        =   30
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd_quot3 
         Caption         =   "View Quotation"
         Height          =   375
         Left            =   5880
         TabIndex        =   29
         Top             =   480
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtp_venf 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_vent 
         Height          =   315
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Vendor"
         Height          =   195
         Left            =   1680
         TabIndex        =   35
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Quotation No."
         Height          =   195
         Left            =   1680
         TabIndex        =   34
         Top             =   960
         Width           =   990
      End
   End
   Begin VB.Frame fq 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quotation"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7575
      Begin VB.CommandButton cmd_quot 
         Caption         =   "Quotation"
         Height          =   375
         Left            =   5880
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.ListBox cbo_quot 
         Height          =   1410
         Left            =   1800
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   480
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker dtp_quotf 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_quott 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   67239939
         CurrentDate     =   38662
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Quotation No."
         Height          =   195
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   345
      End
   End
End
Attribute VB_Name = "ReportQuot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public qtno As String
Public ii As Integer

Private Sub cbo_jobcharge_Click()
cbo_quot2.Clear
Dim msr2 As New ADODB.Recordset
If msr2.State Then msr2.Close
msr2.Open "select DISTINCT(q.qno) from quotation q, prdetails r  where q.rfqno=r.rfqno and (q.qdate)>='" & Format(dtp_jobf.Value, "mm/dd/yyyy") & "' and (q.qdate)<='" & Format(dtp_jobt.Value, "mm/dd/yyyy") & "' and r.jobcharge='" & cbo_jobcharge.Text & "' ", Cn, 3, 2
While Not msr2.EOF
cbo_quot2.AddItem msr2(0)
msr2.MoveNext
Wend
msr2.Close
End Sub

Private Sub cbo_msr_Click()
cbo_quot4.Clear
Dim msr4 As New ADODB.Recordset
If msr4.State Then msr4.Close
msr4.Open "select DISTINCT(qno) from quotationdetails  where  prno='" & cbo_msr.Text & "' ", Cn, 3, 2
While Not msr4.EOF
cbo_quot4.AddItem msr4(0)
msr4.MoveNext
Wend
msr4.Close
End Sub

Private Sub cbo_requestor_Click()
cbo_quot1.Clear
Dim msr1 As New ADODB.Recordset
If msr1.State Then msr1.Close
msr1.Open "select DISTINCT(q.qno) from quotation q, rfq r ,purchaserequisition pr where q.rfqno=r.rfqno and r.prno=pr.prno and (q.qdate)>='" & Format(dtp_reqf.Value, "mm/dd/yyyy") & "' and (q.qdate) <='" & Format(dtp_reqt.Value, "mm/dd/yyyy") & "' and pr.requestor='" & cbo_requestor.Text & "' ", Cn, 3, 2
While Not msr1.EOF
cbo_quot1.AddItem msr1(0)
msr1.MoveNext
Wend
msr1.Close
End Sub

Private Sub cbo_vendor_Click()
cbo_quot3.Clear
Dim msr4 As New ADODB.Recordset
If msr4.State Then msr4.Close
msr4.Open "select DISTINCT(qno) from quotation where (qdate)>='" & Format(dtp_venf.Value, "mm/dd/yyyy") & "' and (qdate) <='" & Format(dtp_vent.Value, "mm/dd/yyyy") & "' and vendor='" & cbo_vendor.Text & "' ", Cn, 3, 2
While Not msr4.EOF
cbo_quot3.AddItem msr4(0)
msr4.MoveNext
Wend
msr4.Close
End Sub
Private Sub cmd_quot_Click()
qtno = ""

ii = 0
For ii = 0 To cbo_quot.ListCount - 1
If cbo_quot.Selected(ii) = True Then
qtno = cbo_quot.List(ii)
    
End If
Next ii
Call loadquotdetails
End Sub

Private Sub cmd_quot1_Click()
qtno = ""

ii = 0
For ii = 0 To cbo_quot1.ListCount - 1
If cbo_quot1.Selected(ii) = True Then
qtno = cbo_quot1.List(ii)
    
End If
Next ii
Call loadquotdetails

End Sub

Private Sub cmd_quot2_Click()
qtno = ""

ii = 0
For ii = 0 To cbo_quot2.ListCount - 1
If cbo_quot2.Selected(ii) = True Then
qtno = cbo_quot2.List(ii)
    
End If
Next ii
Call loadquotdetails
End Sub

Private Sub cmd_quot3_Click()
qtno = ""

ii = 0
For ii = 0 To cbo_quot3.ListCount - 1
If cbo_quot3.Selected(ii) = True Then
qtno = cbo_quot3.List(ii)
    
End If
Next ii
Call loadquotdetails
End Sub
Private Sub cmd_quot4_Click()
qtno = ""

ii = 0
For ii = 0 To cbo_quot4.ListCount - 1
If cbo_quot4.Selected(ii) = True Then
qtno = cbo_quot4.List(ii)
    
End If
Next ii
Call loadquotdetails
End Sub


Private Sub dtp_jobf_Change()
Call loadjobcharge
End Sub

Private Sub dtp_jobf_Click()
Call loadjobcharge
End Sub

Private Sub dtp_msrf_Change()
Call loadqmsr
End Sub

Private Sub dtp_msrf_Click()
Call loadqmsr
End Sub

Private Sub dtp_msrt_Change()
Call loadqmsr
End Sub

Private Sub dtp_msrt_Click()
Call loadqmsr
End Sub

Private Sub dtp_quotf_Change()
Call loadquot
End Sub

Private Sub dtp_quotf_Click()
Call loadquot
End Sub

Private Sub dtp_quott_Change()
Call loadquot
End Sub

Private Sub dtp_quott_Click()
Call loadquot
End Sub

Private Sub dtp_reqf_Change()
Call loadrequestor
End Sub

Private Sub dtp_reqf_Click()
Call loadrequestor
End Sub

Private Sub dtp_reqt_Change()
Call loadrequestor
End Sub

Private Sub dtp_reqt_Click()
Call loadrequestor
End Sub

Private Sub dtp_venf_Change()
Call loadvendor
End Sub

Private Sub dtp_venf_Click()
Call loadvendor
End Sub

Private Sub Form_Load()
dtp_quotf.Value = Format(Date, "dd/MM/yyyy")
dtp_quott.Value = Format(Date, "dd/MM/yyyy")

dtp_reqf.Value = Format(Date, "dd/MM/yyyy")
dtp_reqt.Value = Format(Date, "dd/MM/yyyy")

dtp_venf.Value = Format(Date, "dd/MM/yyyy")
dtp_vent.Value = Format(Date, "dd/MM/yyyy")

dtp_jobf.Value = Format(Date, "dd/MM/yyyy")
dtp_jobt.Value = Format(Date, "dd/MM/yyyy")

dtp_msrf.Value = Format(Date, "dd/MM/yyyy")
dtp_msrt.Value = Format(Date, "dd/MM/yyyy")


fq.Visible = True
fv.Visible = False
fr.Visible = False
fj.Visible = False
fm.Visible = False

Call loadrequestor
Call loadjobcharge
Call loadvendor
Call loadquot
Call loadqmsr


End Sub
Public Sub loadqmsr()
cbo_msr.Clear
Dim rfm As New ADODB.Recordset
If rfm.State Then rfm.Close
rfm.Open "select DISTINCT(r.prno) from quotationdetails q, purchaserequisition r where q.prno=r.prno and (r.prdate)>='" & Format(dtp_msrf.Value, "mm/dd/yyyy") & "' and (r.prdate)<='" & Format(dtp_msrt.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
While Not rfm.EOF
cbo_msr.AddItem rfm(0)
rfm.MoveNext
Wend
rfm.Close

End Sub
Public Sub loadquot()
cbo_quot.Clear
Dim msr As New ADODB.Recordset
If msr.State Then msr.Close
msr.Open "select DISTINCT(qno) from quotation where (qdate)>= '" & Format(dtp_quotf.Value, "mm/dd/yyyy") & "' and (qdate) <= '" & Format(dtp_quott.Value, "mm/dd/yyyy") & "' ", Cn, 3, 2
While Not msr.EOF
cbo_quot.AddItem msr(0)
msr.MoveNext
Wend
msr.Close
End Sub

Public Sub loadquotdetails()
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
rstemp.Fields.Append "lblquot", adVarChar, 100
rstemp.Fields.Append "lbldate", adDate, 100

rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Fields.Append "txtitem", adVarChar, 200
rstemp.Fields.Append "txtqty", adVarChar, 100
rstemp.Fields.Append "txtuom", adVarChar, 100
rstemp.Fields.Append "txtuprice", adDouble, 100
rstemp.Fields.Append "txttprice", adDouble, 100
rstemp.Fields.Append "txtremarks", adVarChar, 500
rstemp.Fields.Append "txtdate", adDate, 500

rstemp.Fields.Append "lbluprice", adDouble, 100
rstemp.Fields.Append "lbltprice", adDouble, 100

rstemp.Fields.Append "lbljobcharge", adVarChar, 100
rstemp.Fields.Append "lbllocations", adVarChar, 100
rstemp.Open
Dim sn As Integer
sn = 0
Dim amt As Double
Dim upr As Double
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from quotation where qno ='" & qtno & "'", Cn, 3, 2
If Not pr.EOF Then
                                  
                 Dim vn As New ADODB.Recordset
                 If vn.State Then vn.Close
                 vn.Open "select * from activityresponsible a, vendor v  where a.avendor=v.code and  a.avendorname='" & pr!vendor & "' ", Cn, 3, 2
                 If Not vn.EOF Then
    RptQuot.Sections("section4").Controls("lblname").Caption = "From: Mr." & vn!rpin
    RptQuot.Sections("section4").Controls("lbladdress").Caption = vn!Name & vn!Address
                 End If
    RptQuot.Sections("section4").Controls("lblquot").Caption = pr("qno")
    RptQuot.Sections("section4").Controls("lbldate").Caption = Format(pr("qdate"), "dd/MMM/yy")
End If
pr.Close
amt = 0: upr = 0
pr.Open "select * from quotationdetails where qno='" & qtno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtitem") = pr("material")
    rstemp("txtqty") = pr("qty")
    rstemp("txtuom") = pr("uom")
    rstemp("txtremarks") = pr("remarks")
    rstemp("txtuprice") = Format(pr("unitrate"), "###,###,##0") ' & " " & pr!Currency
    upr = upr + pr!unitrate
    rstemp("txttprice") = Format(pr("amount"), "###,###,##0") ' & " " & pr!Currency
    amt = amt + pr!amount
    jc(sn) = pr("promisedate")
    lc(sn) = pr("location")
pr.MoveNext
Wend


RptQuot.Sections("section5").Controls("lbluprice").Caption = Format(upr, "###,###,##0")
RptQuot.Sections("section5").Controls("lbltprice").Caption = Format(amt, "###,###,##0")


'promised date
Dim j As Integer
j = 0
Dim cnt As String
Dim jbc As New ADODB.Recordset
If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(promisedate) from quotationdetails where qno= '" & qtno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptQuot.Sections("section5").Controls("lbljobcharge").Caption = "All Items will be Delivered on " & jbc(0)
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

RptQuot.Sections("section5").Controls("lbljobcharge").Caption = "Promised Delivery date for " & cnt
End If




'location

j = 0
cnt = ""
jbc.Close
'If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(location) from quotationdetails where qno= '" & qtno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptQuot.Sections("section5").Controls("lbllocation").Caption = "All Items to be Delivered @ " & jbc(0)
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
RptQuot.Sections("section5").Controls("lbllocation").Caption = "Delivery for " & cnt
End If
rstemp.Update
pr.Close

Set RptQuot.DataSource = rstemp
RptQuot.Show
SetParent RptQuot.hwnd, ReportQuot.hwnd

End Sub
Public Sub loadrequestor()
cbo_requestor.Clear
Dim rrq As New ADODB.Recordset
If rrq.State Then rrq.Close
rrq.Open "select DISTINCT(p.requestor) from purchaserequisition p, rfq r ,quotation q where p.prno=r.prno and r.rfqno=q.rfqno and (q.qdate) >= '" & Format(dtp_reqf.Value, "mm/dd/yyyy") & "' and (q.qdate) <= '" & Format(dtp_reqt.Value, "mm/dd/yyyy") & "'  order by p.requestor", Cn, 3, 2
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
jch.Open "select DISTINCT(p.jobcharge) from quotation q,prdetails p where p.rfqno=q.rfqno and (q.qdate) >= '" & Format(dtp_jobf.Value, "mm/dd/yyyy") & "' and (q.qdate) <= '" & Format(dtp_jobt.Value, "mm/dd/yyyy") & "'  order by p.jobcharge", Cn, 3, 2
While Not jch.EOF
cbo_jobcharge.AddItem jch(0)
jch.MoveNext
Wend
jch.Close
End Sub

Public Sub loadvendor()
cbo_vendor.Clear
Dim msr3 As New ADODB.Recordset
If msr3.State Then msr3.Close
msr3.Open "select DISTINCT(vendor) from quotation  where  (qdate) >= '" & Format(dtp_venf.Value, "mm/dd/yyyy") & "' and (qdate) <= '" & Format(dtp_vent.Value, "mm/dd/yyyy") & "'  ", Cn, 3, 2
While Not msr3.EOF
cbo_vendor.AddItem msr3(0)
msr3.MoveNext
Wend
msr3.Close

End Sub

Private Sub opt_jobcharge_Click()
fq.Visible = False
fv.Visible = False
fr.Visible = False
fj.Visible = True
fm.Visible = False
End Sub

Private Sub opt_msr_Click()
fq.Visible = False
fv.Visible = False
fr.Visible = False
fj.Visible = False
fm.Visible = True
End Sub

Private Sub opt_quot_Click()
fq.Visible = True
fv.Visible = False
fr.Visible = False
fj.Visible = False
fm.Visible = False
End Sub

Private Sub opt_requestor_Click()
fq.Visible = False
fv.Visible = False
fr.Visible = True
fj.Visible = False
fm.Visible = False
End Sub

Private Sub opt_vendor_Click()
fq.Visible = False
fv.Visible = True
fr.Visible = False
fj.Visible = False
fm.Visible = False
End Sub
