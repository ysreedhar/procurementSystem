VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Reportmsr 
   BackColor       =   &H00FFFFFF&
   Caption         =   "View / Print MSR"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9885
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H80000009&
      Caption         =   "View By"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   11415
      Begin VB.OptionButton opt_jobcharge 
         BackColor       =   &H80000009&
         Caption         =   "Jobcharge"
         Height          =   255
         Left            =   7320
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_requestor 
         BackColor       =   &H80000009&
         Caption         =   "Requestor"
         Height          =   255
         Left            =   4500
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton opt_msr 
         BackColor       =   &H80000009&
         Caption         =   "MSR"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame f3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JobCharge"
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   11415
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   5880
         TabIndex        =   32
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton opt_s 
            BackColor       =   &H00FF8080&
            Caption         =   "Summary"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   35
            Top             =   1320
            Width           =   1815
         End
         Begin VB.OptionButton opt_l 
            BackColor       =   &H00FF8080&
            Caption         =   "Landscape"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton opt_p 
            BackColor       =   &H00FF8080&
            Caption         =   "Potrait"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.ListBox cbo_msr2 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   1200
         Width           =   3975
      End
      Begin VB.CommandButton cmd_msr2 
         Caption         =   "View MSR "
         Height          =   375
         Left            =   9000
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cbo_jobcharge 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   480
         Width           =   3975
      End
      Begin MSComCtl2.DTPicker dtp_msr2 
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_msr2a 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MSR No."
         Height          =   195
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "JobCharge"
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame f2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Requestor"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   11415
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   5880
         TabIndex        =   36
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton opt_pr 
            BackColor       =   &H00FF8080&
            Caption         =   "Potrait"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   39
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton opt_lr 
            BackColor       =   &H00FF8080&
            Caption         =   "Landscape"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton opt_sr 
            BackColor       =   &H00FF8080&
            Caption         =   "Summary"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   1320
            Width           =   1815
         End
      End
      Begin VB.ListBox cbo_msr1 
         Height          =   1410
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   1200
         Width           =   3975
      End
      Begin VB.ComboBox cbo_requestor 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton cmd_msr1 
         Caption         =   "View MSR"
         Height          =   375
         Left            =   9000
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtp_msr1 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_msr1a 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Requestor"
         Height          =   195
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MSR No."
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   960
         Width           =   660
      End
   End
   Begin VB.Frame f1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Month / Year"
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11415
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   5760
         TabIndex        =   40
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton opt_sm 
            BackColor       =   &H00FF8080&
            Caption         =   "Summary"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   43
            Top             =   1320
            Width           =   1815
         End
         Begin VB.OptionButton opt_lm 
            BackColor       =   &H00FF8080&
            Caption         =   "Landscape"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   42
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton opt_pm 
            BackColor       =   &H00FF8080&
            Caption         =   "Potrait"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.ListBox cbo_msr 
         Height          =   1860
         Left            =   1800
         Style           =   1  'Checkbox
         TabIndex        =   21
         Top             =   480
         Width           =   3855
      End
      Begin VB.CommandButton cmd_msr 
         Caption         =   "View MSR"
         Height          =   375
         Left            =   9000
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker dtp_msr 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin MSComCtl2.DTPicker dtp_msra 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "To"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "MSR No."
         Height          =   195
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "From"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   345
      End
   End
End
Attribute VB_Name = "Reportmsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Public msrno As String
Public ii As Integer
Dim rstemp As New Recordset
Private Sub cbo_jobcharge_Click()
Call loadmsr2
End Sub

Private Sub cbo_msr_ItemCheck(Item As Integer)
If opt_s.Value = True Then
ElseIf opt_sm.Value = True Then
ElseIf opt_sr.Value = True Then
Else
If cbo_msr.SelCount > 1 Then
MsgBox " Multiple MSR's cannot be viewed"
cbo_msr.Selected(Item) = False
End If
End If

End Sub

Private Sub cbo_msr1_ItemCheck(Item As Integer)
If opt_s.Value = True Then
ElseIf opt_sm.Value = True Then
ElseIf opt_sr.Value = True Then
Else
If cbo_msr1.SelCount > 1 Then
MsgBox " Multiple MSR's cannot be viewed"
cbo_msr1.Selected(Item) = False
End If
End If

End Sub

Private Sub cbo_msr2_ItemCheck(Item As Integer)
If opt_s.Value = True Then
ElseIf opt_sm.Value = True Then
ElseIf opt_sr.Value = True Then
Else
If cbo_msr2.SelCount > 1 Then
MsgBox " Multiple MSR's cannot be viewed"
cbo_msr2.Selected(Item) = False
End If
End If
End Sub

Private Sub cbo_requestor_Click()
Call loadmsr1
End Sub
Private Sub cmd_msr_Click()
msrno = ""
ii = 0
For ii = 0 To cbo_msr.ListCount - 1
If cbo_msr.Selected(ii) = True Then
msrno = cbo_msr.List(ii)
                        If opt_pm.Value = True Then
                        Call loadmsrdetails
                        ElseIf opt_lm.Value = True Then
                        Call loadmsrls
                        ElseIf opt_sm.Value = True Then
                        Call loadmsrsummary
                        End If
End If
Next ii
End Sub
Private Sub cmd_msr1_Click()
msrno = ""
ii = 0
For ii = 0 To cbo_msr1.ListCount - 1
If cbo_msr1.Selected(ii) = True Then
msrno = cbo_msr1.List(ii)
                        If opt_pr.Value = True Then
                        Call loadmsrdetails
                        ElseIf opt_lr.Value = True Then
                        Call loadmsrls
                        ElseIf opt_sr.Value = True Then
                        Call loadmsrsummary
                        End If
End If
Next ii
End Sub
Private Sub cmd_msr2_Click()
msrno = ""
ii = 0
For ii = 0 To cbo_msr2.ListCount - 1
If cbo_msr2.Selected(ii) = True Then
msrno = cbo_msr2.List(ii)
                        If opt_p.Value = True Then
                        
                        Call loadmsrdetails
                        ElseIf opt_l.Value = True Then
                        
                        Call loadmsrls
                        ElseIf opt_s.Value = True Then
                        Call loadmsrsummary
                        End If
End If
Next ii
End Sub
Private Sub dtp_msr_Change()
Call loadmsr
End Sub
Private Sub dtp_msr1_Change()
Call loadrequestor
End Sub

Private Sub dtp_msr1a_Change()
Call loadrequestor
End Sub

Private Sub dtp_msr2_Change()
Call loadjobcharge
End Sub

Private Sub dtp_msr2a_Change()
Call loadjobcharge
End Sub
Private Sub dtp_msra_Change()
Call loadmsr
End Sub
Private Sub Form_Load()
On Error Resume Next
dtp_msr.Value = Format(Date, "dd/MMM/yyyy")
dtp_msra.Value = Format(Date, "dd/MMM/yyyy")
dtp_msr1.Value = Format(Date, "dd/MMM/yyyy")
dtp_msr2.Value = Format(Date, "dd/MMM/yyyy")
dtp_msr1a.Value = Format(Date, "dd/MMM/yyyy")
dtp_msr2a.Value = Format(Date, "dd/MMM/yyyy")
Call loadmsr
Call loadrequestor
Call loadjobcharge
f1.Visible = False
f2.Visible = False
f3.Visible = False
opt_p.Value = True
opt_pr.Value = True
opt_pm.Value = True
End Sub
Public Sub loadmsr()
cbo_msr.Text = ""
Dim msr As New ADODB.Recordset
If msr.State Then msr.Close
msr.Open "select DISTINCT(prno) from purchaserequisition where (prdate) between '" & Format(dtp_msr.Value, "MM/dd/yyyy") & "' and  '" & Format(dtp_msra.Value, "MM/dd/yyyy") & "' ", Cn, 3, 2
While Not msr.EOF
cbo_msr.AddItem msr(0)
msr.MoveNext
Wend
msr.Close
End Sub
Public Sub loadmsrdetails()
On Error Resume Next



rstemp.Close
rstemp.Fields.Append "lblrequestor", adVarChar, 100
rstemp.Fields.Append "lbldepartment", adVarChar, 100
rstemp.Fields.Append "lbllocation", adVarChar, 100
rstemp.Fields.Append "lblmsr", adVarChar, 100
rstemp.Fields.Append "lbldate", adDate, 100
rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Fields.Append "txtitem", adVarChar, 255
rstemp.Fields.Append "txtqty", adVarChar, 100
rstemp.Fields.Append "txtuom", adVarChar, 100
rstemp.Fields.Append "txtremarks", adVarChar, 255
rstemp.Fields.Append "lbljobcharge", adVarChar, 100
rstemp.Fields.Append "lbllocations", adVarChar, 100
rstemp.Fields.Append "txtcostcode", adVarChar, 255
rstemp.Fields.Append "txtrqddate", adDate, 100

'new
rstemp.Fields.Append "lblrequested", adVarChar, 100
rstemp.Fields.Append "lblrecommend", adVarChar, 100
rstemp.Fields.Append "lblapproved", adVarChar, 100
rstemp.Fields.Append "lbldesigrequested", adVarChar, 100
rstemp.Fields.Append "lbldesigrecommend", adVarChar, 100
rstemp.Fields.Append "lbldesigapproved", adVarChar, 100
rstemp.Fields.Append "lbldaterequested", adDate, 100
rstemp.Fields.Append "lbldaterecommend", adDate, 100
rstemp.Fields.Append "lbldateapproved", adDate, 100

rstemp.Open



Dim sn As Integer
sn = 0
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from purchaserequisition  where  prno='" & msrno & "'", Cn, 3, 2
If Not rsnn.EOF Then
            
    RptMSR.Sections("section4").Controls("lblrequestor").Caption = pr("requestor")
    RptMSR.Sections("section4").Controls("lbldepartment").Caption = pr("department")
    RptMSR.Sections("section4").Controls("lblmsr").Caption = pr("prno")
    RptMSR.Sections("section4").Controls("lbldate").Caption = Format(pr("prdate"), "dd/MMM/yy")
    'new
    RptMSR.Sections("section3").Controls("lblrequested").Caption = pr("requestor")
    RptMSR.Sections("section3").Controls("lbldesigrequested").Caption = pr("department")
    RptMSR.Sections("section3").Controls("lbldaterequested").Caption = Format(pr("prdate"), "dd/MMM/yy")
    
    Dim rls As New ADODB.Recordset
    If rls.State Then rls.Close
    rls.Open "select d.dname from prauth p, releasedetails r, designation d where p.at_user=r.rs_code and r.rs_desig=d.dcode and p.at_msrno='" & msrno & "' and rs_rlcode='RCM'", Cn, 3, 2
    If Not rls.EOF Then
    RptMSR.Sections("section3").Controls("lblrecommend").Caption = pr("recommendor")
    RptMSR.Sections("section3").Controls("lbldesigrecommend").Caption = rls(0)
    RptMSR.Sections("section3").Controls("lbldaterecommend").Caption = Format(pr("rdate"), "dd/MMM/yy")
    End If
    rls.Close
    rls.Open "select d.dname from prauth p, releasedetails r, designation d where p.at_user=r.rs_code and r.rs_desig=d.dcode and p.at_msrno='" & msrno & "' and rs_rlcode='APP'", Cn, 3, 2
    If Not rls.EOF Then
    RptMSR.Sections("section3").Controls("lblapproved").Caption = pr("approver")
    RptMSR.Sections("section3").Controls("lbldesigapproved").Caption = rls(0)
    RptMSR.Sections("section3").Controls("lbldateapproved").Caption = Format(pr("adate"), "dd/MMM/yy")
    End If
    rls.Close
    
    
End If
pr.Close
pr.Open "select * from prdetails where prno='" & msrno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtitem") = pr("material")
    rstemp("txtqty") = pr("qty")
    rstemp("txtuom") = pr("uom")
    rstemp("txtremarks") = pr("remarks")
    sp = Split(pr("jobcharge"), "  -  ", Len(pr("jobcharge")), vbTextCompare)
    Dim jb As New ADODB.Recordset
    If jb.State Then jb.Close
    jb.Open "select costcode from jobcharge where job_code='" & sp(0) & "'", Cn, 3, 2
    If Not jb.EOF Then
    rstemp("txtcostcode") = "BgtCde: " & jb(0)
    End If
    jc(sn) = pr("jobcharge")
    lc(sn) = pr("location")
pr.MoveNext
Wend
'nn = RptMSR.Sections("section5").Controls("lbljobcharge").Caption
Dim j As Integer
j = 0
Dim cnt As String
Dim jbc As New ADODB.Recordset
If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(jobcharge) from prdetails where prno= '" & msrno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptMSR.Sections("section5").Controls("lbljobcharge").Caption = "All Items for " & jbc(0)
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

RptMSR.Sections("section5").Controls("lbljobcharge").Caption = "Jobcharge for " & cnt
End If
'location
j = 0
cnt = ""
jbc.Close
jbc.Open "select DISTINCT(location) from prdetails where prno= '" & msrno & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptMSR.Sections("section5").Controls("lbllocation").Caption = "All Items to be Delivered @ " & jbc(0)
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
RptMSR.Sections("section5").Controls("lbllocation").Caption = "Delivery for " & cnt
End If

'---------------


rstemp.UpdateBatch
'pr.Close

Printer.Orientation = vbPRORPortrait
Set RptMSR.DataSource = rstemp

RptMSR.Show

'SetParent RptMSR.hwnd, Reportmsr.hwnd

'-----------------------------------------------------------
'RptMSR.ExportReport rptKeyHTML, "C:\MSR-Reports\" & msrno, True, , _
'rptRangeAllPages


'-----------------------------------------------------------





End Sub
Public Sub loadrequestor()
cbo_requestor.Clear
Dim rrq As New ADODB.Recordset
If rrq.State Then rrq.Close
rrq.Open "select DISTINCT(requestor) from purchaserequisition where (prdate) between '" & Format(dtp_msr1.Value, "MM/dd/yyyy") & "' and '" & Format(dtp_msr1a.Value, "MM/dd/yyyy") & "'  order by requestor", Cn, 3, 2
While Not rrq.EOF
cbo_requestor.AddItem rrq(0)
rrq.MoveNext
Wend
rrq.Close
End Sub
Public Sub loadmsr1()
cbo_msr1.Clear
Dim msr1 As New ADODB.Recordset
If msr1.State Then msr1.Close
msr1.Open "select DISTINCT(prno) from purchaserequisition where (prdate) between '" & Format(dtp_msr1.Value, "MM/dd/yyyy") & "' and '" & Format(dtp_msr1a.Value, "MM/dd/yyyy") & "'  and requestor='" & cbo_requestor.Text & "' ", Cn, 3, 2
While Not msr1.EOF
cbo_msr1.AddItem msr1(0)
msr1.MoveNext
Wend
msr1.Close
End Sub
Public Sub loadjobcharge()
cbo_jobcharge.Clear
Dim jch As New ADODB.Recordset
If jch.State Then jch.Close
jch.Open "select DISTINCT(pd.jobcharge) from purchaserequisition p,prdetails pd where p.prno=pd.prno and (p.prdate) between '" & Format(dtp_msr2.Value, "MM/dd/yyyy") & "' and '" & Format(dtp_msr2a.Value, "MM/dd/yyyy") & "'   order by pd.jobcharge", Cn, 3, 2
While Not jch.EOF
cbo_jobcharge.AddItem jch(0)
jch.MoveNext
Wend
jch.Close
End Sub

Public Sub loadmsr2()
cbo_msr2.Clear
Dim msr2 As New ADODB.Recordset
If msr2.State Then msr2.Close
msr2.Open "select DISTINCT(p.prno) from purchaserequisition p,prdetails pd where p.prno=pd.prno and (p.prdate) between '" & Format(dtp_msr2.Value, "MM/dd/yyyy") & "' and '" & Format(dtp_msr2a.Value, "MM/dd/yyyy") & "'  and pd.jobcharge ='" & cbo_jobcharge.Text & "' ", Cn, 3, 2
While Not msr2.EOF
cbo_msr2.AddItem msr2(0)
msr2.MoveNext
Wend
msr2.Close

End Sub

Private Sub opt_jobcharge_Click()
f1.Visible = False
f2.Visible = False
f3.Visible = True
End Sub

Private Sub opt_l_Click()
On Error Resume Next
ChngPrinterOrientationLandscape Me
End Sub

Private Sub opt_lm_Click()
On Error Resume Next
ChngPrinterOrientationLandscape Me
End Sub

Private Sub opt_lr_Click()
On Error Resume Next
ChngPrinterOrientationLandscape Me
End Sub

Private Sub opt_msr_Click()
f1.Visible = True
f2.Visible = False
f3.Visible = False
End Sub

Private Sub opt_p_Click()
On Error Resume Next
ChngPrinterOrientationPortrait Me
End Sub

Private Sub opt_pm_Click()
On Error Resume Next
ChngPrinterOrientationPortrait Me
End Sub

Private Sub opt_pr_Click()
On Error Resume Next
ChngPrinterOrientationPortrait Me
End Sub

Private Sub opt_requestor_Click()
f1.Visible = False
f2.Visible = True
f3.Visible = False
End Sub

Public Sub loadmsrls()
On Error Resume Next



If rstemp.State Then rstemp.Close
rstemp.Fields.Append "lblrequestor", adVarChar, 100
rstemp.Fields.Append "lbldepartment", adVarChar, 100
rstemp.Fields.Append "lblmsr", adVarChar, 100
rstemp.Fields.Append "lbldate", adDate, 100
rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Fields.Append "txtitem", adVarChar, 255
rstemp.Fields.Append "txtqty", adVarChar, 100
rstemp.Fields.Append "txtuom", adVarChar, 100
rstemp.Fields.Append "txtremarks", adVarChar, 255
rstemp.Fields.Append "txtjob", adVarChar, 100
rstemp.Fields.Append "txtlocation", adVarChar, 255
rstemp.Fields.Append "txtcostcode", adVarChar, 100
rstemp.Fields.Append "txtrqddate", adDate

'new
rstemp.Fields.Append "lblrequested", adVarChar, 100
rstemp.Fields.Append "lblrecommend", adVarChar, 100
rstemp.Fields.Append "lblapproved", adVarChar, 100
rstemp.Fields.Append "lbldesigrequested", adVarChar, 100
rstemp.Fields.Append "lbldesigrecommend", adVarChar, 100
rstemp.Fields.Append "lbldesigapproved", adVarChar, 100
rstemp.Fields.Append "lbldaterequested", adDate, 100
rstemp.Fields.Append "lbldaterecommend", adDate, 100
rstemp.Fields.Append "lbldateapproved", adDate, 100
rstemp.Open

Dim sn As Integer
sn = 0
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from purchaserequisition  where  prno='" & msrno & "'", Cn, 3, 2
If Not pr.EOF Then
    RptMsrLs.Sections("section4").Controls("lblrequestor").Caption = pr("requestor")
    RptMsrLs.Sections("section4").Controls("lbldepartment").Caption = pr("department")
    RptMsrLs.Sections("section4").Controls("lblmsr").Caption = pr("prno")
    RptMsrLs.Sections("section4").Controls("lbldate").Caption = Format(pr("prdate"), "dd/MMM/yy")
    
    RptMsrLs.Sections("section3").Controls("lblrequested").Caption = pr("requestor")
    RptMsrLs.Sections("section3").Controls("lbldesigrequested").Caption = pr("department")
    RptMsrLs.Sections("section3").Controls("lbldaterequested").Caption = Format(pr("prdate"), "dd/MMM/yy")
    
    Dim rls As New ADODB.Recordset
    If rls.State Then rls.Close
    rls.Open "select d.dname from prauth p, releasedetails r, designation d where p.at_user=r.rs_code and r.rs_desig=d.dcode and p.at_msrno='" & msrno & "' and rs_rlcode='RCM'", Cn, 3, 2
    If Not rls.EOF Then
    RptMsrLs.Sections("section3").Controls("lblrecommend").Caption = pr("recommendor")
    RptMsrLs.Sections("section3").Controls("lbldesigrecommend").Caption = rls(0)
    RptMsrLs.Sections("section3").Controls("lbldaterecommend").Caption = Format(pr("rdate"), "dd/MMM/yy")
    End If
    rls.Close
    rls.Open "select d.dname from prauth p, releasedetails r, designation d where p.at_user=r.rs_code and r.rs_desig=d.dcode and p.at_msrno='" & msrno & "' and rs_rlcode='APP'", Cn, 3, 2
    If Not rls.EOF Then
    RptMsrLs.Sections("section3").Controls("lblapproved").Caption = pr("approver")
    RptMsrLs.Sections("section3").Controls("lbldesigapproved").Caption = rls(0)
    RptMsrLs.Sections("section3").Controls("lbldateapproved").Caption = Format(pr("adate"), "dd/MMM/yy")
    End If
    rls.Close
    
End If
pr.Close
pr.Open "select * from prdetails where prno='" & msrno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtitem") = pr("material")
    rstemp("txtqty") = pr("qty")
    rstemp("txtuom") = pr("uom")
    rstemp("txtremarks") = pr("remarks")
    Dim sp As Variant
    sp = Split(pr("jobcharge"), "  -  ", Len(pr("jobcharge")), vbTextCompare)
    Dim jb As New ADODB.Recordset
    If jb.State Then jb.Close
    jb.Open "select costcode from jobcharge where job_code='" & sp(0) & "'", Cn, 3, 2
    If Not jb.EOF Then
    rstemp("txtcostcode") = "BgtCde: " & jb(0)
    End If
    rstemp("txtjob") = pr("jobcharge")
    rstemp("txtlocation") = pr("location")
    rstemp("txtrqddate") = Format(pr("reqdate"), "dd/MM/yyyy")
pr.MoveNext
Wend
'nn = RptMsrLs.Sections("section5").Controls("lbljobcharge").Caption

rstemp.UpdateBatch
pr.Close



Set RptMsrLs.DataSource = rstemp
'RptMsrLs.Orientation = rptOrientLandscape
'Printer.Orientation = 2
RptMsrLs.Show
'SetParent RptMsrLs.hwnd, Reportmsr.hwnd

'-----------------------------------------------------------
'RptMsrLs.ExportReport rptKeyHTML, "C:\MSR-Reports\" & msrno, True, , _
'rptRangeAllPages



End Sub

Public Sub loadmsrsummary()
On Error Resume Next

rstemp.Close
rstemp.Fields.Append "txtrequestor", adVarChar, 100
rstemp.Fields.Append "txtmsrdesc", adVarChar, 100
rstemp.Fields.Append "txtmsr", adVarChar, 100
rstemp.Fields.Append "txtdate", adDate, 100
rstemp.Fields.Append "txtsno", adVarChar, 100
rstemp.Open


For ls = 0 To cbo_msr2.ListCount - 1
If cbo_msr2.Selected(ls) = True Then
msrno = ""
msrno = cbo_msr2.List(ls)
Call msrsummary
End If
Next
ls = 0

For ls = 0 To cbo_msr1.ListCount - 1
If cbo_msr1.Selected(ls) = True Then
msrno = ""
msrno = cbo_msr1.List(ls)
Call msrsummary
End If
Next

ls = 0

For ls = 0 To cbo_msr.ListCount - 1
If cbo_msr.Selected(ls) = True Then
msrno = ""
msrno = cbo_msr.List(ls)
Call msrsummary
End If
Next



rstemp.UpdateBatch
pr.Close
Printer.Orientation = vbPRORPortrait
Set RptMsrSummary.DataSource = rstemp

RptMsrSummary.Show



End Sub

Private Sub opt_s_Click()
On Error Resume Next
ChngPrinterOrientationPortrait Me
End Sub

Public Sub msrsummary()


Dim ls As Integer
ls = 0
Dim sn As Integer
sn = 0
Dim jc(20) As String
Dim lc(20) As String
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from purchaserequisition where prno='" & msrno & "' ", Cn, 3, 2
While Not pr.EOF
sn = sn + 1
rstemp.AddNew
    rstemp("txtsno") = sn
    rstemp("txtmsr") = pr("prno")
    rstemp("txtmsrdesc") = pr("notes")
    rstemp("txtrequestor") = pr("requestor") & " ( " & pr("department") & ")"
    rstemp("txtdate") = Format(pr("prdate"), "dd/MMM/yy")
    
pr.MoveNext
Wend


End Sub
