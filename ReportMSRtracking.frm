VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReportMSRtracking 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   9330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtp_from 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67174401
      CurrentDate     =   38674
   End
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View"
      Height          =   255
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   255
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Height          =   255
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cbo_msr 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8145
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   12225
      ExtentX         =   21564
      ExtentY         =   14367
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComCtl2.DTPicker dtp_to 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67174401
      CurrentDate     =   38674
   End
End
Attribute VB_Name = "ReportMSRtracking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nic As String
Private Sub Check1_Click()
 
Call nocolor
 
End Sub

Private Sub cmd_close_Click()
Unload Me
End Sub

Private Sub cmd_print_Click()
On Error GoTo XIT
WebBrowser.ExecWB 6, OLECMDEXECOPT_DODEFAULT
XIT:
End Sub

Private Sub cmd_show_Click()
 
'''Load frmBusy
'''frmBusy.Show
'''frmBusy.lblBusyString = "Please Wait Report Under Process......"
Call nocolor
'Unload frmBusy
 
End Sub

Private Sub dtp_from_Change()
Call msrload
End Sub

Private Sub dtp_from_Click()
Call msrload
End Sub

Private Sub dtp_to_Change()
Call msrload
End Sub

Private Sub dtp_to_Click()
Call msrload
End Sub

Private Sub Form_Load()
Call connect
main.lbltitle.Caption = "MSR Tracking"
 dtp_from.Value = Format(Date, "dd/MM/yyyy")
 dtp_to.Value = Format(Date, "dd/MM/yyyy")
Me.Top = 10
Me.Left = 10
 
WebBrowser.Navigate "About:Blank"
 Call msrload
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
Public Sub nocolor()
On Error Resume Next
Dim fso As New FileSystemObject
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
            
fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
   fs.WriteLine "      BORDER-BOTTOM: Wheat 1px solid;"
   fs.WriteLine "      BORDER-LEFT: Wheat 1px solid;"
   fs.WriteLine "      BORDER-RIGHT: Wheat 1px solid;"
   fs.WriteLine "      BORDER-TOP: Wheat 1px solid"
   fs.WriteLine "    }"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   fs.WriteLine "        FONT-FAMILY: Tahoma;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bolder;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "    }"
   fs.WriteLine "    .TrFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: black;"
   fs.WriteLine "        FONT-FAMILY: Tahoma;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "   }"
   fs.WriteLine "</style>"

fs.WriteLine "<body>"
fs.WriteLine "<table width=623 cellpadding=0 cellspacing=0 border=0 bordercolor=#acacac>"
fs.WriteLine " <tr  class=TableFont>"
fs.WriteLine " <td width=618 height=37 align=left><font size=3 FAMILY=Tahoma><b><u>MSR TRACKING between  </u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_from.Value, "dd/MMM/yy") & " and</u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_to.Value, "dd/MMM/yy") & " </u></b></td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'header


fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569
fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap align=center><font color=white>S.No</td>"
fs.WriteLine "    <td nowrap><font color=white>MSR No.</td>"
fs.WriteLine "   <td nowrap><font color=white>MSR Date</td>"
fs.WriteLine "   <td colspan=8></td></tr>"
fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td colspan=2>&nbsp;</td>"
fs.WriteLine "    <td nowrap><font color=white>Item</td>"
fs.WriteLine "    <td nowrap><font color=white>Required Date</td>"
fs.WriteLine "    <td nowrap><font color=white>MSR App</td>"
fs.WriteLine "    <td nowrap><font color=white>Buyer Assign</td>"
fs.WriteLine "   <td colspan=5></td></tr>"
fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td colspan=3>&nbsp;</td>"
fs.WriteLine "    <td nowrap><font color=white>RFQ No.</td>"
fs.WriteLine "    <td nowrap><font color=white>RFQ Date</td>"
fs.WriteLine "    <td nowrap><font color=white>Vendor</td>"

fs.WriteLine "    <td nowrap><font color=white>Quot No.</td>"
fs.WriteLine "    <td nowrap><font color=white>Quot Date</td>"
fs.WriteLine "    <td nowrap><font color=white>Promised Date</td>"
fs.WriteLine "    <td nowrap><font color=white>PO No.</td>"
fs.WriteLine "    <td nowrap><font color=white>PO Date</td>"

fs.WriteLine "  </tr>" 'item header
Dim sn As Integer
sn = 0

sn = sn + 1
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "   <td align=center>" & sn & "</td>"
Dim invdet As New ADODB.Recordset
If invdet.State Then invdet.Close
invdet.Open "select DISTINCT(p.prno),p.prdate from purchaserequisition p,prdetails pd where p.prno=pd.prno and p.prno='" & cbo_msr.Text & "' ", Cn, 3, 2
If Not invdet.EOF Then
fs.WriteLine "     <td nowrap>" & invdet(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(1) & "</td>"
'fs.WriteLine "     <td align=center>" & invdet(2) & "</td>"
fs.WriteLine "     <td colspan=8>&nbsp;</td>"
End If
invdet.Close
invdet.Open "select DISTINCT(p.material),p.reqdate,p.status,p.buyer,p.rfqno from prdetails p  where  p.prno='" & cbo_msr.Text & "' ", Cn, 3, 2
While Not invdet.EOF
If invdet.RecordCount > 1 Then
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "     <td colspan=2>&nbsp;</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(2) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(3) & "</td>"
fs.WriteLine "     <td colspan=5>&nbsp;</td>"
fs.WriteLine "   </tr>"
Dim rf As New ADODB.Recordset
If rf.State Then rf.Close
rf.Open "select r.rfqno, r.rfqdate, v.vendor  from rfq r ,prdetails p ,vendorgroup v where r.rfqno=p.rfqno and r.vendorgroup=v.vgroup and r.rfqno='" & invdet(4) & "' and p.material='" & invdet(0) & "'", Cn, 3, 2
While Not rf.EOF
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "     <td colspan=3>&nbsp;</td>"
fs.WriteLine "     <td align=left nowrap>" & rf(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & rf(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & rf(2) & "</td>"
Dim qt As New ADODB.Recordset
If qt.State Then qt.Close
qt.Open "select q.qno,q.qdate,qd.promisedate,qd.pono from quotation q, quotationdetails qd where q.q_id=qd.qtid and q.rfqno='" & invdet(4) & "' and q.vendor = '" & rf(2) & "' ", Cn, 3, 2
If Not qt.EOF Then
fs.WriteLine "     <td align=left nowrap>" & qt(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & qt(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & qt(2) & "</td>"
        If qt(3) <> "" Then
        fs.WriteLine "     <td align=left nowrap>" & qt(3) & "</td>"
        Dim rspo1 As New ADODB.Recordset
        If rspo1.State Then rspo1.Close
        rspo1.Open "select podate from po where pono='" & qt(3) & "' ", Cn, 3, 2
        If Not rspo1.EOF Then
        fs.WriteLine "     <td align=left nowrap>" & rspo1(0) & "</td>"
        End If
        Else
        fs.WriteLine "     <td align=left nowrap>NA</td>"
        fs.WriteLine "     <td align=left nowrap>NA</td>"
        End If

Else
fs.WriteLine "     <td align=centre >NA</td>"
fs.WriteLine "     <td align=centre >NA</td>"
fs.WriteLine "     <td align=centre >NA</td>"
fs.WriteLine "     <td align=centre >NA</td>"
End If
qt.Close
fs.WriteLine "   </tr>"
rf.MoveNext
Wend
rf.Close

Else
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "     <td align=left nowrap>" & invdet(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & invdet(2) & "</td>"
fs.WriteLine "   </tr>"
Dim rf1 As New ADODB.Recordset
If rf1.State Then rf1.Close
rf1.Open "select r.rfqno, r.rfqdate, v.vendor  from rfq r ,prdetails p ,vendorgroup v where r.rfqno=p.rfqno and r.vendorgroup=v.vgroup and r.rfqno='" & invdet(4) & "' and p.material='" & invdet(0) & "'", Cn, 3, 2
While Not rf1.EOF
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "     <td colspan=4>&nbsp;</td>"
fs.WriteLine "     <td align=left nowrap>" & rf1(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & rf1(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & rf1(2) & "</td>"
Dim qt1 As New ADODB.Recordset
If qt1.State Then qt1.Close
qt1.Open "select q.qno,q.qdate,qd.promisedate,qd.pono from quotation q, quotationdetails qd where q.q_id=qd.qt_id and q.rfqno='" & invdet(4) & "' and q.vendor = '" & rf(2) & "' ", Cn, 3, 2
If Not qt1.EOF Then
fs.WriteLine "     <td align=left nowrap>" & qt1(0) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & qt1(1) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & qt1(2) & "</td>"
        If qt1(3) <> "" Then
        fs.WriteLine "     <td align=left nowrap>" & qt1(3) & "</td>"
        Dim rspo As New ADODB.Recordset
        If rspo.State Then rspo.Close
        rspo.Open "select podate from po where pono='" & qt1(3) & "' ", Cn, 3, 2
        If Not rspo.EOF Then
        fs.WriteLine "     <td align=left nowrap>" & rspo(0) & "</td>"
        End If
        Else
        fs.WriteLine "     <td align=left nowrap>NA</td>"
        fs.WriteLine "     <td align=left nowrap>NA</td>"
        End If
Else
fs.WriteLine "     <td align=centre colspan=4>NA</td>"
End If
qt1.Close
fs.WriteLine "   </tr>"
rf1.MoveNext
Wend
rf1.Close


End If
invdet.MoveNext
Wend
invdet.Close

fs.WriteLine "   </tr>"

fs.WriteLine "</table>"
WebBrowser.Navigate App.Path & "\rep.html"
fs.WriteLine "<p>&nbsp;</p>"
fs.WriteLine "</body>"
fs.WriteLine "</html>"



End Sub
Public Sub msrload()
cbo_msr.Clear
Dim inv As New ADODB.Recordset
 If inv.State Then inv.Close
 inv.Open "select Distinct(prno) from purchaserequisition where prdate between '" & Format(dtp_from.Value, "MM/dd/yyyy") & "' and '" & Format(dtp_to.Value, "MM/dd/yyyy") & "' order by prno desc", Cn, 3, 2 ' where status <> 'P'
 While Not inv.EOF
 cbo_msr.AddItem inv(0)
 inv.MoveNext
 Wend
 inv.Close
End Sub
