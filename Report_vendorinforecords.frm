VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Report_vendorinforecords 
   BackColor       =   &H8000000E&
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   11895
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View"
      Height          =   255
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   255
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Height          =   255
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cbo_material 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8745
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   11865
      ExtentX         =   20929
      ExtentY         =   15425
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
      Location        =   ""
   End
   Begin MSComCtl2.DTPicker dtp_from 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67305473
      CurrentDate     =   38674
   End
   Begin MSComCtl2.DTPicker dtp_to 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   67305473
      CurrentDate     =   38674
   End
End
Attribute VB_Name = "Report_vendorinforecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nic As String


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



Private Sub Form_Load()
Call connect
main.lbltitle.Caption = "Vendor InfoRecords"
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
fs.WriteLine " <td width=618 height=37 align=left><font size=3 FAMILY=Tahoma><b><u>VENDOR INFO-RECORDS between  </u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_from.Value, "dd/MMM/yy") & " and</u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_to.Value, "dd/MMM/yy") & " </u></b></td>"
'fs.WriteLine " <td width=618 height=37 align=left><font size=1 FAMILY=Tahoma><b><u>" & Format(dtp_from.Value, "dd/MMM/yy") & " and</u></b><b><u>" & Format(dtp_to.Value, "dd/MMM/yy") & " </u></b></td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'header


fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569

fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap colspan=3><font color=white>Vendor Name</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Vendor Code</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Reg.No</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Cmp Status</td>"
fs.WriteLine "    <td nowrap colspan=3><font color=white>Equity Status(%)</td>"
fs.WriteLine "  </tr>" 'item header
mn = Split(cbo_material.Text, "  -  ", Len(cbo_material.Text), vbTextCompare)
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select * from vendor where code='" & mn(1) & "'", Cn, 3, 2
If Not vn.EOF Then

fs.WriteLine "  <tr bgcolor=white class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap colspan=3>" & vn!Name & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & vn!code & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & vn!regno & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & vn!cmpstatus & "</td>"
fs.WriteLine "    <td nowrap colspan=1>BM: " & vn!bmstatus & "</td>"
fs.WriteLine "    <td nowrap colspan=1>NB: " & vn!nbstatus & "</td>"
fs.WriteLine "    <td nowrap colspan=1>FL: " & vn!flstatus & "</td>"
fs.WriteLine "  </tr>" 'item header

End If

fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap align=center><font color=white>S.No</td>"
fs.WriteLine "    <td nowrap><font color=white>RFQ No.</td>"
fs.WriteLine "    <td nowrap><font color=white>RFQ Date</td>"
fs.WriteLine "   <td ><font color=white>Quot Closing Date</td>"
fs.WriteLine "    <td nowrap><font color=white>Quot No</td>"
fs.WriteLine "    <td nowrap><font color=white>Quot Date</td>"
fs.WriteLine "    <td nowrap><font color=white>Award Order</td>"
fs.WriteLine "    <td nowrap><font color=white>PO No</td>"
fs.WriteLine "    <td nowrap><font color=white>PO Date</td>"
fs.WriteLine "  </tr>" 'item header

Dim sn As Integer

sn = 1


Dim pif As New ADODB.Recordset
If pif.State Then pif.Close
pif.Open "select DISTINCT(r.rfqno),r.rfqdate,r.closingdate from rfq r, vendorgroup v where r.prno=v.vprno and v.vendor='" & mn(0) & "' and r.rfqdate between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' order BY r.rfqno", Cn, 3, 2
While Not pif.EOF
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "   <td align=center>" & sn & "</td>"
fs.WriteLine "     <td nowrap>" & pif(0) & "</td>"
fs.WriteLine "     <td nowrap>" & pif(1) & "</td>"
fs.WriteLine "     <td  nowrap>" & pif(2) & "</td>"
                 Dim qt As New ADODB.Recordset
                 If qt.State Then qt.Close
                 qt.Open "select qno,qdate,award from quotation where rfqno='" & pif(0) & "' and vendor='" & mn(0) & "' ", Cn, 3, 2
                    If Not qt.EOF Then
                    fs.WriteLine "     <td nowrap>" & qt(0) & "</td>"
                    fs.WriteLine "     <td  nowrap>" & qt(1) & "</td>"
                    fs.WriteLine "     <td  nowrap>" & qt(2) & "</td>"
                    Else
                    fs.WriteLine "     <td nowrap>NA</td>"
                    fs.WriteLine "     <td  nowrap>NA</td>"
                    fs.WriteLine "     <td  nowrap>NA</td>"
                    End If
                    If qt(0) = "" Or qt(0) = Null Then
                    fs.WriteLine "     <td nowrap>NA</td>"
                    fs.WriteLine "     <td  nowrap>NA</td>"
                    Else
                    Dim pp As New ADODB.Recordset
                    If pp.State Then pp.Close
                    pp.Open "select pono,podate from po where qno='" & qt(0) & "'", Cn, 3, 2
                    If Not pp.EOF Then
                    fs.WriteLine "     <td nowrap>" & pp(0) & "</td>"
                    fs.WriteLine "     <td  nowrap>" & pp(1) & "</td>"
                    Else
                    fs.WriteLine "     <td nowrap>NA</td>"
                    fs.WriteLine "     <td  nowrap>NA</td>"
                    End If
                    End If
                    pp.Close
                    qt.Close

fs.WriteLine "   </tr>"
sn = sn + 1
pif.MoveNext
Wend

fs.WriteLine "</table>"
WebBrowser.Navigate App.Path & "\rep.html"
fs.WriteLine "<p>&nbsp;</p>"
fs.WriteLine "</body>"
fs.WriteLine "</html>"



End Sub
Public Sub msrload()
cbo_material.Clear
Dim med As New ADODB.Recordset
If med.State Then med.Close
med.Open "select DISTINCT(name),code from vendor order by name", Cn, 3, 2
While Not med.EOF
cbo_material.AddItem med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close
End Sub
