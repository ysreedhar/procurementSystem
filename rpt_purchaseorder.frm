VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form rpt_purchaseorder 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View"
      Height          =   255
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   255
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Height          =   255
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cbo_invoice 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8145
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   10665
      ExtentX         =   18812
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
      Location        =   ""
   End
End
Attribute VB_Name = "rpt_purchaseorder"
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

Private Sub Form_Load()
Call connect
main.lbltitle.Caption = "Purchase Order"
 
Me.Top = 10
Me.Left = 10
 
WebBrowser.Navigate "About:Blank"
 Dim inv As New ADODB.Recordset
 If inv.State Then inv.Close
 inv.Open "select Distinct(pono) from purchaseorder ", Cn, 3, 2 ' where status <> 'P'
 While Not inv.EOF
 cbo_invoice.AddItem inv(0)
 inv.MoveNext
 Wend
End Sub

Private Sub Form_Unload(Cancel As Integer)
main.lbltitle.Caption = ""
End Sub
Public Sub nocolor()
Dim fso As New FileSystemObject
   Set fs = fso.CreateTextFile(App.Path & "\rep.html")
   
fs.WriteLine " <html> "
   fs.WriteLine "<style>"
   fs.WriteLine "    BODY INPUT"
   fs.WriteLine "    {"
   fs.WriteLine "      BACKGROUND-IMAGE: url(file://C:\WINNT\FeatherTexture.bmp);"
   'fs.WriteLine "      BORDER-BOTTOM: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-LEFT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-RIGHT: Wheat 1px solid;"
   'fs.WriteLine "      BORDER-TOP: Wheat 1px solid"
   fs.WriteLine "    }"
   fs.WriteLine "    .TableFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: Black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   'fs.WriteLine "        'FONT-WEIGHT: bolder;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "    }"
   fs.WriteLine "    .TrFont"
   fs.WriteLine "    {"
   fs.WriteLine "        COLOR: black;"
   fs.WriteLine "        FONT-FAMILY: Arial Narrow;"
   fs.WriteLine "        FONT-SIZE: 8pt;"
   fs.WriteLine "        TEXT-TRANSFORM: capitalize;"
   fs.WriteLine "        CURSOR:HAND;"
   fs.WriteLine "   }"
   fs.WriteLine "</style>"

fs.WriteLine "<body>"
fs.WriteLine "<table width=623 cellpadding=0 cellspacing=0 border=1 bordercolor=black>"
fs.WriteLine " <tr class=TableFont>"
fs.WriteLine " <td width=617 height=37 align=center><font size=5><b>PURCHASE ORDER</b></td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'header

Dim invpan As New ADODB.Recordset
If invpan.State Then invpan.Close
invpan.Open "select v.personincharge,v.name,v.address,p.pono,p.podate from purchaseorder p,vendor v where p.vendor=v.name and p.pono='" & cbo_invoice.Text & "' ", Cn, 3, 2
If Not invpan.EOF Then
fs.WriteLine " <table width=623  cellpadding=0 cellspacing=0 border=1 bordercolor=black>" 'height=95

fs.WriteLine " <tr valign=0 class=TableFont>"
fs.WriteLine " <td width=382 >&nbsp;&nbsp;&nbsp;Attn:<b>" & invpan(0) & "</b> <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & invpan(1) & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & invpan(2) & "</td>"
fs.WriteLine " <td width=238>&nbsp;&nbsp;&nbsp;<b>PO No:</b>&nbsp;&nbsp;&nbsp;" & invpan(3) & " <br>&nbsp;&nbsp;&nbsp;<b>PO Date:</b>&nbsp;" & Format(invpan(4), "dd/MMM/yyyy") & "</td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'to address
End If
fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=black>" 'height=569
fs.WriteLine "  <tr bgcolor=gray class=TableFont>"
fs.WriteLine "    <td width=33 height=29><font color=white>&nbsp; S.No</td>"
fs.WriteLine "    <td width=335><font color=white>&nbsp; Item Name</td>"
fs.WriteLine "    <td  width=45 align=center><font color=white>Reqd.Date</td>"
fs.WriteLine "   <td width=45 align=center><font color=white> UOM</td>"
fs.WriteLine "   <td width=45 align=center><font color=white> Qty</td>"
fs.WriteLine "   <td width=45 align=center><font color=white> UnitRate</td>"
fs.WriteLine "    <td  width=45 align=center><font color=white>Amount</td>"

fs.WriteLine "  </tr>" 'item header
Dim sn As Integer
sn = 0
Dim rc As Integer
rc = 0
Dim i As Integer
i = 0
Dim ttot As Double
ttot = 0
Dim invdet As New ADODB.Recordset
If invdet.State Then invdet.Close
invdet.Open "select * from purchaseorder p,podetails pd where p.pono=pd.pono and p.pono='" & cbo_invoice.Text & "' ", Cn, 3, 2
If Not invdet.EOF Then
rc = invdet.RecordCount
Dim qty As Double
Dim urate As Double
Dim amt As Double
qty = 0: urate = 0: amt = 0

While Not invdet.EOF
sn = sn + 1
fs.WriteLine "   <tr valign=0 class=TableFont>"
fs.WriteLine "   <td >&nbsp;" & sn & "</td>"
fs.WriteLine "     <td>&nbsp;" & invdet!Name & "</td>"
fs.WriteLine "     <td align=center>" & Format(invdet!reqdate, "dd/MMM/yy") & "</td>"
fs.WriteLine "     <td align=right>&nbsp;" & invdet!uom & "</td>"
fs.WriteLine "     <td align=right>&nbsp;" & invdet!qty & "</td>"
fs.WriteLine "     <td align=right>&nbsp;" & Format(invdet!unitrate, "###,###,##0.00") & "</td>"
fs.WriteLine "     <td align=right>&nbsp;" & Format(invdet!amount, "###,###,##0.00") & "</td>"

qty = qty + invdet("qty")
urate = urate + invdet!unitrate
amt = amt + invdet!amount
fs.WriteLine "   </tr>"
invdet.MoveNext
Wend
For i = 0 To (20 - rc)
fs.WriteLine "   <tr valign=0 class=TableFont>"
fs.WriteLine "   <td >&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "     <td>&nbsp;</td>"
fs.WriteLine "   </tr>"
Next
End If
fs.WriteLine "   <tr  height=28 class=TableFont>"
fs.WriteLine "     <td colspan=4><b>Total </td>"
fs.WriteLine "     <td align=right><b>" & Format(qty, "###,###,##0.00") & "</td>"
fs.WriteLine "     <td align=right><b>" & Format(urate, "###,###,##0.00") & "</td>"
fs.WriteLine "     <td align=right><b>" & Format(amt, "###,###,##0.00") & "</td>"
fs.WriteLine "   </tr>"
 


fs.WriteLine "</table>"
WebBrowser.Navigate App.Path & "\rep.html"
fs.WriteLine "<p>&nbsp;</p>"
fs.WriteLine "</body>"
fs.WriteLine "</html>"



End Sub




