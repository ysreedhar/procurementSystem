VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form report_purchaseinforecord 
   BackColor       =   &H8000000E&
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo_material 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   8415
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Height          =   255
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   255
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View"
      Height          =   255
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1455
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
      Format          =   67371009
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
      Format          =   67371009
      CurrentDate     =   38674
   End
End
Attribute VB_Name = "report_purchaseinforecord"
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
main.lbltitle.Caption = "Purchase InfoRecords"
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
fs.WriteLine " <td width=618 height=37 align=left><font size=3 FAMILY=Tahoma><b><u>PURCHASE INFO-RECORDS between  </u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_from.Value, "dd/MMM/yy") & " and</u></b><b><u><font size=1 FAMILY=Tahoma>" & Format(dtp_to.Value, "dd/MMM/yy") & " </u></b></td>"
'fs.WriteLine " <td width=618 height=37 align=left><font size=1 FAMILY=Tahoma><b><u>" & Format(dtp_from.Value, "dd/MMM/yy") & " and</u></b><b><u>" & Format(dtp_to.Value, "dd/MMM/yy") & " </u></b></td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'header


fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569

fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap colspan=2><font color=white>Item Description</td>"
fs.WriteLine "    <td nowrap colspan=2><font color=white>Item Code</td>"
fs.WriteLine "    <td nowrap colspan=2><font color=white>Mfr. Ref Code</td>"
fs.WriteLine "  </tr>" 'item header
mn = Split(cbo_material.Text, "  -  ", Len(cbo_material.Text), vbTextCompare)
Dim StrMn As String
StrMn = ""
StrMn = mn(0) & "  -  " & mn(1)
fs.WriteLine "  <tr bgcolor=white class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap colspan=2>" & StrMn & "</td>"
fs.WriteLine "    <td nowrap colspan=2>" & mn(2) & "</td>"
fs.WriteLine "    <td nowrap colspan=2>" & mn(3) & "</td>"
fs.WriteLine "  </tr>" 'item header



fs.WriteLine "  <tr bgcolor=black class=TableFont bgcolor=black>"
fs.WriteLine "    <td Nowrap align=center><font color=white>S.No</td>"
fs.WriteLine "    <td nowrap><font color=white>PO No.</td>"
fs.WriteLine "    <td nowrap><font color=white>PO Date</td>"
fs.WriteLine "   <td nowrap><font color=white>Amount</td>"
fs.WriteLine "    <td nowrap><font color=white>Vendor</td>"
fs.WriteLine "  </tr>" 'item header

Dim sn As Integer

sn = 1


Dim pif As New ADODB.Recordset
If pif.State Then pif.Close
pif.Open "select DISTINCT(pd.pono),SUM(pd.amount),p.podate,p.vendor from podetails pd,po p where p.pono=pd.pono and pd.material='" & StrMn & "' and p.podate between '" & Format(dtp_from.Value, "mm/dd/yyyy") & "' and '" & Format(dtp_to.Value, "mm/dd/yyyy") & "' Group By pd.pono,p.podate,p.vendor", Cn, 3, 2
While Not pif.EOF
fs.WriteLine "   <tr valign=0 class=TableFont >"
fs.WriteLine "   <td align=center>" & sn & "</td>"
fs.WriteLine "     <td nowrap>" & pif(0) & "</td>"
fs.WriteLine "     <td nowrap>" & pif(2) & "</td>"
fs.WriteLine "     <td align=left nowrap>" & pif(1) & "</td>"
fs.WriteLine "     <td nowrap>" & pif(3) & "</td>"
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
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4  order by ml4name", Cn, 3, 2 'where ml4type= '" & mty(0) & "'
While Not med.EOF
cbo_material.AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close
End Sub

