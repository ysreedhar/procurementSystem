VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form Report_BidEvaluation 
   BackColor       =   &H8000000E&
   Caption         =   "BID Evaluation"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_show 
      BackColor       =   &H00C0FFC0&
      Caption         =   "View"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmd_print 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Print"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmd_close 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Close"
      Height          =   255
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cbo_rfqno 
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
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   8745
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   14745
      ExtentX         =   26009
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Report_BidEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_show_Click()
Call BidEvaluationProc
End Sub

Public Sub BidEvaluationProc()

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
fs.WriteLine " <td width=618 height=37 align=left><font size=3 FAMILY=Tahoma><b><u>BID Evaluation</u></b></td>"
fs.WriteLine " </tr>"
fs.WriteLine " </table>" 'header


fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569

fs.WriteLine "  <tr bgcolor=black class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=1><font color=white>Item Description</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Qty</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>UOM</td>"
Dim Vven(25) As String
Dim Vcnt As Integer
Dim vend As New ADODB.Recordset
If vend.State Then vend.Close
vend.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
For i = 0 To vend.RecordCount - 1
Vcnt = vend.RecordCount - 1
Vven(i) = vend(1)
fs.WriteLine "    <td nowrap colspan=1><font color=white>" & vend(0) & "</td>"
vend.MoveNext
Next i
fs.WriteLine "  </tr>" 'item header


'First
Dim cnt As Integer

cnt = 1
Dim rfqd As New ADODB.Recordset
If rfqd.State Then rfqd.Close
rfqd.Open "select DISTINCT(p.location) from quotation q,prdetails p where q.rfqno=p.rfqno  and q.rfqno='" & cbo_rfqno.Text & "'", Cn, 3, 2
  
        While Not rfqd.EOF
        
fs.WriteLine "  <tr bgcolor=white class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=3>Delivery To : & " & rfqd(0) & "</td>"
fs.WriteLine "  </tr>" 'item header
                         
    Dim Vprid As Double
    Dim ven As New ADODB.Recordset
    If ven.State Then ven.Close
                            
    ven.Open "select DISTINCT(material),qty,uom,pr_id from prdetails  where rfqno='" & cbo_rfqno.Text & "' and location='" & rfqd(0) & "' ", Cn, 3, 2
    While Not ven.EOF

fs.WriteLine "  <tr bgcolor=white class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=1>" & ven(0) & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & ven(1) & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & ven(2) & "</td>"
Vprid = ven(3)

    i = 0
    
    For i = 0 To Vcnt
    Dim vn As New ADODB.Recordset
    If vn.State Then vn.Close
    vn.Open "select status from  quotationdetails where vendor ='" & Vven(i) & "'  and  rfqno='" & cbo_rfqno.Text & "'  and prid='" & Vprid & "' ", Cn, 3, 2
    If Not vn.EOF Then
    fs.WriteLine "    <td nowrap colspan=1>" & vn(0) & "</td>"
    End If
    Next i
    ven.MoveNext
    Wend
    ven.Close
     
    fs.WriteLine "  </tr>"
                          
                          
                          
                                                      
        rfqd.MoveNext
    Wend
    rfqd.Close
fs.WriteLine "</table>"

''price

fs.WriteLine "  <tr bgcolor=black class=TableFont ><br><br></tr>"
fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569

fs.WriteLine "  <tr bgcolor=black class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=1><font color=white>Item Description</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Qty</td>"
fs.WriteLine "    <td nowrap colspan=1><font color=white>UOM</td>"
Dim Vvenb(25) As String
Dim Vcntb As Integer
Dim ib As Integer
ib = 0
Dim vendb As New ADODB.Recordset
If vendb.State Then vendb.Close
vendb.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
For ib = 0 To vendb.RecordCount - 1
Vcntb = vendb.RecordCount - 1
Vvenb(ib) = vendb(1)
fs.WriteLine "    <td nowrap colspan=1><font color=white>" & vendb(0) & "</td>"
vendb.MoveNext
Next ib
fs.WriteLine "  </tr>" 'item header


'First
 
Dim rfqdb As New ADODB.Recordset
If rfqdb.State Then rfqdb.Close
rfqdb.Open "select DISTINCT(p.location) from quotation q,prdetails p where q.rfqno=p.rfqno  and q.rfqno='" & cbo_rfqno.Text & "'", Cn, 3, 2
  
        While Not rfqdb.EOF
        
fs.WriteLine "  <tr bgcolor=white class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=3>Delivery To : & " & rfqdb(0) & "</td>"
fs.WriteLine "  </tr>" 'item header
                         
    Dim Vpridb As Double
    Dim venb As New ADODB.Recordset
    If venb.State Then venb.Close
                            
    venb.Open "select DISTINCT(material),qty,uom,pr_id from prdetails  where rfqno='" & cbo_rfqno.Text & "' and location='" & rfqdb(0) & "' ", Cn, 3, 2
    While Not venb.EOF

fs.WriteLine "  <tr bgcolor=white class=TableFont >"
fs.WriteLine "    <td Nowrap colspan=1>" & venb(0) & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & venb(1) & "</td>"
fs.WriteLine "    <td nowrap colspan=1>" & venb(2) & "</td>"
Vpridb = venb(3)

    ib = 0
    
    For ib = 0 To Vcnt
    Dim vnb As New ADODB.Recordset
    If vnb.State Then vnb.Close
    vnb.Open "select (unitrate),amount from  quotationdetails where vendor ='" & Vven(ib) & "'  and  rfqno='" & cbo_rfqno.Text & "'  and prid='" & Vpridb & "' ", Cn, 3, 2
    If Not vnb.EOF Then
    fs.WriteLine "    <td nowrap colspan=1>" & vnb(0) & "-" & vnb(1) & "</td>"
    End If
    Next ib
    venb.MoveNext
    Wend
    venb.Close
                          
    fs.WriteLine "  </tr>"
                          
                          
                          
                                                      
        rfqdb.MoveNext
    Wend
    rfqdb.Close
fs.WriteLine "</table>"

'price --end

'terms
               
fs.WriteLine "  <tr bgcolor=black class=TableFont ><br><br></tr>"
fs.WriteLine " <table width=623  border=1 cellpadding=0 cellspacing=0 bordercolor=#bcbcbc>" 'height=569

fs.WriteLine "  <tr bgcolor=black class=TableFont >"
fs.WriteLine "    <td nowrap colspan=1><font color=white>Terms</td>"
Dim Vvenc(25) As String
Dim Vcntc As Integer
Dim ic As Integer
ic = 0
Dim vendc As New ADODB.Recordset
If vendc.State Then vendc.Close
vendc.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
For ic = 0 To vendc.RecordCount - 1
Vcntc = vendc.RecordCount - 1
Vvenc(ic) = vendc(0)

fs.WriteLine "    <td nowrap colspan=1><font color=white>" & vendc(0) & "</td>"
vendc.MoveNext
Next ic
fs.WriteLine "  </tr>"
ic = 0
For ic = 0 To Vcnt
                  Dim vnc As New ADODB.Recordset
                  If vnc.State Then vnc.Close
                  vnc.Open "select DISTINCT(qt.terms),qt.termsdesc from quotation q, quotationterms qt ,vendor v where q.qno=qt.qno and q.vendor=v.name and v.code ='" & Vvenc(ic) & "' and qt.chq='Yes' and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
                  While Not vnc.EOF
                  fs.WriteLine "    <tr bgcolor=white class=TableFont>"
                  fs.WriteLine "    <td nowrap colspan=1>" & vnc(0) & "</td>"
                  fs.WriteLine "    <td nowrap colspan=1>" & vnc(1) & "</td>"
                  fs.WriteLine "  </tr>"
                  vnc.MoveNext
                  Wend

Next ic

fs.WriteLine "</table>"

WebBrowser.Navigate App.Path & "\rep.html"
fs.WriteLine "<p>&nbsp;</p>"
fs.WriteLine "</body>"
fs.WriteLine "</html>"


End Sub

Private Sub Form_Load()
Call connect
main.lbltitle.Caption = "BID Evaluation"
 
Me.Top = 10
Me.Left = 10
 Dim rfq As New ADODB.Recordset
If rfq.State Then rfq.Close
rfq.Open "select DISTINCT(rfqno) from quotation where award ='No' order by rfqno", Cn, 3, 2
While Not rfq.EOF
cbo_rfqno.AddItem rfq(0)
rfq.MoveNext
Wend
rfq.Close
WebBrowser.Navigate "About:Blank"
End Sub
