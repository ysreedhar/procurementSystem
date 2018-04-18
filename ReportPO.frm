VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form ReportPO 
   BackColor       =   &H8000000E&
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   9345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9345
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11535
      Begin VB.ComboBox cbo_po 
         Height          =   315
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.CommandButton cmd_po 
         Caption         =   "View PO"
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtp_po 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   28377091
         CurrentDate     =   38662
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Month/Year"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "PO No."
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   120
         Width           =   525
      End
   End
End
Attribute VB_Name = "ReportPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_po_Change()
Call loadpodetails
End Sub
Private Sub cmd_po_Click()
Call loadpodetails
End Sub
Private Sub dtp_po_Click()
Call loadpo
End Sub
Private Sub Form_Load()
dtp_po.Value = Format(Date, "MMM/yyyy")
Call loadpo
End Sub
Public Sub loadpo()
Dim msr As New ADODB.Recordset
If msr.State Then msr.Close
msr.Open "select DISTINCT(pono) from po where Month(podate)='" & Format(dtp_po.Value, "MM") & "' and Year(podate)='" & Format(dtp_po.Value, "yyyy") & "' ", Cn, 3, 2
While Not msr.EOF
cbo_po.AddItem msr(0)
msr.MoveNext
Wend
msr.Close
End Sub
Public Sub loadpodetails()
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
rstemp.Fields.Append "lblpo", adVarChar, 100
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
pr.Open "select * from po where pono ='" & cbo_po.Text & "'", Cn, 3, 2
If Not pr.EOF Then
                 Dim vn As New ADODB.Recordset
                 If vn.State Then vn.Close
                 vn.Open "select * from vendor where name='" & pr!vendor & "' ", Cn, 3, 2
                 If Not vn.EOF Then
    RptPO.Sections("section4").Controls("lblname").Caption = "Attn: Mr." & vn!personincharge
    RptPO.Sections("section4").Controls("lbladdress").Caption = vn!Name & vn!Address
                 End If
    RptPO.Sections("section4").Controls("lblpo").Caption = pr("pono")
    RptPO.Sections("section4").Controls("lbldate").Caption = Format(pr("podate"), "dd/MMM/yy")
End If
pr.Close
amt = 0: upr = 0
pr.Open "select * from podetails where pono='" & cbo_po.Text & "' ", Cn, 3, 2
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


RptPO.Sections("section5").Controls("lbluprice").Caption = Format(upr, "###,###,##0")
RptPO.Sections("section5").Controls("lbltprice").Caption = Format(amt, "###,###,##0")


'promised date
Dim j As Integer
j = 0
Dim cnt As String
Dim jbc As New ADODB.Recordset
If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(promisedate) from podetails where pono= '" & cbo_po.Text & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptPO.Sections("section5").Controls("lbljobcharge").Caption = "All Items to be Delivered on " & jbc(0)
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

RptPO.Sections("section5").Controls("lbljobcharge").Caption = "Promised Delivery date for " & cnt
End If

'location
j = 0
cnt = ""
jbc.Close
'If jbc.State Then jbc.Close
jbc.Open "select DISTINCT(location) from podetails where pono= '" & cbo_po.Text & "'", Cn, 3, 2
If jbc.RecordCount <= 1 Then
RptPO.Sections("section5").Controls("lbllocation").Caption = "All Items to be Delivered @ " & jbc(0)
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
RptPO.Sections("section5").Controls("lbllocation").Caption = "Delivery for " & cnt
End If
rstemp.Update
pr.Close

Set RptPO.DataSource = rstemp
RptPO.Show
SetParent RptPO.hwnd, ReportPO.hwnd

End Sub



