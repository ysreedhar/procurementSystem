VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form quotationselection 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Material & Service Evaluation"
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   10155
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_refresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   6600
      TabIndex        =   14
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9540
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   19000
      _ExtentX        =   33523
      _ExtentY        =   16828
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Technical"
      TabPicture(0)   =   "quotationselection.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Price"
      TabPicture(1)   =   "quotationselection.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other Aspects"
      TabPicture(2)   =   "quotationselection.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame3 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   -75000
         TabIndex        =   6
         Top             =   300
         Width           =   15015
         Begin VB.CheckBox chk_app2 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid flex_grid2 
            Height          =   9195
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   19000
            _ExtentX        =   33523
            _ExtentY        =   16219
            _Version        =   393216
            Rows            =   3
            Cols            =   6
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            BackColorBkg    =   16777215
            SelectionMode   =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   -75000
         TabIndex        =   5
         Top             =   300
         Width           =   15735
         Begin VB.CheckBox Chk_app1 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chk_col1 
            BackColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSFlexGridLib.MSFlexGrid flex_grid1 
            Height          =   9075
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   19000
            _ExtentX        =   33523
            _ExtentY        =   16007
            _Version        =   393216
            Rows            =   3
            Cols            =   9
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            BackColorBkg    =   16777215
            SelectionMode   =   1
            AllowUserResizing=   3
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   9255
         Left            =   0
         TabIndex        =   3
         Top             =   300
         Width           =   15015
         Begin VB.CheckBox chk_col 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid flex_grid 
            Height          =   9075
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   19000
            _ExtentX        =   33523
            _ExtentY        =   16007
            _Version        =   393216
            Rows            =   3
            FixedCols       =   0
            RowHeightMin    =   300
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            BackColorBkg    =   16777215
            SelectionMode   =   1
            AllowUserResizing=   3
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.CommandButton cmd_supplyorder 
      Caption         =   "Award Supply Order"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2175
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
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "quotationselection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public st As String
Private Sub cbo_rfqno_Change()
flex_grid.Clear
flex_grid1.Clear
flex_grid2.Clear
cbo_rfqno.Enabled = False
Call flexdisplay
Call flexdisplay1
Call flexdisplay2

End Sub
Private Sub cbo_rfqno_Click()
flex_grid.Clear
flex_grid1.Clear
flex_grid2.Clear
cbo_rfqno.Enabled = False
Call flexdisplay
Call flexdisplay1
Call flexdisplay2
End Sub
Private Sub chk_app_Click(Index As Integer)
If chk_app(Index) = 1 Then
            Chk_app1(Index) = 1
            Dim ind1 As Integer
            ind1 = 6
                    For ind1 = 6 To Index - 1
                    Chk_app1(ind1) = 0
                    Next ind1
                ind = 6
                    For ind1 = (Index + 1) To flex_grid.Cols - 1
                    Chk_app1(ind1) = 0
                    Next ind1
                    
                    'terms
               chk_app2(Index) = 1
                        ind = 6
                    For ind = 6 To Index - 1
                    chk_app2(ind) = 0
                    Next ind
                ind = 6
                    For ind = (Index + 1) To flex_grid2.Cols - 1
                    chk_app2(ind) = 0
                    Next ind
       Else
            Chk_app1(Index) = 0
       End If
End Sub
Private Sub Chk_app1_Click(Index As Integer)

        If Chk_app1(Index) = 1 Then
            chk_app(Index) = 1
            Dim ind As Integer
            ind = 6
                    For ind = 6 To Index - 1
                    chk_app(ind) = 0
                    Next ind
                ind = 6
                    For ind = (Index + 1) To flex_grid.Cols - 1
                    chk_app(ind) = 0
                    Next ind
                    
               'terms
               chk_app2(Index) = 1
                        ind = 6
                    For ind = 6 To Index - 1
                    chk_app2(ind) = 0
                    Next ind
                ind = 6
                    For ind = (Index + 1) To flex_grid2.Cols - 1
                    chk_app2(ind) = 0
                    Next ind
                    
       Else
            chk_app(Index) = 0
            chk_app2(Index) = 0
       End If

End Sub

Private Sub chk_app2_Click(Index As Integer)

        If chk_app2(Index) = 1 Then
            chk_app(Index) = 1
            Dim ind As Integer
            ind = 6
                    For ind = 6 To Index - 1
                    chk_app(ind) = 0
                    Next ind
                ind = 6
                    For ind = (Index + 1) To flex_grid.Cols - 1
                    chk_app(ind) = 0
                    Next ind
                    
               'terms
               Chk_app1(Index) = 1
                        ind = 6
                    For ind = 6 To Index - 1
                    Chk_app1(ind) = 0
                    Next ind
                ind = 6
                    For ind = (Index + 1) To flex_grid1.Cols - 1
                    Chk_app1(ind) = 0
                    Next ind
                    
       Else
            chk_app(Index) = 0
            Chk_app1(Index) = 0
       End If

End Sub

Private Sub chk_col_Click(Index As Integer)
If chk_col(Index) = 1 Then
            chk_col1(Index) = 1
            Dim ind As Integer
       Else
            chk_col1(Index) = 0
       End If
End Sub

Private Sub chk_col1_Click(Index As Integer)
If chk_col1(Index) = 1 Then
            chk_col(Index) = 1
            Dim ind As Integer
          
       Else
            chk_col(Index) = 0
       End If

End Sub

Private Sub cmd_refresh_Click()
flex_grid.Clear
flex_grid1.Clear
flex_grid2.Clear


flex_grid.Rows = 1
flex_grid1.Rows = 1
flex_grid2.Rows = 1
cbo_rfqno.Enabled = True
End Sub

Private Sub cmd_supplyorder_Click()
On Error Resume Next
Call generatepo
MsgBox "Supply Order Awarded to : " & st

Dim a As Integer
a = 0
        For a = 6 To flex_grid.Cols - 1
        chk_app(a).Value = 0
        Next a
a = 0
        For a = 2 To flex_grid.Rows - 1
        chk_col(a).Value = 0
        Next a
        
a = 0
        For a = 6 To flex_grid1.Cols - 1
        Chk_app1(a).Value = 0
        Next a
a = 0
        For a = 2 To flex_grid1.Rows - 1
        chk_col1(a).Value = 0
        Next a
        
a = 0
        For a = 6 To flex_grid1.Cols - 1
        chk_app2(a).Value = 0
        Next a

        
End Sub


Private Sub flex_grid_Click()
On Error Resume Next
'back color
 
Static vprev As Integer

current = flex_grid.Row

'Reset to previous row
If vprev > 0 Then
    flex_grid.Row = vprev
    flex_grid.Col = 1
    Set flex_grid.CellPicture = LoadPicture()
    
    For i = 1 To flex_grid.Cols - 1
    flex_grid.Col = i
    flex_grid.CellBackColor = vbWhite
Next
End If

If flex_grid.Row <> 0 Then
'Current  row
flex_grid.Row = current
For i = 1 To flex_grid.Cols - 1
flex_grid.Col = i
flex_grid.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_grid.Col = 1
End If

vprev = flex_grid.Row

End Sub

Private Sub flex_grid_SelChange()
flex_grid.ToolTipText = flex_grid.TextMatrix(flex_grid.Row, flex_grid.Col)
End Sub

Private Sub Form_Load()
Dim rfq As New ADODB.Recordset
If rfq.State Then rfq.Close
rfq.Open "select DISTINCT(rfqno) from quotation where award ='No' order by rfqno", Cn, 3, 2
While Not rfq.EOF
cbo_rfqno.AddItem rfq(0)
rfq.MoveNext
Wend
rfq.Close

Call flex_title
End Sub
Public Sub flexdisplay()
Dim cnt As Integer

cnt = 1
Dim rfqd As New ADODB.Recordset
If rfqd.State Then rfqd.Close
rfqd.Open "select DISTINCT(p.location) from quotation q,prdetails p where q.rfqno=p.rfqno  and q.rfqno='" & cbo_rfqno.Text & "'", Cn, 3, 2
With flex_grid
    
        While Not rfqd.EOF
        .Rows = .Rows + 1
          
                         .TextMatrix(.Rows - 1, 4) = "Delivery To :" & rfqd(0)
                         .Col = 4
                         .Row = .Rows - 1
                         .CellFontBold = True
                          Dim ven As New ADODB.Recordset
                          If ven.State Then ven.Close
                                                  
                          ven.Open "select DISTINCT(material),qty,uom,pr_id from prdetails where rfqno='" & cbo_rfqno.Text & "' and location='" & rfqd(0) & "' ", Cn, 3, 2
                          While Not ven.EOF
                          cnt = cnt + 1
                          .Rows = .Rows + 1
                          
'check box to select items
On Error Resume Next
Load chk_col(.Rows - 1)
.Col = 1
.Row = .Rows - 1
chk_col(.Rows - 1).Left = .Left + .CellLeft
chk_col(.Rows - 1).Top = .Top + .CellTop
chk_col(.Rows - 1).Height = .CellHeight
chk_col(.Rows - 1).Width = .CellWidth
chk_col(.Rows - 1).ZOrder 0
chk_col(.Rows - 1).Visible = True

'-------------------------------
                                .TextMatrix(.Rows - 1, 0) = ven(3)
                                .TextMatrix(.Rows - 1, 2) = ven(2)
                                .TextMatrix(.Rows - 1, 3) = ven(1)
                                .TextMatrix(.Rows - 1, 4) = ven(0)
                          ven.MoveNext
                          Wend
                          ven.Close
                          
                          
                          cnt = 4
                          ven.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
                          While Not ven.EOF
                          cnt = cnt + 1
                          Dim i As Integer
                          i = 0
                          Dim vn As New ADODB.Recordset
                          If vn.State Then vn.Close
                          
                          For i = 2 To flex_grid.Rows - 1
                          vn.Open "select qd.status from quotation q, quotationdetails qd , prdetails p where q.qno=qd.qno and qd.prno=p.prno and q.vendor ='" & ven(1) & "'  and  q.rfqno='" & cbo_rfqno.Text & "' and p.pr_id='" & flex_grid.TextMatrix(i, 0) & "' ", Cn, 3, 2
                          If Not vn.EOF Then
                          If InStr(flex_grid.TextMatrix(i, 4), "Delivery To :") Then
                          i = i + 1
                          End If
                          flex_grid.TextMatrix(i, cnt) = vn(0)
                          
                          vn.MoveNext
                          End If
                          Next i
                          
                          If cnt = 1 Then
                              .TextMatrix(0, 5) = ven(0)
                                    On Error Resume Next
                                    Load chk_app(.Cols - 1)
                                    .Col = 5
                                    .Row = 0
                                    chk_app(.Cols - 1).Left = .Left + .CellLeft
                                    chk_app(.Cols - 1).Top = .Top + .CellTop
                                    chk_app(.Cols - 1).Height = .CellHeight
                                    chk_app(.Cols - 1).Width = .CellWidth
                                    chk_app(.Cols - 1).ZOrder 0
                                    chk_app(.Cols - 1).Visible = True
                    chk_app(.Cols - 1).Caption = ven(0)
                              
                          ElseIf cnt > 1 Then
                          .Cols = .Cols + 1
                             .TextMatrix(0, cnt) = ven(0)
                             .ColWidth(cnt) = 1000
                                             
                                             
                     On Error Resume Next
                     Load chk_app(.Cols - 1)
                    .Col = cnt
                    .Row = 0
                    chk_app(.Cols - 1).Left = .Left + .CellLeft
                    chk_app(.Cols - 1).Top = .Top + .CellTop
                    chk_app(.Cols - 1).Height = .CellHeight
                    chk_app(.Cols - 1).Width = .CellWidth
                    chk_app(.Cols - 1).ZOrder 0
                    chk_app(.Cols - 1).Visible = True
                    chk_app(.Cols - 1).Caption = ven(0)
                    End If
                          
                         
                          ven.MoveNext
                          Wend
                          ven.Close
                          
                                                      
        rfqd.MoveNext
    Wend
    rfqd.Close
End With

End Sub

Public Sub flex_title()
With flex_grid
    .Rows = 1
    .Cols = 6
    .ColWidth(0) = 0
    .ColWidth(1) = 300
    .ColWidth(2) = 500
    .ColWidth(3) = 500
    .ColWidth(4) = 2500
    .ColWidth(5) = 0
                    .TextMatrix(0, 1) = ""
                    .TextMatrix(0, 2) = "Qty"
                    .TextMatrix(0, 3) = "UOM"
                    .TextMatrix(0, 4) = "Description"
                    
        
 End With
 With flex_grid1
    .Rows = 1
    .Cols = 6
    .ColWidth(0) = 0
    .ColWidth(1) = 300
    .ColWidth(2) = 500
    .ColWidth(3) = 500
    .ColWidth(4) = 2500
    .ColWidth(5) = 0
                    .TextMatrix(0, 1) = ""
                    .TextMatrix(0, 2) = "Qty"
                    .TextMatrix(0, 3) = "UOM"
                    .TextMatrix(0, 4) = "Description"
                         
End With
End Sub
Public Sub flexdisplay1()
'Call flex_title
Dim cnt As Integer
cnt = 1
Dim rfqd As New ADODB.Recordset
If rfqd.State Then rfqd.Close
rfqd.Open "select DISTINCT(p.location) from quotation q,prdetails p where q.rfqno=p.rfqno  and q.rfqno='" & cbo_rfqno.Text & "'", Cn, 3, 2
With flex_grid1
    
        While Not rfqd.EOF
        .Rows = .Rows + 1
                         
                         .TextMatrix(.Rows - 1, 4) = "Delivery To :" & rfqd(0)
                         .Col = 4
                         .Row = .Rows - 1
                         .CellFontBold = True
                          
                          Dim ven As New ADODB.Recordset
                          If ven.State Then ven.Close
                          ven.Open "select DISTINCT(material),qty,uom,pr_id from prdetails  where rfqno='" & cbo_rfqno.Text & "' and location='" & rfqd(0) & "' ", Cn, 3, 2
                          While Not ven.EOF
                          cnt = cnt + 1
                          .Rows = .Rows + 1
                          
'check box to select items
On Error Resume Next
Load chk_col1(.Rows - 1)
.Col = 1
.Row = .Rows - 1
chk_col1(.Rows - 1).Left = .Left + .CellLeft
chk_col1(.Rows - 1).Top = .Top + .CellTop
chk_col1(.Rows - 1).Height = .CellHeight
chk_col1(.Rows - 1).Width = .CellWidth
chk_col1(.Rows - 1).ZOrder 0
chk_col1(.Rows - 1).Visible = True

'-------------------------------
                                .TextMatrix(.Rows - 1, 0) = ven(3)
                                .TextMatrix(.Rows - 1, 2) = ven(2)
                                .TextMatrix(.Rows - 1, 3) = ven(1)
                                .TextMatrix(.Rows - 1, 4) = ven(0)
                          ven.MoveNext
                          Wend
                          ven.Close
                          
                          cnt = 4
                          ven.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
                          While Not ven.EOF
                          cnt = cnt + 1
                          Dim i As Integer
                          i = 0
                          Dim vn As New ADODB.Recordset
                          If vn.State Then vn.Close

                          For i = 2 To flex_grid1.Rows - 1
                          vn.Open "select (qd.unitrate),qd.amount from quotation q, quotationdetails qd,prdetails p where q.qno=qd.qno and qd.prno=p.prno and q.vendor ='" & ven(1) & "'  and  q.rfqno='" & cbo_rfqno.Text & "' and p.pr_id='" & flex_grid1.TextMatrix(i, 0) & "'", Cn, 3, 2
                          If Not vn.EOF Then
                          If InStr(flex_grid1.TextMatrix(i, 4), "Delivery To :") Then
                          i = i + 1
                          End If
                          flex_grid1.TextMatrix(i, cnt) = vn(0) & "  -  " & vn(1)
                          vn.MoveNext
                          End If
                          Next i
                          If cnt = 1 Then
                              .TextMatrix(0, 5) = ven(0)
                              .TextMatrix(1, 5) = "U/P   -   T/P"
    On Error Resume Next
    Load Chk_app1(.Cols - 1)
    .Col = cnt
    .Row = 0
    Chk_app1(.Cols - 1).Left = .Left + .CellLeft
    Chk_app1(.Cols - 1).Top = .Top + .CellTop
    Chk_app1(.Cols - 1).Height = .CellHeight
    Chk_app1(.Cols - 1).Width = .CellWidth
    Chk_app1(.Cols - 1).ZOrder 0
    Chk_app1(.Cols - 1).Visible = True
    Chk_app1(.Cols - 1).Caption = ven(0)
    Chk_app1(.Cols - 1).Alignment = Left
                          ElseIf cnt > 1 Then
                          .Cols = .Cols + 1
                             .TextMatrix(0, cnt) = ven(0)
                             .TextMatrix(1, cnt) = "U/P   -   T/P"
                             .ColWidth(cnt) = 1200
    On Error Resume Next
    Load Chk_app1(.Cols - 1)
    .Col = cnt
    .Row = 0
    Chk_app1(.Cols - 1).Left = .Left + .CellLeft
    Chk_app1(.Cols - 1).Top = .Top + .CellTop
    Chk_app1(.Cols - 1).Height = .CellHeight
    Chk_app1(.Cols - 1).Width = .CellWidth
    Chk_app1(.Cols - 1).ZOrder 0
    Chk_app1(.Cols - 1).Visible = True
    Chk_app1(.Cols - 1).Caption = ven(0)
    Chk_app1(.Cols - 1).Alignment = Left

                                             
                          End If
                          ven.MoveNext
                          Wend
                          ven.Close
                          
                                                      
        rfqd.MoveNext
    Wend
    rfqd.Close
End With

End Sub

Public Sub generatepo()
Dim poid As Double
poid = 0
Dim k As Integer
k = 0
For k = 6 To flex_grid.Cols - 1
        If chk_app(k).Value = 1 Then
           st = chk_app(k).Caption
        End If
Next k


'-----------------------
Dim vd As New ADODB.Recordset
If vd.State Then vd.Close
vd.Open "select name from vendor where code='" & st & "' ", Cn, 3, 2
If Not vd.EOF Then
st = vd(0)
End If

Dim i As Integer
   
   i = 1
assad:
Dim X As String
Dim t As String
t = ""
X = "TL-PO000" & i

'Cn.Execute "delete from po where pono='" & X & "'"
'Cn.Execute "delete from podetails where pono='" & X & "'"
'-----------------------
Dim qtid As Double
qtid = 0
Dim qt As New ADODB.Recordset
If qt.State Then qt.Close
qt.Open "select * from quotation where rfqno='" & cbo_rfqno.Text & "' and vendor ='" & st & "'", Cn, 3, 2
If Not qt.EOF Then
qtid = qt!q_id
                   Dim rff As New ADODB.Recordset
                   If rff.State Then rff.Close
                   rff.Open "select * from po where pono='" & X & "'", Cn, 3, 2
                   If rff.EOF Then
                   rff.AddNew
                            rff!pono = "TL-PO000" & i
                            t = "TL-PO000" & i
                            rff!qno = qt!qno
                            rff!vendor = qt!vendor
                            rff!contactperson = qt!contactperson
                            rff!podate = Format(Date, "dd/MM/yyyy")
                            rff!toperson = ""
                            rff!desig = ""
                            rff!dept = ""
                            rff!Mode = ""
                            rff!refno = ""
                            rff!oref = ""
                            rff!yref = ""
                            rff!Notes = ""
                            rff!Status = "Pending"
                   rff.Update
                   
                   
                   Else
                   i = i + 1
                 
                GoTo assad
                  End If
                  rff.Close
  rff.Open "select MAX(po_id) from po where pono='" & t & "'", Cn, 3, 2
  If Not rff.EOF Then
  poid = rff(0)
  End If
  rff.Close
  qt!award = "Yes"
  qt.Update
  
 End If
 '--------------------------------------------------------------------------------

Dim j As Integer
j = 0
k = 0
For k = 6 To flex_grid.Cols - 1
        If chk_app(k).Value = 1 Then
            For j = 2 To flex_grid.Rows - 1
               If chk_col(j).Value = 1 Then
            Cn.Execute "update quotationdetails set award='Yes' , pono ='" & t & "' where vendor='" & st & "' and material ='" & flex_grid.TextMatrix(j, 4) & "'"
               End If
            Next j
        'st = chk_app(k).Caption
        End If
Next k

 
 
 '--------------------------------------------------------------------------------
 
 
 
         Dim qd As New ADODB.Recordset
         If qd.State Then qd.Close
         qd.Open "select * from podetails", Cn, 3, 2
 qt.Close
 qt.Open "select * from quotationdetails where award ='Yes' and pono='" & t & "' and vendor ='" & st & "' and qtid =" & qtid, Cn, 3, 2
 While Not qt.EOF
                  
                        qd.AddNew
                        qd!pono = t
                        qd!qno = qt!qno
                        qd!Status = qt!Status
                        qd!itemid = qt!itemid
                        qd!mrefcode = qt!mrefcode
                        qd!material = qt!material
                        qd!qty = qt!qty
                        qd!uom = qt!uom
                        qd!unitrate = qt!unitrate
                        qd!Currency = qt!Currency
                        qd!xchg = qt!xchg
                        qd!amount = qt!amount
                        qd!reqdate = qt!reqdate
                        qd!promisedate = qt!promisedate
                        qd!remarks = qt!remarks
                        qd!poid = poid
                        qd!postatus = "Pending"
                        qd!prno = qt!prno
                        qd!rfqno = qt!rfqno
                        qd.Update

qt.MoveNext
Wend
qt.Close

End Sub
Public Sub flexdisplay2()
With flex_grid2
                .Rows = 5
                .ColWidth(0) = 2000
                .ColWidth(1) = 0
                .ColWidth(2) = 0
                .ColWidth(3) = 0
                .ColWidth(4) = 0
                .TextMatrix(1, 0) = "Delivery Terms"
                .TextMatrix(2, 0) = "Delivery Period"
                .TextMatrix(3, 0) = "Payment Terms"
                .TextMatrix(4, 0) = "Quotation Validity"
End With
With flex_grid2


                  cnt = 4
                  Dim ven As New ADODB.Recordset
                  If ven.State Then ven.Close
                  ven.Open "select DISTINCT(v.code),v.name from quotation q , vendor v where q.vendor=v.name  and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
                  While Not ven.EOF
                  cnt = cnt + 1
                  
                  If cnt = 1 Then
                           .TextMatrix(0, 1) = ven(0)
                           .ColWidth(1) = 2000
                            On Error Resume Next
                            Load chk_app2(.Cols - 1)
                            .Col = cnt
                            .Row = 0
                            chk_app2(.Cols - 1).Left = .Left + .CellLeft
                            chk_app2(.Cols - 1).Top = .Top + .CellTop
                            chk_app2(.Cols - 1).Height = .CellHeight
                            chk_app2(.Cols - 1).Width = .CellWidth
                            chk_app2(.Cols - 1).ZOrder 0
                            chk_app2(.Cols - 1).Visible = True
                            chk_app2(.Cols - 1).Caption = ven(0)
                      
                  ElseIf cnt > 1 Then
                     .Cols = .Cols + 1
                     .TextMatrix(0, cnt) = ven(0)
                     .ColWidth(cnt) = 2000
                                     
                                     
                    On Error Resume Next
                    Load chk_app2(.Cols - 1)
                    .Col = cnt
                    .Row = 0
                    chk_app2(.Cols - 1).Left = .Left + .CellLeft
                    chk_app2(.Cols - 1).Top = .Top + .CellTop
                    chk_app2(.Cols - 1).Height = .CellHeight
                    chk_app2(.Cols - 1).Width = .CellWidth
                    chk_app2(.Cols - 1).ZOrder 0
                    chk_app2(.Cols - 1).Visible = True
                    chk_app2(.Cols - 1).Caption = ven(0)
                    End If
                    
                    
                    Dim i As Integer
                  i = 0
                  For i = 1 To flex_grid2.Rows - 1
                   
                  
                  Dim vn As New ADODB.Recordset
                  If vn.State Then vn.Close
                  vn.Open "select qt.termsdesc from quotation q, quotationterms qt ,vendor v where q.qno=qt.qno and q.vendor=v.name and v.code ='" & ven(0) & "' and qt.terms ='" & flex_grid2.TextMatrix(i, 0) & "' and q.rfqno='" & cbo_rfqno.Text & "' ", Cn, 3, 2
                  If Not vn.EOF Then
                  flex_grid2.TextMatrix(i, cnt) = vn(0)
                  
                  End If
                  Next i
                                  
                  ven.MoveNext
                  Wend
                  ven.Close
                  
End With

End Sub
