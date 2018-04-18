VERSION 5.00
Begin VB.Form vscrollGRN 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   11865
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2955
      Left            =   11400
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11730
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   8
         Left            =   10560
         TabIndex        =   90
         Top             =   3360
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   7
         Left            =   10560
         TabIndex        =   89
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   6
         Left            =   10560
         TabIndex        =   88
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   5
         Left            =   10560
         TabIndex        =   87
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   4
         Left            =   10560
         TabIndex        =   86
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   3
         Left            =   10560
         TabIndex        =   85
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   2
         Left            =   10560
         TabIndex        =   84
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   1
         Left            =   10560
         TabIndex        =   83
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cbo_rejection 
         Height          =   315
         Index           =   0
         Left            =   10560
         TabIndex        =   82
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   8
         Left            =   9840
         TabIndex        =   79
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   7
         Left            =   9840
         TabIndex        =   78
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   6
         Left            =   9840
         TabIndex        =   77
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   5
         Left            =   9840
         TabIndex        =   76
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   4
         Left            =   9840
         TabIndex        =   75
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   3
         Left            =   9840
         TabIndex        =   74
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   2
         Left            =   9840
         TabIndex        =   73
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   1
         Left            =   9840
         TabIndex        =   72
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrej 
         Height          =   315
         Index           =   0
         Left            =   9840
         TabIndex        =   71
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   8
         Left            =   9120
         TabIndex        =   70
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   7
         Left            =   9120
         TabIndex        =   69
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   6
         Left            =   9120
         TabIndex        =   68
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   5
         Left            =   9120
         TabIndex        =   67
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   4
         Left            =   9120
         TabIndex        =   66
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   3
         Left            =   9120
         TabIndex        =   65
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   2
         Left            =   9120
         TabIndex        =   64
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   1
         Left            =   9120
         TabIndex        =   63
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_qtyrec 
         Height          =   315
         Index           =   0
         Left            =   9120
         TabIndex        =   62
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   61
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   60
         Top             =   3360
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   59
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   58
         Top             =   3000
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   57
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   56
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   55
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   54
         Top             =   2280
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   53
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   52
         Top             =   1920
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   51
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   1560
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   49
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   47
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   46
         Top             =   840
         Width           =   255
      End
      Begin VB.CheckBox chk_qty 
         BackColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   480
         Width           =   255
      End
      Begin VB.CheckBox chk_item 
         BackColor       =   &H8000000E&
         Caption         =   "Check1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   255
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   0
         Left            =   7560
         TabIndex        =   37
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   1
         Left            =   720
         TabIndex        =   36
         Top             =   840
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   35
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   3
         Left            =   720
         TabIndex        =   34
         Top             =   1560
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   4
         Left            =   720
         TabIndex        =   33
         Top             =   1920
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   5
         Left            =   720
         TabIndex        =   32
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   6
         Left            =   720
         TabIndex        =   31
         Top             =   2640
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   7
         Left            =   720
         TabIndex        =   30
         Top             =   3000
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   8
         Left            =   720
         TabIndex        =   29
         Top             =   3360
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   28
         Top             =   480
         Width           =   5295
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   0
         Left            =   6000
         TabIndex        =   27
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   1
         Left            =   6000
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   2
         Left            =   6000
         TabIndex        =   25
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   24
         Top             =   1560
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   4
         Left            =   6000
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   5
         Left            =   6000
         TabIndex        =   22
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   6
         Left            =   6000
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   7
         Left            =   6000
         TabIndex        =   20
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   8
         Left            =   6000
         TabIndex        =   19
         Top             =   3360
         Width           =   1575
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   1
         Left            =   7560
         TabIndex        =   18
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   2
         Left            =   7560
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   3
         Left            =   7560
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   4
         Left            =   7560
         TabIndex        =   15
         Top             =   1920
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   5
         Left            =   7560
         TabIndex        =   14
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   6
         Left            =   7560
         TabIndex        =   13
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   7
         Left            =   7560
         TabIndex        =   12
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   8
         Left            =   7560
         TabIndex        =   11
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   0
         Left            =   8400
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   1
         Left            =   8400
         TabIndex        =   9
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   2
         Left            =   8400
         TabIndex        =   8
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   3
         Left            =   8400
         TabIndex        =   7
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   4
         Left            =   8400
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   5
         Left            =   8400
         TabIndex        =   5
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   6
         Left            =   8400
         TabIndex        =   4
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   7
         Left            =   8400
         TabIndex        =   3
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   8
         Left            =   8400
         TabIndex        =   2
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rjtd Desc"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   10560
         TabIndex        =   91
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty Rjtd"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   9840
         TabIndex        =   81
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty Rcvd"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9120
         TabIndex        =   80
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Qty"
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Itm"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "UOM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7560
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   720
         TabIndex        =   40
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6000
         TabIndex        =   39
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PO Qty"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8400
         TabIndex        =   38
         Top             =   240
         Width           =   510
      End
   End
End
Attribute VB_Name = "vscrollGRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyItemboxarray As Object 'A class level dynamic Array
Private MyBatchboxarray As Object
Private MyUOMboxarray As Object
Private MyQtyboxarray As Object
Private Mychkitemarray As Object
Private Mychkqtyarray As Object
Private MyQtyboxarrayrc As Object
Private MyQtyboxarrayrj As Object
Private Myreasonarray As Object

Private Sub cbo_category_Change(Index As Integer)
 On Error Resume Next
cbo_batchno(Index).Clear
cbo_uom(Index).Clear
sc = Split(cbo_category(Index).Text, "  -  ", Len(cbo_category(Index).Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(0) & "'", Cn, 3, 2
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend

End Sub

Private Sub cbo_category_Click(Index As Integer)
cbo_batchno(Index).Clear
cbo_uom(Index).Clear
sc = Split(cbo_category(Index).Text, "  -  ", Len(cbo_category(Index).Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close

If GRN.cbo_lookup.Text = "Item ID" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(0) & "'", Cn, 3, 2
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend
ElseIf GRN.cbo_lookup.Text = "Mfr PartNo." Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(1) & "'", Cn, 3, 2
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend
ElseIf GRN.cbo_lookup.Text = "Item Description" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(2) & "'", Cn, 3, 2
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend
ElseIf GRN.cbo_lookup.Text = "Search" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(2) & "'", Cn, 3, 2
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend
End If
End Sub

Private Sub cbo_rejection_Click(Index As Integer)
cbo_rejection(Index).ToolTipText = cbo_rejection(Index).Text
End Sub

Private Sub chk_item_Click(Index As Integer)
If chk_item(Index).Value = 1 Then
cbo_category(Index).Enabled = True
Else
cbo_category(Index).Enabled = False
End If
End Sub

Private Sub chk_qty_Click(Index As Integer)
If chk_qty(Index).Value = 1 Then
cbo_uom(Index).Enabled = True
txt_qty(Index).Enabled = False
txt_qtyrec(Index).Enabled = True
txt_qtyrej(Index).Enabled = False

cbo_rejection(Index).AddItem "Scrap"
cbo_rejection(Index).AddItem "Quality"
cbo_rejection(Index).AddItem "Damaged"
cbo_rejection(Index).Enabled = True
Else
cbo_uom(Index).Enabled = False
txt_qty(Index).Enabled = False
txt_qtyrec(Index).Enabled = False
txt_qtyrej(Index).Enabled = False
txt_qtyrec(Index).Text = txt_qty(Index).Text
txt_qtyrej(Index).Text = 0
cbo_rejection(Index).Enabled = False
cbo_rejection(Index).Text = ""
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

Dim VPos As Integer
 
  'Change the following numbers to the Full height and width of your Form
  intFullHeight = 12000 'Maximized the Form and note the Figures
   
  'This is the how much of your Form is displayed
  intDisplayHeight = Me.Height
   

  With VScroll1
    '.Height = Me.ScaleHeight
    .Min = 0
    .Max = intFullHeight - intDisplayHeight
    .SmallChange = Screen.TwipsPerPixelX * 10
    .LargeChange = .SmallChange
  End With
    
    
    
    
    

    
'scroll
End Sub
Sub ScrollForm(Direction As Byte, NewVal As Integer)
  
  Dim CTL As Control
  Static hOldVal As Integer
  Static vOldVal As Integer
  Dim hMoveDiff As Integer 'Diff in the horizontal controls movements
  Dim vMoveDiff As Integer 'Diff in the vertical controls Movements
  
  Select Case Direction
    
  Case 0 'Scroll Vertically
  
    'Check The Direction of the Vertical Scroll & Extract Value Diff
    If NewVal > vOldVal Then 'Scrolled From Top to Bottom
      'Controls MUST move to the TOP, therefore TOP value Decreases
      vMoveDiff = -(NewVal - vOldVal)
      
            '''''''''''''''''
'        pView.Height = pView.Height - 400
        Frame1.Height = Frame1.Height + 400
'        vscrollform.Height = vscrollform.Height + 400
''''''''''''''''
    Else 'Scrolled From Bottom to Top
      'Controls MUST move to the Bottom, therefore TOP value Increases
      vMoveDiff = (vOldVal - NewVal)
      
      '''''''''''''''''
'        pView.Height = pView.Height - 400
        Frame1.Height = Frame1.Height - 400
'        Me.Height = Me.Height - 400
''''''''''''''''
      
      
      
    End If
  
    For Each CTL In Me.Controls
      'Make sure it's not a ScrollBar
      If Not (TypeOf CTL Is VScrollBar) Then
        'If it's a Line then
        If TypeOf CTL Is Line Then
          CTL.Y1 = CTL.Y1 + vMoveDiff '+ VPos - VScroll1.Value
          CTL.Y2 = CTL.Y2 + vMoveDiff '+ VPos - VScroll1.Value
        Else
          CTL.Top = CTL.Top + vMoveDiff '+ VPos - VScroll1.Value
        End If
      End If
    Next
    
      vOldVal = NewVal 'Reset vOldVal to reflect New Pos of ScrollBar
    
     
  End Select

End Sub

Private Sub txt_qtyrec_Change(Index As Integer)
On Error Resume Next
txt_qtyrej(Index).Text = txt_qty(Index).Text - txt_qtyrec(Index)
End Sub

Private Sub txt_qtyrec_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
txt_qtyrej(Index).Text = txt_qty(Index).Text - txt_qtyrec(Index)
End Sub

Private Sub VScroll1_Change()
  
  ScrollForm 0, VScroll1.Value
'''

    With addItembox
          .Top = cbo_category(MyItemboxarray.ubound - 1).Top + cbo_category(MyItemboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          '          .Enabled = False
          .SetFocus
      End With


     With addBatchbox
          .Top = cbo_batchno(MyBatchboxarray.ubound - 1).Top + cbo_batchno(MyBatchboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
           With addUOMbox
          .Top = cbo_uom(MyUOMboxarray.ubound - 1).Top + cbo_uom(MyUOMboxarray.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox
          .Top = txt_qty(MyQtyboxarray.ubound - 1).Top + txt_qty(MyQtyboxarray.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
       With addchkitem
          .Top = chk_item(Mychkitemarray.ubound - 1).Top + chk_item(Mychkitemarray.ubound - 1).Height + 100
                    .Visible = True
                  ' .Enabled = False
                    .Value = False
          .SetFocus
      End With
      With addchkqty
          .Top = chk_qty(Mychkqtyarray.ubound - 1).Top + chk_qty(Mychkqtyarray.ubound - 1).Height + 100
                    .Visible = True
                    '.Enabled = False
                    .Value = False
          .SetFocus
      End With
      
      
      
      With addQtyboxrc
          .Top = txt_qtyrec(MyQtyboxarrayrc.ubound - 1).Top + txt_qtyrec(MyQtyboxarrayrc.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      With addQtyboxrj
          .Top = txt_qtyrej(MyQtyboxarrayrj.ubound - 1).Top + txt_qtyrej(MyQtyboxarrayrj.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      
'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''

End Sub

Private Sub VScroll1_Scroll()
  
 ScrollForm 0, VScroll1.Value
'''

    With addItembox
          .Top = cbo_category(MyItemboxarray.ubound - 1).Top + cbo_category(MyItemboxarray.ubound - 1).Height + 100
                    .Visible = True
                 '   .Enabled = False
                    .Text = ""
          .SetFocus
      End With


     With addBatchbox
          .Top = cbo_batchno(MyBatchboxarray.ubound - 1).Top + cbo_batchno(MyBatchboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
           With addUOMbox
          .Top = cbo_uom(MyUOMboxarray.ubound - 1).Top + cbo_uom(MyUOMboxarray.ubound - 1).Height + 100
                    .Visible = True
                 '   .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox
          .Top = txt_qty(MyQtyboxarray.ubound - 1).Top + txt_qty(MyQtyboxarray.ubound - 1).Height + 100
                    .Visible = True
                  '  .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      With addchkitem
          .Top = chk_item(Mychkitemarray.ubound - 1).Top + chk_item(Mychkitemarray.ubound - 1).Height + 100
                    .Visible = True
                    '.Enabled = False
                    .Value = False
          .SetFocus
      End With
      With addchkqty
          .Top = chk_qty(Mychkqtyarray.ubound - 1).Top + chk_qty(Mychkqtyarray.ubound - 1).Height + 100
                    .Visible = True
                    '.Enabled = False
                    .Value = False
          .SetFocus
      End With
      
      With addQtyboxrc
          .Top = txt_qtyrec(MyQtyboxarrayrc.ubound - 1).Top + txt_qtyrec(MyQtyboxarrayrc.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      With addQtyboxrj
          .Top = txt_qtyrej(MyQtyboxarrayrj.ubound - 1).Top + txt_qtyrej(MyQtyboxarrayrj.ubound - 1).Height + 100
                    .Visible = True
                   ' .Enabled = False
                    .Text = ""
          .SetFocus
      End With
      With addreason
          .Top = cbo_rejection(Myreasonarray.ubound - 1).Top + cbo_rejection(Myreasonarray.ubound - 1).Height + 100
                    .Visible = True
                 '   .Enabled = False
                    .Text = ""
          .SetFocus
      End With
'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''

End Sub

Private Sub Form_Initialize()
    Set MyItemboxarray = Me.Controls("cbo_category")
    Set MyBatchboxarray = Me.Controls("cbo_batchno")
    Set MyUOMboxarray = Me.Controls("cbo_uom")
    Set MyQtyboxarray = Me.Controls("txt_qty")
    Set Mychkitemarray = Me.Controls("chk_item")
    Set Mychkqtyarray = Me.Controls("chk_qty")
    
    Set MyQtyboxarrayrc = Me.Controls("txt_qtyrec")
    Set MyQtyboxarrayrj = Me.Controls("txt_qtyrej")
    Set Myreasonarray = Me.Controls("cbo_rejection")
    End Sub

Public Function addItembox() As ComboBox
   Dim m As Integer
   m = MyItemboxarray.ubound + 1
   Load MyItemboxarray(m)
   Set addItembox = MyItemboxarray(m)
End Function
Public Function addBatchbox() As ComboBox
   Dim m As Integer
   m = MyBatchboxarray.ubound + 1
   Load MyBatchboxarray(m)
   Set addBatchbox = MyBatchboxarray(m)
End Function
Public Function addUOMbox() As ComboBox
   Dim m As Integer
   m = MyUOMboxarray.ubound + 1
   Load MyUOMboxarray(m)
   Set addUOMbox = MyUOMboxarray(m)
End Function
Public Function addQtybox() As TextBox
   Dim m As Integer
   m = MyQtyboxarray.ubound + 1
   Load MyQtyboxarray(m)
   Set addQtybox = MyQtyboxarray(m)
End Function
Public Function addchkitem() As CheckBox
   Dim m As Integer
   m = Mychkitemarray.ubound + 1
   Load Mychkitemarray(m)
   Set addchkitem = Mychkitemarray(m)
End Function
Public Function addchkqty() As CheckBox
   Dim m As Integer
   m = Mychkqtyarray.ubound + 1
   Load Mychkqtyarray(m)
   Set addchkqty = Mychkqtyarray(m)
End Function

Public Function addQtyboxrc() As TextBox
   Dim m As Integer
   m = MyQtyboxarrayrc.ubound + 1
   Load MyQtyboxarrayrc(m)
   Set addQtyboxrc = MyQtyboxarrayrc(m)
End Function
Public Function addQtyboxrj() As TextBox
   Dim m As Integer
   m = MyQtyboxarrayrj.ubound + 1
   Load MyQtyboxarrayrj(m)
   Set addQtyboxrj = MyQtyboxarrayrj(m)
End Function
Public Function addreason() As ComboBox
   Dim m As Integer
   m = Myreasonarray.ubound + 1
   Load Myreasonarray(m)
   Set addreason = Myreasonarray(m)
End Function
