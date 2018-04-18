VERSION 5.00
Begin VB.Form vscrollgoodstransfer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2955
      Left            =   9480
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
      Width           =   9810
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   8
         Left            =   8280
         TabIndex        =   38
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   7
         Left            =   8280
         TabIndex        =   37
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   6
         Left            =   8280
         TabIndex        =   36
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   5
         Left            =   8280
         TabIndex        =   35
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   4
         Left            =   8280
         TabIndex        =   34
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   3
         Left            =   8280
         TabIndex        =   33
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   2
         Left            =   8280
         TabIndex        =   32
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   1
         Left            =   8280
         TabIndex        =   31
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   0
         Left            =   8280
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   8
         Left            =   7200
         TabIndex        =   29
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   7
         Left            =   7200
         TabIndex        =   28
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   6
         Left            =   7200
         TabIndex        =   27
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   5
         Left            =   7200
         TabIndex        =   26
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   4
         Left            =   7200
         TabIndex        =   25
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   3
         Left            =   7200
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   2
         Left            =   7200
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   1
         Left            =   7200
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   8
         Left            =   5400
         TabIndex        =   21
         Top             =   3360
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   7
         Left            =   5400
         TabIndex        =   20
         Top             =   3000
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   6
         Left            =   5400
         TabIndex        =   19
         Top             =   2640
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   5
         Left            =   5400
         TabIndex        =   18
         Top             =   2280
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   4
         Left            =   5400
         TabIndex        =   17
         Top             =   1920
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   3
         Left            =   5400
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   2
         Left            =   5400
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   1
         Left            =   5400
         TabIndex        =   14
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   0
         Left            =   5400
         TabIndex        =   13
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   5295
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   5295
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   0
         Left            =   7200
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Qty"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         TabIndex        =   41
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Batch No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   5400
         TabIndex        =   40
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Item"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "UOM"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7200
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
   End
End
Attribute VB_Name = "vscrollgoodstransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyItemboxarray As Object 'A class level dynamic Array
Private MyBatchboxarray As Object
Private MyUOMboxarray As Object
Private MyQtyboxarray As Object


Private Sub cbo_category_Click(Index As Integer)
On Error Resume Next
cbo_batchno(Index).Text = ""
cbo_uom(Index).Text = ""
sc = Split(cbo_category(Index).Text, "  -  ", Len(cbo_category(Index).Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close

If goodstransfer.cbo_lookup.Text = "Item ID" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(0) & "'", Cn, 3, 2
ElseIf goodstransfer.cbo_lookup.Text = "Mfr PartNo." Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(1) & "'", Cn, 3, 2
ElseIf goodstransfer.cbo_lookup.Text = "Item Description" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(2) & "'", Cn, 3, 2
ElseIf goodstransfer.cbo_lookup.Text = "Search" Then
um.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(2) & "'", Cn, 3, 2
End If
While Not um.EOF
cbo_batchno(Index).AddItem um(0)
cbo_uom(Index).AddItem um(1)
um.MoveNext
Wend
End Sub

Private Sub Form_Load()


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

Private Sub VScroll1_Change()
  
  ScrollForm 0, VScroll1.Value
'''

    With addItembox
          .Top = cbo_category(MyItemboxarray.ubound - 1).Top + cbo_category(MyItemboxarray.ubound - 1).Height + 100
                    .Visible = True
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
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox
          .Top = txt_qty(MyQtyboxarray.ubound - 1).Top + txt_qty(MyQtyboxarray.ubound - 1).Height + 100
                    .Visible = True
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
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox
          .Top = txt_qty(MyQtyboxarray.ubound - 1).Top + txt_qty(MyQtyboxarray.ubound - 1).Height + 100
                    .Visible = True
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

