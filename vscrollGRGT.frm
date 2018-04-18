VERSION 5.00
Begin VB.Form vscrollGRGT 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   2955
      Left            =   11400
      TabIndex        =   71
      Top             =   120
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H8000000E&
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11610
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   8
         Left            =   10440
         TabIndex        =   70
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   7
         Left            =   10440
         TabIndex        =   69
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   6
         Left            =   10440
         TabIndex        =   68
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   5
         Left            =   10440
         TabIndex        =   67
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   4
         Left            =   10440
         TabIndex        =   66
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   3
         Left            =   10440
         TabIndex        =   65
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   2
         Left            =   10440
         TabIndex        =   64
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   1
         Left            =   10440
         TabIndex        =   63
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt_qty1 
         Height          =   315
         Index           =   0
         Left            =   10440
         TabIndex        =   62
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   8
         Left            =   9480
         TabIndex        =   61
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   7
         Left            =   9480
         TabIndex        =   60
         Top             =   3000
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   6
         Left            =   9480
         TabIndex        =   59
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   5
         Left            =   9480
         TabIndex        =   58
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   4
         Left            =   9480
         TabIndex        =   57
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   3
         Left            =   9480
         TabIndex        =   56
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   2
         Left            =   9480
         TabIndex        =   55
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   1
         Left            =   9480
         TabIndex        =   54
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom1 
         Height          =   315
         Index           =   0
         Left            =   9480
         TabIndex        =   53
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   8
         Left            =   8040
         TabIndex        =   52
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   7
         Left            =   8040
         TabIndex        =   51
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   6
         Left            =   8040
         TabIndex        =   50
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   5
         Left            =   8040
         TabIndex        =   49
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   4
         Left            =   8040
         TabIndex        =   48
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   3
         Left            =   8040
         TabIndex        =   47
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   2
         Left            =   8040
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   1
         Left            =   8040
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno1 
         Height          =   315
         Index           =   0
         Left            =   8040
         TabIndex        =   44
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   0
         Left            =   6120
         TabIndex        =   36
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   32
         Top             =   1920
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   30
         Top             =   2640
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   4575
      End
      Begin VB.ComboBox cbo_category 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   4575
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   0
         Left            =   4680
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   1
         Left            =   4680
         TabIndex        =   25
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   2
         Left            =   4680
         TabIndex        =   24
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   3
         Left            =   4680
         TabIndex        =   23
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   4
         Left            =   4680
         TabIndex        =   22
         Top             =   1920
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   5
         Left            =   4680
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   6
         Left            =   4680
         TabIndex        =   20
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   7
         Left            =   4680
         TabIndex        =   19
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ComboBox cbo_batchno 
         Height          =   315
         Index           =   8
         Left            =   4680
         TabIndex        =   18
         Top             =   3360
         Width           =   1455
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   1
         Left            =   6120
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   2
         Left            =   6120
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   3
         Left            =   6120
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   4
         Left            =   6120
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   5
         Left            =   6120
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   6
         Left            =   6120
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   7
         Left            =   6120
         TabIndex        =   11
         Top             =   3000
         Width           =   975
      End
      Begin VB.ComboBox cbo_uom 
         Height          =   315
         Index           =   8
         Left            =   6120
         TabIndex        =   10
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   0
         Left            =   7080
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   1
         Left            =   7080
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   2
         Left            =   7080
         TabIndex        =   7
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   3
         Left            =   7080
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   4
         Left            =   7080
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   5
         Left            =   7080
         TabIndex        =   4
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   6
         Left            =   7080
         TabIndex        =   3
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   7
         Left            =   7080
         TabIndex        =   2
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox txt_qty 
         Height          =   315
         Index           =   8
         Left            =   7080
         TabIndex        =   1
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "R UOM"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9480
         TabIndex        =   43
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "R Batch No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   42
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "R Qty"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   10440
         TabIndex        =   41
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "S UOM"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6120
         TabIndex        =   40
         Top             =   240
         Width           =   645
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "S Batch No"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4680
         TabIndex        =   38
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "S Qty"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7080
         TabIndex        =   37
         Top             =   240
         Width           =   390
      End
   End
End
Attribute VB_Name = "vscrollGRGT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyItemboxarray As Object 'A class level dynamic Array
Private MyBatchboxarray As Object
Private MyUOMboxarray As Object
Private MyQtyboxarray As Object

Private MyBatchboxarray1 As Object
Private MyUOMboxarray1 As Object
Private MyQtyboxarray1 As Object

Private Sub cbo_category_Change(Index As Integer)
On Error Resume Next
cbo_batchno(Index).Clear
cbo_uom(Index).Clear
sc = Split(cbo_category(Index).Text, "  -  ", Len(cbo_category(Index).Text), vbTextCompare)
Dim umc As New ADODB.Recordset
If umc.State Then umc.Close
If sc(0) = "" Then Exit Sub
If sc(0) = Null Then Exit Sub

umc.Open "select DISTINCT(batchno),uom from materialbatch where itemid='" & sc(0) & "'", Cn, 3, 2

While Not umc.EOF
cbo_batchno(Index).AddItem umc(0)
cbo_uom(Index).AddItem umc(1)
cbo_batchno1(Index).AddItem umc(0)
cbo_uom1(Index).AddItem umc(1)
umc.MoveNext
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
      
       With addBatchbox1
          .Top = cbo_batchno1(MyBatchboxarray1.ubound - 1).Top + cbo_batchno1(MyBatchboxarray1.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
           With addUOMbox1
          .Top = cbo_uom1(MyUOMboxarray1.ubound - 1).Top + cbo_uom1(MyUOMboxarray1.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox1
          .Top = txt_qty1(MyQtyboxarray1.ubound - 1).Top + txt_qty1(MyQtyboxarray1.ubound - 1).Height + 100
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
      
             With addBatchbox1
          .Top = cbo_batchno1(MyBatchboxarray1.ubound - 1).Top + cbo_batchno1(MyBatchboxarray1.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
           With addUOMbox1
          .Top = cbo_uom1(MyUOMboxarray1.ubound - 1).Top + cbo_uom1(MyUOMboxarray1.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
      With addQtybox1
          .Top = txt_qty1(MyQtyboxarray1.ubound - 1).Top + txt_qty1(MyQtyboxarray1.ubound - 1).Height + 100
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
    
    Set MyBatchboxarray1 = Me.Controls("cbo_batchno1")
    Set MyUOMboxarray1 = Me.Controls("cbo_uom1")
    Set MyQtyboxarray1 = Me.Controls("txt_qty1")
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

Public Function addBatchbox1() As ComboBox
   Dim m As Integer
   m = MyBatchboxarray1.ubound + 1
   Load MyBatchboxarray1(m)
   Set addBatchbox1 = MyBatchboxarray1(m)
End Function
Public Function addUOMbox1() As ComboBox
   Dim m As Integer
   m = MyUOMboxarray1.ubound + 1
   Load MyUOMboxarray1(m)
   Set addUOMbox1 = MyUOMboxarray1(m)
End Function
Public Function addQtybox1() As TextBox
   Dim m As Integer
   m = MyQtyboxarray1.ubound + 1
   Load MyQtyboxarray1(m)
   Set addQtybox1 = MyQtyboxarray1(m)
End Function
