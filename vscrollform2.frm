VERSION 5.00
Begin VB.Form vscrollform2 
   BackColor       =   &H00A04729&
   BorderStyle     =   0  'None
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16500
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   16500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   3795
      Left            =   11280
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4000
      Width           =   11475
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   16400
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   13080
         TabIndex        =   65
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   13080
         TabIndex        =   64
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   13080
         TabIndex        =   63
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   13080
         TabIndex        =   62
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   13080
         TabIndex        =   61
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   13080
         TabIndex        =   60
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   13080
         TabIndex        =   59
         Top             =   1200
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   13080
         TabIndex        =   58
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   13080
         TabIndex        =   57
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   56
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   54
         Top             =   840
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   52
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   50
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1200
         TabIndex        =   48
         Top             =   1920
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1200
         TabIndex        =   46
         Top             =   2280
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1200
         TabIndex        =   44
         Top             =   2640
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1200
         TabIndex        =   42
         Top             =   3000
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   41
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1200
         TabIndex        =   40
         Top             =   3360
         Width           =   4935
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   9840
         TabIndex        =   38
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   9840
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   9840
         TabIndex        =   36
         Top             =   1200
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   3
         Left            =   9840
         TabIndex        =   35
         Top             =   1560
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   4
         Left            =   9840
         TabIndex        =   34
         Top             =   1920
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   5
         Left            =   9840
         TabIndex        =   33
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   6
         Left            =   9840
         TabIndex        =   32
         Top             =   2640
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   7
         Left            =   9840
         TabIndex        =   31
         Top             =   3000
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   8
         Left            =   9840
         TabIndex        =   30
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   11040
         TabIndex        =   29
         Top             =   480
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   11040
         TabIndex        =   28
         Top             =   840
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   2
         Left            =   11040
         TabIndex        =   27
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   3
         Left            =   11040
         TabIndex        =   26
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   4
         Left            =   11040
         TabIndex        =   25
         Top             =   1920
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   5
         Left            =   11040
         TabIndex        =   24
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   6
         Left            =   11040
         TabIndex        =   23
         Top             =   2640
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   7
         Left            =   11040
         TabIndex        =   22
         Top             =   3000
         Width           =   1935
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   8
         Left            =   11040
         TabIndex        =   21
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   6240
         TabIndex        =   20
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   8160
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   8160
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   6240
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   8160
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   6240
         TabIndex        =   14
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   8160
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   6240
         TabIndex        =   12
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   8160
         TabIndex        =   11
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6240
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   8160
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6240
         TabIndex        =   8
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   8160
         TabIndex        =   7
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   6240
         TabIndex        =   6
         Top             =   3000
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   8160
         TabIndex        =   5
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   6240
         TabIndex        =   4
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   8160
         TabIndex        =   3
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   13080
         TabIndex        =   72
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   1200
         TabIndex        =   71
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00A04729&
         Caption         =   "UOM"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   9840
         TabIndex        =   69
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00A04729&
         Caption         =   "Material Type"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   11040
         TabIndex        =   68
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00A04729&
         Caption         =   "Mfr.Ref Code"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   8160
         TabIndex        =   67
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00A04729&
         Caption         =   "ItemId"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6240
         TabIndex        =   66
         Top             =   240
         Width           =   915
      End
   End
End
Attribute VB_Name = "vscrollform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyTextboxArray As Object 'A class level dynamic Array
Private MyDescboxarray As Object
Private MyRemarkboxarray As Object
Private MyComboboxarray As Object
Private MyTypeboxarray As Object
Private MyItemboxarray As Object
Private MyMfrboxarray As Object
Public conid As String
Public getuom As String
Public gettype As String

Private Sub Form_Load()
''''scroll
'''If ct1 = 1 Then
'''   Frame1.Caption = "Material Category Level-2"
'''   ElseIf ct1 = 2 Then
'''   Frame1.Caption = "Material Category Level-3"
'''   ElseIf ct1 = 3 Then
'''   Frame1.Caption = "Material Category Level-4"
'''   ElseIf ct1 = 4 Then
'''   Frame1.Caption = "Material Master"
'''   Else
'''End If

Dim VPos As Integer
Dim Hpos As Integer
  'Change the following numbers to the Full height and width of your Form
  intFullHeight = 12000 'Maximized the Form and note the Figures
  intFullWidth = 12000
  'This is the how much of your Form is displayed
  intDisplayHeight = Me.Height
  intDisplayWidth = Me.Width

  With VScroll1
    '.Height = Me.ScaleHeight
    .Min = 0
    .Max = intFullHeight - intDisplayHeight
    .SmallChange = Screen.TwipsPerPixelX * 10
    .LargeChange = .SmallChange
  End With
    
    
    
  With HScroll1
    '.Width = Me.ScaleWidth
    .Min = 0
    .Max = intFullWidth - intDisplayWidth
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
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) Then
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
      
      
      
      ''horizatal scrolll
      
      
      Case 1 'Scroll Horizontally
  
    'Check The Direction of the Horizontal Scroll & Extract Value Diff
    If NewVal > hOldVal Then 'Scrolled From Left to Right
      'Controls MUST move to the LEFT, therefore LEFT value Decreases
      hMoveDiff = -(NewVal - hOldVal)
'      Frame1.Width = Frame1.Width + 400
    Else 'Scrolled From Right to Left
      'Controls MUST move to the RIGHT, therefore LEFT value Increases
      hMoveDiff = (hOldVal - NewVal)
'       Frame1.Width = Frame1.Width - 400
    End If
  
    For Each CTL In Me.Controls
      'Make sure it's not a ScrollBar
      If Not (TypeOf CTL Is VScrollBar) And Not _
             (TypeOf CTL Is HScrollBar) Then
        'If it's a Line then
        If TypeOf CTL Is Line Then
          CTL.X1 = CTL.X1 + hMoveDiff
          CTL.X2 = CTL.X2 + hMoveDiff
        Else
          CTL.Left = CTL.Left - hMoveDiff
        End If
      End If
    Next
      
      hOldVal = NewVal 'Reset hOldVal to reflect New Pos of ScrollBar
    
    
     
  End Select

End Sub

Private Sub Text1_GotFocus(Index As Integer)
If Text1(Index).Text <> "" Then
Call gen_itemid
Text4(Index).Text = conid & "-" & text2(Index).Text
Combo1(Index).Text = getuom
Combo2(Index).Text = gettype
Call uom_type
End If
End Sub

Private Sub text2_LostFocus(Index As Integer)
If text2(Index).Text <> "" Then
Call gen_itemid
Text4(Index).Text = conid & "-" & text2(Index).Text
Combo1(Index).Text = getuom
Combo2(Index).Text = gettype
Call uom_type
End If
End Sub

Private Sub VScroll1_Change()
  
  ScrollForm 0, VScroll1.Value
'''

    With addTextbox
          .Top = Text1(MyTextboxArray.ubound - 1).Top + Text1(MyTextboxArray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With


     With adddescbox
          .Top = text2(MyDescboxarray.ubound - 1).Top + text2(MyDescboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
           With addremarkbox
          .Top = Text3(MyRemarkboxarray.ubound - 1).Top + text2(MyRemarkboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
    With addcombobox
    .Top = Combo1(MyComboboxarray.ubound - 1).Top + Combo1(MyComboboxarray.ubound - 1).Height + 100
    .Visible = True
    .Text = ""
    .SetFocus
      End With
      
      
      With addtypebox
    .Top = Combo2(MyTypeboxarray.ubound - 1).Top + Combo2(MyTypeboxarray.ubound - 1).Height + 100
    .Visible = True
    .Text = ""
    .SetFocus
      End With
      
      
      
      With addItembox
        .Top = Text4(MyItemboxarray.ubound - 1).Top + Text4(MyItemboxarray.ubound - 1).Height + 100
                .Visible = True
                .Text = ""
        .SetFocus
        End With
        
        With addMfrbox
        .Top = Text5(MyMfrboxarray.ubound - 1).Top + Text5(MyMfrboxarray.ubound - 1).Height + 100
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

With addTextbox
    .Top = Text1(MyTextboxArray.ubound - 1).Top + Text1(MyTextboxArray.ubound - 1).Height + 100
              .Visible = True
              .Text = ""
    .SetFocus
End With


     With adddescbox
          .Top = text2(MyDescboxarray.ubound - 1).Top + text2(MyDescboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
     With addremarkbox
          .Top = text2(MyRemarkboxarray.ubound - 1).Top + text2(MyRemarkboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
      With addcombobox
          .Top = Combo1(MyComboboxarray.ubound - 1).Top + Combo1(MyComboboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
      With addtypebox
          .Top = Combo2(MyTypeboxarray.ubound - 1).Top + Combo2(MyTypeboxarray.ubound - 1).Height + 100
                    .Visible = True
                    .Text = ""
          .SetFocus
      End With
      
      
      
        With addItembox
        .Top = Text4(MyItemboxarray.ubound - 1).Top + Text4(MyItemboxarray.ubound - 1).Height + 100
                .Visible = True
                .Text = ""
        .SetFocus
        End With
        
        With addMfrbox
        .Top = Text5(MyMfrboxarray.ubound - 1).Top + Text5(MyMfrboxarray.ubound - 1).Height + 100
                .Visible = True
                .Text = ""
        .SetFocus
        End With

      
      
'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''


End Sub

Private Sub HScroll1_Change()
  
  ScrollForm 1, HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()
  
  ScrollForm 1, HScroll1.Value

End Sub


Private Sub Form_Initialize()
    Set MyTextboxArray = Me.Controls("Text1")
    Set MyDescboxarray = Me.Controls("Text2")
    Set MyRemarkboxarray = Me.Controls("Text3")
    Set MyItemboxarray = Me.Controls("Text4")
    Set MyMfrboxarray = Me.Controls("Text5")
    Set MyComboboxarray = Me.Controls("Combo1")
    Set MyTypeboxarray = Me.Controls("Combo2")
    End Sub
Public Function addItembox() As TextBox
   Dim n As Integer
   n = MyItemboxarray.ubound + 1
   Load MyItemboxarray(n)
   Set addItembox = MyItemboxarray(n)
End Function
Public Function addMfrbox() As TextBox
   Dim n As Integer
   n = MyMfrboxarray.ubound + 1
   Load MyMfrboxarray(n)
   Set addMfrbox = MyMfrboxarray(n)
End Function
Public Function addTextbox() As TextBox
   Dim n As Integer
   n = MyTextboxArray.ubound + 1
   Load MyTextboxArray(n)
   Set addTextbox = MyTextboxArray(n)
End Function

Public Function adddescbox() As TextBox
   Dim m As Integer
   m = MyDescboxarray.ubound + 1
   Load MyDescboxarray(m)
   Set adddescbox = MyDescboxarray(m)
End Function
Public Function addremarkbox() As TextBox
   Dim m As Integer
   m = MyRemarkboxarray.ubound + 1
   Load MyRemarkboxarray(m)
   Set addremarkbox = MyRemarkboxarray(m)
End Function

Public Function addcombobox() As ComboBox
   Dim m As Integer
   m = MyComboboxarray.ubound + 1
   Load MyComboboxarray(m)
   Set addcombobox = MyComboboxarray(m)
End Function
Public Function addtypebox() As ComboBox
   Dim m As Integer
   m = MyTypeboxarray.ubound + 1
   Load MyTypeboxarray(m)
   Set addtypebox = MyTypeboxarray(m)
End Function

Public Sub gen_itemid()
conid = ""
getuom = ""
gettype = ""
Dim itm As New ADODB.Recordset
If itm.State Then itm.Close
itm.Open "select ml1.ml1code,ml2.ml2code,ml3.ml3code ,ml3.ml3uom ,ml3.ml3type from ml3,ml2,ml1 where ml1.ml1code=ml2.ml1code and ml2.ml2code=ml3.ml2code and ml3code='" & materiallevel3.txt_categorycode.Text & "' ", Cn, 3, 2
If Not itm.EOF Then
conid = itm(0) & "-" & itm(1) & "-" & itm(2) & ""
getuom = itm(3)
gettype = itm(4)
End If


End Sub
Public Sub uom_type()
  
  Dim um As New ADODB.Recordset
  If um.State Then um.Close
  um.Open "select DISTINCT(mjuom) from uom ", Cn, 3, 2
  While Not um.EOF
  Combo1(Index).AddItem um(0)
  um.MoveNext
  Wend
  
  Dim tp As New ADODB.Recordset
  If tp.State Then tp.Close
  tp.Open "select DISTINCT(mtcode) from materialtype ", Cn, 3, 2
  While Not tp.EOF
  Combo2(Index).AddItem tp(0)
  tp.MoveNext
  Wend
End Sub
