VERSION 5.00
Begin VB.Form vscrollform 
   BackColor       =   &H00A04729&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   11085
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   3675
      Left            =   10800
      TabIndex        =   21
      Top             =   240
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Caption         =   "Material SubCategory"
      ForeColor       =   &H8000000E&
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   6840
         TabIndex        =   31
         Top             =   3360
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   6840
         TabIndex        =   30
         Top             =   3000
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   6840
         TabIndex        =   29
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   6840
         TabIndex        =   28
         Top             =   2280
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   6840
         TabIndex        =   27
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   6840
         TabIndex        =   26
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   6840
         TabIndex        =   25
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   6840
         TabIndex        =   24
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   22
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   18
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   840
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1440
         TabIndex        =   10
         Top             =   1920
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   8
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   6
         Top             =   2640
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1440
         TabIndex        =   4
         Top             =   3000
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   2
         Top             =   3360
         Width           =   5295
      End
      Begin VB.TextBox text2 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   6840
         TabIndex        =   23
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
         Left            =   1440
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "vscrollform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyTextboxArray As Object 'A class level dynamic Array
Private MyDescboxarray As Object
Private MyRemarkboxarray As Object
Private Sub Form_Load()
''''scroll
If ct1 = 1 Then
   Frame1.Caption = "Material Category Level-2"
   ElseIf ct1 = 2 Then
   Frame1.Caption = "Material Category Level-3"
   ElseIf ct1 = 3 Then
   Frame1.Caption = "Material Category Level-4"
   ElseIf ct1 = 4 Then
   Frame1.Caption = "Material Master"
   Else
End If

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
'pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
'Me.Height = Frame1.Height + 400
''''''


End Sub

Private Sub Form_Initialize()
    Set MyTextboxArray = Me.Controls("Text1")
    Set MyDescboxarray = Me.Controls("Text2")
    Set MyRemarkboxarray = Me.Controls("Text3")
    End Sub

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
