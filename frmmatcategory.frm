VERSION 5.00
Begin VB.Form frmmatcategory 
   Caption         =   "Example"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "frmmatcategory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   4200
      Top             =   3000
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Vote for this code >>"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2640
      Width           =   1695
   End
   Begin VB.PictureBox ScrllngFrm1 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   8115
      TabIndex        =   37
      Top             =   120
      Width           =   8175
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   8295
         TabIndex        =   38
         Top             =   -480
         Visible         =   0   'False
         Width           =   8295
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1680
            TabIndex        =   71
            Top             =   960
            Width           =   6135
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   0
            Left            =   1680
            TabIndex        =   73
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   840
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   840
            TabIndex        =   29
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   840
            TabIndex        =   32
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   840
            TabIndex        =   31
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Company Information:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Website:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   61
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone:"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   60
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Address: "
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Business:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton Command7 
            Caption         =   "Poor"
            Height          =   375
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Good"
            Height          =   375
            Index           =   2
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Very good"
            Height          =   375
            Index           =   3
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Excellent"
            Height          =   375
            Index           =   1
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Select an option:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Poor:"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   67
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Good:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Very good: "
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   65
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Excellent:"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   3735
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.Frame Frame1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   3375
            Begin VB.TextBox Text10 
               Height          =   1125
               Left            =   240
               MultiLine       =   -1  'True
               TabIndex        =   44
               Top             =   600
               Width           =   2895
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Comments:"
               Height          =   255
               Left            =   240
               TabIndex        =   45
               Top             =   360
               Width           =   1335
            End
         End
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   0
         ScaleHeight     =   6615
         ScaleWidth      =   3735
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.CommandButton Command10 
            Caption         =   "Command3"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option2"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   19
            Top             =   5640
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option3"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   5880
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option4"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   21
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option5"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   22
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option6"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   23
            Top             =   5640
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option7"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   24
            Top             =   5880
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option8"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   25
            Top             =   5160
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option9"
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   26
            Top             =   5400
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option10"
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   27
            Top             =   5640
            Width           =   1035
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option11"
            Height          =   255
            Index           =   11
            Left            =   2400
            TabIndex        =   28
            Top             =   5880
            Width           =   1035
         End
         Begin VB.Frame Frame2 
            Caption         =   "Frame1"
            Height          =   975
            Left            =   120
            TabIndex        =   48
            Top             =   4080
            Width           =   3015
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Text            =   "Text10"
               Top             =   600
               Width           =   735
            End
         End
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1440
            TabIndex        =   15
            Text            =   "Text9"
            Top             =   3600
            Width           =   1815
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   1920
            Width           =   1575
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   840
            TabIndex        =   14
            Text            =   "Combo1"
            Top             =   3120
            Width           =   2415
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Check2"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   3120
            Width           =   255
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Value           =   1  'Checked
            Width           =   255
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Text            =   "Text8"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Text            =   "Text7"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Text            =   "Text6"
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "Text5"
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Text            =   "Text4"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   1800
            TabIndex        =   5
            Text            =   "Text2"
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton Command121 
            Caption         =   "Command1"
            Height          =   375
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Command2"
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   600
            Width           =   1575
         End
         Begin VB.CommandButton Command49 
            Caption         =   "Command4"
            Height          =   375
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1080
            Width           =   1575
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   240
            Picture         =   "frmmatcategory.frx":014A
            ScaleHeight     =   225
            ScaleWidth      =   240
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   3600
            Width           =   240
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Aditional Information:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Label2"
            Height          =   255
            Left            =   600
            TabIndex        =   49
            Top             =   3600
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Go To >>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Height          =   300
      Left            =   1200
      Picture         =   "frmmatcategory.frx":04F9
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Next >>"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      Height          =   255
      Left            =   2280
      Picture         =   "frmmatcategory.frx":0537
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Delete Page"
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Height          =   255
      Left            =   1680
      Picture         =   "frmmatcategory.frx":05B9
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Add Page"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text9 
      Height          =   300
      Left            =   2520
      TabIndex        =   57
      Text            =   "1"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   300
      Left            =   840
      Picture         =   "frmmatcategory.frx":0695
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Last Page >>|"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Height          =   300
      Left            =   480
      Picture         =   "frmmatcategory.frx":06DB
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "|<< First Page"
      Top             =   2880
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   300
      Left            =   120
      Picture         =   "frmmatcategory.frx":0720
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "<< Previous"
      Top             =   2880
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0 of 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmmatcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Add first four pages...
Private Sub Form_Load()
    frmmatcategory.AddPage Picture1
    frmmatcategory.AddPage Picture2
    frmmatcategory.AddPage Picture3
    frmmatcategory.AddPage Picture4
    
End Sub

'================================
'==== Navigate through pages ====
'================================

'Go to Previous Page...
Private Sub Command1_Click()
    frmmatcategory.PreviousPage
    
End Sub

'Go to First Page...
Private Sub Command2_Click()
    frmmatcategory.FirstPage
    
End Sub

'Go to Last Page...
Private Sub Command5_Click()
    frmmatcategory.LastPage
    
End Sub

'Go to Next Page...
Private Sub Command3_Click()
    frmmatcategory.NextPage
    
End Sub

'Go to page number displayed on TextBox...
Private Sub Command6_Click()
    frmmatcategory.CurrentPage = Text9.Text
    
End Sub

'====================
'==== Edit pages ====
'====================

'Add page 5...
Private Sub Command8_Click()
    frmmatcategory.AddPage Picture5
    
End Sub

'Delete current page...
Private Sub Command9_Click()
    Call frmmatcategory.DeletePage(frmmatcategory.CurrentPage)
    
End Sub

 

 

'Put all the objects on proper
'place when resizing.
Private Sub Form_Resize()
    Dim intTempTop
    Dim intTempSpace
    
    On Error Resume Next
    
    'Prevent user from resizing the
    'Form into a size too small.
    If (Me.Height < 1500) Then
        Me.Height = 1500
    End If
    
    intTempTop = 600
    intTempSpace = frmmatcategory.Top + frmmatcategory.Height + intTempTop
    
    frmmatcategory.Width = Me.Width - 300
    frmmatcategory.Height = Me.Height - (2000 + intTempTop)
    
    Label1.Top = frmmatcategory.Top + frmmatcategory.Height + 100
    Command1.Top = Label1.Top + Label1.Height
    Command2.Top = Label1.Top + Label1.Height
    Command3.Top = Label1.Top + Label1.Height
    Command5.Top = Label1.Top + Label1.Height
    
    Command9.Top = frmmatcategory.Top + frmmatcategory.Height + 220
    Command8.Top = frmmatcategory.Top + frmmatcategory.Height + 220
    
    Command6.Top = Command8.Top + Command8.Height
    Text9.Top = Command8.Top + Command8.Height
    
    Command12.Top = frmmatcategory.Top + frmmatcategory.Height + 230
    
    Label9.Top = frmmatcategory.Top + frmmatcategory.Height + 900
End Sub

'On PageChanged event, update current
'page number, total of pages and if
'Navigation buttons should be enabled
'and/or visible.
Private Sub frmmatcategory_PageChanged()
    Label1.Caption = frmmatcategory.CurrentPage & " of " & frmmatcategory.HowManyPages
    Text9.Text = frmmatcategory.CurrentPage
    
    Command1.Enabled = frmmatcategory.PreviousEnabled
    Command2.Enabled = frmmatcategory.PreviousEnabled
    Command5.Enabled = frmmatcategory.NextEnabled
    Command3.Enabled = frmmatcategory.NextEnabled
    
    If (frmmatcategory.HowManyPages < 2) Then
        Command1.Visible = False
        Command2.Visible = False
        Command5.Visible = False
        Command3.Visible = False
    Else
        Command1.Visible = True
        Command2.Visible = True
        Command5.Visible = True
        Command3.Visible = True
    End If
    
End Sub
