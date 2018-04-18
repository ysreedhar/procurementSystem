VERSION 5.00
Begin VB.Form MaterialCategory1 
   BackColor       =   &H00A04729&
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00A04729&
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9255
      Begin VB.PictureBox pView 
         BackColor       =   &H00A04729&
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   8955
         TabIndex        =   5
         Top             =   120
         Width           =   8955
         Begin VB.CommandButton Command2 
            Caption         =   "Command1"
            Height          =   255
            Left            =   8280
            TabIndex        =   11
            Top             =   720
            Width           =   615
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   255
            Left            =   8280
            TabIndex        =   10
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   7
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   6
            Top             =   360
            Width           =   6615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   0
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   1560
            TabIndex        =   8
            Top             =   120
            Width           =   795
         End
      End
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   6615
   End
   Begin VB.TextBox txt_categorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "MaterialCategory1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyTextboxArray As Object 'A class level dynamic Array
Private MyDescboxarray As Object


Private Sub Command1_Click()

     With addTextbox
          .Top = Text1(MyTextboxArray.ubound - 1).Top + Text1(MyTextboxArray.ubound - 1).Height
                    .Visible = True
          .SetFocus
      End With


     With adddescbox
          .Top = Text2(MyDescboxarray.ubound - 1).Top + Text2(MyDescboxarray.ubound - 1).Height
                    .Visible = True
          .SetFocus
      End With
pView.Height = pView.Height + 400
Frame1.Height = Frame1.Height + 400
Me.Height = Me.Height + 400
 
End Sub
Private Sub Command2_Click()
pView.Height = pView.Height - 400
Frame1.Height = Frame1.Height - 400
Me.Height = Me.Height - 400
 
End Sub

Private Sub Form_Initialize()
    Set MyTextboxArray = Me.Controls("Text1")
    Set MyDescboxarray = Me.Controls("Text2")
    
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

Public Function delTextbox() As TextBox
 
End Function
