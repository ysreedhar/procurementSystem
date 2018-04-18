VERSION 5.00
Begin VB.Form department 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Department"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_dcode 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txt_dname 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code "
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   420
   End
End
Attribute VB_Name = "department"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
