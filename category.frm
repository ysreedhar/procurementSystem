VERSION 5.00
Begin VB.Form category 
   BackColor       =   &H00A04729&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LEVEL-1(PLWC)"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_categorycode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txt_category 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   7695
   End
   Begin VB.TextBox txt_notes 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1800
      Width           =   7695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   795
   End
End
Attribute VB_Name = "category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
