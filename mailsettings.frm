VERSION 5.00
Begin VB.Form mailsettings 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Mail Settings"
   ClientHeight    =   9300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9300
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_save 
      BackColor       =   &H00FFC0C0&
      Caption         =   "<< Save Settings >>"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Picture         =   "mailsettings.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click to Authorize MSR"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txt_replyemail 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox txt_fromname 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   4695
   End
   Begin VB.TextBox txt_fromemail 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txt_smtp 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "cpxeon.crest.com.my"
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reply Email Address"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SMTP"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   450
   End
End
Attribute VB_Name = "mailsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
