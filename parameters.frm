VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form parameters 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parameters"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Parameters"
      TabPicture(0)   =   "parameters.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "parameters.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4095
         Left            =   0
         TabIndex        =   10
         Top             =   360
         Width           =   6855
         Begin VB.TextBox txt_website 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   7
            Top             =   3315
            Width           =   2655
         End
         Begin VB.TextBox txt_email 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   6
            Top             =   3315
            Width           =   2655
         End
         Begin VB.TextBox txt_fax 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2880
            TabIndex        =   5
            Top             =   2715
            Width           =   2655
         End
         Begin VB.TextBox txt_phone 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Top             =   2715
            Width           =   2655
         End
         Begin VB.TextBox txt_name 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3615
         End
         Begin VB.TextBox txt_address 
            Appearance      =   0  'Flat
            Height          =   765
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   3
            Top             =   1560
            Width           =   5415
         End
         Begin VB.TextBox txt_regno 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   1
            Top             =   915
            Width           =   3615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Website"
            Height          =   195
            Left            =   2880
            TabIndex        =   17
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Email"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   3120
            Width           =   2655
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fax"
            Height          =   195
            Left            =   2880
            TabIndex        =   15
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Phone"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Address"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   5415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reg No"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Name"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3735
         Left            =   -75000
         TabIndex        =   8
         Top             =   360
         Width           =   6735
         Begin VB.TextBox txt_notes 
            Height          =   3135
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   9
            Top             =   240
            Width           =   5535
         End
      End
   End
End
Attribute VB_Name = "parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
