VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form stockposting 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stock Posting"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Item Details"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txt_qty 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cbo_batchno 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2835
      End
      Begin VB.ComboBox cbo_uom 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         TabIndex        =   3
         Top             =   1320
         Width           =   1395
      End
      Begin VB.ComboBox cbo_category 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   6915
      End
      Begin MSComCtl2.DTPicker DTP_date 
         Height          =   315
         Left            =   3000
         TabIndex        =   6
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   67305473
         CurrentDate     =   38733
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5760
         TabIndex        =   10
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Posting Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Material batch"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "UOM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   4320
         TabIndex        =   7
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A04729&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   405
      End
   End
End
Attribute VB_Name = "stockposting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbo_category_Change()
sc = Split(cbo_category.Text, "  -  ", Len(cbo_category.Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(2) & "' and ml4name='" & sc(3) & "'   order by ml4uom", Cn, 3, 2
If Not um.EOF Then
cbo_uom.AddItem um(0)
End If
End Sub
