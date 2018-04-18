VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchasereturns 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Purchase Returns"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16744576
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
      TabCaption(0)   =   "Purchase Returns"
      TabPicture(0)   =   "purchasereturns.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Notes"
      TabPicture(1)   =   "purchasereturns.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   6615
         Left            =   -75000
         TabIndex        =   15
         Top             =   360
         Width           =   9735
         Begin VB.TextBox txt_notes 
            Height          =   5535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   16
            Top             =   240
            Width           =   8775
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   7335
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   9735
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   5880
            TabIndex        =   26
            Top             =   240
            Width           =   3735
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reason"
            Height          =   615
            Left            =   2640
            TabIndex        =   21
            Top             =   720
            Width           =   6375
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Others"
               Height          =   255
               Left            =   4320
               TabIndex        =   25
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Option4 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Expired"
               Height          =   255
               Left            =   2920
               TabIndex        =   24
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Option3 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Damaged"
               Height          =   255
               Left            =   1520
               TabIndex        =   23
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton Option1 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Scrap"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select BatchNo."
            Height          =   1335
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   2415
            Begin VB.ListBox List1 
               Height          =   960
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   18
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox txt_prn 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   11
            Top             =   240
            Width           =   1575
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Height          =   4935
            Left            =   120
            TabIndex        =   2
            Top             =   1560
            Width           =   8895
            Begin VB.CommandButton cmd_new 
               BackColor       =   &H00FF8080&
               Caption         =   "New"
               Height          =   275
               Left            =   7920
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   240
               Width           =   735
            End
            Begin VB.CommandButton Command1 
               BackColor       =   &H00FF8080&
               Caption         =   "Save"
               Height          =   275
               Left            =   7920
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txt_tqty 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   6600
               TabIndex        =   19
               Top             =   480
               Width           =   975
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFC0C0&
               Height          =   975
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   1335
               Begin VB.OptionButton opt_med 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Medicine"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   7
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFC0C0&
                  Caption         =   "Other Item"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   6
                  Top             =   600
                  Width           =   1095
               End
            End
            Begin VB.ComboBox cbo_name 
               Height          =   315
               Left            =   1440
               TabIndex        =   4
               Top             =   480
               Width           =   4215
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5715
               TabIndex        =   3
               Top             =   480
               Width           =   615
            End
            Begin MSFlexGridLib.MSFlexGrid flex_med 
               Height          =   3975
               Left            =   0
               TabIndex        =   8
               Top             =   960
               Width           =   8895
               _ExtentX        =   15690
               _ExtentY        =   7011
               _Version        =   393216
               Rows            =   3
               Cols            =   5
               FixedCols       =   0
               RowHeightMin    =   250
               BackColor       =   12640511
               ForeColor       =   12582912
               BackColorFixed  =   16761024
               ForeColorFixed  =   4210816
               BackColorBkg    =   16761024
               AllowUserResizing=   3
               Appearance      =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Total Qty"
               Height          =   195
               Left            =   6600
               TabIndex        =   20
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Medicine"
               Height          =   195
               Left            =   1440
               TabIndex        =   10
               Top             =   240
               Width           =   4215
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Quantity"
               Height          =   195
               Left            =   5715
               TabIndex        =   9
               Top             =   240
               Width           =   615
            End
         End
         Begin MSComCtl2.DTPicker dtp_prn 
            Height          =   285
            Left            =   4320
            TabIndex        =   12
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   20709377
            CurrentDate     =   38455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Vendor"
            Height          =   195
            Left            =   5880
            TabIndex        =   27
            Top             =   0
            Width           =   3735
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "PRN No."
            Height          =   195
            Left            =   2640
            TabIndex        =   14
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "PRN Date"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   4320
            TabIndex        =   13
            Top             =   0
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "purchasereturns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kl As Integer
Private Sub cmd_new_Click()
kl = 1
End Sub

Private Sub Command1_Click()
Dim jj As Integer
jj = 0
        If kl = 1 Then
                    With flex_med
                        
                        .Rows = .Rows + 1
                        If opt_med.Value = True Then
                        .TextMatrix(.Rows - 1, 1) = opt_med.Caption
                        Else
                        .TextMatrix(.Rows - 1, 1) = opt_item.Caption
                        End If
                        .TextMatrix(.Rows - 1, 2) = cbo_name.Text
                        .TextMatrix(.Rows - 1, 3) = txt_qty.Text
                        .TextMatrix(.Rows - 1, 4) = txt_tqty.Text
                    
                    End With
        Else
                      jj = flex_med.Row
                      
                        If opt_med.Value = True Then
                        flex_med.TextMatrix(jj, 1) = opt_med.Caption
                        Else
                        flex_med.TextMatrix(jj, 1) = opt_item.Caption
                        End If
                        flex_med.TextMatrix(jj, 2) = cbo_name.Text
                        flex_med.TextMatrix(jj, 3) = txt_qty.Text
                        flex_med.TextMatrix(jj, 4) = txt_tqty.Text
        
        End If
opt_med.Value = False
opt_item.Value = False
cbo_name.Text = ""
txt_qty.Text = ""
txt_tqty.Text = ""
End Sub
Private Sub Form_Load()
dtp_prn.Value = Date
List1.AddItem "NA"
Dim mdb As New ADODB.Recordset
If mdb.State Then mdb.Close
mdb.Open "select * from medicinebatch  order by batchno", Cn, 3, 2
While Not mdb.EOF
List1.AddItem mdb!batchno
mdb.MoveNext
Wend




On Error Resume Next
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from purchasereturn", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "PRN000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from purchasereturn where prnno='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_prn.Text = "PRN000" & i
   Else
   i = i + 1
 
GoTo assad
  End If
End Sub

Private Sub List1_Click()

End Sub

Private Sub List1_ItemCheck(Item As Integer)
If List1.Selected(Item) = True Then
assad:
Dim pd As New ADODB.Recordset
If pd.State Then pd.Close
pd.Open "select * from prndetails where prnno='" & List1.List(Item) & "' ", Cn, 3, 2

   With flex_med
While Not pd.EOF
                         .Rows = .Rows + 1
                        If pd!moi = "Medicine" Then
                        flex_med.TextMatrix(.Rows - 1, 1) = "Medicine"
                        Else
                        flex_med.TextMatrix(.Rows - 1, 1) = "Other Item"
                        End If
                        flex_med.TextMatrix(.Rows - 1, 2) = pd!Name
                        flex_med.TextMatrix(.Rows - 1, 3) = pd!qty
                        flex_med.TextMatrix(.Rows - 1, 4) = pd!tqty
                        
                           

pd.MoveNext
Wend
End With
Else
flex_med.Clear
flex_med.Rows = 1
GoTo assad

End If
End Sub

Private Sub opt_med_Click()
lblmi.Caption = "Medicine"
End Sub

Private Sub Option2_Click()
lblmi.Caption = "Other Item"
End Sub
Public Sub flex_titlepr()
On Error Resume Next

   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
        .TextMatrix(0, 1) = "Med/Item"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Name"
        .ColWidth(2) = 4500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Qty"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Tqty"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0
         
    End With
End Sub
Private Sub flex_med_Click()
On Error Resume Next
'back color
 
Static vprev As Integer

current = flex_med.Row

'Reset to previous row
If vprev > 0 Then
    flex_med.Row = vprev
    flex_med.Col = 1
    Set flex_med.CellPicture = LoadPicture()
    
    For i = 1 To flex_med.Cols - 1
    flex_med.Col = i
    flex_med.CellBackColor = vbWhite
Next
End If

'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = vbYellow
Next
flex_med.Col = 1


vprev = flex_med.Row
End Sub

Private Sub flex_med_DblClick()
On Error Resume Next
'back color
 
Static vprev As Integer

current = flex_med.Row

'Reset to previous row
If vprev > 0 Then
    flex_med.Row = vprev
    flex_med.Col = 1
    Set flex_med.CellPicture = LoadPicture()
    
    For i = 1 To flex_med.Cols - 1
    flex_med.Col = i
    flex_med.CellBackColor = vbWhite
Next
End If

'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = vbYellow
Next
flex_med.Col = 1

Dim idd As Double
idd = 0
 

 
 If flex_med.TextMatrix(flex_med.Row, 1) = "Medicine" Then
 opt_med.Value = True
 Else
 opt_item.Value = True
 End If
 cbo_name.Text = flex_med.TextMatrix(flex_med.Row, 2)
 txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 3)
 txt_tqty.Text = flex_med.TextMatrix(flex_med.Row, 4)




vprev = flex_med.Row
End Sub
