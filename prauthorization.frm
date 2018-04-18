VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchaseauthorization 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PR Authorization"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   10875
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   10398
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
      TabCaption(0)   =   "PR Authorization"
      TabPicture(0)   =   "prauthorization.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Shipping / Other Details / Remarks"
      TabPicture(1)   =   "prauthorization.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   6015
         Left            =   0
         TabIndex        =   2
         Top             =   300
         Width           =   11055
         Begin VB.TextBox txt_costcode 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   53
            Top             =   2040
            Width           =   2775
         End
         Begin VB.TextBox txt_jobcharge 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4080
            TabIndex        =   52
            Top             =   2040
            Width           =   3855
         End
         Begin VB.TextBox txt_project 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Top             =   2040
            Width           =   3855
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H008080FF&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   8175
            Begin VB.CheckBox chk_all 
               BackColor       =   &H008080FF&
               Caption         =   "Apply to All Line Items"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   240
               TabIndex        =   46
               Top             =   600
               Width           =   2175
            End
            Begin VB.CheckBox chk_ind 
               BackColor       =   &H008080FF&
               Caption         =   "Apply to Individual Line Item"
               ForeColor       =   &H00C00000&
               Height          =   255
               Left            =   4980
               TabIndex        =   45
               Top             =   600
               Value           =   1  'Checked
               Width           =   2295
            End
            Begin VB.ComboBox cbo_astatus 
               Height          =   315
               Left            =   240
               TabIndex        =   42
               Top             =   240
               Width           =   1635
            End
            Begin VB.ComboBox cbo_abuyer 
               Height          =   315
               Left            =   1920
               TabIndex        =   41
               Top             =   240
               Width           =   5355
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Authorization Status"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   240
               TabIndex        =   44
               Top             =   0
               Width           =   1635
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Authorized Buyer"
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   1920
               TabIndex        =   43
               Top             =   0
               Width           =   5355
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   1695
            Left            =   120
            TabIndex        =   11
            Top             =   2280
            Width           =   10695
            Begin VB.TextBox txt_uom 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1330
               TabIndex        =   50
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_material 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4695
               TabIndex        =   49
               Top             =   480
               Width           =   5955
            End
            Begin VB.TextBox txt_subcategory 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2400
               TabIndex        =   48
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_category 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   47
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_remarks 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4260
               TabIndex        =   13
               Top             =   1200
               Width           =   6375
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   12
               Top             =   1200
               Width           =   1080
            End
            Begin MSComCtl2.DTPicker dtp_date 
               Height          =   285
               Left            =   2795
               TabIndex        =   14
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   56033281
               CurrentDate     =   38455
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Reqd. Date"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   195
               Index           =   0
               Left            =   2810
               TabIndex        =   21
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   20
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Material Sub Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   18
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4260
               TabIndex        =   17
               Top             =   960
               Width           =   6375
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1330
               TabIndex        =   16
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   960
               Width           =   1065
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Select    PR No."
            ForeColor       =   &H00000000&
            Height          =   1575
            Left            =   8400
            TabIndex        =   5
            Top             =   0
            Width           =   2415
            Begin VB.ListBox List1 
               Height          =   1185
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   6
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.ComboBox cbo_personincharge 
            Height          =   315
            Left            =   3360
            TabIndex        =   4
            Top             =   360
            Width           =   4935
         End
         Begin VB.TextBox txt_account 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker dtp_pa 
            Height          =   285
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   56033281
            CurrentDate     =   38455
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   1575
            Left            =   120
            TabIndex        =   22
            Top             =   3960
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   2778
            _Version        =   393216
            Rows            =   1
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
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Jobcharge"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   2
            Left            =   4080
            TabIndex        =   39
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Project "
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   1800
            Width           =   3855
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Cost Code"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   3
            Left            =   8040
            TabIndex        =   37
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Person Incharge"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   3360
            TabIndex        =   10
            Top             =   120
            Width           =   4935
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization Date"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   195
            Index           =   1
            Left            =   1800
            TabIndex        =   9
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Authorization ID"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   11055
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Shipping Details"
            Height          =   1815
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   10695
            Begin VB.TextBox txt_telno 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   33
               Top             =   1080
               Width           =   2655
            End
            Begin VB.TextBox txt_contactperson 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   32
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox txt_shiptoparty 
               Appearance      =   0  'Flat
               Height          =   1245
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   1  'Horizontal
               TabIndex        =   31
               Top             =   480
               Width           =   3975
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Tel No"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4200
               TabIndex        =   36
               Top             =   840
               Width           =   2655
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Contact person"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4200
               TabIndex        =   35
               Top             =   240
               Width           =   4455
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Ship-To-Party"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.TextBox txt_notes 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   28
            Top             =   3960
            Width           =   10575
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Height          =   1815
            Left            =   0
            TabIndex        =   23
            Top             =   1800
            Width           =   10695
            Begin VB.ComboBox cbo_recommendedvendor 
               Height          =   315
               Left            =   120
               TabIndex        =   25
               Top             =   1320
               Width           =   10395
            End
            Begin VB.TextBox txt_justification 
               Appearance      =   0  'Flat
               Height          =   645
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   24
               Top             =   360
               Width           =   10455
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Recommended Vendor"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   135
               TabIndex        =   27
               Top             =   1080
               Width           =   10395
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Justification / Purpose"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   120
               Width           =   10455
            End
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Remarks"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   3720
            Width           =   10575
         End
      End
   End
End
Attribute VB_Name = "purchaseauthorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub chk_all_Click()
If chk_all.Value = 1 Then
MsgBox "Authorization Status  :" & cbo_astatus.Text & " and Authorized Buyer  :" & cbo_abuyer.Text & " Has been applied to all the Line Items"
End If
End Sub

Private Sub chk_ind_Click()
If chk_ind.Value = 1 Then
MsgBox "Click on the specific Line Item you want to apply Status and Buyer", vbOKOnly
End If
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
 kl = 0
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
 
cbo_category.Text = flex_med.TextMatrix(flex_med.Row, 1)
cbo_subcategory.Text = flex_med.TextMatrix(flex_med.Row, 2)
cbo_material.Text = flex_med.TextMatrix(flex_med.Row, 3)
txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 4)
cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 5)
dtp_reqd.Value = flex_med.TextMatrix(flex_med.Row, 6)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 7)


vprev = flex_med.Row
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from purchaseauthorization", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "PA000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from purchaseauthorization where pano='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_account.Text = "PA000" & i
   Else
   i = i + 1
 
GoTo assad
  End If

  
  
  
  '------------------
  Dim lst As New ADODB.Recordset
  If lst.State Then lst.Close
  lst.Open "select DISTINCT(prno) from purchaserequisition where status <> 'Authorized' order by prno ", cn3, 2
  While Not ls.EOF
  List1.AddItem lst(0)
  lst.MoveNext
  Wend
    
  
    Call flex_titlepa
    kl = 1
End Sub
Public Sub flex_titlepa()
On Error Resume Next

   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "Category"
        .ColWidth(1) = 1200
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "SubCategory"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Material"
        .ColWidth(3) = 3000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Qty"
        .ColWidth(4) = 1200
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "UOM"
        .ColWidth(5) = 1200
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "ReqDate"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Remarks"
        .ColWidth(7) = 1200
        .ColAlignment(7) = 0
         
    End With
End Sub

Private Sub List1_Click()

End Sub

Private Sub List1_ItemCheck(Item As Integer)
If List1.SelCount > 1 Then
MsgBox "Only One PR can be Authorized at a time"
Exit Sub
End If
If List1.Selected(Item) = True Then
 
Dim pd As New ADODB.Recordset
If pd.State Then pd.Close
pd.Open "select * from prdetails where prno='" & List1.List(Item) & "' ", Cn, 3, 2

With flex_med
While Not pd.EOF
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, 1) = pd!category
                    .TextMatrix(.Rows - 1, 2) = pd!subcategory
                    .TextMatrix(.Rows - 1, 3) = pd!material
                    .TextMatrix(.Rows - 1, 4) = pd!qty
                    .TextMatrix(.Rows - 1, 5) = pd!uom
                    .TextMatrix(.Rows - 1, 6) = pd!reqdate
                    .TextMatrix(.Rows - 1, 7) = pd!remarks
                                         

pd.MoveNext
Wend
End With

Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from purchaserequisition where prno='" & List1.List(Item) & "' ", Cn, 3, 2
If Not pr.EOF Then
txt_project.Text = pr!project
txt_jobcharge.Text = pr!jobcharge
txt_costcode.Text = pr!costcode
txt_shiptoparty.Text = pr!shiptoaddress
txt_contactperson.Text = pr!contactperson
txt_telno.Text = pr!telno

txt_justification.Text = pr!justification
cbo_recommendedvendor.Text = pr!recommendedvendor
txt_notes.Text = pr!notes

End If




End If
End Sub
