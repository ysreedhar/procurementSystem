VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form delivery 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delivery Order"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11456
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
      TabCaption(0)   =   "DO Line Items"
      TabPicture(0)   =   "delivery.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Shipping / Terms and Conditions / Remarks"
      TabPicture(1)   =   "delivery.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00A04729&
         BorderStyle     =   0  'None
         Height          =   6255
         Left            =   0
         TabIndex        =   24
         Top             =   300
         Width           =   11055
         Begin VB.ComboBox cbo_vendor 
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   6735
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00A04729&
            Height          =   2175
            Left            =   120
            TabIndex        =   32
            Top             =   2640
            Width           =   10815
            Begin VB.ComboBox cbo_name 
               Height          =   315
               Left            =   4680
               TabIndex        =   42
               Top             =   480
               Width           =   5955
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   2235
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               Left            =   2400
               TabIndex        =   40
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox Text2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   39
               Top             =   1800
               Width           =   10455
            End
            Begin VB.ComboBox cbo_uom 
               Height          =   315
               Left            =   1290
               TabIndex        =   38
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   37
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox Text3 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2715
               TabIndex        =   36
               Top             =   1200
               Width           =   1080
            End
            Begin VB.TextBox Text4 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6000
               TabIndex        =   35
               Top             =   1200
               Width           =   1800
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               Left            =   3885
               TabIndex        =   34
               Top             =   1200
               Width           =   1095
            End
            Begin VB.TextBox Text5 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   5070
               TabIndex        =   33
               Top             =   1200
               Width           =   840
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   285
               Left            =   7830
               TabIndex        =   43
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   67371009
               CurrentDate     =   38455
            End
            Begin MSComCtl2.DTPicker dtp_expiration 
               Height          =   285
               Left            =   9240
               TabIndex        =   44
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Format          =   67371009
               CurrentDate     =   38455
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
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
               Index           =   1
               Left            =   7845
               TabIndex        =   56
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   55
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Material Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Material Sub Category"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   53
               Top             =   240
               Width           =   2235
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   1560
               Width           =   10455
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1278
               TabIndex        =   51
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   1065
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Promised Date"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Index           =   0
               Left            =   9240
               TabIndex        =   49
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Unit Rate"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   2706
               TabIndex        =   48
               Top             =   960
               Width           =   1080
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Amount(RM)"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   6000
               TabIndex        =   47
               Top             =   960
               Width           =   1800
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Currency"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   3885
               TabIndex        =   46
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00A04729&
               Caption         =   "Exch Rate"
               ForeColor       =   &H00FFC0FF&
               Height          =   195
               Left            =   5070
               TabIndex        =   45
               Top             =   960
               Width           =   840
            End
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   6735
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00FF8080&
            Caption         =   "Save"
            Height          =   275
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2280
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00FF8080&
            Caption         =   "Delete"
            Height          =   275
            Left            =   1080
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2280
            Width           =   855
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00A04729&
            Caption         =   "Select  PO No."
            ForeColor       =   &H8000000E&
            Height          =   1335
            Left            =   8520
            TabIndex        =   27
            Top             =   0
            Width           =   2415
            Begin VB.ListBox List1 
               Height          =   960
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   28
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox txt_account 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Text9 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3360
            TabIndex        =   25
            Top             =   360
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dtp_po 
            Height          =   285
            Left            =   1800
            TabIndex        =   58
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   67371009
            CurrentDate     =   38455
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   1335
            Left            =   120
            TabIndex        =   59
            Top             =   4800
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   2355
            _Version        =   393216
            Rows            =   1
            Cols            =   5
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   10503977
            ForeColorFixed  =   16777215
            BackColorSel    =   10503977
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            Appearance      =   0
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   5400
            TabIndex        =   60
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   67371009
            CurrentDate     =   38455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "Vendor"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   840
            Width           =   6735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "Contact Person"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   6735
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "DO Date"
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
            TabIndex        =   64
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "DO No."
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "DO No."
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Left            =   3360
            TabIndex        =   62
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackColor       =   &H00A04729&
            Caption         =   "DO Date"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0FF&
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   61
            Top             =   120
            Width           =   1455
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
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Terms / Conditions"
            Height          =   2295
            Left            =   120
            TabIndex        =   10
            Top             =   1920
            Width           =   10815
            Begin VB.TextBox Text14 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   16
               Top             =   1920
               Width           =   1095
            End
            Begin VB.TextBox Text15 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   15
               Top             =   1200
               Width           =   10575
            End
            Begin VB.TextBox Text11 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   14
               Top             =   480
               Width           =   10575
            End
            Begin VB.ComboBox Combo6 
               Height          =   315
               Left            =   1320
               TabIndex        =   13
               Top             =   1920
               Width           =   1455
            End
            Begin VB.TextBox Text12 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3240
               TabIndex        =   12
               Top             =   1920
               Width           =   1095
            End
            Begin VB.ComboBox Combo7 
               Height          =   315
               Left            =   4440
               TabIndex        =   11
               Top             =   1920
               Width           =   1455
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Validity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   22
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Delivery"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   960
               Width           =   10575
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Price"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   10575
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Duration"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1320
               TabIndex        =   19
               Top             =   1680
               Width           =   1455
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Terms"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   3240
               TabIndex        =   18
               Top             =   1680
               Width           =   1095
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Duration"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4440
               TabIndex        =   17
               Top             =   1680
               Width           =   1455
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Shipping Details"
            Height          =   1815
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   10815
            Begin VB.TextBox Text6 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   6
               Top             =   1080
               Width           =   2655
            End
            Begin VB.TextBox Text7 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4200
               TabIndex        =   5
               Top             =   480
               Width           =   4455
            End
            Begin VB.TextBox Text8 
               Appearance      =   0  'Flat
               Height          =   1245
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   1  'Horizontal
               TabIndex        =   4
               Top             =   480
               Width           =   3975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Tel No"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4200
               TabIndex        =   9
               Top             =   840
               Width           =   2655
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Contact person"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4200
               TabIndex        =   8
               Top             =   240
               Width           =   4455
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               Caption         =   "Ship-To-Party"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.TextBox txt_notes 
            Height          =   1575
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   2
            Top             =   4560
            Width           =   10815
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            Caption         =   "Remarks"
            ForeColor       =   &H8000000E&
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   4320
            Width           =   10815
         End
      End
   End
End
Attribute VB_Name = "delivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public kl As Integer
Public tqty As Double
Public s9 As Double
Public s2 As Double


Private Sub cbo_name_Change()
cbo_batch.Clear
Dim mdb As New ADODB.Recordset
If mdb.State Then mdb.Close
mdb.Open "select * from medicinebatch  where medicine ='" & cbo_name.Text & "'", Cn, 3, 2
While Not mdb.EOF
cbo_batch.AddItem mdb!batchno
mdb.MoveNext
Wend
End Sub

Private Sub cbo_name_Click()
cbo_batch.Clear
Dim mdb As New ADODB.Recordset
If mdb.State Then mdb.Close
mdb.Open "select * from medicinebatch  where medicine ='" & cbo_name.Text & "'", Cn, 3, 2
While Not mdb.EOF
cbo_batch.AddItem mdb!batchno
mdb.MoveNext
Wend

End Sub

Private Sub cbo_po_Click()
flex_med.Clear
flex_med.Rows = 1
Dim pr As New ADODB.Recordset
If pr.State Then pr.Close
pr.Open "select * from purchaseorder where pono='" & cbo_po.Text & "' ", Cn, 3, 2
If Not pr.EOF Then
cbo_vendor.Text = pr!vendor

     Dim por As New ADODB.Recordset
     If por.State Then por.Close
     por.Open "select * from podetails where pono='" & cbo_po.Text & "'", Cn, 3, 2
With flex_med
While Not por.EOF
                         .Rows = .Rows + 1

                        flex_med.TextMatrix(.Rows - 1, 1) = por!Name
                        flex_med.TextMatrix(.Rows - 1, 2) = por!uom
                        flex_med.TextMatrix(.Rows - 1, 3) = por!qty
                        flex_med.TextMatrix(.Rows - 1, 4) = por!reqdate
                        
                           

por.MoveNext
Wend
End With
End If

Call flex_titlepr

End Sub

Private Sub cmd_new_Click()
kl = 1
cbo_name.Clear
End Sub

Private Sub Command1_Click()
On Error Resume Next

s2 = 0
s2 = s9 + CDbl(txt_qty.Text)
Dim jj As Integer
jj = 0
        If kl = 1 Then
                    With flex_med
                        
                        .Rows = .Rows + 1
 
                        .TextMatrix(.Rows - 1, 1) = cbo_name.Text
                        .TextMatrix(.Rows - 1, 2) = cbo_uom.Text
                        .TextMatrix(.Rows - 1, 3) = txt_qty.Text
                        .TextMatrix(.Rows - 1, 4) = dtp_expiration.Value
                        .TextMatrix(.Rows - 1, 5) = cbo_batch.Text
                    End With
        Else
                      jj = flex_med.Row
 
                        flex_med.TextMatrix(jj, 1) = cbo_name.Text
                        flex_med.TextMatrix(jj, 2) = cbo_uom.Text
                        flex_med.TextMatrix(jj, 3) = txt_qty.Text
                        flex_med.TextMatrix(jj, 4) = dtp_expiration.Value
                        flex_med.TextMatrix(jj, 5) = cbo_batch.Text
        End If
         
 
cbo_name.Text = ""
txt_qty.Text = ""
dtp_expiration.Value = Date
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
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
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
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1

Dim idd As Double
idd = 0
 

 


    cbo_name.Text = flex_med.TextMatrix(flex_med.Row, 1)
    cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 2)
    txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 3)
    dtp_expiration.Value = flex_med.TextMatrix(flex_med.Row, 4)
    cbo_batch.Text = flex_med.TextMatrix(flex_med.Row, 5)
 
tqty = 0
Dim pd As New ADODB.Recordset
If pd.State Then pd.Close
pd.Open "select total  from  medicinebatch  where  batchno='" & flex_med.TextMatrix(flex_med.Row, 5) & "' and medicine ='" & flex_med.TextMatrix(flex_med.Row, 1) & "'", Cn, 3, 2
If Not pd.EOF Then

 tqty = pd(0)

End If

s9 = 0
s9 = CDbl(tqty) - CDbl(txt_qty.Text)

vprev = flex_med.Row
End Sub

Private Sub Form_Load()
cbo_po.AddItem "NA"
Dim po As New ADODB.Recordset
If po.State Then po.Close
po.Open "select * from purchaseorder order by pono", Cn, 3, 2
While Not po.EOF
cbo_po.AddItem po!pono
po.MoveNext
Wend
flex_titlepr
cbo_batch.Text = "NA"

On Error Resume Next
Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from delivery", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "GRN000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from delivery where invoice='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_invoice.Text = "GRN000" & i
   Else
   i = i + 1
 
GoTo assad
  End If

Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(name) from vendor order by name", Cn, 3, 2
While Not vn.EOF
cbo_vendor.AddItem vn(0)
vn.MoveNext
Wend
  Dim med As New ADODB.Recordset
  If med.State Then med.Close
  med.Open "select DISTINCT(medicine) from medicinebatch order by medicine", Cn, 3, 2
  While Not med.EOF
  cbo_name.AddItem med(0)
  med.MoveNext
  Wend
kl = 0

End Sub
Public Sub flex_titlepr()
On Error Resume Next

   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        
 
        .TextMatrix(0, 1) = "Name"
        .ColWidth(1) = 4500
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "UOM"
        .ColWidth(2) = 1200
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Qty"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Req Date"
        .ColWidth(4) = 2000
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "BatchNo"
        .ColWidth(5) = 2000
        .ColAlignment(5) = 0
         
    End With
End Sub



