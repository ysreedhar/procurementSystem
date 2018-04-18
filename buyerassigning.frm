VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form buyerassigning 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buyer Assigning"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "buyerassigning.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   16777215
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Buyer Assigning"
      TabPicture(0)   =   "buyerassigning.frx":10D5F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "MSR Other Details"
      TabPicture(1)   =   "buyerassigning.frx":10D7B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   6855
         Left            =   -75000
         TabIndex        =   11
         Top             =   300
         Width           =   11655
         Begin VB.TextBox txt_notes 
            Height          =   2415
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   17
            Top             =   3600
            Width           =   10695
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            Height          =   3135
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   10695
            Begin VB.ComboBox cbo_recommendedvendor 
               Height          =   315
               Left            =   120
               TabIndex        =   14
               Top             =   2640
               Width           =   10395
            End
            Begin VB.TextBox txt_justification 
               Appearance      =   0  'Flat
               Height          =   1965
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               Top             =   360
               Width           =   10455
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Recommended Vendor"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   135
               TabIndex        =   16
               Top             =   2400
               Width           =   1950
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Justification / Purpose"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   15
               Top             =   120
               Width           =   1875
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00A04729&
            Height          =   2175
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Visible         =   0   'False
            Width           =   10695
            Begin VB.TextBox txt_jobcharge 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   28
               Top             =   1800
               Width           =   2235
            End
            Begin VB.TextBox txt_location 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   27
               Top             =   1800
               Width           =   2235
            End
            Begin VB.TextBox txt_contactperson 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   26
               Top             =   1800
               Width           =   5955
            End
            Begin VB.TextBox txt_uom 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   1330
               TabIndex        =   25
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txt_material 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4695
               TabIndex        =   24
               Top             =   480
               Width           =   5955
            End
            Begin VB.TextBox txt_subcategory 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   2400
               TabIndex        =   23
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_category 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   22
               Top             =   480
               Width           =   2235
            End
            Begin VB.TextBox txt_remarks 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   4260
               TabIndex        =   21
               Top             =   1200
               Width           =   6375
            End
            Begin VB.TextBox txt_qty 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   285
               Left            =   120
               TabIndex        =   20
               Top             =   1200
               Width           =   1080
            End
            Begin MSComCtl2.DTPicker dtp_reqd 
               Height          =   285
               Left            =   2795
               TabIndex        =   29
               Top             =   1200
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   67108865
               CurrentDate     =   38455
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Location"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   39
               Top             =   1560
               Width           =   615
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Jobcharge"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   38
               Top             =   1560
               Width           =   750
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   37
               Top             =   1560
               Width           =   570
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
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
               TabIndex        =   36
               Top             =   960
               Width           =   1305
            End
            Begin VB.Label lblmi 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Material Code/ Desc"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4695
               TabIndex        =   35
               Top             =   240
               Width           =   5955
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "ItemId"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   34
               Top             =   240
               Width           =   435
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Mfr.Ref Code"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   2400
               TabIndex        =   33
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Remarks"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   4260
               TabIndex        =   32
               Top             =   960
               Width           =   6375
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "UOM"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   1330
               TabIndex        =   31
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Quantity"
               ForeColor       =   &H8000000E&
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   960
               Width           =   1065
            End
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   3360
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8415
         Left            =   0
         TabIndex        =   1
         Top             =   300
         Width           =   11055
         Begin VB.CheckBox chk_app 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H8000000E&
            Height          =   1575
            Left            =   0
            TabIndex        =   40
            Top             =   120
            Width           =   10935
            Begin VB.TextBox txt_department 
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
               Height          =   285
               Left            =   120
               TabIndex        =   44
               Top             =   1080
               Width           =   3015
            End
            Begin VB.TextBox txt_requestor 
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
               Height          =   285
               Left            =   3240
               TabIndex        =   43
               Top             =   1080
               Width           =   5175
            End
            Begin VB.TextBox txt_project 
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
               Height          =   285
               Left            =   3240
               TabIndex        =   42
               Top             =   480
               Width           =   5175
            End
            Begin VB.TextBox txt_account 
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
               Height          =   285
               Left            =   120
               TabIndex        =   41
               Top             =   480
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker dtp_pa 
               Height          =   285
               Left            =   1800
               TabIndex        =   45
               Top             =   480
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   503
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   67108865
               CurrentDate     =   38455
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Department"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   50
               Top             =   840
               Width           =   1020
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Requestor"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   3240
               TabIndex        =   49
               Top             =   840
               Width           =   870
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Project "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   3240
               TabIndex        =   48
               Top             =   240
               Width           =   660
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Auth Date"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   1
               Left            =   1800
               TabIndex        =   47
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "MSR No."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   195
               Left            =   120
               TabIndex        =   46
               Top             =   240
               Width           =   720
            End
         End
         Begin VB.ComboBox cbo_personincharge 
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
            Left            =   240
            TabIndex        =   8
            Top             =   4320
            Visible         =   0   'False
            Width           =   8295
         End
         Begin VB.Frame fr_auth 
            BackColor       =   &H00FF8080&
            Caption         =   "BA Section"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   0
            TabIndex        =   2
            Top             =   5760
            Width           =   10935
            Begin VB.OptionButton opt_all 
               BackColor       =   &H00FF8080&
               Caption         =   "Apply All"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   4800
               TabIndex        =   6
               ToolTipText     =   "Click to Authorize all Line Items"
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton opt_ind 
               BackColor       =   &H00FF8080&
               Caption         =   "Clear"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000E&
               Height          =   255
               Left            =   6360
               TabIndex        =   5
               ToolTipText     =   "Click to Uncheck LineItems"
               Top             =   480
               Width           =   975
            End
            Begin VB.ComboBox cbo_buyer 
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
               ItemData        =   "buyerassigning.frx":10D97
               Left            =   120
               List            =   "buyerassigning.frx":10D99
               TabIndex        =   4
               ToolTipText     =   "Select Authorization Status"
               Top             =   480
               Width           =   4155
            End
            Begin VB.CommandButton cmd_confirmation 
               BackColor       =   &H00FFC0C0&
               Caption         =   " << Save Buyer Assignment >>"
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
               Left            =   7560
               Picture         =   "buyerassigning.frx":10D9B
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Click to  Confirm Approval"
               Top             =   195
               Width           =   2895
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Buyer"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   240
               Width           =   405
            End
         End
         Begin MSFlexGridLib.MSFlexGrid flex_med 
            Height          =   4095
            Left            =   0
            TabIndex        =   9
            Top             =   1680
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   7223
            _Version        =   393216
            Rows            =   1
            Cols            =   14
            FixedCols       =   0
            RowHeightMin    =   250
            BackColor       =   16777215
            ForeColor       =   10503977
            BackColorFixed  =   16744576
            ForeColorFixed  =   16777215
            BackColorSel    =   16744576
            BackColorBkg    =   16777215
            AllowUserResizing=   3
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "Person Incharge"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   4080
            Visible         =   0   'False
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "buyerassigning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kk As Integer
Public StrMsr As String
Public StrRcm As String
Public StrBuyer As String
Private Sub chk_app_Click(Index As Integer)
If opt_all.Value = False Then
If frm_buyerassigning.gg <> 0 Then
If chk_app(Index) = 1 Then
flex_med.Row = Index
flex_med.TextMatrix(flex_med.Row, 2) = cbo_buyer.Text

End If

End If
frm_buyerassigning.gg = 1
End If
End Sub
Private Sub cmd_confirmation_Click()
Dim jy As Integer
jy = 0
Dim by As String
With flex_med
For jy = 1 To flex_med.Rows - 1
'chk_app(jy).Enabled = False
Cn.Execute "update prdetails set buyer ='" & flex_med.TextMatrix(jy, 2) & " ' where prno='" & txt_account.Text & "' and pr_id='" & flex_med.TextMatrix(jy, 0) & "'"
Next
End With


'----------------------------------------------------------------------------
'rfq generation

Dim prfq As New ADODB.Recordset
If prfq.State Then prf.Close
prfq.Open "select DISTINCT(buyer) from prdetails where prno='" & txt_account.Text & "' and buyer <> 'NA' ", Cn, 3, 2
While Not prfq.EOF
StrBuyer = prfq(0)
StrMsr = txt_account.Text
StrRcm = "You have been assigned a task to issue RFQ to selected Vendors for MSR: " & StrMsr


Dim rf As New ADODB.Recordset
If rf.State Then rf.Close
rf.Open "select * from rfq", Cn, 3, 2
Dim i As Integer
   
   i = 1
assad:
Dim X As String
Dim t As String
t = ""
X = "TL-RFQ000" & i
   Dim rff As New ADODB.Recordset
   If rff.State Then rff.Close
   rff.Open "select * from rfq where rfqno='" & X & "' ", Cn, 3, 2
   If rff.EOF Then
   rff.AddNew
   rff!rfqno = "TL-RFQ000" & i
   t = "TL-RFQ000" & i
   rff!rfqdate = Format(Date, "dd/MM/yyyy")
   rff!closingdate = Format(Date, "dd/MM/yyyy")
   rff!Buyer = prfq(0)
   rff!prno = txt_account.Text
   Cn.Execute "update prdetails set rfqno = '" & t & "' where buyer =  '" & prfq(0) & "' and prno= '" & txt_account.Text & "'"
   rff.Update
   rff.Close
   Else
   i = i + 1
 
GoTo assad
  End If



'---------------------------------------------------------------------------------------------

prfq.MoveNext
Wend
prfq.Close
If main.VstrEmail = 0 Then
MsgBox "Buyer Assigned, System is sending mail to the Buyer: Kindly wait for confirmation message"
Call assademailservicebuyer

Else
MsgBox "Buyer Assigned"
End If
'end rfq generation
Call frm_buyerassigning.striptab
Call frm_buyerassigning.flex_itemmodi

'cmd_confirmation.Enabled = False

 
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

Private Sub flex_med_SelChange()
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
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1

Dim idd As Double
idd = 0


vprev = flex_med.Row
End Sub
Private Sub Form_Load()
On Error Resume Next
dtp_pa.Value = Format(Date, "dd/MM/yyyy")
 Me.Top = 2000
 Me.Left = 0
 Me.Height = 6990
 Me.Width = 10980
  If frm_buyerassigning.gg <> 0 Then
  '------------------
  Dim lst As New ADODB.Recordset
  If lst.State Then lst.Close
  lst.Open "select DISTINCT(p.prno) from purchaserequisition p,prdetails pr where p.prno=pr.prno and pr.buyer='NA' order by p.prno ", Cn, 3, 2
  While Not lst.EOF
  List1.AddItem lst(0)
  lst.MoveNext
  Wend
  End If
  lst.Close
  lst.Open "select DISTINCT(a_name) from userid where a_userid='" & main.Label2.Caption & "' ", Cn, 3, 2
If Not lst.EOF Then
cbo_personincharge.Text = lst(0)
End If
lst.Close
lst.Open "select DISTINCT(a_name) from userid where a_designation='Buyer'", Cn, 3, 2
While Not lst.EOF
cbo_buyer.AddItem lst(0)
lst.MoveNext
Wend




    Call flex_titlepa
    kl = 1
End Sub
Public Sub flex_titlepa()
On Error Resume Next
flex_med.Rows = 1
   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0
        .ColWidth(1) = 300
        .TextMatrix(0, 2) = "Buyer"
        .ColWidth(2) = 1500
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Status"
        .ColWidth(3) = 1200
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "ItemId"
        .ColWidth(4) = 1500
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "Mfr.Ref Code"
        .ColWidth(5) = 1200
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "Material"
        .ColWidth(6) = 5000
        .ColAlignment(6) = 0
        .TextMatrix(0, 7) = "Qty"
        .ColWidth(7) = 600
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "UOM"
        .ColWidth(8) = 600
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "ReqDate"
        .ColWidth(9) = 800
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Remarks"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0

        .TextMatrix(0, 11) = "Jobcharge"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0
        
        .TextMatrix(0, 12) = "Work Location"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0

        .TextMatrix(0, 13) = "Stor Loc"
        .ColWidth(13) = 1200
        .ColAlignment(13) = 0
         
    End With
End Sub
Private Sub opt_all_Click()
Dim g As Integer
g = 0

With flex_med
For g = 1 To flex_med.Rows - 1
chk_app(g).Value = 1
flex_med.TextMatrix(g, 2) = cbo_buyer.Text
Next
End With
opt_all.Value = False
End Sub
Private Sub opt_ind_Click()
Dim w As Integer
w = 0

With flex_med
For w = 1 To flex_med.Rows - 1
chk_app(w).Value = 0
flex_med.TextMatrix(w, 2) = cbo_buyer.Text
Next
End With
End Sub



