VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form purchaserequisition 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    MATERIALS & SERVICE REQUISITION"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "purchaserequisition.frx":0000
   ScaleHeight     =   8400
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7665
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13520
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " MSR Details"
      TabPicture(0)   =   "purchaserequisition.frx":11409
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Justification / Purpose"
      TabPicture(1)   =   "purchaserequisition.frx":11425
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame7"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Attach Files"
      TabPicture(2)   =   "purchaserequisition.frx":11441
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Line3"
      Tab(2).Control(1)=   "Frame13"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7485
         Left            =   0
         TabIndex        =   9
         Top             =   300
         Width           =   10935
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "INPUT / EDIT LINE ITEM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   2775
            Left            =   120
            TabIndex        =   27
            Top             =   1680
            Width           =   10695
            Begin TabDlg.SSTab SSTab2 
               Height          =   1935
               Left            =   120
               TabIndex        =   32
               Top             =   720
               Width           =   10455
               _ExtentX        =   18441
               _ExtentY        =   3413
               _Version        =   393216
               Style           =   1
               TabHeight       =   520
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Basic Data"
               TabPicture(0)   =   "purchaserequisition.frx":1145D
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Frame10"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Job Charge / Shipping Details"
               TabPicture(1)   =   "purchaserequisition.frx":11479
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame11"
               Tab(1).ControlCount=   1
               TabCaption(2)   =   "Line Item Remarks"
               TabPicture(2)   =   "purchaserequisition.frx":11495
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Frame12"
               Tab(2).ControlCount=   1
               Begin VB.Frame Frame12 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   1695
                  Left            =   -75000
                  TabIndex        =   35
                  Top             =   320
                  Width           =   10455
                  Begin VB.TextBox txt_remarks 
                     BeginProperty Font 
                        Name            =   "Tahoma"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1005
                     Left            =   0
                     TabIndex        =   51
                     Top             =   360
                     Width           =   10455
                  End
                  Begin VB.Label Label6 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Line Item Remarks"
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
                     Left            =   0
                     TabIndex        =   52
                     Top             =   120
                     Width           =   1635
                  End
               End
               Begin VB.Frame Frame11 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   1695
                  Left            =   -75000
                  TabIndex        =   34
                  Top             =   320
                  Width           =   10455
                  Begin VB.ComboBox cbo_jobcharge 
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
                     Left            =   0
                     TabIndex        =   47
                     Top             =   480
                     Width           =   4815
                  End
                  Begin VB.ComboBox cbo_worklocation 
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
                     Left            =   4920
                     TabIndex        =   46
                     Top             =   480
                     Width           =   2895
                  End
                  Begin VB.ComboBox cbo_location 
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
                     Left            =   7920
                     TabIndex        =   45
                     Top             =   480
                     Width           =   2535
                  End
                  Begin VB.Label Label25 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Job Charge No."
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
                     Index           =   3
                     Left            =   0
                     TabIndex        =   50
                     Top             =   240
                     Width           =   1320
                  End
                  Begin VB.Label Label17 
                     BackStyle       =   0  'Transparent
                     Caption         =   "Work Location"
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
                     Height          =   255
                     Left            =   4920
                     TabIndex        =   49
                     Top             =   240
                     Width           =   1455
                  End
                  Begin VB.Label Label18 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Storage Location"
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
                     Left            =   7920
                     TabIndex        =   48
                     Top             =   240
                     Width           =   1440
                  End
               End
               Begin VB.Frame Frame10 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   1695
                  Left            =   0
                  TabIndex        =   33
                  Top             =   320
                  Width           =   10455
                  Begin VB.Frame frame_ms 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Material Service Duration"
                     BeginProperty Font 
                        Name            =   "Verdana"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00808080&
                     Height          =   735
                     Left            =   4560
                     TabIndex        =   55
                     ToolTipText     =   $"purchaserequisition.frx":114B1
                     Top             =   120
                     Width           =   3735
                     Begin MSComCtl2.DTPicker dtp_from 
                        Height          =   285
                        Left            =   120
                        TabIndex        =   56
                        Top             =   360
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
                        Format          =   67502081
                        CurrentDate     =   38455
                     End
                     Begin MSComCtl2.DTPicker dtp_to 
                        Height          =   285
                        Left            =   2280
                        TabIndex        =   57
                        Top             =   360
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
                        Format          =   67502081
                        CurrentDate     =   38455
                     End
                     Begin VB.Label Label5 
                        BackColor       =   &H80000009&
                        Caption         =   "To"
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
                        Height          =   255
                        Left            =   1680
                        TabIndex        =   58
                        Top             =   360
                        Width           =   375
                     End
                  End
                  Begin VB.ComboBox cbo_materialtype 
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
                     ItemData        =   "purchaserequisition.frx":1154A
                     Left            =   1320
                     List            =   "purchaserequisition.frx":1154C
                     TabIndex        =   53
                     Top             =   240
                     Width           =   3075
                  End
                  Begin VB.TextBox txt_qty 
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
                     Left            =   6855
                     TabIndex        =   39
                     Top             =   1200
                     Width           =   960
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
                     Left            =   7875
                     TabIndex        =   38
                     Top             =   1200
                     Width           =   1215
                  End
                  Begin VB.ComboBox cbo_category 
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
                     Left            =   0
                     TabIndex        =   37
                     Top             =   1200
                     Width           =   6795
                  End
                  Begin VB.ComboBox cbo_lookup 
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
                     ItemData        =   "purchaserequisition.frx":1154E
                     Left            =   1320
                     List            =   "purchaserequisition.frx":11550
                     TabIndex        =   36
                     Top             =   840
                     Width           =   3075
                  End
                  Begin MSComCtl2.DTPicker dtp_reqd 
                     Height          =   285
                     Left            =   9150
                     TabIndex        =   40
                     Top             =   1200
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
                     Format          =   67502081
                     CurrentDate     =   38455
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Material Type"
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
                     Left            =   0
                     TabIndex        =   54
                     Top             =   360
                     Width           =   1155
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Quantity"
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
                     Left            =   6840
                     TabIndex        =   44
                     Top             =   960
                     Width           =   720
                  End
                  Begin VB.Label Label2 
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
                     Left            =   7875
                     TabIndex        =   43
                     Top             =   960
                     Width           =   390
                  End
                  Begin VB.Label Label3 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Select Item By"
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
                     Left            =   0
                     TabIndex        =   42
                     Top             =   960
                     Width           =   1275
                  End
                  Begin VB.Label Label25 
                     AutoSize        =   -1  'True
                     BackColor       =   &H00FF8080&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Date Required"
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
                     Index           =   0
                     Left            =   9165
                     TabIndex        =   41
                     Top             =   960
                     Width           =   1230
                  End
               End
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00FF8080&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   10455
               Begin VB.CommandButton Command1 
                  Caption         =   "Save Item"
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   0
                  Style           =   1  'Graphical
                  TabIndex        =   30
                  Top             =   0
                  Width           =   1215
               End
               Begin VB.CommandButton cmd_delete 
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
                  Height          =   300
                  Left            =   1320
                  Style           =   1  'Graphical
                  TabIndex        =   29
                  Top             =   0
                  Width           =   975
               End
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "INPUT / EDIT HEADER"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   1455
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   10695
            Begin VB.ComboBox cbo_expensetype 
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
               Left            =   4680
               TabIndex        =   67
               Top             =   1080
               Width           =   2895
            End
            Begin VB.TextBox txt_dept 
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
               Left            =   2880
               TabIndex        =   21
               Top             =   480
               Width           =   4695
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Same ""Job No"" For All Items ?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   615
               Left            =   7680
               TabIndex        =   18
               ToolTipText     =   $"purchaserequisition.frx":11552
               Top             =   120
               Width           =   2895
               Begin VB.OptionButton Option2 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "No"
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
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   20
                  Top             =   240
                  Width           =   735
               End
               Begin VB.OptionButton Option1 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Yes"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00A04729&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   19
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Same ""Ship To"" For All Items ?"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   615
               Left            =   7680
               TabIndex        =   15
               ToolTipText     =   $"purchaserequisition.frx":115EB
               Top             =   780
               Width           =   2895
               Begin VB.OptionButton Option3 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Yes"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00A04729&
                  Height          =   255
                  Left            =   600
                  TabIndex        =   17
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   735
               End
               Begin VB.OptionButton Option4 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "No"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00A04729&
                  Height          =   255
                  Left            =   1680
                  TabIndex        =   16
                  Top             =   240
                  Width           =   735
               End
            End
            Begin VB.ComboBox cbo_project 
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
               TabIndex        =   14
               Top             =   1080
               Width           =   4455
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
               TabIndex        =   13
               Top             =   480
               Width           =   1215
            End
            Begin MSComCtl2.DTPicker dtp_pr 
               Height          =   285
               Left            =   1440
               TabIndex        =   22
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
               Format          =   67502081
               CurrentDate     =   38455
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Expense Type"
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
               Left            =   4680
               TabIndex        =   68
               Top             =   840
               Width           =   1200
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Project / Cost Center"
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
               TabIndex        =   26
               Top             =   840
               Width           =   1830
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
               ForeColor       =   &H00A04729&
               Height          =   195
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   720
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "MSR Date"
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
               Index           =   1
               Left            =   1440
               TabIndex        =   24
               Top             =   240
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "Department \ Requestor"
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
               Left            =   2880
               TabIndex        =   23
               Top             =   240
               Width           =   2085
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "VIEW LINE ITEM"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   2895
            Left            =   120
            TabIndex        =   10
            Top             =   4440
            Width           =   10695
            Begin MSFlexGridLib.MSFlexGrid flex_med 
               Height          =   2655
               Left            =   0
               TabIndex        =   11
               Top             =   240
               Width           =   10695
               _ExtentX        =   18865
               _ExtentY        =   4683
               _Version        =   393216
               Cols            =   14
               FixedCols       =   0
               RowHeightMin    =   250
               BackColor       =   16777215
               ForeColor       =   10503977
               BackColorFixed  =   16744576
               ForeColorFixed  =   15593194
               BackColorSel    =   16744576
               BackColorBkg    =   16777215
               SelectionMode   =   1
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
         End
         Begin VB.Label lblid 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   31
            Top             =   1560
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   10920
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7455
         Left            =   -75000
         TabIndex        =   1
         Top             =   300
         Width           =   11655
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7215
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   10695
            Begin VB.ComboBox cbo_recommendedvendor 
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
               TabIndex        =   5
               Top             =   3120
               Width           =   10395
            End
            Begin VB.TextBox txt_justification 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2445
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   360
               Width           =   10455
            End
            Begin VB.TextBox txt_notes 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3375
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   3
               Top             =   3720
               Width           =   10455
            End
            Begin VB.Label Label10 
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
               ForeColor       =   &H00A04729&
               Height          =   195
               Left            =   135
               TabIndex        =   8
               Top             =   2880
               Width           =   1950
            End
            Begin VB.Label Label7 
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
               ForeColor       =   &H00A04729&
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   120
               Width           =   1875
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               BackColor       =   &H00FF8080&
               BackStyle       =   0  'Transparent
               Caption         =   "MSR Remarks "
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
               TabIndex        =   6
               Top             =   3480
               Width           =   1260
            End
         End
         Begin VB.Line Line2 
            X1              =   0
            X2              =   10920
            Y1              =   0
            Y2              =   0
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7455
         Left            =   -75000
         TabIndex        =   59
         Top             =   300
         Width           =   10815
         Begin VB.ListBox listURL 
            Height          =   3375
            ItemData        =   "purchaserequisition.frx":11689
            Left            =   360
            List            =   "purchaserequisition.frx":1168B
            OLEDropMode     =   1  'Manual
            TabIndex        =   65
            Top             =   1320
            Width           =   8295
         End
         Begin VB.TextBox txtsavepath 
            Height          =   375
            Left            =   360
            TabIndex        =   64
            Text            =   "\\cpcomm\savefolder"
            Top             =   4920
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   735
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Width           =   10455
            Begin VB.CommandButton Command4 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Browse"
               Height          =   375
               Left            =   4920
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox txtfilename 
               Height          =   375
               Left            =   120
               TabIndex        =   62
               Top             =   120
               Width           =   4695
            End
            Begin VB.CommandButton Command3 
               BackColor       =   &H00FFC0C0&
               Caption         =   "Attach"
               Height          =   375
               Left            =   6360
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   120
               Width           =   1575
            End
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Note: Double Click on the file to view its contents"
            Height          =   1935
            Left            =   8640
            TabIndex        =   66
            Top             =   1320
            Width           =   2175
         End
      End
      Begin VB.Line Line3 
         X1              =   -75000
         X2              =   -64080
         Y1              =   360
         Y2              =   360
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   0
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "purchaserequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public kl As Integer
Public rad As Integer
Public rad1 As Integer
Public FileName
Public Path
Public Source
Public Destination
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long
Private Sub cbo_category_Change()


If cbo_lookup.Text = "Search" Then
'cbo_category.Clear
Dim mat As New ADODB.Recordset
mat.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml3name like'" & cbo_category.Text & "%' or ml4name like'" & cbo_category.Text & "%' order by ml3name,ml4name", Cn, 3, 2
'While Not mat.EOF
Dim i As Integer
i = 0
For i = 0 To mat.RecordCount - 1
'cbo_category.Clear
If cbo_category.List(i) = mat(2) & "  -  " & mat(3) & "  -  " & mat(0) & "  -  " & mat(1) Then
Else
cbo_category.AddItem mat(2) & "  -  " & mat(3) & "  -  " & mat(0) & "  -  " & mat(1)
End If
mat.MoveNext
Next i
'mat.MoveNext
'Wend
mat.Close
End If

End Sub
Private Sub cbo_category_Click()
On Error Resume Next
If cbo_lookup.Text = "" Then
MsgBox "Select Item selection criteria"
cbo_category.Text = ""
cbo_lookup.SetFocus
Exit Sub
End If
sc = Split(cbo_category.Text, "  -  ", Len(cbo_category.Text), vbTextCompare)
Dim um As New ADODB.Recordset
If um.State Then um.Close

If cbo_lookup.Text = "Item ID" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(2) & "' and ml4name='" & sc(3) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(2) & "' and ml4name='" & sc(3) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Item Description" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Enter Item Manually" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2
ElseIf cbo_lookup.Text = "Search" Then
um.Open "select DISTINCT(ml4uom) from ml4 where ml3name= '" & sc(0) & "' and ml4name='" & sc(1) & "'   order by ml4uom", Cn, 3, 2
End If
If Not um.EOF Then
cbo_uom.Text = um(0)
End If
  
Command1.Enabled = True
End Sub

Private Sub cbo_category_KeyPress(KeyAscii As Integer)
If cbo_lookup.Text = "Item ID" Then
KeyAscii = 0
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
KeyAscii = 0
ElseIf cbo_lookup.Text = "Item Description" Then
KeyAscii = 0
End If


End Sub


Private Sub cbo_expensetype_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_jobcharge_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub cbo_location_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_lookup_Click()

If cbo_materialtype.Text = "" Then
MsgBox "Select Material Type"
cbo_lookup.Text = ""
cbo_materialtype.SetFocus
Exit Sub
End If
Call lookupdetails

End Sub

Private Sub cbo_lookup_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_materialtype_Change()
Call mrenproc
End Sub

Private Sub cbo_materialtype_Click()
If cbo_project.Text = "" Then
MsgBox "Select Project"
cbo_materialtype.Text = ""
cbo_project.SetFocus

Exit Sub
End If
If cbo_expensetype.Text = "" Then
MsgBox "Select Expense Type"
cbo_materialtype.Text = ""
cbo_expensetype.SetFocus

Exit Sub
End If
Call mrenproc
End Sub

Private Sub cbo_materialtype_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_project_Click()
sp8 = Split(cbo_project.Text, "  -  ", Len(cbo_project.Text), vbTextCompare)
Dim jc As New ADODB.Recordset
If jc.State Then jc.Close
jc.Open "select DISTINCT(job_code),job_desc from jobcharge where job_proj_key='" & sp8(0) & "' ", Cn, 3, 2
While Not jc.EOF
cbo_jobcharge.AddItem jc(0) & "  -  " & jc(1)
jc.MoveNext
Wend
jc.Close
End Sub

Private Sub cbo_project_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cbo_uom_Click()
If cbo_category.Text = "" Then
MsgBox "Select Material"
cbo_category.Text = ""
cbo_category.SetFocus
Exit Sub
End If
End Sub

Private Sub cbo_worklocation_Click()
cbo_location.Clear
Dim str1 As New ADODB.Recordset
If str1.State Then str1.Close
str1.Open "select DISTINCT(location) from shipping where worklocation='" & cbo_worklocation.Text & "'  ", Cn, 3, 2
While Not str1.EOF
cbo_location.AddItem str1(0)
str1.MoveNext
Wend
str1.Close
End Sub

Private Sub cmd_delete_Click()
cbo_category.Text = ""
txt_qty.Text = ""
cbo_uom.Text = ""
dtp_reqd.Value = Date
txt_remarks.Text = ""
If rad = 0 Then
cbo_jobcharge.Text = ""
End If
If rad1 = 0 Then
cbo_location.Text = ""
cbo_worklocation.Text = ""
End If
Command1.Enabled = True
cbo_category.Enabled = True
End Sub

Private Sub Command1_Click()
If cbo_category.Text = "" Then
MsgBox "Select Material"
Exit Sub
End If

If txt_qty.Text = "" Then
MsgBox "Enter Required Quantity"
txt_qty.SetFocus
Exit Sub
End If
If cbo_uom.Text = "" Then
MsgBox "Select Unit of Measure"
cbo_uom.SetFocus
Exit Sub
End If

If cbo_jobcharge.Text = "" Then
MsgBox "Select JobCharge"
cbo_jobcharge.SetFocus
Exit Sub
End If

If cbo_worklocation.Text = "" Then
MsgBox "Select Work Location"
cbo_worklocation.SetFocus
Exit Sub
End If

If cbo_location.Text = "" Then
MsgBox "Select Storage Location"
cbo_location.SetFocus
Exit Sub
End If

If dtp_reqd.Value < Date Then
MsgBox "Required Date cannot be less then Todays Date"
Exit Sub
End If
If Not cbo_category.Text = "" Then
    If Not txt_qty.Text = "" Then
    spl = Split(cbo_category.Text, "  -  ", Len(cbo_category.Text), vbTextCompare)
    spl1 = Split(cbo_jobcharge.Text, "  -  ", Len(cbo_jobcharge.Text), vbTextCompare)
Dim jj As Integer
jj = 0

        If kl = 1 Then
                    With flex_med
                        
                        .Rows = .Rows + 1

                        If cbo_lookup.Text = "Item ID" Then
                        .TextMatrix(.Rows - 1, 1) = spl(0)
                        .TextMatrix(.Rows - 1, 2) = spl(1)
                        .TextMatrix(.Rows - 1, 3) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Mfr PartNo." Then
                        .TextMatrix(.Rows - 1, 1) = spl(1)
                        .TextMatrix(.Rows - 1, 2) = spl(0)
                        .TextMatrix(.Rows - 1, 3) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Item Description" Then
                        .TextMatrix(.Rows - 1, 3) = spl(0) & "  -  " & spl(1)
                        .TextMatrix(.Rows - 1, 1) = spl(2)
                        .TextMatrix(.Rows - 1, 2) = spl(3)
                        
                        ElseIf cbo_lookup.Text = "Enter Item Manually" Then
                        .TextMatrix(.Rows - 1, 3) = cbo_category.Text
                        .TextMatrix(.Rows - 1, 1) = ""
                        .TextMatrix(.Rows - 1, 2) = ""
                        
                        ElseIf cbo_lookup.Text = "Search" Then
                        .TextMatrix(.Rows - 1, 3) = spl(0) & "  -  " & spl(1)
                        .TextMatrix(.Rows - 1, 1) = spl(2)
                        .TextMatrix(.Rows - 1, 2) = spl(3)
                        
                        End If
                        
                        .TextMatrix(.Rows - 1, 4) = txt_qty.Text
                        .TextMatrix(.Rows - 1, 5) = cbo_uom.Text
                        .TextMatrix(.Rows - 1, 6) = dtp_reqd.Value
                        .TextMatrix(.Rows - 1, 7) = cbo_jobcharge.Text
                        .TextMatrix(.Rows - 1, 8) = cbo_worklocation.Text
                        .TextMatrix(.Rows - 1, 9) = cbo_location.Text
                        .TextMatrix(.Rows - 1, 10) = txt_remarks.Text
                        .TextMatrix(.Rows - 1, 11) = cbo_materialtype.Text
                        .TextMatrix(.Rows - 1, 12) = Format(dtp_from.Value, "dd/MM/yyyy")
                        .TextMatrix(.Rows - 1, 13) = Format(dtp_to.Value, "dd/MM/yyyy")

                    
                    End With
        Else
                      jj = flex_med.Row

                        If cbo_lookup.Text = "Item ID" Then
                        flex_med.TextMatrix(jj, 1) = spl(0)
                        flex_med.TextMatrix(jj, 2) = spl(1)
                        flex_med.TextMatrix(jj, 3) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Mfr PartNo." Then
                        flex_med.TextMatrix(jj, 1) = spl(1)
                        flex_med.TextMatrix(jj, 2) = spl(0)
                        flex_med.TextMatrix(jj, 3) = spl(2) & "  -  " & spl(3)
                        
                        ElseIf cbo_lookup.Text = "Item Description" Then
                        flex_med.TextMatrix(jj, 1) = spl(2) & "  -  " & spl(3)
                        flex_med.TextMatrix(jj, 2) = spl(0)
                        flex_med.TextMatrix(jj, 3) = spl(1)
                        
                        ElseIf cbo_lookup.Text = "Enter Item Manually" Then
                        flex_med.TextMatrix(jj, 3) = cbo_category.Text
                        flex_med.TextMatrix(jj, 1) = ""
                        flex_med.TextMatrix(jj, 2) = ""
                        
                        ElseIf cbo_lookup.Text = "Search" Then
                        flex_med.TextMatrix(jj, 1) = spl(2) & "  -  " & spl(3)
                        flex_med.TextMatrix(jj, 2) = spl(0)
                        flex_med.TextMatrix(jj, 3) = spl(1)
                        End If
                        
                        flex_med.TextMatrix(jj, 4) = txt_qty.Text
                        flex_med.TextMatrix(jj, 5) = cbo_uom.Text
                        flex_med.TextMatrix(jj, 6) = dtp_reqd.Value
                        flex_med.TextMatrix(jj, 7) = cbo_jobcharge.Text
                        flex_med.TextMatrix(jj, 8) = cbo_worklocation.Text
                        flex_med.TextMatrix(jj, 9) = cbo_location.Text
                        flex_med.TextMatrix(jj, 10) = txt_remarks.Text
                        flex_med.TextMatrix(jj, 11) = cbo_materialtype.Text
                        flex_med.TextMatrix(jj, 12) = Format(dtp_from.Value, "dd/MM/yyyy")
                        flex_med.TextMatrix(jj, 13) = Format(dtp_to.Value, "dd/MM/yyyy")
        End If
 kl = 1
cbo_category.Text = ""
txt_qty.Text = ""
cbo_uom.Text = ""
dtp_reqd.Value = Date
txt_remarks.Text = ""
If rad = 0 Then
cbo_jobcharge.Text = ""
End If
If rad1 = 0 Then
cbo_worklocation.Text = ""
cbo_location.Text = ""
End If
End If 'qty
 End If 'category
 Command1.Enabled = False
End Sub

Private Sub Command3_Click()
 Dim fln As New ADODB.Recordset
Dim k As Integer
k = 0
For k = 0 To listURL.ListCount - 1
FileName = ""
sp = Split(listURL.List(k), "\", Len(listURL.List(k)), vbTextCompare)
Dim i As Integer
i = 0
For i = 0 To 10
If sp(i) <> "" Then
    spp = Split(sp(i), ".", Len(sp(i)), vbTextCompare)
           If sp(1) = "savefolder" Then GoTo goout
      Dim b As String
      Dim a As String
      a = "": b = ""
      On Error Resume Next
      If spp(1) <> "" Then a = spp(0): b = spp(1)
      FileName = a & "." & b
     
End If
Next
If InStr(1, FileName, "TL-MSR", vbTextCompare) Then GoTo goout
 Dim fpt As String
 fpt = ""
 fpt = txtsavepath & "\" & txt_account.Text & "--" & FileName
        FileCopy FileName, txtsavepath & "\" & txt_account.Text & "--" & FileName
        fn = Split(FileName, "--", Len(FileName), vbTextCompare)
        fnm = ""
        fnm = txt_account.Text & "--" & FileName
       
        If fln.State Then fln.Close
        fln.Open "select * from fileattach where fname='" & fnm & "'", Cn, 3, 2
        If fln.EOF Then
        
        
    Dim fa As New ADODB.Recordset
        fa.Open "select * from fileattach", Cn, 3, 2
        fa.AddNew
        fa!fprno = txt_account.Text
        fa!fpath = txtsavepath & "\" & txt_account.Text & "--" & FileName
        If X = 1 Then
        fa!fname = FileName
        Else
        fa!fname = txt_account.Text & "--" & FileName
        End If
        fa.Update
        fa.Close
        
        End If
        fln.Close
    
goout:
''''

Next k
MsgBox "Attachment Process Completed"
listURL.Clear
End Sub

Private Sub Command4_Click()
'On Error Resume Next
Dim objfso As New FileSystemObject
Dim sSel As String, sExt As String

cdOpen.ShowOpen

If Not vbCancel Then
txtfilename = cdOpen.FileName
End If
listURL.AddItem txtfilename.Text

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

If flex_med.Row <> 0 Then
'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1
End If

vprev = flex_med.Row


Dim msrapp As New ADODB.Recordset
If msrapp.State Then msrapp.Close
msrapp.Open "select * from purchaserequisition where status='Approved' and prno='" & txt_account.Text & "'", Cn, 3, 2
If Not msrapp.EOF Then
Command1.Enabled = False
End If

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
If flex_med.Row <> 0 Then
'Current  row
flex_med.Row = current
For i = 1 To flex_med.Cols - 1
flex_med.Col = i
flex_med.CellBackColor = RGB(202, 204, 221) 'vbyellow
Next
flex_med.Col = 1
End If
Dim idd As Double
idd = 0
 
cbo_category.Text = flex_med.TextMatrix(flex_med.Row, 1) & "  -  " & flex_med.TextMatrix(flex_med.Row, 2) & "  -  " & flex_med.TextMatrix(flex_med.Row, 3)
txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 4)
cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 5)
dtp_reqd.Value = flex_med.TextMatrix(flex_med.Row, 6)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 10)
cbo_jobcharge.Text = flex_med.TextMatrix(flex_med.Row, 7)
cbo_worklocation.Text = flex_med.TextMatrix(flex_med.Row, 8)
cbo_location.Text = flex_med.TextMatrix(flex_med.Row, 9)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 10)
cbo_materialtype.Text = flex_med.TextMatrix(flex_med.Row, 11)
dtp_from.Value = flex_med.TextMatrix(flex_med.Row, 12)
dtp_to.Value = flex_med.TextMatrix(flex_med.Row, 13)

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
 

cbo_category.Text = flex_med.TextMatrix(flex_med.Row, 1) & "  -  " & flex_med.TextMatrix(flex_med.Row, 2) & "  -  " & flex_med.TextMatrix(flex_med.Row, 3)
txt_qty.Text = flex_med.TextMatrix(flex_med.Row, 4)
cbo_uom.Text = flex_med.TextMatrix(flex_med.Row, 5)
dtp_reqd.Value = flex_med.TextMatrix(flex_med.Row, 6)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 10)
cbo_jobcharge.Text = flex_med.TextMatrix(flex_med.Row, 7)
cbo_worklocation.Text = flex_med.TextMatrix(flex_med.Row, 8)
cbo_location.Text = flex_med.TextMatrix(flex_med.Row, 9)
txt_remarks.Text = flex_med.TextMatrix(flex_med.Row, 10)
cbo_materialtype.Text = flex_med.TextMatrix(flex_med.Row, 11)
dtp_from.Value = flex_med.TextMatrix(flex_med.Row, 12)
dtp_to.Value = flex_med.TextMatrix(flex_med.Row, 13)

vprev = flex_med.Row

End Sub

Private Sub Form_Load()
On Error Resume Next
'cbo_subcategory.Enabled = True
'cbo_material.Enabled = True
Command1.Enabled = True
frame_ms.Visible = False
rad = 0
rad1 = 0
dtp_pr.Value = Format(Date, "dd/MM/yyyy")
dtp_reqd.Value = Format(Date, "dd/MM/yyyy")

dtp_from.Value = Format(Date, "dd/MM/yyyy")
dtp_to.Value = Format(Date, "dd/MM/yyyy")


cbo_expensetype.AddItem "Project Expenses"
cbo_expensetype.AddItem "Capital Expenses"


Dim mb As New ADODB.Recordset
If mb.State Then mb.Close
mb.Open "select * from purchaserequisition", Cn, 3, 2
Dim i As Integer
 

   i = 1
assad:
Dim X As String
X = "TL-MSR000" & i
   Dim mbb As New ADODB.Recordset
   If mbb.State Then mbb.Close
   mbb.Open "select * from purchaserequisition where prno='" & X & "' ", Cn, 3, 2
   If mbb.EOF Then
   txt_account.Text = "TL-MSR000" & i
   Else
   i = i + 1
 
GoTo assad
  End If

cbo_recommendedvendor.Clear
Dim vn As New ADODB.Recordset
If vn.State Then vn.Close
vn.Open "select DISTINCT(regno),name from vendor order by name", Cn, 3, 2
While Not vn.EOF
cbo_recommendedvendor.AddItem vn(1)
vn.MoveNext
Wend
vn.Close
cbo_worklocation.Clear

vn.Open "select DISTINCT(workloc) from worklocation", Cn, 3, 2
While Not vn.EOF
cbo_worklocation.AddItem vn(0)
vn.MoveNext
Wend
vn.Close

 cbo_location.Clear
 Dim sh As New ADODB.Recordset
 If sh.State Then sh.Close
 sh.Open "select DISTINCT(location) from shipping where worklocation='" & cbo_worklocation.Text & "'  order by location", Cn, 3, 2
 While Not sh.EOF
 cbo_location.AddItem sh(0)
 sh.MoveNext
 Wend
 sh.Close

cbo_project.Clear
Dim prj As New ADODB.Recordset
If prj.State Then prj.Close
prj.Open "select DISTINCT(proj_key),proj_title from projectmaster  ", Cn, 3, 2
While Not prj.EOF
cbo_project.AddItem prj(0) & "  -  " & prj(1)
prj.MoveNext
Wend
prj.Close


prj.Open "select DISTINCT(mtcode),mtdesc from materialtype order by mtcode", Cn, 3, 2
While Not prj.EOF
cbo_materialtype.AddItem prj(0) & "  -  " & prj(1)
prj.MoveNext
Wend
prj.Close

 If Option1.Value = True Then
 rad = 1
 ElseIf Option2.Value = True Then
 rad = 0
 End If
 If Option3.Value = True Then
 rad1 = 1
 ElseIf Option4.Value = True Then
 rad1 = 0
 End If
  
  
    Call flex_titlepr
    kl = 1
    
 sh.Open "select DISTINCT(a_department) from userid where a_name='" & main.Label2.Caption & "' ", Cn, 3, 2
 If Not sh.EOF Then
 txt_dept.Text = sh(0) & "\" & main.Label2.Caption
 End If
 
 cbo_lookup.AddItem "Item ID"
 cbo_lookup.AddItem "Mfr PartNo."
 cbo_lookup.AddItem "Item Description"
 cbo_lookup.AddItem "Enter Item Manually"
 cbo_lookup.AddItem "Search"
 
End Sub
Public Sub flex_titlepr()
On Error Resume Next
flex_med.Rows = 1
   With flex_med
        .Row = 0:    .Col = 0
        .ColWidth(0) = 0

        .TextMatrix(0, 1) = "ItemId"
        .ColWidth(1) = 1700
        .ColAlignment(1) = 0
        .TextMatrix(0, 2) = "Mfr.Ref Code"
        .ColWidth(2) = 1700
        .ColAlignment(2) = 0
        .TextMatrix(0, 3) = "Material"
        .ColWidth(3) = 3000
        .ColAlignment(3) = 0
        .TextMatrix(0, 4) = "Qty"
        .ColWidth(4) = 800
        .ColAlignment(4) = 0
        .TextMatrix(0, 5) = "UOM"
        .ColWidth(5) = 800
        .ColAlignment(5) = 0
        .TextMatrix(0, 6) = "ReqDate"
        .ColWidth(6) = 1200
        .ColAlignment(6) = 0
        
        .TextMatrix(0, 10) = "Remarks"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        
        .TextMatrix(0, 7) = "JobCharge"
        .ColWidth(7) = 1200
        .ColAlignment(7) = 0
        .TextMatrix(0, 8) = "Work Location"
        .ColWidth(8) = 1200
        .ColAlignment(8) = 0
        .TextMatrix(0, 9) = "Stor Loc"
        .ColWidth(9) = 1200
        .ColAlignment(9) = 0
        .TextMatrix(0, 10) = "Line Item Text"
        .ColWidth(10) = 1200
        .ColAlignment(10) = 0
        .TextMatrix(0, 11) = "Mat.Type"
        .ColWidth(11) = 1200
        .ColAlignment(11) = 0
        .TextMatrix(0, 12) = "From"
        .ColWidth(12) = 1200
        .ColAlignment(12) = 0
        .TextMatrix(0, 13) = "To"
        .ColWidth(13) = 1200
        .ColAlignment(13) = 0
        
    End With
End Sub

Private Sub listURL_Click()
StartDoc listURL.Text
End Sub

Private Sub Option1_Click()
rad = 1
End Sub
Private Sub Option2_Click()
rad = 0
cbo_jobcharge.Text = ""
End Sub
Private Sub Option3_Click()
rad1 = 1
End Sub
Private Sub Option4_Click()
rad1 = 0
cbo_location.Text = ""
cbo_contactperson.Text = ""
End Sub
Public Sub lookupdetails()
cbo_category.Clear
mty = Split(cbo_materialtype.Text, "  -  ", Len(cbo_materialtype.Text), vbTextCompare)
 Dim med As New ADODB.Recordset
If med.State Then med.Close

 If cbo_lookup.Text = "Item ID" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(0) & "  -  " & med(1) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close
ElseIf cbo_lookup.Text = "Mfr PartNo." Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(1) & "  -  " & med(0) & "  -  " & med(2) & "  -  " & med(3)
med.MoveNext
Wend
med.Close


ElseIf cbo_lookup.Text = "Item Description" Then
cbo_category.Clear
med.Open "select DISTINCT(ml4itemid),ml4mrefcode,ml3name,ml4name from ml4 where ml4type= '" & mty(0) & "' order by ml4name", Cn, 3, 2
While Not med.EOF
cbo_category.AddItem med(2) & "  -  " & med(3) & "  -  " & med(0) & "  -  " & med(1)
med.MoveNext
Wend
med.Close

ElseIf cbo_lookup.Text = "Enter Item Manually" Then


ElseIf cbo_lookup.Text = "Search" Then
cbo_category.Clear
End If
End Sub


Public Sub mrenproc()
On Error Resume Next
spmt = Split(cbo_materialtype.Text, "  -  ", Len(cbo_materialtype.Text), vbTextCompare)
Dim mren As New ADODB.Recordset
If mren.State Then mren.Close
mren.Open "select DISTINCT(rental) from materialtype where mtcode='" & spmt(0) & "'", Cn, 3, 2
If Not mren.EOF Then
If mren(0) = "Yes" Then
frame_ms.Visible = True
Else
frame_ms.Visible = False
End If
End If
End Sub

