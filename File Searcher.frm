VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XTREME File Searcher V2.0"
   ClientHeight    =   7935
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9705
   Icon            =   "File Searcher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   529
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSystem 
      Caption         =   "Only &System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkReadOnly 
      Caption         =   "Only &Read Only"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Only &Archive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   14
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox chkFiles 
      Caption         =   "Only &Files"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.CheckBox chkFolders 
      Caption         =   "Only F&olders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Only &Hidden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpOne 
      Caption         =   "&Up 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7800
      TabIndex        =   10
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   270
      Left            =   8640
      TabIndex        =   9
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   270
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   20
      Left            =   4560
      Top             =   0
   End
   Begin VB.ListBox lstResult 
      Height          =   5325
      ItemData        =   "File Searcher.frx":000C
      Left            =   120
      List            =   "File Searcher.frx":000E
      TabIndex        =   4
      Top             =   1200
      Width           =   9495
   End
   Begin VB.TextBox txtFilter 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   ".vbp"
      Top             =   480
      Width           =   6735
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label lblCurpath 
      Caption         =   "Current path"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   7320
      Width           =   9495
   End
   Begin VB.Label lblFilesFound 
      Caption         =   "Files Found"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6960
      Width           =   9495
   End
   Begin VB.Label lblFilesSearched 
      Caption         =   "Files Searched"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6600
      Width           =   9495
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBrowse 
         Caption         =   "&Browse"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub mnuFileBrowse_Click()
    If Right$(lstResult.Text, 1) = "\" Then
        StartDoc lstResult.Text
    Else
        StartDoc UpOne(lstResult.Text)
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Call FileSearch(lstResult, txtDir, txtFilter, , , CBool(chkFiles), CBool(chkFolders), _
    CBool(chkReadOnly), CBool(chkArchive), CBool(chkHidden), CBool(chkSystem))
End Sub

Private Sub cmdStop_Click()
    Abort = True
End Sub

Private Sub cmdUpOne_Click()
    txtDir = UpOne(txtDir)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lstResult_DblClick()
    StartDoc lstResult.Text
End Sub

Private Sub lstResult_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        If Right$(lstResult.Text, 1) = "\" Then
            If DeleteFolder(lstResult.Text) Then
                lstResult.RemoveItem lstResult.ListIndex
            Else
                MsgBox "Error deleting folder"
            End If
        Else
            If DeleteFile(lstResult.Text) Then
                lstResult.RemoveItem lstResult.ListIndex
            Else
                MsgBox "Error deleting file"
            End If
        End If
    End Select
End Sub

Private Sub mnuFileProperties_Click()
    If lstResult.Text = vbNullString Then
        MsgBox "Please select an item", vbOKOnly, "Error"
        Exit Sub
    End If
    Load frmProperties
    frmProperties.SetFileName (lstResult.Text)
End Sub

Private Sub tmrUpdate_Timer()
    lblFilesFound = "Files Found: " & FilesFound
    lblFilesSearched = "Total Files Searched: " & FileSearchCount
    lblCurpath = CurrentName
End Sub
