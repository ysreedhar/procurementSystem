VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   12000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frm_login.frx":0000
   ScaleHeight     =   12000
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_cancel 
      Height          =   495
      Left            =   11880
      Picture         =   "frm_login.frx":1B36C
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Click to Clear"
      Top             =   4920
      Width           =   495
   End
   Begin VB.CommandButton cmd_submit 
      Height          =   495
      Left            =   11280
      Picture         =   "frm_login.frx":1B92A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Click to Login"
      Top             =   4920
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   2880
   End
   Begin VB.Frame PositionFrame 
      Caption         =   "Position"
      Enabled         =   0   'False
      Height          =   720
      Left            =   720
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   4170
      Begin VB.TextBox CharPosn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   375
         TabIndex        =   24
         Top             =   255
         Width           =   570
      End
      Begin VB.TextBox CharPosn 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1845
         TabIndex        =   23
         Top             =   255
         Width           =   570
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&X:"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   300
         Width           =   270
      End
      Begin VB.Label CharPosnLabel 
         Caption         =   "&Y:"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   1620
         TabIndex        =   25
         Top             =   300
         Width           =   270
      End
   End
   Begin VB.Frame SpeechOutputFrame 
      Caption         =   "Speech &Output"
      Enabled         =   0   'False
      Height          =   2085
      Left            =   720
      TabIndex        =   16
      Top             =   3750
      Visible         =   0   'False
      Width           =   4170
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "A&uto hide"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   420
         TabIndex        =   21
         Top             =   1650
         Width           =   1200
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Auto &pace"
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1605
         TabIndex        =   20
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Si&ze to text"
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2895
         TabIndex        =   19
         Top             =   1665
         Width           =   1095
      End
      Begin VB.CheckBox BalloonStyleOption 
         Caption         =   "Display &word balloon"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1290
         Width           =   1935
      End
      Begin VB.TextBox SpeakText 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   930
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   255
         Width           =   3900
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   360
      Index           =   1
      Left            =   4995
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Play"
      Enabled         =   0   'False
      Height          =   360
      Index           =   0
      Left            =   4980
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Speak"
      Enabled         =   0   'False
      Height          =   360
      Index           =   2
      Left            =   4995
      TabIndex        =   13
      Top             =   3975
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Move"
      Enabled         =   0   'False
      Height          =   360
      Index           =   3
      Left            =   4995
      TabIndex        =   12
      Top             =   6075
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame AnimationFrame 
      Caption         =   "&Animations for"
      Enabled         =   0   'False
      Height          =   2355
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   4155
      Begin VB.ListBox AnimationListBox 
         Enabled         =   0   'False
         Height          =   1620
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   3900
      End
      Begin VB.CheckBox OutputStyleOption 
         Caption         =   "Play sound &effects"
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox OutputStyleOption 
         Caption         =   "Stop &before next action"
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1995
         TabIndex        =   9
         Top             =   1950
         Value           =   1  'Checked
         Width           =   1995
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Text            =   "GestureDown"
      Top             =   345
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txt_password 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00A04729&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   8520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5040
      Width           =   2415
   End
   Begin VB.TextBox cbo_userid 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00A04729&
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   8520
      TabIndex        =   1
      Top             =   4320
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4260
      Top             =   3105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frm_login.frx":1BF5C
      Top             =   10680
      Width           =   480
   End
   Begin AgentObjectsCtl.Agent MyAgent 
      Left            =   4860
      Top             =   3105
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   8280
      Picture         =   "frm_login.frx":1C55B
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Turn Off PMS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      ToolTipText     =   "Turn Off PCIS"
      Top             =   10733
      Width           =   2175
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCUREMENT MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   3960
      TabIndex        =   5
      Top             =   600
      Width           =   7500
   End
   Begin VB.Label l1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TL OFFSHORE SDN BHD"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   9000
      TabIndex        =   0
      Top             =   3360
      Width           =   4395
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Character As IAgentCtlCharacterEx
Dim NewBalloonStyleOption As Integer
Dim CharLoaded As Boolean
Dim IgnoreSizeEvent As Boolean
Dim CurrentIndex As Integer

Const BalloonOn = 1
Const SizeToText = 2
Const AutoHide = 4
Const AutoPace = 8
Public u As Integer
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Sub SetBalloonStyleOptions()
'------------------------------------------------
'-- This subroutine sets the check boxes for the
'-- the word balloon settings
'------------------------------------------------

'-- Check to see if the balloon is on

If Character.Balloon.Style And BalloonOn Then
    BalloonStyleOption(0).Value = 1
Else
    BalloonStyleOption(0).Value = 0
End If

'-- Check to see if Auto-Hide is on

If Character.Balloon.Style And AutoHide Then
    BalloonStyleOption(1).Value = 1
Else
    BalloonStyleOption(1).Value = 0
End If

'-- Check to see if Auto-Pace is on

If Character.Balloon.Style And AutoPace Then
    BalloonStyleOption(2).Value = 1
Else
    BalloonStyleOption(2).Value = 0
End If

'-- Check to see if Size-To-Text is on

If Character.Balloon.Style And SizeToText Then
    BalloonStyleOption(3).Value = 1
Else
    BalloonStyleOption(3).Value = 0
End If


'-- Set the controls based on Advanced Character Options

If Not Character.Balloon.Enabled Then
    BalloonStyleOption(0).Enabled = False
    BalloonStyleOption(1).Enabled = False
    BalloonStyleOption(2).Enabled = False
    BalloonStyleOption(2).Enabled = False
Else
    BalloonStyleOption(0).Enabled = True
    BalloonStyleOption(1).Enabled = True
    BalloonStyleOption(2).Enabled = True
    BalloonStyleOption(2).Enabled = True
End If

End Sub
Function GetWindowsDir() As String
    Dim Temp As String
    Dim Ret As Long
    Const MAX_LENGTH = 145

    Temp = String$(MAX_LENGTH, 0)
    Ret = GetWindowsDirectory(Temp, MAX_LENGTH)
    Temp = Left$(Temp, Ret)
    If Temp <> "" And Right$(Temp, 1) <> "\" Then
        GetWindowsDir = Temp & "\"
    Else
        GetWindowsDir = Temp
    End If
End Function


Private Sub AnimationListBox_Click()

'-- Enable the Play button
Command1(0).Enabled = True

End Sub

Sub EnableControls()
'-----------------------------------
'-- Enable the controls on the page
'-----------------------------------

'-- Enable the Animation List Box
AnimationFrame.Enabled = True
AnimationListBox.Enabled = True


'-- Enable the Stop and Move buttons
Command1(1).Enabled = True
Command1(3).Enabled = True

'-- Enable the Play Sound Effects option only
'-- if enabled in the Advanced Character Options

If MyAgent.AudioOutput.Enabled And MyAgent.AudioOutput.SoundEffects Then
    OutputStyleOption(0).Enabled = True
End If


'-- Enable the Stop Before Play option

OutputStyleOption(1).Enabled = True


'-- Enable the Balloon Style options
BalloonStyleOption(0).Enabled = True
BalloonStyleOption(1).Enabled = True
BalloonStyleOption(2).Enabled = True
BalloonStyleOption(3).Enabled = True


'-- Enable the Speech Text Box
SpeechOutputFrame.Enabled = True
SpeakText.Enabled = True
SpeakText.BackColor = vbWindowBackground


'-- Enable the X,Y position fields
PositionFrame.Enabled = True

CharPosnLabel(0).Enabled = True
CharPosn(0).Enabled = True
CharPosn(0).BackColor = vbWindowBackground

CharPosnLabel(1).Enabled = True
CharPosn(1).Enabled = True
CharPosn(1).BackColor = vbWindowBackground

End Sub

Private Sub cbo_userid_LostFocus()
''On Error Resume Next
''u = 1
'''--------------------------
''If u = 1 Then
''dp = Split(Time, " ", Len(Time), vbTextCompare)
'''-- If Stop Before Play is set, stop the character
'''-- before the next request
''X = dp(1)
'''If (Time) > ("5:00:0 AM") And (Time) < ("11:59:0 AM") Then
''' strg = "Good Morning"
''' ElseIf (Time) >= ("12:00:0 PM") And (Time) < ("16:00:0 PM") Then
''' strg = "Good Afternoon"
''' ElseIf (Time) > ("4:00:0 PM") And (Time) < ("7:00:0 PM") Then
''' strg = "Good Evening"
''' ElseIf (Time) > ("7:00:0 PM") And (Time) < ("4:59:0 AM") Then
''' strg = "Good Night"
''' Else
''' strg = "Welcome"
''' End If
''If X = "AM" Then
''strg = "Good Morning"
''ElseIf X = "PM" Then
''strg = "Good Afternoon"
''Else
''strg = "Welcome"
''End If
''
''Character.Speak strg & "       " & " " & cbo_userid.Text & "         " & " ,Your " & "Login Time:," & " " & Format(Time, "HH:MM:SS")
''
'''-------------------------
''End If
End Sub

Private Sub cmd_cancel_Click()
cbo_userid.SetFocus
cbo_userid.Text = ""
txt_password.Text = ""
End Sub

Private Sub cmd_close_Click()
 
End Sub

Private Sub cmd_submit_Click()
On Error Resume Next
Dim pwd As New ADODB.Recordset
If pwd.State Then pwd.Close
pwd.Open "select * from userid where a_userid='" & cbo_userid.Text & "' and a_password='" & txt_password.Text & "' ", Cn, 3, 2
If Not pwd.EOF Then
main.Label2.Caption = cbo_userid.Text
main.Label1.Caption = "User:" & " " & pwd("a_userid") & "  " & "Login Time:" & " " & Format(Time, "HH:MM:SS")
 


main.Enabled = True
main.Show
Dim rs As New ADODB.Recordset
If rs.State Then rs.Close
rs.Open "select * from login", Cn, 3, 2
rs.AddNew
rs!l_userid = cbo_userid.Text
rs!l_intime = Now
rs.Update
'-----------------------------------
 Set Character = Nothing
 
Unload frm_login
'-------------------------


End If

End Sub

Private Sub Form_Load()
On Error Resume Next

'''''--------------------
'''''----------------------------------------------------------
'''''-- When the form loads, set the IgnoreSizeEvent flag
'''''-- (used to differentiate when the Character Animation
'''''-- Previewer window is restored), set the CharLoaded flag
'''''-- (used to track when a character is loaded),
'''''-- and set the initial state of the status bar.
'''''----------------------------------------------------------
''''IgnoreSizeEvent = True
''''
''''CharLoaded = False
'''''-- Set a flag to track success
''''OpenSuccess = False
''''''
''''CommonDialog1.CancelError = True
'''''
'''''On Error GoTo ErrHandler
'''''
''''CommonDialog1.Flags = cdlOFNHideReadOnly
'''''
''''''-- Get the Windows directory name
''''Dim DirName As String
''''DirName = GetWindowsDir()
'''''
''''''-- Append the Agent Chars subdirectory
''''CommonDialog1.InitDir = DirName + "msagent\chars"
''''''
'''''''-- Add the filter
''''CommonDialog1.Filter = "Microsoft Agent Characters (*.acs)|*.acs"
''''CommonDialog1.FilterIndex = 1
'''''
''''''-- Show the Open dialog
'''''CommonDialog1.ShowOpen
'''''
'''''--Unload the previous character
''''On Error Resume Next
''''Set Character = Nothing
''''MyAgent.Characters.Unload "C:\WINDOWS\msagent\chars\merlin.acs"
''''
'''''-- Load the new character
''''On Error GoTo errhandler
''''MyAgent.Characters.Load "C:\WINDOWS\msagent\chars\merlin.acs"
''''
''''OpenSuccess = True
''''
''''Set Character = MyAgent.Characters("C:\WINDOWS\msagent\chars\merlin.acs")
''''
''''frm_loginCaption = "merlin" + " - Microsoft Character Animation Previewer"
''''
''''
'''''-- Set the character loaded flag
''''CharLoaded = True
''''
'''''-- Set the character's language
''''Character.LanguageID = &H409
''''
'''''-- Update the caption for the animation list box
''''AnimationFrame.Caption = "&Animations for " + "merlin"
''''
'''''-- Disable the Play button to avoid trying to play a null animation selection
''''Command1(0).Enabled = False
''''
'''''-- Load the character's animation into the list box
''''AnimationListBox.Clear
''''For Each AnimationName In Character.AnimationNames
''''        AnimationListBox.AddItem AnimationName
''''Next
''''
'''''-- Move the character to starting position
''''Character.Left = (frm_loginLeft + 7050) / Screen.TwipsPerPixelX
''''Character.Top = (frm_loginTop + 4000) / Screen.TwipsPerPixelY
''''
'''''-- Show the character
''''Character.Show
'''' Character.Play Text1.Text
''''
'''''-- Update the X,Y position fields with the character's
'''''-- current position
''''CharPosn(0).Text = CStr(Character.Left)
''''CharPosn(1).Text = CStr(Character.Top)
''''
'''''-- Update the state of the balloon style options
''''SetBalloonStyleOptions
''''
'''''-- Initialize the pop-up menu commands
'''''InitPopupMenuCmds
''''
'''''-- Update the state of the controls to match the
'''''-- character's settings
''''EnableControls
''''
''''AnimationListBox.SetFocus
''''
''''Exit Sub
''''
''''errhandler:
''''    If (Err.Number <> cdlCancel) Then
''''        If (OpenSuccess = False) Then
''''            MsgBox "There was an error opening the file " & CommonDialog1.FileName
''''
''''        End If
''''
'''''        Set Character = Nothing
''''
''''    End If
'''''-------------------
''''

'lbloff.Visible = False
Call connect
Me.Top = 0
Me.Left = 0
Me.Width = 16000
Me.Height = 16000
' Picture1.Visible = False
' Picture1.Enabled = False
 
 
 l1.Visible = False
 l2.Visible = False
 
 txt_password.Visible = False
 txt_password.Enabled = False
 cbo_userid.Visible = False
 cbo_userid.Enabled = False
 cmd_cancel.Visible = False
 cmd_cancel.Enabled = False
 cmd_submit.Enabled = False
 cmd_submit.Visible = False
 
'main.Enabled = False
'''Dim lg As New ADODB.Recordset
'''If lg.State Then lg.Close
'''lg.Open "select DISTINCT(a_userid) from userid order by a_userid", Cn, 3, 2
'''While Not lg.EOF
'''cbo_userid.AddItem lg(0)
'''lg.MoveNext
'''Wend
'''lg.Close

End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
Label6.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub Image2_Click()
'----------------------
'-- Move the character to starting position
 
'------------------------
l1.Visible = True
l2.Visible = True

txt_password.Visible = True
txt_password.Enabled = True
cbo_userid.Visible = True
cbo_userid.Enabled = True
cmd_cancel.Visible = True
cmd_cancel.Enabled = True
cmd_submit.Enabled = True
cmd_submit.Visible = True
cbo_userid.SetFocus
'Load frmagent
'frm_loginShow
'frm_loginHide
u = 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.ToolTipText = "Click to Enter User Name"
Image2.MousePointer = 14
End Sub

Private Sub Label6_Click()
 Unload Me
End Sub

 
Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = RGB(100, 250, 20)
Image1.BorderStyle = 0
End Sub


