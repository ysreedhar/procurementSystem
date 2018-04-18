VERSION 5.00
Begin VB.Form mailform 
   Caption         =   "mailform"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton mailform 
      Caption         =   "mailform"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "mailform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendMail_Click
End Sub
Private Sub SendMail_Click()

    ' Create a MailMan instance to send email
    Dim mailman As New ChilkatMailMan2
    ' Unlock the email component
    mailman.UnlockComponent "UnlockCode"
    
    ' Change this to your SMTP server.
    ' Depending on your SMTP server, you may not need
    ' the SMTP login or SMTP password.
    mailman.SMTPHost = "cpxeon.crest.com.my"
    mailman.SmtpUsername = "shaik"
    mailman.SmtpPassword = "shaik05"
    
    ' Create an email for sending
    Dim email As New ChilkatEmail2
    ' Set the email subject and body
    email.Subject = "test"
    email.Body = "this is a test" & vbCrLf & "line 2" & vbCrLf & "-Bill"
    email.AddTo "Test", "assad@crest.com.my"
    email.FromAddress = "assad@crest.com.my"
    email.FromName = "Assad"
    
    ' Add a file attachment to the mail.
    'email.AddFileAttachment "dude.gif"
        
    ' Send mail.  Returns 1 if successful, 0 if failed.
    success = mailman.SendEmail(email)
    If (success = 0) Then
        MsgBox mailman.LastErrorText
    End If
    

End Sub

Private Sub mailform_Click()
SendMail_Click
End Sub
