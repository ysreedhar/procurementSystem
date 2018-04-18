VERSION 5.00
Begin VB.Form mail 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Send Mail"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   10545
   Begin VB.CommandButton Cmd_send 
      Caption         =   "Send"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_send_Click()
Set poSendMail = New vbSendMail.clsSendMail
poSendMail.SMTPHost = "mail.bits.com.my" 'txtServer.Text
poSendMail.From = "assad@bits.com.my" 'txtFrom.Text
poSendMail.FromDisplayName = "ASSAD" 'txtFromName.Text
poSendMail.Recipient = "assad@bits.com.my" 'txtTo.Text
poSendMail.RecipientDisplayName = "Test" 'txtToName.Text
poSendMail.ReplyToAddress = "assad@bits.com.my" 'txtFrom.Text
poSendMail.Subject = "Hi" 'txtSubject.Text
'poSendMail.Attachment = txtFileName.Text 'attached file name
poSendMail.Message = "How are you" 'txtMsg.Text
poSendMail.Send
Set poSendMail = Nothing

End Sub

Private Sub Form_Load()
 Set poSendMail = New clsSendMail
End Sub
