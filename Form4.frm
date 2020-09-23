VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Database"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1320
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "To Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Subject:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Mail To:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "FrmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single

Private Sub Command1_Click()
Dim sDNS As String
Dim sMailServer As String
Dim sDomain As String

sDomain = GetDomainFromAddr(Text3.Text)
If sDomain = "" Then
MsgBox "Please Enter a VALID address!", vbCritical, "Error"
Exit Sub
End If

sMailServer = MX1.GetMX

If sMailServer = "" Then
MsgBox "Sorry could not locate the mail server for this address.", vbInformation, "Opps..."
Exit Sub
End If

If MX1.DNSCount = 0 Then
MsgBox "Could not Get Local DNS!", vbCritical, "Error"
Exit Sub
End If

sDNS = MX1.DNS(0)

'Error checking for this demo only
If sDNS = "" Then
MsgBox "Could not Retrive Local DNS Server!" & vbCrLf & "Please check your internet settings as EVERYONE has a DNS!", vbCritical, "Opps..."
Exit Sub
End If
'''''''''''''''''''''''''''''''''''

Label3.Caption = "Using Server = " & sMailServer
SendEmail sMailServer, "billing@graydove.com", "Billing", Text6.Text, Text3.Text, Text1.Text, Text2.Text
MsgBox "Your mail has been sent.", vbInformation, "Send Mail"
Label3.Caption = "Done"
End Sub

Private Function GetMX(sServer As String, sDNS As String) As String
With wsMX
.RemoteHost = sDNS
.RemotePort = 53 'mx lookup port
.connect
End With
End Function

Public Function GetDomainFromAddr(sAddr As String) As String
Dim Ipos As Long
Ipos = InStr(1, sAddr, "@", vbBinaryCompare)
If Ipos > 0 Then
GetDomainFromAddr = Mid(sAddr, Ipos + 1, Len(sAddr))
Exit Function
End If
GetDomainFromAddr = ""
End Function

Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.connect ' Start connection
    
    WaitFor ("220")
    
    Label3.Caption = "Connecting...."
    Label3.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    Label3.Caption = "Connected"
    Label3.Refresh

    Winsock1.SendData (first)

    Label3.Caption = "Sending Message"
    Label3.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    Label3.Caption = "Disconnecting"
    Label3.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + response, 64, MsgTitle
            Exit Sub
        End If
    Wend
response = "" ' Sent response code to blank **IMPORTANT**
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData response ' Check for incoming response *IMPORTANT*
End Sub

