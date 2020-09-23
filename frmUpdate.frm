VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check for Update"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   615
      Left            =   1920
      TabIndex        =   12
      Top             =   6600
      Width           =   1575
      ExtentX         =   2778
      ExtentY         =   1085
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   4815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   6000
      Width           =   4815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5106
      _Version        =   393217
      ReadOnly        =   -1  'True
      TextRTF         =   $"frmUpdate.frx":0CCA
   End
   Begin VB.Label Label5 
      Caption         =   "Current Version"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "File Size"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   4815
   End
   Begin VB.Label Label3 
      Caption         =   "Update URL"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Programmer"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   1815
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

If Command1.Caption = "Exit" Then
Open App.Path & "\accountdb.log" For Append As #1
Print #1, ""
Unload Me
Exit Sub
End If


On Error Resume Next

Close #1
Close #2
Close #3

RichTextBox1.SelColor = vbRed
RichTextBox1.SelText = "Checking For New Updates..." & vbCrLf

Data = GetUrlSource("http://www.geocities.com/hax03d/updates.txt")

If Data = "" Then
MsgBox "Error Server Down, Please check back at a later time, Sorry for the Inconvience", vbCritical, "Server Down"
Print #1, Time & " - Update failed.  Server was unable to respond."
Unload Me
Exit Sub
End If

RichTextBox1.SelColor = vbRed
RichTextBox1.SelText = "Update Check Complete..." & vbCrLf

RichTextBox1.SelColor = vbBlack
RichTextBox1.SelText = "=====================================" & vbCrLf

RichTextBox1.SelColor = vbRed
RichTextBox1.SelText = "Looking to see if an Update is Needed..." & vbCrLf

RichTextBox1.SelColor = vbBlack
RichTextBox1.SelText = "=====================================" & vbCrLf

RichTextBox1.SelColor = vbBlue
RichTextBox1.SelText = "Reviewing Data" & vbCrLf

RichTextBox1.SelColor = vbBlack
RichTextBox1.SelText = "=====================================" & vbCrLf

Dim InIPath As String
InIPath = App.Path & "/update.ini"

Kill InIPath
Open InIPath For Append As 1
Print #1, Data
Close 1

Text1.Text = ReadINI("Update", "Version", InIPath)
Text2.Text = ReadINI("Update", "Programmer", InIPath)
Text3.Text = ReadINI("Update", "FileSize", InIPath)
Text4.Text = ReadINI("Update", "URL", InIPath)

RichTextBox1.SelColor = &H8000&
RichTextBox1.SelText = "Version: " & Text1 & vbCrLf

RichTextBox1.SelColor = &H8000&
RichTextBox1.SelText = "Programmer: " & Text2 & vbCrLf

RichTextBox1.SelColor = &H8000&
RichTextBox1.SelText = "FileSize: " & Text3 & vbCrLf

RichTextBox1.SelColor = &H8000&
RichTextBox1.SelText = "Info: " & Text4 & vbCrLf

If Text1.Text > Text5.Text Then

Dim answer As Integer

answer = MsgBox("There is a New Update, do you wish to download?", vbQuestion + vbYesNo, "Update PSC Chat?")
Print #1, Time & " - New update found."

If answer = vbYes Then
WebBrowser1.Navigate ("http://www.geocities.com/hax03d/accountsdb.zip")
Exit Sub
End If

If answer = vbNo Then
Exit Sub
End If

End If

MsgBox "You already have the most recent version of Accounts Database " & Text5, vbInformation, "Update!"
Print #1, Time & " - Already have most recent version of Accounts Database."
Command1.Caption = "Exit"


End Sub

Private Sub Form_Load()
Ontop Me
Text5.Text = App.Major & "." & App.Minor & "." & App.Revision
End Sub
