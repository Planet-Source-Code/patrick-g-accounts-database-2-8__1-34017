VERSION 5.00
Begin VB.Form FrmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Information"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3705
   Icon            =   "FrmInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   2640
      TabIndex        =   26
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1320
      TabIndex        =   25
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Text            =   "Click here for packages"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Female"
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Male"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   22
      Top             =   120
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   360
      TabIndex        =   19
      Top             =   4680
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Account #:"
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Please verify that all the above information is correct.  If all is correct please click ""Save"""
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   4800
      Width           =   3255
   End
   Begin VB.Label Label8 
      Caption         =   "Package:"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Date:"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Firstname:"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Lastname:"
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Company:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Gender:"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Username:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Password:"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3000
      Picture         =   "FrmInfo.frx":65AA
      Top             =   1560
      Width           =   480
   End
End
Attribute VB_Name = "FrmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public db As Database
Public ws As Workspace
Public rs As Recordset

Dim max As Long
Dim i As Long



Private Sub Command1_Click()
'do i even have to explain
FrmAdmin.Reload
Unload Me
'end
End Sub

Private Sub Command2_Click()
'Edit current information in the database

Dim Gender As String
If Option1.Value = True Then
    Gender = "Male"
Else
    Gender = "Female"
End If

Print #1, Time & " - Modifying user information for: " & Text1.Text & " " & Text2.Text & "."
rs.MoveFirst
rs.Move (FrmAdmin.Text9.Text)
rs.Edit
rs("firstname") = Text1.Text
rs("lastname") = Text2.Text
rs("age") = Text3.Text
rs("gender") = Gender
rs("username") = Text4.Text
rs("password") = Text5.Text
rs("package") = Combo1.Text
rs("date") = Text6.Text
rs("email") = Text7.Text
rs.Update

FrmAdmin.Reload

'end of edit
Unload Me
End Sub

Private Sub Command3_Click()

Print #1, Time & " - Sending personal e-mail to: "; rs("firstname") & " " & rs("lastname")
FrmMail.Show
FrmMail.Text6.Text = Text1.Text & " " & Text2.Text
FrmMail.Text3.Text = Text7.Text
FrmMail.Text1.Text = "To: " & Text1.Text & " " & Text2.Text
FrmMail.Text2.Text = "Dear, " & Text1.Text & " " & Text2.Text & vbCrLf & vbCrLf & "Enter message here!"

End Sub

Private Sub Form_Load()
'load information into the combo box
Ontop Me
Comboload
'end of load
End Sub

Function Comboload()
On Error Resume Next

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Packages.mdb")
Set rs = db.OpenRecordset("Packages", dbOpenTable)

Combo1.Clear
rs.MoveFirst
max = rs.RecordCount

For i = 1 To max
    Combo1.AddItem rs("Package")
    rs.MoveNext
Next i

rs.MoveFirst
rs.Close
db.Close

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(FrmAdmin.Text7.Text)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)
End Function
