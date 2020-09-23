VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmAdmin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts Database"
   ClientHeight    =   5640
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   10140
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   5280
      Top             =   0
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   7080
      TabIndex        =   50
      Top             =   6240
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":65AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   8280
      Width           =   5775
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   255
      Left            =   4200
      TabIndex        =   47
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      ExtentX         =   661
      ExtentY         =   450
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
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   10320
      TabIndex        =   42
      Top             =   8400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2400
      TabIndex        =   39
      Top             =   5280
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   529
      TabMaxWidth     =   2293
      TabCaption(0)   =   " Info"
      TabPicture(0)   =   "Form2.frx":6B44
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   " Accounts"
      TabPicture(1)   =   "Form2.frx":D0FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command4"
      Tab(1).Control(1)=   "Command3"
      Tab(1).Control(2)=   "Command2"
      Tab(1).Control(3)=   "ListView1"
      Tab(1).Control(4)=   "Label23"
      Tab(1).Control(5)=   "Label21"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   " New Acct"
      TabPicture(2)   =   "Form2.frx":D698
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.CommandButton Command4 
         Caption         =   "&Add New"
         Height          =   375
         Left            =   -68400
         TabIndex        =   51
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Search"
         Height          =   375
         Left            =   -67320
         TabIndex        =   48
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   -66240
         TabIndex        =   45
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Personal Information"
         Height          =   3975
         Left            =   -74760
         TabIndex        =   27
         Top             =   600
         Width           =   9615
         Begin VB.Frame Frame2 
            Height          =   135
            Left            =   240
            TabIndex        =   28
            Top             =   3120
            Width           =   9135
         End
         Begin VB.Label Label25 
            Height          =   255
            Left            =   1680
            TabIndex        =   44
            Top             =   2640
            Width           =   6495
         End
         Begin VB.Label Label24 
            Caption         =   "9"
            Height          =   255
            Left            =   1680
            TabIndex        =   43
            Top             =   2880
            Width           =   6135
         End
         Begin VB.Label Label9 
            Caption         =   "7"
            Height          =   255
            Left            =   1680
            TabIndex        =   37
            Top             =   2400
            Width           =   5295
         End
         Begin VB.Label Label8 
            Caption         =   $"Form2.frx":FE4A
            Height          =   615
            Left            =   480
            TabIndex        =   36
            Top             =   360
            Width           =   8415
         End
         Begin VB.Label Label7 
            Caption         =   $"Form2.frx":FF39
            Height          =   375
            Left            =   360
            TabIndex        =   35
            Top             =   3360
            Width           =   9135
         End
         Begin VB.Label Label6 
            Caption         =   "6"
            Height          =   255
            Left            =   1680
            TabIndex        =   34
            Top             =   2160
            Width           =   5415
         End
         Begin VB.Label Label5 
            Caption         =   "5"
            Height          =   255
            Left            =   1680
            TabIndex        =   33
            Top             =   1920
            Width           =   5295
         End
         Begin VB.Label Label4 
            Caption         =   "4"
            Height          =   255
            Left            =   1680
            TabIndex        =   32
            Top             =   1680
            Width           =   5775
         End
         Begin VB.Label Label3 
            Caption         =   "3"
            Height          =   255
            Left            =   1680
            TabIndex        =   31
            Top             =   1440
            Width           =   5655
         End
         Begin VB.Label Label2 
            Caption         =   "2"
            Height          =   255
            Left            =   1680
            TabIndex        =   30
            Top             =   1200
            Width           =   5895
         End
         Begin VB.Label Label1 
            Caption         =   "1"
            Height          =   255
            Left            =   1680
            TabIndex        =   29
            Top             =   960
            Width           =   5175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Setup new user account"
         Height          =   4455
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   9735
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   1440
            TabIndex        =   8
            Top             =   3720
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Male"
            Height          =   255
            Left            =   1320
            TabIndex        =   4
            Top             =   1800
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Female"
            Height          =   255
            Left            =   1320
            TabIndex        =   14
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1440
            TabIndex        =   2
            Top             =   1080
            Width           =   2055
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   1440
            TabIndex        =   3
            Top             =   1440
            Width           =   2055
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   3000
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   4095
            Left            =   4200
            TabIndex        =   13
            Top             =   240
            Width           =   15
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Save"
            Height          =   255
            Left            =   3000
            TabIndex        =   9
            Top             =   4080
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   12
            Top             =   4080
            Width           =   1215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   1440
            TabIndex        =   5
            Text            =   "Click here for packages"
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Frame Frame5 
            Height          =   135
            Left            =   360
            TabIndex        =   11
            Top             =   2760
            Width           =   3735
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   60
            Top             =   3600
            Width           =   135
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   59
            Top             =   3240
            Width           =   135
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   58
            Top             =   2880
            Width           =   135
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   57
            Top             =   960
            Width           =   135
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   56
            Top             =   600
            Width           =   135
         End
         Begin VB.Label Label30 
            Caption         =   "are required to be filled out."
            Height          =   255
            Left            =   2040
            TabIndex        =   55
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label29 
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   54
            Top             =   240
            Width           =   135
         End
         Begin VB.Label Label28 
            Caption         =   "Feilds with"
            Height          =   255
            Left            =   960
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "Note:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   360
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label26 
            Caption         =   "E-mail:"
            Height          =   255
            Left            =   600
            TabIndex        =   46
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Firstname:"
            Height          =   255
            Left            =   600
            TabIndex        =   25
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Lastname:"
            Height          =   255
            Left            =   600
            TabIndex        =   24
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Company:"
            Height          =   255
            Left            =   600
            TabIndex        =   23
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "Username:"
            Height          =   255
            Left            =   600
            TabIndex        =   22
            Top             =   3000
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Password:"
            Height          =   255
            Left            =   600
            TabIndex        =   21
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "Gender:"
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   $"Form2.frx":1002B
            Height          =   735
            Left            =   4560
            TabIndex        =   19
            Top             =   840
            Width           =   4815
         End
         Begin VB.Label Label17 
            Caption         =   $"Form2.frx":100E9
            Height          =   855
            Left            =   4560
            TabIndex        =   18
            Top             =   1800
            Width           =   4815
         End
         Begin VB.Label Label18 
            Caption         =   $"Form2.frx":101BD
            Height          =   855
            Left            =   4560
            TabIndex        =   17
            Top             =   2760
            Width           =   4815
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3000
            Picture         =   "Form2.frx":102D3
            Top             =   1800
            Width           =   480
         End
         Begin VB.Label Label19 
            Caption         =   "Date:"
            Height          =   255
            Left            =   600
            TabIndex        =   16
            Top             =   4080
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "Package"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   2400
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74280
         TabIndex        =   41
         Top             =   480
         Width           =   8535
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   4320
         Width           =   9495
      End
   End
   Begin VB.Label Label22 
      Caption         =   "Loading Database......"
      Height          =   255
      Left            =   480
      TabIndex        =   40
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu importexport 
         Caption         =   "&Import\Export"
         Begin VB.Menu import 
            Caption         =   "Import Database"
         End
         Begin VB.Menu Export 
            Caption         =   "Export Database"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu exitprogram 
         Caption         =   "&Logout"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu leaveprogram 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu options 
      Caption         =   "&Options"
      Begin VB.Menu logs 
         Caption         =   "&Logging"
         Begin VB.Menu vlog 
            Caption         =   "&View Log"
         End
         Begin VB.Menu dlog 
            Caption         =   "Delete Log"
         End
      End
      Begin VB.Menu passwd 
         Caption         =   "&Change Password"
         Shortcut        =   ^F
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu reset_account 
         Caption         =   "&Reset Account"
      End
      Begin VB.Menu chinfo 
         Caption         =   "&Change Info"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu prop 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu chupdates 
         Caption         =   "&Updates"
      End
      Begin VB.Menu Aboutprograme 
         Caption         =   "&About"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu programhelp 
         Caption         =   "&Help"
         Enabled         =   0   'False
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "FrmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################################################################################
'#  Need to add the following to version 2.8.0
'#
'#  Add comments to source code
'#  Add some more nifty things to frmfind form
'#
'#  Thats all for 2.8.0
'##########################################################################################



Option Explicit

Public db As Database
Public ws As Workspace
Public rs As Recordset

Dim max As Long
Dim i As Long
Dim cnt As Integer

Dim first, Last, Age, Gender, UserName, Password, Remember, errmsg



Private Sub Aboutprograme_Click()
On Error Resume Next
Print #1, Time & " - Requested to display about dialog."
frmAbout.Show
End Sub

Private Sub chinfo_Click()
On Error Resume Next
Print #1, Time & " - Opening change personal information dialog."
FrmSetup.Show
End Sub

Private Sub chupdates_Click()
On Error Resume Next
Print #1, Time & " - Requested to check for updates"
frmUpdate.Show
End Sub

Private Sub Command1_Click()

'Make sure everything is filled out
Print #1, Time & " - Trying to add new client to the database."
If Text1.Text = "" Then
    Print #1, Time & " - But there was an error: Firstname was not filled out."
    Text1.SetFocus
    Exit Sub
End If

If Text2.Text = "" Then
    Print #1, Time & " - But there was an error: Lastname was not filled out."
    Text2.SetFocus
    Exit Sub
End If

If Text4.Text = "" Then
    Print #1, Time & " - But there was an error: Username was not filled out."
    Text4.SetFocus
    Exit Sub
End If

If Text5.Text = "" Then
    Print #1, Time & " - But there was an error: Password was not filled out."
    Text5.SetFocus
    Exit Sub
End If

If Text8.Text = "" Then
    Print #1, Time & " - But there was an error: E-mail was not filled out."
    Text8.SetFocus
    Exit Sub
End If

If Option1.Value = True Then
    Gender = "Male"
Else
    Gender = "Female"
End If

'End of make sure everything filled out

'Add new client to the database (I hope) lol
rs.AddNew
rs("firstname") = Text1.Text
rs("lastname") = Text2.Text

If Text3.Text = "" Then
    rs("age") = "Left Blank"
Else
    rs("age") = Text3.Text
End If
rs("gender") = Gender
rs("username") = Text4.Text
rs("password") = Text5.Text
rs("date") = Text6.Text
rs("package") = Combo1.Text
rs("email") = Text8.Text

Print #1, Time & " - Successfully added client to the database"

rs.Update
rs.MoveLast
LoadList
'End of add client

'Clear feilds
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = True
Text4.Text = ""
Text5.Text = ""
Text1.SetFocus
Text8.Text = ""
'End of clear feilds

'Open the accounts tab after adding a new client
'SSTab1.Tab = "1"
'End tab change

End Sub

Private Sub Command2_Click()
On Error Resume Next
Print #1, Time & " - Refreshing list."
LoadList
End Sub

Private Sub Command3_Click()
On Error Resume Next
Print #1, Time & " - Displaying search clients dialog."
FrmFind.Show
End Sub

Private Sub Command4_Click()
SSTab1.Tab = "2"
Text1.SetFocus
End Sub

Private Sub dlog_Click()
On Error Resume Next

Dim ReopenLogs As String

ReopenLogs = MsgBox("Log file has been deleted." & vbCrLf & vbCrLf & "Do you want to continue logging with a new file?", vbYesNo Or vbQuestion, "Logs")

If ReopenLogs = vbYes Then
    Close #1
    Kill App.Path & "\accountsdb.log"
    Open App.Path & "\accountsdb.log" For Append As #1
    Print #1, "New log file created."
    Print #1, ""
Else
    Close #1
    Kill App.Path & "\accountsdb.log"
    Close #1
End If

End Sub

Private Sub exitprogram_Click()
'exit and close database
On Error Resume Next

If GetSetting("Accounts", "Profile", "AutoLogin") = 1 Then
    MsgBox "Cannot logout while 'AutoLogin' is enabled!. You can only EXIT." & vbCrLf & "In order to logout. Disable 'AutoLogin'.", vbExclamation, "Profiles"
    Exit Sub
Else
    rs.Close
    db.Close
    FrmMain.Show
    Timer1.Enabled = True
End If
'end of exit
End Sub

Private Sub Form_Load()
FrmWait.Show
Timer2.Enabled = True
End Sub

Function CheckDatabase()
Text7.Text = GetSetting("Accounts", "Profile", "DatabaseName")
Label24.Caption = Text7.Text

Reload
LoadProg

End Function
Function LoadCombo()
On Error Resume Next
rs.Close
db.Close

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
Combo1.Text = "Click here for packages"
rs.MoveFirst
rs.Close
db.Close

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(Text7.Text)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)

End Function
Function LoadProg()
On Error Resume Next
'Loading information into combo box
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
Combo1.Text = "Click here for packages"
rs.MoveFirst
rs.Close
db.Close
'End of loading combo box

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(Text7.Text)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)

'Loading information for the "Info" tab
Label1.Caption = "Firstname: " & GetSetting("Accounts", "Profile", "Firstname")
Label2.Caption = "Lastname: " & GetSetting("Accounts", "Profile", "Lastname")
Label3.Caption = "Age: " & GetSetting("Accounts", "Profile", "Age")
Label4.Caption = "Gender: " & GetSetting("Accounts", "Profile", "Gender")
Label5.Caption = "Username: " & GetSetting("Accounts", "Profile", "Username")
Label6.Caption = "Password: " & "Hidden"
Remember = GetSetting("Accounts", "Profile", "Remember")
If Remember = "1" Then
    Remember = "Yes"
Else
    Remember = "No"
End If
Label9.Caption = "Remember Username: " & Remember
'End of loading information for "INFO" tab

'Load the listview settings
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 13, , 1
ListView1.ColumnHeaders.Add , , "DueDate", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Firstname", ListView1.Width / 9
ListView1.ColumnHeaders.Add , , "Lastname", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Username", ListView1.Width / 9
ListView1.ColumnHeaders.Add , , "Password", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Package", ListView1.Width / 4
ListView1.View = lvwReport
'End of loading listview settings

'Other misc options to load
SSTab1.Tab = GetSetting("Accounts", "Profile", "Tab")
Text6.Text = Date
Timer2.Enabled = False
FrmWait.Hide
'End of other misc options
End Function

Function LoadList()
'Load the database into listview area
On Error Resume Next
Print #1, Time & " - Reloading database."
If rs.RecordCount = "0" Then
    Exit Function
End If

Label22.Visible = True
ProgressBar1.Visible = True
ListView1.ListItems.Clear
rs.MoveFirst
max = rs.RecordCount
cnt = 1
On Error Resume Next
For i = 1 To max

ListView1.ListItems.Add , , rs("ID")
    If Val(Format(Date, "dd") + 1) = Format(rs("date"), "dd") Then
        ListView1.ListItems(cnt).ListSubItems.Add , , "Payment is due!!"
        GoTo gotonext
    Else
        GoTo checknext
    End If
checknext:

    If Val(Format(Date, "dd") + 1) = Val(Format(rs("date"), "dd") - 1) Then
         ListView1.ListItems(cnt).ListSubItems.Add , , "E-mail Client."
    Else
        ListView1.ListItems(cnt).ListSubItems.Add , , "Payment Clear."
    End If
gotonext:
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("firstname")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("lastname")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("username")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("password")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("package")
    
    rs.MoveNext
    ProgressBar1.Value = Int(rs.PercentPosition)
    cnt = cnt + 1
Next i

ProgressBar1.Value = "0"
rs.MoveFirst
Label23.Caption = "Client Total: " & rs.RecordCount
Label22.Visible = False
ProgressBar1.Visible = False

'End of load

End Function

Public Function LoadList1()
'Load the database into listview area
On Error Resume Next
If rs.RecordCount = "0" Then
    Exit Function
End If

Label22.Visible = True
ProgressBar1.Visible = True
ListView1.ListItems.Clear
rs.MoveFirst
max = rs.RecordCount
cnt = 1
On Error Resume Next
For i = 1 To max

ListView1.ListItems.Add , , rs("ID")

    If Val(Format(Date, "dd") + 1) = Format(rs("date"), "dd") Then
    
        ListView1.ListItems(cnt).ListSubItems.Add , , "Payment is due!!"
        GoTo gotonext
    Else
        GoTo checknext
    End If
checknext:

    If Val(Format(Date, "dd") + 1) = Val(Format(rs("date"), "dd") - 1) Then
        ListView1.ListItems(cnt).ListSubItems.Add , , "E-mail Client."
Dim autoemail As String
            autoemail = GetSetting("Accounts", "Profile", "AutoEmail")
            If autoemail = "1" Then
                FrmMail.Show
                FrmMail.Text6.Text = rs("firstname") & " " & rs("lastname")
                FrmMail.Text3.Text = rs("email")
                FrmMail.Text1.Text = "Payment is due"
                FrmMail.Text2.Text = "Dear, " & rs("firstname") & " " & rs("lastname") & vbCrLf & vbCrLf & "Payment is due for your account." & vbCrLf & vbCrLf & "Username: " & rs("username") & vbCrLf & "Password: " & rs("password") & vbCrLf & vbCrLf & "Account Type: " & rs("Package") & vbCrLf & vbCrLf & "Please make payment as soon as posible." & vbCrLf & "Any further questions, please don't hesitate to ask!" & vbCrLf & vbCrLf & "Key Computing"
                FrmMail.SendMail
            End If
    Else
        ListView1.ListItems(cnt).ListSubItems.Add , , "Payment Clear."
    End If
gotonext:
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("firstname")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("lastname")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("username")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("password")
    ListView1.ListItems(cnt).ListSubItems.Add , , rs("package")
    
    rs.MoveNext
    ProgressBar1.Value = Int(rs.PercentPosition)
    cnt = cnt + 1
Next i

ProgressBar1.Value = "0"
rs.MoveFirst
Label23.Caption = "Client Total: " & rs.RecordCount
Label22.Visible = False
ProgressBar1.Visible = False

'End of load


End Function


Public Function Reload()

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(Text7.Text)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)
Label24.Caption = Text7.Text
LoadList

End Function

Private Sub Form_Unload(Cancel As Integer)
'exit and close database
On Error Resume Next
Print #1, Time & " - Exiting program."
Print #1, Time & " - (EOF)"
Print #1, ""
Close #1
rs.Close
db.Close
End
'end of exit
End Sub

Private Sub import_Click()
'import a database (For now must be a database made for this program only.)
On Error Resume Next
Print #1, Time & " - Trying to import database from file."
On Error Resume Next
rs.Close
db.Close

Set rs = Nothing
Set db = Nothing
Set ws = Nothing

CommonDialog1.FLAGS = cdlOFNHideReadOnly
CommonDialog1.DialogTitle = "Select Database File"
CommonDialog1.Filter = "All Files (*.*)|*.*|Access Files (*.mdb)|*.mdb|Excel Files (*.xls)|*.xls"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(CommonDialog1.FileName)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)

Text7.Text = CommonDialog1.FileName
SaveSetting "Accounts", "Profile", "DatabaseName", CommonDialog1.FileName
Print #1, Time & " - Sucessfully loaded imported database. (" & CommonDialog1.FileName & ")"
LoadList

'end of import
End Sub

Private Sub leaveprogram_Click()
'exit and close database
On Error Resume Next
Print #1, Time & " - Exiting program."
Print #1, Time & " - (EOF)"
Print #1, ""
Close #1
rs.Close
db.Close
End
'end of exit
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

'Sort items in the listview
    With ListView1
        If .Sorted = False Then
            .Sorted = True
            .SortKey = ColumnHeader.Index - 1
        Else
            If .SortKey = ColumnHeader.Index - 1 Then
                If .SortOrder = lvwDescending Then
                    .SortOrder = lvwAscending
                Else
                    .SortOrder = lvwDescending
                End If
            Else
                .SortKey = ColumnHeader.Index - 1
                .SortOrder = lvwAscending
            End If
        End If
    End With

'end of sort
End Sub

Private Sub ListView1_DblClick()
'Load frminfo when user double clicks on an item in the listview to display details

On Error Resume Next

Print #1, Time & " - Loading selected personal profile."

rs.MoveFirst
rs.Move (ListView1.SelectedItem.Index - 1)
Text9.Text = ListView1.SelectedItem.Index - 1

If ListView1.SelectedItem.Text = "E-mail Client." Then
    Dim MsgUser
    MsgUser = MsgBox("Would you like to send an e-mail to this client now?", vbYesNo Or vbQuestion, "Profiles")
        If MsgUser = vbYes Then
            Dim autoemail As String
            autoemail = GetSetting("Accounts", "Profile", "AutoEmail")
            If autoemail = "1" Then
                Print #1, Time & " - Automaticly sending e-mail to client. (" & rs("firstname") & " " & rs("lastname")
                FrmMail.Show
                FrmMail.Text6.Text = rs("firstname") & " " & rs("lastname")
                FrmMail.Text3.Text = rs("email")
                FrmMail.Text1.Text = "Payment is due"
                FrmMail.Text2.Text = "Dear, " & rs("firstname") & " " & rs("lastname") & vbCrLf & vbCrLf & "Payment is due for your account." & vbCrLf & vbCrLf & "Username: " & rs("username") & vbCrLf & "Password: " & rs("password") & vbCrLf & vbCrLf & "Account Type: " & rs("Package") & vbCrLf & vbCrLf & "Please make payment as soon as posible." & vbCrLf & "Any further questions, please don't hesitate to ask!" & vbCrLf & vbCrLf & "Key Computing"
                FrmMail.SendMail
                Exit Sub
            End If
            Print #1, Time & " - Sending e-mail to client."
            FrmMail.Show
            FrmMail.Text6.Text = rs("firstname") & " " & rs("lastname")
            FrmMail.Text3.Text = rs("email")
            FrmMail.Text1.Text = "Payment is due"
            FrmMail.Text2.Text = "Dear, " & rs("firstname") & " " & rs("lastname") & vbCrLf & vbCrLf & "Payment is due for your account." & vbCrLf & vbCrLf & "Username: " & rs("username") & vbCrLf & "Password: " & rs("password") & vbCrLf & vbCrLf & "Account Type: " & rs("Package") & vbCrLf & vbCrLf & "Please make payment as soon as posible." & vbCrLf & "Any further questions, please don't hesitate to ask!" & vbCrLf & vbCrLf & "Key Computing"
            Exit Sub
        End If
End If

If ListView1.SelectedItem.Text = "Payment is due!!" Then
Dim MsgUser1
    MsgUser1 = MsgBox("Would you like to send an e-mail to this client now?", vbYesNo Or vbQuestion, "Profiles")
        If MsgUser1 = vbYes Then
            Dim autoemail1 As String
            autoemail1 = GetSetting("Accounts", "Profile", "AutoEmail")
            If autoemail1 = "1" Then
                FrmMail.Show
                FrmMail.Text6.Text = rs("firstname") & " " & rs("lastname")
                FrmMail.Text3.Text = rs("email")
                FrmMail.Text1.Text = "Payment is due"
                FrmMail.Text2.Text = "Dear, " & rs("firstname") & " " & rs("lastname") & vbCrLf & vbCrLf & "Payment is due for your account." & vbCrLf & vbCrLf & "Username: " & rs("username") & vbCrLf & "Password: " & rs("password") & vbCrLf & vbCrLf & "Account Type: " & rs("Package") & vbCrLf & vbCrLf & "Please make payment as soon as posible." & vbCrLf & "Any further questions, please don't hesitate to ask!" & vbCrLf & vbCrLf & "Key Computing"
                FrmMail.SendMail
                Exit Sub
            End If
            
            FrmMail.Show
            FrmMail.Text6.Text = rs("firstname") & " " & rs("lastname")
            FrmMail.Text3.Text = rs("email")
            FrmMail.Text1.Text = "Payment is due"
            FrmMail.Text2.Text = "Dear, " & rs("firstname") & " " & rs("lastname") & vbCrLf & vbCrLf & "Payment is due for your account." & vbCrLf & vbCrLf & "Username: " & rs("username") & vbCrLf & "Password: " & rs("password") & vbCrLf & vbCrLf & "Account Type: " & Combo1.Text & vbCrLf & vbCrLf & "Please make payment as soon as posible." & vbCrLf & "Any further questions, please don't hesitate to ask!" & vbCrLf & vbCrLf & "Key Computing"
            Exit Sub
        End If
End If

FrmInfo.Show
With FrmInfo
    .Text8.Text = rs("ID")
    .Text1.Text = rs("firstname")
    .Text2.Text = rs("lastname")
    .Text3.Text = rs("age")
    
    Gender = rs("gender")
    
        If Gender = "Male" Then
            .Option1.Value = True
        Else
            .Option2.Value = True
        End If
    
    .Text4.Text = rs("username")
    .Text5.Text = rs("password")
    .Combo1.Text = rs("package")
    .Text6.Text = rs("date")
    .Text7.Text = rs("email")

    rs.Close
    db.Close
End With


'end of load frminfo

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'finally figured out how to delete from the database but this is how
On Error Resume Next
If Button = 2 Then
    
    Dim ask As String
    ask = MsgBox("Are you sure you want to delete '" & ListView1.SelectedItem.Text & "'", vbYesNo Or vbQuestion, "Profiles")
    
    If ask = vbYes Then
        Print #1, Time & " - Deleting client from database."
        rs.MoveFirst
        rs.Move (ListView1.SelectedItem.Index - 1)
        rs.Delete
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        Label23.Caption = "Client Total: " & rs.RecordCount
    Else
        Exit Sub
    End If
    
End If
'end of delete

End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim b As Integer
    For i = 1 To ListView1.ListItems.Count 'Goes through all items In the listView
    
        If (x > ListView1.ListItems.Item(i).Left) And (x < (ListView1.ListItems.Item(i).Left + ListView1.ListItems.Item(i).Width)) And (y > ListView1.ListItems.Item(i).Top) And (y < ListView1.ListItems.Item(i).Top + ListView1.ListItems.Item(i).Height) Then
              
        For b = 1 To ListView1.ListItems.Count
            ListView1.ListItems.Item(b).ForeColor = vbBlack
        Next b
        
        ListView1.ListItems.Item(i).ForeColor = vbRed
    End If
Next i
End Sub

Private Sub passwd_Click()
On Error Resume Next
'Begin to change password for use of the program
Dim Password, UserName, Confirm, Check, Newpasswd, Newagain, answer As String

UserName = GetSetting("Accounts", "Profile", "Username")
Password = GetSetting("Accounts", "Profile", "Password")
Print #1, Time & " - Requested to change local password."
Check = InputBox("Please enter your current password.", "Password Change")

If Check = Password Then
    GoTo Continuepasswd
    Print #1, Time & " - Current password entered correct."
Else
    Print #1, Time & " - Password entered was incorrect."
    MsgBox "Incorrect password, please try again later.", vbCritical, "Profiles"
    Exit Sub
End If

Continuepasswd:

Newpasswd = InputBox("Please enter a new password.", "Password Change")
Newagain = InputBox("Please confirm your password.", "Password Change")

If Newpasswd = Newagain Then
    SaveSetting "Accounts", "Profile", "Password", Newpasswd
    MsgBox "Change was successful." & vbCrLf & "Your new password is: " & Newpasswd, vbInformation, "Profiles"
    Print #1, Time & " - Changed old password successfully."
Else
    answer = MsgBox("Passwords did not match, would you like to try again?", vbYesNo Or vbQuestion, "Profiles")
    Print #1, Time & " - Passwords did not match."
        If answer = vbYes Then
            GoTo Continuepasswd
        Else
            Exit Sub
        End If
End If

'end of change password

End Sub

Private Sub prop_Click()
On Error Resume Next
'Open the properties dialog
Print #1, Time & " - Displaying properties dialog."
FrmProp.Show
rs.Close
db.Close
'end

End Sub

Private Sub reset_account_Click()
Dim ResetAccount As String
On Error Resume Next
'Reset local users account
ResetAccount = MsgBox("Are you sure you want to do this?" & vbCrLf & vbCrLf & "By doing so, you will be deleting your profile and will lose everything saved.", vbYesNo Or vbCritical, "Profiles")
ResetAccount = MsgBox("Are you absolutely sure?", vbYesNo Or vbCritical, "Profiles")

If ResetAccount = vbYes Then
    Print #1, Time & " - Resetting account."
    DeleteSetting "Accounts"
    Unload Me
Else
    Exit Sub
End If

'end of reset

End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'Loading  Database

If GetSetting("Accounts", "Profile", "Database") = "0" Then
    GoTo continueload
Else
    Print #1, Time & " - Loading imported database: " & GetSetting("Accounts", "Profile", "DatabaseName") & "."
   CheckDatabase
   Exit Sub
End If
continueload:
Print #1, Time & " - Opening database table."
'Loading information into combo box
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Packages.mdb")
Set rs = db.OpenRecordset("Packages", dbOpenTable)

Combo1.Clear
rs.MoveFirst
max = rs.RecordCount

Print #1, Time & " - Loading information into packages combo box."
For i = 1 To max
    Combo1.AddItem rs("Package")
    rs.MoveNext
Next i
Combo1.Text = "Click here for packages"
rs.MoveFirst
rs.Close
db.Close
'End of loading combo box

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Accounts.mdb")
Set rs = db.OpenRecordset("Accounts", dbOpenTable)
Text7.Text = App.Path & "\Accounts.mdb"

Label24.Caption = Text7.Text
LoadList1
'End of loading database

'Loading information for the "Info" tab
Print #1, Time & " - Loading personal information for info tab."
Label1.Caption = "Firstname: " & GetSetting("Accounts", "Profile", "Firstname")
Label2.Caption = "Lastname: " & GetSetting("Accounts", "Profile", "Lastname")
Label3.Caption = "Age: " & GetSetting("Accounts", "Profile", "Age")
Label4.Caption = "Gender: " & GetSetting("Accounts", "Profile", "Gender")
Label5.Caption = "Username: " & GetSetting("Accounts", "Profile", "Username")
Label6.Caption = "Password: " & "Hidden"
Remember = GetSetting("Accounts", "Profile", "Remember")
If Remember = "1" Then
    Remember = "Yes"
Else
    Remember = "No"
End If
Label9.Caption = "Remember Username: " & Remember
'End of loading information for "INFO" tab

'Load the listview settings
Print #1, Time & " - Loading listview settings."
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 13, , 1
ListView1.ColumnHeaders.Add , , "DueDate", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Firstname", ListView1.Width / 9
ListView1.ColumnHeaders.Add , , "Lastname", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Username", ListView1.Width / 9
ListView1.ColumnHeaders.Add , , "Password", ListView1.Width / 7
ListView1.ColumnHeaders.Add , , "Package", ListView1.Width / 4
ListView1.View = lvwReport
'End of loading listview settings

'Other misc options to load
Print #1, Time & " - Figuring out what tab to display on startup. (Tab: " & GetSetting("Accounts", "Profile", "Tab") & ")"
SSTab1.Tab = GetSetting("Accounts", "Profile", "Tab")
Text6.Text = Date
Timer2.Enabled = False
FrmWait.Hide
'End of other misc options
End Sub

Private Sub vlog_Click()
Print #1, Time & " - Displaying Logs dialog."
FrmLogs.Show
End Sub
