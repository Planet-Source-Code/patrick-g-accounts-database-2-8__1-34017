VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Properties"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10395
   Icon            =   "FrmProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7440
      TabIndex        =   39
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   375
      Left            =   8400
      TabIndex        =   38
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   9360
      TabIndex        =   37
      Top             =   5280
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "Packages"
      Height          =   5055
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   6120
         TabIndex        =   42
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Add Package"
         Height          =   375
         Left            =   4560
         TabIndex        =   35
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3480
         TabIndex        =   34
         Top             =   3480
         Width           =   2295
      End
      Begin VB.ListBox List1 
         Height          =   3180
         ItemData        =   "FrmProp.frx":65AA
         Left            =   240
         List            =   "FrmProp.frx":65AC
         TabIndex        =   32
         Top             =   1320
         Width           =   3015
      End
      Begin VB.Frame Frame7 
         Height          =   135
         Left            =   120
         TabIndex        =   31
         Top             =   720
         Width           =   7455
      End
      Begin VB.Label Label19 
         Caption         =   "Price:"
         Height          =   255
         Left            =   6120
         TabIndex        =   43
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label18 
         Caption         =   "$"
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         Top             =   3480
         Width           =   135
      End
      Begin VB.Line Line1 
         X1              =   3360
         X2              =   7440
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label17 
         Caption         =   "To remove a package from the list to the right.  Right click on an item using your mouse and follow the instructions."
         Height          =   375
         Left            =   3360
         TabIndex        =   40
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label Label16 
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1080
         Width           =   3015
      End
      Begin VB.Label Label15 
         Caption         =   "Enter new package Below:"
         Height          =   255
         Left            =   3480
         TabIndex        =   33
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label14 
         Caption         =   $"FrmProp.frx":65AE
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   7335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Licence"
      Height          =   5055
      Left            =   2640
      TabIndex        =   23
      Top             =   120
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox Text1 
         Height          =   4575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "FrmProp.frx":6662
         Top             =   240
         Width           =   7455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Startup Options"
      Height          =   5055
      Left            =   2640
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox Check6 
         Caption         =   "Enable logging of all activity"
         Height          =   195
         Left            =   360
         TabIndex        =   45
         Top             =   2760
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Automaticly e-mail client without prompting"
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Top             =   2400
         Width           =   3495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   3120
         Width           =   375
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   4080
         TabIndex        =   27
         Top             =   3120
         Width           =   2655
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Remember last imported database"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3120
         Width           =   2775
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Login automaticly on startup"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Remember username on startup"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Display ""New Acct"" tab on startup"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   4320
         Width           =   2775
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Display ""Accounts"" tab on startup"
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   3960
         Width           =   2775
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Display ""Info"" tab on startup"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   3600
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Display splash screen on startup"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   1320
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.Frame Frame4 
         Height          =   135
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   6495
      End
      Begin VB.Label Label11 
         Caption         =   "Database:"
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Display tab on startup"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   $"FrmProp.frx":AE14
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Information"
      Height          =   4815
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Frame Frame2 
         Height          =   135
         Left            =   240
         TabIndex        =   1
         Top             =   3600
         Width           =   6495
      End
      Begin VB.Label Label10 
         Caption         =   "8"
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   3000
         Width           =   4815
      End
      Begin VB.Label Label9 
         Caption         =   "5"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   $"FrmProp.frx":AECD
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label7 
         Caption         =   $"FrmProp.frx":AFBC
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   3840
         Width           =   6255
      End
      Begin VB.Label Label6 
         Caption         =   "6"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label5 
         Caption         =   "7"
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   1800
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "2"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   1320
         Width           =   3855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProp.frx":B0AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProp.frx":11668
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProp.frx":11C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProp.frx":143B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProp.frx":146CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5055
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8916
      _Version        =   393217
      Indentation     =   443
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   2640
      Picture         =   "FrmProp.frx":17B38
      Top             =   120
      Width           =   7500
   End
End
Attribute VB_Name = "FrmProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset

Dim max As Long
Dim i As Long

Private Sub Check5_Click()
If Check5.Value = "1" Then
    Command4.Visible = True
    Text2.Visible = True
    Label11.Visible = True
Else
    Command4.Visible = False
    Text2.Visible = False
    Label11.Visible = False
End If
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()

On Error Resume Next
Dim Op As String

If Option1.Value = True Then
    Op = "0"
ElseIf Option2.Value = True Then
    Op = "1"
ElseIf Option3.Value = True Then
    Op = "2"
End If

SaveSetting "Accounts", "Profile", "Logo", Check1.Value
SaveSetting "Accounts", "Profile", "Remember", Check2.Value
SaveSetting "Accounts", "Profile", "Rememberpasswd", Check3.Value
SaveSetting "Accounts", "Profile", "AutoLogin", Check4.Value
SaveSetting "Accounts", "Profile", "Tab", Op
SaveSetting "Accounts", "Profile", "Database", Check5.Value
SaveSetting "Accounts", "Profile", "DatabaseName", Text2.Text
SaveSetting "Accounts", "Profile", "AutoEmail", Check3.Value
SaveSetting "Accounts", "Profile", "Logging", Check6.Value

FrmAdmin.Label24.Caption = Text2.Text
If Check5.Value = "1" Then
    Dim question As String
    question = MsgBox("Would you like to import the database now?", vbQuestion Or vbYesNo, "Profiles")
    If question = vbYes Then
        FrmAdmin.Text7.Text = Text2.Text
        FrmAdmin.Reload
        FrmAdmin.LoadCombo
    End If
End If
If Check6.Value = 1 Then
    Open App.Path & "\accountsdb.log" For Append As #1
Else
    Close #1
End If
FrmAdmin.Reload
FrmAdmin.LoadCombo
Unload Me



End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim Op As String

If Option1.Value = True Then
    Op = "0"
ElseIf Option2.Value = True Then
    Op = "1"
ElseIf Option3.Value = True Then
    Op = "2"
End If

SaveSetting "Accounts", "Profile", "Logo", Check1.Value
SaveSetting "Accounts", "Profile", "Remember", Check2.Value
SaveSetting "Accounts", "Profile", "Rememberpasswd", Check3.Value
SaveSetting "Accounts", "Profile", "AutoLogin", Check4.Value
SaveSetting "Accounts", "Profile", "Tab", Op
SaveSetting "Accounts", "Profile", "Database", Check5.Value
SaveSetting "Accounts", "Profile", "DatabaseName", Text2.Text
SaveSetting "Accounts", "Profile", "AutoEmail", Check3.Value
SaveSetting "Accounts", "Profile", "Logging", Check6.Value
FrmAdmin.Label24.Caption = Text2.Text
If Check5.Value = "1" Then
    Dim question As String
    question = MsgBox("Would you like to import the database now?", vbQuestion Or vbYesNo, "Profiles")
    If question = vbYes Then
        FrmAdmin.Text7.Text = Text2.Text
        FrmAdmin.Reload
    End If
End If
If Check6.Value = 1 Then
    Open App.Path & "\accountsdb.log" For Append As #1
Else
    Close #1
End If

End Sub

Private Sub Command4_Click()
'Select a database to be automaticly loaded on load
Text2.Text = ""
CommonDialog1.FLAGS = cdlOFNHideReadOnly
CommonDialog1.DialogTitle = "Select Database File"
CommonDialog1.Filter = "All Files (*.*)|*.*|Access Files (*.mdb)|*.mdb|Excel Files (*.xls)|*.xls"
CommonDialog1.FilterIndex = 2
CommonDialog1.ShowOpen
Text2.Text = CommonDialog1.FileName
'End of auto load
End Sub

Private Sub Command5_Click()
On Error Resume Next
If Text3.Text = "" Then
    Text3.SetFocus
    Exit Sub
End If

If Text4.Text = "" Then
    Text4.SetFocus
    Exit Sub
End If


Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Packages.mdb")
Set rs = db.OpenRecordset("Packages", dbOpenTable)
rs.AddNew
rs("Package") = Text3.Text & " ($" & Text4.Text & ")"
rs.Update
rs.MoveLast
Text3.Text = ""
Text4.Text = ""
LoadPackages

End Sub

Function LoadPackages()
On Error Resume Next
If rs.RecordCount = "0" Then
    Exit Function
End If

List1.Clear
rs.MoveFirst
max = rs.RecordCount

For i = 1 To max
    List1.AddItem rs("Package")
    rs.MoveNext
Next i

rs.MoveFirst
Label16.Caption = "# of packages: " & rs.RecordCount
rs.Close
db.Close
'End of load

End Function


Private Sub Form_Load()
On Error Resume Next
Ontop Me
Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(App.Path & "\Packages.mdb")
Set rs = db.OpenRecordset("Packages", dbOpenTable)
LoadPackages
TreeView1.Nodes.Clear

TreeView1.Nodes.Add , , , "Options", 1
TreeView1.Nodes.Add , , , "Help", 1

TreeView1.Nodes.Add 1, tvwChild, , "Information", 4
TreeView1.Nodes.Add 1, tvwChild, , "Startup", 4
TreeView1.Nodes.Add 1, tvwChild, , "Packages", 4

TreeView1.Nodes.Add 2, tvwChild, , "Licence", 4
TreeView1.Nodes.Add 2, tvwChild, , "About", 4

Dim first, Last, Age, Gender, UserName, Password, Remember, Logo, autoemail As String
first = GetSetting("Accounts", "Profile", "Firstname")
Last = GetSetting("Accounts", "Profile", "Lastname")
Age = GetSetting("Accounts", "Profile", "Age")
Gender = GetSetting("Accounts", "Profile", "Gender")
UserName = GetSetting("Accounts", "Profile", "Username")
Password = "Hidden"
Remember = GetSetting("Accounts", "Profile", "Remember")
Logo = GetSetting("Accounts", "Profile", "Logo")

If Remember = "1" Then
    Remember = "Yes"
Else
    Remember = "No"
End If

If Logo = "1" Then
    Logo = "Yes"
Else
    Logo = "No"
End If
Label1.Caption = "Firstname: " & first
Label2.Caption = "Lastname: " & Last
Label3.Caption = "Age: " & Age
Label4.Caption = "Gender: " & Gender
Label5.Caption = "Username: " & UserName
Label6.Caption = "Password: " & Password
Label9.Caption = "Remember Username: " & Remember
Label10.Caption = "Display Logo: " & Logo

Check1.Value = GetSetting("Accounts", "Profile", "Logo")
Check2.Value = GetSetting("Accounts", "Profile", "Remember")
Check3.Value = GetSetting("Accounts", "Profile", "RememberPasswd")
Check4.Value = GetSetting("Accounts", "Profile", "AutoLogin")
Check5.Value = GetSetting("Accounts", "Profile", "Database")
Check3.Value = GetSetting("Accounts", "Profile", "AutoEmail")
Check6.Value = GetSetting("Accounts", "Profile", "Logging")

Text2.Text = GetSetting("Accounts", "Profile", "DatabaseName")
Dim StartOption As String
StartOption = GetSetting("Accounts", "Profile", "Tab")

If StartOption = "0" Then
    Option1.Value = True
End If

If StartOption = "1" Then
    Option2.Value = True
End If

If StartOption = "2" Then
    Option3.Value = True
End If

If Check5.Value = "1" Then
    Command4.Visible = True
    Text2.Visible = True
    Label11.Visible = True
Else
    Command4.Visible = False
    Text2.Visible = False
    Label11.Visible = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim Op As String

If Option1.Value = True Then
    Op = "0"
ElseIf Option2.Value = True Then
    Op = "1"
ElseIf Option3.Value = True Then
    Op = "2"
End If

SaveSetting "Accounts", "Profile", "Logo", Check1.Value
SaveSetting "Accounts", "Profile", "Remember", Check2.Value
SaveSetting "Accounts", "Profile", "Rememberpasswd", Check3.Value
SaveSetting "Accounts", "Profile", "AutoLogin", Check4.Value
SaveSetting "Accounts", "Profile", "Tab", Op
SaveSetting "Accounts", "Profile", "Database", Check5.Value
SaveSetting "Accounts", "Profile", "DatabaseName", Text2.Text
SaveSetting "Accounts", "Profile", "AutoEmail", Check3.Value
SaveSetting "Accounts", "Profile", "Logging", Check6.Value

FrmAdmin.Label24.Caption = Text2.Text

If Check6.Value = 1 Then
    Open App.Path & "\accountsdb.log" For Append As #1
Else
    Close #1
End If

If Check5.Value = "1" Then
    Dim question As String
    question = MsgBox("Would you like to import the database now?", vbQuestion Or vbYesNo, "Profiles")
    If question = vbYes Then
        FrmAdmin.Text7.Text = Text2.Text
        FrmAdmin.Reload
        Unload Me
    Else
        Unload Me
    End If
End If
End Sub



Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = "2" Then
    Dim answer
    answer = MsgBox("Are you sure you want to remove this item?", vbYesNo Or vbQuestion, "Profiles")
        If answer = vbYes Then
            Set ws = DBEngine.Workspaces(0)
            Set db = ws.OpenDatabase(App.Path & "\Packages.mdb")
            Set rs = db.OpenRecordset("Packages", dbOpenTable)
            rs.MoveFirst
            rs.Move (List1.ListIndex)
            rs.Delete
            List1.RemoveItem (List1.ListIndex)
            Label16.Caption = "# of packages: " & rs.RecordCount
        End If
        
End If

End Sub

Private Sub TreeView1_Click()
On Error Resume Next
'this is to figure out what to show when an option is selected in the treeview
If TreeView1.SelectedItem.Text = "Information" Then

    Image1.Visible = False
    Frame1.Visible = True
    Frame3.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    
ElseIf TreeView1.SelectedItem.Text = "Options" Then

    Frame1.Visible = False
    Image1.Visible = True
    Frame3.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False

ElseIf TreeView1.SelectedItem.Text = "Setup" Then

    Image1.Visible = False
    Frame1.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    
ElseIf TreeView1.SelectedItem.Text = "Startup" Then
    
    Image1.Visible = False
    Frame1.Visible = False
    Frame3.Visible = True
    Frame5.Visible = False
    Frame6.Visible = False

ElseIf TreeView1.SelectedItem.Text = "Main" Then

    Frame1.Visible = False
    Image1.Visible = True
    Frame3.Visible = False
    Frame5.Visible = False
    Frame6.Visible = False
    
ElseIf TreeView1.SelectedItem.Text = "Licence" Then
    Frame1.Visible = False
    Image1.Visible = False
    Frame3.Visible = False
    Frame5.Visible = True
    Frame6.Visible = False

ElseIf TreeView1.SelectedItem.Text = "About" Then
    frmAbout.Show
ElseIf TreeView1.SelectedItem.Text = "Packages" Then
    Frame6.Visible = True
    Frame1.Visible = False
    Image1.Visible = False
    Frame3.Visible = False
    Frame5.Visible = False
End If
'end of selected treeview
End Sub

