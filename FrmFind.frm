VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Client"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6570
   Icon            =   "FrmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2280
      Top             =   0
   End
   Begin VB.CheckBox Check2 
      Caption         =   "If exact match, Open profile"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   14
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   960
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
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
            Picture         =   "FrmFind.frx":65AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2295
      Left            =   0
      TabIndex        =   13
      Top             =   2640
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4048
      Sorted          =   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "All"
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search By"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1815
      Begin VB.OptionButton Option4 
         Caption         =   "E-mail"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Lastname"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Firstname"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Account Number"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Search Results:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Direction:"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Search:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset

Dim i As Long
Dim max As Long
Dim cnt As Integer
Dim SelectedOption As Integer

Dim Text As String
Dim Gender As String


Private Sub Command1_Click()
FrmLoad.Show
Timer3.Enabled = True
End Sub

Private Sub Command2_Click()
Timer1.Enabled = True
FrmAdmin.Reload
End Sub


Private Sub Form_Load()
Ontop Me
ListView1.ColumnHeaders.Clear
ListView1.ColumnHeaders.Add , , "ID", ListView1.Width / 8, , 1
ListView1.ColumnHeaders.Add , , "Firstname", ListView1.Width / 4
ListView1.ColumnHeaders.Add , , "Lastname", ListView1.Width / 4
ListView1.ColumnHeaders.Add , , "E-mail", ListView1.Width / 2
ListView1.View = lvwReport

    Combo1.AddItem ("All")
    Combo1.AddItem ("Up")
    Combo1.AddItem ("Down")
    OpenDB
End Sub


Public Function OpenDB()

Set ws = DBEngine.Workspaces(0)
Set db = ws.OpenDatabase(FrmAdmin.Text7.Text)
Set rs = db.OpenRecordset("Accounts", dbOpenTable)

End Function

Private Sub Form_Unload(Cancel As Integer)
Timer1.Enabled = True
FrmAdmin.Reload

End Sub


Private Sub ListView1_DblClick()
On Error Resume Next
Print #1, Time & " - Displaying found clients personal information."

OpenDB
rs.MoveFirst
max = rs.RecordCount
For i = 1 To max
 
 If ListView1.SelectedItem.Text = FrmAdmin.ListView1.ListItems(i).Text Then
    rs.Move i - 1
    FrmAdmin.Text9.Text = i - 1
    GoTo LoadFrm
 End If

Next i

LoadFrm:
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


End Sub

Private Sub Option1_Click()
Check2.Enabled = True
End Sub

Private Sub Option2_Click()
Check2.Value = 0
Check2.Enabled = False
End Sub

Private Sub Option3_Click()
Check2.Value = 0
Check2.Enabled = False
End Sub

Private Sub Option4_Click()
Check2.Enabled = True
End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    Command1_Click
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

Public Function GetTok(ByVal strVal As String, intIndex As Integer, strDelimiter As String) As String

        Dim strSubString() As String
        Dim intIndex2 As Integer
        Dim i As Integer
        Dim intDelimitLen As Integer
        intIndex2 = 1
        i = 0
        intDelimitLen = Len(strDelimiter)
        Do While intIndex2 > 0
            ReDim Preserve strSubString(i + 1)
            intIndex2 = InStr(1, strVal, strDelimiter)
            If intIndex2 > 0 Then
                strSubString(i) = Mid(strVal, 1, (intIndex2 - 1))
                strVal = Mid(strVal, (intIndex2 + intDelimitLen), Len(strVal))
            Else
                strSubString(i) = strVal
            End If
            i = i + 1
        Loop
        If intIndex > (i + 1) Or intIndex < 1 Then
            GetTok = ""
        Else
            GetTok = strSubString(intIndex - 1)
        End If
End Function

Private Sub Timer2_Timer()
Text1.SetFocus
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Print #1, Time & " - Searching for client...."
If Text1.Text = "" Then
    Text1.SetFocus
    Timer3.Enabled = False
    Exit Sub
End If
FrmLoad.Show
If Option1.Value = True Then
    SelectedOption = 1
End If

If Option2.Value = True Then
    SelectedOption = 2
End If

If Option3.Value = True Then
    SelectedOption = 3
End If

If Option4.Value = True Then
    SelectedOption = 4
End If

Me.Height = 5505
OpenDB
ListView1.ListItems.Clear
rs.MoveLast
rs.MoveFirst
cnt = 1
max = rs.RecordCount
FrmLoad.PB.max = rs.RecordCount
FrmLoad.PB1.max = rs.RecordCount + max
Open App.Path & "/search.txt" For Output As #2

For i = 1 To max
    Print #2, rs("ID") & "%" & rs("firstname") & "%" & rs("lastname") & "%" & rs("email")
    FrmLoad.PB.Value = i
    rs.MoveNext
    FrmLoad.PB1.Value = i
Next i
FrmLoad.PB.Value = 0
Close #2


If Check1.Value = "1" Then
FrmLoad.Label4.Caption = "Displaying Results..."

    Open App.Path & "/search.txt" For Input As #2
    
        Do While Not EOF(2)
            Line Input #2, Text
                If UCase(GetTok(Text, SelectedOption, "%")) = Text1.Text Then
                        ListView1.ListItems.Add , , GetTok(Text, 1, "%"), 1
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 2, "%")
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 3, "%")
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 4, "%")
                        FrmLoad.PB.Value = cnt
                        FrmLoad.PB1.Value = rs.RecordCount + cnt
                        cnt = cnt + 1
                End If
    Loop
    Timer3.Enabled = False
Else
FrmLoad.Label4.Caption = "Displaying Results..."

    Open App.Path & "/search.txt" For Input As #2
    
        Do While Not EOF(2)
            Line Input #2, Text
                If LCase(GetTok(Text, SelectedOption, "%")) = Text1.Text Then
                        ListView1.ListItems.Add , , GetTok(Text, 1, "%"), 1
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 2, "%")
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 3, "%")
                        ListView1.ListItems(cnt).ListSubItems.Add , , GetTok(Text, 4, "%")
                        FrmLoad.PB.Value = cnt
                        FrmLoad.PB1.Value = rs.RecordCount + cnt
                        cnt = cnt + 1
                End If
    Loop
Timer3.Enabled = False
End If

Close #2

rs.Close
db.Close
Label3.Caption = "Found " & ListView1.ListItems.Count & " results."
FrmLoad.PB1.Value = 0
Print #1, Time & " - Found " & ListView1.ListItems.Count & " results."
Kill App.Path & "/search.txt"
On Error Resume Next
If Check2.Value = 1 Then
    ListView1.ListItems.Item(1).ForeColor = vbBlue
    ListView1_DblClick
End If
FrmLoad.Hide
End Sub
