VERSION 5.00
Begin VB.Form FrmStart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   480
      Top             =   4680
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "FrmStart.frx":0000
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "FrmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Ontop Me
Dim Setting As String

Setting = GetSetting("Accounts", "Profile", "Logo")

If Setting = "1" Then
    Exit Sub
Else
    FrmMain.Show
    Unload Me
End If

End Sub

Private Sub Timer1_Timer()
FrmMain.Show
Unload Me
End Sub
