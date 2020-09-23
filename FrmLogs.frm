VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmLogs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log Viewer"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   Icon            =   "FrmLogs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "View Logs"
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   12303
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"FrmLogs.frx":65AA
      End
   End
End
Attribute VB_Name = "FrmLogs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Ontop Me
Close #1
Open App.Path & "\accountsdb.log" For Input As #3
    Dim TxT As String
    Do While Not EOF(3)
        Line Input #3, TxT
            RichTextBox1.SelText = TxT & vbCrLf
    Loop
Close #3
Open App.Path & "\accountsdb.log" For Append As #1
Print #1, ""
End Sub
