Attribute VB_Name = "Gettok"
Option Explicit

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Function ReadINI(Section As String, KeyName As String, FileName As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function


Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Function GetTok(ByVal strVal As String, intIndex As Integer, strDelimiter As String) As String

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
Public Sub Ontop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub Notontop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
