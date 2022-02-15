VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "FuckComputer for Windows"
   ClientHeight    =   3660
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Î¢ÈíÑÅºÚ Light"
      Size            =   12
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   6360
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   6135
   End
   Begin VB.Label lblConsole 
      BackStyle       =   0  'Transparent
      Caption         =   "The Fake Console"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FuckComputer ported to wInDoWs by YidaozhanYa

'Let everyone enjoy the fun of fucking. --Chi_Tang

Private Sub Form_Activate()
List1.Visible = False
lblConsole.Caption = "Let everyone enjoy the fun of fucking."
DoEvents
CheckInternet
PrintW "Checking git..."
Shell "cmd /c git --version > %TEMP%\fuckcomputer.tmp", vbHide
Sleep 100
Open Environ("Temp") & "\fuckcomputer.tmp" For Input As #1
    Dim TmpVal As String
    Input #1, TmpVal
    If InStr(TmpVal, "git version") = False Then
        MsgBox "Git is not found!", vbOKOnly + vbCritical, "FuckComputer for Windows"
    End If
Close #1
Set objJson = JSON.parse(JsonString)
lblConsole.Caption = "Welcome to FuckComputer " & CStr(objJson.Count) & " in 1 (Windows ver)!"
List1.Clear
Dim y As Integer
y = 0
For Each x In objJson.keys
    y = y + 1
    List1.AddItem "(" & y & ") " & x
Next x
List1.Visible = True
DoEvents
End Sub

Private Sub CheckInternet()
On Error GoTo GreatConnection
PrintW "Checking Internet..."
Dim xhr As Object
Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
xhr.open "GET", "https://ssmzhn.github.io/FuckComputer.json", True ' Async
xhr.send
xhr.waitForResponse 10
If xhr.Status = 200 Then
    JsonString = xhr.responseText
    Set xhr = Nothing
    Exit Sub
Else
    GoTo GreatConnection
End If
GreatConnection:
    MsgBox "Great connection!", vbOKOnly + vbCritical, "FuckComputer for Windows"
End Sub

Private Sub List1_Click()
            Dim TmpVal As String
Dim choice As String
List1.Visible = False
choice = Split(List1.Text, ") ")(1)
PrintW "Cloing respository..."
ShellAndWait "cmd /c cd " & Chr(34) & App.Path & Chr(34) & " && git clone https://github.com/FuckComputer/" & choice
Select Case objJson(choice)
    Case "cpp"
        PrintW "Checking g++..."
        Shell "cmd /c g++ --version > %TEMP%\fuckcomputer.tmp", vbHide
        Sleep 100
        Open Environ("Temp") & "\fuckcomputer.tmp" For Input As #1
            Input #1, TmpVal
        Close #1
        If InStr(TmpVal, "g++ (GCC)") = False Then
            MsgBox "G++ is not found!", vbOKOnly + vbCritical, "FuckComputer for Windows"
        Else
            ShellAndWait "cmd /c g++ " & Chr(34) & App.Path & "\" & choice & "\src\main.cpp" & Chr(34)
            Shell "cmd /k " & Chr(34) & App.Path & "\" & choice & "\a.exe" & Chr(34), vbNormalFocus
        End If
    Case "nim"
        PrintW "Checking nim..."
        Shell "cmd /c nim --version > %TEMP%\fuckcomputer.tmp", vbHide
        Sleep 100
        Open Environ("Temp") & "\fuckcomputer.tmp" For Input As #1
            Input #1, TmpVal
        Close #1
        If InStr(TmpVal, "Nim Compiler") = False Then
            MsgBox "Nim is not found!", vbOKOnly + vbCritical, "FuckComputer for Windows"
        Else
            ShellAndWait "cmd /c nim compile -d:release " & Chr(34) & App.Path & "\" & choice & "\src\main.nim" & Chr(34)
            ShellAndWait "cmd /c rename " & Chr(34) & App.Path & "\" & choice & "\src\main" & Chr(34) & " main.exe"
            Shell "cmd /k " & Chr(34) & App.Path & "\" & choice & "\src\main.exe" & Chr(34), vbNormalFocus
        End If
    Case "python"
        PrintW "Checking Python..."
        Shell "cmd /c python --version > %TEMP%\fuckcomputer.tmp", vbHide
        Sleep 100
        Open Environ("Temp") & "\fuckcomputer.tmp" For Input As #1
            Input #1, TmpVal
        Close #1
        If InStr(TmpVal, "Python ") = False Then
            MsgBox "Python is not found!", vbOKOnly + vbCritical, "FuckComputer for Windows"
        Else
            Shell "cmd /k python " & Chr(34) & App.Path & "\" & choice & "\src\main.py" & Chr(34), vbNormalFocus
        End If
End Select
End
End Sub

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
