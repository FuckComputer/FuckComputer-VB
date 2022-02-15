Attribute VB_Name = "theShitCodes"
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public JsonString As String, objJson As Object

Public Sub PrintW(Text As String)
frmMain.lblConsole.Caption = frmMain.lblConsole.Caption & vbCrLf & Text
End Sub

Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub
