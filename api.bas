Attribute VB_Name = "api"
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Private Declare Function SetForegroundWindow& Lib "user32" (ByVal hwnd As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Declare Function GetAsyncKeyState Lib "user32" (ByVal dwMessage As Long) As Integer

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE


Public Active As Boolean

Public Function Amazing_Magic_THING(sName As String, sKeyCode As String) As Boolean

On Error Resume Next

Dim sRandomSeed As String
Dim X As Long
Dim KeyCounter As Long
Dim PrimaryLetter As Long
Dim DecodedKey As String
Dim AntiCodedLetter As Long
Dim sBuffer As String
    
sRandomSeed = Left$(sKeyCode, InStr(sKeyCode, "-") - 1)
sName = UCase$(sName)
sKeyCode = Right$(sKeyCode, 10)
KeyCounter = 1

For X = 1 To Len(sName)
    If Asc(Mid$(sName, X, 1)) >= 65 And Asc(Mid$(sName, X, 1)) <= 90 Then sBuffer = sBuffer & Mid$(sName, X, 1)
Next X
sName = sBuffer

Do While Len(sName) < 10
    sName = sName + Chr$(65)
Loop

For X = 1 To Len(sKeyCode)
    PrimaryLetter = Asc(Mid$(sKeyCode, X, 1))
    AntiCodedLetter = PrimaryLetter - CInt(Mid$(sRandomSeed, KeyCounter, 1))

    If PrimaryLetter = 48 Then 'zero
        DecodedKey = DecodedKey + Mid$(sName, X, 1) 'Take the corresponding letter from the name
    Else
        DecodedKey = DecodedKey + Chr$(AntiCodedLetter)
    End If
        KeyCounter = KeyCounter + 1
        If KeyCounter > 4 Then KeyCounter = 1
Next X

If DecodedKey = Left$(sName, 10) Then
    Amazing_Magic_THING = True
Else
    Amazing_Magic_THING = False
End If

End Function

Public Sub MakeNormal(hwnd As Long)
    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Public Sub MakeTopMost(hwnd As Long)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub
