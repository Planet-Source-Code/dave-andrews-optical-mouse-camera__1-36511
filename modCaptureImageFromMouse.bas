Attribute VB_Name = "modCaptureImageFromMouse"
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDeviceGammaRamp Lib "gdi32" (ByVal hDC As Long, lpv As Any) As Long
Private Declare Function InSendMessage Lib "user32" () As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Function CaptureImage(STAT As Object, IMGARY() As Long) As Boolean
Dim i As Long
Dim j As Long
Dim RET As Long
Dim dX As Long
Dim xY As Long
Dim But As Long
RET = GetDeviceCaps(STAT.hDC, 1)
Do While i < 12000
    Call mouse_event(1, dX, dy, But, j)
    If But = 0 Then
        IMGARY(i) = dX + dy + j
        STAT.Caption = "Capturing: " & i & "/ 12000"
    End If
    DoEvents
    i = i + IMGARY(i) + 12
Loop
CaptureImage = True
End Function


