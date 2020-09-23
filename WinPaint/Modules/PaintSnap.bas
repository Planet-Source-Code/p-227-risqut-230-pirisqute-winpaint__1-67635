Attribute VB_Name = "Paintshot"
'API Calls
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Constants
Const VK_MENU = &H12
Const VK_SNAPSHOT = &H2C
Const KEYEVENTF_KEYUP = &H2
'Sub that gets screen capture
Sub snap1(frm1 As Form)
'Variables
Dim frm As New FrmMain
    ' Presses Alt.
    keybd_event VK_MENU, 0, 0, 0
    'DoEvents
    DoEvents
    ' Presses Print Scrn.
    keybd_event VK_SNAPSHOT, 1, 0, 0
    'DoEvents
    DoEvents
    ' Releases Alt.
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    'DoEvents
    DoEvents
    'Set form picture to screen caption
    frm1.Image1.Picture = Clipboard.GetData(vbCFBitmap)
End Sub
