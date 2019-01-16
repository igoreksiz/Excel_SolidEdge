Attribute VB_Name = "mdl_VK"
Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub SendPaste()
Const KeyEventF_KeyDOWN = 0
Const KeyEventF_KeyUP = &H2
Const vk_Ctrl = 17
Const vk_V = 86

keybd_event vk_Ctrl, 0, KeyEventF_KeyDOWN, 0 '按下Ctrl键
keybd_event vk_V, 0, KeyEventF_KeyDOWN, 0
'Sleep 500 '延时500毫秒
keybd_event vk_V, 0, KeyEventF_KeyUP, 0
keybd_event vk_Ctrl, 0, KeyEventF_KeyUP, 0 '释放Ctrl键

End Sub

