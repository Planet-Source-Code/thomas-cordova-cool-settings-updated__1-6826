Attribute VB_Name = "Module1"
Public Sub Off()
Unload Me
End Sub
Public Sub Shutdown()
Shell "rundll32 user,tilechildwindows"
End Sub
Public Sub Restart()
Shell "rundll32 shell32,SHExitWindowsEx 2"
End Sub
Public Sub Control()
Shell "rundll32 shell32,Control_RunDLL", vbNormalFocus
End Sub
Public Sub Display()
Shell "rundll32 shell32,Control_RunDLL desk.cpl"
End Sub
Public Sub MouseSwap()
Shell "rundll32 user,swapmousebutton"
End Sub
Public Sub PrintTest()
Shell "rundll32 msprint2.dll,RUNDLL_PrintTestPage"
End Sub
Public Sub Hardware()
Shell "rundll32 sysdm.cpl,InstallDevice_Rundll"
End Sub
Public Sub NetworkDrive()
Shell "rundll32 user,wnetconnectdialog"
End Sub
Public Sub Format()
Shell "rundll32 user,wnetconnectdialog"
    End Sub
Public Sub Copy()

    If Text1.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText Text1.SelText
    End If
End Sub
Public Sub Cut()
If Text1.SelText = "" Then
        Exit Sub
    Else
        Clipboard.Clear
        Clipboard.SetText Text1.SelText
        Text1.SelText = ""
    End If
End Sub
Public Sub Paste()
Text1.SelText = Clipboard.GetText
End Sub
Public Sub SelectAll()
Text1.SelStart = 0
 Text1.SelLength = Len(Text1.Text)
 Text1.SetFocus
End Sub
Public Sub ClearText()
 Text1.SelStart = 0
 Text1.SelLength = Len(Text.Text)
 SendMessage Text1.hWnd, WM_CLEAR, 0, 0
End Sub
Public Sub HideTaskbar()
hwnd1 = FindWindow("Shell_traywnd", "")
Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub
Public Sub ShowTaskbar()
Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub
Public Function HideClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 0
End Function
Public Function ShowClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
ShowWindow Handle&, 1
End Function
Public Function DeleteClock()
Dim FindClass As Long, FindParent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", vbNullString)
FindParent& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
Handle& = FindWindowEx(FindParent&, 0, "TrayClockWClass", vbNullString)
SendMessage Handle&, WM_DESTROY, 0, 0
End Function
Public Function HideSystemTray()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 0
End Function
Public Function ShowSystemTray()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
ShowWindow Handle&, 1
End Function
Public Function DeleteSystemTray()
Dim FindClass As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
Handle& = FindWindowEx(FindClass&, 0, "TrayNotifyWnd", vbNullString)
SendMessage Handle&, WM_DESTROY, 0, 0
End Function
Public Function HidePrograms()
Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
ShowWindow Handle&, 0
End Function
Public Function ShowPrograms()
Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
ShowWindow Handle&, 1
End Function
Public Function DeletePrograms()
Dim FindClass As Long, FindClass2 As Long, Parent As Long, Handle As Long
FindClass& = FindWindow("Shell_TrayWnd", "")
FindClass2& = FindWindowEx(FindClass&, 0, "ReBarWindow32", vbNullString)
Parent& = FindWindowEx(FindClass2&, 0, "MSTaskSwWClass", vbNullString)
Handle& = FindWindowEx(Parent&, 0, "SysTabControl32", vbNullString)
SendMessage Handle&, WM_DESTROY, 0, 0
End Function
Public Function BlackOut(TheForm As Form)
NotOnTop TheForm
ShowTaskbar
ShowWindowsToolBar
TheForm.BorderStyle = 3
TheForm.Caption = "Form"
Screen.MousePointer = vbArrow
TheForm.BackColor = &H8000000A
TheForm.Width = Screen.Width / 2
TheForm.Height = Screen.Height / 2
TheForm.Left = Screen.Width / 2 - TheForm.Width / 2
TheForm.Top = Screen.Height / 2 - TheForm.Height / 2
UnPreventFromClosing
EnableCtrlAltDel
End Function
Public Function DisableCtrlAltDel()
Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Function
Public Function EnableCtrlAltDel()
Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Function

