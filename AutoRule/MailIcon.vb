Option Explicit On

Imports System.Runtime.InteropServices

Module MaiIcon
    ' Entry Point is SetNewMailIcon.

    Public Const WUM_RESETNOTIFICATION As Long = &H407

    'Required Public constants, types & declares for the Shell_Notify API method
    Public Const NIM_ADD As Long = &H0
    Public Const NIM_MODIFY As Long = &H1
    Public Const NIM_DELETE As Long = &H2

    Public Const NIF_ICON As Long = &H2 'adding an ICON
    Public Const NIF_TIP As Long = &H4 'adding a TIP
    Public Const NIF_MESSAGE As Long = &H1 'want return messages

    ' Structure needed for Shell_Notify API
    Structure NOTIFYICONDATA
        Dim cbSize As Integer
        Dim hWnd As IntPtr
        Dim uID As Integer
        Dim uFlags As Integer
        Dim uCallbackMessage As Integer
        Dim hIcon As IntPtr
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Dim szTip As String
        Dim dwState As String
        Dim dwStateMask As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Dim szInfo As String
        Dim uTimeout As Integer ' ignore the uVersion union
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=64)> Dim szInfoTitle As String
        Dim dwInfoFlags As Integer
        Dim guidItem As Guid
    End Structure


    <Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
    Private Function SendMessage(ByVal hWnd As IntPtr, _
                                 ByVal msg As Integer, ByVal wParam As Integer, _
                                 ByVal lParam As Integer) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("shell32.dll", SetLastError:=True)> _
    Private Function Shell_NotifyIcon(ByVal dwMessage As Long, ByVal lpData As NOTIFYICONDATA) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
    Private Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function

    <Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)> _
    Private Function FindWindowEx(ByVal hwndParent As IntPtr, _
                                  ByVal hwndChildAfter As IntPtr, ByVal lpszClass As String, _
                                  ByVal lpszWindow As String) As IntPtr
    End Function


    Sub SetNewMailIcon()
        Dim hResult As Long

        Dim OLHwnd As IntPtr = FindWindow("rctrl_renwnd32", "")
        Dim OLEngine As IntPtr = FindWindow("WMS Notif Engine:Dispatch Window Class", "W")

        'Dim Calcu As IntPtr = FindWindow(Nothing, "Calculator")
        'Dim editx As IntPtr = FindWindowEx(Calcu, IntPtr.Zero, "edit", Nothing)

        'SendMessage(OLEngine, &HC1FC, 0, 0)
        SendMessage(OLEngine, &HC1FA, 0, 0)
        SendMessage(OLHwnd, &H44E, 0, 0)

        hResult = NewMailIcon(OLHwnd)
        hResult = NewMailIcon(OLEngine)

    End Sub

    Private Function NewMailIcon(ByVal hwnd As Long) As Boolean
        Dim pShell_Notify As New NOTIFYICONDATA
        Dim hResult As Long

        'setup the Shell_Notify structure
        pShell_Notify.cbSize = Len(pShell_Notify)
        pShell_Notify.hWnd = hwnd
        pShell_Notify.uID = 0

        ' Remove it from the system tray and catch result
        hResult = Shell_NotifyIcon(NIM_ADD, pShell_Notify)
        If (hResult) Then
            NewMailIcon = True
        Else
            NewMailIcon = False
        End If
    End Function
End Module
