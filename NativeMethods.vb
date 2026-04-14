Imports System
Imports System.Runtime.InteropServices

Friend NotInheritable Class NativeMethods
    Private Sub New()
    End Sub

    Public Const WM_HOTKEY As Integer = &H312
    Public Const INPUT_MOUSE As UInteger = 0UI
    Public Const INPUT_KEYBOARD As UInteger = 1UI
    Public Const MOUSEEVENTF_MOVE As UInteger = &H1UI
    Public Const MOUSEEVENTF_LEFTDOWN As UInteger = &H2UI
    Public Const MOUSEEVENTF_LEFTUP As UInteger = &H4UI
    Public Const MOUSEEVENTF_ABSOLUTE As UInteger = &H8000UI
    Public Const MOUSEEVENTF_VIRTUALDESK As UInteger = &H4000UI
    Public Const KEYEVENTF_EXTENDEDKEY As UInteger = &H1UI
    Public Const KEYEVENTF_KEYUP As UInteger = &H2UI
    Public Const KEYEVENTF_SCANCODE As UInteger = &H8UI
    Public Const MAPVK_VK_TO_VSC As UInteger = 0UI
    Public Const SM_XVIRTUALSCREEN As Integer = 76
    Public Const SM_YVIRTUALSCREEN As Integer = 77
    Public Const SM_CXVIRTUALSCREEN As Integer = 78
    Public Const SM_CYVIRTUALSCREEN As Integer = 79

    <StructLayout(LayoutKind.Sequential)>
    Public Structure INPUT
        Public type As UInteger
        Public unionData As InputUnion
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure InputUnion
        <FieldOffset(0)>
        Public mi As MOUSEINPUT

        <FieldOffset(0)>
        Public ki As KEYBDINPUT

        <FieldOffset(0)>
        Public hi As HARDWAREINPUT
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As UInteger
        Public dwFlags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure KEYBDINPUT
        Public wVk As UShort
        Public wScan As UShort
        Public dwFlags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure HARDWAREINPUT
        Public uMsg As UInteger
        Public wParamL As UShort
        Public wParamH As UShort
    End Structure

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As UInteger, vk As UInteger) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetCursorPos(x As Integer, y As Integer) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SendInput(nInputs As UInteger, <MarshalAs(UnmanagedType.LPArray)> pInputs() As INPUT, cbSize As Integer) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=False)>
    Public Shared Sub mouse_event(dwFlags As UInteger, dx As UInteger, dy As UInteger, dwData As UInteger, dwExtraInfo As UIntPtr)
    End Sub

    <DllImport("user32.dll", SetLastError:=False)>
    Public Shared Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As UIntPtr)
    End Sub

    <DllImport("user32.dll", SetLastError:=False)>
    Public Shared Function MapVirtualKey(uCode As UInteger, uMapType As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", SetLastError:=False)>
    Public Shared Function GetSystemMetrics(nIndex As Integer) As Integer
    End Function
End Class
