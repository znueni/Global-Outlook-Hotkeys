Imports System, System.Windows.Forms, System.Diagnostics, System.Runtime.InteropServices, System.Collections.Generic, System.Reflection
#Region "Assemblyinfo"
<Assembly: AssemblyTitle("Global Outlook Hotkeys")>
<Assembly: AssemblyCompany("")>
<Assembly: AssemblyProduct("")>
<Assembly: AssemblyCopyright("Copyright © Sebastian 2018")>
<Assembly: AssemblyTrademark("")>
<Assembly: ComVisible(False)>
<Assembly: Guid("6e595e5c-f4bd-4880-bfba-3a437d8cdee5")>
<Assembly: AssemblyVersion("1.0.0.0")>
<Assembly: AssemblyFileVersion("1.0.0.0")>
<Assembly: AssemblyDescription("Win+F12: Outlook, Win+F11: Outlook calendar, CTRL+Win+F12 or CTRL+Win+SPACE or Win+NUM0: New Mail, SHIFT+Win+F12 or SHIFT+Win+SPACE or Win+NUM1: New Appointment, CTRL+SHIFT+ALT+Win+F12: Exit me")>
#End Region

Public Class Globalhotkeys
#Region "Win API imports"
    'Enumerate Windows:
    Private Delegate Function EnumWindowsCallback(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
    Private Declare Function EnumWindows Lib "user32" (ByVal Adress As EnumWindowsCallback, ByVal y As Integer) As Integer
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpWindowText As String, ByVal cch As Integer) As Integer
    Private Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    Private Const SW_SHOWMINIMIZED As Short = 2
    Private Const SW_SHOWMAXIMIZED As Short = 3
    Private Const SW_SHOWNORMAL As Short = 1

    Private Structure POINTAPI
        Public x As Integer
        Public y As Integer
    End Structure
    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
    Private Structure WINDOWPLACEMENT
        Public Length As Integer
        Public flags As Integer
        Public showCmd As Integer
        Public ptMinPosition As POINTAPI
        Public ptMaxPosition As POINTAPI
        Public rcNormalPosition As RECT
    End Structure
    Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As IntPtr, ByRef lpwndpl As WINDOWPLACEMENT) As Integer
    '  Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Integer, ByVal lpEnumFunc As EnumWindowsCallback, ByVal lParam As Integer) As Integer
    ' Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer


    'Hotkeys
    Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As Integer, ByVal vk As Integer) As Integer
    Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As IntPtr, ByVal id As Integer) As Integer
    Public Const WM_HOTKEY As Integer = &H312
    Enum KeyModifier
        None = 0
        Alt = &H1
        Control = &H2
        Shift = &H4
        Winkey = &H8
    End Enum

    'Move to front
    Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer
    Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

#End Region

    Const WP_F11_CALENDAR As Byte = 11
    Const WP_F12_OUTLOOK As Byte = 12
    Const WP_F12_ALT_EXIT As Byte = 3
    Const WP_NUM0_NMAIL As Byte = 5
    Const WP_NUM1_NAPPOINTMENT As Byte = 6


    Public Shared Sub Main()
        Dim f As New HotKeyform()
        Application.DoEvents()

        'These are the hotkeys we define to be registered.
        'The WP_ constants refer to our functionality, we define. If you want to assign an additional hotkey to a functionality, just add another RegisterHotKey with the selected WP_ constant.
        'If you create a new functionality, create a new WP_ constant and add a Hotkey with RegisterHotKey. The WndProc function below contains the code for each WP_ action.
        RegisterHotKey(f.Handle, WP_F11_CALENDAR, KeyModifier.Winkey, Keys.F11)
        RegisterHotKey(f.Handle, WP_F12_OUTLOOK, KeyModifier.Winkey, Keys.F12)
        RegisterHotKey(f.Handle, WP_NUM0_NMAIL, KeyModifier.Winkey Or KeyModifier.Control, Keys.F12)
        RegisterHotKey(f.Handle, WP_NUM1_NAPPOINTMENT, KeyModifier.Winkey Or KeyModifier.Shift, Keys.F12)
        RegisterHotKey(f.Handle, WP_F12_ALT_EXIT, KeyModifier.Winkey Or KeyModifier.Alt Or KeyModifier.Control Or KeyModifier.Shift, Keys.F12)
        RegisterHotKey(f.Handle, WP_NUM0_NMAIL, KeyModifier.Winkey Or KeyModifier.Control, Keys.Space)
        RegisterHotKey(f.Handle, WP_NUM1_NAPPOINTMENT, KeyModifier.Winkey Or KeyModifier.Shift, Keys.Space)
        RegisterHotKey(f.Handle, WP_NUM0_NMAIL, KeyModifier.Winkey, Keys.NumPad0)
        RegisterHotKey(f.Handle, WP_NUM1_NAPPOINTMENT, KeyModifier.Winkey, Keys.NumPad1)
        Application.Run()
    End Sub


    Public Class HotKeyform
        Inherits Form
        Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
            If m.Msg = WM_HOTKEY Then
                If m.WParam.ToInt32 = WP_F12_ALT_EXIT Then
                    'Unregister all hotkeys and quit
                    UnregisterHotKey(Handle, WP_F11_CALENDAR)
                    UnregisterHotKey(Handle, WP_F12_OUTLOOK)
                    UnregisterHotKey(Handle, WP_F12_ALT_EXIT)
                    UnregisterHotKey(Handle, WP_NUM0_NMAIL)
                    UnregisterHotKey(Handle, WP_NUM1_NAPPOINTMENT)
                    Application.Exit() : Exit Sub
                End If

                'Debug.WriteLine("WM_HOTKEY + " & m.WParam.ToInt32)

                'If outlook is not running, dont do anything
                Dim P() As Process = Process.GetProcessesByName("outlook")
                If P.Length <> 1 Then Exit Sub


                'Determine the window handles
                'But only do this for wparam 11,12 (= actions WP_F11_CALENDAR & WP_F12_OUTLOOK)
                Dim CalendarHandle As Integer = 0
                Dim ExplorerHandle As Integer = 0
                If m.WParam.ToInt32 >= 11 Then
                    Dim OutlookPID As Integer
                    GetWindowThreadProcessId(P(0).MainWindowHandle, OutlookPID)
                    EnumWindows(Function(hwnd, lparam)
                                    If IsWindowVisible(hwnd) Then 'sichtbar
                                        Dim pid As Integer = 0
                                        GetWindowThreadProcessId(hwnd, pid)
                                        If pid = OutlookPID Then 'outlook window
                                            Dim text As New String(" "c, 1024)
                                            GetWindowText(hwnd, text, 1024)
                                            text = text.Trim(vbNullChar, " "c)
                                            If text > "" AndAlso text.EndsWith("Outlook") Then
                                                ' MsgBox($"Window {hwnd} of PID {pid} with text {vbCrLf}{text}{vbCrLf}Outlook PID = {OutlookPID}")
                                                If (text.StartsWith("Kalender") Or text.StartsWith("Calendar")) Then
                                                    CalendarHandle = hwnd
                                                Else
                                                    ExplorerHandle = hwnd
                                                End If
                                            End If
                                        End If
                                    End If
                                    Return True
                                End Function, 0)
                End If

                'Code for each WP_ function we defined:
                Select Case m.WParam
                    Case WP_F11_CALENDAR
                        If CalendarHandle > 0 Then ShowWindowbyHandle(CalendarHandle)

                    Case WP_F12_OUTLOOK
                        If ExplorerHandle > 0 Then
                            ShowWindowbyHandle(ExplorerHandle)
                        ElseIf CalendarHandle > 0 Then
                            ShowWindowbyHandle(CalendarHandle)
                        End If

                    Case WP_NUM0_NMAIL
                        Process.Start(P(0).MainModule.FileName, "/c ipm.note")

                    Case WP_NUM1_NAPPOINTMENT
                        Process.Start(P(0).MainModule.FileName, "/c ipm.appointment")
                End Select
            End If
            MyBase.WndProc(m)
        End Sub

        Private Sub ShowWindowbyHandle(handle As Integer)
            'Helper function to bring a window to front.
            Dim wp As WINDOWPLACEMENT
            wp.Length = System.Runtime.InteropServices.Marshal.SizeOf(wp)
            GetWindowPlacement(handle, wp)
            If wp.showCmd = SW_SHOWMINIMIZED Then ShowWindow(handle, SW_SHOWMAXIMIZED)
            SetForegroundWindow(handle)
        End Sub
    End Class


End Class
