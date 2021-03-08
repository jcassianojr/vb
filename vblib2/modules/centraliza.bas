Attribute VB_Name = "centraliza"
Option Explicit
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long                               ' This is +1 (right - left = width)
    Bottom As Long                               ' This is +1 (bottom - top = height)
End Type

Private Type MONITORINFO
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Private Type MONITORINFOEX
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
    szDevice As String * 32
End Type

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function GetMonitorInfoEx Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFOEX) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long

Public Sub Center(ByRef frm As Form)
    Dim h As Long
    On Error Resume Next
    h = MonitorHandleForHwnd(frm.hWnd)
    CenterWindowOnMonitor frm.hWnd, h
End Sub

Public Function MonitorHandleForHwnd(hWnd As Long) As Long
    Const MONITOR_DEFAULTTONULL = &H0
    MonitorHandleForHwnd = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONULL)
End Function

Public Sub CenterWindowOnMonitor(hWnd As Long, hMonitor As Long, Optional bUseWorkingArea As Boolean = False)
    Dim lLeft As Long
    Dim lTop As Long
    '
    If MonitorWidthPx(hMonitor) = 0 Then Exit Sub ' Bad monitor.
    If bUseWorkingArea Then
        lLeft = (MonitorWidthWorkPx(hMonitor) - WindowWidthPx(hWnd)) \ 2
        lTop = (MonitorHeightWorkPx(hMonitor) - WindowHeightPx(hWnd)) \ 2
    Else
        lLeft = (MonitorWidthPx(hMonitor) - WindowWidthPx(hWnd)) \ 2
        lTop = (MonitorHeightPx(hMonitor) - WindowHeightPx(hWnd)) \ 2
    End If
    PositionWindowOnMonitor hWnd, hMonitor, lLeft, lTop
End Sub

Public Sub PositionWindowOnMonitor(hWnd As Long, hMonitor As Long, ByVal lLeftPixel As Long, ByVal lTopPixel As Long)
    ' This can be used to position windows on other programs so long as you have the hWnd.
    Dim lHeight As Long
    Dim lWidth As Long
    '
    lHeight = WindowHeightPx(hWnd)
    lWidth = WindowWidthPx(hWnd)
    '
    lLeftPixel = lLeftPixel + MonitorLeftPx(hMonitor)
    lTopPixel = lTopPixel + MonitorTopPx(hMonitor)
    '
    MoveWindow hWnd, lLeftPixel, lTopPixel, lWidth, lHeight, 1&
End Sub

Public Function MonitorWidthPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorWidthPx(MonitorWidthPx(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorWidthPx = uMonInfo.rcMonitor.Right - uMonInfo.rcMonitor.Left
End Function

Public Function MonitorHeightPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorHeightPx(MonitorWidthPx(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorHeightPx = uMonInfo.rcMonitor.Bottom - uMonInfo.rcMonitor.Top
End Function

Public Function MonitorWidthWorkPx(hMonitor As Long) As Long
    ' The "Work" verions exclude the taskbar.
    ' If you just have the number, do: MonitorWidthWorkPx(MonitorWidthPx(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorWidthWorkPx = uMonInfo.rcWork.Right - uMonInfo.rcWork.Left
End Function

Public Function MonitorHeightWorkPx(hMonitor As Long) As Long
    ' The "Work" verions exclude the taskbar.
    ' If you just have the number, do: MonitorHeightWorkPx(MonitorWidthPx(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorHeightWorkPx = uMonInfo.rcWork.Bottom - uMonInfo.rcWork.Top
End Function

Public Function WindowHeightPx(hWnd As Long) As Long
    Dim r As RECT
    GetWindowRect hWnd, r
    WindowHeightPx = r.Bottom - r.Top
End Function

Public Function WindowWidthPx(hWnd As Long) As Long
    Dim r As RECT
    GetWindowRect hWnd, r
    WindowWidthPx = r.Right - r.Left
End Function

Public Function MonitorLeftPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorLeftPx(MonitorHandle(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorLeftPx = uMonInfo.rcMonitor.Left
End Function

Public Function MonitorRightPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorRightPx(MonitorHandle(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorRightPx = uMonInfo.rcMonitor.Right
End Function

Public Function MonitorTopPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorTopPx(MonitorHandle(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorTopPx = uMonInfo.rcMonitor.Top
End Function

Public Function MonitorBottomPx(hMonitor As Long) As Long
    ' If you just have the number, do: MonitorBottomPx(MonitorHandle(MonitorNum))
    Dim uMonInfo As MONITORINFO
    uMonInfo.cbSize = LenB(uMonInfo)
    If GetMonitorInfo(hMonitor, uMonInfo) = 0 Then Exit Function
    MonitorBottomPx = uMonInfo.rcMonitor.Bottom
End Function

