Attribute VB_Name = "modAdvMapEditor"
Option Explicit

Private Const GWL_WNDPROC       As Long = (-4)
Private Const GWL_USERDATA      As Long = (-21)
Private Const GWL_STYLE         As Long = (-16)

Private Const WM_DESTROY        As Long = &H2
Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_CAPTURECHANGED As Long = &H215
Private Const WM_GETMINMAXINFO  As Long = &H24
Private Const WM_NCDESTROY      As Long = &H82

Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOMOVE        As Long = &H2

Private Const WS_CAPTION        As Long = &HC00000

'API Declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                        (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32.dll" ( _
     ByVal hwnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, _
ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName _
         As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)
'Types
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
'Mod Globals
Private g_MovingMainWnd As Boolean
Private g_OrigCursorPos As POINTAPI
Private g_OrigWndPos As POINTAPI

Global gHW As Long

Public Sub MapEditorMode(switch As Boolean)
    If switch Then
        frmMain.Width = frmMain.Width - 30
        frmMain.Height = frmMain.Height + 600
        frmMain.picForm.Top = frmMain.picForm.Top + 24 + 40
        
        If frmMain.mapPreviewSwitch.Value Then
            frmMain.mapPreviewSwitch.Picture = LoadResPicture("MAP_DOWN", vbResBitmap)
            frmMapPreview.Show
            frmMapPreview.RecalcuateDimensions
        Else
            frmMain.mapPreviewSwitch.Picture = LoadResPicture("MAP_UP", vbResBitmap)
        End If
    Else
        frmMain.Width = frmMain.Width + 30
        frmMain.Height = frmMain.Height - 600
        frmMain.picForm.Top = frmMain.picForm.Top - 24 - 40
        frmMapPreview.Hide
    End If
    Call FlipBit(WS_CAPTION, Not switch)
End Sub
Private Function FlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
   Dim nStyle As Long
   
   nStyle = GetWindowLong(frmMain.hwnd, GWL_STYLE)
   
   If Value Then
      nStyle = nStyle Or Bit
   Else
      nStyle = nStyle And Not Bit
   End If
   Call SetWindowLong(frmMain.hwnd, GWL_STYLE, nStyle)
   Call Redraw
   
   FlipBit = (nStyle = GetWindowLong(frmMain.hwnd, GWL_STYLE))
End Function
Public Sub Redraw()
   ' Redraw window with new style.
   Const swpFlags As Long = _
      SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
      'SWP_NOZORDER
    SetWindowPos frmMain.hwnd, 0, 0, 0, 0, 0, swpFlags
End Sub

Public Sub MainMouseMove(hwnd As Long)

    If (g_MovingMainWnd) Then

        Dim pt As POINTAPI

        If (GetCursorPos(pt)) Then

            Dim wnd_x As Long, wnd_y As Long

            wnd_x = g_OrigWndPos.x + (pt.x - g_OrigCursorPos.x)
            wnd_y = g_OrigWndPos.y + (pt.y - g_OrigCursorPos.y)
            SetWindowPos frmMain.hwnd, 0, wnd_x, wnd_y, 0, 0, (SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE)
            frmMapPreview.Move frmMain.Left - frmMapPreview.Width - 80, frmMain.Top + 75
        End If
    End If

End Sub
Public Sub MainCaptureChanged(hwnd As Long, lParam As Long)
    g_MovingMainWnd = IIf(lParam = hwnd, True, False)
End Sub
Public Sub MainLButtonUp(hwnd As Long)
    ReleaseCapture
End Sub

Public Sub MainLButtonDown(hwnd As Long)

    If (GetCursorPos(g_OrigCursorPos)) Then

        Dim rt As RECT

        GetWindowRect frmMain.hwnd, rt
        g_OrigWndPos.x = rt.Left
        g_OrigWndPos.y = rt.Top
        g_MovingMainWnd = True
        SetCapture hwnd
    End If

End Sub
Public Sub MainPreventResizing(hwnd As Long, constWidth As Long, constHeight As Long, ByRef lParam As Long)
                 Dim MMI As MINMAXINFO
                  
                  CopyMemory MMI, ByVal lParam, LenB(MMI)
                   With MMI
                      .ptMinTrackSize.x = constWidth
                      .ptMinTrackSize.y = constHeight
                      .ptMaxTrackSize.x = constWidth
                      .ptMaxTrackSize.y = constHeight
                  End With
                  CopyMemory ByVal lParam, MMI, LenB(MMI)
End Sub
