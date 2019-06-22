Attribute VB_Name = "scroll"
Option Explicit

Type POINTAPI
    X As Long
    Y As Long
End Type

Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

#If VBA7 Then
    #If Win64 Then
        Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal Point As LongLong) As LongPtr
    #Else
        Declare PtrSafe Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
    #End If
    Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As LongPtr
    Declare PtrSafe Function SetFocus Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As LongPtr
    Dim hwnd As LongPtr, lMouseHook As LongPtr
#Else
    Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
    Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function GetActiveWindow Lib "user32" () As Long
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
    Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Dim hwnd As Long, lMouseHook As Long
#End If

Const WH_MOUSE_LL = 14
Const WM_MOUSEWHEEL = &H20A
Const HC_ACTION = 0

Dim oComboBox As Object

Sub SetComboBoxHook(ByVal Control As Object)
    Dim tPt As POINTAPI
    Dim sBuffer As String
    Dim lRet As Long
    
    Set oComboBox = Control
    RemoveComboBoxHook
    GetCursorPos tPt
    #If Win64 Then
        Dim lPt As LongPtr
        CopyMemory lPt, tPt, LenB(tPt)
        hwnd = WindowFromPoint(lPt)
    #Else
        hwnd = WindowFromPoint(tPt.X, tPt.Y)
    #End If
    sBuffer = Space(256)
    lRet = GetClassName(GetParent(hwnd), sBuffer, 256)
    If InStr(Left(sBuffer, lRet), "MdcPopup") Then
        SetFocus hwnd
        #If Win64 Then
            lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.HinstancePtr, 0)
        #Else
            lMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf MouseProc, Application.Hinstance, 0)
        #End If
    End If
End Sub

Sub RemoveComboBoxHook()
    UnhookWindowsHookEx lMouseHook
End Sub

#If VBA7 Then
    Function MouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, lParam As MSLLHOOKSTRUCT) As LongPtr
#Else
    Function MouseProc(ByVal nCode As Long, ByVal wParam As Long, lParam As MSLLHOOKSTRUCT) As Long
#End If

    Dim sBuffer As String
    Dim lRet As Long
        
    sBuffer = Space(256)
    lRet = GetClassName(GetActiveWindow, sBuffer, 256)
    If Left(sBuffer, lRet) = "wndclass_desked_gsk" Then Call RemoveComboBoxHook
    If IsWindow(hwnd) = 0 Then Call RemoveComboBoxHook
    
    If (nCode = HC_ACTION) Then
        If wParam = WM_MOUSEWHEEL Then
        #If Win64 Then
            Dim lPt As LongPtr
            CopyMemory lPt, lParam.pt, LenB(lParam.pt)
            If WindowFromPoint(lPt) = hwnd Then
        #Else
            If WindowFromPoint(lParam.pt.X, lParam.pt.Y) = hwnd Then
        #End If
                On Error Resume Next
                    If lParam.mouseData > 0 Then
                        oComboBox.ListIndex = oComboBox.ListIndex - 1
                    Else
                        oComboBox.ListIndex = oComboBox.ListIndex + 1
                    End If
                On Error GoTo 0
            End If
        End If
    End If
    
    MouseProc = CallNextHookEx(lMouseHook, nCode, wParam, ByVal lParam)
End Function
