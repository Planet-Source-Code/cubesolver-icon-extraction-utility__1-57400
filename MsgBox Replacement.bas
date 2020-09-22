Attribute VB_Name = "modMsgBoxReplacement"
Option Explicit

' Code taken from:
' http://vbnet.mvps.org/index.html?code/hooks/messageboxhookcentre.htm
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Modified by CubeSolver Nov 23, 2004
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' In order to pass the form's coordinates to MsgBoxHookProc
Private frmLeft As Long
Private frmTop As Long
Private frmWidth As Long
Private frmHeight As Long

' Misc API constants
Private Const WH_CBT As Long = 5
Private Const GWL_HINSTANCE As Long = (-6)
Private Const HCBT_ACTIVATE As Long = 5

' UDT for passing data through the hook
Private Type MSGBOX_HOOK_PARAMS
  hWndOwner   As Long
  hHook       As Long
End Type

Private Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type

' Need this declared at module level as
' it is used in the call and the hook proc
Private mhp As MSGBOX_HOOK_PARAMS

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Function Msgbox(ByRef frmX As VB.Form, ByVal sPrompt As String, Optional ByVal dwStyle As Long, Optional ByVal sTitle As String) As Long
  ' Replaces VB's built in MsgBox function in VB5/6
  Dim hInstance As Long
  Dim hThreadId As Long

  ' Set up form coordinates
  frmLeft = frmX.Left \ Screen.TwipsPerPixelX
  frmTop = frmX.Top \ Screen.TwipsPerPixelY
  frmWidth = frmX.Width \ Screen.TwipsPerPixelX
  frmHeight = frmX.Height \ Screen.TwipsPerPixelX

  If dwStyle = 0 Then dwStyle = vbOKOnly
  If Len(sTitle) = 0 Then sTitle = App.Title

  'Set up the hook
   hInstance = GetWindowLong(frmX.hWnd, GWL_HINSTANCE)
   hThreadId = GetCurrentThreadId()

  ' Set up the MSGBOX_HOOK_PARAMS values
  ' By specifying a Windows hook as one of the
  ' params, we can intercept messages sent by
  ' Windows and thereby manipulate the dialog
  With mhp
    .hWndOwner = frmX.hWnd
    .hHook = SetWindowsHookEx(WH_CBT, AddressOf MsgBoxHookProc, hInstance, hThreadId)
  End With

  ' Call the MessageBox API and return the
  ' value as the result of this function
  Msgbox = MessageBox(frmX.hWnd, sPrompt, sTitle, dwStyle)
End Function
Private Function MsgBoxHookProc(ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rc As RECT

  ' Temporary vars for demo
  Dim lNewLeft As Long
  Dim lNewTop As Long
  Dim lDlgWidth As Long
  Dim lDlgHeight As Long
  Dim hWndMsgBox As Long

  ' When the message box is about to be shown,
  ' center the dialog
  If uMsg = HCBT_ACTIVATE Then
    'in a HCBT_ACTIVATE message, wParam holds
    'the handle to the messagebox
    hWndMsgBox = wParam

    ' Just as was done in other API hook demos, position
    ' the dialog centered in the calling parent form
    Call GetWindowRect(hWndMsgBox, rc)

    lDlgWidth = rc.lRight - rc.lLeft
    lDlgHeight = rc.lBottom - rc.lTop

    lNewLeft = frmLeft + ((frmWidth - lDlgWidth) \ 2)
    lNewTop = frmTop + ((frmHeight - lDlgHeight) \ 2)

    Call MoveWindow(hWndMsgBox, lNewLeft, lNewTop, lDlgWidth, lDlgHeight, True)

    ' Done with the dialog so release the hook
    Call UnhookWindowsHookEx(mhp.hHook)
  End If

  ' Return False to let normal processing continue
  MsgBoxHookProc = False
End Function
