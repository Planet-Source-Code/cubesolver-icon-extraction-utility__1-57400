Attribute VB_Name = "modCheckPlacement"
Option Explicit

'  ________________________________                          _______
' / modCheckPlacement              \________________________/ v1.04 |
' |                                                                 |
' |       Description:  If your app saves the placement of forms on |
' |                     the screen, you can use this module to      |
' |                     ensure that, when recalled, those settings  |
' |                     aren't placing the form in a strange        |
' |                     location.  Make sure it stays visible to    |
' |                     the user whether it was moved off the       |
' |                     screen or under the taskbar and then saved  |
' |                     with those coordinates or if the screen     |
' |                     resolution has changed since last use.      |
' |                     Also includes center a form over the form   |
' |                     from which it was called.  This is in case  |
' |                     you prefer to place your forms manually     |
' |                     instead of relying on the StartUpPosition   |
' |                     property.  Includes procedures for          |
' |                     centering VB forms.                         |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  November 20, 2003                           |
' |      OS Tested On:  Windows NT 4 SP 6a                          |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.04     CubeSolver  Oct 15, 2004  Added ResizeFormToOjbect   |
'  |                                    sub.                       |
'  | 1.03     CubeSolver  Oct 01, 2004  Added CenterForm and       |
'  |                                    CenterOverWindow subs.     |
'  | 1.02     CubeSolver  Apr 16, 2004  Added CenterOverOwner sub. |
'  | 1.01     CubeSolver  Nov 21, 2003  Eliminate the use of the   |
'  |                                    screen object to allow for |
'  |                                    virtual screen sizes used  |
'  |                                    by multiple monitors.      |
'  |                                    Thanks to Chloe for        |
'  |                                    bringing this issue to     |
'  |                                    light.                     |
'  | 1.00     CubeSolver  Nov 20, 2003  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type

Private Const HEIGHT_MIN As Long = 3000         ' Minimum height of form - used in ResizeFormToObject sub
Private Const SPI_GETWORKAREA As Long = 48
Private Const WIDTH_MIN As Long = 3700          ' Minimum width of form - used in ResizeFormToObject sub

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Sub CenterForm(ByRef frmX As VB.Form)
  ' Center a VB form on the work area
  Dim lLeft As Long, lTop As Long
  Dim tWA As RECT

  ' Get the work area dimensions
  Call SystemParametersInfo(SPI_GETWORKAREA, 0, tWA, 0)

  ' Convert the virtual coordinates (pixels) to scale coordinates (twips)
  tWA.lLeft = tWA.lLeft * Screen.TwipsPerPixelX
  tWA.lRight = tWA.lRight * Screen.TwipsPerPixelX
  tWA.lTop = tWA.lTop * Screen.TwipsPerPixelY
  tWA.lBottom = tWA.lBottom * Screen.TwipsPerPixelY

  ' Calculate the new coordinates
  lTop = (tWA.lBottom \ 2) - (frmX.Height \ 2)
  lLeft = (tWA.lRight \ 2) - (frmX.Width \ 2)

  ' Place the form using the new coordinates
  frmX.Move lLeft, lTop
End Sub
Public Sub CenterOverOwner(ByRef frmToCenter As VB.Form, ByRef frmOwner As VB.Form)
  ' Center a VB form over another VB form
  Dim lOwnerLeft As Long, lOwnerTop As Long
  Dim lOwnerWidth As Long, lOwnerHeight As Long
  Dim lMoveLeft As Long, lMoveTop As Long

  ' Get current placement and size of owner form
  lOwnerLeft = frmOwner.Left
  lOwnerTop = frmOwner.Top
  lOwnerWidth = frmOwner.Width
  lOwnerHeight = frmOwner.Height

  ' Modify form's placement on the x axis
  If lOwnerWidth < frmToCenter.Width Then
    lMoveLeft = (frmToCenter.Width - lOwnerWidth) \ 2
    frmToCenter.Left = lOwnerLeft - lMoveLeft
  ElseIf lOwnerWidth > frmToCenter.Width Then
    lMoveLeft = (lOwnerWidth - frmToCenter.Width) \ 2
    frmToCenter.Left = lOwnerLeft + lMoveLeft
  Else
    frmToCenter.Left = lOwnerLeft
  End If

  ' Modify form's placement on the y axis
  If lOwnerHeight < frmToCenter.Height Then
    lMoveTop = (frmToCenter.Height - lOwnerHeight) \ 2
    frmToCenter.Top = lOwnerTop - lMoveTop
  ElseIf lOwnerHeight > frmToCenter.Height Then
    lMoveTop = (lOwnerHeight - frmToCenter.Height) \ 2
    frmToCenter.Top = lOwnerTop + lMoveTop
  Else
    frmToCenter.Top = lOwnerTop
  End If
  Call CheckPlacement(frmToCenter)
End Sub
Public Sub CenterOverWindow(ByRef frmToCenter As VB.Form, ByVal lWindowhWnd As Long)
  ' Center a VB form over any other form whose hWnd has been passed
  Dim lOwnerWidth As Long, lOwnerHeight As Long
  Dim lMoveLeft As Long, lMoveTop As Long
  Dim tWindow As RECT

  ' Get current placement and size of owner form
  Call GetWindowRect(lWindowhWnd, tWindow)
  lOwnerWidth = (tWindow.lRight - tWindow.lLeft) * Screen.TwipsPerPixelX
  lOwnerHeight = (tWindow.lBottom - tWindow.lTop) * Screen.TwipsPerPixelY
  tWindow.lLeft = tWindow.lLeft * Screen.TwipsPerPixelX
  tWindow.lTop = tWindow.lTop * Screen.TwipsPerPixelY

  ' Modify form's placement on the x axis
  If lOwnerWidth < frmToCenter.Width Then
    lMoveLeft = (frmToCenter.Width - lOwnerWidth) \ 2
    frmToCenter.Left = tWindow.lLeft - lMoveLeft
  ElseIf lOwnerWidth > frmToCenter.Width Then
    lMoveLeft = (lOwnerWidth - frmToCenter.Width) \ 2
    frmToCenter.Left = tWindow.lLeft + lMoveLeft
  Else
    frmToCenter.Left = tWindow.lLeft
  End If

  ' Modify form's placement on the y axis
  If lOwnerHeight < frmToCenter.Height Then
    lMoveTop = (frmToCenter.Height - lOwnerHeight) \ 2
    frmToCenter.Top = tWindow.lTop - lMoveTop
  ElseIf lOwnerHeight > frmToCenter.Height Then
    lMoveTop = (lOwnerHeight - frmToCenter.Height) \ 2
    frmToCenter.Top = tWindow.lTop + lMoveTop
  Else
    frmToCenter.Top = tWindow.lTop
  End If
  Call CheckPlacement(frmToCenter)
End Sub
Public Sub CheckPlacement(ByRef frmX As VB.Form)
  ' Make sure the VB form is on the screen and not under the taskbar
  Dim tWA As RECT             ' Hold the virtual screen dimensions

  Const SPACER As Long = 15   ' Default distance away from edge

  ' Get the work area dimensions
  Call SystemParametersInfo(SPI_GETWORKAREA, 0, tWA, 0)

  ' Convert the virtual coordinates (pixels) to scale coordinates (twips)
  tWA.lLeft = tWA.lLeft * Screen.TwipsPerPixelX
  tWA.lRight = tWA.lRight * Screen.TwipsPerPixelX
  tWA.lTop = tWA.lTop * Screen.TwipsPerPixelY
  tWA.lBottom = tWA.lBottom * Screen.TwipsPerPixelY

  ' Check if x axis modification is necessary
  If frmX.Width + frmX.Left > tWA.lRight Then
    frmX.Left = tWA.lRight - frmX.Width - SPACER
  ElseIf frmX.Left < tWA.lLeft Then
    frmX.Left = tWA.lLeft + SPACER
  End If

  ' Check if y axis modification is necessary
  If frmX.Height + frmX.Top > tWA.lBottom Then
    frmX.Top = tWA.lBottom - frmX.Height - SPACER
  ElseIf frmX.Top < tWA.lTop Then
    frmX.Top = tWA.lTop + SPACER
  End If
End Sub
Public Sub ResizeFormToObject(ByRef frmX As VB.Form, ByRef objX As Object, Optional ByVal lWOffset As Long = 0, Optional ByVal lHOffset As Long = 0, Optional ByVal lWMin As Long = WIDTH_MIN, Optional ByVal lHMin As Long = HEIGHT_MIN)
  ' When you change the size of an object, such as allowing AutoSize on a PictureBox, you can
  ' use this sub to resize the form accordingly
  Dim lNewHeight As Long, lNewWidth As Long
  Dim tWA As RECT

  ' Get the work area dimensions
  Call SystemParametersInfo(SPI_GETWORKAREA, 0, tWA, 0)

  ' Convert the virtual coordinates (pixels) to scale coordinates (twips)
  tWA.lLeft = tWA.lLeft * Screen.TwipsPerPixelX
  tWA.lRight = tWA.lRight * Screen.TwipsPerPixelX
  tWA.lTop = tWA.lTop * Screen.TwipsPerPixelY
  tWA.lBottom = tWA.lBottom * Screen.TwipsPerPixelY

  lNewHeight = objX.Height + lHOffset
  lNewWidth = objX.Width + lWOffset

  If lNewHeight > tWA.lBottom - tWA.lTop Then
    lNewHeight = tWA.lBottom - tWA.lTop
  End If
  If lNewWidth > tWA.lRight - tWA.lLeft Then
    lNewWidth = tWA.lRight - tWA.lLeft
  End If
  If lNewHeight < lHMin Then
    lNewHeight = lHMin
  End If
  If lNewWidth < lWMin Then
    lNewWidth = lWMin
  End If

  frmX.Height = lNewHeight
  frmX.Width = lNewWidth
End Sub
