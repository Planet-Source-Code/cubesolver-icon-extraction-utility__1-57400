VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIcons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Extraction Utility"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7635
   Icon            =   "Icons.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7635
   Begin VB.PictureBox picPort 
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   7305
      TabIndex        =   2
      Top             =   120
      Width           =   7360
      Begin VB.PictureBox picContainer 
         BorderStyle     =   0  'None
         Height          =   4335
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   4335
         ScaleWidth      =   7365
         TabIndex        =   3
         Top             =   0
         Width           =   7360
         Begin VB.PictureBox picIcon 
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   -1120
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   4
            Top             =   120
            Width           =   480
         End
         Begin MSComDlg.CommonDialog cdlOpen 
            Left            =   6840
            Top             =   4200
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Shape shpHilite 
            BackColor       =   &H000000C0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000C0&
            Height          =   510
            Left            =   120
            Top             =   4560
            Visible         =   0   'False
            Width           =   510
         End
      End
   End
   Begin VB.VScrollBar vsScroll 
      Height          =   4335
      LargeChange     =   3
      Left            =   7370
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   5760
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4515
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7673
            MinWidth        =   7673
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2293
            MinWidth        =   2293
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1587
            MinWidth        =   1587
            Object.ToolTipText     =   "Time spent extracting"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5:22 AM"
         EndProperty
      EndProperty
   End
   Begin VB.Line linMenu 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   0
      X2              =   7680
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Line linMenu 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   7680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDeselect 
         Caption         =   "&Deselect Icon"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Selected Icon"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save &All Icons"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePrinterSetup 
         Caption         =   "Prin&ter Setup"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRUSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSize 
         Caption         =   "&Large Icons"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuViewSize 
         Caption         =   "&Small Icons"
         Enabled         =   0   'False
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpShow 
         Caption         =   "&Show Help Page"
      End
      Begin VB.Menu mnuHelpSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Icon Extraction Utility"
      End
   End
End
Attribute VB_Name = "frmIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / frmIcons                       \________________________/ v1.03 |
' |                                                                 |
' |       Description:  Extract icons from executables, dynamic     |
' |                     libraries and Active-X files.               |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  September 18, 2003                          |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.03     CubeSolver  Jan 26, 2005  Allow scrolling for large  |
'  |                                    number of icons.           |
'  | 1.02     CubeSolver  Nov 18, 2004  Use picturebox array,      |
'  |                                    loading and destroying     |
'  |                                    as needed.                 |
'  | 1.01     CubeSolver  Sep 02, 2004  Added MRU class.           |
'  | 1.00     CubeSolver  Sep 18, 2003  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private bUseMRU As Boolean
Private cMRU As clsMRU
Private lLargeIcons() As Long
Private lSmallIcons() As Long
Private lIconCount As Long
Private lSelected As Long
Private sFileName As String
Private sIniFile As String
Private sInitDir As String
Private sSaveDir As String

Private Const LARGE_ICON As Long = 32
Private Const SMALL_ICON As Long = 16
Private Const DI_NORMAL As Long = 3
Private Const OFFSET As Long = 120

Private Type RECT
  lLeft As Long
  lTop As Long
  lRight As Long
  lBottom As Long
End Type
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Function CheckPrinterCount() As Long
  CheckPrinterCount = Printers.Count
  mnuFilePrinterSetup.Enabled = CBool(CheckPrinterCount)
End Function
Private Sub DisplayIcons()
  Dim lBegin As Long, lEnd As Long
  Dim lLoop As Long

  lBegin = GetTickCount   ' Time how long this operation takes

  ' Clear our picturebox array
  If lIconCount > 0 Then
    Call PicArrayClear(lIconCount)
  End If

  ' Get the total number of icons in the file
  lIconCount = ExtractIconEx(sFileName, -1, 0, 0, 0)
  If lIconCount = 0 Then
    Msgbox Me, "There are no icons in the file you selected.", vbInformation, "Cannot Open"
    Exit Sub
  End If

  ' Let the user know we're working
  Screen.MousePointer = vbHourglass
  Call UpdateStatus(1, "Extracting...")
  Call UpdateStatus(2, vbNullString)
  Call LockWindowUpdate(picContainer.hWnd)

  ' Create pictureboxes to display all the icons
  Call PicArrayInitialize(lIconCount)

  ' Let the user see that we've cleaned up any old icons
  Call LockWindowUpdate(0)
  picContainer.Refresh
  Call LockWindowUpdate(picContainer.hWnd)

  ' Update GUI
  mnuFileClose.Enabled = True
  mnuViewSize(0).Enabled = True
  mnuViewSize(1).Enabled = True
  Call UpdateStatus(3, vbNullString)
  shpHilite.Visible = False

  ' So we don't think any icons are selected
  lSelected = -1

  ReDim lLargeIcons(lIconCount)
  ReDim lSmallIcons(lIconCount)
  imlIcons.ListImages.Clear

  For lLoop = 0 To lIconCount - 1
    Call GetIcon(lLoop)
  Next

  lEnd = GetTickCount

  ' Display information about the file
  Call UpdateStatus(1, sFileName)
  Call UpdateStatus(2, lIconCount & " icon" & IIf(lIconCount = 1, vbNullString, "s") & " found")
  Call UpdateStatus(3, lEnd - lBegin & " ms")

  Call LockWindowUpdate(0)
  Screen.MousePointer = vbDefault
End Sub
Private Sub GetIcon(ByVal lIndex As Long)
  On Error GoTo ErrorHandler

  ' Get the handle of the icon indicated by lIndex
  Call ExtractIconEx(sFileName, lIndex, lLargeIcons(lIndex), lSmallIcons(lIndex), 1)

  With picIcon(lIndex + 1)
    Set .Picture = LoadPicture(vbNullString)
'    .AutoRedraw = True
    If mnuViewSize(0).Checked = True Then
      Call DrawIconEx(.hDC, 0, 0, lLargeIcons(lIndex), LARGE_ICON, LARGE_ICON, 0, 0, DI_NORMAL)
    Else
      Call DrawIconEx(.hDC, 0, 0, lSmallIcons(lIndex), SMALL_ICON, SMALL_ICON, 0, 0, DI_NORMAL)
    End If
'    .Refresh
  End With

  ' Store in image list for saving process
  imlIcons.ListImages.Add lIndex + 1, , picIcon(lIndex + 1).Image

  Exit Sub

ErrorHandler:
  Select Case Err.Number
    Case Else
      Msgbox Me, "Error #" & Err.Number & ":  " & Err.Description, vbExclamation, Err.Source
  End Select
End Sub
Private Sub PicArrayInitialize(ByVal lPicCount As Long)
  ' Create the necessary number of pictureboxes and place them appropriately
  Dim lLeft As Long, lTop As Long
  Dim lCol As Long, lRow As Long
  Dim lLoop As Long
  Dim sTemp As String

  For lLoop = 1 To lPicCount
    lCol = (lLoop Mod 12) - 1
    If lCol = -1 Then
      lCol = 11
    End If
    sTemp = Format$(lLoop / 12, "0.000")
    If Right$(sTemp, 3) = "000" Then
      lRow = Val(sTemp) - 1
    Else
      lRow = Val(Left$(sTemp, InStr(sTemp, ".") - 1))
    End If
    lTop = (600 * lRow) + OFFSET
    lLeft = (600 * lCol) + OFFSET
    Load picIcon(lLoop)
    picIcon(lLoop).Move lLeft, lTop
    picIcon(lLoop).Visible = True
    picIcon(lLoop).AutoRedraw = True
  Next

  ' Now setup viewport
  picContainer.Top = 0
  picContainer.Height = (OFFSET * 2) + (600 * (lRow + 1))
  If lRow + 1 > 7 Then
    vsScroll.Max = (lRow + 1) - 7
  Else
    vsScroll.Max = 0
  End If
End Sub
Private Sub PicArrayClear(ByVal lPicCount As Long)
  ' Clean up our picturebox array
  Dim lLoop As Long

  If lPicCount > 0 Then
    For lLoop = 1 To lPicCount
      Set picIcon(lLoop) = Nothing
      Unload picIcon(lLoop)
    Next
  End If
End Sub
Private Function IsCursorOverPic(ByRef picX As VB.PictureBox) As Boolean
  ' Determine if the cursor is over a certain picturebox
  Dim tWinRect As RECT, tCursorPoint As POINTAPI

  Call GetCursorPos(tCursorPoint)
  Call GetWindowRect(picX.hWnd, tWinRect)
  If tCursorPoint.X > tWinRect.lLeft And tCursorPoint.X < tWinRect.lRight Then
    If tCursorPoint.Y > tWinRect.lTop And tCursorPoint.Y < tWinRect.lBottom Then
      ' Cursor is over the picturebox
      IsCursorOverPic = True
    End If
  End If
End Function
Private Sub UpdateStatus(ByVal iPanel As Integer, ByVal sMsg As String)
  ' Easier way to update the statusbar panels
  sbrInfo.Panels.Item(iPanel).Text = sMsg
End Sub
Private Sub Form_Load()
  Dim bPlaced As Boolean, bRecallPlacement As Boolean
  Dim lRecent As Long
  Dim sXLeft As String, sXTop As String

  Set cIni = New clsIni
  Set cMRU = New clsMRU

  ' Location of the ini file
  sIniFile = App.Path
  sIniFile = sIniFile & IIf(Right$(sIniFile, 1) = "\", vbNullString, "\")
  sIniFile = sIniFile & App.Title & ".ini"

  ' Set up the Ini class
  cIni.IniArea = "Settings"
  cIni.IniFile = sIniFile

  ' Prepare for MRU class
  lRecent = Val(cIni.ReadEntry("MRU Size"))
  If Len(Trim$(cIni.ReadEntry("Recent List"))) = 0 Then
    ' First time set up, set default size
    lRecent = 4
    cIni.WriteEntry "MRU Size", "4"
    cIni.WriteEntry "Recent List", "1"
    bUseMRU = True
  Else
    bUseMRU = CBool(cIni.ReadEntry("Recent List"))
  End If

  ' Set up the MRU class
  cMRU.IniArea = "Recent"
  cMRU.IniFile = sIniFile
  cMRU.Max = 9
  cMRU.MRUSize = lRecent
  cMRU.FormName = Me
  If bUseMRU Then
    cMRU.GetMRU
  End If

  ' Recall form location
  sXLeft = Trim$(cIni.ReadEntry("Left"))
  sXTop = Trim$(cIni.ReadEntry("Top"))
  bRecallPlacement = CBool(cIni.ReadEntry("Place Form", "1"))
  If Not Len(sXLeft) = 0 And Not Len(sXTop) = 0 Then
    bPlaced = True
  End If

  Call CheckPrinterCount

  ' Restore form location
  If bPlaced Then
    If bRecallPlacement Then
'      Me.Move Val(sXLeft), Val(sXTop), 7455, 5430
      Me.Move Val(sXLeft), Val(sXTop), 7725, 5430
    Else
      Call CenterForm(Me)
    End If
  Else
    Call CenterForm(Me)
  End If
  Call CheckPlacement(Me)
End Sub
Private Sub Form_Resize()
  On Error Resume Next

  ' For some reason, Me.Move will occasionally resize the form
  If Me.Height < 5430 Then
    Me.Height = 5430
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  ' Clean up before exiting
  Erase lLargeIcons, lSmallIcons

  cIni.WriteEntry "Left", CStr(Me.Left)
  cIni.WriteEntry "Top", CStr(Me.Top)

  Set cIni = Nothing
  Set cMRU = Nothing
  Set frmIcons = Nothing
End Sub
Private Sub mnuFileClose_Click()
  mnuFileClose.Enabled = False
  mnuFileDeselect.Enabled = False
  mnuFilePrint.Enabled = False
  mnuFileSave.Enabled = False
  mnuFileSaveAll.Enabled = False
  mnuViewSize(0).Enabled = False
  mnuViewSize(1).Enabled = False
  Call UpdateStatus(1, vbNullString)
  Call UpdateStatus(2, vbNullString)
  Call UpdateStatus(3, vbNullString)
  shpHilite.Visible = False
  lSelected = -1

  Call PicArrayClear(lIconCount)
  lIconCount = 0
End Sub
Private Sub mnuFileDeselect_Click()
  Call UpdateStatus(2, vbNullString)
  mnuFileDeselect.Enabled = False
  mnuFilePrint.Enabled = False
  mnuFileSave.Enabled = False
  shpHilite.Visible = False
  lSelected = -1
End Sub
Private Sub mnuFileExit_Click()
  Unload Me
End Sub
Private Sub mnuFileMRU_Click(Index As Integer)
  ' Retrieve the full file name from the ini file
  sFileName = cMRU.ReadFromIni("MRU " & Index)

  ' Set the new initial directory for the common dialog
  cMRU.InitDir = Left$(sFileName, Len(sFileName) - InStr(1, StrReverse(sFileName), "\"))

  ' We have a file, so let's add it to the top of the MRU list
  cMRU.AddToMRUList sFileName

  mnuFileSaveAll.Enabled = True
  Call DisplayIcons
End Sub
Private Sub mnuFileOpen_Click()
  On Error GoTo ErrorHandler

  ' Display the File Open dialog
  ' Filter out all files except exes and dlls
  cdlOpen.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly
  cdlOpen.FileName = vbNullString
  cdlOpen.Filter = "Exe, Dll and Ocx Files (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|Executable Files (*.exe)|*.exe|Dll Files (*.dll)|*.dll|OCX Files (*.ocx)|*.ocx"
  If Len(sFileName) > 0 Then
    cdlOpen.InitDir = Left$(sFileName, Len(sFileName) - InStr(1, StrReverse(sFileName), "\"))
  End If
  cdlOpen.Action = 1
  sFileName = cdlOpen.FileName
  If Len(sFileName) = 0 Then
    Exit Sub
  End If
  sInitDir = Left$(sFileName, Len(sFileName) - InStr(1, StrReverse(sFileName), "\"))
  mnuFileSaveAll.Enabled = True

  If bUseMRU Then
    ' Remember the current folder
    cMRU.InitDir = sInitDir

    ' We have a file, so let's add it to the MRU list
    cMRU.AddToMRUList sFileName
  End If

  Call DisplayIcons

ErrorHandler:
End Sub
Private Sub mnuFilePrint_Click()
  ' Print the selected icon
  Printer.PaintPicture picIcon(lSelected).Picture, 1150, 950
  Printer.EndDoc
End Sub
Private Sub mnuFilePrinterSetup_Click()
  ' Show the printer setup dialog
  On Error GoTo ErrorHandler
  cdlOpen.CancelError = True

  cdlOpen.Flags = cdlPDNoSelection Or cdlPDNoPageNums Or _
  cdlPDHidePrintToFile Or cdlPDUseDevModeCopies
  cdlOpen.ShowPrinter

  Exit Sub

ErrorHandler:
'  Call Shell("rundll32.exe shell32.dll,SHHelpShortcuts_RunDLL PrintersFolder")
End Sub
Private Sub mnuFileSave_Click()
  ' Save the selected icon
  Dim sSaveAs As String, sTemp As String
  Dim picX As Picture

  On Error GoTo ErrorHandler

  sTemp = sFileName
  sTemp = Mid$(sTemp, Len(sTemp) - InStr(1, StrReverse(sTemp), "\") + 2)
  sTemp = Mid$(sTemp, 1, InStr(1, sTemp, ".") - 1) & "_" & lSelected  ' & ".ico"

  ' Get name to save file as
  cdlOpen.Filter = "Icon File (*.ico)|*.ico*"
  cdlOpen.FileName = sTemp
  If Len(sSaveDir) > 0 Then
    cdlOpen.InitDir = Left$(sSaveDir, Len(sSaveDir) - InStr(1, StrReverse(sSaveDir), "\"))
  End If
  cdlOpen.CancelError = True
  cdlOpen.ShowSave
  If Len(cdlOpen.FileName) = 0 Then
    Exit Sub
  Else
    sSaveAs = cdlOpen.FileName
    sSaveDir = Left$(sSaveAs, Len(sSaveAs) - InStr(1, StrReverse(sSaveAs), "\"))
  End If

  ' Make sure the file name has a .ico extension
  If LCase$(Right$(sSaveAs, 4)) <> ".ico" Then
    sSaveAs = sSaveAs & ".ico"
  End If

'  Call SavePicture(picIcon(lSelected).Image, sSaveAs)

  ' This method saves better than the above method (Note MaskColor of ImageList)
  Set picX = imlIcons.ListImages.Item(lSelected).ExtractIcon
  Call SavePicture(picX, sSaveAs)
  Set picX = Nothing

  Call UpdateStatus(3, "Saved")
  Exit Sub

ErrorHandler:
  Select Case Err.Number
    Case 32755
      ' Do nothing, Cancel was selected
    Case Else
      Msgbox Me, "Error #" & Err.Number & ":  " & Err.Description, vbExclamation, Err.Source
  End Select
End Sub
Private Sub mnuFileSaveAll_Click()
  ' Save all the icons in the current file
  Dim lLoop As Long
  Dim picX As Picture
  Dim sFolder As String, sIcon As String
  Dim sOldStatus As String, sTemp As String

  On Error GoTo ErrorHandler

  Screen.MousePointer = vbHourglass
  sOldStatus = sbrInfo.Panels.Item(1).Text
  Call UpdateStatus(1, "Saving all icons...")

  sFolder = BrowseForFolder(Me)
  If Len(Trim$(sFolder)) = 0 Then
    Exit Sub
  End If

  sFolder = sFolder & IIf(Right$(sFolder, 1) = "\", vbNullString, "\")

  sIcon = sFileName
  sIcon = Mid$(sIcon, Len(sIcon) - InStr(1, StrReverse(sIcon), "\") + 2)
  sIcon = Mid$(sIcon, 1, InStr(1, sIcon, ".") - 1) & "_"

  For lLoop = 1 To lIconCount
    sTemp = sFolder & sIcon & lLoop & ".ico"
    If CBool(PathFileExists(sTemp)) Then
      Kill sTemp
    End If
    Set picX = imlIcons.ListImages.Item(lLoop).ExtractIcon
    Call SavePicture(picX, sTemp)
    Set picX = Nothing
  Next

ErrorHandler:
  Call UpdateStatus(1, sOldStatus)
  Screen.MousePointer = vbDefault
End Sub
Private Sub mnuHelpAbout_Click()
  Msgbox Me, "Icon Extraction Utility v" & App.Major & "." & Format$(App.Minor, "00") & vbNewLine & vbNewLine & "     by CubeSolver", vbInformation + vbOKOnly, "About..."
End Sub
Private Sub mnuHelpOptions_Click()
  ' Show the options form
  frmOptions.Show vbModal, Me

  ' Clean the list if the user selected that in Options
  If bClearMRU Then
    bClearMRU = False   ' Turn the flag off
    cMRU.WriteToIni vbNullString, vbNullString
    cMRU.HideMRU
  End If

  ' Reset the size of the list
  cMRU.MRUSize = cIni.ReadEntry("MRU Size")
  bUseMRU = CBool(cIni.ReadEntry("Recent List"))
  If bUseMRU Then
    cMRU.GetMRU     ' Repopulate the MRU list
  Else
    cMRU.HideMRU
  End If
End Sub
Private Sub mnuHelpShow_Click()
  ' Show the help form
  Call CreateHelpFile("Icon Extraction Help")
  frmHelp.Show vbModeless, Me
End Sub
Private Sub mnuViewSize_Click(Index As Integer)
  ' Only change when a non-checked menu item is clicked
  If mnuViewSize(Index).Checked = False Then
    mnuViewSize(0).Checked = Not (mnuViewSize(0).Checked)
    mnuViewSize(1).Checked = Not (mnuViewSize(1).Checked)
  End If

  mnuFilePrint.Enabled = False
  mnuFileSave.Enabled = False
  shpHilite.Visible = False
  lSelected = -1

  If lIconCount > 0 Then
    Call DisplayIcons
  End If
End Sub
Private Sub picContainer_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Allow executables and dlls to be dropped directly onto the icon display area
  Dim sDroppedFile As String
  Dim lLoop As Long

  If Not Data.GetFormat(vbCFFiles) Then Exit Sub
  If Data.Files.Count > 1 Then
    For lLoop = 1 To Data.Files.Count
      sDroppedFile = sDroppedFile & Data.Files.Item(lLoop) & vbCrLf
    Next
    DoEvents
    Msgbox Me, "You have attempted to open more than one file." & vbCrLf & vbCrLf & sDroppedFile, vbExclamation + vbOKOnly, "Cannot Open"
    Exit Sub
  End If
  sDroppedFile = Data.Files.Item(1)   ' Only grab the first item

  Select Case LCase$(Right$(sDroppedFile, 4))
    Case ".exe", ".dll", ".ocx"
      If bUseMRU Then
        ' We have a file, so let's add it to the MRU list
        cMRU.AddToMRUList sDroppedFile
      End If
      sFileName = sDroppedFile
      Call DisplayIcons
    Case Else
      Msgbox Me, Data.Files.Item(1) & vbNewLine & "is an invalid file type.", vbExclamation + vbOKOnly, "Cannot Open"
  End Select

' ### Remove commented code below
'  If LCase$(Right$(sDroppedFile, 4)) = ".exe" Or LCase$(Right$(sDroppedFile, 4)) = ".dll" Then
'    If bUseMRU Then
'      ' We have a file, so let's add it to the MRU list
'      cMRU.AddToMRUList sDroppedFile
'    End If
'    sFileName = sDroppedFile
'    Call DisplayIcons
'  Else
'    Msgbox Me, Data.Files.Item(1) & vbCrLf & "is an invalid file type.", vbExclamation + vbOKOnly, "Cannot Open"
'  End If
End Sub
Private Sub picIcon_DblClick(Index As Integer)
  If lSelected > -1 And lSelected = Index Then
    Call mnuFileSave_Click
  End If
End Sub
Private Sub picIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  ' Only process this sub if the cursor is still positioned over the picturebox
  If Not IsCursorOverPic(picIcon(Index)) Then
    Exit Sub
  End If

  If Shift = 1 Or Shift = 2 Then      ' Shift or Ctrl keys
    Msgbox Me, "Multiple icons cannot be selected.", vbInformation, "No Multiple Selection"
  ElseIf Button = vbRightButton Then
    If lSelected = Index Then
      Call UpdateStatus(2, vbNullString)
      mnuFileDeselect.Enabled = False
      mnuFilePrint.Enabled = False
      mnuFileSave.Enabled = False
      shpHilite.Visible = False
      lSelected = -1
    End If
  Else
    If shpHilite.Visible = False Then
      shpHilite.Visible = True
    End If

    Call UpdateStatus(2, "Icon " & Index)
    Call UpdateStatus(3, vbNullString)
    shpHilite.Move picIcon(Index).Left - 20, picIcon(Index).Top - 20
    shpHilite.Visible = True
    mnuFileDeselect.Enabled = True
    mnuFileSave.Enabled = True
    mnuFilePrint.Enabled = CBool(CheckPrinterCount)
    lSelected = Index
  End If
End Sub
Private Sub vsScroll_Change()
  picContainer.Top = 0 - (vsScroll.Value * 600)
End Sub
