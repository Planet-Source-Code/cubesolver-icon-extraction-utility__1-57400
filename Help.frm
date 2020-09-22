VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmHelp 
   Caption         =   "Help"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   6960
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser brwHelp 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      ExtentX         =   12303
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / frmHelp                        \________________________/ v1.00 |
' |                                                                 |
' |       Description:  Simple form for displaying HTML help file.  |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  September 24, 2003                          |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.00     CubeSolver  Sep 24, 2003  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Sub Form_Load()
  If CBool(PathFileExists(sHelpFile)) Then
    brwHelp.Silent = True
    brwHelp.Navigate sHelpFile
  Else
    Msgbox Me, "Missing help file!", vbExclamation, "Error"
  End If
  Me.Caption = sHelpTitle
End Sub
Private Sub Form_Resize()
  On Error Resume Next

  brwHelp.Move 0, 0, Me.Width - 100, Me.Height - 400
End Sub
Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next

  Call Kill(sHelpFile)
  sHelpFile = vbNullString

  Set frmHelp = Nothing
End Sub
