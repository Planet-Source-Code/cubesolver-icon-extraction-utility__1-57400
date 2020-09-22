VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3630
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame fraGeneral 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.Frame fraMRU 
         Caption         =   "MRU"
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         Begin VB.VScrollBar vsMRU 
            Height          =   235
            Left            =   2640
            Max             =   0
            Min             =   9
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   390
            Width           =   255
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear List"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            ToolTipText     =   "Remove the contents of the MRU list"
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkRecent 
            Caption         =   "Recently used file list:"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Mark to use the most recently used file list"
            Top             =   360
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.TextBox txtRecent 
            Height          =   285
            Left            =   2160
            MaxLength       =   3
            TabIndex        =   3
            Text            =   "4"
            ToolTipText     =   "Number of entries to keep in the history"
            Top             =   360
            Width           =   760
         End
      End
      Begin VB.Frame fraPlacement 
         Caption         =   "Form placement when started:"
         Height          =   1215
         Left            =   120
         TabIndex        =   6
         Top             =   1560
         Width           =   3135
         Begin VB.OptionButton optPlace 
            Caption         =   "Recall last position"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton optPlace 
            Caption         =   "Always center on screen"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / frmOptions                     \________________________/ v1.00 |
' |                                                                 |
' |       Description:  Simple form for setting application         |
' |                     options.                                    |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  November 19, 2004                           |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.00     CubeSolver  Nov 19, 2004  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Sub cmdCancel_Click()
  bClearMRU = False     ' Make sure the flag isn't set

  Unload Me
End Sub
Private Sub cmdClear_Click()
  If Msgbox(Me, "Are you sure you want to clean the MRU list?", vbQuestion + vbYesNo, "Confirm Clear") = vbYes Then
    bClearMRU = True    ' Turn the flag on
  End If
End Sub
Private Sub cmdOK_Click()
  ' Save all settings and exit the form
  If Val(txtRecent.Text) > 9 Then
    Msgbox Me, "The number must be between 1 and 9. Try again" & vbNewLine & _
           "by entering a number in this range.", vbExclamation, "Invalid Entry"
    txtRecent.SetFocus
    Exit Sub
  End If

  cIni.WriteEntry "MRU Size", txtRecent.Text
  cIni.WriteEntry "Recent List", CStr(chkRecent.Value)

  If optPlace(0).Value = True Then
    cIni.WriteEntry "Place Form", "0"
  Else
    cIni.WriteEntry "Place Form", "1"
  End If

  Unload Me
End Sub
Private Sub Form_Load()
  chkRecent.Value = Val(cIni.ReadEntry("Recent List", "1"))
  vsMRU.Value = Val(cIni.ReadEntry("MRU Size"))
  optPlace(cIni.ReadEntry("Place Form", "1")).Value = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Set frmOptions = Nothing
End Sub
Private Sub txtRecent_GotFocus()
  txtRecent.SelStart = 0
  txtRecent.SelLength = Len(txtRecent.Text)
End Sub
Private Sub txtRecent_KeyPress(KeyAscii As Integer)
  Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyBack    ', vbKeyReturn
      ' Valid keys
    Case vbKeyReturn
      cmdOK.SetFocus
      KeyAscii = 0
    Case Else
      KeyAscii = 0
  End Select
End Sub
Private Sub txtRecent_KeyUp(KeyCode As Integer, Shift As Integer)
  ' Allow the user to alter the value by using the arrow keys
  Dim lNum As Long

  lNum = vsMRU.Value

  Select Case KeyCode
    Case 38, 39   ' Up and right arrows
      If lNum < 9 Then
        lNum = lNum + 1
      End If
    Case 37, 40   ' Down and left arrows
      If lNum > 0 Then
        lNum = lNum - 1
      End If
  End Select
  KeyCode = 0
  vsMRU.Value = CStr(lNum)
End Sub
Private Sub vsMRU_Change()
  txtRecent.Text = CStr(vsMRU.Value)
  If Me.Visible Then
    chkRecent.Value = Abs(CInt(CBool(vsMRU.Value)))
  End If
End Sub
