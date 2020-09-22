VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / clsIni                         \________________________/ v1.00 |
' |                                                                 |
' |       Description:  Simple class for reading ini files          |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  October 14, 2003                            |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.00     CubeSolver  Oct 14, 2003  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private m_sIniFile As String, m_sIniArea As String

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Function ReadEntry(ByVal sKey As String, Optional sDefaultVal As String = vbNullString) As String
  Dim sBuffer As String
  Dim lRet As Long

  sBuffer = Space$(255)

  lRet = GetPrivateProfileString(m_sIniArea, sKey, vbNullString, sBuffer, 255, m_sIniFile)
  If lRet = 0 Then
    ReadEntry = sDefaultVal
  Else
    ReadEntry = Left$(sBuffer, lRet)
  End If
End Function
Public Sub WriteEntry(ByVal sKey As String, ByVal sValue As String)
  Call WritePrivateProfileString(m_sIniArea, sKey, sValue, m_sIniFile)
End Sub
Public Property Get IniArea() As String
  IniArea = m_sIniArea
End Property
Public Property Let IniArea(ByVal sArea As String)
  m_sIniArea = sArea
End Property
Public Property Get IniFile() As String
  IniFile = m_sIniFile
End Property
Public Property Let IniFile(ByVal sFile As String)
  m_sIniFile = sFile
End Property
