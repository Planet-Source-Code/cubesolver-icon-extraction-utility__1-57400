VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'  ________________________________                          _______
' / clsMRU                         \________________________/ v1.03 |
' |                                                                 |
' | Class Description:  MRU - Most Recently Used (A persistent list |
' |                     of the most recently used files.)           |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  October 08, 2003                            |
' |      OS Tested On:  Windows NT 4 SP 6a                          |
' |     Notes for Use:  - Must add as many blank menu items to the  |
' |                       form as you need called mnuFileMRU with   |
' |                       index numbers starting at zero. Visible   |
' |                       property on all should be set to False.   |
' |                     - Add a menu separator at the bottom of the |
' |                       MRU menus list called mnuFileMRUSep,      |
' |                       also with Visible property set to False.  |
' |                     - Set the Max property to the max number of |
' |                       mnuFileMRU menus you created on the form. |
' |                     - Set the MRUSize property to the max       |
' |                       amount of MRU menus you wish to allow or  |
' |                       less.                                     |
' |                     - Point IniFile to the physical location on |
' |                       which to save/read the MRU list.          |
' |                     - Set IniArea to the heading in the ini     |
' |                       file that you'd prefer to use.            |
' |                     - Set the FormName property to the form on  |
' |                       which the MRU menus reside.               |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.03     CubeSolver  Nov 19, 2004  Added HideMRU sub.         |
'  | 1.02     CubeSolver  Jul 27, 2004  Fixed problem with MRU     |
'  |                                    dup entries.               |
'  | 1.01     CubeSolver  Oct 09, 2003  Use ini instead of         |
'  |                                    registry                   |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private m_frmX As VB.Form     ' Form that is using the MRU list
Private m_lMax As Long        ' Max size of MRU list
Private m_lMRUSize As Long    ' Current size of the MRU list
Private m_sIniArea As String  ' Section inside ini file
Private m_sIniFile As String  ' Location of ini file
Private m_sInitDir As String  ' Keep track of the MRU directory for common dialog

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Function AbbreviateString(ByVal sEntry As String) As String
  ' Very crude method of condensing a string to fit in a small area
  Dim lLoc As Long, lLoop As Long, lLen As Long
  Dim sFirst As String, sLast As String, sSegment As String
  Dim sTemp As String, sFill As String
  Dim bLoop As Long

  If Len(sEntry) < 48 Then
    AbbreviateString = sEntry
    Exit Function
  End If

  sFirst = Left$(sEntry, 3) & "...\"
  lLoc = Len(sEntry) - InStr(1, StrReverse(sEntry), "\")
  sLast = Mid$(sEntry, lLoc + 2)
  sTemp = Left$(sEntry, lLoc)

  bLoop = True

  ' Work backwards through the string
  Do While bLoop
    lLoc = Len(sTemp) - InStr(1, StrReverse(sTemp), "\")
    sSegment = Mid$(sTemp, lLoc + 2)
    sTemp = Left$(sTemp, lLoc)
    If lLoc = 0 Then
      bLoop = False
    ElseIf Len(sSegment & "\" & sLast) < 36 Then
      sLast = sSegment & "\" & sLast
    ElseIf Len(sSegment & "\" & sLast) = 36 Then
      sLast = sSegment & "\" & sLast
      bLoop = False
    Else
      ' Get the remainder of the last segment
      lLen = (36 - Len("\" & sLast))
      For lLoop = 1 To lLen
        sFill = sFill & Mid$(sSegment, lLoop, 1)
      Next
      sLast = sFill & "...\" & sLast
      bLoop = False
    End If
  Loop
  AbbreviateString = sFirst & sLast
End Function
Public Sub AddToMRUList(ByVal sFile As String)
  Dim sMRUList() As String, sMRU As String
  Dim lListSize As Long, lLoop As Long, lDup As Long
  Dim bFound As Boolean

  ' Because our arrays are zero based
  lListSize = MRUSize - 1

  ReDim sMRUList(lListSize)

  ' Retrieve all saved paths
  For lLoop = 0 To lListSize
    sMRU = ReadFromIni("MRU " & lLoop)
    If Len(sMRU) > 0 And Not IsInArray(sMRU, sMRUList) Then
      sMRUList(lLoop) = sMRU
    End If
  Next
  Call FillArrayGaps(sMRUList)

  ' Now make sure this is not something that has already been opened
  If IsInArray(sFile, sMRUList) Then
    ' We need to move it to the top of the list
    lDup = InArraySlot(sFile, sMRUList)         ' Find its location
    For lLoop = lDup To 1 Step -1               ' Move all above down one slot
      sMRUList(lLoop) = sMRUList(lLoop - 1)
      sMRUList(lLoop - 1) = vbNullString
    Next
  Else
    ' Make room at the top of the list
    For lLoop = lListSize To 1 Step -1
      sMRUList(lLoop) = sMRUList(lLoop - 1)
      sMRUList(lLoop - 1) = vbNullString
    Next
  End If

  ' Prep work has been done, add it to the top of the list
  sMRUList(0) = sFile

  ' Now save in the ini file and display the menus
  For lLoop = 0 To lListSize
    If Len(sMRUList(lLoop)) > 0 Then
      Call WriteToIni("MRU " & lLoop, sMRUList(lLoop))
      FormName.mnuFileMRU(lLoop).Caption = "&" & lLoop + 1 & ". " & AbbreviateString(sMRUList(lLoop))
      FormName.mnuFileMRU(lLoop).Visible = True
      bFound = True
    Else
      Call WriteToIni("MRU " & lLoop, vbNullString)
    End If
  Next
  If bFound Then
    FormName.mnuFileMRUSep.Visible = True
  End If

  Erase sMRUList
End Sub
Private Sub FillArrayGaps(ByRef sFill() As String)
  ' Move all data up to the lower numbered array elements
  ' eliminating all empty rows
  Dim lOuterLoop As Long, lFindData As Long

  For lOuterLoop = 0 To UBound(sFill)
    If Len(Trim$(sFill(lOuterLoop))) = 0 Then
      ' Find a row that has data and move it up
      For lFindData = lOuterLoop To UBound(sFill)
        If Len(Trim$(sFill(lFindData))) > 0 Then
          sFill(lOuterLoop) = sFill(lFindData)
          sFill(lOuterLoop) = sFill(lFindData)
          sFill(lFindData) = vbNullString
          sFill(lFindData) = vbNullString
          Exit For
        End If
      Next
    End If
  Next
End Sub
Public Sub GetMRU()
  Dim lListSize As Long, lLoop As Long
  Dim sMRU As String, sMRUList() As String
  Dim bFound As Boolean

  ' Because our arrays are zero based
  lListSize = MRUSize - 1

  ' Prepare the menus (specifically for list shrinking)
  For lLoop = 0 To Max - 1
    FormName.mnuFileMRU(lLoop).Caption = vbNullString
    FormName.mnuFileMRU(lLoop).Visible = False
  Next

  ' Make sure the user wants to see the list
  If lListSize < 0 Then
    FormName.mnuFileMRUSep.Visible = False
    Exit Sub
  End If

  ReDim sMRUList(lListSize)

  ' Retrieve all saved paths
  For lLoop = 0 To lListSize
    sMRU = ReadFromIni("MRU " & lLoop)
    If Len(sMRU) > 0 And Not IsInArray(sMRU, sMRUList) Then
      sMRUList(lLoop) = sMRU
    End If
  Next

  ' Clean up if any have been erased from the ini file
  Call FillArrayGaps(sMRUList)

  For lLoop = 0 To lListSize
    If Len(sMRUList(lLoop)) > 0 Then
      If lLoop = 0 Then
        ' Set the initial directory of the common dialog control
        ' to that of the most recently viewed file
        InitDir = Left$(sMRUList(lLoop), Len(sMRUList(lLoop)) - InStr(1, StrReverse(sMRUList(lLoop)), "\"))
      End If
      FormName.mnuFileMRU(lLoop).Caption = "&" & lLoop + 1 & ". " & AbbreviateString(sMRUList(lLoop))
      FormName.mnuFileMRU(lLoop).Visible = True
      bFound = True
    End If
  Next
  If bFound Then
    FormName.mnuFileMRUSep.Visible = True
  End If
End Sub
Public Sub HideMRU()
  Dim lLoop As Long

  For lLoop = 0 To Max - 1
    FormName.mnuFileMRU(lLoop).Caption = vbNullString
    FormName.mnuFileMRU(lLoop).Visible = False
  Next
  FormName.mnuFileMRUSep.Visible = False
End Sub
Private Function InArraySlot(ByVal sValue As String, ByRef aSearch As Variant) As Long
  ' Determine the slot of an array in which a value resides
  Dim lLoop As Long

  For lLoop = 0 To UBound(aSearch)
    If aSearch(lLoop) = sValue Then
      InArraySlot = lLoop
      Exit For
    End If
  Next
End Function
Private Function IsInArray(ByVal sValue As String, ByRef aSearch As Variant) As Boolean
  ' Determine whether or not a value is in an array
  If Not IsArray(aSearch) Then Exit Function
  IsInArray = InStr(1, vbNullChar & Join(aSearch, vbNullChar) & vbNullChar, vbNullChar & sValue & vbNullChar) > 0
End Function
Public Function ReadFromIni(ByVal sKey As String, Optional sDefaultVal As String = vbNullString) As String
  ' Retrieve a value from an ini file
  Dim sBuffer As String
  Dim lRet As Long

  sBuffer = Space$(255)

  lRet = GetPrivateProfileString(m_sIniArea, sKey, vbNullString, sBuffer, 255, m_sIniFile)
  If lRet = 0 Then
    ReadFromIni = sDefaultVal
  Else
    ReadFromIni = Left$(sBuffer, lRet)
  End If
End Function
Public Sub WriteToIni(ByVal sKey As String, ByVal sValue As String)
  ' Write a value to an ini file
  Call WritePrivateProfileString(m_sIniArea, sKey, sValue, m_sIniFile)
End Sub
Public Property Get FormName() As VB.Form
  Set FormName = m_frmX
End Property
Public Property Let FormName(ByRef frmX As VB.Form)
  Set m_frmX = frmX
End Property
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
  ' If this does not get set, Win.ini will be the file that gets read
  m_sIniFile = sFile
End Property
Public Property Get InitDir() As String
  InitDir = m_sInitDir
End Property
Public Property Let InitDir(ByVal sDir As String)
  m_sInitDir = sDir
End Property
Public Property Get Max() As Long
  Max = m_lMax
End Property
Public Property Let Max(ByVal lMax As Long)
  If lMax <= 0 Then lMax = 1                          ' Don't allow max to be smaller than 1
  m_lMax = lMax
  If m_lMax < m_lMRUSize Then m_lMRUSize = m_lMax     ' Don't allow value to surpass max
End Property
Public Property Get MRUSize() As Long
  MRUSize = m_lMRUSize
End Property
Public Property Let MRUSize(ByVal lSize As Long)
  If lSize > m_lMax Then lSize = m_lMax
  If lSize < 0 Then lSize = 0
  m_lMRUSize = lSize
End Property
