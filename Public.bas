Attribute VB_Name = "modPublic"
Option Explicit

Public bClearMRU As Boolean
Public cIni As clsIni

Private Type BrowseInfo
  hWndOwner As Long
  pidlRoot As Long
  pszDisplayName  As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Function BrowseForFolder(ByRef frmX As VB.Form) As String
  Dim iNull As Integer, lpIDList As Long
  Dim sPath As String, tBI As BrowseInfo

  Const BIF_RETURNONLYFSDIRS As Long = &H1
  Const MAX_PATH As Long = 260

  With tBI
    ' Set the owner window
    .hWndOwner = frmX.hWnd
    .lpszTitle = "Select the folder containing the images you wish to view"
    ' Return only if the user selected a directory
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With

  ' Show the 'Browse for folder' dialog
  lpIDList = SHBrowseForFolder(tBI)
  If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    ' Get the path from the IDList
    Call SHGetPathFromIDList(lpIDList, sPath)
    ' Free the block of memory
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then
      BrowseForFolder = Left$(sPath, iNull - 1)
    End If
  End If
End Function
