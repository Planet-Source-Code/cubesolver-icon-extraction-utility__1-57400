Attribute VB_Name = "modHelpFiles"
Option Explicit

'  ________________________________                          _______
' / modHelpFiles                   \________________________/ v1.00 |
' |                                                                 |
' |       Description:  Simple module for generating HTML help      |
' |                     file.                                       |
' |                                                                 |
' |   Original Author:  CubeSolver                                  |
' |      Date Created:  September 26, 2003                          |
' |      OS Tested On:  Windows NT 4 SP 6a, Windows XP              |
' |                  _____________________________                  |
' |_________________/                             \_________________|
'  | °         ° \___________________________________/ °         ° |
'  |              ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯              |
'  |---------------------[ Revision History ]----------------------|
'  | °                                                           ° |
'  | Version  Who         Date          Comment                    |
'  | -------  ----------  ------------  -------------------------- |
'  | 1.00     CubeSolver  Sep 26, 2003  Original version.          |
'  \_______________________________________________________________/
'                                       \ASCII Art by Cubesolver/
'                                        ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Public sHelpTitle As String, sHelpFile As String

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Sub CreateHelpFile(ByVal sName As String)
  Dim iOut As Integer
  Dim sPath As String, sFile As String

  sHelpTitle = sName
  sPath = TempDirIs
  sFile = Format$(Date, "mmddyyyy") & Format$(Time, "hhmmss") & ".html"
  sHelpFile = sPath & sFile

  iOut = FreeFile

  Open sHelpFile For Output As #iOut
    Print #iOut, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">"
    Print #iOut, "<HTML>"
    Print #iOut, "<HEAD>"
    Print #iOut, "<META HTTP-EQUIV=""Content-Type"" Content=""text/html; charset=Windows-1252"">"
    Print #iOut, "<STYLE TYPE=""text/css"">"
    Print #iOut, "<!--"
    Print #iOut, "  BODY          {font-family: Arial, Verdana, Tahoma, Helvetica; font-size: 75%;}"
    Print #iOut, "  TD            {font-size: 80%;}"
    Print #iOut, "  TH            {background-color: #EEEEEE; font-size: 80%;}"
    Print #iOut, "  B             {color: #5080B0;}"
'    Print #iOut, "  H3            {color: #4B0082;}"
    Print #iOut, "  .Red          {color: #A50000;}"
    Print #iOut, "  .TrailerText  {font-size: 80%; text-align: right;}"   'center;}"
    Print #iOut, "-->"
    Print #iOut, "</STYLE>"
    Print #iOut,
    Print #iOut, "  <SCRIPT LANGUAGE=""JavaScript"">"
    Print #iOut, "    function click() {"
    Print #iOut, "      if ((event.button == 2) || (event.button == 3))"
    Print #iOut, "        alert (""Right-click function is not available in Help."");"
    Print #iOut, "    }"
    Print #iOut, "    document.onmousedown = click;"
    Print #iOut, "  </SCRIPT>"
    Print #iOut,
    Print #iOut, "  <TITLE>" & sHelpTitle & "</TITLE>"
    Print #iOut, "</HEAD>"
    Print #iOut,
    Print #iOut, "<BODY BGCOLOR=#FFFFFF TEXT=#000000>"
    Print #iOut,
    Print #iOut, "<H3>Icon Extraction Utility</H3>"
    Print #iOut,
    Print #iOut, "<P>You can use this utility to extract large (32 x 32) or small (16 x 16) icons from executables, dynamic link libraries and Active-X files.</P>"
    Print #iOut,
    Print #iOut, "<P>This table shows the options and functions available in the <B CLASS=""Red"">Icon Extraction</B> utility.</P>"
    Print #iOut,
    Print #iOut, "<TABLE CELLSPACING=5 COLS=2>"
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TH ALIGN=left WIDTH=42%>Use this</TH>"
    Print #iOut, "    <TH ALIGN=left WIDTH=58%>To do this</TH>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Close</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Close any previously opened executable, dynamic link library or Active-X file.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Exit</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Shutdown the icon extraction utility.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Open</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Choose an executable, dynamic link library or Active-X file from which icons will be extracted (assuming any exist). You can also open by dragging and dropping them onto the display area.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Print</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Send the highlighted icon to the printer.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Printer Setup</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Choose the printer and update its configuration.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Deselect Icon</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Remove highlighting from selected icon.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Save Selected Icon</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Saves the highlighted icon, chosen by the user from an open executable, dynamic link library or Active-X file, as a .ico type file onto the chosen disk and folder.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Save All Icons</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Saves all icons in the currently open file. You will be prompted to choose a folder into which the icons are to be saved.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>View Large Icons / Small Icons</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Choose whether to view large or small icons within an executable, dynamic link library or Active-X file.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Click Icon</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Using the left mouse button, click on any single icon in the display area to select it. A red border will appear around the icon signifying that the selection has been made. This action will enable the print and save functions.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut,
    Print #iOut, "  <TR VALIGN=""top"">"
    Print #iOut, "    <TD WIDTH=42%><B>Right-Click Icon</B></TD>"
    Print #iOut, "    <TD WIDTH=58%>Using the right mouse button, click on any single icon in the display area that has been previously selected (indicated by a red border) to deselect it. This action will disable the print and save functions.</TD>"
    Print #iOut, "  </TR>"
    Print #iOut, "</TABLE><BR>"
    Print #iOut,
'    Print #iOut, "<HR COLOR=#660000 SIZE=1 NOSHADE>"
    Print #iOut, "<HR COLOR=#4C4C4C SIZE=1 NOSHADE>"
    Print #iOut,
    Print #iOut, "<DIV CLASS =""TrailerText"">Icon Extraction Help File v1.0</DIV>"
    Print #iOut,
    Print #iOut, "</BODY>"
    Print #iOut, "</HTML>"
  Close #iOut
End Sub
Private Function TempDirIs() As String
  Dim sBuffer As String

  ' Get Temporary path
  sBuffer = Space$(255)
  Call GetTempPath(Len(sBuffer), sBuffer)
  TempDirIs = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
  TempDirIs = TempDirIs & IIf(Right$(TempDirIs, 1) = "\", vbNullString, "\")
End Function
