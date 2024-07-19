Attribute VB_Name = "SpecialFolderModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API functions used by this program.
Private Const CSIDL_ADMINTOOLS As Long = &H30
Private Const CSIDL_ALTSTARTUP As Long = &H1D
Private Const CSIDL_APPDATA As Long = &H1A
Private Const CSIDL_BITBUCKET As Long = &HA
Private Const CSIDL_CDBURN_AREA As Long = &H3B
Private Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F
Private Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E
Private Const CSIDL_COMMON_APPDATA As Long = &H23
Private Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Private Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Private Const CSIDL_COMMON_FAVORITES As Long = &H1F
Private Const CSIDL_COMMON_MUSIC As Long = &H35
Private Const CSIDL_COMMON_PICTURES As Long = &H36
Private Const CSIDL_COMMON_PROGRAMS As Long = &H17
Private Const CSIDL_COMMON_STARTMENU As Long = &H16
Private Const CSIDL_COMMON_STARTUP As Long = &H18
Private Const CSIDL_COMMON_TEMPLATES As Long = &H2D
Private Const CSIDL_COMMON_VIDEO As Long = &H37
Private Const CSIDL_COMPUTERSNEARME As Long = &H3D
Private Const CSIDL_CONNECTIONS As Long = &H31
Private Const CSIDL_CONTROLS As Long = &H3
Private Const CSIDL_COOKIES As Long = &H21
Private Const CSIDL_DESKTOP As Long = &H0
Private Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Private Const CSIDL_DRIVES As Long = &H11
Private Const CSIDL_FAVORITES As Long = &H6
Private Const CSIDL_FLAG_CREATE As Long = &H8000
Private Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000
Private Const CSIDL_FONTS As Long = &H14
Private Const CSIDL_HISTORY As Long = &H22
Private Const CSIDL_INTERNET As Long = &H1
Private Const CSIDL_INTERNET_CACHE As Long = &H20
Private Const CSIDL_LOCAL_APPDATA As Long = &H1C
Private Const CSIDL_MYDOCUMENTS As Long = &HC
Private Const CSIDL_MYMUSIC As Long = &HD
Private Const CSIDL_MYPICTURES As Long = &H27
Private Const CSIDL_MYVIDEO As Long = &HE
Private Const CSIDL_NETHOOD As Long = &H13
Private Const CSIDL_NETWORK As Long = &H12
Private Const CSIDL_PERSONAL As Long = &H5
Private Const CSIDL_PHOTOALBUMS As Long = &H45
Private Const CSIDL_PLAYLISTS As Long = &H3F
Private Const CSIDL_PRINTERS As Long = &H4
Private Const CSIDL_PRINTHOOD As Long = &H1B
Private Const CSIDL_PROFILE As Long = &H28
Private Const CSIDL_PROGRAM_FILES As Long = &H26
Private Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B
Private Const CSIDL_PROGRAMS As Long = &H2
Private Const CSIDL_RECENT As Long = &H8
Private Const CSIDL_RESOURCES As Long = &H38
Private Const CSIDL_SAMPLE_MUSIC As Long = &H40
Private Const CSIDL_SAMPLE_PLAYLISTS As Long = &H41
Private Const CSIDL_SAMPLE_PICTURES As Long = &H42
Private Const CSIDL_SAMPLE_VIDEOS As Long = &H43
Private Const CSIDL_SENDTO As Long = &H9
Private Const CSIDL_STARTMENU As Long = &HB
Private Const CSIDL_STARTUP As Long = &H7
Private Const CSIDL_SYSTEM As Long = &H25
Private Const CSIDL_TEMPLATES As Long = &H15
Private Const CSIDL_WINDOWS As Long = &H24
Private Const MAX_PATH As Long = 260

Private Declare Sub CoTaskMemFree Lib "Ole32.dll" (ByVal pvoid As Long)
Private Declare Function SHGetFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long

'This procedure is executed when this program is started.
Private Sub Main()
   MsgBox GetSpecialFolder(CSIDL_DESKTOP), vbInformation
End Sub


'This procedure gets the path to the special folder referred to by the specified CSIDL.
Private Function GetSpecialFolder(CSIDL As Long) As String
Dim Path As String
Dim PidlListOut As Long
Dim ReturnValue As Long

   Path = Space$(MAX_PATH)
   ReturnValue = SHGetFolderLocation(CLng(0), CSIDL, CLng(0), CLng(0), PidlListOut)
   If ReturnValue = 0 Then ReturnValue = SHGetPathFromIDListA(PidlListOut, Path)
   CoTaskMemFree PidlListOut
   
   GetSpecialFolder = Trim$(Path)
End Function


