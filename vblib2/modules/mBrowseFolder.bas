Attribute VB_Name = "mBrowseFolder"
Private Type BROWSEINFO
  hwndOwner As Long
  pidlRoot As Long
  pszDisplayName As Long
  lpszTitle As Long
  ulFlags As Long
  lpfnCallback As Long
  lParam As Long
  iImage As Long
End Type

Public Enum BrowseType
  BrowseForFolders = &H1
  BrowseForComputers = &H1000
  BrowseForPrinters = &H2000
  BrowseForEverything = &H4000
End Enum

Public Enum FolderType
  CSIDL_BITBUCKET = 10
  CSIDL_CONTROLS = 3
  CSIDL_DESKTOP = 0
  CSIDL_DRIVES = 17
  CSIDL_FONTS = 20
  CSIDL_NETHOOD = 12
  CSIDL_NETWORK = 19
  'CSIDL_PERSONAL = 5
  CSIDL_PRINTERS = 4
  CSIDL_PROGRAMS = 2
  CSIDL_RECENT = 8
  CSIDL_SENDTO = 9
  CSIDL_STARTMENU = 11
End Enum

Private Const MAX_PATH = 260
#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If

#If VBA7 Then
Private Declare  PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem  As LongPtr)
Private Declare  PtrSafe Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare  PtrSafe Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Private Declare  PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList  As LongPtr, ByVal lpBuffer As String) As Long
Private Declare  PtrSafe Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner  As LongPtr, ByVal nFolder  As LongPtr, ListId  As LongPtr) As Long
#Else
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ListId As Long) As Long
#End If
Public PathTujuan As String

Public Function BrowseFolders(hwndOwner As Long, sMessage As String, Browse As BrowseType, ByVal RootFolder As FolderType) As String
  Dim Nullpos As Integer
  Dim lpIDList As Long
  Dim res As Long
  Dim sPath As String
  Dim BInfo As BROWSEINFO
  Dim RootID As Long

  SHGetSpecialFolderLocation hwndOwner, RootFolder, RootID
  BInfo.hwndOwner = hwndOwner
  BInfo.lpszTitle = lstrcat(sMessage, "")
  BInfo.ulFlags = Browse
  If RootID <> 0 Then BInfo.pidlRoot = RootID
  lpIDList = SHBrowseForFolder(BInfo)
  If lpIDList <> 0 Then
    sPath = String(MAX_PATH, 0)
    res = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    Nullpos = InStr(sPath, vbNullChar)
    If Nullpos <> 0 Then
      sPath = Left(sPath, Nullpos - 1)
    End If
  End If
  BrowseFolders = sPath
  PathTujuan = BrowseFolders
End Function

