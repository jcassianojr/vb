Attribute VB_Name = "mOpenWith"
Option Explicit

'Module: mOpenWith
'Version: 0.1
'Author: fafalone
'Purpose: Vista and above provides an API call for the Open With dialog, which offers more options
'         than the previous typical method of using rundll

Public Const S_OK = 0           ' indicates success
Public Const S_FALSE = 1&   ' special HRESULT value

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

Public Declare Function SHOpenWithDialog Lib "shell32" (ByVal hWnd As Long, poainfo As OPENASINFO) As Long

Public Enum OPEN_AS_INFO_FLAGS
  OAIF_ALLOW_REGISTRATION = &H1  'Enable the "always use this program" checkbox. If not passed, it will be disabled.
  OAIF_REGISTER_EXT = &H2  'Do the registration after the user hits the OK button.
  OAIF_EXEC = &H4  'Execute file after registering.
  OAIF_FORCE_REGISTRATION = &H8  'Force the Always use this program checkbox to be checked. Typically, you won't use the OAIF_ALLOW_REGISTRATION flag when you pass this value.
  OAIF_HIDE_REGISTRATION = &H20  'Introduced in Windows Vista. Hide the Always use this program checkbox. If this flag is specified, the OAIF_ALLOW_REGISTRATION and OAIF_FORCE_REGISTRATION flags will be ignored.
  OAIF_URL_PROTOCOL = &H40  'Introduced in Windows Vista. The value for the extension that is passed is actually a protocol, so the Open With dialog box should show applications that are registered as capable of handling that protocol.
  OAIF_FILE_IS_URI = &H80  'Introduced in Windows 8. The location pointed to by the pcszFile parameter is given as a URI.
End Enum

Public Type OPENASINFO
  pcszFile As Long
  pcszClass As Long  'file type description for registering the type with 'always open', if not set uses extension, as in 'XYZ File'
  oafInFlags As OPEN_AS_INFO_FLAGS
End Type

' Esta funcao abre o openwith do windows deixando escolher qual aplicativos cadastrados para a extensao
' nao pode ser trocada por shellexecute pois ela abre o programa padrao para a extensao
' Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd 'escolhe o aplicativo da extensao
' ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.hWnd 'abre o aplicativo padrao da extensao


Public Declare Function GetBinaryType Lib "kernel32" Alias "GetBinaryTypeW" (ByVal lpApplicationName As Long, lpBinaryType As Long) As Long
Public Const SCS_32BIT_BINARY = 0
Public Const SCS_64BIT_BINARY = 6
Public Const SCS_DOS_BINARY = 1
Public Const SCS_OS216_BINARY = 5
Public Const SCS_PIF_BINARY = 3
Public Const SCS_POSIX_BINARY = 4
Public Const SCS_WOW_BINARY = 2

Public Function OpenWith(sFile As String, lFlags As OPEN_AS_INFO_FLAGS, Optional hWndParent As Long, Optional sClass As String) As Long
  Dim oai As OPENASINFO
  oai.pcszFile = StrPtr(sFile)
  oai.oafInFlags = lFlags
  If sClass <> "" Then oai.pcszClass = StrPtr(sClass)
  OpenWith = SHOpenWithDialog(hWndParent, oai)
End Function

Public Function IsExecutable(sFile As String) As Boolean
  Dim lType As Long
  Dim hr As Long

  hr = GetBinaryType(StrPtr(sFile), lType)
  Debug.Print "hr=" & hr & "; lType=" & lType
  If hr = 0 Then
    IsExecutable = False
  Else
    IsExecutable = True
  End If

End Function
Public Sub DisplayOpenWith(strFile As String)

'***PURPOSE: DISPLAY OPEN WITH DIALOG:
' PASS IT A FILE NAME
' e.g., DisplayOpenWith "C:\FileWithNoDefaultApplication.bvq"
'**************************************

  On Error Resume Next
  Shell "rundll32.exe shell32.dll, OpenAs_RunDLL " & strFile

End Sub
