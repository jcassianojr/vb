Attribute VB_Name = "mFileAPI"
Option Explicit
'
'ISextensao(cARQ, cEXT As String) verifica extensao arquivo Ex. ("teste.txt",".txt") Ret. true
'TrocaExt(ByVal cARQ As Variant, ByVal cEXT As String) As String
'EXTENSAO(ByVal eARQ As String) As String
'NomeArq(ByVal eARQ As Variant, Optional ByVal lTIRAEXT As Boolean = False) As String
'CreateNewDirectory(ByVal NewDirectory As String)
'FileOpen(frmOwner As Form, _
'                         Optional ByVal sFilters As String, _
'                         Optional ByVal nFilterIndex As Long = 1, _
'                         Optional ByVal sDefaultFileName As String, _
'                         Optional ByVal sDefaultExtension As String, _
'                         Optional ByVal sStartingFolder As String, _
'                         Optional ByVal sTitle As String) As String
'FileSave(frmOwner As Form, _
'                         Optional ByVal sFilters As String, _
'                         Optional ByVal nFilterIndex As Long, _
'                         Optional ByVal sDefaultExtension As String, _
'                         Optional ByVal sDefaultFileName As String, _
'                         Optional ByVal sStartingFolder As String, _
'                         Optional ByVal sTitle As String) As String
'OpenArqExt(oFORM As Form, ByVal cARQ As String, ByVal cEXT As String, ByVal cTITULO As String) As String
'SaveArqExt(oFORM As Form, ByVal cARQ As String, ByVal cEXT As String, ByVal cTITULO As String) As String
'ImgFILTER() As String
'ImgFILTER2() As String

'parsefile(ByVal archivo As String, ByVal parte As String) As String
' c:\Users\jcass\Downloads\lembra.txt
'"D" 'drive                       c:
'"R" 'raiz                        c:\Users\jcass\Downloads\
'"C" 'completo sem extensao       c:\Users\jcass\Downloads\lembra
'"N" 'SO Nome -sem extensao       lembra
'"A" 'Arquivo nome+exensao        lembra.TXT
'"E" ' Extensao                   TXT
'"PD"'parente dir                 Downloads
'"PP"'Parente path                c:\Users\jcass\Downloads

'FixPath(ByVal cARQ As String) As String
'ShortSpec(ByVal sFileSpec As String) As String
'OpenStreamFile(FileName$, Mode%, RLock%, RecordLen%) As Integer
'FileErrors(ErrVal As Integer) As Integer
Public Const OFN_ALLOWMULTISELECT As Long = &H200
Public Const OFN_CREATEPROMPT As Long = &H2000
Public Const OFN_ENABLEHOOK As Long = &H20
Public Const OFN_ENABLETEMPLATE As Long = &H40
Public Const OFN_ENABLETEMPLATEHANDLE As Long = &H80
Public Const OFN_EXPLORER As Long = &H80000
Public Const OFN_EXTENSIONDIFFERENT As Long = &H400
Public Const OFN_FILEMUSTEXIST As Long = &H1000
Public Const OFN_HIDEREADONLY As Long = &H4
Public Const OFN_LONGNAMES As Long = &H200000
Public Const OFN_NOCHANGEDIR As Long = &H8
Public Const OFN_NODEREFERENCELINKS As Long = &H100000
Public Const OFN_NOLONGNAMES As Long = &H40000
Public Const OFN_NONETWORKBUTTON As Long = &H20000
Public Const OFN_NOREADONLYRETURN As Long = &H8000&  'see comments
Public Const OFN_NOTESTFILECREATE As Long = &H10000
Public Const OFN_NOVALIDATE As Long = &H100
Public Const OFN_OVERWRITEPROMPT As Long = &H2
Public Const OFN_PATHMUSTEXIST As Long = &H800
Public Const OFN_READONLY As Long = &H1
Public Const OFN_SHAREAWARE As Long = &H4000
Public Const OFN_SHAREFALLTHROUGH As Long = 2
Public Const OFN_SHAREWARN As Long = 0
Public Const OFN_SHARENOWARN As Long = 1
Public Const OFN_SHOWHELP As Long = &H10
Public Const OFS_MAXPATHNAME As Long = 260

Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_ALLOWUNDO = &H40

Private Const FO_MOVE As Long = &H1
Private Const FO_RENAME As Long = &H4



Private Const GCT_LFNCHAR = &H1
Private Const GCT_SHORTCHAR = &H2

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
       Or OFN_LONGNAMES _
       Or OFN_CREATEPROMPT _
       Or OFN_NODEREFERENCELINKS

Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
       Or OFN_LONGNAMES _
       Or OFN_OVERWRITEPROMPT _
       Or OFN_HIDEREADONLY


Public Type OPENFILENAME
  nStructSize As Long
  hWndOwner As Long
  hInstance As Long
  SFILTER As String
  sCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  sFile As String
  nMaxFile As Long
  sFileTitle As String
  nMaxTitle As Long
  sInitialDir As String
  sDialogTitle As String
  Flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  sDefFileExt As String
  nCustData As Long
  fnHook As Long
  sTemplateName As String
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAborted As Boolean
  hNameMaps As Long
  sProgress As String
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type



Public OFN As OPENFILENAME

Enum TFileFlags
    FOF_ALLOWUNDO = &H40
    FOF_FILESONLY = &H80
    FOF_MULTIDESTFILES = &H1
    FOF_NOCONFIRMATION = &H10
    FOF_NOCONFIRMMKDIR = &H200
    FOF_SIMPLEPROGRESS = &H100
    FOF_SILENT = &H4
End Enum

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
Public Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" _
                                        Alias "GetOpenFileNameA" _
                                        (pOpenfilename As OPENFILENAME) As Long

Public Declare PtrSafe Function GetSaveFileName Lib "comdlg32.dll" _
                                        Alias "GetSaveFileNameA" _
                                        (pOpenfilename As OPENFILENAME) As Long

Public Declare PtrSafe Function GetShortPathName Lib "kernel32" _
                                         Alias "GetShortPathNameA" _
                                         (ByVal lpszLongPath As String, _
                                          ByVal lpszShortPath As String, _
                                          ByVal cchBuffer As LongPtr) As Long

Private Declare PtrSafe Function SHFileOperation _
                          Lib "shell32.dll" Alias "SHFileOperationA" ( _
                              lpFileOp As SHFILEOPSTRUCT _
                            ) As Long

Public Declare PtrSafe Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As LongPtr, ByVal lpszShortPath As LongPtr, ByVal cchBuffer As LongPtr) As Long

Private Declare PtrSafe Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Longptr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32.dll" () As Long


#Else

Private Declare Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Long) As Long

Public Declare Function GetOpenFileName Lib "comdlg32.dll" _
                                        Alias "GetOpenFileNameA" _
                                        (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetSaveFileName Lib "comdlg32.dll" _
                                        Alias "GetSaveFileNameA" _
                                        (pOpenfilename As OPENFILENAME) As Long

Public Declare Function GetShortPathName Lib "kernel32" _
                                         Alias "GetShortPathNameA" _
                                         (ByVal lpszLongPath As String, _
                                          ByVal lpszShortPath As String, _
                                          ByVal cchBuffer As Long) As Long

Private Declare Function SHFileOperation _
                          Lib "shell32.dll" Alias "SHFileOperationA" ( _
                              lpFileOp As SHFILEOPSTRUCT _
                            ) As Long

Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
#End If

Public Function SH_Copy(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Copy files or folders
    shFileOp.hwnd = GetDesktopWindow
    shFileOp.wFunc = FO_COPY
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Copy = SHFileOperation(shFileOp)
    
End Function
Public Function SH_Move(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Move files or folders
    shFileOp.hwnd = GetDesktopWindow
    shFileOp.wFunc = FO_MOVE
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Move = SHFileOperation(shFileOp)
    
End Function

Public Function SH_Rename(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Rename files or folders
    shFileOp.hwnd = GetDesktopWindow
    shFileOp.wFunc = FO_RENAME
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Rename = SHFileOperation(shFileOp)
    
End Function

Public Function SH_Delete(src As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Delete files or folders
    shFileOp.hwnd = GetDesktopWindow
    shFileOp.wFunc = FO_DELETE
    shFileOp.pFrom = src
    shFileOp.fFlags = cFlags
    
    SH_Delete = SHFileOperation(shFileOp)
    
End Function

Public Function FileExists(ByRef sFileName As String) As Boolean
    On Error Resume Next
    FileExists = PathFileExists(sFileName)
End Function


Public Function IsExtensao(ByVal cARQ As String, cEXT As String) As Boolean
 IsExtensao = False
  cARQ = UCase(FixStr(cARQ))
  cEXT = UCase(cEXT)
   If parsefile(cARQ, "E") = cEXT Then
      IsExtensao = True
   End If
End Function
Public Function TrocaExt(ByVal cARQ As Variant, ByVal cEXT As String) As String
  TrocaExt = parsefile(cARQ, "C") & "." & cEXT
End Function

Public Function EXTENSAO(ByVal eARQ As String) As String
  EXTENSAO = parsefile(eARQ, "E")
End Function
Public Function NomeArq(ByVal eARQ As Variant, Optional ByVal lTIRAEXT As Boolean = False) As String
  If lTIRAEXT Then
     NomeArq = parsefile(eARQ, "N")
  Else
     NomeArq = parsefile(eARQ, "A")
  End If
End Function
Public Function CopyFileWindowsWay(ByVal SourceFile As String, ByVal DestinationFile As String, Optional ByVal lAPAGA As Boolean = False) As Long
  Dim lngReturn As Long
  Dim typFileOperation As SHFILEOPSTRUCT
  If PathFileExists(DestinationFile) Then
      If lAPAGA Then
          DeleteFile DestinationFile  'Kill DestinationFile
      End If
      Exit Function
  End If
  With typFileOperation
    .hwnd = 0
    .wFunc = FO_COPY
    .pFrom = SourceFile & vbNullChar & vbNullChar  'source file
    .pTo = DestinationFile & vbNullChar & vbNullChar  'destination file
    .fFlags = FO_ALLOWUNDO
  End With
  lngReturn = SHFileOperation(typFileOperation)

  CopyFileWindowsWay = lngReturn

  If lngReturn <> 0 Then                       'Operation failed
    MsgBox "Copiando " & SourceFile & " " & DestinationFile
  Else                                         'Aborted
    If typFileOperation.fAnyOperationsAborted = True Then
      MsgBox "Copiando " & SourceFile & " " & DestinationFile
    End If
  End If
End Function

Public Sub CreateNewDirectory(ByVal NewDirectory As String)
  Dim sDirTest As String
  Dim SecAttrib As SECURITY_ATTRIBUTES
  Dim bSuccess As Boolean
  Dim sPath As String
  Dim iCounter As Integer
  Dim sTempDir As String
  Dim iFlag As Integer
  iFlag = 0
  sPath = NewDirectory

  If Right(sPath, Len(sPath)) <> "\" Then
    sPath = sPath & "\"
  End If

  iCounter = 1
  Do Until InStr(iCounter, sPath, "\") = 0
    iCounter = InStr(iCounter, sPath, "\")
    sTempDir = Left(sPath, iCounter)
    sDirTest = Dir(sTempDir)
    iCounter = iCounter + 1

    SecAttrib.lpSecurityDescriptor = &O0
    SecAttrib.bInheritHandle = False
    SecAttrib.nLength = Len(SecAttrib)
    bSuccess = CreateDirectory(sTempDir, SecAttrib)
  Loop
End Sub

Public Function GetTempDirectory() As String
  Dim tempPath As String
  Dim sLen As Integer
  tempPath = String(255, 0)
  sLen = GetTempPath(256, tempPath)
  tempPath = Left(tempPath, sLen)
  GetTempDirectory = tempPath
End Function

Public Function FileOpen(frmOwner As Form, _
                         Optional ByVal sFilters As String, _
                         Optional ByVal nFilterIndex As Long = 1, _
                         Optional ByVal sDefaultFileName As String, _
                         Optional ByVal sDefaultExtension As String, _
                         Optional ByVal sStartingFolder As String, _
                         Optional ByVal sTitle As String) As String



  With OFN
    .nStructSize = Len(OFN)
    .hWndOwner = CLng(frmOwner.hwnd)
    .SFILTER = sFilters & vbNullChar & vbNullChar
    .nFilterIndex = nFilterIndex
    .sFile = sDefaultFileName & Space$(1024) & vbNullChar & vbNullChar
    .nMaxFile = Len(.sFile)
    .sDefFileExt = sDefaultExtension & vbNullChar & vbNullChar
    .sFileTitle = vbNullChar & Space$(512) & vbNullChar & vbNullChar
    .nMaxTitle = Len(OFN.sFileTitle)
    .sInitialDir = sStartingFolder & vbNullChar & vbNullChar
    .sDialogTitle = sTitle
    .Flags = OFS_FILE_OPEN_FLAGS
  End With

  If GetOpenFileName(OFN) Then
    FileOpen = TrimNull(OFN.sFile)
  End If

End Function

Public Function FileSave(frmOwner As Form, _
                         Optional ByVal sFilters As String, _
                         Optional ByVal nFilterIndex As Long, _
                         Optional ByVal sDefaultExtension As String, _
                         Optional ByVal sDefaultFileName As String, _
                         Optional ByVal sStartingFolder As String, _
                         Optional ByVal sTitle As String) As String



  With OFN

    .nStructSize = Len(OFN)
    .hWndOwner = CLng(frmOwner.hwnd)
    .SFILTER = sFilters & vbNullChar & vbNullChar
    .nFilterIndex = nFilterIndex
    .sFile = sDefaultFileName & Space$(1024) & _
             vbNullChar & vbNullChar
    .nMaxFile = Len(.sFile)
    .sDefFileExt = sDefaultExtension & vbNullChar & vbNullChar
    .sFileTitle = vbNullChar & Space$(512) & _
                  vbNullChar & vbNullChar
    .nMaxTitle = Len(OFN.sFileTitle)
    .sInitialDir = sStartingFolder & vbNullChar & vbNullChar
    .sDialogTitle = sTitle
    .Flags = OFS_FILE_SAVE_FLAGS

  End With

  If GetSaveFileName(OFN) Then
    FileSave = TrimNull(OFN.sFile)
  End If

End Function

Function OpenArqExt(oFORM As Form, ByVal cARQ As String, ByVal cEXT As String, ByVal cTITULO As String) As String
  Dim sFileName As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim SFILTER As String
  OpenArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  SFILTER = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileOpen(oFORM, SFILTER, 1, sRECENTFILE, cEXT, sPath, "Escolher " & cTITULO)
  If Len(sFileName) = 0 Then
    lRETU = False
    Exit Function
  End If
  lRETU = True
  OpenArqExt = sFileName
End Function

Function SaveArqExt(oFORM As Form, ByVal cARQ As String, ByVal cEXT As String, ByVal cTITULO As String) As String
  Dim sFileName As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim SFILTER As String
  SaveArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  SFILTER = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileSave(oFORM, SFILTER, 1, cEXT, "Novo", sPath, cTITULO)
  If Len(sFileName) = 0 Then
    lRETU = False
    Exit Function
  End If
  lRETU = True
  SaveArqExt = sFileName
End Function

Public Function ImgFILTER() As String
  Dim SFILTER As String
  SFILTER = "Windows ou OS/2 Bitmap File (*.BMP)" & vbNullChar & "*.BMP" & vbNullChar
  SFILTER = SFILTER & "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  SFILTER = SFILTER & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  SFILTER = SFILTER & "Tagged Imag (*.TIF, *.TIFF)" & vbNullChar & "*.TIF;*.TIFF" & vbNullChar
  SFILTER = SFILTER & "Zsoft Paintbrush PCX bitmap format (*.PCX)" & vbNullChar & "*.PCX" & vbNullChar
  SFILTER = SFILTER & "Windows Icon (*.ICO)" & vbNullChar & "*.ICO" & vbNullChar
  SFILTER = SFILTER & "Graphics Interchange Format (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  SFILTER = SFILTER & "Adobe Photoshop (*.PSD)" & vbNullChar & "*.PSD" & vbNullChar
  SFILTER = SFILTER & "Dr. Halo (*.CUT)" & vbNullChar & "*.CUT" & vbNullChar
  SFILTER = SFILTER & "DirectDraw Surface (*.DDS)" & vbNullChar & "*.DDS" & vbNullChar
  SFILTER = SFILTER & "High Dynamic Range (*.HDR)" & vbNullChar & "*.HDR" & vbNullChar
  SFILTER = SFILTER & "Amiga IFF (*.IFF, *.LBM)" & vbNullChar & "*.IFF;*.LBM" & vbNullChar
  SFILTER = SFILTER & "JPEG Network Graphics (*.JNG)" & vbNullChar & "*.JPG" & vbNullChar
  SFILTER = SFILTER & "Commodore 64 Koala format (*.KOA)" & vbNullChar & "*.KOA" & vbNullChar
  SFILTER = SFILTER & "Multiple Network Graphics (*.MNG)" & vbNullChar & "*.MNG" & vbNullChar
  SFILTER = SFILTER & "Portable Bitmap (ASCII) (*.PBM)" & vbNullChar & "*.PBM" & vbNullChar
  SFILTER = SFILTER & "Portable Bitmap (BINARY) (*.PBM)" & vbNullChar & "*.OBM" & vbNullChar
  SFILTER = SFILTER & "Kodak PhotoCD (*.PCD)" & vbNullChar & "*.PCD" & vbNullChar
  SFILTER = SFILTER & "Portable Graymap (ASCII) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  SFILTER = SFILTER & "Portable Graymap (BINARY) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  SFILTER = SFILTER & "Portable Pixelmap (ASCII) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  SFILTER = SFILTER & "Portable Pixelmap (BINARY) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  SFILTER = SFILTER & "Sun Rasterfile (*.RAS)" & vbNullChar & "*.RAS" & vbNullChar
  SFILTER = SFILTER & "Truevision Targa files (*.TGA, *.TARGA)" & vbNullChar & "*.TGA" & vbNullChar
  SFILTER = SFILTER & "Wireless Bitmap (*.WBMP)" & vbNullChar & "*.WBMP" & vbNullChar
  SFILTER = SFILTER & "X11 Bitmap Format (*.XBM)" & vbNullChar & "*.XBM" & vbNullChar
  SFILTER = SFILTER & "X11 Pixmap Format (*.XPM)" & vbNullChar & "*.XPM" & vbNullChar
  SFILTER = SFILTER & "All Files" & vbNullChar & "*.*"
  ImgFILTER = SFILTER
End Function

Public Function ImgFILTER2() As String
  Dim SFILTER As String
  SFILTER = "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  SFILTER = SFILTER & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  SFILTER = SFILTER & "GIF (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  ImgFILTER2 = SFILTER
End Function

Public Function parsefile(ByVal archivo As String, ByVal parte As String) As String
  Dim PosSlash As Integer
  Dim PosPunto As Integer
  Dim PosColon As Integer
  Dim LastSlash As Integer
  Dim i As Integer
  Dim C As String
  PosSlash = 0
  LastSlash = 0                                'un slash antes del ultimo
  PosPunto = 0
  PosColon = 0
  For i = 1 To Len(archivo)
    C = Mid(archivo, i, 1)
    If C = "." Then PosPunto = i             'Posicion del ULTIMO punto
    If C = "\" Then
      LastSlash = PosSlash
      PosSlash = i                         'Posicion del ULTIMO backslash
    End If
    If C = ":" Then PosColon = i             'Posicion del :
  Next
  Select Case UCase(parte)
  Case "D"                                     'drive
    parsefile = Mid(archivo, 1, PosColon)
  Case "R"                                     'raiz
    parsefile = Mid(archivo, 1, PosSlash)
  Case "C"                                     'completo sem extensao
    parsefile = Mid(archivo, 1, PosPunto - 1)
  Case "N"                                     'SO Nome -sem extensao
    parsefile = Mid(archivo, PosSlash + 1, PosPunto - PosSlash - 1)
  Case "A"                                     'Arquivo nome+exensao
    parsefile = Mid(archivo, PosSlash + 1)
  Case "E"                                     ' Extensao
    parsefile = Mid(archivo, PosPunto + 1)
  Case "PD"                                    'parente dir
    If LastSlash > 0 Then
      parsefile = Mid(archivo, LastSlash + 1, PosSlash - LastSlash - 1)
    Else
      parsefile = Mid(archivo, 1, PosSlash)
    End If
  Case "PP"                                    'Parente path
    parsefile = Mid(archivo, 1, PosSlash - 1)
  Case Else
    parsefile = ""
  End Select
End Function

Function FixPath(ByVal cARQ As String) As String
  If Right(cARQ, 1) = "\" Then
    FixPath = cARQ
  Else
    FixPath = cARQ + "\"
  End If
End Function
Function GetDriveUNC(DriveString As String) As String
'GetDriveUNC by Gavin Bollard 2000
'---------------------------------
'This function is designed to be used when you want
'your program to be network aware and be able to use
'either the drive letter reference or the UNC Name.
'For example when creating a CD that needs to
'reference it's own drive letter, you might use this
'code to anticipate it being shared on a Network CD
'Tower or Drive.
'
'Example Usage:  sDriveLetter = GetDriveUNC(App.Path)
'
'Reads a String and Returns either...
'1. Drive Letter,Colon, backslash
'2. UNC Name ending in Backslash
'3. Empty String (if not a drive letter or UNC Name)

  Dim DriveText As String
  Dim ThirdSlashPos As Integer
  Dim FourthSlashPos As Integer

  DriveString = Trim$(DriveString)
  If Mid$(DriveString, 2, 1) = ":" Then
    DriveText = Left$(DriveString, 2) + "\"
  Else
    If Left$(DriveString, 2) = "\\" Then
      ThirdSlashPos = InStr(3, DriveString, "\", _
                            vbTextCompare)
      FourthSlashPos = InStr(ThirdSlashPos + 1, _
                             DriveString, "\", vbTextCompare)
      If FourthSlashPos > 5 Then
        DriveText = Left$(DriveString, FourthSlashPos)
      Else
        If (FourthSlashPos = 0) And (ThirdSlashPos > 3) _
           Then
          DriveText = DriveString + "\"
        Else
          DriveText = ""
        End If
      End If
    Else
      DriveText = ""
    End If
  End If
  GetDriveUNC = DriveText
End Function
Public Function ShortSpec(ByVal sFileSpec As String) As String
' This works with drive assignments or UNC paths.
' If it fails, an empty string is returned.
' Not designed to work with "relative" or "partial" paths.
  Dim lRet As Long
  Dim bUnc As Boolean
  '
  If Left$(sFileSpec, 2) = "\\" Then
    sFileSpec = "\\?\UNC\" & Mid$(sFileSpec, 3)
    bUnc = True
  Else
    sFileSpec = "\\?\" & sFileSpec
  End If
  lRet = GetShortPathNameW(StrPtr(sFileSpec), 0, 0)
  If lRet Then
    ShortSpec = Space$(lRet - 1)
    lRet = GetShortPathNameW(StrPtr(sFileSpec), StrPtr(ShortSpec), lRet)
    If lRet Then
      If bUnc Then
        ShortSpec = "\\" & Mid$(ShortSpec, 9)  ' Strip the "\\?\UNC\".
      Else
        ShortSpec = Mid$(ShortSpec, 5)   ' Strip the "\\?\".
      End If
    Else
      ShortSpec = vbNullString
    End If
  End If
End Function
Public Function FolderExists(sDir As String) As Boolean
  Dim s As String
  s = sDir
  If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1)
  On Error GoTo FileExistsError
  ' If no error then something existed.
  FolderExists = ((GetAttr(s) And vbDirectory) = vbDirectory)
  Exit Function
FileExistsError:
  FolderExists = False
  Exit Function
End Function
Public Function OpenStreamFile(filename$, Mode%, RLock%, RecordLen%) As Integer
  Const REPLACEFILE = 1, READAFILE = 2, ADDTOFILE = 3
  Const RANDOMFILE = 4, BINARYFILE = 5
  Const NOLOCK = 0, RDLOCK = 1, WRLOCK = 2, RWLOCK = 3
  Dim FileNum%
  Dim Action%
  FileNum% = FreeFile
  On Error GoTo OpenErrors
  Select Case Mode
  Case REPLACEFILE
    Select Case RLock%
    Case NOLOCK
      Open filename For Output Shared As FileNum%
    Case RDLOCK
      Open filename For Output Lock Read As FileNum%
    Case WRLOCK
      Open filename For Output Lock Write As FileNum%
    Case RWLOCK
      Open filename For Output Lock Read Write As FileNum%
    End Select
  Case READAFILE
    Select Case RLock%
    Case NOLOCK
      Open filename For Input Shared As FileNum%
    Case RDLOCK
      Open filename For Input Lock Read As FileNum%
    Case WRLOCK
      Open filename For Input Lock Write As FileNum%
    Case RWLOCK
      Open filename For Input Lock Read Write As FileNum%
    End Select
  Case ADDTOFILE
    Select Case RLock%
    Case NOLOCK
      Open filename For Append Shared As FileNum%
    Case RDLOCK
      Open filename For Append Lock Read As FileNum%
    Case WRLOCK
      Open filename For Append Lock Write As FileNum%
    Case RWLOCK
      Open filename For Append Lock Read Write As FileNum%
    End Select
  Case RANDOMFILE
    Select Case RLock%
    Case NOLOCK
      Open filename For Random Shared As FileNum% Len = RecordLen%
    Case RDLOCK
      Open filename For Random Lock Read As FileNum% Len = RecordLen%
    Case WRLOCK
      Open filename For Random Lock Write As FileNum% Len = RecordLen%
    Case RWLOCK
      Open filename For Random Lock Read Write As FileNum% Len = RecordLen%
    End Select
  Case BINARYFILE
    Select Case RLock%
    Case NOLOCK
      Open filename For Binary Shared As FileNum%
    Case RDLOCK
      Open filename For Binary Lock Read As FileNum%
    Case WRLOCK
      Open filename For Binary Lock Write As FileNum%
    Case RWLOCK
      Open filename For Binary Lock Read Write As FileNum%
    End Select
  Case Else
    Exit Function
  End Select
  OpenStreamFile = FileNum%
  Exit Function
OpenErrors:
  Action% = FileErrors(Err)
  Select Case Action%
  Case 0
    Resume                                   'Resumes at line where ERROR occured
  Case 1
    Resume Next                              'Resumes at line after ERROR
  Case 2
    OpenStreamFile = 0                       'Unrecoverable ERROR-reports error, exits function with error code
    Exit Function
  Case Else
    'MsgBox error$(Err) + vbCrLf + "After line " + CStr(Erl) + vbCrLf + "Program will TERMINATE!"

    'Unrecognized ERROR-reports error and terminates.
    'End
  End Select
End Function

Private Function FileErrors(ErrVal As Integer) As Integer
'Return Value 0=Resume,              1=Resume Next,
'             2=Unrecoverable Error, 3=Unrecognized Error
  Dim MsgType%
  Dim Msg$
  Dim Response%
  MsgType% = 48
  Select Case ErrVal
  Case 68
    Msg$ = "That device appears Unavailable."
    MsgType% = MsgType% + 4
  Case 71
    Msg$ = "Insert a Disk in the Drive"
  Case 53
    Msg$ = "Cannot Find File"
    MsgType% = MsgType% + 5
  Case 57
    Msg$ = "Internal Disk Error."
    MsgType% = MsgType% + 4
  Case 61
    Msg$ = "Disk is Full.  Continue?"
    MsgType% = 35
  Case 64, 52
    Msg$ = "That Filename is Illegal!"
    MsgType% = MsgType% + 5
  Case 70
    Msg$ = "File in use by another user!"
    MsgType% = MsgType% + 5
  Case 76
    Msg$ = "Path does not Exist!"
    MsgType% = MsgType% + 2
  Case 54
    Msg$ = "Bad File Mode!"
  Case 55
    Msg$ = "File is Already Open."
  Case 62
    Msg$ = "Read Attempt Past End of File."
  Case Else
    FileErrors = 3
    Exit Function
  End Select
  Response% = MsgBox(Msg$, MsgType%, "Disk Error")
  Select Case Response%
  Case 1, 4
    FileErrors = 0
  Case 5
    FileErrors = 1
  Case 2, 3
    FileErrors = 2
  Case Else
    FileErrors = 3
  End Select
End Function


Public Function DeleteFile(ByVal sFileName As String, Optional vRecycleBin As Boolean) As Boolean
  Dim fo As SHFILEOPSTRUCT

  With fo
    .pFrom = sFileName & vbNullChar
    .wFunc = FO_DELETE
    If vRecycleBin Then .fFlags = FO_ALLOWUNDO    ' Send to RecycleBin
  End With

  Call SHFileOperation(fo)
  DeleteFile = Abs(fo.fAborted) - 1
End Function


'*****************************************
' ValidFileName: returns a valid file name String

' The cases where an invalid file name can be returned are:
' 1) The DefaultFileName parameter is a null String ("") and ProposedFileName was completely invalid or also a null string.
' 2) The DefaultFileName parameter is a null String ("") and ProposedFileName was only an extension (".something").
' In both above cases it will return a null string (""). So if you set DefaultFileName to "" you should check that the returned value is not "". In all other cases some valid file name will be returned.

' Parameters
' ProposedFileName: the String that is proposed to use as file name
' ReplacementChar (Optional): a character to be used as a replacement for invalid characters, the default is nothing (a null string)
' DefaultFileName (Optional): the file name that will be used when the is not file name
' ForOldFileFormat_8Dot3 (Optional): set it to True in the case that you need to support the old file format convention that was used in the D.O.S. or if for some reason you want to restrict the file name to that format
' AllowExtension (Optional): specifies if the string can supply not only the file name but also the file extension. The default is True.
' [out] HasExtension (Optional): It is a return value. It is passed ByRef, it returns if the file name included an Extension.
'*****************************************
Public Function ValidFileName(ByVal ProposedFileName As String, Optional ByVal ReplacementChar As String = "", Optional DefaultFileName As String = "Untitled", Optional ForOldFileFormat_8Dot3 As Boolean = False, Optional AllowExtension As Boolean = True, Optional ByRef HasExtension As Boolean) As String
    Dim iChar As String
    Dim C  As Long
    Dim iFlag As Long
    Dim iName As String
    Dim iExt As String
    Dim iDotPos As Long
    Dim iNameLen As Long
    Dim iExtLen As Long
    Dim iFileName As String
    
    If ForOldFileFormat_8Dot3 Then
        iFlag = GCT_SHORTCHAR
    Else
        iFlag = GCT_LFNCHAR
    End If
     
     ' Validate the replacement char (thanks Lavolpe)
    If ReplacementChar <> "" Then
        If Not ((PathGetCharType(AscW(ReplacementChar)) And iFlag) = iFlag) Then
            Err.Raise 2069, App.Title & "ValidFileName", "ReplacementChar is not valid."
            Exit Function
        End If
    End If
   
    ProposedFileName = Trim$(ProposedFileName)
    If InStr(ProposedFileName, "/") Then ProposedFileName = Replace(ProposedFileName, "/", "-")  ' to preserve a date formatting
    If InStr(ProposedFileName, """") Then ProposedFileName = Replace(ProposedFileName, """", "'") ' convert double quotes to single quotes to preserve quotation
    
    'strip out not allowed characters in all the file name:
    iFileName = ""
    For C = 1 To Len(ProposedFileName)
        iChar = Mid$(ProposedFileName, C, 1)
        If (PathGetCharType(AscW(iChar)) And iFlag) = iFlag Then
            iFileName = iFileName & iChar
        Else
            iFileName = iFileName & ReplacementChar
        End If
    Next C
    
    ' strip out illegal characters at the end:
    Do
        iChar = Right$(iFileName, 1)
        Select Case iChar
            Case " ", "."
                iFileName = Left(iFileName, Len(iFileName) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    
    ' separate name and (optional) extension
    iDotPos = InStrRev(iFileName, ".")
    If iDotPos > 0 Then
        iName = Left(iFileName, iDotPos - 1)
        iExt = Mid(iFileName, iDotPos + 1)
    Else
        iName = iFileName
        iExt = ""
    End If
    
    ' strip out illegal characters at the beginning of the name
    For C = 1 To Len(iName)
        iChar = Left(iName, 1)
        Select Case iChar
            Case " ", "."
                iName = Mid(iName, 2)
            Case Else
                Exit For
        End Select
    Next C
    
    ' don't permit too long file names (or estensions in the case of ForOldFileFormat_8Dot3)
    iNameLen = Len(iName)
    iExtLen = Len(iExt)
    If ForOldFileFormat_8Dot3 Then
        If iNameLen > 8 Then
            iName = Left(iName, 8)
        End If
        If iExtLen > 3 Then
            iExt = Left(iExt, 3)
        End If
    Else
        If iExtLen > 0 Then
            If iNameLen > 258 Then
                iName = Left(iName, 258)
                iNameLen = 258
            End If
        End If
        If (iNameLen + iExtLen + 1) > 260 Then
            iExt = Left(iExt, 260 - iNameLen - 1)
        End If
    End If
    
    ' don't permit forbidden file names
    Select Case UCase(iName)
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            iName = ""
    End Select
    If iName = "" Then
        'if there is not a valid file name, then use the default
        iName = DefaultFileName
    End If
    
    ' compose file name again
    ValidFileName = iName
    If iDotPos > 0 Then
        If AllowExtension Then
            If ValidFileName <> "" Then
                ValidFileName = ValidFileName & "." & iExt
                HasExtension = True
            End If
        End If
    End If

End Function

'Reads a text file and returns the contents as a string
' - Returns an empty string if the file does not exist or is empty
Function FileRead(FName As String) As String
    On Error GoTo ErroRead
    Dim fso As Object
    Dim streamIn As Object
    
    FileRead = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Se o arquivo não existir, sai da função retornando vazio com segurança
    If Not fso.FileExists(FName) Then
        Set fso = Nothing
        Exit Function
    End If
    
    ' Abre o arquivo no modo de leitura (1 = ForReading)
    Set streamIn = fso.OpenTextFile(FName, 1, False)
    
    ' Lê todo o conteúdo de uma vez só
    FileRead = streamIn.ReadAll
    
    ' Fecha e limpa os objetos da memória
    streamIn.Close
    Set streamIn = Nothing
    Set fso = Nothing
    Exit Function

ErroRead:
    Set streamIn = Nothing
    Set fso = Nothing
End Function
'Writes a text file with the contents of a string
'   - Creates the file if it does not exist
'   - Overwrites the contents of an existing file without warning
'   - Returns true if successful

'Appends the contents to the end of a file
' - if the file does not exist, it is created
' - a new line is implicitly added after the contents
'   `- this means that FileAppend may be repeatedly called without passing any vbCrLf's
'Appends the contents to the end of a file
' - If the file does not exist, it is created automatically
' - A new line is implicitly added after the contents
' - Returns true if successful
Function FileAppend(FName As String, Contents As String) As Boolean
    On Error GoTo ErroAppend
    Dim fso As Object
    Dim streamOut As Object
    
    FileAppend = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre no modo de anexo (8 = ForAppending), e o parâmetro True garante
    ' que o FSO criará o arquivo do zero caso ele ainda não exista!
    Set streamOut = fso.OpenTextFile(FName, 8, True)
    
    ' Grava o conteúdo seguido de uma quebra de linha (conforme a lógica original da sua função)
    streamOut.WriteLine Contents
    
    ' Fecha e encerra o canal de I/O
    streamOut.Close
    Set streamOut = Nothing
    Set fso = Nothing
    
    FileAppend = True
    Exit Function

ErroAppend:
    Set streamOut = Nothing
    Set fso = Nothing
End Function
'Writes a text file with the contents of a string
'   - Creates the file if it does not exist
'   - Overwrites the contents of an existing file safely using FSO
'   - Returns true if successful
Function FileWrite(FName As String, Contents As String) As Boolean
    On Error GoTo ErroWrite
    Dim fso As Object
    Dim streamOut As Object
    
    FileWrite = False
    
    ' Instancia o FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' O parâmetro "True" no final força o FSO a sobrescrever o arquivo
    ' automaticamente se ele já existir, sem necessidade de apagá-lo antes!
    Set streamOut = fso.OpenTextFile(FName, 2, True)
    
    ' Grava o conteúdo de forma limpa
    streamOut.Write Contents
    
    ' Fecha o canal de escrita com segurança
    streamOut.Close
    
    ' Libera os objetos da memória
    Set streamOut = Nothing
    Set fso = Nothing
    
    FileWrite = True
    Exit Function

ErroWrite:
    ' Tratamento preventivo de objetos em caso de erro físico de disco (ex: falta de permissão)
    Set streamOut = Nothing
    Set fso = Nothing
    ' Opcional: Se quiser debugar em tempo de desenvolvimento, descomente a linha abaixo:
    ' Debug.Print "Erro FileWrite: " & Err.Description
End Function
