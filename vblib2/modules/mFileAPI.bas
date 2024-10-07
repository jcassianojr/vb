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


Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER _
       Or OFN_LONGNAMES _
       Or OFN_CREATEPROMPT _
       Or OFN_NODEREFERENCELINKS

Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER _
       Or OFN_LONGNAMES _
       Or OFN_OVERWRITEPROMPT _
       Or OFN_HIDEREADONLY

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40

Public Type OPENFILENAME
  nStructSize As Long
  hwndOwner As Long
  hInstance As Long
  sFILTER As String
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
  hWnd As Long
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
Public Function IsExtensao(ByVal cARQ As String, cEXT As String) As Boolean
 IsExtensao = False
  cARQ = UCase(FixStr(cARQ))
  cEXT = UCase(cEXT)
'  If InStr(1, cARQ, cEXT, vbTextCompare) > 0 Then
'    Extensao = True
'  End If
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
  If FileExists(DestinationFile) Then
      If lAPAGA Then
          DeleteFile DestinationFile  'Kill DestinationFile
      End If
      Exit Function
  End If
  With typFileOperation
    .hWnd = 0
    .wFunc = FO_COPY
    .pFrom = SourceFile & vbNullChar & vbNullChar  'source file
    .pTo = DestinationFile & vbNullChar & vbNullChar  'destination file
    .fFlags = FOF_ALLOWUNDO
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
  Dim tempPath As String, sLen As Integer
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


  'Dim Pos As Long
  'Dim buff As String
  'Dim sLongname As String
  'Dim sShortname As String


  With OFN
    .nStructSize = Len(OFN)
    .hwndOwner = frmOwner.hWnd
    .sFILTER = sFilters & vbNullChar & vbNullChar
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


''Dim buff As String
''Dim sLname As String
''Dim sSname As String

  With OFN

    .nStructSize = Len(OFN)
    .hwndOwner = frmOwner.hWnd
    .sFILTER = sFilters & vbNullChar & vbNullChar
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
  Dim sFILTER As String
  OpenArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  sFILTER = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileOpen(oFORM, sFILTER, 1, sRECENTFILE, cEXT, sPath, "Escolher " & cTITULO)
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
  Dim sFILTER As String
  SaveArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  sFILTER = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileSave(oFORM, sFILTER, 1, cEXT, "Novo", sPath, cTITULO)
  If Len(sFileName) = 0 Then
    lRETU = False
    Exit Function
  End If
  lRETU = True
  SaveArqExt = sFileName
End Function

Public Function ImgFILTER() As String
  Dim sFILTER As String
  sFILTER = "Windows ou OS/2 Bitmap File (*.BMP)" & vbNullChar & "*.BMP" & vbNullChar
  sFILTER = sFILTER & "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  sFILTER = sFILTER & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  sFILTER = sFILTER & "Tagged Imag (*.TIF, *.TIFF)" & vbNullChar & "*.TIF;*.TIFF" & vbNullChar
  sFILTER = sFILTER & "Zsoft Paintbrush PCX bitmap format (*.PCX)" & vbNullChar & "*.PCX" & vbNullChar
  sFILTER = sFILTER & "Windows Icon (*.ICO)" & vbNullChar & "*.ICO" & vbNullChar
  sFILTER = sFILTER & "Graphics Interchange Format (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  sFILTER = sFILTER & "Adobe Photoshop (*.PSD)" & vbNullChar & "*.PSD" & vbNullChar
  sFILTER = sFILTER & "Dr. Halo (*.CUT)" & vbNullChar & "*.CUT" & vbNullChar
  sFILTER = sFILTER & "DirectDraw Surface (*.DDS)" & vbNullChar & "*.DDS" & vbNullChar
  sFILTER = sFILTER & "High Dynamic Range (*.HDR)" & vbNullChar & "*.HDR" & vbNullChar
  sFILTER = sFILTER & "Amiga IFF (*.IFF, *.LBM)" & vbNullChar & "*.IFF;*.LBM" & vbNullChar
  sFILTER = sFILTER & "JPEG Network Graphics (*.JNG)" & vbNullChar & "*.JPG" & vbNullChar
  sFILTER = sFILTER & "Commodore 64 Koala format (*.KOA)" & vbNullChar & "*.KOA" & vbNullChar
  sFILTER = sFILTER & "Multiple Network Graphics (*.MNG)" & vbNullChar & "*.MNG" & vbNullChar
  sFILTER = sFILTER & "Portable Bitmap (ASCII) (*.PBM)" & vbNullChar & "*.PBM" & vbNullChar
  sFILTER = sFILTER & "Portable Bitmap (BINARY) (*.PBM)" & vbNullChar & "*.OBM" & vbNullChar
  sFILTER = sFILTER & "Kodak PhotoCD (*.PCD)" & vbNullChar & "*.PCD" & vbNullChar
  sFILTER = sFILTER & "Portable Graymap (ASCII) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  sFILTER = sFILTER & "Portable Graymap (BINARY) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  sFILTER = sFILTER & "Portable Pixelmap (ASCII) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  sFILTER = sFILTER & "Portable Pixelmap (BINARY) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  sFILTER = sFILTER & "Sun Rasterfile (*.RAS)" & vbNullChar & "*.RAS" & vbNullChar
  sFILTER = sFILTER & "Truevision Targa files (*.TGA, *.TARGA)" & vbNullChar & "*.TGA" & vbNullChar
  sFILTER = sFILTER & "Wireless Bitmap (*.WBMP)" & vbNullChar & "*.WBMP" & vbNullChar
  sFILTER = sFILTER & "X11 Bitmap Format (*.XBM)" & vbNullChar & "*.XBM" & vbNullChar
  sFILTER = sFILTER & "X11 Pixmap Format (*.XPM)" & vbNullChar & "*.XPM" & vbNullChar
  sFILTER = sFILTER & "All Files" & vbNullChar & "*.*"
  ImgFILTER = sFILTER
End Function

Public Function ImgFILTER2() As String
  Dim sFILTER As String
  sFILTER = "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  sFILTER = sFILTER & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  sFILTER = sFILTER & "GIF (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  ImgFILTER2 = sFILTER
End Function

Public Function parsefile(ByVal archivo As String, ByVal parte As String) As String
  Dim PosSlash, PosPunto, PosColon, LastSlash, i As Long
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
  Dim S As String
  S = sDir
  If Right$(S, 1) = "\" Then S = Left$(S, Len(S) - 1)
  On Error GoTo FileExistsError
  ' If no error then something existed.
  FolderExists = ((GetAttr(S) And vbDirectory) = vbDirectory)
  Exit Function
FileExistsError:
  FolderExists = False
  Exit Function
End Function
Public Function OpenStreamFile(FileName$, Mode%, RLock%, RecordLen%) As Integer
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
      Open FileName For Output Shared As FileNum%
    Case RDLOCK
      Open FileName For Output Lock Read As FileNum%
    Case WRLOCK
      Open FileName For Output Lock Write As FileNum%
    Case RWLOCK
      Open FileName For Output Lock Read Write As FileNum%
    End Select
  Case READAFILE
    Select Case RLock%
    Case NOLOCK
      Open FileName For Input Shared As FileNum%
    Case RDLOCK
      Open FileName For Input Lock Read As FileNum%
    Case WRLOCK
      Open FileName For Input Lock Write As FileNum%
    Case RWLOCK
      Open FileName For Input Lock Read Write As FileNum%
    End Select
  Case ADDTOFILE
    Select Case RLock%
    Case NOLOCK
      Open FileName For Append Shared As FileNum%
    Case RDLOCK
      Open FileName For Append Lock Read As FileNum%
    Case WRLOCK
      Open FileName For Append Lock Write As FileNum%
    Case RWLOCK
      Open FileName For Append Lock Read Write As FileNum%
    End Select
  Case RANDOMFILE
    Select Case RLock%
    Case NOLOCK
      Open FileName For Random Shared As FileNum% Len = RecordLen%
    Case RDLOCK
      Open FileName For Random Lock Read As FileNum% Len = RecordLen%
    Case WRLOCK
      Open FileName For Random Lock Write As FileNum% Len = RecordLen%
    Case RWLOCK
      Open FileName For Random Lock Read Write As FileNum% Len = RecordLen%
    End Select
  Case BINARYFILE
    Select Case RLock%
    Case NOLOCK
      Open FileName For Binary Shared As FileNum%
    Case RDLOCK
      Open FileName For Binary Lock Read As FileNum%
    Case WRLOCK
      Open FileName For Binary Lock Write As FileNum%
    Case RWLOCK
      Open FileName For Binary Lock Read Write As FileNum%
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
    If vRecycleBin Then .fFlags = FOF_ALLOWUNDO    ' Send to RecycleBin
  End With

  Call SHFileOperation(fo)
  DeleteFile = Abs(fo.fAborted) - 1
End Function
