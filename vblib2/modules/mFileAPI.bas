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
  sFilter As String
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

#If VBA7 Or Win64 Then
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
Private Declare PtrSafe Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As LongPtr, ByVal lpszShortPath As LongPtr, ByVal cchBuffer As LongPtr) As Long

Private Declare PtrSafe Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Longptr) As Long
Private Declare PtrSafe Function GetDesktopWindow Lib "user32.dll" () As Long

Public Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
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
Private Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32.dll" () As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If

Public Function SH_Copy(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Copy files or folders
    shFileOp.hWnd = GetDesktopWindow
    shFileOp.wFunc = FO_COPY
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Copy = SHFileOperation(shFileOp)
    
End Function
Public Function SH_Move(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Move files or folders
    shFileOp.hWnd = GetDesktopWindow
    shFileOp.wFunc = FO_MOVE
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Move = SHFileOperation(shFileOp)
    
End Function

Public Function SH_Rename(src As String, dest As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Rename files or folders
    shFileOp.hWnd = GetDesktopWindow
    shFileOp.wFunc = FO_RENAME
    shFileOp.pFrom = src
    shFileOp.pTo = dest
    shFileOp.fFlags = cFlags
    
    SH_Rename = SHFileOperation(shFileOp)
    
End Function

Public Function SH_Delete(src As String, cFlags As TFileFlags) As Long
Dim shFileOp As SHFILEOPSTRUCT
    'Delete files or folders
    shFileOp.hWnd = GetDesktopWindow
    shFileOp.wFunc = FO_DELETE
    shFileOp.pFrom = src
    shFileOp.fFlags = cFlags
    
    SH_Delete = SHFileOperation(shFileOp)
    
End Function

Public Function FileExists(ByRef sFileName As String) As Boolean
    Dim fAttr As VbFileAttribute
    Const errBadFileNameOrNumber As Long = 52
    
    If LenB(sFileName) = 0 Then Exit Function
    
    On Error Resume Next
    ' Tenta o método nativo do VB6 primeiro
    fAttr = GetAttr(sFileName)
    
    If Err.Number = 0 Then
        ' Se não deu erro, verifica se é realmente um arquivo (e não um diretório)
        FileExists = Not CBool(fAttr And vbDirectory)
    ElseIf Err.Number = errBadFileNameOrNumber Then
        ' Caso contenha caracteres Unicode ou inválidos para o Dir/GetAttr tradicional,
        ' podemos usar o FSO como alternativa segura no Windows para VB6:
        FileExists = GetFSOFileExists(sFileName)
    Else
        FileExists = False
    End If
    On Error GoTo 0
End Function

' Função auxiliar para usar o FileSystemObject de forma segura no VB6
Private Function GetFSOFileExists(ByRef sFileName As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then
        GetFSOFileExists = fso.FileExists(sFileName)
    End If
    Set fso = Nothing
    On Error GoTo 0
End Function
Public Function CopyFolder(ByRef sourcePath As String, ByRef destinationPath As String, Optional ByVal includeSubFolders As Boolean = True) As Boolean
    Dim fso As Object
    On Error GoTo ErrorHandler
    
    If LenB(sourcePath) = 0 Or LenB(destinationPath) = 0 Then Exit Function
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Garante que a pasta de destino existe (aproveitando sua função CreateNewDirectory ou FSO)
    If Not fso.FolderExists(destinationPath) Then
        CreateNewDirectory destinationPath
    End If
    
    ' Copia a pasta e seu conteúdo (o parâmetro True sobrescreve se necessário)
    fso.CopyFolder sourcePath, destinationPath, True
    
    Set fso = Nothing
    CopyFolder = True
    Exit Function

ErrorHandler:
    Set fso = Nothing
    CopyFolder = False
End Function
Public Function DeleteFolder(ByRef FolderPath As String, Optional ByVal deleteContents As Boolean = True, Optional ByVal failIfMissing As Boolean = False) As Boolean
    Dim fso As Object
    Dim fixedPath As String
    
    DeleteFolder = False
    If LenB(FolderPath) = 0 Then Exit Function
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Padroniza o caminho usando sua função FixPath existente
    fixedPath = FixPath(FolderPath)
    
    ' Remove a barra final exigida pelo FSO em alguns contextos de verificação
    If Right$(fixedPath, 1) = "\" And Len(fixedPath) > 3 Then
        fixedPath = Left$(fixedPath, Len(fixedPath) - 1)
    End If
    
    If Not fso.FolderExists(fixedPath) Then
        DeleteFolder = Not failIfMissing
        Set fso = Nothing
        Exit Function
    End If
    
    ' Se deleteContents for True, força a exclusão mesmo se houver arquivos/subpastas
    ' O parâmetro True no final do DeleteFolder do FSO indica "force" (ignora atributos de somente leitura)
    fso.DeleteFolder fixedPath, True
    
    Set fso = Nothing
    DeleteFolder = True
    Exit Function

ErrorHandler:
    Set fso = Nothing
    DeleteFolder = False
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
    .hWnd = 0
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


Public Function GetTempDirectory() As String
    Dim BufferSize As Long
    Dim tempPath As String
    
    ' 1. Descobre o tamanho exato do buffer necessário chamando a API com tamanho 0
    BufferSize = GetTempPath(0, vbNullString)
    
    If BufferSize > 0 Then
        ' 2. Aloca a string com o tamanho exato retornado pela API
        tempPath = Space$(BufferSize)
        
        ' 3. Preenche o buffer com o caminho real da pasta temporária
        If GetTempPath(BufferSize, tempPath) > 0 Then
            ' Remove o caractere nulo de terminação padrão do Windows
            GetTempDirectory = TrimNull(tempPath)
            
            ' Garante que o caminho retorne padronizado com a barra no final (utilizando sua função FixPath)
            If LenB(GetTempDirectory) > 0 Then
                GetTempDirectory = FixPath(GetTempDirectory)
            End If
            Exit Function
        End If
    End If
    
    ' Fallback de segurança caso a API falhe por algum motivo crítico
    GetTempDirectory = "C:\Temp\"
End Function
Public Sub CreateNewDirectory(ByVal NewDirectory As String)
    Dim fso As Object
    Dim fixedPath As String
    
    NewDirectory = Trim$(NewDirectory)
    If LenB(NewDirectory) = 0 Then Exit Sub
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Padroniza o caminho
    fixedPath = NewDirectory
    If Right$(fixedPath, 1) = "\" And Len(fixedPath) > 3 Then
        fixedPath = Left$(fixedPath, Len(fixedPath) - 1)
    End If
    
    ' Se a pasta ainda não existir, o FSO cria toda a árvore hierárquica de uma só vez
    If Not fso.FolderExists(fixedPath) Then
        fso.CreateFolder fixedPath
    End If
    
    Set fso = Nothing
    Exit Sub

ErrorHandler:
    Set fso = Nothing
    ' Opcional: Tratar erro de permissão caso necessário
End Sub
Public Function FileOpen(frmOwner As Form, _
                         Optional ByVal sFilters As String, _
                         Optional ByVal nFilterIndex As Long = 1, _
                         Optional ByVal sDefaultFileName As String, _
                         Optional ByVal sDefaultExtension As String, _
                         Optional ByVal sStartingFolder As String, _
                         Optional ByVal sTitle As String) As String



  With OFN
    .nStructSize = Len(OFN)
    .hWndOwner = CLng(frmOwner.hWnd)
    .sFilter = sFilters & vbNullChar & vbNullChar
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
    .hWndOwner = CLng(frmOwner.hWnd)
    .sFilter = sFilters & vbNullChar & vbNullChar
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
  Dim sFilter As String
  OpenArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  sFilter = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileOpen(oFORM, sFilter, 1, sRECENTFILE, cEXT, sPath, "Escolher " & cTITULO)
  If Len(sFileName) = 0 Then
    lRETU = False
    Exit Function
  End If
  lRETU = True
  OpenArqExt = sFileName
End Function
Public Function SelecionarPasta(ByVal Titulo As String) As String
    Dim objShell As Object
    Dim objFolder As Object
    
    Set objShell = CreateObject("Shell.Application")
    ' 1 = BIF_RETURNONLYFSDIRS (Apenas pastas do sistema de arquivos)
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 1)
    
    
    If Not objFolder Is Nothing Then
        SelecionarPasta = objFolder.Items.item.Path
    Else
        SelecionarPasta = ""
    End If
    
    Set objFolder = Nothing
    Set objShell = Nothing
End Function
Function SaveArqExt(oFORM As Form, ByVal cARQ As String, ByVal cEXT As String, ByVal cTITULO As String) As String
  Dim sFileName As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFilter As String
  SaveArqExt = ""
  If Len(cARQ) > 0 Then
    sRECENTFILE = cARQ
  Else
    sPath = App.Path
  End If
  sFilter = cTITULO & vbNullChar & "*." & cEXT & vbNullChar
  sFileName = FileSave(oFORM, sFilter, 1, cEXT, "Novo", sPath, cTITULO)
  If Len(sFileName) = 0 Then
    lRETU = False
    Exit Function
  End If
  lRETU = True
  SaveArqExt = sFileName
End Function

Public Function ImgFILTER() As String
  Dim sFilter As String
  sFilter = "Windows ou OS/2 Bitmap File (*.BMP)" & vbNullChar & "*.BMP" & vbNullChar
  sFilter = sFilter & "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  sFilter = sFilter & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  sFilter = sFilter & "Tagged Imag (*.TIF, *.TIFF)" & vbNullChar & "*.TIF;*.TIFF" & vbNullChar
  sFilter = sFilter & "Zsoft Paintbrush PCX bitmap format (*.PCX)" & vbNullChar & "*.PCX" & vbNullChar
  sFilter = sFilter & "Windows Icon (*.ICO)" & vbNullChar & "*.ICO" & vbNullChar
  sFilter = sFilter & "Graphics Interchange Format (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  sFilter = sFilter & "Adobe Photoshop (*.PSD)" & vbNullChar & "*.PSD" & vbNullChar
  sFilter = sFilter & "Dr. Halo (*.CUT)" & vbNullChar & "*.CUT" & vbNullChar
  sFilter = sFilter & "DirectDraw Surface (*.DDS)" & vbNullChar & "*.DDS" & vbNullChar
  sFilter = sFilter & "High Dynamic Range (*.HDR)" & vbNullChar & "*.HDR" & vbNullChar
  sFilter = sFilter & "Amiga IFF (*.IFF, *.LBM)" & vbNullChar & "*.IFF;*.LBM" & vbNullChar
  sFilter = sFilter & "JPEG Network Graphics (*.JNG)" & vbNullChar & "*.JPG" & vbNullChar
  sFilter = sFilter & "Commodore 64 Koala format (*.KOA)" & vbNullChar & "*.KOA" & vbNullChar
  sFilter = sFilter & "Multiple Network Graphics (*.MNG)" & vbNullChar & "*.MNG" & vbNullChar
  sFilter = sFilter & "Portable Bitmap (ASCII) (*.PBM)" & vbNullChar & "*.PBM" & vbNullChar
  sFilter = sFilter & "Portable Bitmap (BINARY) (*.PBM)" & vbNullChar & "*.OBM" & vbNullChar
  sFilter = sFilter & "Kodak PhotoCD (*.PCD)" & vbNullChar & "*.PCD" & vbNullChar
  sFilter = sFilter & "Portable Graymap (ASCII) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  sFilter = sFilter & "Portable Graymap (BINARY) (*.PGM)" & vbNullChar & "*.PGM" & vbNullChar
  sFilter = sFilter & "Portable Pixelmap (ASCII) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  sFilter = sFilter & "Portable Pixelmap (BINARY) (*.PPM)" & vbNullChar & "*.PPM" & vbNullChar
  sFilter = sFilter & "Sun Rasterfile (*.RAS)" & vbNullChar & "*.RAS" & vbNullChar
  sFilter = sFilter & "Truevision Targa files (*.TGA, *.TARGA)" & vbNullChar & "*.TGA" & vbNullChar
  sFilter = sFilter & "Wireless Bitmap (*.WBMP)" & vbNullChar & "*.WBMP" & vbNullChar
  sFilter = sFilter & "X11 Bitmap Format (*.XBM)" & vbNullChar & "*.XBM" & vbNullChar
  sFilter = sFilter & "X11 Pixmap Format (*.XPM)" & vbNullChar & "*.XPM" & vbNullChar
  sFilter = sFilter & "All Files" & vbNullChar & "*.*"
  ImgFILTER = sFilter
End Function

Public Function ImgFILTER2() As String
  Dim sFilter As String
  sFilter = "Independent JPEG Group (*.JPG, *.JIF, *.JPEG, *.JPE)" & vbNullChar & "*.JPG;*.JIF;*.JPEG;*.JPE" & vbNullChar
  sFilter = sFilter & "Portable Network Graphics (*.PNG)" & vbNullChar & "*.PNG" & vbNullChar
  sFilter = sFilter & "GIF (*.GIF)" & vbNullChar & "*.GIF" & vbNullChar
  ImgFILTER2 = sFilter
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

Public Function FixPath(ByVal pathToFix As String) As String
    Const PS As String = "\"
    Dim resultPath As String
    Dim isUNC As Boolean
    
    If LenB(pathToFix) = 0 Then Exit Function
    
    ' Padroniza todas as barras para o formato do Windows (\) e identifica se é UNC (ex: \\servidor\pasta)
    resultPath = Replace(pathToFix, "/", PS)
    isUNC = (Left$(resultPath, 2) = "\\")
    
    ' Remove duplicidades excessivas de separadores, preservando o prefixo UNC se houver
    resultPath = RemoveDuplicatePS(resultPath, isUNC)
    
    ' Garante que o caminho termine com a barra separadora padrão
    If Right$(resultPath, 1) <> PS Then
        resultPath = resultPath & PS
    End If
    
    FixPath = resultPath
End Function
Public Function AddDirSep(strPathName As String) As String
    AddDirSep = Trim$(strPathName)
    If Right$(AddDirSep, 1) <> "\" Then AddDirSep = AddDirSep & "\"
End Function
' Função auxiliar inspirada diretamente na lógica de remoção de duplicidade da LibFileTools
Private Function RemoveDuplicatePS(ByVal pathToFix As String, ByVal isUNC As Boolean) As String
    Const PS As String = "\"
    Dim startPos As Long
    Dim currPos As Long
    Dim prevPos As Long
    Dim diff As Long
    Dim i As Long
    
    If isUNC Then currPos = 2 ' Pula o prefixo UNC inicial: \\
    RemoveDuplicatePS = pathToFix
    
    Do
        prevPos = currPos
        currPos = InStr(currPos + 1, pathToFix, PS)
        If startPos = 0 Then startPos = prevPos + 1
        If currPos - prevPos <= 1 Then
            diff = currPos - startPos
            If currPos = 0 Then diff = diff + Len(pathToFix) + 1
            If startPos * Sgn(i * diff) > 1 Then
                Mid$(RemoveDuplicatePS, i) = Mid$(pathToFix, startPos, diff)
                i = i + diff
            End If
            If i = 0 Then i = (startPos + diff) * Sgn(prevPos)
            startPos = 0
        End If
    Loop Until currPos = 0
    
    If i > 1 Then RemoveDuplicatePS = Left$(RemoveDuplicatePS, i - 1)
End Function
Public Function GetDriveUNC(ByVal DriveString As String) As String
    Dim fso As Object
    Dim driveName As String
    Dim fsDrive As Object
    
    DriveString = Trim$(DriveString)
    If LenB(DriveString) = 0 Then Exit Function
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Extrai a raiz/drive (ex: "C:" ou "\\servidor\compartilhamento")
    driveName = fso.GetDriveName(DriveString)
    
    If LenB(driveName) = 0 Then
        ' Se não conseguir pelo FSO, tenta padronizar via string limpa
        GetDriveUNC = ""
        Set fso = Nothing
        Exit Function
    End If
    
    ' Se começar com \\, já é um caminho UNC válido
    If Left$(driveName, 2) = "\\" Then
        GetDriveUNC = FixPath(driveName)
        Set fso = Nothing
        Exit Function
    End If
    
    ' Verifica se é uma unidade mapeada (ex: Z:) para buscar o ShareName correspondente
    Set fsDrive = fso.GetDrive(driveName)
    If Not fsDrive Is Nothing Then
        If LenB(fsDrive.shareName) > 0 Then
            GetDriveUNC = FixPath(fsDrive.shareName)
        Else
            ' Se for uma unidade local pura (ex: C:\), retorna a letra com barra
            GetDriveUNC = FixPath(fsDrive.driveLetter & ":")
        End If
    End If
    
    Set fsDrive = Nothing
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Set fsDrive = Nothing
    Set fso = Nothing
    GetDriveUNC = ""
End Function

Public Function FolderExists(ByRef sDir As String) As Boolean
    Dim fso As Object
    Dim fixedPath As String
    
    FolderExists = False
    If LenB(sDir) = 0 Then Exit Function
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Padroniza o caminho removendo barra final excessiva para o FSO validar corretamente
    fixedPath = Trim$(sDir)
    If Right$(fixedPath, 1) = "\" And Len(fixedPath) > 3 Then
        fixedPath = Left$(fixedPath, Len(fixedPath) - 1)
    End If
    
    FolderExists = fso.FolderExists(fixedPath)
    
    Set fso = Nothing
    Exit Function

ErrorHandler:
    Set fso = Nothing
    FolderExists = False
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
    If vRecycleBin Then .fFlags = FO_ALLOWUNDO    ' Send to RecycleBin
  End With

  Call SHFileOperation(fo)
  DeleteFile = Abs(fo.fAborted) - 1
End Function
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
    Dim nameLen As Long
    Dim i As Long
    
    If ForOldFileFormat_8Dot3 Then
        iFlag = GCT_SHORTCHAR
    Else
        iFlag = GCT_LFNCHAR
    End If
     
    ' Valida o caractere de substituição
    If ReplacementChar <> "" Then
        If Not ((PathGetCharType(AscW(ReplacementChar)) And iFlag) = iFlag) Then
            Err.Raise 2069, App.Title & "ValidFileName", "ReplacementChar is not valid."
            Exit Function
        End If
    End If
   
    ProposedFileName = Trim$(ProposedFileName)
    If InStr(ProposedFileName, "/") Then ProposedFileName = Replace(ProposedFileName, "/", "-")
    If InStr(ProposedFileName, """") Then ProposedFileName = Replace(ProposedFileName, """", "'")
    
    ' Remove caracteres não permitidos em todo o nome do arquivo[cite: 1]
    iFileName = ""
    For C = 1 To Len(ProposedFileName)
        iChar = Mid$(ProposedFileName, C, 1)
        If (PathGetCharType(AscW(iChar)) And iFlag) = iFlag Then
            iFileName = iFileName & iChar
        Else
            iFileName = iFileName & ReplacementChar
        End If
    Next C
    
    ' Remove caracteres ilegais ou espaços/pontos nas extremidades (melhoria inspirada na LibFileTools)[cite: 1, 2]
    nameLen = Len(iFileName)
    i = nameLen
    If i > 0 Then
        Do While i > 0
            Select Case Mid$(iFileName, i, 1)
                Case " ", "."
                    i = i - 1
                Case Else
                    Exit Do
            End Select
        Loop
        If i < nameLen Then
            iFileName = Left$(iFileName, i)
        End If
    End If
    
    ' Separa nome e extensão (opcional)[cite: 1]
    iDotPos = InStrRev(iFileName, ".")
    If iDotPos > 0 Then
        iName = Left(iFileName, iDotPos - 1)
        iExt = Mid(iFileName, iDotPos + 1)
    Else
        iName = iFileName
        iExt = ""
    End If
    
    ' Remove caracteres ilegais no início do nome[cite: 1]
    For C = 1 To Len(iName)
        iChar = Left(iName, 1)
        Select Case iChar
            Case " ", "."
                iName = Mid(iName, 2)
            Case Else
                Exit For
        End Select
    Next C
    
    ' Impede nomes ou extensões longas demais[cite: 1]
    iNameLen = Len(iName)
    iExtLen = Len(iExt)
    If ForOldFileFormat_8Dot3 Then
        If iNameLen > 8 Then iName = Left(iName, 8)
        If iExtLen > 3 Then iExt = Left(iExt, 3)
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
    
    ' Protege contra nomes reservados do Windows[cite: 1]
    Select Case UCase(iName)
        Case "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
            iName = ""
    End Select
    If iName = "" Then
        iName = DefaultFileName
    End If
    
    ' Reconstrói o nome final do arquivo[cite: 1]
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
Public Function GetFullPath(ByVal cPath As String) As String
    ' Garante que caminhos relativos não falhem em 64-bit
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFullPath = fso.GetAbsolutePathName(cPath)
    Set fso = Nothing
End Function
