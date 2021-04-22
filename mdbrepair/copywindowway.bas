Attribute VB_Name = "copywindowway"
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Function CopyFileWindowsWay(ByVal SourceFile As String, ByVal DestinationFile As String, Optional ByVal lAPAGA As Boolean = False) As Long
    Dim lngReturn As Long
    Dim typFileOperation As SHFILEOPSTRUCT
    If lAPAGA Then
        If FileExists(DestinationFile) Then
            Kill DestinationFile
        Else
            MsgBox "ja existe arquivo destino" & DestinationFile
        End If
    End If
    With typFileOperation
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
        .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
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
