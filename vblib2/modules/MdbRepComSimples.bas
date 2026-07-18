Attribute VB_Name = "MdbRepComSimples"
Option Explicit

Public Enum enumDatabaseType
  [MSAccess2000] ' .mdb
  [MSAccess2007] ' .accdb
End Enum
''' <summary>
''' Compacta e repara uma base de dados Access (.mdb ou .accdb) usando a tecnologia adequada
''' </summary>
Public Function CompactAccess(xDatabaseLocation As String, _
                              Optional xDatabasePassword As String) As Boolean

    Dim jro As Object
    Dim daoEngine As Object
    Dim strTempFile As String
    Dim strPass As String
    Dim isAccdb As Boolean

    On Error GoTo CompactAccess_Err
    CompactAccess = False

    ' 1. Verifica a extensão para definir o tipo
    isAccdb = (InStr(LCase(xDatabaseLocation), ".accdb") > 0)
    strTempFile = xDatabaseLocation & ".temp"
    
    If Dir(strTempFile) <> "" Then Kill strTempFile

    ' 2. Executa a compactação condicional
    If Not isAccdb Then
        ' Usa JRO para .mdb
        strPass = IIf(Len(xDatabasePassword) > 0, ";Jet OLEDB:Database Password=" & xDatabasePassword, "")
        Set jro = CreateObject("JRO.JetEngine")
        jro.CompactDataBase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xDatabaseLocation & strPass, _
                           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTempFile & strPass
    Else
        ' Usa DAO para .accdb
        strPass = IIf(Len(xDatabasePassword) > 0, ";PWD=" & xDatabasePassword, "")
        Set daoEngine = CreateObject("DAO.DBEngine.120")
        daoEngine.CompactDataBase xDatabaseLocation, strTempFile, , , strPass
    End If

    ' 3. Substitui o original pelo compactado
    Kill xDatabaseLocation
    Name strTempFile As xDatabaseLocation

    CompactAccess = True
    
CleanExit:
    Set jro = Nothing
    Set daoEngine = Nothing
    Exit Function

CompactAccess_Err:
    MsgBox "Erro ao compactar base de dados:" & vbCrLf & Err.Description, vbCritical, "Compactação"
    If Dir(strTempFile) <> "" Then Kill strTempFile
    Resume CleanExit
End Function



