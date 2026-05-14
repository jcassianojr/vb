Attribute VB_Name = "dcompareMain"
Sub Main()
'InitXPStyles

  '''DAO.DBEngine.SetOption dbExclusiveAsyncDelay, 2000
  '''DAO.DBEngine.SetOption dbSharedAsyncDelay, 0
  '''DAO.DBEngine.SetOption dbFlushTransactionTimeout, 500
  '
  '''DAO.DBEngine.SetOption dbUserCommitSync, "yes"
  '''DAO.DBEngine.SetOption dbImplicitCommitSync, "yes"
  ''DAO.DBEngine.SetOption dbLockRetry, 20
  'DAO.DBEngine.SetOption dbPageTimeout, 5000
  'DAO.DBEngine.SetOption dbMaxLocksPerFile, 9500
  '''DAO.DBEngine.SetOption dbLockDelay, 100
  '''DAO.DBEngine.SetOption dbRecycleLVs, 0
  
  ' Aumentar o limite de travas para evitar erro "MaxLocksPerFile exceeded" em grandes correções
    '''DAO.DBEngine.SetOption dbMaxLocksPerFile, 15000
    
    ' Desativar sincronização implícita pode acelerar gravações em massa (use com cautela)
  '  DAO.DBEngine.SetOption dbImplicitCommitSync, "no"
    
    ' Aumentar o timeout para redes ou discos lentos
    '''DAO.DBEngine.SetOption dbPageTimeout, 6000

' Reative apenas o essencial. Se fechar, comente de novo.
   ' On Error Resume Next
   ' DAO.DBEngine.setOption dbMaxLocksPerFile, 15000
   ' On Error GoTo 0

 ' EnabledVisualStyles
  dCompare.Show
End Sub
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
  'Case "E"                                     ' Extensao
    '//parsefile = Mid(archivo, PosPunto + 1)
  Case "E"
    PosPunto = InStrRev(archivo, ".")
    If PosPunto > 0 Then parsefile = Mid(archivo, PosPunto + 1)
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

Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
  MDG = False
  eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
  If eRESP = vbYes Then
    MDG = True
  End If
End Function


Public Function GeraConexao(ByVal Caminho As String)
    Dim sProvider As String
    If LCase(Right(Caminho, 5)) = "accdb" Then
        sProvider = "Microsoft.ACE.OLEDB.12.0"
    Else
        sProvider = "Microsoft.Jet.OLEDB.4.0"
    End If
    GeraConexao = "Provider=" & sProvider & ";Data Source=" & Caminho
End Function

Public Function Alert(ByVal cDIZ As String, Optional ByVal cTITLE As String = "Informacao")
  MsgBox cDIZ, vbOKOnly, cTITLE
End Function

Public Function ArquivoExiste(ByVal FilePath As String) As Boolean
    ArquivoExiste = (Dir(FilePath, vbNormal + vbHidden + vbSystem) <> "")
End Function

Public Function TipoDado2(ByVal intType As Integer) As String
  Select Case intType
  Case adSmallInt, adInteger, adSingle, adDouble, adCurrency, adBigInt, adBinary, _
       adDecimal, adNumeric, adTinyInt, adUnsignedBigInt, adUnsignedInt, _
       adUnsignedSmallInt, adUnsignedTinyInt
    TipoDado2 = "N"                          ''E um Numero
  Case adDate, adDBDate, adDBTimeStamp
    TipoDado2 = "D"                          ''Tipos Data
  Case adBoolean
    TipoDado2 = "B"                          ''Tipos Boolean
  Case adVarChar, adLongVarChar, adChar, adLongVarWChar, adVarWChar, adWChar
    TipoDado2 = "S"                          ''Tipos String
  Case adDBTime
    TipoDado2 = "T"                          ''Tipos Time
  Case adBSTR, adChapter, adFileTime, adGUID, adIDispatch, adUserDefined
    TipoDado2 = "X"                          ''Outros Tipos
  Case adEmpty
    TipoDado2 = "U"                          ''uNDEFINIDOS VAZIOS
  Case adError, adIUnknown
    TipoDado2 = "E"                          ''Erros
  Case adLongVarBinary
    TipoDado2 = "O"                          ''Ole
  Case adPropVariant, adVarBinary, adVariant   ''Variantes
    TipoDado2 = "V"
  Case Else
    TipoDado2 = "U"
  End Select
End Function


Public Function TipoDados(ByVal intType As Integer) As String
''EMPTY 0
''DECIMAL 14
''DBTIMESTAMP 135
''TINYINT 16
''NUMERIC 131
''BSTR 8
''SMALLINT 2 BOOLEAN 11 CHAR 129
''INTEGER 3 ERROR 10 VARCHAR 200
''BIGINT 20 USERDEFINED 132 LONGVARCHAR 201
''UNSIGNEDTINYINT 17 VARIANT 12 WCHAR 130
''UNSIGNEDSMALLINT 18 IDISPATCH 9 VARWCHAR 202
''UNSIGNEDINT 19 IUNKNOWN 13 LONGVARWCHAR 203
''UNSIGNEDBIGINT 21 GUID 72 BINARY 128
''SINGLE 4 DATE 7 VARBINARY 204
''DOUBLE 5 DBDATE 133 LONGVARBINARY 205
''CURRENCY 6 DBTIME 134 CHAPTER 136

  Select Case intType
  Case adSmallInt
    TipoDados = "SmallInt"
  Case adInteger
    TipoDados = "Integer"
  Case adSingle
    TipoDados = "Single"
  Case adDouble
    TipoDados = "Double"
  Case adDate
    TipoDados = "Date"
  Case adBoolean
    TipoDados = "Boolean"
  Case adVarChar
    TipoDados = "VarChar"
  Case adLongVarChar
    TipoDados = "LongVarChar"
  Case adBigInt
    TipoDados = "BigInt"
  Case adChar
    TipoDados = "Char"
  Case adDBDate
    TipoDados = "DBDate"
  Case adDBTime
    TipoDados = "DBTime"
  Case adDBTimeStamp
    TipoDados = "DBTimeStamp"
  Case adDecimal
    TipoDados = "Decimal"
  Case adFileTime
    TipoDados = "FileTime"
  Case adNumeric
    TipoDados = "Numeric"
  Case adPropVariant
    TipoDados = "PropVariant"
  Case adTinyInt
    TipoDados = "TinyInt"
  Case adUnsignedBigInt
    TipoDados = "UnsignedBigInt"
  Case adUnsignedInt
    TipoDados = "UnsignedInt"
  Case adUnsignedSmallInt
    TipoDados = "UnsignedSmallInt"
  Case adUnsignedTinyInt
    TipoDados = "UnsignedTinyInt"
  Case adLongVarBinary
    TipoDados = "LongVarBinary(Ole)"
  Case adLongVarWChar
    TipoDados = "LongVarWChar(Memo)"
  Case adVarNumeric
    TipoDados = "VarNumeric"
  Case adVarWChar
    TipoDados = "VarWChar"
  Case adWChar
    TipoDados = "WChar"
  Case adCurrency
    TipoDados = "Currency"
  Case adBinary
    TipoDados = "Binary"
  Case adBSTR
    TipoDados = "BSTR"
  Case adChapter
    TipoDados = "Chapter"
  Case adBinary
    TipoDados = "Binary"
  Case adBSTR
    TipoDados = "BSTR"
  Case adChapter
    TipoDados = "Chapter"
  Case adGUID
    TipoDados = "GUID"
  Case adIDispatch
    TipoDados = "IDispatch"
  Case adIUnknown
    TipoDados = "IUnknown"
  Case adGUID
    TipoDados = "GUID"
  Case adIDispatch
    TipoDados = "IDispatch"
  Case adIUnknown
    TipoDados = "IUnknown"

  Case Else
    TipoDados = "Tipo de dados não reconhecido"
  End Select
End Function


