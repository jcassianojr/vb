Attribute VB_Name = "AdoLib"
'MaxBufferSize=2048;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes
'
'mysql
'Your table must have a primary key
'Your connection string must have Option=3
'Driver={MySQL ODBC 5.2 ANSI Driver};Server=myServerAddress;Database=myDataBase;UID=myUsername;PWD=myPassword;Option=3;
'
'const JET_ENGINETYPE abaixo
'DoConvertMDB "C:\data\MyDB97.MDB", "C:\data\MyDB2000.MDB", Jet4x
'Sub DoConvertMDB(SourceDB, DestDB, Format)
'Dim EngineS
'    Set Engine = CreateObject("JRO.JetEngine")
'    Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SourceDB, _
     '      "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=" & Format & ";Data Source=" & DestDB
'End Sub
'myData.Properties("Update Criteria").value = adCriteriaKey

Option Explicit
Public Const cJetPro = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const cJetA12 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Public Const cJetA16 = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source="
Public Const cJetADV = "Provider=Advantage.OLEDB.1;Data Source="
Public Const cJetExt = ";Extended Properties="
Public Const JET_ENGINETYPE_UNKNOWN = 0
Public Const JET_ENGINETYPE_JET10 = 1
Public Const JET_ENGINETYPE_JET11 = 2
Public Const JET_ENGINETYPE_JET2X = 3
Public Const JET_ENGINETYPE_JET3X = 4
Public Const JET_ENGINETYPE_JET4X = 5
Public Const JET_ENGINETYPE_DBASE3 = 10
Public Const JET_ENGINETYPE_DBASE4 = 11
Public Const JET_ENGINETYPE_DBASE5 = 12
Public Const JET_ENGINETYPE_EXCEL30 = 20
Public Const JET_ENGINETYPE_EXCEL40 = 21
Public Const JET_ENGINETYPE_EXCEL50 = 22
Public Const JET_ENGINETYPE_EXCEL80 = 23
Public Const JET_ENGINETYPE_EXCEL90 = 24
Public Const JET_ENGINETYPE_EXCHANGE4 = 30
Public Const JET_ENGINETYPE_LOTUSWK1 = 40
Public Const JET_ENGINETYPE_LOTUSWK3 = 41
Public Const JET_ENGINETYPE_LOTUSWK4 = 42
Public Const JET_ENGINETYPE_PARADOX3X = 50
Public Const JET_ENGINETYPE_PARADOX4X = 51
Public Const JET_ENGINETYPE_PARADOX5X = 52
Public Const JET_ENGINETYPE_PARADOX7X = 53
Public Const JET_ENGINETYPE_TEXT1X = 60
Public Const JET_ENGINETYPE_HTML1X = 70

'' rs.Open "SELECT * FROM [Salary$A1:B2] ", cn, adOpenDynamic, adLockOptimistic excel select
''                         pasta$faixa
''Geracarq Gera Um Conecao conforme tipo dbf->foxpro mdb->oledbacess ...
'' convertendo [JETFOX],[JETMDB]...
''Geraconn gera um conecao conforme tipo dbf->[FOXPRO] mdb->[JETMDB]oledbacess ...
''Tipo con gera a matrix com o tipo conecao,string concecao,tipo interno conecao
''Tipodado2 retorna o tipo C aracter D ata ...
''Tipodado retorno o tipo conforme o padrao ado
Public Function GeracArq(ByVal cARQ As String, Optional ByVal cTIPO As String = "", Optional ByVal lWRITE As Boolean = True)
  Dim aRETU As Variant
  cARQ = GeraConn(cARQ, cTIPO)
  aRETU = TipoConn(cARQ, , , lWRITE)
  GeracArq = aRETU(1)
End Function

Public Function GeraConn(ByVal cARQ As String, Optional cTIPO As String = "") As String
  Dim nPOS As Long
  Dim cARQTMP As String
  
  GeraConn = cARQ
  cARQTMP = UCase(cARQ) 'usado no instr maiscula mas sempre atribui na carq conecao e case sensitive
  If InStr(cARQTMP, "[") > 0 Then                 ''ja e uma connecao
    Exit Function
  End If
  If cTIPO = "JETMDB" Or cTIPO = "MDB" Or InStr(cARQTMP, ".MDB") > 0 Then
    GeraConn = "[JETMDB]" & cARQ
    Exit Function
  End If
  If cTIPO = "JETFOX" Or cTIPO = "FOX" Or cTIPO = "DBF" Or cTIPO = "SDECDX" Or cTIPO = "SDEFOX" Then
    GeraConn = "[JETFOX]" & cARQ
    Exit Function
  End If
    If cTIPO = "SQLITE" Or InStr(LCase(cARQ), ".sqlite") > 0 Or InStr(LCase(cARQ), ".sqlite3") > 0 _
                 Or InStr(LCase(cARQ), ".fossil") > 0 Or InStr(LCase(cARQ), ".db3") > 0 _
                 Or InStr(LCase(cARQ), ".db") > 0 Then
    GeraConn = "[SQLITE]" & cARQ
    Exit Function
  End If

  If InStr(cARQTMP, ".DBF") > 0 Then  'DBF
    nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
    cARQ = Mid(cARQ, 1, nPOS)
    GeraConn = "[JETFOX]" & cARQ
    Exit Function
  End If

  If cTIPO = "A16MDB" Or cTIPO = "A16JETMDB" Then
    GeraConn = "[A16MDB]" & cARQ
    Exit Function
  End If
  
  If InStr(cARQTMP, ".PD") > 0 Then   ' paradox
    nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
    cARQ = Mid(cARQ, 1, nPOS)
    GeraConn = "[JETPDX5]" & cARQ
    Exit Function
  End If
  If cTIPO = "A12MDB" Or cTIPO = "A12JETMDB" Then
    GeraConn = "[A12MDB]" & cARQ
    Exit Function
  End If
  
Select Case cTIPO
    Case "XLS"
      GeraConn = "[XLS]" & cARQ
    Case "XLSX"
      GeraConn = "[XLSX]" & cARQ
    Case "XLSDRV"
      GeraConn = "[XLSDRV]" & cARQ
    Case "SQLITE"
      GeraConn = "[SQLITE]" & cARQ
    Case "JETTXT"
      GeraConn = "[JETTXT]" & cARQ
    Case "DBFIII"
      GeraConn = "[JETDBFIII]" & cARQ
    Case "ADSDX"
      GeraConn = "[ADSCDX]" & cARQ
    Case "ADSNTX" Or "SDENTX"
      GeraConn = "[ADSNTX]" & cARQ
    Case "ADSADT"
      GeraConn = "[ADSADT]" & cARQ
    Case "ADSVFP"
      GeraConn = "[ADSVFP]" & cARQ
    Case "PDX3"
      GeraConn = "[JETPDX3]" & cARQ
    Case "PDX4"
      GeraConn = "[JETPDX4]" & cARQ
    Case "PDX5"
      GeraConn = "[JETPDX5]" & cARQ
    Case "A12DBFIII"
      GeraConn = "[A12DBFIII]" & cARQ
    Case "A12PDX3"
      GeraConn = "[A12PDX3]" & cARQ
    Case "A12PDX4"
      GeraConn = "[A12PDX4]" & cARQ
    Case "A12PDX5"
      GeraConn = "[A12PDX5]" & cARQ
    Case "A16XLS"
      GeraConn = "[A16XLS]" & cARQ
    Case "A16XLSX"
      GeraConn = "[A16XLSX]" & cARQ
    Case "A16XLSM"
      GeraConn = "[A16XLSM]" & cARQ
    Case "A16XLSB"
      GeraConn = "[A16XLSB]" & cARQ
    Case "A16DBFIII"
      GeraConn = "[A16DBFIII]" & cARQ
    Case "A16PDX3"
      GeraConn = "[A16PDX3]" & cARQ
    Case "A16PDX4"
      GeraConn = "[A16PDX4]" & cARQ
    Case "A16PDX5"
      GeraConn = "[A16PDX5]" & cARQ
End Select
End Function

Public Function TipoConn(ByVal cARQ As String, Optional ByVal cUSER As String = "", _
                         Optional ByVal cPASS As String = "", Optional ByVal lWRITE As Boolean = True) As Variant
  Dim cARQTMP As String
  Dim cJETUSO As String
  Dim lTEMMDB As Boolean
  Dim lTEMSQLITE As Boolean
  Dim cADSTIP As String
  Dim cADSNOM As String
  Dim cXLSVER As String '
  lTEMMDB = False
  lTEMSQLITE = False
  TipoConn = Array("ADO", cARQ, "???")
  cARQTMP = UCase(cARQ)
  
  '
  'access
  '
  If InStr(cARQTMP, ".MDB") > 0 Then
    lTEMMDB = True
    TipoConn = Array("ADO", cARQ, "MDB")
  End If
  '
  'sqlite
  '
  If InStr(cARQTMP, ".SQLITE") > 0 Then
    lTEMSQLITE = True
    TipoConn = Array("ADO", cARQ, "SQLITE")
  End If
  '
  'checando provider
  '
  If InStr(cARQ, "[") = 0 Then
    If InStr(cARQTMP, "PROVIDER") > 0 Then
      Exit Function
    End If
    If lTEMMDB Then
      cARQ = "[JETMDB]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
     If lTEMSQLITE Then
      cARQ = "[SQLITE]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
  End If
'
'access
'
  If InStr(cARQTMP, "[JETMDB]") > 0 Then
    cARQ = Replace(cARQ, "[JETMDB]", "")
    If InStr("PROVIDER", cARQTMP) = 0 Then
       cARQ = cJetPro & cARQ
    End If
    If Len(cUSER) > 0 Then
      cARQ = cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
    End If
    If Not lWRITE Then
      cARQ = cARQ & ";Mode=Share Deny Write"  '";Mode=Read"
    Else
      cARQ = cARQ & ";Mode=Share Deny None"  '";Mode=Read"
    End If
    TipoConn = Array("ADO", cARQ, "MDB")
    Exit Function                            '' encerra agilizar
  End If
  
  '
  ' SQLITE
  '
  If InStr(cARQTMP, "[SQLITE]") > 0 Then  'c:\Program Files (x86)\SQLite ODBC Driver\readme.txt http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe
    cARQ = Replace(cARQ, "[SQLITE]", "")
    If InStr(cARQTMP, "SQLITE3 ODBC DRIVER") = 0 Then
       cARQ = "Driver={SQLite3 ODBC Driver};Database=" + cARQ + ";"
    End If
    TipoConn = Array("ADO", cARQ, "SQLITE")
    Exit Function
  End If

  '
  'jetfox
  '
  If InStr(cARQTMP, "[JETFOX]") > 0 Or InStr(cARQTMP, "[SDECDX]") > 0 Then
    cARQ = Replace(cARQ, "[JETFOX]", "")
    cARQ = Replace(cARQ, "[SDECDX]", "")

    ' PROVIDER=VFPOLEDB.1;Data Source=caminho
    ' ;SourceType=dbf;Deleted=Yes;Mode=ReadWrite|Share Deny None;Mode=Share Deny None
    ' ;DELETED=True;CODEPAGE=1252;MVCOUNT=16384;ENGINEBEHAVIOR=90
    ' ;TABLEVALIDATE=0;REFRESH=5;VARCHARMAPPING=False
    ' ;ANSI=True;REPROCESS=5;OLE DB Services = 0;
    '   MVCOUNT - sets the maximum number of variables that Visual FoxPro can maintain.
    '   ENGINEBEHAVIOR - specifies the SQL Engine compatibility level.
    '   TABLEVALIDATE - specifies the level of table validation to perform.
    '   REFRESH - specifies how often local memory buffers are refreshed with changes from other users on the network.
    '   DELETED=true/false
    '   ANSI=true/false
    '   REPROCESS=nnn
    '   CODEPAGE=nnnn
    '   VARCHARMAPPING=true/false
    '   Collating Sequence=cSequenceName
    If lWRITE Then  'deleted=true tambem no where pode incluir  and deleted()
      cARQ = "Provider=VFPOLEDB.1;Data Source=" & cARQ & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False;Collating Sequence=MACHINE;DELETED=True;"  'NULL=NO"
    Else
      cARQ = "Provider=VFPOLEDB.1;Data Source=" & cARQ & ";Mode=Read|Share Deny None;Persist Security Info=False;Collating Sequence=MACHINE;DELETED=True;"  'NULL=NO"
    End If
    TipoConn = Array("ADO", cARQ, "DBF")
    Exit Function
  End If
  If InStr(cARQTMP, "[XLSDRV]") > 0 Then
    cARQ = Replace(cARQ, "[XLSDRV]", "")
    cARQ = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & cARQ
    TipoConn = Array("ADO", cARQ, "XLSDRV")
    Exit Function
  End If
  If InStr(cARQTMP, "[CONN]") > 0 Then
    cARQ = Replace(cARQ, "[CONN]", "")
    If lTEMMDB > 0 Then
      TipoConn = Array("ADO", cARQ, "MDB")
    Else
      TipoConn = Array("ADO", cARQ, "CONN")
    End If
    Exit Function
  End If
  If InStr(cARQTMP, "[SQLSERVER]") > 0 Then
    cARQ = Replace(cARQ, "[SQLSERVER]", "")
    TipoConn = Array("ADO", cARQ, "SQLSERVER")
    Exit Function
  End If

  If InStr(cARQTMP, "[A16MDB]") > 0 Then
    cARQ = Replace(cARQ, "[A16MDB]", "")
    If Len(cUSER) > 0 Then
      cARQ = cJetA16 & cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
    Else
      cARQ = cJetA16 & cARQ                ''& ";Persist Security Info=False"
    End If
    If Not lWRITE Then
      cARQ = cARQ & ";Mode=Share Deny Write"  '";Mode=Read"
    Else
      cARQ = cARQ & ";Mode=Share Deny None"  '";Mode=Read"
    End If
    TipoConn = Array("ADO", cARQ, "MDB")
    Exit Function
  End If

  If InStr(cARQTMP, "[A12MDB]") > 0 Then
    cARQ = Replace(cARQ, "[A12MDB]", "")
    If Len(cUSER) > 0 Then
      cARQ = cJetA12 & cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
    Else
      cARQ = cJetA12 & cARQ                ''& ";Persist Security Info=False"
    End If
    TipoConn = Array("ADO", cARQ, "MDB")
    Exit Function
  End If
  
  If InStr(cARQTMP, "[MYSQL]") > 0 Then
    cARQ = Replace(cARQ, "[MYSQL]", "")
    TipoConn = Array("ADO", cARQ, "MYSQL")
    Exit Function
  End If
  If InStr(cARQTMP, "[INFORMIX]") > 0 Then
    cARQ = Replace(cARQ, "[INFORMIX]", "")
    TipoConn = Array("ADO", cARQ, "INFORMIX")
    Exit Function
  End If
  If InStr(cARQTMP, "[FIREBIRD]") > 0 Then
    cARQ = Replace(cARQ, "[FIREBIRD]", "")
    TipoConn = Array("ADO", cARQ, "FIREBIRD")
    Exit Function
  End If
  If InStr(cARQTMP, "[INTERBASE]") > 0 Then
    cARQ = Replace(cARQ, "[INTERBASE]", "")
    TipoConn = Array("ADO", cARQ, "INTERBASE")
    Exit Function
  End If
  If InStr(cARQTMP, "[POSTGRESQL]") > 0 Then
    cARQ = Replace(cARQ, "[POSTGRESQL]", "")
    TipoConn = Array("ADO", cARQ, "POSTGRESQL")
    Exit Function
  End If
  If InStr(cARQTMP, "[UDL]") > 0 Then
    cARQ = Replace(cARQ, "[UDL]", "")
    cARQ = "File Name=" & cARQ & ";"
    TipoConn = Array("ADO", cARQ, "UDL")
    Exit Function
  End If
  'ADS
  If InStr(cARQTMP, "[ADS") > 0 Then
    If InStr(cARQTMP, "[ADSCDX]") > 0 Then
      cADSTIP = "ADS_CDX"
      cADSNOM = "ADSCDX"
    End If
    If InStr(cARQTMP, "[ADSNTX]") > 0 Or InStr(cARQTMP, "[SDENTX]") > 0 Then
      cADSTIP = "ADS_NTX"
      cADSNOM = "ADSCDX"
    End If
    If InStr(cARQTMP, "[ADSADT]") > 0 Then
      cADSTIP = "ADS_ADT"
      cADSNOM = "ADSADT"
    End If
    If InStr(cARQTMP, "[ADSVFP]") > 0 Then
      cADSTIP = "ADS_VFP"
      cADSNOM = "ADSVFP"
    End If

    cARQ = Replace(cARQ, "[" + cADSNOM + "]", "")
    cARQ = "Provider=Advantage.OLEDB.1;" & _
           "Mode=Share Deny None;" & _
           "Show Deleted Records in DBF Tables with Advantage=False;" & _
           "Data Source=" & cARQ & ";Advantage Server Type=ADS_Local_Server;" & _
           "TableType=" + cADSTIP + ";Security Mode=ADS_IGNORERIGHTS;" & _
           "Lock Mode=Compatible;" & _
           "Use NULL values in DBF Tables with Advantage=True;" & _
           "Exclusive=No;Deleted=No;" & _
           "Persist Security Info=False;Advantage Character Data Type=ADS_OEM;"
    TipoConn = Array("ADO", cARQ, cADSNOM)
    Exit Function
  End If
  'tratamentos JET
  cJETUSO = cJetPro  ''usa padrao jet mas a12 a 16 tem especifico
  'a12MDB a16MDB acima
  'aqui a12 a16 pdx dbiii seta o jetuso
  'retorno abaixo com o jet jeta12 ou jeTa16
  cXLSVER = "8.0"
  If InStr(cARQTMP, "[A12") > 0 Then
    cARQTMP = Replace(cARQTMP, "[A12", "[JET")
    cARQ = Replace(cARQ, "[A12", "[JET")
    cJETUSO = cJetA12
    cXLSVER = "12.0"
  End If
  If InStr(cARQTMP, "[A16") > 0 Then
    cARQTMP = Replace(cARQTMP, "[A16", "[JET")
    cARQ = Replace(cARQ, "[A16", "[JET")
    cJETUSO = cJetA16
    cXLSVER = "12.0"
  End If
  If InStr(cARQTMP, "[XLS]") > 0 Then
    cARQ = Replace(cARQ, "[XLS]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "Excel " + cXLSVER + ";HDR=Yes" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "XLS")
    Exit Function
  End If
  If InStr(cARQTMP, "[XLSX]") > 0 Then
    cARQ = Replace(cARQ, "[XLSX]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "Excel " + cXLSVER + ";HDR=Yes" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "XLSX")
    Exit Function
  End If
  If InStr(cARQTMP, "[XLSM]") > 0 Then
    cARQ = Replace(cARQ, "[XLSM]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "Excel " + cXLSVER + ";HDR=Yes" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "XLSM")
    Exit Function
  End If
  If InStr(cARQTMP, "[XLSB]") > 0 Then
    cARQ = Replace(cARQ, "[XLSB]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "Excel 12.0;HDR=Yes" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "XLSB")
    Exit Function
  End If

  If InStr(cARQTMP, "[JETTXTPIPE]") > 0 Then
    'Delimited(x)   File is considered as a delimited file with delimited character ‘x’.
    cARQ = Replace(cARQ, "[JETTXT]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "text;HDR=Yes;FMT=Delimited(|)" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "JETXT")
    Exit Function
  End If
  If InStr(cARQTMP, "[JETTXTTAB]") > 0 Then
    'TabDelimited   File is considered as a tab delimited file.
    cARQ = Replace(cARQ, "[JETTXT]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "text;HDR=Yes;FMT=TabDelimited" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "JETXT")
    Exit Function
  End If
  If InStr(cARQTMP, "[JETTXTCSV]") > 0 Or InStr(cARQTMP, "[JETTXT]") > 0 Then
    'Delimited      File is considered as a comma delimited file. Comma is the default delimited character.
    cARQ = Replace(cARQ, "[JETTXT]", "")
    cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "text;HDR=Yes;FMT=Delimited" & Chr(34) & ";"
    TipoConn = Array("ADO", cARQ, "JETXT")
    Exit Function
  End If
  'dbfiii com jet a12 ou a16
  If InStr(cARQTMP, "[JETDBFIII]") > 0 Then
    cARQ = Replace(cARQ, "[JETDBFIII]", "")
    cARQ = cJETUSO & cARQ & cJetExt & "DBASE III" & ";"
    TipoConn = Array("ADO", cARQ, "DBF")
    Exit Function
  End If
  'paradox com jet a12 ou a16
  If InStr(cARQTMP, "[JETPD") > 0 Then
    If InStr(cARQTMP, "[JETPDX3]") > 0 Then
      cARQ = Replace(cARQ, "[JETPDX3]", "")
      cARQ = cJETUSO & cARQ & cJetExt & "Paradox 3.X" & ";"
      TipoConn = Array("ADO", cARQ, "DB")
      Exit Function
    End If
    If InStr(cARQTMP, "[JETPDX4]") > 0 Then
      cARQ = Replace(cARQ, "[JETPDX4]", "")
      cARQ = cJETUSO & cARQ & cJetExt & "Paradox 4.X" & ";"
      TipoConn = Array("ADO", cARQ, "DB")
      Exit Function
    End If
    If InStr(cARQTMP, "[JETPDX5]") > 0 Then
      cARQ = Replace(cARQ, "[JETPDX5]", "")
      cARQ = cJETUSO & cARQ & cJetExt & "Paradox 5.X" & ";"
      TipoConn = Array("ADO", cARQ, "DB")
      Exit Function
    End If
  End If
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

Public Function ADORsStatus(ByRef eSTATUS)
  ADORsStatus = "Status Nao Especificado"
  Select Case eSTATUS
  Case adRecOK
    ADORsStatus = "Registro Gravado Corretamente"
  Case adRecNew
    ADORsStatus = "Novo Registro"
  Case adRecModified
    ADORsStatus = "Registro Modificado"
  Case adRecDeleted
    ADORsStatus = "Registro Deletado"
  Case adRecUnmodified
    ADORsStatus = "Registro nao modificado"
  Case adRecInvalid
    ADORsStatus = "Nao Salvo - Invalido BookMark"
  Case adRecMultipleChanges
    ADORsStatus = "Nao Salvo - Afetou muitos Registro"
  Case adRecPendingChanges
    ADORsStatus = "Nao Salvo - Insercao Pendente"
  Case adRecCanceled
    ADORsStatus = "Nao Salvo - Operacao Cancelada"
  Case adRecCantRelease
    ADORsStatus = "Nao Salvo - Existe registro travado"
  Case adRecConcurrencyViolation
    ADORsStatus = "Nao Salvo - Cursor Optimista em Uso"
  Case adRecIntegrityViolation
    ADORsStatus = "Nao Salvo - Violacao de Integridade"
  Case adRecMaxChangesExceeded
    ADORsStatus = "Nao Salvo - Muitas Pendencias"
  Case adRecObjectOpen
    ADORsStatus = "Nao Salvo - Falha no meio de Armazenamento"
  Case adRecOutOfMemory
    ADORsStatus = "Nao Salvo - Erro Memoria Computador"
  Case adRecPermissionDenied
    ADORsStatus = "Nao Salvo - Falta permissao"
  Case adRecSchemaViolation
    ADORsStatus = "Nao Salvo - Violacao da Estrutura do Arquivo"
  Case adRecDBDeleted
    ADORsStatus = "Registro Ja Deletado"
  Case Else
    ADORsStatus = "Status Nao Especificado"
  End Select

End Function

Function ADOErro(ByRef oErro As Variant, Optional ByVal cERRO As String = "")
  Dim errorObject As ADODB.error
  For Each errorObject In oErro
    cERRO = cERRO & " Ado Erro Numero: " & errorObject.Number & vbCrLf
    cERRO = cERRO & " Ado Descricao  : " & errorObject.Description & vbCrLf
    cERRO = cERRO & " Ado Fonte      : " & errorObject.Source & vbCrLf
    cERRO = cERRO & " Ado SQL        : " & errorObject.SQLState & vbCrLf
    cERRO = cERRO & " Ado Erro Nativo: " & errorObject.NativeError & vbCrLf
  Next
  SayErro cERRO, True
End Function

Public Function ADO_IsRecordsetEmpty(ByRef oRS As ADODB.Recordset) As Boolean
  Dim cERRO As String
  On Error GoTo trataerro
  ADO_IsRecordsetEmpty = False
  If Not ADO_IsOpen(oRS) Then
    ADO_IsRecordsetEmpty = True
    Exit Function
  End If
  If oRS.EOF And oRS.BOF Then
    ADO_IsRecordsetEmpty = True
    Exit Function
  End If
  If oRS.RecordCount = 0 Then
    ADO_IsRecordsetEmpty = True
    Exit Function
  End If
  Exit Function
trataerro:
  cERRO = "ADO_IsRecordsetEmpty"
  cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
  Select Case Err.Number
  Case Else
    ADOErro oRS.ActiveConnection.Errors, cERRO
  End Select
End Function

Public Function ADO_FieldValueToString(ByRef FLD As ADODB.Field, Optional ByVal sNullRepresentation As String = "(null)") As String
'Call TraceEnters(MODULE_NAME & "::ADO_FieldValueToString")
'TraceDetail = "To convert the value of a field into a string"
  On Error GoTo trataerro

  ADO_FieldValueToString = ""
  Select Case (FLD.Type)
    '~~~~~~
  Case adBSTR, adChar, adVarChar, adVarWChar, adWChar, adLongVarChar, adLongVarWChar, adGUID:
    ADO_FieldValueToString = Nz(FLD.Value, sNullRepresentation)
    '~~~~~~
  Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adBoolean:
    ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adVariant:
    ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adDate, adDBDate, adDBTime, adDBTimeStamp:
    ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adBinary, adVarBinary, adLongVarBinary:
    ADO_FieldValueToString = IIf(IsNull(FLD.Value), sNullRepresentation, BytesToHexString(FLD.Value))
    '~~~~~~
  Case Else:
    'Err.Raise 13, MODULE_NAME & "::ADO_FieldValueToString", "Sorry, unable to convert fields of type " & fld.Type & ", (" & ADO_DataTypeEnumToEnglish(fld.Type) & ") to string."
  End Select
  Exit Function
trataerro:
  SayErro "ADO_FieldValueToString"
  Exit Function
End Function

Public Sub ADO_FreeRecordset(ByRef rs As ADODB.Recordset)
  On Error Resume Next
  If rs.State = adStateOpen Then
    rs.Close
  End If
  Set rs = Nothing
  On Error GoTo 0
End Sub

Public Function ADO_IsOpen(ByRef oADOObject As Object) As Boolean
' Purpose: To determine if a connection or a recordset is open
' !! Assumes/Pre: Nothing
' Parameters:
'  oADOObject as Object  -
' Returns: Boolean
'       Success-
'       Failure- Raises error on failure
' Revision history:
'   2004-Feb-20 10:47 [Michael Johnson] Initial creation
'Call TraceEnters(MODULE_NAME & "::ADO_IsOpen")
'TraceDetail = "To determine if a connection or a recordset is open"
  On Error GoTo trataerro

  ADO_IsOpen = False
  If oADOObject Is Nothing Then
    Exit Function
  End If

  If oADOObject.State = adStateOpen Then
    ADO_IsOpen = True
    Exit Function
  End If
  Exit Function

trataerro:
  SayErro "ADO_ISOPEN"

End Function

Public Function Nz( _
       vValue As Variant, _
       Optional vReplacementIfNull As Variant = 0 _
     ) As Variant
' Purpose: To replace a NULL with another value, if the value is Null.
' Example/Note:     sResult = Nz(rs.Fields(sFieldName), "") ' See MS Access VBA for Nz documentation
' !! Assumes/Pre: Nothing
' Parameters:
'   vValue- Value to evaluate if null
'   vReplacementIfNull - What should replace a Null value
' Returns: Variant
'       Success- If not null, returns the supplied value, else returns the replacement
'       Failure- Raises error on failure
' Revision history:
'   Michael Johnson     2002-Mar-12 1243     Initial creation
'   2003-Aug-22 10:17 [Michael B. Johnson] Abreviated and changed to variants
'   2004-Feb-18 16:36 [Michael B. Johnson] Changed from using TypeName() to IsNull()
'Call TraceEnters(MODULE_NAME & "::Nz")
'TraceDetail = "To replace a NULL with a string, if the value is Null."

  If IsNull(vValue) Then
    Nz = vReplacementIfNull
  Else
    Nz = vValue
  End If

  Exit Function
End Function

Public Function BytesToHexString(vaBytes As Variant) As String
' Purpose: To translate a Byte() Array into human readable Format
' Example/Note: BytesToHexString
' !! Assumes/Pre: Nothing
' Parameters:
'   vaBytes - Variant byte array
' Returns: String
'       Success- String with leading "0x" to denote hexadecimal format.
'       Failure- Raises error on failure
' Dependencies: None
'        mod->Sub
' Revision history:
'   Michael Johnson     2000/Aug/01 13:51     Initial creation
  On Error GoTo trataerro

  Dim sAccumulator As String, lCtr As Long
  Dim sHex As String, sFormatted As String

  BytesToHexString = ""
  If Not TypeName(vaBytes) = "Byte()" Then
    BytesToHexString = ""
    Exit Function
  End If

  For lCtr = 0 To UBound(vaBytes)
    sHex = Hex(vaBytes(lCtr))
    sFormatted = Format(sHex, "@@")
    sAccumulator = sAccumulator & Replace(sFormatted, " ", "0")
  Next
  sAccumulator = "0x" & sAccumulator
  BytesToHexString = sAccumulator
  Exit Function
trataerro:
  BytesToHexString = ""
  Exit Function
End Function
Public Function TemTabelaADO(ByVal cARQ As String, ByVal cTabela As String, Optional ByVal lMES As Boolean = True) As Boolean
  Dim oCat As ADOX.Catalog
  Dim oTabela As ADOX.Table
  On Error GoTo trataerro

  TemTabelaADO = False
  Set oCat = New ADOX.Catalog
  oCat.ActiveConnection = cARQ

  For Each oTabela In oCat.Tables
    If UCase(oTabela.Name) = UCase(cTabela) Then
      TemTabelaADO = True
      Exit For
    End If
  Next

  If lMES And Not TemTabelaADO Then
    Alert ("Tabela nao Encontrada" & cTabela & Chr(13) & Chr(10) & cARQ)
  End If
trataerro:
  Select Case Err.Number
  Case Else
    SayErro "Tem Tabela Ado :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTabela & Chr(13) & Chr(10)
    Exit Function
  End Select

End Function

Public Function AdoNewTable(ByVal cARQORI As String, Optional ByVal lCRIA As Boolean = False, _
                            Optional ByVal Ntipo As Integer = 5) As Boolean
  Dim cat As New ADOX.Catalog
  On Error GoTo trataerro
  AdoNewTable = False
  If Not FileExist(cARQORI, True) Then
    If lCRIA Or MDG("Criar Arquivo " & cARQORI) Then
      cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cARQORI & _
                 ";Jet OLEDB:Engine Type=" & Ntipo & ";"
      AdoNewTable = True
    End If
  End If
  Exit Function
trataerro:
  Select Case Err.Number
  Case Else
    SayErro "ADO Novo Arquivo Access/MDB:" & Chr(13) & Chr(10) & cARQORI
    Exit Function
  End Select
End Function



