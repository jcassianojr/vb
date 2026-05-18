Attribute VB_Name = "AdoLib"
Option Explicit
'
'There is no provider named 'Microsoft.ACE.OLEDB.14.0' even though it's Access 2010 (aka version number 14) the provider
'that should be used
'still is named with version number 12.
'So change from 'Microsoft.ACE.OLEDB.14.0' to 'Microsoft.ACE.OLEDB.12.0' and it will probably work better!
'
Public Const cJetA12 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Public Const cJetADV = "Provider=Advantage.OLEDB.1;Data Source="
Public Const cJetExt = ";Extended Properties="
Public Const cJetPro = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

Public Const JET_ENGINETYPE_UNKNOWN = 0
Public Const JET_ENGINETYPE_JET10 = 1
Public Const JET_ENGINETYPE_JET11 = 2
Public Const JET_ENGINETYPE_JET2X = 3
Public Const JET_ENGINETYPE_JET3X = 4
Public Const JET_ENGINETYPE_JET4X = 5
Public Const Jet_EngineType_Ace12 = 6 'accdb
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
Public Function GeracArq(ByVal cARQ As String, Optional ByVal cTipo As String = "", Optional ByVal lWRITE As Boolean = True) As String
  Dim aRETU As Variant
  cARQ = GeraConn(cARQ, cTipo)
  aRETU = TipoConn(cARQ, , , lWRITE)
  GeracArq = aRETU(1)
End Function


Public Function GeraConn(ByVal cARQ As String, Optional cTipo As String = "") As String
  Dim nPOS As Long
  Dim cARQTMP As String
  
  GeraConn = cARQ
  cARQTMP = UCase(cARQ) 'usado no instr maiscula mas sempre atribui na carq conecao e case sensitive
  If InStr(cARQTMP, "[") > 0 Then                 ''ja e uma connecao
    Exit Function
  End If
  If cTipo = "JETMDB" Or cTipo = "MDB" Or InStr(cARQTMP, ".MDB") > 0 Then
    GeraConn = "[JETMDB]" & cARQ
    Exit Function
  End If
  If cTipo = "JETFOX" Or cTipo = "FOX" Or cTipo = "DBF" Or InStr(cARQTMP, ".DBF") > 0 Then
    GeraConn = "[JETFOX]" & cARQ
    Exit Function
  End If
  
'
' foxpro .dbf 'vir antes do sqlite para nao confundir com .db tambem paradox.db
'
  
  If InStr(cARQTMP, ".DBF") > 0 Then  'DBF
    nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
    cARQ = Mid(cARQ, 1, nPOS)
    GeraConn = "[JETFOX]" & cARQ
    Exit Function
  End If
  
  If cTipo = "SQLITE" Or InStr(LCase(cARQ), ".sqlite") > 0 Or InStr(LCase(cARQ), ".sqlite3") > 0 _
                 Or InStr(LCase(cARQ), ".fossil") > 0 Or InStr(LCase(cARQ), ".db3") > 0 _
                 Or (InStr(LCase(cARQ), ".db") > 0 And InStr(cARQTMP, ".DBF") = 0) Then
    GeraConn = "[SQLITE]" & cARQ
    Exit Function
  End If

  If LCase(Right(cARQ, 6)) = ".accdb" Then  'InStr(LCase(cARQ), ".accdb")>0
    GeraConn = "[ACCDB]" & cARQ
    Exit Function
  End If
  
  If InStr(LCase(cARQ), ".mariadb") > 0 Then
    GeraConn = "[MARIADB]" & cARQ
    Exit Function
  End If
  
  If InStr(LCase(cARQ), ".oracle") > 0 Then
    GeraConn = "[ORACLE]" & cARQ
    Exit Function
  End If
  
  
  
  If InStr(LCase(cARQ), ".pgsql") > 0 Or InStr(LCase(cARQ), ".postgresql") > 0 Then
    GeraConn = "[PGSQL]" & cARQ
    Exit Function
  End If
  
  
  If InStr(LCase(cARQ), ".mysql") > 0 Then
    GeraConn = "[MYSQL]" & cARQ
    Exit Function
  End If
      
  If InStr(LCase(cARQ), ".oracle") > 0 Then
    GeraConn = "[ORACLE]" & cARQ
    Exit Function
  End If
    

  
  If InStr(cARQTMP, ".PD") > 0 Then   ' paradox
    nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
    cARQ = Mid(cARQ, 1, nPOS)
    GeraConn = "[JETPDX5]" & cARQ
    Exit Function
  End If
  
  If Mid(cTipo, 1, 5) = "ACCDB" Then ' ACCDBDBFIII  ACCDBMDB PX3 PX4 PX5 XLS XLXS XLXB, CSV... _
                                   cada versao jet suporta extensoes em outros formatos de arquivos _
                                   posicao 2,3 versao 12,13,14,115 _
                                   posicao 4 MDB ACCDB formato extensao
     GeraConn = "[ACCDB" & Mid(cTipo, 6) & "]"
     Exit Function
  End If
 
  
If Len(cTipo) > 0 Then
  Select Case cTipo
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
      Case "ADSNTX"
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
  End Select
End If
End Function

Public Function TipoConn(ByVal cARQ As String, Optional ByVal cUSER As String = "", _
                         Optional ByVal cPASS As String = "", Optional ByVal lWRITE As Boolean = True) As Variant
  Dim cARQTMP As String
  Dim cJETUSO As String
  Dim lTEMMDB As Boolean
  Dim lTEMSQLITE As Boolean
  Dim lTEMMARIADB As Boolean
  Dim lTEMPG As Boolean
  Dim lTEMMYSQL As Boolean
  Dim lTEMORACLE As Boolean
  
  Dim cADSTIP As String
  Dim cADSNOM As String
  Dim cXLSVER As String '
  Dim aCONN As Variant
  
  'usa boolean para agilizar if,case... no lugar de comparacao com string
  lTEMMDB = False
  lTEMSQLITE = False
  lTEMMARIADB = False
  lTEMMYSQL = False
  lTEMPG = False
  lTEMORACLE = False
  
  
  'inicial valores padrao
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
  'mariadb
  '
  If InStr(cARQTMP, ".MARIADB") > 0 Or InStr(cARQTMP, "{MARIADB") > 0 Then
    lTEMMARIADB = True
    TipoConn = Array("ADO", cARQ, "MARIADB")
  End If
  
  '
  'mariadb
  '
  If InStr(cARQTMP, ".ORACLE") > 0 Or InStr(cARQTMP, "{ORACLE") > 0 Then
    lTEMORACLE = True
    TipoConn = Array("ADO", cARQ, "ORACLE")
  End If
   
  '
  'mysql
  '
  If InStr(cARQTMP, ".MYSQL") > 0 Or InStr(cARQTMP, "{MYSQL") > 0 Then
    lTEMMYSQL = True
    TipoConn = Array("ADO", cARQ, "MYSQL")
  End If
  
    '
  'postgresql pgsql
  '
  If InStr(cARQTMP, ".PGSQL") > 0 Or InStr(cARQTMP, ".POSTGRESQL") > 0 Or InStr(cARQTMP, "{POSTGRESQL") > 0 Then
    lTEMPG = True
    TipoConn = Array("ADO", cARQ, "PGSQL")
  End If
    '
 
  '
  'checando provider,driver connecao
  '
  If InStr(cARQ, "[") = 0 Then 'inclui se nao tiver [ evita duplicar [USO][USO]
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
    If lTEMMARIADB Then
      cARQ = "[MARIADB]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
   If lTEMORACLE Then
      cARQ = "[ORACLE]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
    If lTEMMYSQL Then
      cARQ = "[MYSQL]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
    If lTEMPG Then
      cARQ = "[PGSQL]" & cARQ
      cARQTMP = UCase(cARQ)
    End If
    
  End If
'
'access
'
  If lTEMMDB Then
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
  If lTEMSQLITE Then  'c:\Program Files (x86)\SQLite ODBC Driver\readme.txt http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe
    cARQ = Replace(cARQ, "[SQLITE]", "")
    If InStr(cARQTMP, "SQLITE3 ODBC DRIVER") = 0 Then
       cARQ = "Driver={SQLite3 ODBC Driver};Database=" + cARQ + ";"
    End If
    TipoConn = Array("ADO", cARQ, "SQLITE")
    Exit Function
  End If
  
  If lTEMMARIADB Then
    cARQ = Replace(cARQ, "[MARIADB]", "")
    If InStr(cARQTMP, "MARIADB ODBC") = 0 Then 'geracom se nao passado
       aCONN = Split(cARQ, ".") 'localhost.port.mariadb.banco 'localhost.3306.mariadb.citacao
       If Len(cUSER) > 0 Then
         cARQ = "DRIVER={MariaDB ODBC 3.2 Driver};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PASSWORD=" + cPASS + ";"
       Else
         cARQ = "DRIVER={MariaDB ODBC 3.2 Driver};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=root;PASSWORD=admin;"
       End If
    End If
    TipoConn = Array("ADO", cARQ, "MARIADB")
    Exit Function
  End If
  
  
  
 If lTEMORACLE Then
'    "Provider=OraOLEDB.Oracle;data source=" & _
 '   sWorld & ".World;User id=" & sUID & ";password=" & sPWD & ";"

    cARQ = Replace(cARQ, "[ORACLE]", "")
    If InStr(cARQTMP, "ODBC For Oracle") = 0 Then 'geracom se nao passado
       aCONN = Split(cARQ, ".") 'localhost.port.mariadb.banco 'localhost.3306.mariadb.citacao
       If Len(cUSER) > 0 Then
         cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PWD=" + cPASS + ";"
       Else
         cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=root;PASSWORD=admin;"
       End If
    End If
    TipoConn = Array("ADO", cARQ, "ORACLE")
    Exit Function
  End If
  
  
  
  

  If lTEMMYSQL Then
    cARQ = Replace(cARQ, "[MYSQL]", "")
    If InStr(cARQTMP, "MYSQL ODBC") = 0 Then 'geracom se nao passado
       aCONN = Split(cARQ, ".") 'localhost.port.mysql.banco 'localhost.5432.mysql.citacao
       If Len(cUSER) > 0 Then
          cARQ = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=" + cUSER + ";Pwd=" + cPASS + ";"     '32 driver versao 8
       Else
         cARQ = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=root;Pwd=admin;"    '32 driver versao 8
       End If
    End If
    TipoConn = Array("ADO", cARQ, "MYSQL")
    Exit Function
  End If

   If lTEMPG Then
    cARQ = Replace(cARQ, "[PGSQL]", "")
    cARQ = Replace(cARQ, "[POSTGRESQL]", "")
    If InStr(cARQTMP, "POSTGRESQL ANSI") = 0 Then 'geracom se nao passado
       aCONN = Split(cARQ, ".") 'localhost.port.postgresql.banco 'localhost.5432.postgresql.citacao
       If Len(cUSER) > 0 Then
          cARQ = "Driver={PostgreSQL ANSI};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=" + cUSER + ";Pwd=" + cPASS + ";"
       Else
          cARQ = "Driver={PostgreSQL ANSI};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=postgres;Pwd=admin;"
       End If
    End If
    TipoConn = Array("ADO", cARQ, "PGSQL")
    Exit Function
  End If

  '
  'jetfox
  '
  If InStr(cARQTMP, "[JETFOX]") > 0 Then
    cARQ = Replace(cARQ, "[JETFOX]", "")

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
    If lTEMMDB Then
      TipoConn = Array("ADO", cARQ, "MDB")
    Else
      TipoConn = Array("ADO", cARQ, "CONN")
    End If
    Exit Function
  End If
  If InStr(cARQTMP, "[SQLSERVER]") > 0 Or InStr(cARQTMP, "[MSSQL]") > 0 Then
     cARQ = Replace(cARQ, "[SQLSERVER]", "")
     cARQ = Replace(cARQ, "[MSSQL]", "")
       If Len(cUSER) > 0 Then
         cARQ = "Provider=" + MSSqlOledbProvider(1) + ";Server=" + cARQ + ";Uid=" + cUSER + ";Pwd=" + cPASS + ";"
       Else
         cARQ = "Provider=" + MSSqlOledbProvider(1) + ";Server=" + cARQ + ";"
       End If
    TipoConn = Array("ADO", cARQ, "SQLSERVER")
    Exit Function
  End If

  If InStr(cARQTMP, "[ACCDB") > 0 Then
    cJETUSO = cJetA12
    cARQ = Replace(cARQ, "[ACCDB", "")
    cARQ = Replace(cARQ, "MDB]", "")
    cARQ = Replace(cARQ, "]", "")
    If Len(cUSER) > 0 Then
      cARQ = cJETUSO & cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
    Else
      cARQ = cJETUSO & cARQ                ''& ";Persist Security Info=False"
    End If
    If Not lWRITE Then
      cARQ = cARQ & ";Mode=Share Deny Write"  '";Mode=Read"
    Else
      cARQ = cARQ & ";Mode=Share Deny None"  '";Mode=Read"
    End If
    TipoConn = Array("ADO", cARQ, "ACEOLEDB")
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
    If InStr(cARQTMP, "[ADSNTX]") > 0 Then
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
  cJETUSO = cJetPro
  cXLSVER = "8.0"
  If InStr(cARQTMP, "12") > 0 Then
    cJETUSO = cJetA12
    cXLSVER = "12.0"
  End If
  If InStr(cARQTMP, "14") > 0 Then
    cJETUSO = cJetA12
    cXLSVER = "14.0"
  End If
  If InStr(cARQTMP, "15") > 0 Then
    cJETUSO = cJetA12
    cXLSVER = "15.0"
  End If
  If InStr(cARQTMP, "16") > 0 Then
    cJETUSO = cJetA12
    cXLSVER = "16.0"
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

Public Function ADORsStatus(ByRef eSTATUS) As String
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
  Dim errorObject As ADODB.Error
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
  On Error GoTo TrataErro
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
TrataErro:
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
  On Error GoTo TrataErro

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
TrataErro:
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

    On Error GoTo TrataErro  'resume Next 'Evita quebra se o objeto estiver em estado inválido
        ADO_IsOpen = False
    If oADOObject Is Nothing Then Exit Function
    
    'Verifica se o estado tem o bit de "Open" ligado
    ADO_IsOpen = ((oADOObject.State And adStateOpen) = adStateOpen)
        
 ' On Error GoTo trataerro

 ' ADO_IsOpen = False
 ' If oADOObject Is Nothing Then
 '   Exit Function
 ' End If

 ' If oADOObject.State = adStateOpen Then
 '   ADO_IsOpen = True
 '   Exit Function
 ' End If
  Exit Function

TrataErro:
  SayErro "ADO_ISOPEN"

End Function
'existe na commom.bas
'Public Function Nz( _
'       vValue As Variant, _
'       Optional vReplacementIfNull As Variant = 0 _
'     ) As Variant
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

 ' If IsNull(vValue) Then
 '   Nz = vReplacementIfNull
 ' Else
 '   Nz = vValue
 ' End If

  'Exit Function
'End Function

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
  On Error GoTo TrataErro

  Dim sAccumulator As String
  Dim lCtr As Long
  Dim sHex As String
  Dim sFormatted As String

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
TrataErro:
  BytesToHexString = ""
  Exit Function
End Function
Public Function TemTabelaADO(ByVal cARQ As String, ByVal cTABELA As String, Optional ByVal lMES As Boolean = True) As Boolean
  Dim oCat As ADOX.Catalog
  Dim oTabela As ADOX.Table
  On Error GoTo TrataErro

  TemTabelaADO = False
  Set oCat = New ADOX.Catalog
  oCat.ActiveConnection = cARQ

  For Each oTabela In oCat.Tables
    If UCase(oTabela.Name) = UCase(cTABELA) Then
      TemTabelaADO = True
      Exit For
    End If
  Next

  If lMES And Not TemTabelaADO Then
    Alert ("Tabela nao Encontrada" & cTABELA & Chr(13) & Chr(10) & cARQ)
  End If
TrataErro:
  Select Case Err.Number
  Case Else
    SayErro "Tem Tabela Ado :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTABELA & Chr(13) & Chr(10)
    Exit Function
  End Select

End Function

Public Function CriaMdbAccess(ByVal cARQORI As String, Optional ByVal lCRIA As Boolean = False, _
                            Optional ByVal nTIPO As Integer = 5) As Boolean
                            
    'OLEDB:Engine Type=5
  'Unknown                      0
  'Microsoft Jet 1.0            1
  'Microsoft Jet 1.1            2
  'Microsoft Jet 2.0            3
  'Microsoft Jet 3.x(97)        4
  'Microsoft Jet 4.x(2000)      5
  'JetEngineType_Ace12 =        6         accdb
                            
                            
  Dim cat As New ADOX.Catalog
  On Error GoTo TrataErro
  CriaMdbAccess = False
  If Not FileConnExist(cARQORI, True) Then
    If lCRIA Or MDG("Criar Arquivo " & cARQORI) Then
      If nTIPO = 6 Or InStr(LCase(cARQORI), ".accdb") > 0 Then
           cat.Create ("Provider=Microsoft.ACE.OLEDB.12;" & _
                        "Data Source=" & cARQORI & ";" & _
                        "JET OLEDB:Engine Type=6;")
      Else
          cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cARQORI & _
                  ";Jet OLEDB:Engine Type=" & nTIPO & ";"
      End If
      CriaMdbAccess = True
    End If
  End If
  Exit Function
TrataErro:
  Select Case Err.Number
  Case Else
    SayErro "ADO Novo Arquivo Access/MDB:" & Chr(13) & Chr(10) & cARQORI
    Exit Function
  End Select
End Function


Public Function MSSqlOdbcDriver() As String
    
    'Customize this list to include the drivers you want to support
    Dim SupportedDrivers As Variant
    SupportedDrivers = Array( _
        "ODBC Driver 17 for SQL Server", _
        "ODBC Driver 13 for SQL Server", _
        "SQL Server Native Client 11.0", _
        "SQL Server")   'the "SQL Server" driver is the legacy driver that is included with Windows
    
    'We're using the registry to check for Drivers, so we'll just create the registry object
    '   once and use it each time through the loop below
    Dim oReg As Object
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Const RegPath As String = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    Const HKEY_LOCAL_MACHINE = &H80000002
    
    
    Dim i As Long
    For i = 0 To UBound(SupportedDrivers)
        Dim driverName As String: driverName = SupportedDrivers(i)
        
        Dim KeyExists As Boolean, RegValue As String
        KeyExists = (oReg.GetStringValue(HKEY_LOCAL_MACHINE, RegPath, driverName, RegValue) = 0)
        
        Dim DriverIsInstalled
        DriverIsInstalled = (KeyExists And (RegValue = "Installed"))
        
        If DriverIsInstalled Then
            'Return the first Driver from the list that is installed on the computer
            MSSqlOdbcDriver = driverName
            Set oReg = Nothing
            Exit Function
        End If
    Next i
    Set oReg = Nothing
    
    'If we get here it means there were no matches found;
    '   until we work out all the kinks, we may be better served by returning the default Driver
    '   here rather than throwing an error
    MsgBox "No ODBC Drivers found (not even the default that ships with Windows!)"
End Function

' ----------------------------------------------------------------
' Procedure : GetLatestOledbProvider --MSSqlOledbProvider
' Date      : 12/14/2022
' Author    : Mike Wolfe
' Source    : https://nolongerset.com/getlatestoledbprovider/
' Purpose   : Iterates through a custom list of OLEDB providers
'               and returns the first installed match.
' ----------------------------------------------------------------
Public Function MSSqlOledbProvider(Optional ByVal nTIPO As Integer = 1) As String
    Dim SupportedProviders As Variant
    SupportedProviders = Array( _
        "MSOLEDBSQL19", "Microsoft OLE DB Driver 19 for SQL Server", "EE5DE99A-4453-4C96-861C-F8832A7F59FE", _
        "MSOLEDBSQL", "Microsoft OLE DB Driver for SQL Server", "5A23DE84-1D7B-4A16-8DED-B29C09CB648D", _
        "SQLNCLI11", "SQL Server Native Client 11.0", "397C2819-8272-4532-AD3A-FB5E43BEAA39", _
        "SQLNCLI10", "SQL Server Native Client 10.0", "8F4A6B68-4F36-4e3c-BE81-BC7CA4E9C45C", _
        "SQLOLEDB", "Microsoft OLE DB Provider for SQL Server", "0C7FF16C-38E3-11d0-97AB-00C04FC2AD98")
    
    'We're using the registry to check for providers, so we'll just create the registry object
    '   once and use it each time through the loop below
    Dim oReg As Object
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Const HKEY_CLASSES_ROOT = &H80000000
    
    Dim i As Long
    For i = 0 To UBound(SupportedProviders) Step 3
        Dim ProviderName As String: ProviderName = SupportedProviders(i)
        Dim ProviderDesc As String: ProviderDesc = SupportedProviders(i + 1)
        Dim ProviderUID As String: ProviderUID = SupportedProviders(i + 2)
        
        Dim ProviderIsInstalled As Boolean, SubkeyPath As String
        SubkeyPath = "CLSID\{" & ProviderUID & "}\"
        ProviderIsInstalled = (oReg.EnumKey(HKEY_CLASSES_ROOT, SubkeyPath, "", "") = 0)
        
        If ProviderIsInstalled Then
            'Return the first provider from the list that is installed on the computer
            Select Case nTIPO
                   Case 1
                       MSSqlOledbProvider = ProviderName
                   Case 2
                       MSSqlOledbProvider = ProviderDesc
                   Case 3
                       MSSqlOledbProvider = ProviderUID
            End Select
            Set oReg = Nothing
            Exit Function
        End If
    Next i
    Set oReg = Nothing
    
    'If we get here it means there were no matches found
    MsgBox "No OLE DB providers found (not even the default that ships with Windows!)"
End Function


Public Function ado_GetCursorType(CursorType As Integer) As String

    ' *******************
    ' CursorTypeEnum
    '
    ' adOpenDynamic = 2
    ' adOpenForwardOnly = 0
    ' adOpenKeyset = 1
    ' adOpenStatic = 3
    ' adOpenUnspecified = -1
    '*******************
    
    Select Case CursorType
        Case 2
            ado_GetCursorType = "Dynamic"
        Case 0
            ado_GetCursorType = "ForwardOnly"
       Case 1
            ado_GetCursorType = "Keyset"
        Case 3
            ado_GetCursorType = "Static"
        Case -1
            ado_GetCursorType = "Unspecified"
    End Select
End Function

Function ado_GetLockType(LockType As Integer) As String
    
    '*******************
    ' LockTypeEnum
    '
    ' adLockBatchOptimistic = 4
    ' adLockOptimistic = 3
    ' adLockPessimistic = 2
    ' adLockReadOnly = 1
    ' adLockUnspecified = -1
    '*******************
    
    Select Case LockType
        Case 4
            ado_GetLockType = "BatchOptimistic"
        Case 3
            ado_GetLockType = "Optimistic"
       Case 2
            ado_GetLockType = "Pessimistic"
        Case 1
            ado_GetLockType = "ReadOnly "
        Case -1
            ado_GetLockType = "Unspecified"
    End Select

End Function
Public Function ExtraiWhere(ByVal cSQL As String) As String
    Dim nPOS As Long
    Dim sUpperSQL As String
    
    sUpperSQL = UCase(cSQL)
    nPOS = InStr(sUpperSQL, "WHERE ")
    
    If nPOS > 0 Then
        ' Retorna do "WHERE" até ao fim da string
        ExtraiWhere = Trim(Mid(cSQL, nPOS))
    Else
        ' Se não houver WHERE, retorna vazio (CUIDADO: isso afetaria a tabela toda)
        ExtraiWhere = ""
    End If
End Function


Public Function ExtraiTabela(ByVal cSQL As String) As String
    ' Busca o nome da tabela após o "FROM"
    Dim nPOS As Long
    cSQL = UCase(cSQL)
    nPOS = InStr(cSQL, "FROM ")
    If nPOS > 0 Then
        ExtraiTabela = Trim(Mid(cSQL, nPOS + 5))
        ' Remove o resto se houver WHERE ou ORDER
        If InStr(ExtraiTabela, " ") > 0 Then
            ExtraiTabela = Left(ExtraiTabela, InStr(ExtraiTabela, " ") - 1)
        End If
    End If
End Function


Public Function PrepararValorSQL(ByVal v As Variant) As String
    If IsNull(v) Then
        PrepararValorSQL = "NULL"
    ElseIf IsNumeric(v) Then
        ' No SQLite, decimais devem usar PONTO, não vírgula
        PrepararValorSQL = Replace(CStr(v), ",", ".")
    ElseIf IsDate(v) Then
        PrepararValorSQL = "'" & Format(v, "yyyy-mm-dd") & "'"
    Else
        ' Texto: dobra as aspas simples para não quebrar o SQL
        PrepararValorSQL = "'" & Replace(v, "'", "''") & "'"
    End If
End Function
