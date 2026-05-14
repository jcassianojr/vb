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

  EnabledVisualStyles
  dCompare.Show
End Sub

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



