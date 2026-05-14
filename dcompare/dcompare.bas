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



