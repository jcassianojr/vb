Attribute VB_Name = "dcompareMain"
Sub Main()
'InitXPStyles

  DAO.DBEngine.SetOption dbExclusiveAsyncDelay, 2000
  DAO.DBEngine.SetOption dbSharedAsyncDelay, 0
  DAO.DBEngine.SetOption dbFlushTransactionTimeout, 500
  '
  DAO.DBEngine.SetOption dbUserCommitSync, "yes"
  DAO.DBEngine.SetOption dbImplicitCommitSync, "yes"
  DAO.DBEngine.SetOption dbLockRetry, 20
  DAO.DBEngine.SetOption dbPageTimeout, 5000
  DAO.DBEngine.SetOption dbMaxLocksPerFile, 9500
  DAO.DBEngine.SetOption dbLockDelay, 100
  DAO.DBEngine.SetOption dbRecycleLVs, 0


  '    adodb.E

  EnabledVisualStyles
  dCompare.Show
End Sub

Function formatarg()
  formatarg = True
End Function
