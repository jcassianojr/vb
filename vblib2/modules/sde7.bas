Attribute VB_Name = "SDE7"
' SDE7.DLL Version 7.0.0.0
' ----------

' file open modes
' ---------------
Global Const READWRITE = 0
Global Const READONLY = 1
Global Const EXCLUSIVE = 2

' SDE RDE types
' -------------
Global Const SDENTX = 1
Global Const SDEFOX = 2
Global Const SDENSX = 3
Global Const SDENSX_DBT = 4
Global Const SDEVFOX = 5

' SDE text file types
' -------------------
Global Const COMMA_DELIM = 21
Global Const SDF_FILE = 22
Global Const TAB_DELIM = 23
Global Const SPACE_DELIM = 24
Global Const USER_DELIM = 25

' OEM Source types for AppendFrom
' -------------------------------
Global Const OEMNTX = 31
Global Const OEMFOX = 32
Global Const OEMNSX = 33

' SDE Index Types
' ------------------
Global Const INDEX_STANDARD = 1
Global Const INDEX_STANDARD_UNIQUE = 2
Global Const INDEX_CONDITIONAL = 3
Global Const INDEX_CONDITIONAL_UNIQUE = 4

' Country (for sx_SetDateFormat)
' ------------------------------
Global Const AMERICAN = 0
Global Const ANSI = 1
Global Const BRITISH = 2
Global Const FRENCH = 3
Global Const GERMAN = 4
Global Const ITALIAN = 5
Global Const SPANISH = 6
Global Const WIN_DEFAULT = 99

' Epoch Default
' -------------
Global Const EPOCH_DEFAULT = 1980

' Data type identifiers for sx_Replace
' ------------------------------------
Global Const R_INTEGER = 1
Global Const R_LONG = 2
Global Const R_DOUBLE = 8
Global Const R_JULIAN = 32
Global Const R_LOGICAL = 128
Global Const R_CHAR = 1024
Global Const R_DATESTR = 1056
Global Const R_MEMO = 3072
Global Const R_BITMAP = 4096
Global Const R_BLOBFILE = 8192
Global Const R_GENERAL = 8195

' sx_QueryTest Results
' --------------------
Global Const OPTIMIZE_NONE = 0
Global Const OPTIMIZE_PART = 1
Global Const OPTIMIZE_FULL = 2

' sx_EvalTest Results
' -------------------
Global Const EVAL_CHARACTER = 1
Global Const EVAL_NUMERIC = 2
Global Const EVAL_LOGICAL = 3
'Global Const EVAL_DATESTRING = 4       ' not implemented, same as EVAL_CHARACTER

'SDE Index(tag) types
'--------------------
Global Const IDX_NONE = 0
Global Const IDX_UNIQUE = 1
Global Const IDX_EMPTY = 2

'SDE Errorlevel Constants
'------------------------
Global Const ERRLEVEL_NONE = 0
Global Const ERRLEVEL_FATAL = 1
Global Const ERRLEVEL_STANDARD = 2

' RYO BOOL Operations for RYOFilterActivate
' -----------------------------------------
Global Const RYOFILTER_NEW = 1
Global Const RYOFILTER_AND = 2
Global Const RYOFILTER_OR = 3
Global Const RYOFILTER_XOR = 4
Global Const RYOFILTER_ANDNOT = 5
Global Const RYOFILTER_ORNOT = 6
Global Const RYOFILTER_XORNOT = 7

' Collation rule type ( ... sx_SetCollationRule )
' -----------------------------------------
Global Const ALPHABETICAL = 0                    ' usual linguistic
Global Const SPELLING = 1                        ' == Duden
Global Const EXPANDING = 2                       ' additonal groups coalltion rule
Global Const MACHINE = 3                         ' simple value ordering
' ... and collation rule order
Global Const DEFAULT_SET = 0                     ' ALPHABETICAL or duden or expanding default


'* ******************************** */
'* sx_SysProp Constants             */
'* ******************************** */
'  Global Task Information

Global Const SDE_SP_GETSOFTSEEK = 1000           'Get the softseek flag
Global Const SDE_SP_SETSOFTSEEK = 1001           'Set the softseek flag
Global Const SDE_SP_GETEXACT = 1002              'Get the extact flag
Global Const SDE_SP_SETEXACT = 1003              'Set the extact flag
Global Const SDE_SP_GETTRANSLATE = 1004          'Get the translate flag
Global Const SDE_SP_GETDELETED = 1006            'Get the deleted flag
Global Const SDE_SP_PUTOBUFFER = 1007            'Write the optimistic buffer on commit
Global Const SDE_SP_GETOBUFFER = 1008            'Get the optimistic buffer flag
Global Const SDE_SP_SETOBUFFER = 1009            'Set the optimistic buffer flag
Global Const SDE_SP_GETSTRINGTYPE = 1010         'Get the stringtype flag
Global Const SDE_SP_SETSTRINGTYPE = 1011         'Set the stringtype flag
Global Const SDE_SP_GETDISABLEAUTO = 1012        'Get the disable auto open flag
Global Const SDE_SP_SETDISABLEAUTO = 1013        'Set the disable auto open flag

Global Const SDE_SP_SETOEMCOLLATE = 1101         ' Set the collation sequence for OEM tables.
Global Const SDE_SP_GETOEMCOLLATE = 1111         ' Get the collation sequence for OEM tables.
Global Const SDE_SP_SETCHRCOLLATE = 1102         ' Set the collation sequence for Win tables.
Global Const SDE_SP_GETCHRCOLLATE = 1122         ' Get the collation sequence for Win tables.
Global Const SDE_SP_SETLGTRCOLLATE = 1103        ' Set the ligatures collation dimmension
Global Const SDE_SP_GETLGTRCOLLATE = 1133        ' Get the ligatures collation dimmension

Global Const SDE_SP_GETDUDENCOLLATE = 1104       ' get the German DUDEN collate flag
Global Const SDE_SP_SETDUDENCOLLATE = 1105       ' set the German DUDEN collate flag
Global Const SDE_SP_GETLIMITCASECONV = 1106      ' limit case conv to A-Z, a-z if TRUE
Global Const SDE_SP_SETLIMITCASECONV = 1107      ' limit case conv to A-Z, a-z if TRUE

Global Const SDE_SP_SETSPECIALCOLLATE = 1108     ' Set the international collation like DUDEN collate flag
Global Const SDE_SP_GETSPECIALCOLLATE = 1109     ' Set the international collation like DUDEN collate flag
Global Const SDE_SP_GETLANGUAGECOLLATE = 1110    ' Get language, according to collation done

'Behavior settings which bridge the differences between 1.40 and 2.00
Global Const SDE_SP_GETADDQUERY = 1300           'Get the AddQueryFlag
Global Const SDE_SP_SETADDQUERY = 1301           'Set the AddQueryFlag
Global Const SDE_SP_GETUSECONDITIONAL = 1302     'Get the bUseConditional flag
Global Const SDE_SP_SETUSECONDITIONAL = 1303     'Get the bUseConditional flag
Global Const SDE_SP_SETWRITEBLOBHDR = 1305       'Set the bWriteHeader flag
Global Const SDE_SP_GETQUERYRELAXFLAG = 1306     'Get flag that dictates rules of query
Global Const SDE_SP_SETQUERYRELAXFLAG = 1307     'Set flag that dictates rules of query

'WorkArea information
Global Const SDE_SP_GETDRIVER = 2000             'Get the active driver
Global Const SDE_SP_SETSTRDEFLEN = 2001          ' Set the default lenght for STR, if 2nd parameter is absent and field lenght zero
Global Const SDE_SP_SETSTRDEFDEC = 2002          ' Set the default decimals for STR, if 3d parameter is absent and field lenght zero
Global Const SDE_SP_SETDEFAPPEND = 2003          ' Set default behavior for ordering ordering for non-unique key like FOX/Clipper
Global Const SDE_SP_SETDBTNSX = 2004             ' MED: 03-23-01, added NSX_DBT support // differs from SDE7.h
Global Const SDE_SP_BDESPECIFIC = 2005           ' Set way for the treatment of LIKE operator in accoring to BDE rules
Global Const SDE_SP_DBASEDATEHEADER = 2006       ' Set way for the using of DBF header in according to DbaseIII+ specification
Global Const SDE_SP_SETAUTOPAD = 2007            ' Turn on auto padding for resulting character key during indexing
Global Const SDE_SP_GETAUTOPAD = 2008            ' Get flag state for auto padding
Global Const SDE_SP_SETNOVELLFIX = 2009
Global Const SDE_SP_SETORIGINDATE = 2010

'Index information for current workarea
Global Const SDE_SP_GETINDEXCOUNT = 3000         ' Get the number of indexes
Global Const SDE_SP_GETDESCENDING = 3002         ' Get the descending flag
Global Const SDE_SP_GETEMPTY = 3004              ' Get the empty index flag

' -----------------------------------
' SDE FUNCTIONS AND PROCEDURES
' -----------------------------------
Declare Sub sx_AddDudenCollation Lib "SDE7.dll" ()
Declare Sub sx_AddEtecCollation Lib "SDE7.dll" ()
Declare Function sx_Alias Lib "SDE7.dll" (ByVal uiWorkArea As Integer) As String
Declare Sub sx_Append Lib "SDE7.dll" ()
Declare Function sx_AppendEx Lib "SDE7.dll" () As Integer
Declare Sub sx_AppendBlank Lib "SDE7.dll" ()
Declare Function sx_AppendBlankEx Lib "SDE7.dll" () As Integer
Declare Function sx_AppendFrom Lib "SDE7.dll" (ByVal cpFilename As String, ByVal iSourceType As Integer, ByVal cpScopeExpr As Any) As Long
Declare Function sx_BaseDate Lib "SDE7.dll" () As String
Declare Function sx_BaseName Lib "SDE7.dll" () As String
Declare Function sx_BlobToFile Lib "SDE7.dll" (ByVal cpFieldName As String, ByVal cpFilename As String) As Long
Declare Function sx_Bof Lib "SDE7.dll" () As Long
Declare Sub sx_Close Lib "SDE7.dll" ()
Declare Sub sx_CloseAll Lib "SDE7.dll" ()
Declare Sub sx_CloseIndexes Lib "SDE7.dll" ()
Declare Sub sx_Commit Lib "SDE7.dll" ()
Declare Function sx_CommitLevel Lib "SDE7.dll" (ByVal nNewLevel As Long) As Long
Declare Function sx_CopyFile Lib "SDE7.dll" (ByVal cpToFileName As String) As Long
Declare Function sx_CopyFileText Lib "SDE7.dll" (ByVal cpTextFileName As String, ByVal iFileType As Integer) As Long
Declare Function sx_CopyStructure Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String) As Long
Declare Function sx_CopyStructureExtended Lib "SDE7.dll" (ByVal cpFilename As String) As Long
Declare Function sx_Count Lib "SDE7.dll" () As Long
Declare Function sx_CreateExec Lib "SDE7.dll" () As Long
Declare Sub sx_CreateField Lib "SDE7.dll" (ByVal cpName As String, ByVal cpType As String, ByVal iLength As Integer, ByVal iDecimals As Integer)
Declare Function sx_CreateFrom Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String, ByVal cpStruFile As String, ByVal iRdeType As Integer) As Long
Declare Function sx_CreateNew Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String, ByVal iRdeType As Integer, ByVal iNumFields As Integer) As Integer
Declare Function sx_CreateNewEx Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String, ByVal iRdeType As Integer, ByVal iNumFields As Integer, ByVal nModeFlag As Long) As Integer
Declare Function sx_DbfDecrypt Lib "SDE7.dll" () As Long
Declare Function sx_DbfEncrypt Lib "SDE7.dll" () As Long
Declare Function sx_DBFilter Lib "SDE7.dll" () As String
Declare Sub sx_DBRlockList Lib "SDE7.dll" (vpVar As Any)
Declare Function sx_Decrypt Lib "SDE7.dll" (ByVal cpBuffer As Any, ByVal cpPassword As String, ByVal iBuffLen As Long) As String
Declare Sub sx_Delete Lib "SDE7.dll" ()
Declare Function sx_Deleted Lib "SDE7.dll" () As Long
Declare Function sx_Descend Lib "SDE7.dll" (ByVal cpKeyString As String) As String
Declare Function sx_Empty Lib "SDE7.dll" (ByVal cpFieldName As String) As Long
Declare Function sx_Encrypt Lib "SDE7.dll" (ByVal cpBuffer As Any, ByVal cpPassword As String, ByVal iBuffLen As Long) As String
Declare Function sx_Eof Lib "SDE7.dll" () As Long
Declare Function sx_ErrorLevel Lib "SDE7.dll" (ByVal iErrorLevel As Integer) As Integer
Declare Function sx_ErrorTest Lib "SDE7.dll" () As Long
Declare Function sx_EvalLogical Lib "SDE7.dll" (ByVal cpExpression As String) As Long
Declare Function sx_EvalNumeric Lib "SDE7.dll" (ByVal cpExpression As String) As Double
Declare Function sx_EvalString Lib "SDE7.dll" (ByVal cpExpression As String) As String
Declare Function sx_EvalTest Lib "SDE7.dll" (ByVal cpExpression As String) As Integer
Declare Function sx_FieldCount Lib "SDE7.dll" () As Integer
Declare Function sx_FieldDecimals Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_FieldName Lib "SDE7.dll" (ByVal uiFieldNum As Integer) As String
Declare Function sx_FieldNum Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_FieldOffset Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_FieldType Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Function sx_FieldWidth Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_FilterAlias Lib "SDE7.dll" (ByVal cpAliasName As String, ByVal cpFieldName As String) As Long
Declare Function sx_FilterDlg Lib "SDE7.dll" (ByVal hWnd As Long, ByVal cpExpr As Any, ByVal cpCaption As Any, ByVal iHasIndexList As Integer) As String
Declare Sub sx_FinalizeSession Lib "SDE7.dll" ()
Declare Function sx_Flock Lib "SDE7.dll" () As Long
Declare Function sx_Found Lib "SDE7.dll" () As Long
Declare Function sx_GetBitMap Lib "SDE7.dll" (ByVal cpFieldName As String, ByVal hWnd As Long) As Long
Declare Function sx_GetBlob Lib "SDE7.dll" (ByVal cpFieldName As String, vpVar As Any) As Long
Declare Function sx_GetBlobLength Lib "SDE7.dll" (ByVal cpFieldName As String) As Long
Declare Function sx_GetByte Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Function sx_GetCommitLevel Lib "SDE7.dll" (ByVal iWorkareaScope As Integer) As Long
Declare Function sx_GetDateJulian Lib "SDE7.dll" (ByVal cpFieldName As String) As Long
Declare Function sx_GetDateString Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Function sx_GetDateFormat Lib "SDE7.dll" () As Long
Declare Function sx_GetDouble Lib "SDE7.dll" (ByVal cpFieldName As String) As Double
Declare Function sx_GetInteger Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_GetLogical Lib "SDE7.dll" (ByVal cpFieldName As String) As Integer
Declare Function sx_GetLong Lib "SDE7.dll" (ByVal cpFieldName As String) As Long
Declare Function sx_GetMemo Lib "SDE7.dll" (ByVal cpFieldName As String, ByVal LineWidth As Integer) As String
Declare Function sx_GetQueryBit Lib "SDE7.dll" (ByVal lRecNum As Long) As Long
Declare Sub sx_GetRecord Lib "SDE7.dll" (cpRecord As Any)
Declare Function sx_GetScope Lib "SDE7.dll" (ByVal iWhichScope As Integer) As String
Declare Function sx_GetString Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Function sx_GetSystemCharOrder Lib "SDE7.dll" () As String
Declare Function sx_GetSystemLocale Lib "SDE7.dll" () As String
Declare Function sx_GetTrimString Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Sub sx_GetUDFPath Lib "SDE7.dll" (ByVal cpUDFPath As String, ByVal lMaxBufLen As Long)
Declare Function sx_GetVariant Lib "SDE7.dll" (ByVal cpFieldName As String) As String
Declare Sub sx_Go Lib "SDE7.dll" (ByVal lRecNum As Long)
Declare Sub sx_GoBottom Lib "SDE7.dll" ()
Declare Sub sx_GoTop Lib "SDE7.dll" ()
Declare Function sx_Index Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpExpr As String, ByVal iOption As Integer, ByVal bDescend As Long, ByVal cpCondition As Any) As Integer
Declare Sub sx_IndexClose Lib "SDE7.dll" ()
Declare Function sx_IndexCondition Lib "SDE7.dll" () As String
Declare Function sx_IndexFlip Lib "SDE7.dll" () As Long
Declare Function sx_IndexKey Lib "SDE7.dll" () As String
Declare Function sx_IndexKeyField Lib "SDE7.dll" () As String
Declare Function sx_IndexName Lib "SDE7.dll" (ByVal iIndex As Integer) As String
Declare Function sx_IndexOpen Lib "SDE7.dll" (ByVal cpFilename As String) As Integer
Declare Function sx_IndexOrd Lib "SDE7.dll" () As Integer
Declare Function sx_IndexTag Lib "SDE7.dll" (ByVal cpFilename As Any, ByVal cpTagName As String, ByVal cpExpr As String, ByVal iOption As Integer, ByVal bDescend As Long, ByVal cpCondition As Any) As Integer
Declare Function sx_IndexType Lib "SDE7.dll" () As Integer
Declare Function sx_IsEncrypted Lib "SDE7.dll" (ByVal iFileOrRec As Integer) As Long
Declare Function sx_IsNull Lib "SDE7.dll" (ByVal cpFieldName As String) As Long
Declare Function sx_KeyAdd Lib "SDE7.dll" (ByVal cpTagName As Any) As Long
Declare Function sx_KeyData Lib "SDE7.dll" () As String
Declare Function sx_KeyDrop Lib "SDE7.dll" (ByVal cpTagName As Any) As Long
Declare Function sx_KeyLength Lib "SDE7.dll" () As Integer
Declare Function sx_Locate Lib "SDE7.dll" (ByVal cpExpression As Any, ByVal iDirection As Integer, ByVal bContinue As Long) As Long
Declare Function sx_LockCount Lib "SDE7.dll" () As Integer
Declare Function sx_Locked Lib "SDE7.dll" (ByVal lRecNum As Long) As Long
Declare Function sx_OpenMode Lib "SDE7.dll" () As Integer
Declare Function sx_OrderPosGet Lib "SDE7.dll" () As Integer
Declare Sub sx_OrderPosSet Lib "SDE7.dll" (dPosition As Double)
Declare Function sx_OrderRecNo Lib "SDE7.dll" () As Long
Declare Sub sx_Pack Lib "SDE7.dll" ()
Declare Function sx_PutBlob Lib "SDE7.dll" (ByVal cpFieldName As String, vpVar As Any, ByVal lSize As Long) As Long
Declare Sub sx_PutRecord Lib "SDE7.dll" (cpRecord As Any)
Declare Sub sx_PutVariant Lib "SDE7.dll" (ByVal cpFieldName As String, lpVariant As Variant)
Declare Function sx_Query Lib "SDE7.dll" (ByVal cpExpression As Any) As Long
Declare Function sx_QueryRecCount Lib "SDE7.dll" () As Long
Declare Function sx_QuerySetExact Lib "SDE7.dll" (ByVal bSetQueryExact As Long) As Long
Declare Function sx_QueryTest Lib "SDE7.dll" (ByVal cpExpression As Any) As Integer
Declare Sub sx_Recall Lib "SDE7.dll" ()
Declare Function sx_RecCount Lib "SDE7.dll" () As Long
Declare Function sx_RecNo Lib "SDE7.dll" () As Long
Declare Function sx_RecSize Lib "SDE7.dll" () As Long
Declare Function sx_RecToString Lib "SDE7.dll" (cpRecStruc As Any, ByVal iLength As Integer) As String
Declare Sub sx_Reindex Lib "SDE7.dll" ()
Declare Sub sx_Replace Lib "SDE7.dll" (ByVal cpFieldName As String, ByVal iDataType As Integer, vpData As Any)
Declare Function sx_Rlock Lib "SDE7.dll" (ByVal lRecNum As Long) As Long
Declare Function sx_RYOFilterActivate Lib "SDE7.dll" (ByVal iFilterHandle As Integer, ByVal iBoolOperation As Integer) As Long
Declare Function sx_RYOFilterCopy Lib "SDE7.dll" () As Integer
Declare Function sx_RYOFilterCreate Lib "SDE7.dll" () As Integer
Declare Function sx_RYOFilterDestroy Lib "SDE7.dll" (ByVal iFilterHandle As Integer) As Long
Declare Function sx_RYOFilterGetBit Lib "SDE7.dll" (ByVal iFilterHandle As Integer, ByVal lRecNo As Long) As Long
Declare Function sx_RYOFilterRestore Lib "SDE7.dll" (ByVal cpFilename As String) As Long
Declare Function sx_RYOFilterSave Lib "SDE7.dll" (ByVal iFilterHandle As Integer, ByVal cpFilename As String) As Long
Declare Function sx_RYOFilterSetBit Lib "SDE7.dll" (ByVal iFilterHandle As Integer, ByVal lRecNo As Long, ByVal iOnOrOff As Integer) As Long
Declare Function sx_RYOKeyAdd Lib "SDE7.dll" (ByVal cpTagName As String, ByVal cpKey As String) As Long
Declare Function sx_RYOKeyDrop Lib "SDE7.dll" (ByVal cpTagName As String) As Long
Declare Function sx_Seek Lib "SDE7.dll" (ByVal cpKeyValue As String) As Long
Declare Function sx_SeekBin Lib "SDE7.dll" (ByVal cpKeyValue As String, ByVal iLength As Integer) As Long
Declare Function sx_Select Lib "SDE7.dll" (ByVal uiBaseArea As Integer) As Integer
Declare Sub sx_SetCentury Lib "SDE7.dll" (ByVal iValue As Integer)
Declare Function sx_SetCollationRule Lib "SDE7.dll" (ByVal lRuleType As Long, ByVal cpSrcSymbSet As String, ByVal cpDstSymbSet As String, ByVal bResetPrev As Long, ByVal bOem As Long, ByVal lReserved As Long) As Long
Declare Sub sx_SetDateFormat Lib "SDE7.dll" (ByVal uiDateType As Integer)
Declare Sub sx_SetDeleted Lib "SDE7.dll" (ByVal uiDeleted As Integer)
Declare Function sx_SetEpoch Lib "SDE7.dll" (ByVal iBaseYear As Integer) As Integer
Declare Function sx_SetErrorFunc Lib "SDE7.dll" (pFunc As Any, pData As Any) As Long
Declare Sub sx_SetExact Lib "SDE7.dll" (ByVal uiOnOff As Integer)
Declare Sub sx_SetFilter Lib "SDE7.dll" (ByVal cpExpression As Any)
Declare Sub sx_SetGaugeHook Lib "SDE7.dll" (ByVal hwndGauge As Long)
Declare Sub sx_SetLockTimeout Lib "SDE7.dll" (ByVal iSeconds As Integer)
Declare Sub sx_SetMachineCollation Lib "SDE7.dll" () 'simple char value ordering
Declare Function sx_SetMemoBlockSize Lib "SDE7.dll" (ByVal iBlockSize As Integer) As Integer
Declare Function sx_SetOrder Lib "SDE7.dll" (ByVal iIndex As Integer) As Integer
Declare Sub sx_SetPassword Lib "SDE7.dll" (ByVal cpEncodeKey As Any)
Declare Sub sx_SetQueryBit Lib "SDE7.dll" (ByVal lRecNum As Long, ByVal bValue As Long)
Declare Sub sx_SetRelation Lib "SDE7.dll" (ByVal uiChildArea As Integer, ByVal cpKeyExpr As String)
Declare Function sx_SetScope Lib "SDE7.dll" (ByVal cpLowVal As Any, ByVal cpHighVal As Any) As Long
Declare Sub sx_SetSoftSeek Lib "SDE7.dll" (ByVal uiOnOff As Integer)
Declare Sub sx_SetStringType Lib "SDE7.dll" (ByVal uiStringType As Integer)
Declare Sub sx_SetSystemCollation Lib "SDE7.dll" () ' default system linguistic collation
Declare Sub sx_SetTranslate Lib "SDE7.dll" (ByVal uiOnOff As Integer)
Declare Sub sx_SetTurboRead Lib "SDE7.dll" (ByVal uiOnOff As Integer)
Declare Function sx_SetUDFPath Lib "SDE7.dll" (ByVal cpTagName As String)
Declare Function sx_SetUserDelimiter Lib "SDE7.dll" (ByVal cDelimiter As Byte, bMakeTrim As Long, bReflectMemo As Long) As Byte
Declare Sub sx_Skip Lib "SDE7.dll" (ByVal lNumRecs As Long)
Declare Function sx_SysProp Lib "SDE7.dll" (ByVal uiOnOff As Integer, vpData As Any) As Long
Declare Function sx_TagArea Lib "SDE7.dll" (ByVal cpTagName As Any) As Integer
Declare Function sx_TagDelete Lib "SDE7.dll" (ByVal cpTagName As String) As Long
Declare Function sx_TagName Lib "SDE7.dll" (ByVal iIndex As Integer) As String
Declare Sub sx_Unlock Lib "SDE7.dll" (ByVal lRecNum As Long)
Declare Function sx_Use Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String, ByVal iOpenMode As Integer, ByVal iRdeType As Integer) As Integer
Declare Function sx_UseEx Lib "SDE7.dll" (ByVal cpFilename As String, ByVal cpAlias As String, ByVal iOpenMode As Integer, ByVal iRdeType As Integer, ByVal nModeFlag As Long) As Integer
Declare Function sx_Version Lib "SDE7.dll" () As String
Declare Function sx_WorkArea Lib "SDE7.dll" (ByVal cpAlias As Any) As Integer
Declare Sub sx_Zap Lib "SDE7.dll" ()

'===============================================================================================

' ---------------------------------------------------------------------
' Obsolete Declarations **DO NOT USE**
' ---------------------------------------------------------------------

'Declare Sub sx_SetErrorHook Lib "SDE7.dll" (ByVal uiErrorHook As Integer)
'Declare Function sx_SetGaugeFunc LIB "SDE7.dll" (BYVAL pGaugeFunc As Integer) As Integer
'Declare Function sx_SetHandles Lib "SDE7.dll" (ByVal iNumHandles As Integer) As Integer


' Outdated SDE RDE types
' ----------------
'Global Const DBFNTX = 1
'Global Const DBFIDX = 2
'Global Const DBFNSX = 3
'Global Const SIXNTX = 1
'Global Const SIXFOX = 2
'Global Const SIXNSX = 3
