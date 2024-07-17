VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIniSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const defPath As String = "arquivo.ini"


Private m_sPath As String
Private m_sKey As String
Private m_sSection As String
Private m_sDefault As String
Private m_lLastReturnCode As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Sub Class_Initialize()

  m_sPath = defPath

End Sub

Private Function CLngDefault(ByVal sString As String, Optional ByVal lDefault As Long = 0) As Long
  Dim lr As Long
  On Error Resume Next
  lr = CLng(sString)
  If (Err.Number <> 0) Then
    CLngDefault = lDefault
  Else
    CLngDefault = lr
  End If
End Function

Private Property Let Default(sDefault As String)
  m_sDefault = sDefault
End Property

Private Property Get Default() As String
  Default = m_sDefault
End Property

Private Sub DeleteKey()
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, 0&, m_sPath)
End Sub

Private Sub DeleteSection()
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, 0&, m_sPath)
End Sub

Public Function DeleteSetting(Section$, Optional Key$ = "") As Boolean



  m_sSection = Section$
  If Len(Key$) = 0 Then
    DeleteSection
  Else
    m_sKey = Key$
    DeleteKey
  End If

  If Success Then
    DeleteSetting = True
  End If



End Function

Private Sub EnumerateAllSections(ByRef sSections() As String, ByRef iCount As Long)
  Dim sIniFile As String
  Dim iPos As Long
  Dim iNextPos As Long
  Dim sCur As String

  iCount = 0
  Erase sSections
  sIniFile = Sections
  If (Len(sIniFile) > 0) Then
    iPos = 1
    iNextPos = InStr(iPos, sIniFile, Chr$(0))
    Do While iNextPos <> 0
      If (iNextPos <> iPos) Then
        sCur = Mid$(sIniFile, iPos, (iNextPos - iPos))
        iCount = iCount + 1
        ReDim Preserve sSections(1 To iCount) As String
        sSections(iCount) = sCur
      End If
      iPos = iNextPos + 1
      iNextPos = InStr(iPos, sIniFile, Chr$(0))
    Loop
  End If

End Sub

Private Sub EnumerateCurrentSection(ByRef sKey() As String, ByRef iCount As Long)
  Dim sSection As String
  Dim iPos As Long
  Dim iNextPos As Long
  Dim sCur As String

  iCount = 0
  Erase sKey
  sSection = INISection
  If (Len(sSection) > 0) Then
    iPos = 1
    iNextPos = InStr(iPos, sSection, Chr$(0))
    Do While iNextPos <> 0
      sCur = Mid$(sSection, iPos, (iNextPos - iPos))
      If (sCur <> Chr$(0)) Then
        iCount = iCount + 1
        ReDim Preserve sKey(1 To iCount) As String
        sKey(iCount) = Mid$(sSection, iPos, (iNextPos - iPos))
        iPos = iNextPos + 1
        iNextPos = InStr(iPos, sSection, Chr$(0))
      End If
    Loop
  End If
End Sub

Public Function GetAllSections() As Variant

  Dim Sections() As String
  Dim SectionCount As Long
  Dim i As Long
  Dim ArrOut()

  EnumerateAllSections Sections(), SectionCount
  If SectionCount > 0 Then
    ReDim Preserve ArrOut(SectionCount - 1)
    For i = 1 To SectionCount
      ArrOut(i - 1) = Sections(i)
    Next i
    GetAllSections = ArrOut
  End If



End Function

Public Function GetAllSettings(Section$) As Variant

  Dim Keys() As String
  Dim KeyCount As Long
  Dim iKey As Long
  Dim ArrOut()




  m_sSection = Section$
  EnumerateCurrentSection Keys(), KeyCount
  If KeyCount > 0 Then
    ReDim Preserve ArrOut(KeyCount - 1, 1)
    For iKey = 1 To KeyCount
      m_sKey = Keys(iKey)


      ArrOut(iKey - 1, 0) = m_sKey
      ArrOut(iKey - 1, 1) = Value

    Next iKey
    GetAllSettings = ArrOut
  End If


End Function

Public Function GetSetting(Section$, Key$, Optional Default$ = "") As String

  Dim A$



  m_sSection = Section$
  m_sKey = Key$
  m_sDefault = Default$

  A$ = Value

  If Not (Success) Then
    A$ = Default$
  End If



  GetSetting = A$

End Function

Private Property Get INISection() As String
  Dim sBuf As String
  Dim iSize As String
  Dim iRetCode As Integer

  sBuf = Space$(8192)
  iSize = Len(sBuf)
  iRetCode = GetPrivateProfileString(m_sSection, 0&, m_sDefault, sBuf, iSize, m_sPath)
  If (iSize > 0) Then
    INISection = Left$(sBuf, iRetCode)
  Else
    INISection = ""
  End If

End Property

Private Property Let INISection(sSection As String)
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, 0&, sSection, m_sPath)
End Property

Private Property Let Key(sKey As String)
  m_sKey = sKey
End Property

Private Property Get Key() As String
  Key = m_sKey
End Property

Private Property Get LastReturnCode() As Long
  LastReturnCode = m_lLastReturnCode
End Property

Property Let Path(sPath As String)
  m_sPath = sPath
End Property

Property Get Path() As String
  Path = m_sPath
End Property

Public Function SaveSetting(Section$, Key$, Setting$) As Boolean



  m_sSection = Section$
  m_sKey = Key$
  Value = Setting
  If Success Then
    SaveSetting = True
  End If



End Function

Private Property Let Section(sSection As String)
  m_sSection = sSection
End Property

Private Property Get Section() As String
  Section = m_sSection
End Property

Private Property Get Sections() As String
  Dim sBuf As String
  Dim iSize As String
  Dim iRetCode As Integer

  sBuf = Space$(8192)
  iSize = Len(sBuf)
  iRetCode = GetPrivateProfileString(0&, 0&, m_sDefault, sBuf, iSize, m_sPath)
  If (iSize > 0) Then
    Sections = Left$(sBuf, iRetCode)
  Else
    Sections = ""
  End If

End Property

Private Property Get Success() As Boolean
  Success = (m_lLastReturnCode <> 0)
End Property

Private Property Get Value() As String

  Dim sBuf As String
  Dim iSize As String


  sBuf = Space$(255)
  iSize = Len(sBuf)
  m_lLastReturnCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)
  If (iSize > 0) Then
    Value = Left$(sBuf, m_lLastReturnCode)
  Else
    Value = ""
  End If

End Property

Private Property Let Value(sValue As String)

  Dim iPos As Integer


  iPos = InStr(sValue, Chr$(0))
  Do While iPos <> 0
    sValue = Left$(sValue, (iPos - 1)) & Mid$(sValue, (iPos + 1))
    iPos = InStr(sValue, Chr$(0))
  Loop
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
End Property

