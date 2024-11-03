Attribute VB_Name = "modIniFiles"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_filename As String
Private Const MAX_SIZE = 2048

Enum TCompare
    BinCompare = 0
    TxtCompare = 1
End Enum

Public Property Get Filename() As Variant
    Filename = m_filename
End Property

Public Property Let Filename(ByVal vNewValue As Variant)
    m_filename = vNewValue
End Property

Public Function ReadString(Selection As String, Key As String, Default As String) As String
Dim RetVal As Long
Dim buffer As String
    'Create buffer to hold the string
    buffer = String$(MAX_SIZE, Chr$(0))
    'Read the string from the ini file
    RetVal = GetPrivateProfileString(Selection, Key, Default, _
    buffer, MAX_SIZE, Filename)
    'Extract the string from the buffer
    ReadString = Left$(buffer, RetVal)
End Function

Public Function ReadInteger(Selection As String, Key As String, Default As Integer) As Integer
    ReadInteger = CInt(ReadString(Selection, Key, CInt(Default)))
End Function

Public Function ReadFloat(Selection As String, Key As String, Default As Double) As Double
    ReadFloat = CDbl(ReadString(Selection, Key, CDbl(Default)))
End Function

Public Function ReadBoolean(Selection As String, Key As String, Default As Boolean) As Boolean
    ReadBoolean = CBool(ReadString(Selection, Key, CBool(Default)))
End Function

Public Function WriteString(Selection As String, Key As String, Value As String) As Long
    'Write value to ini file.
    WriteString = WritePrivateProfileString(Selection, Key, Value, Filename)
End Function

Public Function WriteInteger(Selection As String, Key As String, Value As Integer) As Long
    'Write integer to ini file.
    WriteInteger = WriteString(Selection, Key, CInt(Value))
End Function

Public Function WriteFloat(Selection As String, Key As String, Value As Double) As Long
    'Write float to ini file.
    WriteFloat = WriteString(Selection, Key, CDbl(Value))
End Function

Public Function WriteBoolean(Selection As String, Key As String, Value As Boolean) As Long
    'Write bool to ini file.
    WriteBoolean = WriteString(Selection, Key, CBool(Value))
End Function

Public Function DeleteSelection(Selection As String) As Long
    'Delete a selection from ini file.
    DeleteSelection = WritePrivateProfileString(Selection, ByVal 0&, ByVal 0&, Filename)
End Function

Public Function DeleteKey(Selection As String, Key As String) As Long
    'Delete a key and it's value
    DeleteKey = WritePrivateProfileString(Selection, Key, ByVal 0&, Filename)
End Function

Public Function SelectionFound(Selection As String) As Boolean
Dim c As New Collection
Dim I As Long
Dim Flag As Boolean
    
    Flag = False
    'Read selections
    Set c = ReadSelections()
    
    For I = 1 To c.Count
        'Compare selection names
        If LCase$(c(I)) = LCase$(Selection) Then
            'Selection match found exit loop
            Flag = True
            Exit For
        End If
    Next I
    'Set obj c to nothing
    Set c = Nothing
    SelectionFound = Flag
End Function

Public Function KeyFound(Selection As String, Key As String) As Boolean
Dim c As New Collection
Dim I As Integer
Dim Flag As Boolean

    'Find a key in the ini file
    Flag = False
    
    Set c = ReadSelectionKeys(Selection)
    
    For I = 1 To c.Count
        If LCase$(c(I)) = LCase$(Key) Then
            Flag = True
            Exit For
        End If
    Next I
    
    Set c = Nothing
    KeyFound = Flag
End Function

Public Function ReadSelections() As Collection
Dim fp As Long
Dim c As New Collection
Dim sLine As String

    If FileFound(Filename) Then
        fp = FreeFile
        Open Filename For Input As #fp
            Do While Not EOF(fp)
                'Read input line
                Line Input #fp, sLine
                'Trim the line
                sLine = Trim$(sLine)
     
                If sLine <> "" Then
                    'Check for selection
                    If ((Left$(sLine, 1) = "[") And (Right$(sLine, 1) = "]")) Then
                        'Add selection name to collection
                        Call c.Add(Mid$(sLine, 2, Len(sLine) - 2))
                    End If
                End If
            Loop
        Close #fp
    End If
    'Return collection
    Set ReadSelections = c
End Function

Public Function ValueFound(Selection As String, Key As String, Checkvalue As String _
, ICompare As TCompare) As Boolean
Dim S0 As String
Dim S1 As String
    'Check if a key contains a value
    
    'Get value
    S0 = ReadString(Selection, Key, "")
    S1 = Checkvalue
    'Compare binary
    If (ICompare = BinCompare) Then
        ValueFound = (S0 = S1)
    End If
    'Compare text
    If (ICompare = TxtCompare) Then
        ValueFound = (UCase$(S0) = UCase$(S1))
    End If
    
End Function

Public Function ReadSectionValues(Selection As String) As Collection
Dim c As New Collection
Dim vals As New Collection
Dim I As Integer
    'Read election values
    
    'Read selection keys
    Set c = ReadSelectionKeys(Selection)
    
    For I = 1 To c.Count
       Call vals.Add(ReadString(Selection, c(I), ""))
    Next I
    
    Set c = Nothing
    'Return values
    Set ReadSectionValues = vals
End Function

Public Function ReadSelectionKeys(Selection As String) As Collection
Dim fp As Long
Dim c As New Collection
Dim inSelection As Boolean
Dim sLine As String
Dim sSel As String
Dim ePos As Integer
    
    inSelection = False
    
    'Check if file is found
    If FileFound(Filename) Then
        fp = FreeFile
        Open Filename For Input As #fp
            Do While Not EOF(fp)
                'Read input line
                Line Input #fp, sLine
                'Trim the line
                sLine = Trim$(sLine)
     
                If sLine <> "" Then
                    'Check for selection
                    If ((Left$(sLine, 1) = "[") And (Right$(sLine, 1) = "]")) Then
                        'Get selection name
                        sSel = UCase$(Trim$(Mid$(sLine, 2, Len(sLine) - 2)))
                        If (sSel = UCase$(Selection)) Then
                            'Set found selection to true
                            inSelection = True
                            'Get next line
                            Line Input #fp, sLine
                            'Trim line
                            sLine = Trim$(sLine)
                        Else
                            'Not in selection
                            inSelection = False
                        End If
                    End If
                        'Check if we have a selection match
                        If inSelection Then
                            'Get position of = in sline
                            ePos = InStr(sLine, "=")
                            If (ePos > 0) Then
                                'Extract the key name and add to collection
                                 Call c.Add(Trim$(Left$(sLine, ePos - 1)))
                            End If
                        End If
                End If
            Loop
        Close #fp
    End If
    'Return collection
    Set ReadSelectionKeys = c
End Function

'Private tools
Private Function FileFound(File As String) As Boolean
    If (Dir$(File)) = "" Then
        FileFound = False
    Else
        FileFound = True
    End If
End Function
