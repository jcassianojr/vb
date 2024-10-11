Attribute VB_Name = "Module1"
Option Explicit
Public lRETU As Boolean

Public Function FixStr(ByVal eVAR As Variant, _
                       Optional ByVal ePAD As Variant = "", _
                       Optional ByVal coper As String = "", _
                       Optional ByVal nLEN As Integer = 0) As Variant
  On Error GoTo errhandler
  If IsNull(eVAR) Then
    If ePAD <> "" Then
      eVAR = ePAD
    End If
  End If
  If Not IsNull(eVAR) Then
     FixStr = CStr(eVAR)
  End If
  If UCase(eVAR) = "NULL" Then
    FixStr = ePAD
  End If
  If InStr(coper, "TRIM") > 0 Then
    eVAR = Trim(eVAR)
    FixStr = eVAR
  End If
  If nLEN > 0 And Len(eVAR) > nLEN Then
    eVAR = Mid(eVAR, 1, nLEN)
    FixStr = eVAR
  End If
  Exit Function
errhandler:
  FixStr = ""
  Resume Next
End Function
Public Function TrimNull(ByVal sTxt As String) As String
  Dim Arr() As String
  Arr() = Split(sTxt, Chr$(0))
  If UBound(Arr) >= 0 Then
    TrimNull = Arr(0)
  Else
    TrimNull = sTxt
  End If
End Function
Public Function FileExists(ByVal PathName As String) As Boolean
On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
End Function

