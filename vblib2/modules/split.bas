Attribute VB_Name = "Strings"
' Strings.bas
Option Explicit

Public Enum SplitCompareMethod
  [Split BinaryCompare] = VbCompareMethod.vbBinaryCompare         ' InStrB
  '    [Split TextCompare] = VbCompareMethod.vbTextCompare             ' InStr(TextCompare)
  [Split CharacterCompare] = VbCompareMethod.vbDatabaseCompare    ' InStr(BinaryCompare)
End Enum

#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If
#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If

#If VBA7 Then
Private Declare  PtrSafe Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
'Private Declare  PtrSafe Sub GetMem4 Lib "msvbvm60" (ByVal Ptr As LongPtr, Value As LongPtr) 'imcompatibilidade vba7
'Private Declare  PtrSafe Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As LongPtr, ByVal Value As LongPtr) 'imcompatibilidade vba7
'Private Declare  PtrSafe Sub PutMem8 Lib "msvbvm60" (ByVal Ptr As LongPtr, ByVal Value As Currency) 'imcompatibilidade vba7
Private Declare  PtrSafe Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As LongPtr, ByVal Length As LongPtr) As Long
Private Declare  PtrSafe Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As LongPtr, ByVal Length As LongPtr) As Long
Private Declare  PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare  PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare  PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As LongPtr, ByVal bInheritHandle As LongPtr, ByVal dwProcessId As LongPtr) As Long
Private Declare  PtrSafe Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As LongPtr, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As LongPtr, Optional lpNumberOfBytesWritten As LongPtr) As Long
#Else
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Arr() As Any) As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Ptr As Long, Value As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Currency)
Private Declare Function SysAllocStringByteLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, Optional lpNumberOfBytesWritten As Long) As Long
#End If



Private Function InIDE(Optional IDE) As Boolean
  If IsMissing(IDE) Then Debug.Assert Not InIDE(InIDE) Else IDE = True
End Function

Private Property Get Procedure(ByVal AddressOfDest As Long) As Long
  Procedure = AddressOfDest
End Property

Private Property Let Procedure(ByVal AddressOfDest As Long, ByVal AddressOfSrc As Long)
  Dim JMP As Currency, PID As Long
  ' get process handle
  PID = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)
  If PID Then
    If InIDE Then
      ' get correct pointers to procedures in IDE
      GetMem4 AddressOfDest + &H16&, AddressOfDest
      GetMem4 AddressOfSrc + &H16&, AddressOfSrc
    End If
    Debug.Assert App.hInstance
    ' ASM JMP (0xE9) followed by bytes to jump in memory
    JMP = (&HE9& * 0.0001@) + (AddressOfSrc - AddressOfDest - 5@) * 0.0256@
    ' write the JMP over the destination procedure
    WriteProcessMemory PID, ByVal AddressOfDest, JMP, 5
    ' close process handle
    CloseHandle PID
  End If
End Property

Public Function Split(ByVal Expression, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As SplitCompareMethod) As String()
'expression as string deixado so expression pois recebee as vezes matrizes exemplo auso(2)
' general variables that we need
  Dim D() As Integer, E() As Integer, H() As Long, HE() As Long, HS() As Long, P() As Long, R() As String, S() As Long
  Dim C As Long, i As Long, J As Long, K As Long, N As Long, LD As Long, LE As Long, PL As Long, V As Long
  Dim PD As Long, PE As Long, PS As Long
  ' length information
  If Compare = [Split BinaryCompare] Then
    LE = LenB(Expression)
    LD = LenB(Delimiter)
  Else
    LE = Len(Expression)
    LD = Len(Delimiter)
  End If
  ' unlimited or limited?
  If Limit = -1 Then Limit = LE \ LD + 1
  ' validate lengths and limit
  If LE > 0 And LD > 0 And Limit >= 0 Then
    ' find the first item
    If Limit > 1 Then
      Select Case Compare
      Case [Split BinaryCompare]
        Do: i = InStrB(i + 1, Expression, Delimiter)
        Loop Until (i And 1) = 1 Or (i = 0)
      Case Else
        i = InStr(Expression, Delimiter)
      End Select
    End If
    ' did we find an item?
    If i Then
      ' space for knowing the positions
      PL = Limit \ 80
      ReDim P(0 To PL)
      ' InStrB?
      If Compare = [Split BinaryCompare] Then
        Do
          ' remember position
          P(C) = i
          ' find next
          i = i + LD - 1
          Do: i = InStrB(i + 1, Expression, Delimiter)
          Loop Until (i And 1) = 1 Or (i = 0)
          ' increase counter
          C = C + 1
          If C > PL Then PL = PL + C: ReDim Preserve P(PL)
        Loop While i > 0 And C <= Limit
        P(C) = LE + 1
        ' make space for the new items
        ReDim R(0 To C)
        If C < 32 Then
          ' there is not a lot to do so keep it simple!
          i = 1
          For C = 0 To C
            J = P(C) - i
            R(C) = MidB$(Expression, i, J)
            i = P(C) + LD
          Next C
        Else
          ' okay, now there should be enough to work with to have use of safe arrays...
          ReDim H(0 To 5)
          HS = H
          HS(0) = 1: HS(1) = 4: HS(4) = 1
          ' S = current R array item (Long)
          PS = ArrPtr(S)
          PutMem4 PS, VarPtr(HS(0))
          H(0) = 1: H(1) = 2: H(4) = LE
          HE = H
          HE(3) = StrPtr(Expression)
          ' D = current R item's data (Integer array)
          PD = ArrPtr(D)
          PutMem4 PD, VarPtr(H(0))
          ' E = Expression's data (Integer array)
          PE = ArrPtr(E)
          PutMem4 PE, VarPtr(HE(0))
          V = VarPtr(R(0))
          ' then start working...
          i = 1
          For C = 0 To C
            J = P(C) - i
            If J Then
              If J > 191 Then
                ' MidB$ is faster than Integer array with long copy
                R(C) = MidB$(Expression, i, J)
              Else
                ' Integer array is faster with a short copy
                HS(3) = V + C * 4
                S(0) = SysAllocStringByteLen(0, J)
                H(3) = S(0)
                N = (i - 1) \ 2
                For K = 0 To (J - 1) \ 2
                  D(K) = E(N + K)
                Next K
              End If
            End If
            i = P(C) + LD
          Next C
          PutMem4 PS, 0
          PutMem4 PD, 0
          PutMem4 PE, 0
        End If
      Else  ' InStr
        Do
          ' remember position
          P(C) = i
          ' find next
          i = InStr(i + LD, Expression, Delimiter)
          ' increase counter
          C = C + 1
          If C > PL Then PL = PL + C: ReDim Preserve P(PL)
        Loop While i > 0 And C <= Limit
        P(C) = LE + 1
        ' make space for the new items
        ReDim R(0 To C)
        If C < 32 Then
          ' there is not a lot to do so keep it simple!
          i = 1
          For C = 0 To C
            J = P(C) - i
            R(C) = Mid$(Expression, i, J)
            i = P(C) + LD
          Next C
        Else
          ' okay, now there should be enough to work with to have use of safe arrays...
          ReDim H(0 To 5)
          HS = H
          HS(0) = 1: HS(1) = 4: HS(4) = 1
          ' S = current R array item (Long)
          PS = ArrPtr(S)
          PutMem4 PS, VarPtr(HS(0))
          H(0) = 1: H(1) = 2: H(4) = LE
          HE = H
          HE(3) = StrPtr(Expression)
          ' D = current R item's data (Integer array)
          PD = ArrPtr(D)
          PutMem4 PD, VarPtr(H(0))
          ' E = Expression's data (Integer array)
          PE = ArrPtr(E)
          PutMem4 PE, VarPtr(HE(0))
          V = VarPtr(R(0))
          ' then start working...
          i = 1
          For C = 0 To C
            J = P(C) - i
            If J Then
              If J > 191 Then
                ' MidB$ is faster than Integer array with long copy
                R(C) = Mid$(Expression, i, J)
              Else
                ' Integer array is faster with a short copy
                HS(3) = V + C * 4
                S(0) = SysAllocStringLen(0, J)
                H(3) = S(0)
                N = i - 1
                For K = 0 To J
                  D(K) = E(N + K)
                Next K
              End If
            End If
            i = P(C) + LD
          Next C
          PutMem4 PS, 0
          PutMem4 PD, 0
          PutMem4 PE, 0
        End If
      End If
    Else
      ' one item
      ReDim R(0)
      R(0) = Expression
    End If
    ' swap arrays!
    SwapStrArr R, Split
  Else
    ' safe array header for an empty string array
    ReDim H(0 To 6): H(0) = vbString: H(1) = &H1800001: H(2) = 4
    ' H(1) becomes ArrPtr; H(0) is a negative item telling array datatype
    PutMem4 StrArrPtr(Split), VarPtr(H(1))
    ' remove items from H array
    PutMem8 (Not Not H) + 12, 0
    ' VB6 IDE mysterious bug fix with Not Array
    Debug.Assert App.hInstance
  End If
End Function

Public Function StrArrPtr(Arr() As String) As Long
  Procedure(AddressOf Strings.StrArrPtr) = Procedure(AddressOf Strings.z_ArrPtr)
  StrArrPtr = Strings.StrArrPtr(Arr)
End Function

Public Sub SwapStrArr(Arr1() As String, Arr2() As String)
  Procedure(AddressOf Strings.SwapStrArr) = Procedure(AddressOf Strings.z_Swap)
  Strings.SwapStrArr Arr1, Arr2
End Sub

Public Function z_ArrPtr(ByVal Value As Long) As Long
  z_ArrPtr = Value
End Function

Public Sub z_Swap(Swap1 As Long, Swap2 As Long)
  Dim lngSwap As Long
  lngSwap = Swap1
  Swap1 = Swap2
  Swap2 = lngSwap
End Sub
