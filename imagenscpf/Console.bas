Attribute VB_Name = "basConsole"
Option Explicit

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
Declare  PtrSafe Function AllocConsole Lib "kernel32" () As Long
Declare  PtrSafe Function FreeConsole Lib "kernel32" () As Long
Declare  PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Declare  PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As LongPtr) As Long
Declare  PtrSafe Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As LongPtr, lpBuffer As Any, ByVal nNumberOfCharsToWrite As LongPtr, lpNumberOfCharsWritten As LongPtr, lpReserved As Any) As Long
Declare  PtrSafe Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As LongPtr, ByVal Add As LongPtr) As Long
Public Declare  PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias _
                                                "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                            ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                                            ByVal lpReturnedString As String, ByVal nSize As LongPtr, _
                                                                            ByVal lpFileName As String) As Long
#Else
Declare Function AllocConsole Lib "kernel32" () As Long
Declare Function FreeConsole Lib "kernel32" () As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, lpBuffer As Any, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Declare Function SetConsoleCtrlHandler Lib "kernel32" (ByVal HandlerRoutine As Long, ByVal Add As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                                                "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                            ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                                            ByVal lpReturnedString As String, ByVal nSize As Long, _
                                                                            ByVal lpFileName As String) As Long
#End If

Public Const STD_OUTPUT_HANDLE = -11&
Public hConsole As Long
Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
  Dim z As Long
  Dim scaminho As String * 255
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, App.Path + "\" & App.EXEName & ".INI")
  If Len(ePAD) = 0 Then
    ePAD = App.Path & "\"
  End If
  PegPath = IIf(Asc(Left(scaminho, 1)) = "0", ePAD, Left(scaminho, z))

End Function
Public Function ConsoleHandler(ByVal CtrlType As Long) As Long

  ConsoleHandler = 1  'Keep track of the console, it'll survive even Ctrl+Break

End Function
