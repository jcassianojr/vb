Attribute VB_Name = "modCloseBtn"
'usado ao pedir senha desabilita o x na barra do form
'evitando assim que usuario feche a tela de login
'forcando a pedir usuario e senha

Option Explicit

Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Type MENUITEMINFO
  cbSize As Long
  fMask As Long
  fType As Long
  fState As Long
  wID As Long
  hSubMenu As Long
  hbmpChecked As Long
  hbmpUnchecked As Long
  dwItemData As Long
  dwTypeData As String
  cch As Long
End Type

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


#If VBA7 Or Win64 Then
Private Declare  PtrSafe Function GetSystemMenu Lib "user32" ( _
                                       ByVal hWnd As LongPtr, ByVal bRevert As LongPtr) As Long

Private Declare  PtrSafe Function GetMenuItemInfo Lib "user32" Alias _
                                         "GetMenuItemInfoA" (ByVal hMenu As LongPtr, ByVal un As LongPtr, _
                                                             ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare  PtrSafe Function SetMenuItemInfo Lib "user32" Alias _
                                         "SetMenuItemInfoA" (ByVal hMenu As LongPtr, ByVal un As LongPtr, _
                                                             ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare  PtrSafe Function SendMessage Lib "user32" Alias _
                                     "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As LongPtr, _
                                                     ByVal wParam As LongPtr, lParam As Any) As Long

Private Declare  PtrSafe Function IsWindow Lib "user32" _
                                  (ByVal hWnd As LongPtr) As Long


#Else
Private Declare Function GetSystemMenu Lib "user32" ( _
                                       ByVal hwnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
                                         "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
                                                             ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
                                         "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
                                                             ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
                                     "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                                     ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
                                  (ByVal hwnd As Long) As Long
 #End If
                                  
                                  

'*******************************************************************************
' Enables / Disables the close button on the titlebar and in the system menu
' of the form window passed.
'-------------------------------------------------------------------------------
' Return Values:
'
'    0  Close button state changed succesfully / nothing to do.
'   -1  Invalid Window Handle (hWnd argument) Passed to the function
'   -2  Failed to switch command ID of Close menu item in system menu
'   -3  Failed to switch enabled state of Close menu item in system menu
'
'-------------------------------------------------------------------------------
' Parameters:
'
'   hWnd    The window handle of the form whose close button is to be enabled/
'           disabled / greyed out.
'
'   Enable  True if the close button is to be enabled, or False if it is to
'           be disabled / greyed out.
'
'-------------------------------------------------------------------------------
' Example:
'
' Add a form window to your project, and place a button on the form. Add the
' following in the form's code window:
'
'    Option Explicit
'
'    Private m_blnCloseEnabled As Boolean
'
'    Private Sub Form_Load()
'        m_blnCloseEnabled = True
'        Command1.Caption = "Disable"
'    End Sub
'
'    Private Sub Command1_Click()
'        m_blnCloseEnabled = Not m_blnCloseEnabled
'        EnableCloseButton clng(Me.hwnd), m_blnCloseEnabled
'
'        If m_blnCloseEnabled Then
'            Command1.Caption = "Disable"
'        Else
'            Command1.Caption = "Enable"
'        End If
'    End Sub
'
'-------------------------------------------------------------------------------

Public Function EnableCloseButton(ByVal hwnd As Long, Enable As Boolean) _
       As Integer
  Const xSC_CLOSE As Long = -10

  ' Check that the window handle passed is valid

137:   EnableCloseButton = -1
138:   If IsWindow(hwnd) = 0 Then Exit Function

  ' Retrieve a handle to the window's system menu

  Dim hMenu As Long
143:   hMenu = GetSystemMenu(hwnd, 0)

  ' Retrieve the menu item information for the close menu item/button

  Dim MII As MENUITEMINFO
148:   MII.cbSize = Len(MII)
149:   MII.dwTypeData = String(80, 0)
150:   MII.cch = Len(MII.dwTypeData)
151:   MII.fMask = MIIM_STATE

153:   If Enable Then
154:     MII.wID = xSC_CLOSE
155:   Else
156:     MII.wID = SC_CLOSE
157:   End If

159:   EnableCloseButton = -0
160:   If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function

  ' Switch the ID of the menu item so that VB can not undo the action itself

  Dim lngMenuID As Long
165:   lngMenuID = MII.wID

167:   If Enable Then
168:     MII.wID = SC_CLOSE
169:   Else
170:     MII.wID = xSC_CLOSE
171:   End If

173:   MII.fMask = MIIM_ID
174:   EnableCloseButton = -2
175:   If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function

  ' Set the enabled / disabled state of the menu item

179:   If Enable Then
180:     MII.fState = (MII.fState Or MFS_GRAYED)
181:     MII.fState = MII.fState - MFS_GRAYED
182:   Else
183:     MII.fState = (MII.fState Or MFS_GRAYED)
184:   End If

186:   MII.fMask = MIIM_STATE
187:   EnableCloseButton = -3
188:   If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function

  ' Activate the non-client area of the window to update the titlebar, and
  ' draw the close button in its new state.

193:   SendMessage hwnd, WM_NCACTIVATE, True, 0

195:   EnableCloseButton = 0

End Function

'*******************************************************************************
'
'-------------------------------------------------------------------------------
