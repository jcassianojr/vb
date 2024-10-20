Attribute VB_Name = "Mod_Functions"
Option Explicit

' Declare Public Type
Public Type RECT
    Left              As Long
    Top               As Long
    Right             As Long
    Bottom            As Long
End Type

Public Type POINTAPI
    X                 As Long
    Y                 As Long
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

' Declare API functions
#If VBA7 Then
Public Declare PtrSafe  Function CreatePen Lib "gdi32" (ByVal nPenStyle As LongPtr, ByVal nWidth As LongPtr, ByVal crColor As LongPtr) As Long
Public Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Public Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As Long
Public Declare PtrSafe  Function LineTo Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr) As Long
Public Declare PtrSafe  Function MoveToEx Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, lpPoint As POINTAPI) As Long
Public Declare PtrSafe Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As LongPtr, ByVal lpStr As String, ByVal nCount As LongPtr, lpRect As RECT, ByVal wFormat As LongPtr) As Long
Public Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr) As Long
Public Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As Long
Public Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal hSrcDC As LongPtr, ByVal xSrc As LongPtr, ByVal ySrc As LongPtr, ByVal nSrcWidth As LongPtr, ByVal nSrcHeight As LongPtr, ByVal dwRop As LongPtr) As Long
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

#Else
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If


'***********************************************************************
' created date: 2004-10-30 17:15
' Purpose: How to draw caption using API (using drawtext method)
'***********************************************************************
Public Sub fDrawText(hDC As Long, StrCaption As String, lLeft As Long, _
        lTop As Long, lRight As Long, lBottom As Long)
        
'=Variable Declaration
Dim lLength As Long
Dim Rc As RECT

    ' set rect info
    With Rc
        .Left = lLeft
        .Top = lTop
        .Right = lRight
        .Bottom = lBottom
    End With
    
    ' get length of string
    lLength = Len(StrCaption)
    
    '= Draw Caption (wordwrap)
    DrawText hDC, StrCaption, lLength, Rc, (&H0 Or &H10)

End Sub

'*************************************************************************
' created date: 2004-10-30 17:17
' Purpose: How to create line using API
'*************************************************************************
Public Sub fDrawLine(DestDC As Long, X As Single, _
    Y As Single, X1 As Single, Y1 As Single, _
    oColor As OLE_COLOR, Optional iWidth As Long = 1)

  Dim pt    As POINTAPI
  Dim iPen  As Long
  Dim iPen1 As Long

    iPen = CreatePen(0, iWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    MoveToEx DestDC, X, Y, pt
    LineTo DestDC, X1, Y1
    SelectObject DestDC, iPen1
    DeleteObject iPen

End Sub

'*******************************************************************************
' created date: 2004-10-30 17:22
' Purpose: How to draw picture using API?
'*******************************************************************************
Public Sub fTransBlt(ByVal dstDC As Long, ByVal DstX As Long, _
        ByVal DstY As Long, ByVal SrcPic As StdPicture, _
        Optional DstW As Long, Optional DstH As Long)
        
    ' local variable declaration
    Dim OriW As Long, OriH As Long
    Dim SrcDC As Long, SrcBmp As Long, SrcObj As Long

    ' get original sizes of picture
    OriW = SrcPic.Width
    OriH = SrcPic.Height
    
    ' check destination sizes
    If DstW = 0 Then DstW = SrcPic.Width
    If DstH = 0 Then DstH = SrcPic.Height
    
    
    SrcDC = CreateCompatibleDC(dstDC)
    SrcBmp = CreateCompatibleBitmap(dstDC, DstW, DstH)
    SrcObj = SelectObject(SrcDC, SrcPic)

    StretchBlt dstDC, DstX, DstY, DstW, DstH, SrcDC, 0, 0, OriW, OriH, vbSrcCopy
    
    ' clean up
    DeleteObject SrcBmp
    DeleteDC SrcDC

End Sub

'******************************************************************************
' created date: 2004-10-30 17:38
' Purpose: Draw welcome page
'******************************************************************************
Public Sub fDrawWelcome(iPic As PictureBox, OptObject)
On Error Resume Next
    With iPic
        ' pre process
        .AutoRedraw = True
        .Cls
        .Height = 313
        .ScaleMode = 3 ' vbpixel
        
        ' process below:
        
       
        ' Draw Line
        fDrawLine CLng(.hDC), 0, .Height - 1, .Width, .Height - 1, RGB(154, 154, 156)
        
        ' Set Font Properties
        .FontBold = True
        .FontSize = 13
        .ForeColor = vbBlack
        
        ' draw Title
        fDrawText CLng(.hDC), LoadResString(102), 175, 15, 400, 90
        
        ' back to normal font
        .FontSize = 8
        .FontBold = False
        
        ' draw caption for description about this wizard
        fDrawText CLng(.hDC), LoadResString(103), 175, 80, 470, 500
        fDrawText CLng(.hDC), LoadResString(104), 175, 130, 460, 220
        fDrawText CLng(.hDC), LoadResString(105), 175, 230, 470, 246
        
        ' display scenario option
        ' 1. Move to MySQL server directly
        OptObject(0).Caption = LoadResString(106)
        OptObject(0).Left = 176
        OptObject(0).Top = 250
        OptObject(0).Width = 250
        
        ' 2. Store into dump file
        Load OptObject(1)
        OptObject(1).Caption = LoadResString(107)
        OptObject(1).Top = 266
        OptObject(1).Visible = True
        
        ' 3. Both
        Load OptObject(2)
        OptObject(2).Caption = LoadResString(108)
        OptObject(2).Top = 282
        OptObject(2).Visible = True
       
        ' display default scenario
        OptObject(0).Value = True ' Move to MySQL server directly
        
        ' Drawline
        fDrawLine CLng(.hDC), 0, 311, .Width, 311, RGB(145, 154, 156)
        ' post process
        .Refresh
        .AutoRedraw = False
    End With
End Sub


Public Sub fDrawHeaderPage(iPic As PictureBox, Index As Integer)
On Error Resume Next

    With iPic
        .AutoRedraw = True
        .Cls
        .Height = 60
        
        ' draw image
        If Index = 1 Then ' MySQL logo
            fTransBlt CLng(.hDC), 420, 10, LoadResPicture(102, 0)
        Else
            fTransBlt CLng(.hDC), 440, 10, LoadResPicture(101 + Index, 0)
        End If
        ' draw line
        fDrawLine CLng(.hDC), 0, .Height - 1, .Width, .Height - 1, RGB(146, 154, 158)
        
        ' Draw title
        .FontBold = True
        
        ' draw title page
        fDrawText CLng(.hDC), LoadResString(109 + Index), 22, 12, 400, 25
        
        '= Draw Description
        .FontBold = False
        If 114 + Index <> 118 Then
           fDrawText CLng(.hDC), LoadResString(114 + Index), 32, 28, 500, 60
        End If
        .Refresh
        .AutoRedraw = False
        
    End With
    
End Sub

Public Sub DrawSeparator(frm_Convert As Form)
'***********************************************
' you can change the separator title
Const MyCompany As String = "Advance Access to MySQL wizard"
'***********************************************
' local variable
Dim lRight As Long

    With frm_Convert
           .AutoRedraw = True
           .Cls
           .ForeColor = vbWhite
           .FontBold = False
           .FontSize = 8
           lRight = .TextWidth(MyCompany) + 19
           ' Draw caption on line separator
           fDrawText CLng(.hDC), MyCompany, 17, 297, lRight, 309
           .ForeColor = vbHighlight
           fDrawText CLng(.hDC), MyCompany, 16, 296, lRight, 309
           ' draw line separator
           fDrawLine CLng(.hDC), CSng(lRight), 305, .ScaleWidth - 16, 305, vbWhite
           fDrawLine CLng(.hDC), CSng(lRight), 304, .ScaleWidth - 15, 304, RGB(154, 156, 158)
           .Refresh
           .AutoRedraw = False
    End With
    
End Sub













