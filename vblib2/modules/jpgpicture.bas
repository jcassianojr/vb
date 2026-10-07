Attribute VB_Name = "jpgpicture"
Option Explicit

' ==============================================================================
' ENUMS POSICIONADOS NO TOPO ABSOLUTO (Evita erro de compilação)
' ==============================================================================
Public Enum ExifOrientationEnum
  Exif_Orientation_Normal = 1
  Exif_Orientation_MirrorHorizontal = 2
  Exif_Orientation_Rotate180 = 3
  Exif_Orientation_MirrorVertical = 4
  Exif_Orientation_MirrorHorizontalRotate270 = 5
  Exif_Orientation_Rotate90 = 6
  Exif_Orientation_MirrorHorizontalRotate90 = 7
  Exif_Orientation_Rotate270 = 8
End Enum

Public Enum RotationEnum
    Rotate_0 = 0
    Rotate_90 = 90
    Rotate_180 = 180
    Rotate_270 = 270
End Enum

Public Const EXIF_ORIENTATION As Long = 274

' ==============================================================================
' MANUTENÇÃO DA RETROCOMPATIBILIDADE GDI (Para StretchSourcePictureFromPicture)
' ==============================================================================
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
Private Const STRETCH_HALFTONE = 4

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
  Size As Long
  PictureType As Long
  hPic As LongPtr
  hPal As LongPtr
End Type

#If VBA7 Or Win64 Then
    Private Type BITMAP
        BMType As Long: BMWidth As Long: BMHeight As Long: BMWidthBytes As Long
        BMPlanes As Integer: BMBitsPixel As Integer: BMBits As LongPtr
    End Type
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function CopyImage Lib "user32" (ByVal hImage As LongPtr, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal fuFlags As Long) As LongPtr
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As GUID) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef iPic As StdPicture) As Long
    Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Public Declare PtrSafe Function SetStretchBltMode Lib "gdi32" (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
    Public Declare PtrSafe Function StretchBlt Lib "gdi32" (ByVal hdcDest As LongPtr, _
         ByVal xDest As Long, ByVal yDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
         ByVal hdcSrc As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal dwRop As Long) As Long
#Else
    Private Type BITMAP
        BMType As Long: BMWidth As Long: BMHeight As Long: BMWidthBytes As Long
        BMPlanes As Integer: BMBitsPixel As Integer: BMBits As Long
    End Type
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function CopyImage Lib "user32" (ByVal hImage As Long, ByVal uType As Long, ByVal CX As Long, ByVal CY As Long, ByVal fuFlags As Long) As Long
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As GUID) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (ByRef PicDesc As PICTDESC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef iPic As StdPicture) As Long
    Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
    Public Declare Function StretchBlt Lib "gdi32" (ByVal hdcDest As Long, _
         ByVal xDest As Long, ByVal yDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
         ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal dwRop As Long) As Long
#End If

Public Function PictureWithOwnedBitmap(ByVal sourcePicture As StdPicture) As StdPicture
  Const IMAGE_BITMAP As Long = 0
  Const LR_CREATEDIBSECTION As Long = &H2000
  Const PICTYPE_BITMAP As Long = 1
  Dim pictureIID As GUID
  Dim pictureDesc As PICTDESC
  Dim bitmapCopy As LongPtr
  Dim result As Long

  If sourcePicture Is Nothing Then
    Err.Raise 5, "jpgpicture.PictureWithOwnedBitmap", "A source picture is required."
  End If
  If sourcePicture.Handle = 0 Then
    Err.Raise 5, "jpgpicture.PictureWithOwnedBitmap", "The source picture has no bitmap handle."
  End If
  If IIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), pictureIID) <> 0 Then
    Err.Raise vbObjectError + 1000, "jpgpicture.PictureWithOwnedBitmap", "Could not resolve the StdPicture interface."
  End If

  bitmapCopy = CopyImage(sourcePicture.Handle, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION)
  If bitmapCopy = NULL_PTR Then
    Err.Raise vbObjectError + 1001, "jpgpicture.PictureWithOwnedBitmap", "Could not duplicate the image bitmap."
  End If

  With pictureDesc
    .Size = Len(pictureDesc)
    .PictureType = PICTYPE_BITMAP
    .hPic = bitmapCopy
    .hPal = NULL_PTR
  End With

  result = OleCreatePictureIndirect(pictureDesc, pictureIID, 1, PictureWithOwnedBitmap)
  If result <> 0 Then
    Call DeleteObject(bitmapCopy)
    Err.Raise vbObjectError + 1002, "jpgpicture.PictureWithOwnedBitmap", "Could not create an independently owned StdPicture."
  End If
End Function

' ==============================================================================
' 1. MÉTODOS REFATORADOS PARA USO DA CLASSE stdImage (Fim do PicSaveLoad)
' ==============================================================================

Public Function StretchSourcePictureFromFile(ByVal FileName As String, ByRef picDest As PictureBox) As StdPicture
  Dim imgObj As Object
  
  If Len(FileName) = 0 Then
    Beep
    Exit Function
  End If

  Set imgObj = stdImage.CreateFromFile(FileName)
  Set StretchSourcePictureFromFile = PictureWithOwnedBitmap(imgObj.ToStdPicture())

  Call StretchSourcePictureFromPicture(StretchSourcePictureFromFile, picDest)
End Function

Public Function lerarquivoimagem(ByVal STMPFILE As String, ByRef Picture1 As PictureBox, ByRef Picture2 As PictureBox) As Boolean
  lerarquivoimagem = False
  If Len(STMPFILE) > 0 Then
    If FixInt(FileLen(STMPFILE)) > 500000 Then
      Alert ("Imagem Muito Grande,Ajuste o tamanho")
      If Not MDG("Anexar mesmo assim-NAO RECOMENDADO") Then Exit Function
    End If
    
    Dim imgObj As Object
    Set imgObj = stdImage.CreateFromFile(STMPFILE)
    Set Picture1.Picture = PictureWithOwnedBitmap(imgObj.ToStdPicture())
    
    StretchSourcePictureFromPicture Picture1.Picture, Picture2
    lerarquivoimagem = True
  End If
End Function

Public Function salvarpict(oFORM As Form, ByVal Picture1 As Variant, _
                           Optional ByVal sFileName As String = "imagem", _
                           Optional ByVal sPath As String = "")
  Dim sFilter As String, cEXTENSAO As String
  Dim imgObj As Object

  If Len(sPath) = 0 Then sPath = App.Path
  
  sFilter = ImgFILTER2()
  sFileName = FileSave(oFORM, sFilter, 1, , sFileName, sPath, "Salvar Imagem")
  cEXTENSAO = parsefile(sFileName, "E")

  If Len(sFileName) <= 0 Then
    Alert ("Nome do arquivo nao Preenchido")
    Exit Function
  End If
  
  If FileConnExist(sFileName, False) Then
    If MDG("Arquivo de Destino Ja existe Sobrepor") Then
      DeleteFile sFileName, True
    Else
      Exit Function
    End If
  End If

  Set imgObj = stdImage.CreateFromStdPicture(Picture1.Picture)
  
  Select Case UCase(cEXTENSAO)
    Case "JPG", "JPEG": imgObj.ToFile sFileName, stdImgFormatJPEG, 70
    Case "PNG": imgObj.ToFile sFileName, stdImgFormatPNG
    Case "GIF": imgObj.ToFile sFileName, stdImgFormatGIF
    Case "BMP": imgObj.ToFile sFileName, stdImgFormatBMP
    Case Else: imgObj.ToFile sFileName, stdImgFormatDefault
  End Select
End Function

' ==============================================================================
' 2. PROCESSAMENTO TOTAL EM MEMÓRIA (Base64 sem MSXML2)
' ==============================================================================

Public Sub SalvarBase64ComoImagem(ByVal sBase64 As String, ByVal sCaminhoDestino As String)
    Dim imgObj As Object
    
    If InStr(sBase64, "data:image") = 0 Then
        sBase64 = "data:image/png;base64," & sBase64
    End If
    
    Set imgObj = stdImage.CreateFromDataURL(sBase64)
    imgObj.ToFile sCaminhoDestino
End Sub

Public Sub Base64ParaPictureBox(ByVal sBase64 As String, ByRef picDest As PictureBox)
    Dim imgObj As Object
    If InStr(sBase64, "data:image") = 0 Then sBase64 = "data:image/png;base64," & sBase64
    
    Set imgObj = stdImage.CreateFromDataURL(sBase64)
    Set picDest.Picture = PictureWithOwnedBitmap(imgObj.ToStdPicture())
End Sub

' ==============================================================================
' 3. EXPANSÃO DE CAPACIDADES (Pronto para Uso na UI)
' ==============================================================================

Public Sub AreaDeTransferenciaParaPictureBox(ByRef picDest As PictureBox)
    Dim imgObj As Object
    Set imgObj = stdImage.CreateFromClipboard()
    Set picDest.Picture = PictureWithOwnedBitmap(imgObj.ToStdPicture())
End Sub

Public Sub CapturarTelaParaPictureBox(ByRef picDest As PictureBox)
    Dim imgObj As Object
    Set imgObj = stdImage.CreateFromScreen()
    Set picDest.Picture = PictureWithOwnedBitmap(imgObj.ToStdPicture())
End Sub

' ==============================================================================
' 4. MÉTODOS INTOCADOS PARA RETROCOMPATIBILIDADE EXATA
' ==============================================================================

Public Sub StretchSourcePictureFromPicture(ByVal picSrc As StdPicture, ByRef picDest As PictureBox)
  Dim hMemDC As LongPtr, hDesktopDC As LongPtr, hOldBmp As LongPtr
  Dim hDesktopWnd As LongPtr, hMemWdth As Long, hMemHght As Long
  Dim Bmp As BITMAP, nRetVal As Long
  Dim OldSM As ScaleModeConstants, OldAR As Boolean, ScaleFactor As Double
  Dim ShowLeft As Long, ShowTop As Long, ShowWidth As Long, ShowHeight As Long
  Dim scaleModeChanged As Boolean, autoRedrawChanged As Boolean
  Dim errorNumber As Long, errorSource As String, errorDescription As String

  If picSrc.Handle = 0 Then Exit Sub

  On Error GoTo ErrorHandler
  hDesktopWnd = GetDesktopWindow()
  hDesktopDC = GetDC(hDesktopWnd)
  If hDesktopDC = NULL_PTR Then Exit Sub

  hMemDC = CreateCompatibleDC(hDesktopDC)
  If hMemDC = NULL_PTR Then
    Call ReleaseDC(hDesktopWnd, hDesktopDC)
    Exit Sub
  End If

  hOldBmp = SelectObject(hMemDC, picSrc.Handle)
  If hOldBmp = NULL_PTR Or hOldBmp = -1 Then
    Call DeleteDC(hMemDC)
    Call ReleaseDC(hDesktopWnd, hDesktopDC)
    Exit Sub
  End If
  nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
  hMemWdth = Bmp.BMWidth
  hMemHght = Bmp.BMHeight

  If (hMemWdth > 0) And (hMemHght > 0) Then
    With picDest
      OldSM = .ScaleMode: .ScaleMode = vbPixels
      scaleModeChanged = True
      ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / .ScaleHeight)
      ShowWidth = hMemWdth / ScaleFactor: ShowHeight = hMemHght / ScaleFactor
      ShowLeft = (.ScaleWidth - ShowWidth) / 2: ShowTop = (.ScaleHeight - ShowHeight) / 2
      OldAR = .AutoRedraw: .AutoRedraw = True
      autoRedrawChanged = True
      .Cls
      nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
      nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, vbSrcCopy)
      .Refresh
    End With
  End If

CleanUp:
  On Error GoTo 0
  Call SelectObject(hMemDC, hOldBmp)
  Call DeleteDC(hMemDC)
  Call ReleaseDC(hDesktopWnd, hDesktopDC)
  If autoRedrawChanged Then picDest.AutoRedraw = OldAR
  If scaleModeChanged Then picDest.ScaleMode = OldSM
  If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorDescription
  Exit Sub

ErrorHandler:
  errorNumber = Err.Number
  errorSource = Err.Source
  errorDescription = Err.Description
  Resume CleanUp
End Sub

Private Sub ScaleForBestFit(ByVal picSrc As StdPicture, ByRef picDest As PictureBox)
    Dim aspRatio As Single, oWid As Long, oHgt As Long, dWidth As Long, dHeight As Long

    oWid = picSrc.Picture.Width / Screen.TwipsPerPixelX
    oHgt = picSrc.Picture.Height / Screen.TwipsPerPixelY
    dWidth = picDest.ScaleWidth: dHeight = picDest.ScaleHeight
    aspRatio = oWid / oHgt
    
    If oWid > dWidth Or oHgt > dHeight Then
        If dWidth / dHeight > aspRatio Then
            dWidth = aspRatio * dHeight
        Else
            dHeight = dWidth / aspRatio
        End If
    Else
        dWidth = oWid: dHeight = oHgt
    End If
   
   picDest.Cls
   picDest.PaintPicture picSrc.Picture, 0, 0, dWidth, dHeight
   picDest.Refresh
End Sub

Private Function Biggest(Val1 As Double, Val2 As Double) As Double
  Biggest = IIf(Val1 >= Val2, Val1, Val2)
End Function

Public Sub RotateJpegFile(ByVal FilePath As String, Optional ByVal DegreesRotate As RotationEnum = Rotate_0)
    Dim img As Object, proc As Object
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile FilePath
    
    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add proc.FilterInfos("RotateFlip").FilterID
    proc.Filters(1).Properties("RotationAngle").Value = DegreesRotate
    
    Set img = proc.Apply(img)
    img.SaveFile FilePath
End Sub

Public Function RotateImg(ByVal sImgPath As String, ByVal angle As Long) As Boolean
    On Error GoTo ErroHandler
    Dim img As Object, proc As Object
    
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile sImgPath
    
    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add proc.FilterInfos("RotateFlip").FilterID
    proc.Filters(1).Properties("RotationAngle").Value = angle
    
    Set img = proc.Apply(img)
    
    If Len(Dir(sImgPath)) > 0 Then Kill sImgPath
    img.SaveFile sImgPath
    RotateImg = True
    Exit Function

ErroHandler:
    RotateImg = False
End Function

Public Function GetExifOrientation(ByVal FilePath As String) As ExifOrientationEnum
  Dim img As Object, prop As Object
  On Error GoTo errhandler
  GetExifOrientation = -1
  Set img = CreateObject("WIA.ImageFile")
  img.LoadFile FilePath
  For Each prop In img.Properties
    If prop.propertyId = EXIF_ORIENTATION Then
      GetExifOrientation = CLng(prop.Value)
      Exit Function
    End If
  Next prop
errhandler:
End Function
