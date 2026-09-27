Attribute VB_Name = "jpgpicture"
Option Explicit

' http://www.codenewsgroups.net/group/microsoft.public.vb.general.discussion/topic3286.aspx
' http://edais.mvps.org/


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
' --- ESTRUTURA DE COMPATIBILIDADE ---
#If VBA7 Or Win64 Then
    ' --- VERSÃO 64-BIT / TWINBASIC / VBA7 ---
    Private Type BITMAP
        BMType As Long
        BMWidth As Long
        BMHeight As Long
        BMWidthBytes As Long
        BMPlanes As Integer
        BMBitsPixel As Integer
        BMBits As LongPtr ' Ajustado para LongPtr em 64-bit
    End Type

    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As LongPtr
    Private Declare PtrSafe Function GetObject Lib "gdi32" Alias "GetObjectA" _
        (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
        Public Declare PtrSafe Function SetStretchBltMode Lib "gdi32" _
        (ByVal hDC As LongPtr, ByVal nStretchMode As Long) As Long
   Public Declare PtrSafe Function StretchBlt Lib "gdi32" _
        (ByVal hdcDest As LongPtr, _
         ByVal xDest As Long, ByVal yDest As Long, _
         ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
         ByVal hdcSrc As LongPtr, _
         ByVal xSrc As Long, ByVal ySrc As Long, _
         ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
         ByVal dwRop As Long) As Long
#Else
    ' --- VERSÃO 32-BIT CLÁSSICA (VB6) ---
    Private Type BITMAP
        BMType As Long
        BMWidth As Long
        BMHeight As Long
        BMWidthBytes As Long
        BMPlanes As Integer
        BMBitsPixel As Integer
        BMBits As Long
    End Type

    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Public Declare Function GetDesktopWindow Lib "user32" () As Long
    Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
        (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Public Declare Function SetStretchBltMode Lib "gdi32" _
        (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
    Public Declare Function StretchBlt Lib "gdi32" _
        (ByVal hdcDest As Long, _
         ByVal xDest As Long, ByVal yDest As Long, _
         ByVal nWidthDest As Long, ByVal nHeightDest As Long, _
         ByVal hdcSrc As Long, _
         ByVal XSrc As Long, ByVal YSrc As Long, _
         ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
         ByVal dwRop As Long) As Long
#End If


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
Public Const EXIF_ORIENTATION As Long = 274
Public Enum RotationEnum
    Rotate_0 = 0
    Rotate_90 = 90
    Rotate_180 = 180
    Rotate_270 = 270
End Enum

Public Sub RotateJpegFile(ByVal FilePath As String, Optional ByVal DegreesRotate As RotationEnum = Rotate_0)

    Dim img As Object
    Dim proc As Object
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile FilePath

    Set proc = CreateObject("WIA.ImageProcess")
    proc.Filters.Add proc.FilterInfos("RotateFlip").FilterID
    proc.Filters(1).Properties("RotationAngle").Value = DegreesRotate

    Set img = proc.Apply(img)
    img.SaveFile FilePath
End Sub

Public Function GetExifOrientation(ByVal FilePath As String) As ExifOrientationEnum
  Dim img As Object
  Dim prop As Object
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
  Debug.Print Err.Number & " - " & Err.Description
End Function







Public Function StretchSourcePictureFromFile(ByVal filename As String, ByRef picDest As PictureBox) As StdPicture
  Dim hMemDC As Long
  Dim hOldBmp As Long
  Dim hMemWdth As Long
  Dim hMemHght As Long
  Dim Bmp As BITMAP
  Dim nRetVal As Long
  Dim picSrc As StdPicture
  Dim OldSM As ScaleModeConstants
  Dim OldAR As Boolean
  Dim ScaleFactor As Double
  Dim ShowLeft As Long
  Dim ShowTop As Long
  Dim ShowWidth As Long
  Dim ShowHeight As Long

  If Len(filename) = 0 Then
    Beep
    Exit Function
  End If

  'Create the memory DC
  hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
  'Load the picture
  'Set picSrc = LoadPictureEx(FileName)
  Set picSrc = LoadPicture(filename)

  'Assign the picture to the memory DC
  hOldBmp = SelectObject(hMemDC, picSrc.Handle)

  'Get the sizes of the picture
  nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
  hMemWdth = Bmp.BMWidth
  hMemHght = Bmp.BMHeight

  'Make sure there is a picture
  If (hMemWdth > 0) And (hMemHght > 0) Then

    'Stretch the picture to the picturebox
    With picDest
      'Save the PictureBox's ScaleMode and set it to vbPixels
      OldSM = .ScaleMode
      .ScaleMode = vbPixels

      'Get the largest possible scaling factor
      ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / _
                                                    .ScaleHeight)

      'Get the positions and sizes for the destination picture
      ShowWidth = hMemWdth / ScaleFactor
      ShowHeight = hMemHght / ScaleFactor
      ShowLeft = (.ScaleWidth - ShowWidth) / 2
      ShowTop = (.ScaleHeight - ShowHeight) / 2

      'Save the PictureBox's AutoRedraw and set it to True
      OldAR = .AutoRedraw
      .AutoRedraw = True

      '.Picture = LoadPictureEx()
      .Picture = LoadPicture()
      .Cls
      nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
      nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, _
                           ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, _
                           vbSrcCopy)

      If (nRetVal = 0) Then
        Debug.Print "StretchBlt() Error Code " & _
                    Err.LastDllError
      End If

      .Refresh
    End With

    'Reset the PictureBox's ScaleMode
    picDest.ScaleMode = OldSM

    'Reset the PictureBox's AutoRedraw
    picDest.AutoRedraw = OldAR
  End If

  'Return the picture object
  Set StretchSourcePictureFromFile = picSrc

  'Clean up the used memory
  Call SelectObject(hMemDC, hOldBmp)
  Call DeleteDC(hMemDC)
  Set picSrc = Nothing
End Function

Public Function lerarquivoimagem(ByVal STMPFILE, ByRef Picture1 As PictureBox, ByRef Picture2 As PictureBox)
  lerarquivoimagem = False
  If Len(STMPFILE) > 0 Then
    If FixInt(FileLen(STMPFILE)) > 500000 Then
      Alert ("Imagem Muito Grande,Ajuste o tamanho")
      If Not MDG("Anexar mesmo assim-NAO RECOMENDADO") Then
        Exit Function
      Else
        Exit Function                    'nao permitindo aumentando o banco e travando relatorio crystal
      End If
    End If
    Picture1.Picture = LoadPicture(STMPFILE)
    StretchSourcePictureFromPicture Picture1, Picture2
    lerarquivoimagem = True
  End If
End Function
Private Sub ScaleForBestFit(ByVal picSrc As StdPicture, _
                                           ByRef picDest As PictureBox)
    Dim aspRatio As Single, oWid As Long, oHgt As Long
    Dim dWidth As Long, dHeight As Long

 'no form propriedades no componente ou no load
 'Picture1.ScaleMode = 3                            ' Pixels
  '  Picture2.ScaleMode = 3
 '   Picture1.AutoRedraw = True
  

    ' Get original image dimensions
    oWid = picSrc.Picture.Width / Screen.TwipsPerPixelX ' Convert to pixels
    oHgt = picSrc.Picture.Height / Screen.TwipsPerPixelY
    
    ' Desired dimensions (Picture1 area)
    dWidth = picDest.ScaleWidth
    dHeight = picDest.ScaleHeight
   
    ' Calculate aspect ratio
    aspRatio = oWid / oHgt
    
    ' Calculate best fit
    If oWid > dWidth Or oHgt > dHeight Then
        If dWidth / dHeight > aspRatio Then
            dWidth = aspRatio * dHeight
        Else
            dHeight = dWidth / aspRatio
        End If
    Else
        dWidth = oWid
        dHeight = oHgt
    End If
   
   ' Draw result to destination picturebox (Picture1)
   picDest.Cls
   picDest.PaintPicture picSrc.Picture, 0, 0, dWidth, dHeight
   picDest.Refresh

End Sub




Public Sub StretchSourcePictureFromPicture(ByVal picSrc As StdPicture, _
                                           ByRef picDest As PictureBox)
  Dim hMemDC As Long
  Dim hOldBmp As Long
  Dim hMemWdth As Long
  Dim hMemHght As Long
  Dim Bmp As BITMAP
  Dim nRetVal As Long
  Dim OldSM As ScaleModeConstants
  Dim OldAR As Boolean
  Dim ScaleFactor As Double
  Dim ShowLeft As Long
  Dim ShowTop As Long
  Dim ShowWidth As Long
  Dim ShowHeight As Long

  'Make sure we have a valid picture
  If picSrc.Handle = 0 Then
    Beep
    Exit Sub
  End If

  'Create the memory DC
  hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
  'Assign the picture to the memory DC
  hOldBmp = SelectObject(hMemDC, picSrc.Handle)

  'Get the sizes of the picture
  nRetVal = GetObject(picSrc.Handle, Len(Bmp), Bmp)
  hMemWdth = Bmp.BMWidth
  hMemHght = Bmp.BMHeight

  'Make sure there is a picture
  If (hMemWdth > 0) And (hMemHght > 0) Then

    'Stretch the picture to the picturebox
    With picDest
      'Save the PictureBox's ScaleMode and set it to vbPixels
      OldSM = .ScaleMode
      .ScaleMode = vbPixels

      'Get the largest possible scaling factor
      ScaleFactor = Biggest(hMemWdth / .ScaleWidth, hMemHght / _
                                                    .ScaleHeight)

      'Get the positions and sizes for the destination picture
      ShowWidth = hMemWdth / ScaleFactor
      ShowHeight = hMemHght / ScaleFactor
      ShowLeft = (.ScaleWidth - ShowWidth) / 2
      ShowTop = (.ScaleHeight - ShowHeight) / 2

      'Save the PictureBox's AutoRedraw and set it to True
      OldAR = .AutoRedraw
      .AutoRedraw = True

      .Cls
      nRetVal = SetStretchBltMode(.hDC, STRETCH_HALFTONE)
      nRetVal = StretchBlt(.hDC, ShowLeft, ShowTop, ShowWidth, _
                           ShowHeight, hMemDC, 0, 0, hMemWdth, hMemHght, _
                           vbSrcCopy)

      If (nRetVal = 0) Then
        Debug.Print "StretchBlt() Error Code " & _
                    Err.LastDllError
      End If

      .Refresh
    End With

    'Reset the PictureBox's ScaleMode
    picDest.ScaleMode = OldSM

    'Reset the PictureBox's AutoRedraw
    picDest.AutoRedraw = OldAR
  End If

  'Clean up the used memory
  Call SelectObject(hMemDC, hOldBmp)
  Call DeleteDC(hMemDC)
End Sub

Private Function Biggest(Val1 As Double, Val2 As Double) As Double
  Biggest = IIf(Val1 >= Val2, Val1, Val2)
End Function

Public Function salvarpict(oFORM As Form, ByVal Picture1 As Variant, _
                           Optional ByVal sFileName As String = "imagem", _
                           Optional ByVal sPath As String = "")
  Dim sFILTER As String
  Dim cEXTENSAO As String

  If Len(sPath) = 0 Then
    sPath = App.Path
  End If
  sFILTER = ImgFILTER2()
  sFileName = FileSave(oFORM, sFILTER, 1, , sFileName, sPath, "Salvar Imagem")

  cEXTENSAO = parsefile(sFileName, "E")

  If Len(sFileName) <= 0 Then
    Alert ("Nome do arquivo nao Preenchido")
    Exit Function
  End If
  If FileConnExist(sFileName, False) Then
    If MDG("Arquivo de Destino Ja existe Sobrepor") Then
      DeleteFile sFileName, True  'Kill sFILENAME
    Else
      Exit Function
    End If
  End If
  Select Case cEXTENSAO
  Case "JPG"
    PicSaveLoad.SavePicture Picture1.Picture, sFileName, fmtJPEG, 70
  Case "PNG"
    PicSaveLoad.SavePicture Picture1.Picture, sFileName, fmtPNG
  Case "GIF"
    PicSaveLoad.SavePicture Picture1.Picture, sFileName, fmtGIF
  Case Else 'usa default vb
    SavePicture Picture1.Picture, sFileName
  End Select
End Function
Private Function Base64ToByte(ByVal sBase64 As String) As Byte()
    Dim oXML As Object
    Dim oNode As Object
    
    Set oXML = CreateObject("MSXML2.DOMDocument")
    Set oNode = oXML.createElement("b64")
    oNode.DataType = "bin.base64"
    oNode.Text = sBase64
    Base64ToByte = oNode.nodeTypedValue
    Set oNode = Nothing
    Set oXML = Nothing
End Function
Public Sub SalvarBase64ComoImagem(ByVal sBase64 As String, ByVal sCaminhoDestino As String)
    Dim sDados As String
    Dim baImagem() As Byte
    Dim hFile As Integer
    
    ' 1. Remove o prefixo (ex: "data:image/png;base64,")
    ' Encontra a posição da vírgula
    If InStr(sBase64, ",") > 0 Then
        sDados = Mid(sBase64, InStr(sBase64, ",") + 1)
    Else
        sDados = sBase64
    End If
    
    ' 2. Converte a string Base64 para Array de Bytes
    ' Nota: Você precisará de uma função "Base64ToByte" (abaixo)
    baImagem = Base64ToByte(sDados)
    
    ' 3. Grava no disco
    hFile = FreeFile
    Open sCaminhoDestino For Binary Access Write As #hFile
    Put #hFile, , baImagem
    Close #hFile
End Sub

 Public Function RotateImg(ByVal sImgPath As String, ByVal angle As Long) As Boolean
        Dim pFact As ShellImageDataFactory
        Dim pShImg As IShellImageData
        Set pFact = New ShellImageDataFactory
        pFact.CreateImageFromFile StrPtr(sImgPath), pShImg
        pShImg.Decode SHIMGDEC_DEFAULT, 0, 0
        pShImg.Rotate angle
        Dim ipf As IPersistFile
        Set ipf = pShImg
        ipf.Save sImgPath, 1
    End Function
