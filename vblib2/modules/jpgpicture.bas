Attribute VB_Name = "jpgpicture"
Option Explicit




' http://www.codenewsgroups.net/group/microsoft.public.vb.general.discussion/topic3286.aspx
' http://edais.mvps.org/

'---------------------------------Start Code------------------
Private Type BITMAP                              '14 bytes
  BMType As Long
  BMWidth As Long
  BMHeight As Long
  BMWidthBytes As Long
  BMPlanes As Integer
  BMBitsPixel As Integer
  BMBits As Long
End Type

'Win32 API Function Declarations
Private Declare Function CreateCompatibleDC _
                          Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC _
                          Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject _
                          Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC _
                          Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDesktopWindow _
                          Lib "user32" () As Long

Private Declare Function GetObject _
                          Lib "gdi32" _
                              Alias "GetObjectA" (ByVal hObject As Long, _
                                                  ByVal nCount As Long, _
                                                  lpObject As Any) As Long

Private Declare Function SelectObject _
                          Lib "gdi32" (ByVal hDC As Long, _
                                       ByVal hObject As Long) As Long

Private Declare Function SetStretchBltMode _
                          Lib "gdi32" (ByVal hDC As Long, _
                                       ByVal nStretchMode As Long) As Long

Private Declare Function StretchBlt _
                          Lib "gdi32" (ByVal hDC As Long, _
                                       ByVal x As Long, _
                                       ByVal y As Long, _
                                       ByVal nWidth As Long, _
                                       ByVal nHeight As Long, _
                                       ByVal hSrcDC As Long, _
                                       ByVal XSrc As Long, _
                                       ByVal YSrc As Long, _
                                       ByVal nSrcWidth As Long, _
                                       ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

'Win32 API Constant Declarations
Private Const STRETCH_HALFTONE = 4

Public Function ADOPegBlob(ByRef cPICTURE, ByVal cARQ As String, ByVal cTABLE As String, Optional ByVal cWHERE As String, _
                           Optional ByVal cCAMPO As String = "IMAGEM") As Boolean
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim lFileLength As Long
  Dim abBytes() As Byte
  Dim iFileNum As Integer
  Dim sTEMPFILE As String
  Dim aRETU
  Dim cSQL As String


  On Error GoTo errhandler

  ADOPegBlob = False

  lOPEN = False
  lRSOP = False

  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1) '//GeracArq(cARQ, , False)
  
    sTEMPFILE = App.Path & "\" & Format(Now, "yyyymmddhhnnss") & ".jpg"
    If FileExists(sTEMPFILE) Then 'arquivo temporario pode apagar
       DeleteFile sTEMPFILE
    End If
  
  cSQL = cTABLE
  If cWHERE <> "" Then
     Select Case aRETU(2)
          Case "SQLITE"
               cSQL = "select BLOB_EXPORT(" + cCAMPO + ",'" + sTEMPFILE + "' )  as imagem from " + cTABLE + "  WHERE " & cWHERE
          Case Else
               cSQL = "select " + cCAMPO + " from " + cTABLE + "  WHERE " & cWHERE
     End Select
  End If
  
  
  '"select IMAGEM from imagens  WHERE CODIGO='" & ZGRP & "'"
  

  Set oDB = New ADODB.Connection
  oDB.CursorLocation = adUseClient
  oDB.ConnectionTimeout = 120
  oDB.Open cARQ
  lOPEN = True
  eRETU01 = 0

  Set oRS = New ADODB.Recordset
  oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
  lRSOP = True
  If Not oRS.EOF Then
    If Not IsNull(oRS(cCAMPO)) Then
    
       Select Case aRETU(2)
         Case "SQLITE"
              If FileExists(sTEMPFILE) Then
                 eRETU01 = FileLen(sTEMPFILE)
                 Set cPICTURE.Picture = LoadPicture(sTEMPFILE)
                 ADOPegBlob = True
                 Kill sTEMPFILE
              End If
         Case Else
           lFileLength = LenB(oRS(cCAMPO))
           If lFileLength > 1 Then
              iFileNum = FreeFile
              Open sTEMPFILE For Binary As #iFileNum
              abBytes = oRS(cCAMPO).GetChunk(lFileLength)
              Put #iFileNum, , abBytes()
              Close #iFileNum
      
              If Not sTEMPFILE = "" Then
                 Set cPICTURE.Picture = LoadPicture(sTEMPFILE)
                eRETU01 = FileLen(sTEMPFILE)
                ADOPegBlob = True
              End If
             DeleteFile sTEMPFILE
            End If
        End Select
    End If
  End If

  oRS.Close
  oDB.Close
  Set oRS = Nothing
  Set oDB = Nothing
  Exit Function




errhandler:
  cERRO = "AdoPegBlob" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
  If lRSOP Then
    cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
  End If
  Select Case Err.Number
  Case Else
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Function
  End Select

End Function

Public Function ADOGrvBlob(ByVal cARQ As String, ByVal cSQL As String, _
                           ByRef cPICTURE, Optional ByVal cCAMPO As String = "IMAGEM") As Boolean
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim lFileLength As Long
  Dim abBytes() As Byte
  Dim iFileNum As Integer
  Dim sTEMPFILE As String



  On Error GoTo errhandler

  ADOGrvBlob = False

  cARQ = GeracArq(cARQ)

  lOPEN = False
  lRSOP = False

  Set oDB = New ADODB.Connection
  oDB.CursorLocation = adUseClient
  oDB.ConnectionTimeout = 120
  oDB.Open cARQ



  lOPEN = True

  Set oRS = New ADODB.Recordset
  oRS.Open cSQL, oDB, adOpenKeyset, adLockOptimistic  ''adOpenStatic


  lRSOP = True
  If Not oRS.EOF Then
    If cPICTURE.Picture.Height = 0 Then
      oRS.Fields(cCAMPO) = Null
      oRS.Update
      ADOGrvBlob = True
    Else
      sTEMPFILE = App.Path & "\" & Format(Now, "yyyymmddhhnnss") & ".jpg"

      PicSave.SavePicture cPICTURE.Picture, sTEMPFILE, fmtJPEG, 70
      'SavePictureEx cPICTURE.Picture, sTEMPFILE, FIF_JPEG


      iFileNum = FreeFile
      Open sTEMPFILE For Binary Access Read As #iFileNum
      lFileLength = LOF(iFileNum)
      ReDim abBytes(lFileLength)
      Get #iFileNum, , abBytes()
      Close #iFileNum
      'put byte array contents into db field
      If Not sTEMPFILE = "" Then
        oRS.Fields(cCAMPO).AppendChunk abBytes()
        oRS.Update
        ADOGrvBlob = True
      End If
      'sTEMPFILE = sTEMPFILE & ".jpg"
      DeleteFile sTEMPFILE  'Kill sTEMPFILE
    End If
  End If

  oRS.Close
  oDB.Close
  Set oRS = Nothing
  Set oDB = Nothing
  Exit Function




errhandler:
  cERRO = "AdoGRVBlob" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
  If lRSOP Then
    cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
  End If
  Select Case Err.Number
  Case Else
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Function
  End Select

End Function

Public Function StretchSourcePictureFromFile(ByVal FileName As String, ByRef picDest As PictureBox) As StdPicture
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

  If Len(FileName) = 0 Then
    Beep
    Exit Function
  End If

  'Create the memory DC
  hMemDC = CreateCompatibleDC(GetDC(GetDesktopWindow()))
  'Load the picture
  'Set picSrc = LoadPictureEx(FileName)
  Set picSrc = LoadPicture(FileName)

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
                           Optional ByVal sFILENAME As String = "imagem", _
                           Optional ByVal sPath As String = "")
  Dim sFILTER As String
  Dim cEXTENSAO As String

  If Len(sPath) = 0 Then
    sPath = App.Path
  End If
  sFILTER = ImgFILTER2()                       ''ImgFILTER()
  sFILENAME = FileSave(oFORM, sFILTER, 1, , sFILENAME, sPath, "Salvar Imagem")

  cEXTENSAO = parsefile(sFILENAME, "E")

  If Len(sFILENAME) <= 0 Then
    Alert ("Nome do arquivo nao Preenchido")
    Exit Function
  End If
  If FileExist(sFILENAME, False) Then
    If MDG("Arquivo de Destino Ja existe Sobrepor") Then
      DeleteFile sFILENAME  'Kill sFILENAME
    Else
      Exit Function
    End If
  End If
  Select Case cEXTENSAO
  Case "JPG"
    PicSave.SavePicture Picture1.Picture, sFILENAME, fmtJPEG, 70
  Case "PNG"
    PicSave.SavePicture Picture1.Picture, sFILENAME, fmtPNG
  Case "GIF"
    PicSave.SavePicture Picture1.Picture, sFILENAME, fmtGIF


  Case Else
    SavePicture Picture1.Picture, sFILENAME
  End Select
End Function


