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
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As LongPtr ' Ajustado para LongPtr em 64-bit
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
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
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
         ByVal xSrc As Long, ByVal ySrc As Long, _
         ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
         ByVal dwRop As Long) As Long
#End If




Public Function ADOPegBlob(ByRef cPICTURE, ByVal cARQ As String, ByVal cTable As String, Optional ByVal cWHERE As String, _
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
  Dim aRETU As Variant
  Dim cSQL As String
  Dim mystream As New ADODB.Stream
  mystream.Type = adTypeBinary


  On Error GoTo errhandler

  ADOPegBlob = False

  lOPEN = False
  lRSOP = False

  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1) '//GeracArq(cARQ, , False)
  
    sTEMPFILE = App.Path & "\" & Format(Now, "yyyymmddhhnnss") & ".jpg"
    If FileExists(sTEMPFILE) Then 'arquivo temporario pode apagar
       DeleteFile sTEMPFILE, True
    End If
  
 'Oracle 8.1 to store image you need to create a field of LongRAW type
  cSQL = cTable
  If cWHERE <> "" Then
     Select Case aRETU(2)
          Case "SQLITE"
               cSQL = "select BLOB_EXPORT(" + cCAMPO + ",'" + sTEMPFILE + "' )  as imagem from " + cTable + "  WHERE " & cWHERE
          Case Else
               cSQL = "select " + cCAMPO + " from " + cTable + "  WHERE " & cWHERE
     End Select
  End If
  
  If aRETU(2) = "PGSQL" Then
     cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
  End If

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
         Case "MYSQL", "MARIADB", "PGSQL"
               mystream.Open
               mystream.Write oRS.Fields(0)
               mystream.SaveToFile sTEMPFILE, adSaveCreateOverWrite
               If FileExists(sTEMPFILE) Then
                 eRETU01 = FileLen(sTEMPFILE)
                 Set cPICTURE.Picture = LoadPicture(sTEMPFILE)
                 ADOPegBlob = True
                 Kill sTEMPFILE
              End If
              mystream.Close
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
             DeleteFile sTEMPFILE, True
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

Public Function ADOGrvBlob(ByVal cARQ As String, ByVal cTable As String, _
                           ByRef cPICTURE, Optional ByVal cCAMPO As String = "IMAGEM", _
                           Optional ByVal cWHERE As String = "") As Boolean
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim oCMD As ADODB.Command
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim lFileLength As Long
  Dim abBytes() As Byte
  Dim iFileNum As Integer
  Dim sTEMPFILE As String
  Dim aRETU As Variant
  Dim cCMD As String
  Dim cSQL As String
  Dim mystream As New ADODB.Stream
  mystream.Type = adTypeBinary


  On Error GoTo errhandler

  ADOGrvBlob = False

  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1) ' GeracArq(cARQ)

  lOPEN = False
  lRSOP = False

  cSQL = cTable
  If cWHERE <> "" Then
    cSQL = "select " + cCAMPO + " from " + cTable + "  WHERE " & cWHERE
  End If
  
  'Oracle 8.1 to store image you need to create a field of LongRAW type
  If aRETU(2) = "MYSQL" Or aRETU(2) = "MYSQL" Then ' chave indice precisam estar o recordset
      cSQL = "select * from " + cTable + "  WHERE " & cWHERE
  End If
  
  If aRETU(2) = "PGSQL" Then
      cSQL = "select * from " + Chr(34) + UCase(cTable) + Chr(34) + "  WHERE " + Chr(34) + cWHERE 'fazendo double quotes
      cSQL = Replace(cSQL, "=", Chr(34) + "=")
  End If
  
  
  If cPICTURE.Picture.Height > 0 Then
     sTEMPFILE = App.Path & "\" & Format(Now, "yyyymmddhhnnss") & ".jpg"
     PicSaveLoad.SavePicture cPICTURE.Picture, sTEMPFILE, fmtJPEG, 70
  End If
  
  Set oDB = New ADODB.Connection
  oDB.CursorLocation = adUseClient
  oDB.ConnectionTimeout = 120
  oDB.Open cARQ

  lOPEN = True
  
Select Case aRETU(2)
     Case "SQLITE"
          '"Update IMAGENS  SET IMAGEM = blob_import('c:\temp\testered.jpg') where   codigo='red'"
          If cPICTURE.Picture.Height = 0 Then
             cCMD = "Update " + cSQL + "  SET " + cCAMPO + " = NULL where  " + cWHERE
          Else
             cCMD = "Update " + cSQL + "  SET " + cCAMPO + " = blob_import('" + sTEMPFILE + "') where  " + cWHERE
          End If
          Set oCMD = New ADODB.Command
          oCMD.ActiveConnection = oDB
          oCMD.CommandText = cCMD
          oCMD.Execute
          Set oCMD = Nothing
     Case "MYSQL", "MARIADB", "PGSQL"
        If FileExists(sTEMPFILE) Then
          Set oRS = New ADODB.Recordset
          oRS.Open cSQL, oDB, adOpenStatic, adLockOptimistic  ''adOpenStatic ADOPENKEYSET
          lRSOP = True
          If Not oRS.EOF Then
             mystream.Open
             mystream.LoadFromFile sTEMPFILE
             oRS.Fields(cCAMPO) = mystream.Read
             mystream.Close
             oRS.Update
             ADOGrvBlob = True
          End If
          DeleteFile sTEMPFILE, True
          oRS.Close
        End If
     Case Else
        Set oRS = New ADODB.Recordset
        oRS.Open cSQL, oDB, adOpenKeyset, adLockOptimistic  ''adOpenStatic
      
      
        lRSOP = True
        If Not oRS.EOF Then
          If cPICTURE.Picture.Height = 0 Then
            oRS.Fields(cCAMPO) = Null
            oRS.Update
            ADOGrvBlob = True
          Else
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
            DeleteFile sTEMPFILE, True  'Kill sTEMPFILE
          End If
        End If
        oRS.Close
        Set oRS = Nothing
  End Select
  oDB.Close
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
  hMemWdth = Bmp.bmWidth
  hMemHght = Bmp.bmHeight

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
  hMemWdth = Bmp.bmWidth
  hMemHght = Bmp.bmHeight

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


