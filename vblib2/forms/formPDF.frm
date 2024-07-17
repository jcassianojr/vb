VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Begin VB.Form formConvertToPDF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text-PDF v1.0"
   ClientHeight    =   4980
   ClientLeft      =   2625
   ClientTop       =   1395
   ClientWidth     =   8655
   Icon            =   "formPDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   8655
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   7680
      Top             =   720
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin XPControls.XPButton btnConvert 
      Height          =   435
      Left            =   7200
      TabIndex        =   24
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "formPDF.frx":08CA
      Caption         =   "Gerar Pdf"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox ChkPaisagem 
      Caption         =   "Paisagem"
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox cmbPageSize 
      Height          =   315
      ItemData        =   "formPDF.frx":0E64
      Left            =   3720
      List            =   "formPDF.frx":0E71
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2400
      Width           =   1305
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      ItemData        =   "formPDF.frx":0E98
      Left            =   2280
      List            =   "formPDF.frx":0EB4
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2400
      Width           =   690
   End
   Begin VB.ComboBox cmbRotation 
      Height          =   315
      ItemData        =   "formPDF.frx":0EE6
      Left            =   3000
      List            =   "formPDF.frx":0EF6
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2400
      Width           =   690
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      ItemData        =   "formPDF.frx":0F0F
      Left            =   120
      List            =   "formPDF.frx":0F1C
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtTitle 
      Height          =   375
      Left            =   1470
      TabIndex        =   9
      Top             =   1920
      Width           =   5100
   End
   Begin VB.TextBox txtOutputFile 
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   6960
   End
   Begin VB.TextBox txtFilename 
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   6960
   End
   Begin VB.TextBox txtSubject 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1440
      Width           =   5100
   End
   Begin VB.TextBox txtKeywords 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   5100
   End
   Begin VB.TextBox txtCreator 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   5100
   End
   Begin VB.TextBox txtAuthor 
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   45
      Width           =   5100
   End
   Begin XPControls.XPButton CmdVisua 
      Height          =   435
      Left            =   3840
      TabIndex        =   19
      Top             =   3840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "formPDF.frx":0F3D
      Caption         =   "Visualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton btnClose 
      Height          =   435
      Left            =   7320
      TabIndex        =   20
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "formPDF.frx":14D7
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton btnSave 
      Height          =   435
      Left            =   1800
      TabIndex        =   21
      Top             =   3840
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      Picture         =   "formPDF.frx":1A71
      Caption         =   "Salvar &PDF Como"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton btnOpen 
      Height          =   435
      Left            =   2040
      TabIndex        =   22
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "formPDF.frx":200B
      Caption         =   "Abrir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdSendMail 
      Height          =   435
      Left            =   5520
      TabIndex        =   23
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "formPDF.frx":25A5
      Caption         =   "EnviarEmail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   435
      Left            =   3480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "formPDF.frx":2B3F
      Caption         =   "Abrir Com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdPreviewOrigem 
      Height          =   435
      Left            =   4920
      TabIndex        =   26
      Top             =   2880
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "formPDF.frx":30D9
      Caption         =   "Visualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Caption         =   "Titulo:"
      Height          =   255
      Left            =   150
      TabIndex        =   8
      Top             =   1950
      Width           =   1230
   End
   Begin VB.Label lblOutputFile 
      Caption         =   "Arquivo Destino"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblFilename 
      Caption         =   "Arquivo para converter"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblSubject 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto:"
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   1470
      Width           =   1230
   End
   Begin VB.Label lblKeyword 
      Alignment       =   1  'Right Justify
      Caption         =   "Palavras Chave:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   1005
      Width           =   1230
   End
   Begin VB.Label lblCreator 
      Alignment       =   1  'Right Justify
      Caption         =   "Autor:"
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   540
      Width           =   1230
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "formConvertToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Position As Long
Dim pageNo As Long
Dim lineNo As Long
Dim pageHeight As Long
Dim pageWidth As Long
Dim location(1 To 5000) As Long
Dim pageObj(1 To 5000) As Long
Dim LINES As Long
Dim obj As Long
Dim Tpages As Long
Dim encoding As Long
Dim resources As Long
'Dim Pages As Variant
Dim author As String
Dim creator As String
Dim keywords As String
Dim subject As String
Dim Title As String
Dim BaseFont As String
Dim pointSize As Currency
Dim vertSpace As Currency
Dim rotate As Integer
Dim info As Long
Dim root As Long
Dim npagex As Double
Dim npagey As Long
Dim filetxt As String
Dim filepdf As String
Dim linelen As Long
Dim cache As String
'Dim cmdline As String

Const AppName = "Text-PDF v1.0"

Private Sub CmdAbrirCom_Click()
  If FileExist(txtFilename.tEXT, True) Then
    Call OpenWith(txtFilename.tEXT, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
  End If
End Sub

Private Sub cmdPreviewOrigem_Click()
  PrintPreview1.ShowPreview
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  MyPrintingTXT
End Sub

Private Sub CmdSendMail_Click()
  Dim cARQUIVO As String
  cARQUIVO = FixStr(txtOutputFile.tEXT)
  'servidor,porta,from,to,assunto,anexos,mensagem,enviar e sair
  ePASS01 = Array("", _
                  "", _
                  "", _
                  "", _
                  "", _
                  cARQUIVO, _
                  "", _
                  False)
  FrmSendMail.Show vbModal, Me
End Sub

Private Sub CmdVisua_Click()
  Dim cARQSHELL As String
  cARQSHELL = FixStr(txtOutputFile.tEXT)
  If FileExist(cARQSHELL) Then
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , Me.hWnd
  End If
End Sub

Private Sub Form_Load()

'  CentralizaJanela Me
  CenterFormToScreen Me
  '    txtCreator.tEXT = AppName
  cmbFont.ListIndex = 1                        ' 10 pt
  cmbFontSize.ListIndex = 0
  cmbRotation.ListIndex = 0
  cmbPageSize.ListIndex = 0

  txtCreator.tEXT = zNOMEFOLHA



  If Len(ePASS01) > 0 Then
    If FileExist(ePASS01) Then
      txtTitle.tEXT = NomeArq(ePASS01, True)
      txtFilename.tEXT = ePASS01
      txtOutputFile.tEXT = TrocaExt(ePASS01, "PDF")
      btnOpen.Enabled = False
      txtFilename.Enabled = False
      '        btnConvert_Click
    End If
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Private Sub txtAuthor_GotFocus()
  txtAuthor.SelStart = 0
  txtAuthor.SelLength = Len(txtAuthor.tEXT)
End Sub

Private Sub txtCreator_GotFocus()
  txtCreator.SelStart = 0
  txtCreator.SelLength = Len(txtCreator.tEXT)
End Sub

Private Sub txtSubject_GotFocus()
  txtSubject.SelStart = 0
  txtSubject.SelLength = Len(txtSubject.tEXT)
End Sub

Private Sub txtTitle_GotFocus()
  txtTitle.SelStart = 0
  txtTitle.SelLength = Len(txtTitle.tEXT)
End Sub

Private Sub txtKeywords_GotFocus()
  txtKeywords.SelStart = 0
  txtKeywords.SelLength = Len(txtKeywords.tEXT)
End Sub

Private Sub txtFilename_GotFocus()
  txtFilename.SelStart = 0
  txtFilename.SelLength = Len(txtFilename.tEXT)
End Sub

Private Sub txtOutputFile_GotFocus()
  txtOutputFile.SelStart = 0
  txtOutputFile.SelLength = Len(txtOutputFile.tEXT)
End Sub

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnOpen_Click()
  On Local Error Resume Next


  txtFilename.tEXT = OpenArqExt(Me, txtFilename.tEXT, "txt", "Arquivos txt")

  If Not Len(txtFilename.tEXT) = 0 Then
    txtOutputFile.tEXT = Left(txtFilename.tEXT, Len(txtFilename.tEXT) - 3) & "pdf"
  End If


End Sub

Private Sub btnSave_Click()
  Dim sFILTER, FileName As String
  sFILTER = "Arquivos PDF (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  FileName = FileSave(Me, sFILTER, 1, "TXT", cARQRTF, cARQRTF, "Salvar PDF")
  If Not Len(FileName) = 0 Then
    txtOutputFile.tEXT = FileName
  End If
End Sub

Private Sub btnSource_Click()
  On Local Error Resume Next
End Sub

Private Sub btnConvert_Click()
  If txtFilename.tEXT <> "" And txtOutputFile.tEXT <> "" Then
    ConvertToPDF txtFilename.tEXT, txtOutputFile.tEXT, _
                 txtAuthor.tEXT, txtCreator.tEXT, txtKeywords.tEXT, _
                 txtSubject.tEXT, txtTitle.tEXT, _
                 cmbFont.tEXT, Val(cmbFontSize.tEXT), Val(cmbRotation.tEXT), _
                 Val(cmbPageSize.tEXT), Val(Right(cmbPageSize.tEXT, 3)), ChkPaisagem.Value
    If FileExist(txtOutputFile.tEXT) Then
      Alert "Conversao Concluida"
    End If

  Else
    MsgBox "Informe o nome do arquivo."
  End If
End Sub

Public Function ConvertToPDF(FileName As String, outputfile As String, _
                             Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
                             Optional TextSubject As String, Optional TextTitle As String, _
                             Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
                             Optional pwidth As Single = 8.5, Optional pheight As Single = 11, _
                             Optional lPAISAGEM As Boolean = False)

  Dim tpwidth As Long
  Dim tpheight As Long
  On Error GoTo er
  If Not FileExist(FileName) Then
    MsgBox "Arquivo '" & FileName & "' n„o existe."
    Exit Function
  ElseIf FileExist(outputfile) Then
    If MDG("Arquivo Ja Existente" & outputfile, "Confirma Exclusao") Then
      DeleteFile outputfile  'Kill outputfile
    End If
  End If

  If lPAISAGEM Then
    tpwidth = pwidth
    tpheight = pheight
    pwidth = tpheight
    pheight = tpwidth
  End If


  initialize FontName, FontSize, Rotation, pwidth, pheight

  author = TextAuthor
  creator = TextCreator
  keywords = TextKeywords
  subject = TextSubject
  Title = TextTitle
  filetxt = FileName
  filepdf = outputfile

  Call WriteStart
  Call WriteHead
  Call WritePages
  Call endpdf
  Exit Function
er:
  MsgBox Err.Description
End Function

Private Sub initialize(FontName As String, FontSize As Integer, Rotation As Integer, pwidth As Single, pheight As Single)
  pageHeight = 72 * pheight
  pageWidth = 72 * pwidth

  BaseFont = FontName                          ' Courier, Times-Roman, Arial
  pointSize = FontSize                         ' Font Size; n„o altere
  vertSpace = FontSize * 1.2                   ' Vertical spacing
  rotate = Rotation                            ' degrees to rotate; try setting 90,180,etc
  LINES = (pageHeight - 72) / vertSpace        ' no of lines on one page

  Select Case LCase(FontName)
  Case "courier": linelen = 1.5 * pageWidth / pointSize
  Case "arial": linelen = 2 * pageWidth / pointSize
    'Case "Times-Roman": linelen = 2.2 * pageWidth / pointSize
  Case Else: linelen = 2.2 * pageWidth / pointSize
  End Select

  obj = 0
  npagex = pageWidth / 2
  npagey = 25
  pageNo = 0
  Position = 0
  cache = ""
End Sub

Private Sub writepdf(stre As String, Optional flush As Boolean)
  On Local Error Resume Next
  Position = Position + Len(stre)
  cache = cache & stre & vbCr
  If Len(cache) > 32000 Or flush Then
    Open filepdf For Append As #1
    Print #1, cache;
    Close #1
    cache = ""
  End If
End Sub

Private Sub WriteStart()
  writepdf ("%PDF-1.2")
  writepdf ("%‚„œ”")
End Sub

Private Sub WriteHead()
  Dim CreationDate As String
  On Error GoTo er
  CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
  obj = obj + 1
  location(obj) = Position
  info = obj

  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Author (" & author & ")")
  writepdf ("/CreationDate (" & CreationDate & ")")
  writepdf ("/Creator (" & creator & ")")
  writepdf ("/Producer (" & AppName & ")")
  writepdf ("/Title (" & Title & ")")
  writepdf ("/Subject (" & subject & ")")
  writepdf ("/Keywords (" & keywords & ")")
  writepdf (">>")
  writepdf ("endobj")

  obj = obj + 1
  root = obj
  obj = obj + 1
  Tpages = obj
  encoding = obj + 2
  resources = obj + 3

  obj = obj + 1
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Font")
  writepdf ("/Subtype /Type1")
  writepdf ("/Name /F1")
  writepdf ("/Encoding " & encoding & " 0 R")
  writepdf ("/BaseFont /" & BaseFont)
  writepdf (">>")
  writepdf ("endobj")

  obj = obj + 1
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Encoding")
  writepdf ("/BaseEncoding /WinAnsiEncoding")
  writepdf (">>")
  writepdf ("endobj")

  obj = obj + 1
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("  /Font << /F1 " & obj - 2 & " 0 R >>")
  writepdf ("  /ProcSet [ /PDF /Text ]")
  writepdf (">>")
  writepdf ("endobj")
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Private Sub WritePages()
  Dim i As Integer
  Dim line As String, tmpline As String, beginstream As String
  On Error GoTo er
  Open filetxt For Input As #2
  beginstream = StartPage
  lineNo = -1
  Do Until EOF(2)
    Line Input #2, line
    lineNo = lineNo + 1

    'quebra de p·gina
    If lineNo >= LINES Or InStr(line, Chr(12)) > 0 Then
      writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
      writepdf ("(" & pageNo & ") Tj")
      writepdf ("/F1 " & pointSize & " Tf")
      endpage (beginstream)
      beginstream = StartPage
    End If

    line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
    line = Trim(line)

    If Len(line) > linelen Then

      'quebra de linha
      Do While Len(line) > linelen
        tmpline = Left(line, linelen)
        For i = Len(tmpline) To Len(tmpline) \ 2 Step -1
          If InStr("*&^%$#,. ;<=>[])}!""", Mid(tmpline, i, 1)) Then
            tmpline = Left(tmpline, i)
            Exit For
          End If
        Next

        line = Mid$(line, Len(tmpline) + 1)
        writepdf ("T* (" & tmpline & vbCrLf & ") Tj")
        lineNo = lineNo + 1

        'quebra de p·gina
        If lineNo >= LINES Or InStr(line, Chr(12)) > 0 Then
          writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
          writepdf ("(" & pageNo & ") Tj")
          writepdf ("/F1 " & pointSize & " Tf")
          endpage (beginstream)
          beginstream = StartPage
        End If
      Loop

      lineNo = lineNo + 1
      writepdf ("T* (" & line & vbCrLf & ") Tj")

    Else

      writepdf ("T* (" & line & vbCrLf & ") Tj")

    End If
  Loop
  Close #2
  writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
  writepdf ("(" & pageNo & ") Tj")
  writepdf ("/F1 " & pointSize & " Tf")
  endpage (beginstream)
  Exit Sub
er:
  MsgBox Err.Description
  Close
End Sub

Private Function StartPage() As String
  Dim strmpos As Long
  On Error GoTo er
  obj = obj + 1
  location(obj) = Position
  pageNo = pageNo + 1
  pageObj(pageNo) = obj

  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Page")
  writepdf ("/Parent " & Tpages & " 0 R")
  writepdf ("/Resources " & resources & " 0 R")
  obj = obj + 1
  writepdf ("/Contents " & obj & " 0 R")
  writepdf ("/Rotate " & rotate)
  writepdf (">>")
  writepdf ("endobj")

  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Length " & obj + 1 & " 0 R")
  writepdf (">>")
  writepdf ("stream")
  strmpos = Position
  writepdf ("BT")
  writepdf ("/F1 " & pointSize & " Tf")
  writepdf ("1 0 0 1 50 " & pageHeight - 40 & " Tm")
  writepdf (vertSpace & " TL")

  StartPage = strmpos
  Exit Function
er:
  MsgBox Err.Description
End Function

Function endpage(streamstart As Long) As String
  Dim streamEnd As Long
  On Error GoTo er
  writepdf ("ET")
  streamEnd = Position
  writepdf ("endstream")
  writepdf ("endobj")
  obj = obj + 1
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf (streamEnd - streamstart)
  writepdf "endobj"
  lineNo = 0
  Exit Function
er:
  MsgBox Err.Description
End Function

Sub endpdf()
  Dim ty As String, i As Integer, xreF As Long
  On Error GoTo er
  location(root) = Position
  writepdf (root & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Catalog")
  writepdf ("/Pages " & Tpages & " 0 R")
  writepdf (">>")
  writepdf ("endobj")
  location(Tpages) = Position
  writepdf (Tpages & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Pages")
  writepdf ("/Count " & pageNo)
  writepdf ("/MediaBox [ 0 0 " & pageWidth & " " & pageHeight & " ]")
  ty = ("/Kids [ ")
  For i = 1 To pageNo
    ty = ty & pageObj(i) & " 0 R "
  Next i
  ty = ty & "]"
  writepdf (ty)
  writepdf (">>")
  writepdf ("endobj")
  xreF = Position
  writepdf ("0 " & obj + 1)
  writepdf ("0000000000 65535 f ")
  For i = 1 To obj
    writepdf (Format(location(i), "0000000000") & " 00000 n ")
  Next i
  writepdf ("trailer")
  writepdf ("<<")
  writepdf ("/Size " & obj + 1)
  writepdf ("/Root " & root & " 0 R")
  writepdf ("/Info " & info & " 0 R")
  writepdf (">>")
  writepdf ("startxref")
  writepdf (xreF)
  writepdf "%%EOF", True
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Public Function ReplaceText(tEXT As String, TextToReplace As String, NewText As String) As String
  Dim mtext As String, SpacePos As Long
  mtext = tEXT
  SpacePos = InStr(mtext, TextToReplace)
  Do While SpacePos
    mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
    SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
  Loop
  ReplaceText = mtext
End Function
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  Dim cARQOPEN As String
  cARQOPEN = txtFilename.tEXT
  If Not FileExist(cARQOPEN, True) Then
    Exit Sub
  End If
  fileFile = FreeFile
  Open cARQOPEN For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
End Sub

