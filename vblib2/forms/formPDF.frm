VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form formConvertToPDF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text-PDF v1.0"
   ClientHeight    =   5772
   ClientLeft      =   2628
   ClientTop       =   1392
   ClientWidth     =   11004
   Icon            =   "formPDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   11004
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   9360
      Top             =   4320
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin XPControls.XPButton btnConvert 
      Height          =   552
      Left            =   6720
      TabIndex        =   21
      Top             =   3240
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   974
      Picture         =   "formPDF.frx":08CA
      Caption         =   "Criar Pdf (Conversor)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
   Begin VBCCR18.ComboBoxW cmbPageSize 
      Height          =   288
      Left            =   3720
      TabIndex        =   17
      Top             =   2400
      Width           =   1305
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin VBCCR18.ComboBoxW cmbFontSize 
      Height          =   288
      Left            =   2280
      TabIndex        =   15
      Top             =   2400
      Width           =   690
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin VBCCR18.ComboBoxW cmbRotation 
      Height          =   288
      Left            =   3000
      TabIndex        =   16
      Top             =   2400
      Width           =   690
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
   End
   Begin VBCCR18.ComboBoxW cmbFont 
      Height          =   288
      Left            =   120
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
      _ExtentX        =   0
      _ExtentY        =   0
      Style           =   2
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
      Left            =   1800
      TabIndex        =   13
      Top             =   3960
      Width           =   6480
   End
   Begin VB.TextBox txtFilename 
      Height          =   345
      Left            =   1920
      TabIndex        =   11
      Top             =   2880
      Width           =   6480
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
   Begin XPControls.XPButton btnSave 
      Height          =   432
      Left            =   8400
      TabIndex        =   19
      Top             =   3960
      Width           =   2172
      _ExtentX        =   3831
      _ExtentY        =   762
      Picture         =   "formPDF.frx":0E64
      Caption         =   "Salvar Como"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton btnOpen 
      Height          =   432
      Left            =   3000
      TabIndex        =   20
      Top             =   3240
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   762
      Picture         =   "formPDF.frx":13FE
      Caption         =   "Abrir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   432
      Left            =   4200
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   762
      Picture         =   "formPDF.frx":1998
      Caption         =   "Abrir Com"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdGeraPorPreview 
      Height          =   552
      Left            =   5400
      TabIndex        =   23
      Top             =   3240
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   974
      Picture         =   "formPDF.frx":1F32
      Caption         =   "Cria Pdf   (Preview)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdAbrirPdf 
      Height          =   432
      Left            =   2040
      TabIndex        =   24
      Top             =   4320
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   762
      Picture         =   "formPDF.frx":24CC
      Caption         =   "Abrir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.CommandButtonW btnclose 
      Height          =   612
      Left            =   9000
      TabIndex        =   25
      Top             =   240
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "formPDF.frx":2A66
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdimp 
      Height          =   612
      Left            =   5040
      TabIndex        =   26
      Top             =   4320
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Imprimir"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "formPDF.frx":3000
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdVisua 
      Height          =   612
      Left            =   3840
      TabIndex        =   27
      Top             =   4320
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Visualizar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "formPDF.frx":359A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdPreviewOrigem 
      Height          =   612
      Left            =   1920
      TabIndex        =   28
      Top             =   3240
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Visualizar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "formPDF.frx":3B34
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdSendMail 
      Height          =   612
      Left            =   6240
      TabIndex        =   29
      Top             =   4320
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Email"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "formPDF.frx":40CE
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
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
      Height          =   228
      Left            =   120
      TabIndex        =   12
      Top             =   4080
      Width           =   1212
   End
   Begin VB.Label lblFilename 
      Caption         =   "Arquivo para converter"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
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
Dim Obj As Long
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
  If FileConnExist(txtFilename.Text, True) Then
    Call OpenWith(txtFilename.Text, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, CLng(Me.hwnd))
  End If
End Sub

Private Sub CmdAbrirPdf_Click()
  Dim cARQSHELL As String
  cARQSHELL = FixStr(txtOutputFile.Text)
  If FileConnExist(cARQSHELL) Then
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , CLng(Me.hwnd)
  End If
End Sub

Private Sub CmdGeraPorPreview_Click()
 If PrintPreview1.PrinterExists("Microsoft Print to PDF") Then
        PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
        If PrintPreview1.Canceled Then
            MsgBox "Error, could not save the file, check if to are allowed to write at that location", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "'Microsoft Print to PDF' driver is not installed.", vbExclamation
        Exit Sub
    End If
btnfalse
End Sub

Private Sub cmdimp_Click()
  Dim cARQSHELL As String
  cARQSHELL = FixStr(txtOutputFile.Text)
  If FileConnExist(cARQSHELL) Then
    ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , "print", CLng(Me.hwnd)
  End If
End Sub

Private Sub cmdPreviewOrigem_Click()
  PrintPreview1.ShowPreview
End Sub

Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  MyPrintingTXT
End Sub

Private Sub CmdSendMail_Click()
  Dim cARQUIVO As String
  cARQUIVO = FixStr(txtOutputFile.Text)
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
 cARQRTF = FixStr(txtOutputFile.Text)
 FrmPreview.Show vbModal, Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub
Private Sub Form_Load()

  CenterFormToScreen Me
  cmbFont.ListIndex = 1                        ' 10
  cmbFontSize.ListIndex = 0
  cmbRotation.ListIndex = 0
  cmbPageSize.ListIndex = 0

  txtCreator.Text = zNOMEFOLHA



  If Len(ePASS01) > 0 Then
    If FileConnExist(ePASS01) Then
      txtTitle.Text = NomeArq(ePASS01, True)
      txtFilename.Text = ePASS01
      txtOutputFile.Text = TrocaExt(ePASS01, "PDF")
      btnOpen.Enabled = False
      txtFilename.Enabled = False
    End If
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
  CmdVisua.Enabled = False
  CmdSendMail.Enabled = False
  cmdimp.Enabled = False
  CmdAbrirPdf.Enabled = False


  
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub txtAuthor_GotFocus()
  txtAuthor.SelStart = 0
  txtAuthor.SelLength = Len(txtAuthor.Text)
End Sub
Private Sub txtCreator_GotFocus()
  txtCreator.SelStart = 0
  txtCreator.SelLength = Len(txtCreator.Text)
End Sub

Private Sub txtSubject_GotFocus()
  txtSubject.SelStart = 0
  txtSubject.SelLength = Len(txtSubject.Text)
End Sub

Private Sub txtTitle_GotFocus()
  txtTitle.SelStart = 0
  txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub txtKeywords_GotFocus()
  txtKeywords.SelStart = 0
  txtKeywords.SelLength = Len(txtKeywords.Text)
End Sub

Private Sub txtFilename_GotFocus()
  txtFilename.SelStart = 0
  txtFilename.SelLength = Len(txtFilename.Text)
End Sub

Private Sub txtOutputFile_GotFocus()
  txtOutputFile.SelStart = 0
  txtOutputFile.SelLength = Len(txtOutputFile.Text)
End Sub

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnOpen_Click()
  On Local Error Resume Next
  txtFilename.Text = OpenArqExt(Me, txtFilename.Text, "txt", "Arquivos txt")

  If Not Len(txtFilename.Text) = 0 Then
    txtOutputFile.Text = TrocaExt(txtFilename.Text, "pdf")
  End If


End Sub

Private Sub btnSave_Click()
  Dim sFILTER As String
  Dim filename As String
  sFILTER = "Arquivos PDF (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  filename = FileSave(Me, sFILTER, 1, "PDF", cARQRTF, cARQRTF, "Salvar PDF")
  If Not Len(filename) = 0 Then
    txtOutputFile.Text = filename
  End If
End Sub

Private Sub btnSource_Click()
  On Local Error Resume Next
End Sub
Private Sub btnfalse()
      CmdVisua.Enabled = True
       CmdSendMail.Enabled = True
       cmdimp.Enabled = True
       CmdAbrirPdf.Enabled = True
       btnConvert.Enabled = False
       btnSave.Enabled = False
       CmdGeraPorPreview.Enabled = False
       txtOutputFile.Enabled = False
End Sub
Private Sub btnConvert_Click()
' Desativa os controles para proteção visual durante o processo
   ' cmdGerar.Enabled = False
    Screen.MousePointer = vbHourglass

    ' A SUA CHAMADA ORIGINAL FICA ASSIM, PERFEITA:
    ConvertToPDF txtFilename.Text, txtOutputFile.Text, _
                 txtAuthor.Text, txtCreator.Text, txtKeywords.Text, _
                 txtSubject.Text, txtTitle.Text, _
                 cmbFont.Text, Val(cmbFontSize.Text), Val(cmbRotation.Text), _
                 Val(cmbPageSize.Text), Val(Right(cmbPageSize.Text, 3)), ChkPaisagem.Value

    ' Restaura os controles
    Screen.MousePointer = vbDefault
   ' cmdGerar.Enabled = True
End Sub
Public Sub ConvertToPDF(ByVal strSourceFile As String, _
                         ByVal strOutputFile As String, _
                         Optional ByVal strAuthor As String = "", _
                         Optional ByVal strCreator As String = "", _
                         Optional ByVal strKeywords As String = "", _
                         Optional ByVal strSubject As String = "", _
                         Optional ByVal strTitle As String = "", _
                         Optional ByVal strFontName As String = "Arial", _
                         Optional ByVal sngFontSize As Single = 10, _
                         Optional ByVal lRotation As Long = 0, _
                         Optional ByVal lPageSizeCode As Long = 0, _
                         Optional ByVal lPageSizeVal As Long = 0, _
                         Optional ByVal iPaisagemValue As Integer = 0)

    Dim fso As Object
    Dim streamIn As Object
    Dim cLINHA As String
    Dim pdf As ClsFPDF
    
    ' Variáveis de paginação e controle vertical
    Dim nMargemEsquerda As Single
    Dim nLinhaAtual As Single
    Dim nLimiteInferior As Single
    Dim nAlturaLinha As Single
    Dim cOrientacao As String
    
    ' Validações padrão que você já possui
    If Not FileConnExist(strSourceFile, True) Then Exit Sub
    If Len(strOutputFile) = 0 Then strOutputFile = TrocaExt(strSourceFile, "PDF")
    
    If FileConnExist(strOutputFile, False) Then
        Alert ("Arquivo Destino Ja existe")
        Exit Sub
    End If
    
    On Error GoTo TrataErro
    
    ' 1. Determina a orientação ("L" para Paisagem, "P" para Retrato)
    ' No VB6, um CheckBox marcado (vbChecked) tem o valor 1
    If iPaisagemValue = 1 Then
        cOrientacao = "L"
        nLimiteInferior = 190 ' Limite em milímetros para folha deitada (A4 menor eixo)
    Else
        cOrientacao = "P"
        nLimiteInferior = 275 ' Limite padrão para folha em pé
    End If
    
    ' 2. Abre o arquivo texto via FSO (Streaming rápido por linha)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(strSourceFile, 1, False)
    
    ' 3. Inicializa o motor da classe ClsFPDF
    Set pdf = New ClsFPDF
    pdf.CreatePDF cOrientacao, "mm"
    pdf.AddPage
    
    ' Aplica a fonte e tamanho capturados do Combo da sua tela
    pdf.SetFont strFontName, sngFontSize
    
    ' 4. Configuração das margens e espaçamento proporcional dinâmico
    nMargemEsquerda = 15
    nLinhaAtual = 15
    
    ' Calcula a altura ideal de cada linha para o texto não sobrepor
    nAlturaLinha = (sngFontSize * 1.3) / (72 / 25.4)
    
    ' 5. Loop de leitura e escrita
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' Desenha o texto no PDF
        pdf.Cell cLINHA, nMargemEsquerda, nLinhaAtual
        nLinhaAtual = nLinhaAtual + nAlturaLinha
        
        ' Se estourar a margem inferior, quebra a página mantendo a fonte definida
        If nLinhaAtual > nLimiteInferior Then
            pdf.AddPage
            pdf.SetFont strFontName, sngFontSize
            nLinhaAtual = 15
        End If
    Loop
    
    ' 6. Salva o arquivo finalizado
    pdf.Output strOutputFile
    
    ' Opcional: Feedback que seu sistema costuma dar
    ' Alert "Conversão concluída com sucesso!"

Fim:
    If Not streamIn Is Nothing Then streamIn.Close
    Set streamIn = Nothing
    Set fso = Nothing
    Set pdf = Nothing
    Exit Sub

TrataErro:
    Alert "Erro na conversão para PDF: " & Err.Description
    Resume Fim
End Sub
Public Sub MyPrintingTXT()
  Dim fso As Object
  Dim streamIn As Object
  Dim strBuffer As String
  Dim cARQOPEN As String
  
  cARQOPEN = txtFilename.Text
  If Not FileConnExist(cARQOPEN, True) Then
    Exit Sub
  End If
  
  On Error GoTo TrataErro

  ' SUBSTITUIÇÃO DO OPEN: Inicializa o fluxo contínuo via FSO
  Set fso = CreateObject("Scripting.FileSystemObject")
  Set streamIn = fso.OpenTextFile(cARQOPEN, 1, False)
  
  ' SUBSTITUIÇÃO DO DO WHILE NOT EOF: Loop usando o Stream
  Do While Not streamIn.AtEndOfStream
    
    ' SUBSTITUIÇÃO DO INPUT: Pega a linha a linha na mesma variável STRBUFFER
    strBuffer = streamIn.ReadLine
    
    ' MANTIDO EXATAMENTE IGUAL: Envia a linha para a impressão/componente
    Printer.Print strBuffer
    
  Loop
  
Fim:
  ' SUBSTITUIÇÃO DO CLOSE: Fecha o arquivo e limpa a memória de forma segura
  If Not streamIn Is Nothing Then streamIn.Close
  Set streamIn = Nothing
  Set fso = Nothing
  Exit Sub

TrataErro:
  MsgBox "Erro na leitura do arquivo para impressão: " & Err.Description, vbCritical
  Resume Fim
End Sub

