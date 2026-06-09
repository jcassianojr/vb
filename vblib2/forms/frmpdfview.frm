VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Object = "{C7880F7C-F318-471C-8BB7-25660C4BE24A}#1.1#0"; "OrdoPDFReader.v3.ocx"
Begin VB.Form frmPDFView 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "PDFReader"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   13092
   Icon            =   "frmpdfview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   13092
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   732
      Left            =   11400
      TabIndex        =   1
      Top             =   240
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1291
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmpdfview.frx":038A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin OrdoPDF_Reader_3.OrdoPDFReaderV3 PDFReader1 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11145
      _ExtentX        =   19664
      _ExtentY        =   12086
      BorderStyle     =   1
      LanguageUsed    =   "auto"
   End
End
Attribute VB_Name = "frmPDFView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bPrimeiraCarga As Boolean

Private Sub Encerrar_Click()
 Screen.MousePointer = vbDefault
  Unload Me
End Sub

'Private Sub Command1_Click()
'  cARQRTF = "C:/temp/citacao.pdf"
'  PDFReader1.Load cARQRTF
'  PDFReader1.FitControl
'End Sub
 
 Private Sub Form_Activate()
    ' Se não for a primeira vez que a tela ganha foco (ex: o usuário minimizou e maximizou),
    ' ignora para não recarregar o PDF à toa.
    If Not m_bPrimeiraCarga Then Exit Sub
    m_bPrimeiraCarga = False
    
    On Error GoTo TrataErro
    
    ' Dá um pequeno respiro de milissegundos para o Windows estabilizar os ponteiros da DLL
    DoEvents
    
    ' 1. Validação básica de preenchimento do caminho do arquivo
    If Len(Trim(cARQRTF)) = 0 Then
        MsgBox "Nenhum arquivo PDF foi especificado para visualização.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' 2. Verifica se o arquivo gerado pela ClsFPDF está pronto no disco
    If Dir(cARQRTF) = "" Then
        MsgBox "O arquivo PDF não foi localizado no caminho:" & vbCrLf & cARQRTF, vbCritical, "Erro de Arquivo"
        Exit Sub
    End If
    
    ' 3. BLINDAGEM CONTRA ARQUIVOS FALSOS: Testa a assinatura binária do PDF
    If Not IsValidPDF(cARQRTF) Then
        MsgBox "O arquivo selecionado não é um documento PDF válido!" & vbCrLf & _
               "Pode tratar-se de um arquivo de texto (TXT) ou documento (DOC) renomeado indevidamente.", _
               vbCritical, "Erro de Validação de Formato"
        Exit Sub
    End If
    
    ' 3. Agora sim, com o componente 100% carregado e ativo, faz o Load seguro
    Screen.MousePointer = vbHourglass
    PDFReader1.Load cARQRTF
    PDFReader1.FitControl
    Screen.MousePointer = vbDefault
    
    Exit Sub

TrataErro:
    Screen.MousePointer = vbDefault
    MsgBox "Erro ao carregar o visualizador de PDF: " & Err.Description, vbCritical, "Erro de Inicialização"
End Sub
   
Private Sub Form_Load()
  'da erro quando load aqui
  'da erro quando chama command1_click
  'deixando inicialmente para o usuario carregar o pdf
  m_bPrimeiraCarga = True
  'PDFReader1.Load cARQRTF
  'PDFReader1.FitControl
End Sub

'End Sub

'Private Sub Command2_Click()
'    PDFReader1.SelectPDFFile App.Path
'End Sub

'Private Sub Command3_Click()
'    PDFReader1.FitControl
'End Sub

'Private Sub Command4_Click()
'    PDFReader1.Zoom = 100
'End Sub

'Private Sub Command5_Click()
'    Dim Numero As Integer
'    Numero = Val(InputBox("Go to the page (1 à " & PDFReader1.GetPagesCount & ")", "Enter page number"))
    
'    If Numero > 0 Then PDFReader1.DisplayedPage = Numero
'End Sub

'Private Sub Command6_Click()
'    PDFReader1.SelectPDFFileForOCR
'End Sub

'Private Sub Command7_Click()
'    PDFReader1.PrintPDF , , 1, 1
'End Sub

'Private Sub Form_Load()
 'PDFReader1.Load cARQRTF
' Command1_Click
'End Sub

'Private Sub Command8_Click()
'    PDFReader1.CopyPageToClipboard
'End Sub

Private Sub Form_Resize()
   On Error Resume Next
   PDFReader1.Height = Me.ScaleHeight
   PDFReader1.Width = Me.ScaleWidth - PDFReader1.Left - 1700 '1700 espacaos dos comandos a direita
End Sub
Private Sub PDFReader1_PageChanged(PageViewed As Integer)
    Debug.Print PageViewed
End Sub
'Private Sub Option1_Click(Index As Integer)
'    If Option1(0).Value Then PDFReader1.IsToolbarVisible = True
'    If Option1(1).Value Then PDFReader1.IsToolbarVisible = False
'    If Option1(2).Value Then PDFReader1.IsStatusBarVisible = False
'    If Option1(3).Value Then PDFReader1.IsStatusBarVisible = True
'    If Option1(4).Value Then PDFReader1.IsPDFButtonVisible = True
'    If Option1(5).Value Then PDFReader1.IsPDFButtonVisible = False
'End Sub
Private Sub PDFReader1_PDFLoaded(filename As String, FilePath As String)
    Debug.Print filename, FilePath
    Me.Caption = "PDFReader" & " : " & filename & " (" & FilePath & ")"
End Sub

'Private Sub Picture1_Click(Index As Integer)
'    PDFReader1.BackColor = Picture1(Index).BackColor
'End Sub
Public Function IsValidPDF(ByVal strFilePath As String) As Boolean
    Dim fNum As Integer
    Dim strHeader As String * 4
    
    ' Se o arquivo não existir fisicamente, já aborta
    If Dir(strFilePath) = "" Then
        IsValidPDF = False
        Exit Function
    End If
    
    On Error GoTo Erro
    
    ' Abre o arquivo em modo binário puro para leitura rápida
    fNum = FreeFile
    Open strFilePath For Binary Access Read As #fNum
    
    ' Lê apenas os primeiros 4 bytes do arquivo
    Get #fNum, 1, strHeader
    Close #fNum
    
    ' Verifica se o cabeçalho começa com a assinatura mundial do PDF: "%PDF"
    If strHeader = "%PDF" Then
        IsValidPDF = True
    Else
        IsValidPDF = False
    End If
    
    Exit Function

Erro:
    ' Se der erro ao abrir (arquivo bloqueado, etc), assume falso por segurança
    If fNum > 0 Then Close #fNum
    IsValidPDF = False
End Function
