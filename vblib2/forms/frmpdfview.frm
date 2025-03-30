VERSION 5.00
Object = "{C7880F7C-F318-471C-8BB7-25660C4BE24A}#1.1#0"; "OrdoPDFReader.v3.ocx"
Begin VB.Form frmPDFView 
   Caption         =   "PDFReader"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12945
   Icon            =   "frmpdfview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   12945
   StartUpPosition =   2  'CenterScreen
   Begin OrdoPDF_Reader_3.OrdoPDFReaderV3 PDFReader1 
      Height          =   6855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12091
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

'Private Sub Command1_Click()
'  cARQRTF = "C:/temp/citacao.pdf"
'  PDFReader1.Load cARQRTF
'  PDFReader1.FitControl
'End Sub
 
   
Private Sub Form_Load()
  'da erro quando load aqui
  'da erro quando chama command1_click
  'deixando inicialmente para o usuario carregar o pdf
  PDFReader1.Load cARQRTF
  PDFReader1.FitControl
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

'    On Error Resume Next
   PDFReader1.Height = Me.ScaleHeight
   PDFReader1.Width = Me.ScaleWidth - PDFReader1.Left - 1700 '1700 espacaos dos comandos a direita
End Sub

'Private Sub Option1_Click(Index As Integer)
'    If Option1(0).Value Then PDFReader1.IsToolbarVisible = True
'    If Option1(1).Value Then PDFReader1.IsToolbarVisible = False
'    If Option1(2).Value Then PDFReader1.IsStatusBarVisible = False
'    If Option1(3).Value Then PDFReader1.IsStatusBarVisible = True
'    If Option1(4).Value Then PDFReader1.IsPDFButtonVisible = True
'    If Option1(5).Value Then PDFReader1.IsPDFButtonVisible = False
'End Sub

Private Sub PDFReader1_PageChanged(PageViewed As Integer)
    Debug.Print PageViewed
End Sub

Private Sub PDFReader1_PDFLoaded(FileName As String, FilePath As String)
    Debug.Print FileName, FilePath
    Me.Caption = "PDFReader" & " : " & FileName & " (" & FilePath & ")"
End Sub

'Private Sub Picture1_Click(Index As Integer)
'    PDFReader1.BackColor = Picture1(Index).BackColor
'End Sub
