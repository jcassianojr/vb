Attribute VB_Name = "ModBackupPDFOriginal"
Option Explicit

' =========================================================================
' VARIÁVEIS E MATRIZES GLOBAIS DO MOTOR DE IMPRESSÃO ORIGINAL (BACKUP)
' =========================================================================
Public location(1 To 50000) As Long
Public pageObj(1 To 5000) As Long
Public pageNo As Long
Public Obj As Long
Public Position As Long
Public stream_len As Long
Public start_stream As Long
Public end_stream As Long
Public title As String
Public author As String
Public root As Long
Public info As Long
Public pages As Long

' =========================================================================
' ROTINAS ORIGINAIS EXTRAÍDAS DO FORMPDF.FRM
' =========================================================================

Public Sub writepdf(ByVal str As String, Optional ByVal ln As Boolean = False)
    On Error GoTo er
    ' Simulação/Mantenedor da escrita sequencial binária original
    ' Esta sub controlava a escrita física medindo a propriedade Position
    Position = Position + Len(str) + 2
    Exit Sub
er:
    MsgBox "Erro no writepdf (Backup): " & Err.Description
End Sub

Public Sub setobject()
    On Error GoTo er
    Obj = Obj + 1
    location(Obj) = Position
    writepdf (Obj & " 0 obj")
    Exit Sub
er:
    MsgBox "Erro no setobject (Backup): " & Err.Description
End Sub

Public Sub pdfHeader()
    On Error GoTo er
    Position = 0
    Obj = 0
    pageNo = 0
    writepdf ("%PDF-1.4")
    
    ' Catálogo Root
    setobject
    root = Obj
    writepdf ("<<")
    writepdf ("/Type /Catalog")
    writepdf ("/Pages 2 0 R")
    writepdf (">>")
    writepdf ("endobj")
    
    ' Árvore de páginas (Objeto 2)
    Obj = Obj + 1
    pages = Obj
    location(Obj) = Position
    writepdf (Obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Pages")
    Exit Sub
er:
    MsgBox "Erro no pdfHeader (Backup): " & Err.Description
End Sub

Public Sub pdfFooter()
    On Error GoTo er
    Dim ty As String
    Dim i As Long
    Dim xreF As Long
    
    ' Fecha o Objeto Info de Metadados original
    setobject
    info = Obj
    writepdf ("<<")
    writepdf ("/Title (" & title & ")")
    writepdf ("/Author (" & author & ")")
    writepdf ("/Creator (Visual Basic 6.0)")
    writepdf ("/Producer (Pure VB6 PDF Engine)")
    writepdf (">>")
    writepdf ("endobj")
    
    ' Atualiza a árvore estrutural de páginas (Objeto 2)
    location(pages) = Position
    writepdf (pages & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Pages")
    writepdf ("/Count " & pageNo)
    
    ty = ("/Kids [ ")
    For i = 1 To pageNo
        ty = ty & pageObj(i) & " 0 R "
    Next i
    ty = ty & "]"
    
    writepdf (ty)
    writepdf (">>")
    writepdf ("endobj")
    
    ' Escrita da tabela de referências cruzadas (xref)
    xreF = Position
    writepdf ("0 " & Obj + 1)
    writepdf ("0000000000 65535 f ")
    For i = 1 To Obj
        writepdf (Format(location(i), "0000000000") & " 00000 n ")
    Next i
    
    ' Trailer de encerramento do arquivo
    writepdf ("trailer")
    writepdf ("<<")
    writepdf ("/Size " & Obj + 1)
    writepdf ("/Root " & root & " 0 R")
    writepdf ("/Info " & info & " 0 R")
    writepdf (">>")
    writepdf ("startxref")
    writepdf (xreF)
    writepdf "%%EOF", True
    Exit Sub
er:
    MsgBox "Erro no pdfFooter (Backup): " & Err.Description
End Sub

Public Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
    Dim mtext As String
    Dim SpacePos As Long
    mtext = Text
    SpacePos = InStr(mtext, TextToReplace)
    Do While SpacePos
        mtext = Left(mtext, SpacePos - 1) & NewText & Right(mtext, Len(mtext) - SpacePos - Len(TextToReplace) + 1)
        SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
    Loop
    ReplaceText = mtext
End Function
