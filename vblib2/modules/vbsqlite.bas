Attribute VB_Name = "vbsqlite"
Option Explicit
' Reference=*\G{7CC1A5F1-A0FF-4546-A0F1-FBFE744A4522}#1.1#0#..\..\..\..\..\..\WINDOWS\system32\VBSQLite12.DLL#VB SQLite Library 1.2
' Notas de Implementação:
' 1. Requer as classes cSQLiteConnection e cSQLiteCursor do projeto VBSQLite.
' 2. Mantém compatibilidade com a função SQLDialeto do adolib.bas.

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: ADOComando
'---------------------------------------------------------------------------------------
Public Function SQLiteComando(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim loConn As New SQLiteConnection
    On Error GoTo Erro
    
    ' Limpeza da Tag e tradução do dialeto
    cCON = LimpaTag(cCON)
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB cCON
    loConn.Execute cSQL
    
    SQLiteComando = True
    loConn.CloseDB
    Exit Function
Erro:
    SQLiteComando = False
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegSQLAdo[cite: 2]
'---------------------------------------------------------------------------------------
Public Function PegSQLite(ByVal cCON As String, ByVal cSQL As String, ByRef aDados As Variant) As Boolean
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim i As Integer
    
    On Error GoTo Erro
    
    cCON = LimpaTag(cCON)
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB cCON
    Set loCursor = loConn.CreateCursor(cSQL)
    
    If Not loCursor.EOF Then
        ReDim aDados(loCursor.columnCount - 1)
        For i = 0 To loCursor.columnCount - 1
            ' FVar é sua função de tratamento de dados[cite: 2]
            aDados(i) = FVar(loCursor.Value(i))
        Next
        PegSQLite = True
    Else
        PegSQLite = False
    End If
    
    loConn.CloseDB
    Exit Function
Erro:
    PegSQLite = False
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: GrvSQLado[cite: 2]
'---------------------------------------------------------------------------------------
Public Function GrvSQLite(ByVal cCON As String, ByVal cSQL As String, ByRef aDados As Variant) As Boolean
    Dim loConn As New cSQLiteConnection
    Dim loCursor As cSQLiteCursor
    Dim i As Integer
    
    On Error GoTo Erro
    
    cCON = LimpaTag(cCON)
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB cCON
    ' Abre um cursor editável (Updateable)
    Set loCursor = loConn.CreateCursor(cSQL)
    
    If Not loCursor.EOF Then
        For i = 0 To loCursor.columnCount - 1
            ' Simula o preenchimento de campos do ADO[cite: 2]
            loCursor.Value(i) = aDados(i)
        Next
        ' No VBSQLite, a gravação pode exigir um comando UPDATE manual ou método da classe
        ' Dependendo da versão da classe, usamos Execute:
        loConn.Execute GerarUpdateManual(cSQL, aDados, loCursor)
        GrvSQLite = True
    End If
    
    loConn.CloseDB
    Exit Function
Erro:
    GrvSQLite = False
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: IncluiSQLAdo[cite: 2]
'---------------------------------------------------------------------------------------
Public Function IncluiSQLite(ByVal cCON As String, ByVal cTABELA As String, ByVal aCAMPOS As Variant, ByVal aVALORES As Variant, Optional ByRef nID As Long) As Boolean
    Dim loConn As New cSQLiteConnection
    Dim cSQL As String
    Dim i As Integer
    
    On Error GoTo Erro
    
    cCON = LimpaTag(cCON)
    
    ' Montagem do INSERT manual (mais seguro para SQLite Nativo)
    cSQL = "INSERT INTO " & cTABELA & " ("
    For i = LBound(aCAMPOS) To UBound(aCAMPOS)
        cSQL = cSQL & aCAMPOS(i) & IIf(i < UBound(aCAMPOS), ",", "")
    Next
    cSQL = cSQL & ") VALUES ("
    For i = LBound(aVALORES) To UBound(aVALORES)
        cSQL = cSQL & FormataParaSQL(aVALORES(i)) & IIf(i < UBound(aVALORES), ",", "")
    Next
    cSQL = cSQL & ")"
    
    loConn.OpenDB cCON
    loConn.Execute cSQL
    
    ' Captura o ID gerado automaticamente
    nID = loConn.LastInsertRowID
    
    IncluiSQLite = True
    loConn.CloseDB
    Exit Function
Erro:
    IncluiSQLite = False
End Function

'---------------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES DE SUPORTE
'---------------------------------------------------------------------------------------

Private Function LimpaTag(ByVal cCON As String) As String
    ' Remove a tag customizada para obter o path puro do arquivo
    LimpaTag = Replace(cCON, "[SQLITEDIRETO]", "")
End Function

Private Function FormataParaSQL(ByVal vValor As Variant) As String
    ' Garante que o dado vá no formato correto para o SQLite[cite: 1, 2]
    If IsNull(vValor) Or vValor = "" Then
        FormataParaSQL = "NULL"
    ElseIf IsDate(vValor) Then
        FormataParaSQL = "'" & Format(vValor, "YYYY-MM-DD HH:MM:SS") & "'"
    ElseIf IsNumeric(vValor) Then
        FormataParaSQL = Replace(CStr(vValor), ",", ".")
    Else
        ' Trata aspas simples dentro de strings[cite: 1]
        FormataParaSQL = "'" & Replace(vValor, "'", "''") & "'"
    End If
End Function

Private Function GerarUpdateManual(cSQL_Original As String, aDados As Variant, loCursor As Object) As String
    ' Função interna para converter um comando SELECT em UPDATE caso a classe não suporte .Update direto
    ' Baseado na lógica de GrvSQLado[cite: 2]
    ' ... lógica de parsing de string aqui ...
End Function
