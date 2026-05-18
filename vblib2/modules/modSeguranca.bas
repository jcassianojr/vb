Attribute VB_Name = "modSeguranca"
Option Explicit

' --- DECLARAÃ‡Ã•ES DA WINDOWS CRYPTOAPI (NATIVA) ---
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal cbData As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long

' APIs para manipulaÃ§Ã£o do arquivo config.dat
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
' Constantes da CryptoAPI
Private Const PROV_RSA_FULL As Long = 1
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_MD5 As Long = 32771
Private Const CALG_RC2 As Long = 26114

' Chave Mestra para derivar a chave de criptografia
Private Const CHAVE_MESTRA As String = "Diretoria@Segura#2026!bancos"

Public Function CaminhoArquivoCofre() As String
    Dim Caminho As String
    Caminho = App.Path
    If Right$(Caminho, 1) <> "\" Then Caminho = Caminho & "\"
    CaminhoArquivoCofre = Caminho & "config.dat"
End Function

' --- FUNÃ‡ÃƒO PARA CRIPTOGRAFAR (CryptoAPI) ---
Public Function CriptografarForte(ByVal TextoPuro As String) As String
    If TextoPuro = "" Then Exit Function
    
    Dim hProv As Long, hHash As Long, hKey As Long
    Dim strBuffer As String
    Dim LngDataLen As Long, LngBufLen As Long
    Dim i As Long, StrHex As String
    
    ' 1. Adquire o contexto criptogrÃ¡fico do Windows
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        ' 2. Cria um hash MD5 para derivar a chave a partir da nossa string mestra
        If CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, CHAVE_MESTRA, Len(CHAVE_MESTRA), 0) <> 0 Then
                ' 3. Deriva uma chave RC2 estÃ¡vel baseada no Hash da chave mestra
                If CryptDeriveKey(hProv, CALG_RC2, hHash, 0, hKey) <> 0 Then
                    
                    ' Prepara o buffer preenchendo com espaÃ§os (mÃºltiplo de 8 bytes para cifra de bloco)
                    LngDataLen = Len(TextoPuro)
                    LngBufLen = ((LngDataLen \ 8) + 1) * 8
                    strBuffer = TextoPuro & Space$(LngBufLen - LngDataLen)
                    
                    ' 4. Executa a criptografia nativa do Windows
                    If CryptEncrypt(hKey, 0, 1, 0, strBuffer, LngDataLen, LngBufLen) <> 0 Then
                        ' Converte os caracteres resultantes para Hexadecimal limpo
                        For i = 1 To LngDataLen
                            StrHex = StrHex & Right$("00" & Hex$(Asc(Mid$(strBuffer, i, 1))), 2)
                        Next i
                        CriptografarForte = StrHex
                    End If
                    
                    CryptDestroyKey hKey
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
End Function

' --- FUNÃ‡ÃƒO PARA DESCRIPTOGRAFAR (CryptoAPI) ---
Public Function DescriptografarForte(ByVal TextoCriptografado As String) As String
    If TextoCriptografado = "" Then Exit Function
    
    Dim hProv As Long, hHash As Long, hKey As Long
    Dim strBuffer As String
    Dim LngDataLen As Long
    Dim i As Long, ByteChar As String
    
    ' Reverte a string Hexadecimal de volta para caracteres binÃ¡rios no buffer
    For i = 1 To Len(TextoCriptografado) Step 2
        ByteChar = Chr$(Val("&H" & Mid$(TextoCriptografado, i, 2)))
        strBuffer = strBuffer & ByteChar
    Next i
    LngDataLen = Len(strBuffer)
    
    ' 1. Adquire o contexto criptogrÃ¡fico do Windows
    If CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT) <> 0 Then
        ' 2. Recria o mesmo hash MD5 da chave mestra
        If CryptCreateHash(hProv, CALG_MD5, 0, 0, hHash) <> 0 Then
            If CryptHashData(hHash, CHAVE_MESTRA, Len(CHAVE_MESTRA), 0) <> 0 Then
                ' 3. Deriva a mesma chave RC2
                If CryptDeriveKey(hProv, CALG_RC2, hHash, 0, hKey) <> 0 Then
                    
                    ' 4. Executa a descriptografia nativa do Windows
                    If CryptDecrypt(hKey, 0, 1, 0, strBuffer, LngDataLen) <> 0 Then
                        ' Retorna a string limpa (removendo os espaÃ§os de preenchimento do bloco)
                        DescriptografarForte = Left$(strBuffer, LngDataLen)
                    End If
                    
                    CryptDestroyKey hKey
                End If
            End If
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
End Function

' --- INTERFACE DE LEITURA E GRAVAÃ‡ÃƒO NO ARQUIVO ---
Public Sub GravarNoCofre(ByVal SecaoBanco As String, ByVal Chave As String, ByVal Valor As String)
    Dim ValorCripto As String
    ValorCripto = CriptografarForte(Valor)
    Call WritePrivateProfileString(SecaoBanco, Chave, ValorCripto, CaminhoArquivoCofre)
End Sub

Public Function LerDoCofre(ByVal SecaoBanco As String, ByVal Chave As String) As String
    Dim Retorno As String * 255
    Dim QtdCaracteres As Long
    QtdCaracteres = GetPrivateProfileString(SecaoBanco, Chave, "", Retorno, 255, CaminhoArquivoCofre)
    Dim ValorCripto As String
    ValorCripto = Left$(Retorno, QtdCaracteres)
    LerDoCofre = DescriptografarForte(ValorCripto)
End Function

' Adicione esta declaração no topo


' Adicione esta função ao final do módulo
Public Sub ListarBancosNoCombo(ByRef Combo As Object)
    Dim strBuffer As String * 1024
    Dim lngLen As Long
    Dim vBancos As Variant
    Dim i As Integer
    
    Combo.Clear
    lngLen = GetPrivateProfileSectionNames(strBuffer, 1024, CaminhoArquivoCofre)
    
    If lngLen > 0 Then
        ' A API retorna as seções separadas por um caractere nulo (Chr(0))
        vBancos = Split(Left$(strBuffer, lngLen - 1), Chr$(0))
        For i = LBound(vBancos) To UBound(vBancos)
            If vBancos(i) <> "" Then Combo.AddItem vBancos(i)
        Next i
    End If
End Sub

Public Sub ExcluirBanco(ByVal NomeSecao As String)
    ' No Windows API, enviar vbNullString para a seção exclui a seção inteira
    Call WritePrivateProfileString(NomeSecao, vbNullString, vbNullString, CaminhoArquivoCofre)
End Sub

