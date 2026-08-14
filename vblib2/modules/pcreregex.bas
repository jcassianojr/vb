Option Explicit

' --- DeclaraÃ§Ãµes da API C da pcre2-8.dll ---
#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function pcre2_compile_8 Lib "pcre2-8.dll" Alias "pcre2_compile_8" (ByVal pattern As LongPtr, ByVal length As LongPtr, ByVal options As Long, ByRef errorcode As Long, ByRef erroroffset As LongPtr, ByVal ccontext As LongPtr) As LongPtr
    Public Declare PtrSafe Sub pcre2_code_free_8 Lib "pcre2-8.dll" Alias "pcre2_code_free_8" (ByVal code As LongPtr)
    Public Declare PtrSafe Function pcre2_match_8 Lib "pcre2-8.dll" Alias "pcre2_match_8" (ByVal code As LongPtr, ByVal subject As LongPtr, ByVal length As LongPtr, ByVal startoffset As LongPtr, ByVal options As Long, ByVal match_data As LongPtr, ByVal mcontext As LongPtr) As Long
    Public Declare PtrSafe Function pcre2_match_data_create_from_pattern_8 Lib "pcre2-8.dll" Alias "pcre2_match_data_create_from_pattern_8" (ByVal code As LongPtr, ByVal gcontext As LongPtr) As LongPtr
    Public Declare PtrSafe Sub pcre2_match_data_free_8 Lib "pcre2-8.dll" Alias "pcre2_match_data_free_8" (ByVal match_data As LongPtr)
#Else
    Public Declare Function pcre2_compile_8 Lib "pcre2-8.dll" Alias "pcre2_compile_8" (ByVal pattern As Long, ByVal length As Long, ByVal options As Long, ByRef errorcode As Long, ByRef erroroffset As Long, ByVal ccontext As Long) As Long
    Public Declare Sub pcre2_code_free_8 Lib "pcre2-8.dll" Alias "pcre2_code_free_8" (ByVal code As Long)
    Public Declare Function pcre2_match_8 Lib "pcre2-8.dll" Alias "pcre2_match_8" (ByVal code As Long, ByVal subject As Long, ByVal length As Long, ByVal startoffset As Long, ByVal options As Long, ByVal match_data As Long, ByVal mcontext As Long) As Long
    Public Declare Function pcre2_match_data_create_from_pattern_8 Lib "pcre2-8.dll" Alias "pcre2_match_data_create_from_pattern_8" (ByVal code As Long, ByVal gcontext As Long) As Long
    Public Declare Sub pcre2_match_data_free_8 Lib "pcre2-8.dll" Alias "pcre2_match_data_free_8" (ByVal match_data As Long)
#End If

Private Const PCRE2_ZERO_TERMINATED As LongPtr = -1

' ==============================================================================
' FunÃ§Ã£o substituta usando diretamente a pcre2-8.dll compilada por vocÃª
' ==============================================================================
Public Function CheckRegEx(ByVal texto As String, ByVal padrao As String) As Boolean
    Dim errCode As Long
    Dim errOffset As LongPtr
    Dim pCode As LongPtr
    Dim pMatchData As LongPtr
    Dim matchResult As Long
    
    CheckRegEx = False
    
    ' 1. Compila o padrÃ£o da expressÃ£o regular usando a DLL do PCRE2
    pCode = pcre2_compile_8(StrPtr(padrao), PCRE2_ZERO_TERMINATED, 0, errCode, errOffset, 0&)
    
    If pCode = 0 Then Exit Function ' Erro de sintaxe na expressÃ£o regular
    
    ' 2. Cria o bloco de dados para o resultado da busca baseado no padrÃ£o compilado
    pMatchData = pcre2_match_data_create_from_pattern_8(pCode, 0&)
    
    If pMatchData <> 0 Then
        ' 3. Executa a validaÃ§Ã£o (match) do texto contra o padrÃ£o compilado
        matchResult = pcre2_match_8(pCode, StrPtr(texto), Len(texto), 0&, 0&, pMatchData, 0&)
        
        ' Se o resultado for maior ou igual a 0, encontrou correspondÃªncia
        If matchResult >= 0 Then
            CheckRegEx = True
        End If
        
        ' Libera o bloco de dados da memÃ³ria
        pcre2_match_data_free_8 pMatchData
    End If
    
    ' 4. Libera o ponteiro do cÃ³digo compilado da memÃ³ria
    pcre2_code_free_8 pCode
End Function