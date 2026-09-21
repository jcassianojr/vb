# modIni.bas

Módulo para leitura e gravação de arquivos INI em VBA/VB6, com compatibilidade para 32/64 bits e helpers para trabalhar com seções, chaves, caminhos e conversões de tipos.

Arquivo principal:
- `vblib2/modules/modIni.bas`

## Objetivo

O módulo centraliza acesso a arquivos INI usando as APIs do Windows:
- `GetPrivateProfileString`
- `WritePrivateProfileString`
- `GetPrivateProfileSectionNames`

Ele oferece:
- leitura e escrita direta de valores em INI;
- funções com nomes compatíveis com códigos antigos;
- conversão automática de tipos (`String`, `Long`, `Date`, `Boolean`);
- acesso rápido a caminho e valores padrão.

---

## Declarações principais

O módulo possui compatibilidade com VBA7 e versões antigas:

```vb
#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As LongPtr, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As LongPtr, ByVal lpFileName As String) As Long
#Else
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If
```

Essas declarações permitem que o módulo funcione tanto em Office 32-bit quanto 64-bit.

---

## Funções básicas

### GetIni
Lê um valor em um arquivo INI.

```vb
Public Function GetIni(Section As String, Item As String, Default As String, IniName As String) As String
```

Exemplo:

```vb
Dim sHost As String
sHost = GetIni("DB", "SERVER", "localhost", "C:\app\config.ini")
```

### SetIni
Grava um valor em um arquivo INI.

```vb
Public Sub SetIni(Section As String, Item As String, ByVal Value As String, IniName As String)
```

Exemplo:

```vb
SetIni "DB", "SERVER", "10.0.0.15", "C:\app\config.ini"
```

---

## Helpers compatíveis com código legado

Essas funções foram mantidas para compatibilidade com chamadas antigas do projeto.

### PegINIVAL
Lê um valor do INI e retorna um valor padrão quando o retorno for vazio.

```vb
Public Function PegINIVAL(ByVal cARQINI As String, ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
```

Exemplo:

```vb
Dim sBanco As String
sBanco = PegINIVAL("C:\app\config.ini", "DB", "NOME", "default")
```

### PegPath
Lê um valor de caminho do INI, usando `App.Path` como fallback.

```vb
Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "", Optional ByVal cARQINI As String = "") As String
```

Exemplo:

```vb
Dim sCaminho As String
sCaminho = PegPath("PATH", "ARQUIVOS", App.Path & "\", "C:\app\config.ini")
```

### PegTable
Retorna um caminho de tabela / diretório a partir do arquivo INI padrão do executável.

```vb
Public Function PegTable(ByVal cGRUPO As String, ByVal cCAMPO As String)
```

Exemplo:

```vb
Dim sDir As String
sDir = PegTable("PATH", "BANCODADOS")
```

---

## Funções de alto nível (`Profile*`)

Essas funções foram pensadas para um uso mais moderno e padronizado, com conversão automática de tipos.

### ProfileString
Lê um valor string, com nome de arquivo e seção normalizados.

```vb
Public Function ProfileString(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal cDefault As Variant = "") As String
```

Se `cINIFile` não possuir extensão, o módulo adiciona `.INI` automaticamente.

Exemplo:

```vb
Dim sValor As String
sValor = ProfileString("meuapp", "DB", "SERVER", "localhost")
```

Isso tenta abrir algo como:

```text
MEUAPP.INI
```

### SetProfile
Grava valores com conversão automática de tipos.

```vb
Public Function SetProfile(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, ByVal xValue As Variant) As Boolean
```

Conversões suportadas:
- `String` → valor direto
- `Integer/Long/Single/Double/Decimal/Currency` → string numérica
- `Boolean` → `1` ou `0`
- `Date` → `yyyymmdd`

Exemplo:

```vb
Dim ok As Boolean
ok = SetProfile("meuapp", "DB", "PORT", 3306)
ok = SetProfile("meuapp", "DB", "AUTOLOGIN", True)
```

### ProfileNum
Lê valor numérico e converte para `Double`.

```vb
Public Function ProfileNum(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal nDefault As Variant = 0) As Double
```

Exemplo:

```vb
Dim nPorta As Double
nPorta = ProfileNum("meuapp", "DB", "PORT", 3306)
```

### ProfileDate
Lê um valor de data gravado como `yyyymmdd`.

```vb
Public Function ProfileDate(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal dDefault As Variant) As Date
```

Exemplo:

```vb
Dim dUltimoAcesso As Date
dUltimoAcesso = ProfileDate("config", "APP", "LAST_ACCESS", Date)
```

### ProfileLogical
Lê texto e converte para `Boolean`.

```vb
Public Function ProfileLogical(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal lDefault As Variant = False) As Boolean
```

Valores aceitos como verdadeiros:
- `S`, `SIM`, `T`, `.T.`, `TRUE`, `YES`, `ON`, `Y`, `1`

Exemplo:

```vb
Dim bAtivo As Boolean
bAtivo = ProfileLogical("meuapp", "APP", "ATIVO", True)
```

---

## Exemplo completo de uso

Arquivo INI:

```ini
[DB]
SERVER=localhost
PORT=3306
AUTOLOGIN=1

[APP]
ATIVO=TRUE
LAST_ACCESS=20260921
```

Código VB/VBA:

```vb
Option Explicit

' Importa o módulo
' Global modIni

Private Sub TestaINI()
    Dim sServer As String
    Dim nPorta As Double
    Dim bAutoLogin As Boolean
    Dim dUltimoAcesso As Date
    Dim sCaminho As String

    sServer = ProfileString("meuapp", "DB", "SERVER", "localhost")
    nPorta = ProfileNum("meuapp", "DB", "PORT", 3306)
    bAutoLogin = ProfileLogical("meuapp", "DB", "AUTOLOGIN", False)
    dUltimoAcesso = ProfileDate("meuapp", "APP", "LAST_ACCESS", Date)

    MsgBox "Servidor: " & sServer & vbCrLf & _
           "Porta: " & nPorta & vbCrLf & _
           "AutoLogin: " & bAutoLogin & vbCrLf & _
           "Ultimo acesso: " & dUltimoAcesso

    SetProfile "meuapp", "APP", "ATIVO", True
    SetProfile "meuapp", "DB", "PORT", 5432

    sCaminho = PegPath("PATH", "BASE", App.Path & "\", App.Path & "\meuapp.ini")
    MsgBox sCaminho
End Sub
```

---

## Observações importantes

### 1. Caminhos e nomes de arquivo

Algumas funções assumem arquivo padrão no diretório do executável:

```vb
App.Path & "\" & App.EXEName & ".INI"
```

Quando você usa `ProfileString`, `SetProfile`, `ProfileNum`, etc., se o nome do INI não vier com extensão, o módulo acrescenta `.INI` automaticamente.

### 2. Compatibilidade com 64 bits

O módulo usa `#If VBA7 Or Win64 Then` para ajustar as declarações da API. Isso evita erros em versões modernas do Office.

### 3. Armazenamento de datas

A função `SetProfile` usa:

```vb
Format$(xValue, "yyyymmdd")
```

e `ProfileDate` lê esse formato para recriar um `Date`.

### 4. Conversão booleana

Para acesso rápido a flags, o módulo interpreta valores como `True` quando forem:

```vb
"S", "SIM", "T", ".T.", "TRUE", "YES", "ON", "Y", "1"
```

---

## Dicas de uso

- Prefira `ProfileString` e `SetProfile` para novo código.
- Use `GetIni` / `SetIni` quando quiser manter compatibilidade com funções antigas.
- Para valores de caminho, `PegPath` é útil quando o caminho está em um INI e precisa ter fallback para o diretório do programa.
- Sempre mantenha seções e chaves consistentes, por exemplo:

```vb
"DB", "SERVER"
"APP", "VERSION"
"PATH", "ARQUIVOS"
```


---

## Resumo

O módulo `modIni.bas` é uma base sólida para gerenciar configurações em INI em aplicações VB6/VBA. Ele combina compatibilidade com legado e funções modernas, sendo útil para:

- arquivos de configuração;
- caminhos de pasta e banco de dados;
- parâmetros de execução;
- flags e preferências do sistema.

Se o projeto usa muitos parâmetros de ambiente, essa abordagem mantém os dados externos bem organizados e fáceis de manter.
y