# Módulo `mFileAPI.bas` — Documentação Técnica & Guia de Uso

O módulo **`mFileAPI.bas`** é uma biblioteca altamente otimizada, robusta e voltada para o ecossistema **Visual Basic 6 (VB6) no Windows**. Ele reúne funções nativas de API do Windows (`comdlg32`, `kernel32`, `shell32`) e o motor inteligente do *FileSystemObject (FSO)* para realizar manipulações avançadas de arquivos, pastas, diálogos do sistema e normalização de caminhos de rede.

---

## 📋 Índice
1. [Visão Geral e Arquitetura](#visão-geral-e-arquitetura)
2. [Constantes e Estruturas de Dados (API)](#constantes-e-estruturas-de-dados-api)
3. [Funções de Manipulação de Caminhos e Nomes](#funções-de-manipulação-de-caminhos-e-nomes)
4. [Operações de Arquivos e Pastas](#operações-de-arquivos-e-pastas)
5. [Leitura e Escrita de Arquivos (Streams)](#leitura-e-escrita-de-arquivos-streams)
6. [Diálogos do Windows (Open / Save / Folders)](#diálogos-do-windows-open--save--folders)
7. [Operações em Lote via Shell (`SHFILEOPSTRUCT`)](#operações-em-lote-via-shell-shfileopstruct)

---

## 🛠️ Visão Geral e Arquitetura

O módulo foi refatorado para eliminar falhas comuns do VB6 clássico (como estouramento de buffers de API, erros em caracteres Unicode/especiais e caminhos malformados com barras duplicadas). Ele combina o melhor de dois mundos: **APIs de baixo nível do Windows** para máxima performance e **FSO (FileSystemObject)** para segurança em árvores de diretórios complexas e redes corporativas (UNC).

---

## ⚙️ Constantes e Estruturas de Dados (API)

O módulo declara flags padrão do Windows para caixas de diálogo comuns (`OPENFILENAME`) e estruturas de manipulação de arquivos via Shell:

* **`OPENFILENAME`**: Estrutura usada pelas APIs `GetOpenFileName` e `GetSaveFileName`.
* **`SHFILEOPSTRUCT`**: Estrutura usada para operações de cópia, movimentação, exclusão e renomeação gerenciadas pelo explorador de arquivos do Windows (`SHFileOperation`).
* **`TFileFlags`**: Enumeração com flags de controle para o Shell (`FOF_ALLOWUNDO`, `FOF_NOCONFIRMATION`, etc.).

---

## 📁 Funções de Manipulação de Caminhos e Nomes

### `FixPath(ByVal pathToFix As String) As String`
Normaliza qualquer caminho de diretório, convertendo barras invertidas incorretas (`/` para `/`), removendo duplicidades excessivas (preservando o prefixo de rede `//` UNC) e garantindo que o caminho termine com a barra separadora (`/`).

### `GetDriveUNC(ByVal DriveString As String) As String`
Identifica se uma unidade de disco (ex: `Z:`) está mapeada para um caminho de rede e retorna o endereço UNC correspondente (ex: `//servidor/compartilhamento`), ou retorna a unidade local formatada.

### `ValidFileName(...) As String`
Sanitiza uma string proposta para garantir que se torne um nome de arquivo válido no Windows. 
* Remove caracteres proibidos (`*`, `?`, `<`, `>`, `|`, etc.).
* Trata espaços e pontos ilegais nas extremidades (início e fim do nome).
* Protege contra nomes reservados do sistema operacional (`CON`, `PRN`, `AUX`, `NUL`, `COM1`-`COM9`, `LPT1`-`LPT9`).
* Controla o comprimento máximo e suporta o formato antigo `8.3`.

### `parsefile(ByVal archivo As String, ByVal parte As String) As String`
Função modular que extrai partes específicas de um caminho completo com base em uma letra de parâmetro:
* `"D"`: Drive (ex: `c:`)
* `"R"`: Raiz (ex: `c:/Users/Download/`)
* `"C"`: Caminho completo sem extensão
* `"N"`: Nome do arquivo sem extensão
* `"A"`: Nome do arquivo com extensão (`arquivo.txt`)
* `"E"`: Extensão pura (`txt`)
* `"PD"`: Diretório parente
* `"PP"`: Caminho parente completo

---

## 📂 Operações de Arquivos e Pastas

### `FileExists(ByRef sFileName As String) As Boolean`
Verifica se um arquivo existe de forma híbrida. Tenta primeiro a checagem nativa rápida via `GetAttr` do VB6 e faz fallback automático para o FSO caso o caminho contenha caracteres especiais ou Unicode que gerem erro de compilação/sistema.

### `FolderExists(ByRef sDir As String) As Boolean`
Valida a existência de um diretório no disco, limpando barras finais excedentes para garantir compatibilidade com o validador do FSO.

### `CreateNewDirectory(ByVal NewDirectory As String)`
Cria uma árvore inteira de diretórios de forma recursiva (`C:/PastaA/PastaB/PastaC`) em uma única chamada, utilizando o FSO.

### `CopyFolder(ByVal sourcePath As String, ByVal destinationPath As String, ...)`
Clona pastas inteiras e subestruturas recursivamente, criando o destino automaticamente se ele não existir.

### `DeleteFolder(ByRef FolderPath As String, ...)`
Remove pastas por completo, incluindo todo o conteúdo interno de subpastas e arquivos, com suporte a exclusão forçada (ignorando atributos de *Somente Leitura*).

### `GetTempDirectory() As String`
Recupera o diretório temporário do sistema operacional (`Temp`) utilizando alocação dinâmica de buffer via API do Windows (`GetTempPath`), eliminando riscos de corte em perfis de usuário com caminhos longos. Já retorna formatado com `FixPath`.

---

## 📄 Leitura e Escrita de Arquivos (Streams)

O módulo inclui funções otimizadas baseadas no FSO para gerenciar arquivos de texto de forma limpa e segura:

* **`FileRead(FName As String) As String`**: Lê todo o conteúdo de um arquivo de texto de uma só vez. Retorna vazio se o arquivo não existir.
* **`FileWrite(FName As String, Contents As String) As Boolean`**: Cria ou sobrescreve um arquivo de texto com o conteúdo fornecido.
* **`FileAppend(FName As String, Contents As String) As Boolean`**: Adiciona conteúdo ao final de um arquivo existente (ou o cria se não existir), adicionando uma quebra de linha implícita.
* **`OpenStreamFile(...) As Integer`**: Gerencia abertura de arquivos binários ou randômicos de baixo nível com controle de concorrência (`Shared`, `Lock Read`, `Lock Write`, etc.) e tratamento centralizado de erros de disco (`FileErrors`).

---

## 🪟 Diálogos do Windows (Open / Save / Folders)

* **`FileOpen(...)` & `FileSave(...)`**: Abre as caixas de diálogo nativas do Windows (`comdlg32.dll`) para seleção de arquivos de abertura ou salvamento, suportando filtros personalizados.
* **`OpenArqExt` & `SaveArqExt`**: Atalhos parametrizados para abrir/salvar arquivos filtrados por uma extensão específica (ex: `.txt`, `.csv`).
* **`SelecionarPasta(ByVal Titulo As String) As String`**: Exibe a janela padrão do Windows (`Shell.Application`) para o usuário escolher um diretório de forma interativa.
* **`ImgFILTER()` & `ImgFILTER2()`**: Retornam strings de filtros formatadas com múltiplos tipos de imagens bitmap, JPEG, PNG, GIF, TIFF, etc., prontas para uso em caixas de diálogo de abertura.

---

## 🔄 Operações em Lote via Shell (`SHFILEOPSTRUCT`)

Funções de alto desempenho integradas diretamente ao Shell do Windows para gerenciar grandes volumes de arquivos com suporte a desfazer (`Undo`), progresso e tratamento de erros do sistema:
* `SH_Copy`
* `SH_Move`
* `SH_Rename`
* `SH_Delete`
* `CopyFileWindowsWay`
