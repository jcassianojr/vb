# ImagensMain - Automação de Processamento de Imagens (VB6)

Este projeto é um utilitário desenvolvido em **Visual Basic 6.0** focado na gestão e migração de imagens entre sistemas de arquivos e bases de dados relacionais. Ele permite automatizar o carregamento de arquivos JPG para campos binários (BLOB) em diferentes tipos de bases de dados.

## 🚀 Funcionalidades

* **Conexão Híbrida**: Suporte automático para bancos de dados Microsoft Access (`.mdb` e `.accdb`) e SQLite (`.sqlite`, `.db`).
* **Inteligência de Driver**: Identifica se o caminho fornecido no arquivo de configuração é apenas um caminho de ficheiro ou uma string de conexão completa, montando o provedor OLEDB ou Driver ODBC adequado.
* **Validação de Ambiente**: Verifica se o Driver SQLite ODBC (32 bits) está instalado no Registro do Windows (incluindo suporte a sistemas 64-bit via WOW6432Node) antes de iniciar o processo.
* **Processamento em Lote**: Percorre a tabela de imagens, localiza o ficheiro físico correspondente ao código e utiliza o método `AppendChunk` para atualizar o banco de dados.
* **Saída via Console**: Integração com APIs do Windows para fornecer feedback em tempo real através de uma janela de console dedicada.

## 🛠️ Requisitos Técnicos

* **IDE**: Visual Basic 6.0.
* **Bibliotecas (Referências)**:
    * `Microsoft ActiveX Data Objects 6.1 Library` (ADO)
    * `RC6` (RichClient6)
    * `OLEEXP` (Modern Shell Interfaces)
* **Dependências de Runtime**:
    * Para ficheiros `.accdb`: Microsoft Access Database Engine (ACE).
    * Para ficheiros `.sqlite`: SQLite3 ODBC Driver (32 bits).

## ⚙️ Configuração (imgmpcpf.ini)

O sistema utiliza um ficheiro `.ini` para definir os caminhos de trabalho. Se apenas o caminho do ficheiro for fornecido em `IMGCOM`, o sistema gera a string de conexão automaticamente:

```ini
[PATH]
IMGJPG=C:\Fotos\          ; Pasta onde estão os ficheiros .JPG
IMGCOM=C:\Dados\base.db   ; Caminho do ficheiro (MDB, ACCDB ou SQLite) ou String ADO completa