# Sistema de Gestão de Imagens e Integração SQL (VB6)

Este projeto é uma biblioteca de utilitários e formulários desenvolvida em **Visual Basic 6.0 (VB6)**, focada no gerenciamento de imagens (JPEG/Blob), manipulação de bases de dados via ADO/ADOX e integração com diversos drivers modernos e legados.

## 🚀 Funcionalidades Principais

### 1. Gestão e Conversão de Imagens (Blob para JPEG)
O sistema possui um motor robusto para manipular imagens armazenadas em bases de dados (campos BLOB/Long Binary):
- **Exportação em Massa:** Através do formulário `frmexpjpg`, o sistema varre a base de dados e exporta fotos (funcionários, peças, produtos) para ficheiros físicos `.jpg`.
- **Redimensionamento e Qualidade:** Integração para definir a qualidade do JPEG (ex: 70%) e visualização prévia em tempo real.
- **Suporte a ERPs de Mercado:** Lógica customizada para estruturas de dados de sistemas como Logix e Datamace.

### 2. Abstração de Conectividade (Multi-Database)
A biblioteca `adolib.bas` atua como uma camada de compatibilidade para diversos motores:
- **Suporte SQLite:** Funções específicas para integração com bases de dados SQLite.
- **Motores Microsoft:** Conexões prontas para Microsoft Jet (`.mdb`) e ACE (`.accdb`).
- **Visual FoxPro (VFP):** Tratamento especial para comandos de drivers VFP (ex: `SET DELETED ON`).
- **Gestão de Strings de Conexão:** Centralização de provedores (OLEDB/ODBC) para facilitar a troca de ambientes.

### 3. Utilitários SQL e Automatização
O módulo `SqlFuncoesAdo.bas` simplifica a escrita de código VB6:
- **Execução Simplificada:** Funções para executar comandos SQL (`INSERT`, `UPDATE`, `DELETE`) com tratamento de erros.
- **Cópia de Recordsets:** Lógica para clonar dados entre diferentes tabelas ou bases de dados de forma automatizada.
- **Tratamento de Tipos:** Preparação de valores SQL para evitar erros de formatação de data e decimais entre diferentes dialetos.

### 4. Interface e Experiência do Utilizador
- **Grelhas Avançadas:** Uso do `VBFlexGrid` para exibição de dados com suporte a filtragem e ordenação.
- **Modernização Visual:** Implementação de `EnabledVisualStyles` para suporte ao aspeto visual do Windows 10/11.
- **Visualização de Impressão:** Integração com o componente `BSPrinter` para gerar previews de documentos com imagens.

## ⚙️ Características Técnicas

- **Linguagem:** Visual Basic 6.0 (VB6).
- **Arquitetura:** Modular (`.bas` para lógica, `.frm` para interface).
- **Acesso a Dados:** ActiveX Data Objects (ADO) - compatível com versões 2.8 e 6.1.
- **Dependências Principais:** - `XpControls.ocx` (Interface)
  - `VBFLXGRD18.OCX` (VBFlexGrid)
  - `vbccr18.ocx` (Common Controls)
  - `BSPrin10.ocx` (Impressão)

## 🔧 Notas de Instalação e Migração

1. **Referências:** Ao abrir o projeto, certifique-se de que as referências ao *Microsoft ActiveX Data Objects 6.1 Library* estão ativas no menu `Project -> References`.
2. **Drivers:** Para o correto funcionamento do SQLite, o driver ODBC correspondente deve estar instalado na arquitetura 32-bit (SysWOW64).
3. **Ambiente:** O ficheiro `imagens.bas` contém a `Sub Main` que inicializa os estilos visuais e as variáveis globais de conexão.

---
*Projeto desenvolvido para manutenção de sistemas legados e integração de ferramentas modernas.*