Aqui está a documentação técnica consolidada, incluindo as strings de conexão necessárias para o seu projeto.

---

# Guia de Implementação: OLEDB para CSV

Este guia descreve como utilizar o Microsoft.ACE.OLEDB para manipular arquivos CSV como tabelas, garantindo flexibilidade através do schema.ini.

## 1. Connection Strings (VB.NET)

Dependendo da arquitetura do seu sistema (32-bit vs 64-bit) e do driver instalado, utilize uma das opções abaixo.

### Opção A: Driver ACE (Recomendado para sistemas modernos)

Este driver é o mais versátil e suporta arquivos .xlsx e .csv.

```vbnet
Dim pasta As String = "C:\Caminho\Dos\Arquivos\"
Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & pasta & _
                           ";Extended Properties='text;HDR=Yes;FMT=Delimited';"

```

### Opção B: Driver Jet (Legado, apenas 32-bit)

Use apenas se estiver mantendo compatibilidade restrita com sistemas muito antigos (VB6).

```vbnet
Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pasta & _
                           ";Extended Properties='text;HDR=Yes;FMT=Delimited';"

```

> **Nota:** HDR=Yes indica que a primeira linha contém os nomes das colunas.

---

## 2. Estrutura do schema.ini

Coloque este arquivo na mesma pasta do seu CSV. O driver o lerá automaticamente ao abrir a conexão.

```ini
[SeuArquivo.csv]
ColNameHeader=True
Format=Delimited(;)
MaxScanRows=0
CharacterSet=ANSI
DateTimeFormat=dd/mm/yyyy

; Mapeamento (Ordem Física das colunas no arquivo)
Col1=ID Long
Col2=CodigoNCM Text Width 12
Col3=Descricao Memo
Col4=Aliquota Double
Col5=ValorBase Currency
Col6=DataVigencia DateTime
Col7=Ativo Bit

```

---

## 3. Resumo de Tipos e Parâmetros

| Tipo | Descrição |
| --- | --- |
| **Long** | Inteiro 32-bit (Chaves/IDs) |
| **Double** | Ponto flutuante (Cálculos/Percentuais) |
| **Currency** | Monetário (Precisão fixa) |
| **DateTime** | Data/Hora (Formatado) |
| **Text Width N** | Texto com tamanho fixo |
| **Memo** | Texto longo (sem limite) |

**Parâmetros fundamentais:**

* **MaxScanRows=0**: Faz o driver ler o arquivo inteiro para inferir tipos. **Essencial** para evitar erros de leitura.
* **Format=Delimited(;)**: Define o separador. Altere para TabDelimited ou CSVDelimited conforme a origem.
* **DecimalSymbol=.**: Use se o CSV utilizar ponto como separador decimal.

---

## 4. Dicas Estratégicas para o seu Sistema

1. **Versatilidade Dinâmica:** Como seu sistema lida com diferentes layouts de impostos, implemente uma função que **gere o arquivo schema.ini dinamicamente** antes de abrir a OleDbConnection. Isso permite que seu sistema se adapte instantaneamente a qualquer alteração de colunas do fornecedor.
2. **Segurança:** Use SELECT via SQL para extrair apenas as colunas que seu sistema conhece. Se o arquivo do fornecedor mudar a ordem das colunas, o schema.ini manterá o mapeamento correto e seu código permanecerá intacto.
3. **Performance:** Para processar grandes volumes, leia o CSV via OleDbDataAdapter para um DataTable. Uma vez em memória (DataTable), você terá flexibilidade total para filtros e manipulações sem onerar o arquivo original ou o driver OLEDB.

---

*Referência para desenvolvimento em VB.NET / Integração de sistemas.*