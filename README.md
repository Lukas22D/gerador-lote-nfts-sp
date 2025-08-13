# Gerador de Lote NFTS - São Paulo

Aplicação para gerar arquivos de lote NFTS (Nota Fiscal de Serviços) para a Prefeitura de São Paulo, agora com funcionalidade de importação de dados via Excel.

## Funcionalidades

- **Importação Excel**: Importa notas fiscais de planilhas Excel com formato padronizado
- **Preenchimento Automático**: Código do serviço é preenchido automaticamente com base no Item/SubItem
- **Validação Completa**: Validação rigorosa de todos os campos obrigatórios
- **Interface Moderna**: Interface gráfica atualizada com PyQt5
- **Arquitetura Segregada**: Código organizado com responsabilidades bem definidas

## Instalação

1. Instale as dependências:
```bash
pip install -r requirements.txt
```

2. Execute a aplicação:
```bash
python3 src/app.py
```

## Formato do Excel

O arquivo Excel deve conter as seguintes colunas (exatamente com estes nomes):

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| Tipo do Documento | Tipo da nota fiscal | "01 - Dispensado" |
| Numero do Documento | Número da nota | "1001" |
| Série do Documento | Série (obrigatória para tipo 02) | "1" |
| Data da Prestação | Data da prestação | "15/12/2024" |
| Tributação do Serviço | Tipo de tributação | "T - Operação Normal" |
| Código do Serviço | CNAE (preenchido automaticamente) | "2919" |
| Item/SubItem | Item/subitem da classificação | "0107" |
| Valor | Valor da nota | "1500,00" |
| Aliquota | Alíquota do ISS | "500" (5%) |
| ISS Retido pelo Tomador | ISS retido | "Sim" ou "Não" |
| Tipo de Prestador | Tipo de documento | "2 - CNPJ" |
| CNPJ/CPF do Prestador | Documento do prestador | "12.345.678/0001-90" |
| Cidade | Cidade do prestador | "São Paulo" |
| UF | Estado | "SP" |
| CEP | CEP do prestador | "01234-567" |
| Discriminação dos Serviços | Descrição dos serviços | "Serviços de TI" |

## Preenchimento Automático do Código do Serviço

Quando você preencher o campo "Item/SubItem" com pelo menos 4 dígitos (ex: "0107"), o sistema automaticamente:

1. Remove zeros à esquerda do valor
2. Busca a classificação correspondente no arquivo `aliquotas.json`
3. Preenche o "Código do Serviço" com o CNAE correspondente

**Exemplo**: Item/SubItem "0107" → Código do Serviço "02919"

## Como Usar

1. **Prepare seu Excel**: Use o arquivo `exemplo_notas_fiscais.xlsx` como modelo
2. **Execute a aplicação**: `python3 src/app.py`
3. **Selecione CCM**: Escolha o contribuinte
4. **Importe Excel**: Clique em "Selecionar Arquivo Excel" e escolha sua planilha
5. **Revise dados**: Verifique as notas importadas na tabela
6. **Gere arquivo**: Clique em "Gerar Arquivo" para salvar o lote NFTS

## Geração de Arquivo de Exemplo

Para gerar um arquivo Excel de exemplo:

```bash
python3 exemplo_excel.py
```

Isso criará o arquivo `exemplo_notas_fiscais.xlsx` com 3 notas de exemplo.

## Arquitetura do Código

O código foi reestruturado com separação de responsabilidades:

- **`ExcelImporter`**: Responsável por importar e processar dados do Excel
- **`DataValidator`**: Responsável por validar todos os dados
- **`FileGenerator`**: Responsável por gerar o arquivo de saída
- **`NoteDialog`**: Interface para adicionar/editar notas manualmente
- **`MainWindow`**: Interface principal da aplicação

## Validações

O sistema valida automaticamente:
- Formato de datas, valores e documentos
- Campos obrigatórios
- Limites de caracteres
- Códigos válidos para tipos de documento, tributação, etc.
- Consistência entre dados relacionados

## Observações

- O arquivo `aliquotas.json` deve estar presente no diretório raiz
- Valores monetários podem usar vírgula ou ponto como separador decimal
- Documentos (CNPJ/CPF) podem ter máscara, serão automaticamente limpos
- CEP pode ter máscara (formato 12345-678)
- Campos vazios são tratados adequadamente 