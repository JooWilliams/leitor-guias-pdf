# leGuias - Extrator de Guias TISS (PDF → Excel)

Ferramenta para extrair dados de guias médicas no padrão TISS (SP/SADT) em PDF e exportar para planilha Excel organizada.

## Funcionalidades

- Leitura automática de todos os PDFs dentro da pasta `guias/`
- Extração de dados do beneficiário, solicitante, procedimentos e números de guia
- Suporte a diferentes layouts de guia (Bradesco, Unimed, etc.)
- Fallback inteligente: busca número da guia nos campos 3 e 2
- Exportação para Excel com duas abas: **Dados Gerais** e **Procedimentos**

## Requisitos

- Python 3.8+
- pdfplumber
- openpyxl

## Instalação

```bash
pip install pdfplumber openpyxl
```

## Estrutura do Projeto

```
leGuias/
├── leitura2.py          # Script principal
├── guias/               # Pasta com os PDFs das guias
│   ├── guia1.pdf
│   ├── guia2.pdf
│   └── ...
├── guias_extraidas.xlsx  # Arquivo gerado (após execução)
├── .gitignore
└── README.md
```

## Como Usar

1. Coloque os PDFs das guias dentro da pasta `guias/`
2. Execute o script:

```bash
python leitura2.py
```

3. O arquivo `guias_extraidas.xlsx` será gerado com os dados extraídos.

## Campos Extraídos

### Aba "Dados Gerais"
| Campo | Descrição |
|---|---|
| Registro ANS | Código da operadora na ANS |
| Número da Guia Principal | Número principal da guia (campo 3 ou 2) |
| Número da Guia | Número atribuído pela operadora |
| Número da Carteira | Carteira do beneficiário |
| Nome Beneficiário | Nome completo do paciente |
| Data Solicitação | Data da solicitação médica |
| Indicação Clínica | Motivo/indicação clínica |

### Aba "Procedimentos"
| Campo | Descrição |
|---|---|
| Beneficiário | Nome do paciente |
| Nº Guia Principal | Número da guia vinculada |
| Descrição | Nome do procedimento (código TUSS) |
| Qtde Autorizada | Quantidade autorizada pela operadora |

## Observações

- Os PDFs não são versionados por conterem dados sensíveis de pacientes (LGPD)
- O arquivo Excel gerado também é ignorado pelo `.gitignore`
- Diferentes operadoras podem ter layouts distintos — ajuste os regex conforme necessário