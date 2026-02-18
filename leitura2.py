import pdfplumber
import re
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

def extrair_dados_guia(page):
    """Extrai dados estruturados de uma página da guia TISS usando pdfplumber."""
    texto = page.extract_text() or ''
    dados = {}

    # --- Dados Gerais ---
    dados['Registro ANS'] = extrair_campo(texto, r'1 - Registro ANS.*?(\d{5,6})')
    # Tenta primeiro o campo 3, se não encontrar, busca no campo 2
    dados['Número da Guia Principal'] = extrair_campo(texto, r'3\s*-\s*Número da Guia Principal\s*\n?\s*\d+\s+(\d{5,})')
    if not dados['Número da Guia Principal']:
        # Formato alternativo: "2-N° da Guia no Prestador" seguido do número na próxima linha
        match = re.search(r'2\s*-\s*N[°º]?\s*(?:da\s+)?Guia\s+no\s+Prestador\n.*?(\d{5,})', texto)
        if match:
            dados['Número da Guia Principal'] = match.group(1)
        else:
            dados['Número da Guia Principal'] = ''
    dados['Número da Guia'] = extrair_campo(texto, r'7 - Número da Guia.*?\n(\d+)')
    dados['Número da Carteira'] = extrair_campo(texto, r'8 - Número da Carteira\s*\n?\s*(\d+)')
    dados['Senha'] = extrair_campo(texto, r'5 - Senha.*?(\d+)')
    dados['Data Autorização'] = extrair_campo(texto, r'4 - Data da Autorização\s*\n?\s*(\d{2}/\d{2}/\d{4})')
    dados['Data Validade Senha'] = extrair_campo(texto, r'6 - Data de Validade da Senha\s*\n?\s*(\d{2}/\d{2}/\d{4})')
    dados['Validade Carteira'] = extrair_campo(texto, r'9 - Validade da Carteira\s*\n?\s*(\d{2}/\d{2}/\d{4})')

    # --- Dados do Beneficiário ---
    # Tenta capturar o nome que aparece antes de "4 - Data" ou após "10 - Nome"
    dados['Nome Beneficiário'] = extrair_campo(
        texto, r'(?:10 - Nome.*?\n|Beneficiário.*?\n)\s*([A-ZÀ-Ú ]{5,})'
    )
    # Fallback: pegar nome em CAPS que parece nome de pessoa
    if not dados['Nome Beneficiário']:
        dados['Nome Beneficiário'] = extrair_campo(
            texto, r'([A-ZÀ-Ú]{3,}(?:\s+(?:DO|DA|DOS|DAS|DE)?\s+[A-ZÀ-Ú]{3,}\s+[A-ZÀ-Ú]{3,}))'
        )
    dados['Recém Nascido'] = extrair_campo(texto, r'12 - Recém Nascido\s*\n?\s*(Sim|Não)')

    # --- Dados do Solicitante ---
    dados['Contratado Solicitante'] = extrair_campo(texto, r'14 - Nome do Contratado\s*\n?\s*(.+?)(?=\s*19|\s*\n)')
    dados['Nome Profissional'] = extrair_campo(texto, r'15 - Nome do Profissional Solicitante.*?\n\s*(.+?)(?=\s*DF|\s*\n)')
    dados['Conselho'] = extrair_campo(texto, r'16 - Conselho Profissional\s*\n?\s*(CRM|CRO|CRP|CREFITO)')
    dados['Número Conselho'] = extrair_campo(texto, r'17 - Número do Conselho\s*\n?\s*(\d+)')
    dados['UF Conselho'] = extrair_campo(texto, r'18 - UF\s*(\w{2})')
    dados['Código CBO'] = extrair_campo(texto, r'19 - Código\s*CBO\s*\n?\s*(\w+)')

    # --- Dados da Solicitação ---
    dados['Caráter Atendimento'] = extrair_campo(texto, r'21 - Caráter.*?(ELETIVO|URGENCIA|URGÊNCIA)')
    dados['Data Solicitação'] = extrair_campo(texto, r'22 - Data da Solicitação.*?(\d{2}/\d{2}/\d{4})')
    dados['Indicação Clínica'] = extrair_campo(texto, r'23 - Indicação Clínica\s*\n?\s*(.+?)(?=\s*\d{2}/\d{2}/\d{4}|\n)')

    # --- Procedimentos via extract_tables (mais confiável) ---
    dados['Procedimentos'] = extrair_procedimentos_tabela(page, texto)

    return dados


def extrair_procedimentos_tabela(page, texto_fallback):
    """
    Tenta extrair procedimentos via tabela do pdfplumber.
    Se não conseguir, faz fallback com regex no texto.
    """
    procedimentos = []

    # Tentar extrair tabelas da página
    tabelas = page.extract_tables()
    for tabela in tabelas:
        for linha in tabela:
            if not linha:
                continue
            # Limpar células None
            celulas = [str(c).strip() if c else '' for c in linha]
            texto_linha = ' '.join(celulas)

            # Procurar linhas com código de 8 dígitos (padrão TUSS)
            match = re.search(r'(\d{1})\s*(\d{8})\s+(.+?)\s+(\d+)\s+(\d+)', texto_linha)
            if match:
                procedimentos.append({
                    'Tabela': match.group(1),
                    # 'Código': match.group(2),
                    'Descrição': match.group(3).strip(),
                    'Qtde Solicitada': int(match.group(4)),
                    'Qtde Autorizada': int(match.group(5)),
                })

    # Fallback: regex no texto bruto
    if not procedimentos:
        matches = re.findall(r'(\d)\s+(\d{8})\s+(.+?)\s+(\d+)\s+(\d+)', texto_fallback)
        for m in matches:
            procedimentos.append({
                'Tabela': m[0],
                # 'Código': m[1],
                'Descrição': m[2].strip(),
                'Qtde Solicitada': int(m[3]),
                'Qtde Autorizada': int(m[4]),
            })

    return procedimentos


def extrair_campo(texto, pattern):
    """Helper para extrair um campo via regex."""
    match = re.search(pattern, texto, re.IGNORECASE | re.DOTALL)
    if match and match.lastindex:
        return match.group(1).strip()
    elif match:
        return match.group(0).strip()
    return ''


def exportar_para_excel(lista_dados, nome_arquivo='guias_extraidas.xlsx'):
    """Exporta dados das guias para Excel com formatação."""
    wb = Workbook()

    # Estilos
    header_font = Font(bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # =============================================
    # ABA 1 - DADOS GERAIS
    # =============================================
    ws = wb.active
    ws.title = 'Dados Gerais'

    campos = [
        'Registro ANS', 'Número da Guia Principal', 'Número da Guia', 'Número da Carteira', 'Senha',
        'Nome Beneficiário', 'Contratado Solicitante', 'Nome Profissional',
        'Conselho', 'Número Conselho', 'UF Conselho',
        'Caráter Atendimento', 'Data Solicitação', 'Indicação Clínica',
    ]

    # Cabeçalho
    ws.append(campos)
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    # Dados
    for dados in lista_dados:
        ws.append([dados.get(c, '') for c in campos])

    # Ajustar largura
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 3, 40)

    # =============================================
    # ABA 2 - PROCEDIMENTOS
    # =============================================
    ws2 = wb.create_sheet('Procedimentos')
    cols_proc = [
        'Beneficiário', 'Nº Guia Principal', 'Descrição', 'Qtde Autorizada'
    ]
    ws2.append(cols_proc)
    for cell in ws2[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align

    for dados in lista_dados:
        for proc in dados.get('Procedimentos', []):
            ws2.append([
                dados.get('Nome Beneficiário', ''),
                dados.get('Número da Guia Principal', ''),
                proc['Descrição'],
                proc['Qtde Autorizada'],
            ])

    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = min(max_len + 3, 50)

    wb.save(nome_arquivo)
    print(f'\n✅ Arquivo salvo: {nome_arquivo}')


# =============================================================
# EXECUÇÃO - PROCESSAR UM OU VÁRIOS PDFs
# =============================================================

# Opção 1: Um único PDF
# arquivos = ['ALICE PITOMBEIRA PONTES_2.pdf']

# Opção 2: Todos os PDFs de uma pasta
pasta = 'guias'  # pasta atual para o caminho desejado
arquivos = [f for f in os.listdir(pasta) if f.lower().endswith('.pdf')]

todas_guias = []

for nome_arquivo in arquivos:
    caminho = os.path.join(pasta, nome_arquivo)
    print(f'\n📄 Processando: {nome_arquivo}')

    with pdfplumber.open(caminho) as pdf:
        for i, page in enumerate(pdf.pages):
            dados = extrair_dados_guia(page)
            dados['Arquivo Origem'] = nome_arquivo
            dados['Página'] = i + 1
            todas_guias.append(dados)

            # Preview no terminal
            print(f'  Página {i+1}:')
            for k, v in dados.items():
                if k not in ('Procedimentos', 'Arquivo Origem', 'Página'):
                    if v:
                        print(f'    {k}: {v}')
            for proc in dados.get('Procedimentos', []):
                print(f'    → {proc["Descrição"]} '
                      f'| Aut: {proc["Qtde Autorizada"]}')

# Exportar
if todas_guias:
    exportar_para_excel(todas_guias)
else:
    print('⚠️ Nenhuma guia encontrada nos PDFs.')