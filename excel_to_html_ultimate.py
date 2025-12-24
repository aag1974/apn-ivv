"""
excel_to_html_report_final_v2.py

Gera um dashboard HTML executivo (Chart.js) a partir de um Excel exportado do relatório.

NOVO: Adicionado card de resumo executivo como primeira seção e novo sistema de insights.

Melhorias incluídas nesta versão:
- Card de resumo executivo com principais métricas do mês mais atual (11 métricas + setas de tendência)
- Sistema avançado de insights que extrai dados MoM/YoY diretamente do Excel
- Novos formatos de insights baseados no tipo de dados (percentual, absoluto, monetário)
- Setas de tendência visuais em cada métrica do card
- Legendas padronizadas com "bolinhas" (usePointStyle + circle) em todos os gráficos
- Destaque automático do ano corrente (última série) em gráficos de linha
- Consolidações anuais com cores por ano (mesma paleta do template)
- Insights já injetados no HTML (sem placeholders manuais)

Formato dos novos insights:
Para dados em percentual (IVV):
- Variação MoM: Out/2025 - Set/2025: 1,3%
- Variação YoY: Out/2025 - Out/2024: 48,4%
- Pico: XX,X% (Mai/25)
- Média Trimestral: XX,X%
- Média anual: XX,X%

Para dados absolutos (Ofertas, Vendas, Lançamentos):
- Variação MoM: Out/2025 - Set/2025: 1,3%
- Variação YoY: Out/2025 - Out/2024: 48,4%
- Pico: X.XXX (Mai/25)
- Média Trimestral: X.XXX
- Média anual: X.XXX

Para dados monetários (Preços):
- Variação MoM: Out/2025 - Set/2025: 1,3%
- Variação YoY: Out/2025 - Out/2024: 48,4%
- Pico: R$ X.XXX (Mai/25)
- Média Trimestral: R$ X.XXX
- Média anual: R$ X.XXX

Para dados monetários em milhões (VGV, VGL):
- Variação MoM: Out/2025 - Set/2025: 1,3%
- Variação YoY: Out/2025 - Out/2024: 48,4%
- Pico (R$ Mi): R$ X.XXX (Mai/25)
- Média Trimestral (R$ Mi): R$ X.XXX
- Média anual (R$ Mi): R$ X.XXX

Uso:
    python3 excel_to_html_report_final_v2.py <input_excel.xlsx> <output_html.html>

Requisitos:
    pandas, numpy
"""

import sys
import os
import re
import json
import pandas as pd
import numpy as np
from datetime import datetime

def br_int(value: float) -> str:
    """12.345"""
    return f"{value:,.0f}".replace(",", ".")


def br_float(value: float, decimals: int = 1) -> str:
    """1.234,5"""
    s = f"{value:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def br_percent(value: float, decimals: int = 1) -> str:
    """7,6%"""
    return br_float(value, decimals) + "%"


def br_currency(value: float, decimals: int = 2) -> str:
    """R$ 1.234.567,89"""
    return "R$ " + br_float(value, decimals)


# Paleta de cores (mesma lógica do template)
COLOURS = [
    '#e74c3c',  # red
    '#f39c12',  # orange
    '#9b59b6',  # purple
    '#3498db',  # blue
    '#27ae60',  # green
    '#e67e22',  # carrot
    '#1abc9c',  # turquoise
]


# -------------------------
# Parsers
# -------------------------
def parse_percentage(value) -> float | None:
    """Converte '8,6%' em 8.6"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, (int, float)) and not (isinstance(value, float) and pd.isna(value)):
        return float(value)
    if not isinstance(value, str):
        value = str(value)
    v = value.strip().replace('%', '').replace('.', '').replace(',', '.')
    try:
        return float(v)
    except ValueError:
        return None


def parse_number(value) -> float | None:
    """
    Converte números considerando formatação BR.

    Corrige o caso comum do Excel/Pandas interpretar "4.693" como 4.693,
    quando na verdade deveria ser 4693.
    Também trata valores como 21.340, 478.250 etc, que são interpretados
    como 21.34 ou 478.25, mas significam 21 340 ou 478 250.

    A heurística aqui é:
      * Se o valor é um float e tem parte decimal de 3 dígitos com pelo menos
        um dígito diferente de zero (ex.: "4.693", "328.584"), remove o
        separador e trata como milhar.
      * Se o valor é um float e tem parte decimal de 2 dígitos com pelo menos
        um dígito diferente de zero e a parte inteira é >= 10 (ex.: "21.34", "478.25"),
        multiplica por 1000 se a multiplicação resultar em um inteiro.
      * Casos como "244.000" (parte decimal de 3 dígitos, mas todos zeros)
        **não** são tratados como milhar e permanecem como 244.

    Para strings, remove pontos e converte vírgulas em ponto para obter
    floats, ignorando qualquer conteúdo dentro de colchetes.
    """
    # Tratar valores nulos ou NaN
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None

    # Se já é um número (int ou float)
    if isinstance(value, (int, float)):
        # Tratamento especial para floats
        if isinstance(value, float):
            s = f"{value}"
            if '.' in s:
                integer_part, decimal_part = s.split('.')
                # Caso com exatamente 3 dígitos decimais e pelo menos um dígito != '0'
                if len(decimal_part) == 3 and decimal_part.isdigit() and any(ch != '0' for ch in decimal_part):
                    try:
                        return float(integer_part + decimal_part)
                    except ValueError:
                        pass
                # Caso com 2 dígitos decimais (parte inteira >= 10) e pelo menos um dígito != '0'
                elif len(decimal_part) == 2 and decimal_part.isdigit() and any(ch != '0' for ch in decimal_part):
                    try:
                        if integer_part.isdigit() and int(integer_part) >= 10:
                            approx = value * 1000
                            # se aproximado é inteiro, tratar como milhar
                            if abs(approx - round(approx)) < 1e-6:
                                return float(int(round(approx)))
                    except ValueError:
                        pass
                # Caso com 1 dígito decimal (parte inteira >= 10) e decimal não zero
                elif len(decimal_part) == 1 and decimal_part.isdigit() and decimal_part != '0':
                    try:
                        if integer_part.isdigit() and int(integer_part) >= 10:
                            approx = value * 1000
                            if abs(approx - round(approx)) < 1e-6:
                                return float(int(round(approx)))
                    except ValueError:
                        pass
            # Para floats com 1 decimal ou outros casos, não tratar como milhares
        # Para floats e ints não tratados acima, retorna normalmente
        try:
            return float(value)
        except ValueError:
            return None

    # Caso seja string ou outro tipo: limpar formatação brasileira
    text = str(value)
    # Remove conteúdo entre colchetes (observações como "230 [3]")
    text = re.sub(r"\s*\[.*\]", "", text)
    # Remove separadores de milhar (pontos) e troca vírgula por ponto
    text = text.replace('.', '').replace(',', '.')
    try:
        return float(text)
    except ValueError:
        return None

# -------------------------
# New Parser for bracket numbers
# -------------------------
def parse_bracket_number(value) -> float | None:
    """
    Extrai o número inteiro entre colchetes em strings como "230 [3]".
    Retorna None se não houver colchetes ou se a entrada for vazia.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    # Converter para string para buscar padrão
    s = str(value)
    match = re.search(r"\[(\d+)\]", s)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    return None


# -------------------------
# Cleaning
# -------------------------
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas inválidas (metadados, vazios etc.)."""
    if df is None or df.empty:
        return pd.DataFrame()

    df_clean = df.copy()
    first_col = df_clean.columns[0]

    valid_idx = []
    for idx, row in df_clean.iterrows():
        first_val = row[first_col]
        if pd.isna(first_val) or first_val is None:
            continue
        s = str(first_val).strip().lower()
        if not s or s.isspace():
            continue
        if any(k in s for k in ['variações', 'variacao', 'observação', 'observacao', 'nan']):
            continue
        valid_idx.append(idx)

    if not valid_idx:
        return pd.DataFrame()
    return df_clean.iloc[valid_idx].reset_index(drop=True)


# -------------------------
# Regional tables (Ofertas/Vendas/Preços)
# -------------------------
def is_regional_data(df: pd.DataFrame) -> bool:
    """
    Verifica se uma planilha contém dados regionais (não mensais).
    
    Retorna True se a primeira coluna contém nomes de regiões,
    False se contém meses ou outros dados.
    """
    if df.empty or df.shape[1] < 2:
        return False
    
    # Pegar primeira coluna após limpeza básica
    first_col = df.iloc[:, 0].dropna().astype(str).str.strip().str.lower()
    
    # Palavras que indicam dados mensais (não regionais)
    monthly_indicators = [
        'jan', 'fev', 'mar', 'abr', 'mai', 'jun',
        'jul', 'ago', 'set', 'out', 'nov', 'dez',
        'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
        'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro',
        '1t', '2t', '3t', '4t',  # trimestres
        '2021', '2022', '2023', '2024', '2025',  # anos
        'período', 'periodo', 'mês', 'mes', 'month'
    ]
    
    # Palavras que indicam dados regionais
    regional_indicators = [
        'região', 'regiao', 'area', 'área', 'zona', 'setor',
        'bairro', 'distrito', 'localidade', 'regional',
        'centro', 'norte', 'sul', 'leste', 'oeste',
        'brotas', 'asa', 'taguatinga', 'ceilândia', 'sobradinho',
        'águas claras', 'samambaia', 'planaltina', 'gama', 'santa maria'
    ]
    
    monthly_count = 0
    regional_count = 0
    
    # Analisar primeiras 20 linhas da primeira coluna
    for value in first_col.head(20):
        value_lower = str(value).lower()
        
        # Contar indicadores mensais
        if any(indicator in value_lower for indicator in monthly_indicators):
            monthly_count += 1
            
        # Contar indicadores regionais  
        if any(indicator in value_lower for indicator in regional_indicators):
            regional_count += 1
    
    # Se encontrou mais indicadores mensais que regionais, não é regional
    if monthly_count > regional_count:
        return False
        
    # Se encontrou pelo menos alguns indicadores regionais, é regional
    if regional_count > 0:
        return True
        
    # Se não encontrou nenhum indicador claro, verificar padrão de colunas
    # Dados regionais típicos têm: Região, 1qto, 2qtos, 3qtos, 4+qtos, Total
    if df.shape[1] >= 6:
        # Verificar se tem padrão de quartos (1 qto, 2 qtos, etc)
        col_names = []
        if df.shape[0] > 0:
            # Tentar usar primeira linha como cabeçalho
            col_names = df.iloc[0].astype(str).str.lower()
        
        quarters_pattern = ['1', '2', '3', '4', 'total', 'qto', 'qtos']
        pattern_matches = sum(1 for col in col_names if any(q in str(col) for q in quarters_pattern))
        
        if pattern_matches >= 3:  # Se encontrou pelo menos 3 padrões de quartos
            return True
    
    return False


def parse_region_table(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara tabela de regiões para ordenação e exibição.

    A planilha original possui uma linha de cabeçalho com nomes de colunas
    (Região, 1 qto, 2 qtos, 3 qtos, 4+ qtos, Total). As linhas seguintes
    contêm os dados. Esta função remove linhas vazias, renomeia as colunas,
    converte os valores numéricos considerando formatação brasileira e
    ordena as linhas pela coluna Total de forma decrescente, exceto a
    linha "Total Geral", que sempre é colocada no final.
    """
    if df is None or df.empty:
        return pd.DataFrame()

    df_region = df.copy()
    # Manter somente as 6 primeiras colunas
    df_region = df_region.iloc[:, :6]
    expected_cols = ['Região', '1 qto', '2 qtos', '3 qtos', '4+ qtos', 'Total']
    # Atribuir nomes de colunas padronizados
    df_region.columns = expected_cols
    # Verificar se a primeira linha contém texto de cabeçalho e descartá-la
    if not df_region.empty:
        first_val = str(df_region.at[0, 'Região']).strip().lower() if df_region.at[0, 'Região'] is not None else ''
        if first_val in ['região', 'regiao', 'regi']:
            df_region = df_region.drop(index=0)
    df_region = df_region.reset_index(drop=True)
    # Remover linhas com região vazia
    df_region = df_region[df_region['Região'].notna()].copy()
    # Converter valores para float para ordenação
    def to_float(val: any) -> float:
        v = parse_number(val)
        return v if v is not None else 0.0
    df_region['Total_num'] = df_region['Total'].apply(to_float)
    # Identificar linhas de totais (qualquer ocorrência de "total")
    regiao_norm = df_region['Região'].astype(str).str.strip().str.lower()
    mask_total = regiao_norm.str.contains('total')
    total_rows = df_region.loc[mask_total].copy()
    df_region_no_total = df_region.loc[~mask_total].copy()
    # Ordenar linhas restantes pelo total decrescente
    df_region_no_total = df_region_no_total.sort_values(by='Total_num', ascending=False)
    # Concatenar mantendo totais no final
    df_sorted = pd.concat([df_region_no_total, total_rows], ignore_index=True)
    # Remover coluna auxiliar
    df_sorted = df_sorted.drop(columns=['Total_num'])
    return df_sorted


def create_region_table_html(df: pd.DataFrame, title: str) -> str:
    """
    Converte um DataFrame de regiões em uma tabela HTML.

    Inclui título e um subtítulo padronizado. Assume-se que os dados
    já estão ordenados corretamente.

    O subtítulo "Distribuição por número de quartos" passa a ser
    utilizado no lugar de "Distribuição por tipo de quarto (ordenado por Total)"
    conforme solicitação do usuário.
    """
    if df is None or df.empty:
        return ''
    headers = list(df.columns)
    # Função auxiliar para formatar células: remove casas decimais de inteiros
    def format_cell(val: any) -> str:
        """Formata uma célula para exibição em HTML.

        - Para valores nulos/NaN retorna vazio.
        - Para números inteiros retorna com separador de milhares.
        - Para demais valores retorna a representação original (por exemplo,
          strings já formatadas com duas casas decimais para tabelas de preço).
        """
        if val is None or (isinstance(val, float) and pd.isna(val)):
            return ''
        val_str = str(val)
        # Tentar converter para número
        num = parse_number(val)
        # Se for string com vírgula (indicando formatação decimal) ou for traço,
        # retornar como está para preservar as casas decimais ou o marcador de zero
        if isinstance(val_str, str):
            stripped = val_str.strip()
            if ',' in stripped or stripped == '-':
                return val_str
        # Se o valor é um número e é inteiro (ou quase inteiro), aplicar formatação
        if num is not None and abs(num - round(num)) < 1e-6:
            try:
                return br_int(int(round(num)))
            except Exception:
                pass
        # Caso contrário, retornar a string original (útil para valores
        # previamente formatados ou não numéricos)
        return val_str

    rows_html = []
    for _, row in df.iterrows():
        cols_html = ''.join(f'<td>{format_cell(row[col])}</td>' for col in headers)
        rows_html.append(f'      <tr>{cols_html}</tr>')
    # Subtítulo padronizado
    subtitle = 'Distribuição por número de quartos'
    table_html = [
        '<div class="chart-container">',
        f'  <div class="chart-title">{title}</div>',
        f'  <div class="chart-subtitle">{subtitle}</div>',
        '  <div class="table-wrapper">',
        '    <table class="region-table">',
        '      <tr>' + ''.join(f'<th>{h}</th>' for h in headers) + '</tr>',
        *rows_html,
        '    </table>',
        '  </div>',
        '</div>'
    ]
    return '\n'.join(table_html)


def insert_region_tables(html_content: str, region_tables: dict[str, str]) -> str:
    """
    Injeta as tabelas de regiões nas seções corretas do HTML.

    A inserção é feita logo antes do início da próxima seção, garantindo
    que cada tabela apareça apenas na sua respectiva view.
    """
    # Definições de seções e seu próximo id  
    insertion_specs = [
        ('ofertas', 'vendas', region_tables.get('ofertas', '')),
        ('vendas', 'lancamentos', region_tables.get('vendas', '')),
        # ORDEM GARANTIDA: Preços de OFERTA sempre antes de VENDA
        ('precos', 'vgv', ''), # Será preenchido abaixo
    ]
    
    # Garantir ordem específica para tabelas de preços
    precos_content = ""
    if 'precos_oferta' in region_tables:
        precos_content += region_tables['precos_oferta']
    if 'precos_venda' in region_tables:
        precos_content += region_tables['precos_venda']
    
    # Atualizar o insertion_specs com o conteúdo na ordem correta
    for i, (section_id, next_id, content) in enumerate(insertion_specs):
        if section_id == 'precos':
            insertion_specs[i] = (section_id, next_id, precos_content)
    
    new_html = html_content
    
    for section_id, next_id, insertion in insertion_specs:
        if not insertion:
            print(f"⚠️  Nenhuma tabela para inserir na seção '{section_id}'")
            continue
            
        print(f"📋 Inserindo tabela na seção '{section_id}' (antes de '{next_id}')...")
        
        # Localizar início da seção atual
        start_idx = new_html.find(f'<div id="{section_id}"')
        if start_idx == -1:
            print(f"❌ Seção '{section_id}' não encontrada no HTML!")
            continue
            
        # Localizar início da próxima seção
        next_idx = new_html.find(f'<div id="{next_id}"', start_idx + 1)
        if next_idx == -1:
            print(f"⚠️  Próxima seção '{next_id}' não encontrada, inserindo no final da seção")
            next_idx = len(new_html)
            
        # Procurar o último fechamento de </div> entre a seção atual e a próxima
        closing_idx = new_html.rfind('</div>', start_idx, next_idx)
        if closing_idx == -1:
            # fallback: insere antes da próxima seção
            insertion_point = next_idx
        else:
            insertion_point = closing_idx
            
        # Inserir a tabela
        new_html = new_html[:insertion_point] + '\n' + insertion + '\n' + new_html[insertion_point:]
        print(f"✅ Tabela inserida na seção '{section_id}' na posição {insertion_point}")
        
    return new_html


# -------------------------
# Insights (YTD, Pico, Tendência, YoY)
# -------------------------
def _valid(vals: list) -> list:
    return [v for v in vals if v is not None and not (isinstance(v, float) and np.isnan(v))]


def calc_ytd(values: list) -> float | None:
    vals = _valid(values)
    if not vals:
        return None
    return float(np.mean(vals))


def calc_peak(values: list) -> float | None:
    vals = _valid(values)
    if not vals:
        return None
    return float(np.max(vals))


def calc_trend(values: list) -> str:
    """
    Tendência via regressão linear simples (numpy polyfit).
    Retorna: 'Alta', 'Queda' ou 'Estável'.
    """
    vals = _valid(values)
    if len(vals) < 4:
        return "Estável"
    x = np.arange(len(vals), dtype=float)
    y = np.array(vals, dtype=float)
    slope = np.polyfit(x, y, 1)[0]
    # limiar conservador para evitar ruído
    if slope > 0.01:
        return "Alta"
    if slope < -0.01:
        return "Queda"
    return "Estável"


def calc_yoy(current_values: list, prev_values: list, is_percent: bool) -> str:
    """
    Comparativo YoY baseado na média dos meses disponíveis do ano corrente,
    comparando com os mesmos meses do ano anterior (quando existirem).
    """
    cur = current_values
    prev = prev_values
    n = min(len(cur), len(prev))
    if n <= 0:
        return "n/d"

    cur_cut = [cur[i] for i in range(n)]
    prev_cut = [prev[i] for i in range(n)]

    cur_vals = _valid(cur_cut)
    prev_vals = _valid(prev_cut)

    if not cur_vals or not prev_vals:
        return "n/d"

    cur_mean = float(np.mean(cur_vals))
    prev_mean = float(np.mean(prev_vals))

    if is_percent:
        # diferença em pontos percentuais
        diff_pp = cur_mean - prev_mean
        sign = "+" if diff_pp >= 0 else ""
        return f"{sign}{br_float(abs(diff_pp), 1)} p.p." if diff_pp >= 0 else f"{br_float(diff_pp, 1)} p.p."
    else:
        if prev_mean == 0:
            return "n/d"
        chg = (cur_mean / prev_mean - 1.0) * 100.0
        sign = "+" if chg >= 0 else ""
        return f"{sign}{br_float(abs(chg), 1)}%" if chg >= 0 else f"{br_float(chg, 1)}%"


# -------------------------
# Datasets Chart.js
# -------------------------
def build_monthly_dataset(df: pd.DataFrame, is_percent: bool = False) -> dict:
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return {'labels': [], 'datasets': []}

    labels = df_clean.iloc[:, 0].astype(str).tolist()
    datasets = []
    years = [
        c for c in df_clean.columns[1:]
        if re.fullmatch(r"\d{4}", str(c).strip())
    ]

    for idx, year in enumerate(years):
        colour = COLOURS[idx % len(COLOURS)]
        data = []
        for val in df_clean[year]:
            parsed = parse_percentage(val) if is_percent else parse_number(val)
            data.append(parsed)

        is_current = (idx == len(years) - 1)

        dataset = {
            'label': str(year).strip(),
            'data': data,
            'borderColor': colour,
            'backgroundColor': f"rgba({int(colour[1:3],16)}, {int(colour[3:5],16)}, {int(colour[5:7],16)}, 0.10)",
            'borderWidth': 2,                 # igual a todos
            'tension': 0.4,
            'pointRadius': 3,                 # igual a todos
            'pointHoverRadius': 5,
            'pointStyle': 'circle',
        }

        # Ano corrente apenas TRACEJADO
        if is_current:
            dataset['borderDash'] = [6, 4]

        datasets.append(dataset)

    return {'labels': labels, 'datasets': datasets}


def build_quarterly_dataset(df: pd.DataFrame, is_percent: bool = False) -> dict:
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return {'labels': [], 'datasets': []}

    labels = df_clean.iloc[:, 0].astype(str).tolist()
    datasets = []
    for idx, year in enumerate(df_clean.columns[1:]):
        colour = COLOURS[idx % len(COLOURS)]
        data = []
        for val in df_clean[year]:
            parsed = parse_percentage(val) if is_percent else parse_number(val)
            data.append(parsed)

        datasets.append({
            'label': str(year).strip(),
            'data': data,
            'backgroundColor': f"rgba({int(colour[1:3],16)}, {int(colour[3:5],16)}, {int(colour[5:7],16)}, 0.80)",
            'borderColor': colour,
            'borderWidth': 2,
            'pointStyle': 'circle',
        })
    return {'labels': labels, 'datasets': datasets}


def build_yearly_dataset(df: pd.DataFrame, is_percent: bool = False) -> tuple:
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return ({'labels': [], 'datasets': []}, [])

    labels, values, variations = [], [], []

    for _, row in df_clean.iterrows():
        year = str(row.iloc[0]).strip()
        labels.append(year)

        val = parse_percentage(row.iloc[1]) if is_percent else parse_number(row.iloc[1])
        values.append(val)

        if len(row) > 2 and not pd.isna(row.iloc[2]):
            variations.append(str(row.iloc[2]).strip())
        else:
            variations.append('-')

    colors = [COLOURS[i % len(COLOURS)] for i in range(len(values))]

    data = {
        'labels': labels,
        'datasets': [{
            'label': 'Valor',
            'data': values,
            'backgroundColor': colors,
            'borderColor': colors,
            'borderWidth': 1,
            'pointStyle': 'circle',
        }]
    }
    return data, variations

# === Novos construtores para números entre colchetes (empreendimentos) ===
def build_monthly_dataset_bracket(df: pd.DataFrame) -> dict:
    """
    Constrói dataset mensal para valores de empreendimentos (números entre colchetes).
    O DataFrame deve ter a primeira coluna com labels (meses) e as demais com anos.
    """
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return {'labels': [], 'datasets': []}

    labels = [str(v).strip() for v in df_clean.iloc[:, 0]]
    datasets = []
    for idx, col in enumerate(df_clean.columns[1:]):
        values = [parse_bracket_number(v) for v in df_clean[col]]
        color = COLOURS[idx % len(COLOURS)]
        # Definir propriedades semelhantes aos demais gráficos de linha
        dataset = {
            'label': str(col),
            'data': values,
            'borderColor': color,
            # utilizar leve transparência para área preenchida
            'backgroundColor': f"rgba({int(color[1:3],16)}, {int(color[3:5],16)}, {int(color[5:7],16)}, 0.10)",
            'borderWidth': 2,
            'tension': 0.4,
            'pointRadius': 3,
            'pointHoverRadius': 5,
            'pointStyle': 'circle',
        }
        # Ano corrente (última série) em tracejado
        if idx == len(df_clean.columns[1:]) - 1:
            dataset['borderDash'] = [6, 4]
        datasets.append(dataset)
    return {'labels': labels, 'datasets': datasets}

def build_quarterly_dataset_bracket(df: pd.DataFrame) -> dict:
    """
    Constrói dataset trimestral para valores de empreendimentos.
    """
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return {'labels': [], 'datasets': []}

    labels = [str(v).strip() for v in df_clean.iloc[:, 0]]
    datasets = []
    for idx, col in enumerate(df_clean.columns[1:]):
        values = [parse_bracket_number(v) for v in df_clean[col]]
        color = COLOURS[idx % len(COLOURS)]
        datasets.append({
            'label': str(col),
            'data': values,
            # usar cores com transparência semelhante aos demais gráficos de barras
            'backgroundColor': f"rgba({int(color[1:3],16)}, {int(color[3:5],16)}, {int(color[5:7],16)}, 0.80)",
            'borderColor': color,
            'borderWidth': 2,
            'pointStyle': 'circle',
        })
    return {'labels': labels, 'datasets': datasets}

def build_yearly_dataset_bracket(df: pd.DataFrame) -> tuple:
    """
    Constrói dataset anual e variações para empreendimentos (números entre colchetes).
    Retorna (data, variations)
    """
    df_clean = clean_dataframe(df)
    if df_clean.empty:
        return ({'labels': [], 'datasets': []}, [])
    labels, values, variations = [], [], []
    for _, row in df_clean.iterrows():
        year = str(row.iloc[0]).strip()
        labels.append(year)
        val = parse_bracket_number(row.iloc[1])
        values.append(val)
        if len(row) > 2 and not pd.isna(row.iloc[2]):
            variations.append(str(row.iloc[2]).strip())
        else:
            variations.append('-')
    colors = [COLOURS[i % len(COLOURS)] for i in range(len(values))]
    data = {
        'labels': labels,
        'datasets': [{
            'label': 'Valor',
            'data': values,
            'backgroundColor': colors,
            'borderColor': colors,
            'borderWidth': 1,
            'pointStyle': 'circle',
        }]
    }
    return data, variations


# -------------------------
# HTML generation
# -------------------------
# -------------------------
# Função para extrair valores mais atuais para o card de resumo
# -------------------------
def extract_summary_values(data_dict, highlights):
    """Extrai os valores mais atuais de cada métrica para o card de resumo."""
    summary = {}
    
    def get_trend_arrow(data_key):
        """
        Extrai seta de tendência baseada nos highlights.
        Se houver setas (🟢, 🔴 ou 🟡️) no texto de trend dos highlights, retorna essa seta.
        Caso contrário, retorna string vazia.
        """
        trend = highlights.get(f'{data_key} Trend', '')
        if '🟢' in trend:
            return '🟢'
        if '🔴' in trend:
            return '🔴'
        if '🟡️' in trend:
            return '🟡️'
        return ''

    def compute_arrow_from_series(series):
        """
        Dada uma lista de valores (possivelmente contendo None),
        encontra os dois últimos valores não nulos e retorna uma
        seta de tendência comparando o valor mais recente ao anterior.

        Retorna:
          '🟢' se o último valor for maior que o penúltimo;
          '🔴' se o último valor for menor que o penúltimo;
          '🟡️' se forem iguais;
          ''  se não houver dados suficientes.
        """
        # Filtrar valores válidos preservando a ordem (evitar None)
        valid = [v for v in series if v is not None and not (isinstance(v, float) and np.isnan(v))]
        if len(valid) < 2:
            return ''
        last = valid[-1]
        prev = valid[-2]
        try:
            if last > prev:
                return '🟢'
            elif last < prev:
                return '🔴'
            else:
                return '🟡️'
        except Exception:
            return ''
    
    # IVV
    if 'IVV Monthly' in data_dict:
        ivv_data = data_dict['IVV Monthly']
        if ivv_data['datasets']:
            last_dataset = ivv_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['ivv'] = br_percent(last_dataset['data'][i])
                    # Calcular seta comparando com mês anterior
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    # Se não houver dados suficientes, usar setas dos highlights
                    summary['ivv_trend'] = arrow if arrow else get_trend_arrow('IVV')
                    break
    
    # Unidades ofertadas
    if 'Ofertas Monthly' in data_dict:
        ofertas_data = data_dict['Ofertas Monthly']
        if ofertas_data['datasets']:
            last_dataset = ofertas_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['ofertas'] = br_int(last_dataset['data'][i])
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['ofertas_trend'] = arrow if arrow else get_trend_arrow('Ofertas')
                    break
    
    # Unidades vendidas
    if 'Vendas Monthly' in data_dict:
        vendas_data = data_dict['Vendas Monthly']
        if vendas_data['datasets']:
            last_dataset = vendas_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['vendas'] = br_int(last_dataset['data'][i])
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['vendas_trend'] = arrow if arrow else get_trend_arrow('Vendas')
                    break
    
    # Unidades lançadas
    lancamentos_valor = None
    empreendimentos_valor = None
    
    if 'Lanc Monthly' in data_dict:
        lanc_data = data_dict['Lanc Monthly']
        if lanc_data['datasets']:
            last_dataset = lanc_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    lancamentos_valor = last_dataset['data'][i]
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['lancamentos_trend'] = arrow if arrow else get_trend_arrow('Lanc')
                    break
    
    # Número de empreendimentos lançados
    if 'Lanc Proj Monthly' in data_dict:
        lanc_proj_data = data_dict['Lanc Proj Monthly']
        if lanc_proj_data['datasets']:
            last_dataset = lanc_proj_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    empreendimentos_valor = int(last_dataset['data'][i])
                    break
    
    # Montar texto dos lançamentos
    if lancamentos_valor is not None and empreendimentos_valor is not None:
        if lancamentos_valor == 0:
            summary['lancamentos'] = f"- [{empreendimentos_valor}]"
        else:
            summary['lancamentos'] = f"{br_int(lancamentos_valor)} [{empreendimentos_valor}]"
    elif lancamentos_valor is not None:
        if lancamentos_valor == 0:
            summary['lancamentos'] = "-"
        else:
            summary['lancamentos'] = br_int(lancamentos_valor)
    else:
        summary['lancamentos'] = 'n/d'
    
    # Oferta em m²
    if 'OfertaM2 Monthly' in data_dict:
        oferta_m2_data = data_dict['OfertaM2 Monthly']
        if oferta_m2_data['datasets']:
            last_dataset = oferta_m2_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['oferta_m2'] = f"{br_int(last_dataset['data'][i])} m²"
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['oferta_m2_trend'] = arrow if arrow else get_trend_arrow('OfertaM2')
                    break
    
    # Venda em m²
    if 'VendaM2 Monthly' in data_dict:
        venda_m2_data = data_dict['VendaM2 Monthly']
        if venda_m2_data['datasets']:
            last_dataset = venda_m2_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['venda_m2'] = f"{br_int(last_dataset['data'][i])} m²"
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['venda_m2_trend'] = arrow if arrow else get_trend_arrow('VendaM2')
                    break
    
    # Preço de oferta
    if 'Precos Oferta Monthly' in data_dict:
        preco_oferta_data = data_dict['Precos Oferta Monthly']
        if preco_oferta_data['datasets']:
            last_dataset = preco_oferta_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['preco_oferta'] = br_currency(last_dataset['data'][i], 0)
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['preco_oferta_trend'] = arrow if arrow else get_trend_arrow('PrecosOferta')
                    break
    
    # Preço de venda
    if 'Precos Venda Monthly' in data_dict:
        preco_venda_data = data_dict['Precos Venda Monthly']
        if preco_venda_data['datasets']:
            last_dataset = preco_venda_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['preco_venda'] = br_currency(last_dataset['data'][i], 0)
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['preco_venda_trend'] = arrow if arrow else get_trend_arrow('PrecosVenda')
                    break
    
    # VGV
    if 'VGV Monthly' in data_dict:
        vgv_data = data_dict['VGV Monthly']
        if vgv_data['datasets']:
            last_dataset = vgv_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['vgv'] = f"{br_currency(last_dataset['data'][i], 0)}M"
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['vgv_trend'] = arrow if arrow else get_trend_arrow('VGV')
                    break
    
    # VGL
    if 'VGL Monthly' in data_dict:
        vgl_data = data_dict['VGL Monthly']
        if vgl_data['datasets']:
            last_dataset = vgl_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['vgl'] = f"{br_currency(last_dataset['data'][i], 0)}M"
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['vgl_trend'] = arrow if arrow else get_trend_arrow('VGL')
                    break
    
    # Distratos
    if 'Distratos Monthly' in data_dict:
        distratos_data = data_dict['Distratos Monthly']
        if distratos_data['datasets']:
            last_dataset = distratos_data['datasets'][-1]
            for i in range(len(last_dataset['data']) - 1, -1, -1):
                if last_dataset['data'][i] is not None:
                    summary['distratos'] = br_int(last_dataset['data'][i])
                    arrow = compute_arrow_from_series(last_dataset['data'])
                    summary['distratos_trend'] = arrow if arrow else get_trend_arrow('Distratos')
                    break
    
    # Valores padrão se não encontrados
    summary.setdefault('ivv', 'n/d')
    summary.setdefault('ofertas', 'n/d')
    summary.setdefault('vendas', 'n/d')
    summary.setdefault('lancamentos', 'n/d')
    summary.setdefault('oferta_m2', 'n/d')
    summary.setdefault('venda_m2', 'n/d')
    summary.setdefault('preco_oferta', 'n/d')
    summary.setdefault('preco_venda', 'n/d')
    summary.setdefault('vgv', 'n/d')
    summary.setdefault('vgl', 'n/d')
    summary.setdefault('distratos', 'n/d')
    
    # Se a seta de tendência estiver vazia após o cálculo, considerar
    # estável (➡️) por padrão. Isto evita exibir caixas vazias.
    trend_keys = [
        'ivv_trend', 'ofertas_trend', 'vendas_trend', 'lancamentos_trend',
        'oferta_m2_trend', 'venda_m2_trend', 'preco_oferta_trend',
        'preco_venda_trend', 'vgv_trend', 'vgl_trend', 'distratos_trend'
    ]
    for key in trend_keys:
        if not summary.get(key):
            summary[key] = '➡️'
    
    return summary


# -------------------------
# Novas funções para extrair dados MoM/YoY e insights avançados
# -------------------------

def extract_mom_yoy_from_sheet(df: pd.DataFrame):
    """Extrai dados MoM e YoY das últimas linhas do Excel no formato real."""
    if df.empty:
        return None, None, None, None
    
    # Procurar pela linha "Variações" no DataFrame original (não limpo)
    for idx in range(len(df)):
        row = df.iloc[idx]
        first_cell = str(row.iloc[0]).lower() if not pd.isna(row.iloc[0]) else ""
        
        if 'variações' in first_cell or 'variacao' in first_cell:
            # Extrair o texto da segunda coluna
            variacao_texto = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            
            if variacao_texto and variacao_texto != 'nan':
                # Extrair MoM e YoY usando regex
                import re
                
                # Formato: "Out/2025 - Set/2025: 1,3% | Out/2025 - Out/2024: 48,4%"
                mom_match = re.search(r'(\w+/\d+)\s*-\s*(\w+/\d+):\s*([-+]?[\d,]+%)', variacao_texto)
                
                if '|' in variacao_texto:
                    yoy_part = variacao_texto.split('|')[1].strip()
                    yoy_match = re.search(r'(\w+/\d+)\s*-\s*(\w+/\d+):\s*([-+]?[\d,]+%)', yoy_part)
                else:
                    yoy_match = None
                
                mom_value = None
                yoy_value = None
                mom_label = None
                yoy_label = None
                
                if mom_match:
                    mom_label = f"{mom_match.group(1)} - {mom_match.group(2)}"
                    mom_percent = mom_match.group(3).replace(',', '.').replace('%', '')
                    try:
                        mom_value = float(mom_percent)
                    except ValueError:
                        mom_value = None
                
                if yoy_match:
                    yoy_label = f"{yoy_match.group(1)} - {yoy_match.group(2)}"
                    yoy_percent = yoy_match.group(3).replace(',', '.').replace('%', '')
                    try:
                        yoy_value = float(yoy_percent)
                    except ValueError:
                        yoy_value = None
                
                return mom_value, yoy_value, mom_label, yoy_label
    
    return None, None, None, None


def find_peak_with_month(df: pd.DataFrame, data_type='number'):
    """Encontra o pico e o mês correspondente."""
    df_clean = clean_dataframe(df)
    if df_clean.empty or df_clean.shape[1] < 2:
        return None, None
    
    # Usar apenas as linhas de dados (excluir variações, observações etc)
    data_rows = []
    for idx in range(len(df_clean)):
        first_cell = str(df_clean.iloc[idx, 0]).lower()
        if first_cell not in ['nan', 'variações', 'observação', 'variacao', 'observacao'] and not pd.isna(df_clean.iloc[idx, 0]):
            data_rows.append(idx)
    
    if not data_rows:
        return None, None
    
    labels = [str(df_clean.iloc[idx, 0]) for idx in data_rows]
    last_col_idx = len(df_clean.columns) - 1  # Index da última coluna
    
    if data_type == 'percent':
        values = [parse_percentage(df_clean.iloc[idx, last_col_idx]) for idx in data_rows]
    else:
        values = [parse_number(df_clean.iloc[idx, last_col_idx]) for idx in data_rows]
    
    peak_value = None
    peak_month = None
    for i, val in enumerate(values):
        if val is not None:
            if peak_value is None or val > peak_value:
                peak_value = val
                peak_month = labels[i] if i < len(labels) else 'n/d'
    
    return peak_value, peak_month


def calculate_averages(df: pd.DataFrame, data_type='number'):
    """Calcula médias trimestral e anual com base nos dados reais."""
    df_clean = clean_dataframe(df)
    if df_clean.empty or df_clean.shape[1] < 2:
        return None, None
    
    # Usar apenas as linhas de dados (excluir variações, observações etc)
    data_rows = []
    for idx in range(len(df_clean)):
        first_cell = str(df_clean.iloc[idx, 0]).lower()
        if first_cell not in ['nan', 'variações', 'observação', 'variacao', 'observacao'] and not pd.isna(df_clean.iloc[idx, 0]):
            data_rows.append(idx)
    
    if not data_rows:
        return None, None
    
    last_col_idx = len(df_clean.columns) - 1  # Index da última coluna
    
    if data_type == 'percent':
        values = [parse_percentage(df_clean.iloc[idx, last_col_idx]) for idx in data_rows]
    else:
        values = [parse_number(df_clean.iloc[idx, last_col_idx]) for idx in data_rows]
    
    valid_values = [v for v in values if v is not None]
    
    if not valid_values:
        return None, None
    
    # Média trimestral: últimos 3 meses válidos
    quarterly_avg = sum(valid_values[-3:]) / len(valid_values[-3:]) if len(valid_values) >= 3 else None
    
    # Média anual: todos os meses válidos disponíveis no ano atual
    yearly_avg = sum(valid_values) / len(valid_values) if valid_values else None
    
    return quarterly_avg, yearly_avg


def extract_observation_from_sheet(df: pd.DataFrame):
    """Extrai observações sobre dados incompletos das planilhas trimestrais/anuais."""
    if df.empty:
        return None
    
    # Procurar pela linha "Observação" no DataFrame original
    for idx in range(len(df)):
        row = df.iloc[idx]
        first_cell = str(row.iloc[0]).lower() if not pd.isna(row.iloc[0]) else ""
        
        if 'observação' in first_cell or 'observacao' in first_cell:
            # Extrair o texto da segunda coluna
            observacao_texto = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            
            if observacao_texto and observacao_texto != 'nan':
                return observacao_texto
    
    return None


def find_best_quarter_with_performance(df: pd.DataFrame, data_type='number'):
    """Encontra o trimestre com MELHOR performance (maior valor)."""
    df_clean = clean_dataframe(df)
    if df_clean.empty or df_clean.shape[1] < 2:
        return None, None
    
    # Usar apenas as linhas de dados trimestrais (1T, 2T, 3T, 4T)
    quarter_rows = []
    for idx in range(len(df_clean)):
        first_cell = str(df_clean.iloc[idx, 0]).upper()
        if first_cell in ['1T', '2T', '3T', '4T']:
            quarter_rows.append((idx, first_cell))
    
    if not quarter_rows:
        return None, None
    
    last_col_idx = len(df_clean.columns) - 1  # Index da última coluna
    
    best_value = None
    best_quarter = None
    
    for idx, quarter_label in quarter_rows:
        if data_type == 'percent':
            value = parse_percentage(df_clean.iloc[idx, last_col_idx])
        else:
            value = parse_number(df_clean.iloc[idx, last_col_idx])
        
        if value is not None:
            if best_value is None or value > best_value:
                best_value = value
                best_quarter = quarter_label
    
    return best_value, best_quarter


def format_new_insights(df: pd.DataFrame, data_type='number', is_millions=False, month_ref=''):
    """
    Formata os novos insights no padrão solicitado usando dados reais do Excel.
    data_type: 'percent', 'number', 'currency'
    """
    if df.empty or df.shape[1] < 2:
        return {
            'mom': 'n/d',
            'yoy': 'n/d', 
            'peak': 'n/d',
            'yearly_avg': 'n/d'
        }
    
    # Extrair MoM e YoY do Excel
    mom_value, yoy_value, mom_label, yoy_label = extract_mom_yoy_from_sheet(df)
    
    # Encontrar pico
    peak_value, peak_month = find_peak_with_month(df, data_type)
    
    # Calcular apenas média anual (sem trimestral)
    _, yearly_avg = calculate_averages(df, data_type)
    
    # Função para formatar valores
    def format_value(val, include_month=False, month=''):
        if val is None:
            return 'n/d'
        
        if data_type == 'percent':
            formatted = br_percent(val)
        elif data_type == 'currency':
            formatted = br_currency(val, 0)
        else:  # number
            formatted = br_int(val)
        
        if include_month and month:
            formatted += f" ({month})"
            
        return formatted
    
    # Montar strings de resultado
    mom_str = f"Variação MoM: {mom_label}: {br_percent(mom_value)}" if mom_value is not None and mom_label else "Variação MoM: n/d"
    yoy_str = f"Variação YoY: {yoy_label}: {br_percent(yoy_value)}" if yoy_value is not None and yoy_label else "Variação YoY: n/d"
    
    peak_str = format_value(peak_value, True, peak_month)
    yearly_avg_str = format_value(yearly_avg)
    
    # Adicionar prefixos apropriados
    if data_type == 'currency' and is_millions:
        peak_label = f"Pico (R$ Mi): {peak_str}"
        yearly_label = f"Média anual (R$ Mi): {yearly_avg_str}"
    elif data_type == 'currency':
        peak_label = f"Pico: {peak_str}"
        yearly_label = f"Média anual: {yearly_avg_str}"
    else:
        peak_label = f"Pico: {peak_str}"
        yearly_label = f"Média anual: {yearly_avg_str}"
    
    return {
        'mom': mom_str,
        'yoy': yoy_str,
        'peak': peak_label,
        'yearly_avg': yearly_label
    }


def extract_month_ref(filename: str) -> str:
    """Ex.: Relatorio_Completo_Residencial_2025_08.xlsx -> Ago/25"""
    pattern = r'(\d{4})_(\d{2})'
    match = re.search(pattern, filename)
    months = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun',
              'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    if match:
        year = match.group(1)
        month_num = int(match.group(2))
        if 1 <= month_num <= 12:
            return f"{months[month_num - 1]}/{year[-2:]}"
    now = datetime.now()
    return f"{months[now.month - 1]}/{str(now.year)[-2:]}"


def _to_js_json(obj):
    def clean_obj(o):
        if isinstance(o, dict):
            return {k: clean_obj(v) for k, v in o.items()}
        if isinstance(o, list):
            return [clean_obj(x) for x in o]
        if isinstance(o, float) and (pd.isna(o) or np.isnan(o)):
            return None
        return o
    return json.dumps(clean_obj(obj), ensure_ascii=False, default=str)


def generate_html(data_dict: dict, report_date: str, month_ref: str, highlights: dict) -> str:
    # Extrair valores para o card de resumo
    summary = extract_summary_values(data_dict, highlights)
    
    # Helpers de legendas (bolinhas)
    # Obs: aplicado em TODOS os charts via config base.
    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Dashboard IVV - Apresentação Executiva</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/3.9.1/chart.min.js"></script>
  <style>
    * {{ margin:0; padding:0; box-sizing:border-box; }}
    body {{ font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif; background:linear-gradient(135deg,#478DDC 0%,#2c5aa0 100%); color:#333; line-height:1.6; }}
    .container {{ max-width:1400px; margin:0 auto; padding:20px; }}
    .nav-bar {{ background:rgba(255,255,255,0.1); backdrop-filter:blur(10px); border-radius:15px; padding:20px; margin-bottom:30px; text-align:center; }}
    .nav-buttons {{ display:flex; justify-content:center; gap:15px; flex-wrap:wrap; }}
    .nav-btn {{ background:rgba(255,255,255,0.2); color:white; padding:12px 24px; border:none; border-radius:25px; cursor:pointer; font-size:14px; font-weight:600; transition:all .3s ease; text-decoration:none; display:inline-block; }}
    .nav-btn:hover {{ background:rgba(255,255,255,0.3); transform:translateY(-2px); box-shadow:0 5px 15px rgba(0,0,0,0.2); }}
    .nav-btn.active {{ background:#27ae60; box-shadow:0 5px 15px rgba(39,174,96,0.3); }}
    .section {{ display:none; }}
    .section.active {{ display:block; }}
    .header {{ text-align:center; color:white; margin-bottom:40px; padding:30px 0; position:relative; }}
    .header-content {{ display:flex; align-items:center; justify-content:flex-start; gap:30px; padding:0 40px; }}
    .logo {{ height:100px; width:auto; filter:drop-shadow(2px 2px 4px rgba(0,0,0,0.3)) drop-shadow(0 0 8px rgba(255,255,255,0.8)); flex-shrink:0; }}
    .header-text {{ flex-grow:1; text-align:center; }}
    .month-ref {{ position:absolute; top:20px; right:40px; background:rgba(255,255,255,0.2); padding:8px 15px; border-radius:20px; font-weight:bold; font-size:14px; backdrop-filter:blur(10px); }}
    .header h1 {{ font-size:2.5em; margin-bottom:10px; text-shadow:2px 2px 4px rgba(0,0,0,0.3); }}
    .header p {{ font-size:1.2em; opacity:.9; }}
    .chart-container {{ background:white; border-radius:15px; padding:30px; margin-bottom:30px; box-shadow:0 10px 30px rgba(0,0,0,0.2); backdrop-filter:blur(10px); }}
    .chart-title {{ font-size:1.8em; margin-bottom:20px; color:#2c3e50; text-align:center; font-weight:600; }}
    .chart-subtitle {{ font-size:1em; color:#7f8c8d; text-align:center; margin-bottom:30px; }}
    .chart-wrapper {{ position:relative; height:400px; margin-bottom:20px; }}
    .chart-wrapper.small {{ height:300px; }}
    .insights {{ background:#f8f9fa; border-radius:10px; padding:20px; margin-top:20px; border-left:5px solid #3498db; }}
    .insights h4 {{ color:#2c3e50; margin-bottom:10px; }}
    .insights ul {{ list-style-type:none; padding-left:0; }}
    .insights li {{ padding:5px 0; color:#555; }}
    .insights li:before {{ content:"▸ "; color:#3498db; font-weight:bold; }}
    .grid {{ display:grid; grid-template-columns:1fr 1fr; gap:30px; margin-bottom:30px; }}
    /* ===== RESPONSIVIDADE MOBILE COMPLETA ===== */
    
    /* Tablets e telas médias */
    @media (max-width: 1024px) {{
      .container {{ padding: 15px; }}
      .header h1 {{ font-size: 2.2em; }}
      .summary-card {{ padding: 30px; }}
      .metrics-grid {{ grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }}
      .chart-container {{ padding: 25px; }}
    }}
    
    /* Tablets em retrato e smartphones grandes */
    @media (max-width: 768px) {{
      .grid {{ grid-template-columns: 1fr; }}
      .header h1 {{ font-size: 2em; }}
      .chart-container {{ padding: 20px; }}
      .nav-buttons {{ gap: 10px; }}
      .nav-btn {{ 
        padding: 14px 20px; 
        font-size: 15px; 
        min-height: 48px; /* Touch-friendly */ 
        border-radius: 30px;
      }}
      .summary-card {{ 
        padding: 25px; 
        border-radius: 15px; 
      }}
      .summary-title {{ 
        font-size: 1.8em; 
        margin-bottom: 25px; 
      }}
      .metrics-grid {{ 
        grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); 
        gap: 12px; 
      }}
      .metric-item {{ 
        padding: 15px; 
        border-radius: 12px; 
      }}
      .metric-value {{ 
        font-size: 1.4em; 
      }}
      .metric-label {{ 
        font-size: 0.8em; 
      }}
      .chart-title {{ 
        font-size: 1.3em; 
      }}
      .chart-subtitle {{ 
        font-size: 0.9em; 
      }}
    }}
    
    /* Smartphones */
    @media (max-width: 480px) {{
      .container {{ 
        padding: 10px; 
        margin: 0; 
      }}
      .nav-bar {{ 
        padding: 15px; 
        margin-bottom: 20px; 
        border-radius: 12px; 
      }}
      .nav-buttons {{ 
        flex-direction: column; 
        gap: 8px; 
      }}
      .nav-btn {{ 
        width: 100%; 
        padding: 16px; 
        font-size: 16px; 
        min-height: 52px; /* Ainda mais touch-friendly */
        border-radius: 25px;
        text-align: center;
      }}
      .header {{ 
        padding: 20px; 
        border-radius: 12px; 
      }}
      .header h1 {{ 
        font-size: 1.6em; 
        line-height: 1.2; 
      }}
      .logo {{ 
        height: 70px; 
        margin-right: 10px; 
      }}
      .summary-card {{ 
        padding: 20px; 
        border-radius: 12px; 
        margin-bottom: 20px; 
      }}
      .summary-title {{ 
        font-size: 1.5em; 
        margin-bottom: 20px; 
      }}
      .metrics-grid {{ 
        grid-template-columns: 1fr; 
        gap: 10px; 
      }}
      .metric-item {{ 
        padding: 12px; 
        border-radius: 10px; 
      }}
      .metric-value {{ 
        font-size: 1.2em; 
      }}
      .metric-label {{ 
        font-size: 0.75em; 
      }}
      .chart-container {{ 
        padding: 15px; 
        margin-bottom: 20px; 
        border-radius: 12px; 
      }}
      .chart-title {{ 
        font-size: 1.1em; 
        margin-bottom: 8px; 
      }}
      .chart-subtitle {{ 
        font-size: 0.8em; 
      }}
      .chart-wrapper {{ 
        height: 280px; /* Altura otimizada para mobile */
      }}
      .chart-wrapper.small {{ 
        height: 220px; 
      }}
      .insights {{ 
        padding: 12px; 
        border-radius: 8px; 
      }}
      .insights h4 {{ 
        font-size: 1em; 
        margin-bottom: 8px; 
      }}
      .insights ul {{ 
        font-size: 0.85em; 
      }}
      .highlight-box {{ 
        padding: 12px; 
        border-radius: 8px; 
        margin: 15px 0; 
      }}
      .highlight-box h3 {{ 
        font-size: 1.2em; 
      }}
      .region-table th, .region-table td {{ 
        padding: 6px 8px; 
        font-size: 0.85em; 
        min-width: 60px; /* Reduzir largura mínima em mobile */
      }}
      
      .region-table th:first-child,
      .region-table td:first-child {{
        min-width: 90px; /* Largura reduzida para região em mobile */
      }}
      
      .table-wrapper {{
        border-radius: 6px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.1);
      }}
      
      .region-table {{
        min-width: 500px; /* Largura mínima reduzida para mobile */
      }}
    }}
    
    /* Smartphones pequenos */
    @media (max-width: 360px) {{
      .container {{ padding: 8px; }}
      .header h1 {{ font-size: 1.4em; }}
      .nav-btn {{ font-size: 15px; padding: 14px; }}
      .summary-title {{ font-size: 1.3em; }}
      .metric-value {{ font-size: 1.1em; }}
      .chart-wrapper {{ height: 250px; }}
      .chart-wrapper.small {{ height: 200px; }}
      .logo {{ height: 60px; }}
    }}
    
    /* ===== OTIMIZAÇÕES PARA APRESENTAÇÃO MOBILE ===== */
    @media (max-width: 768px) {{
      #presentationContainer .slide {{
        padding: 20px 10px;
      }}
      
      #presentationContainer .chart-container {{
        padding: 15px;
        margin-bottom: 15px;
      }}
      
      #presentationContainer .chart-title {{
        font-size: 1.2em;
        margin-bottom: 8px;
      }}
      
      #presentationContainer .chart-subtitle {{
        font-size: 0.85em;
      }}
      
      #presentationContainer .chart-wrapper {{
        height: 60vh; /* Usa viewport height para mobile */
        max-height: 400px;
      }}
      
      #presentationContainer .highlight-box {{
        padding: 10px;
        margin: 10px 0;
      }}
      
      #presentationContainer .insights {{
        padding: 10px;
        font-size: 0.9em;
      }}
      
      /* Navegação da apresentação otimizada para touch */
      #presentationContainer .slide.active {{
        overflow-y: auto; /* Permite scroll se necessário */
        -webkit-overflow-scrolling: touch; /* Smooth scroll no iOS */
      }}
    }}
    
    /* ===== MELHORIAS TOUCH-FRIENDLY ===== */
    
    /* Aumentar área de toque para todos os botões */
    .nav-btn, button {{
      min-height: 44px; /* Padrão de acessibilidade Apple/Google */
      min-width: 44px;
    }}
    
    /* OCULTAR APRESENTAÇÃO EM MOBILE */
    @media (max-width: 768px) {{
      #presentationButton {{
        display: none !important; /* Oculta completamente o botão em mobile */
      }}
    }}
    
    /* Melhorar legibilidade em telas pequenas */
    @media (max-width: 768px) {{
      body {{
        font-size: 16px; /* Evita zoom no iOS */
        line-height: 1.5;
      }}
      
      /* Evitar zoom em inputs (se houver) */
      input, select, textarea {{
        font-size: 16px;
      }}
    }}
    
    /* ===== ORIENTAÇÃO ===== */
    @media screen and (orientation: landscape) and (max-height: 500px) {{
      .header h1 {{ font-size: 1.3em; }}
      .nav-btn {{ padding: 8px 16px; }}
      .chart-wrapper {{ height: 250px; }}
      .summary-card {{ padding: 15px; }}
    }}
    .highlight-box {{ background:linear-gradient(45deg,#27ae60,#2ecc71); color:white; padding:15px; border-radius:10px; text-align:center; margin:20px 0; }}
    .highlight-box h3 {{ font-size:1.5em; margin-bottom:10px; }}
    
    /* Estilos para o card de resumo */
    .summary-card {{ background:linear-gradient(135deg,#ffffff 0%,#f8f9fa 100%); border-radius:20px; padding:40px; margin-bottom:30px; box-shadow:0 15px 35px rgba(0,0,0,0.1); }}
    .summary-title {{ font-size:2.2em; color:#2c3e50; text-align:center; margin-bottom:30px; font-weight:700; }}
    .metrics-grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(250px, 1fr)); gap:20px; }}
    .metric-item {{ background:white; border-radius:15px; padding:20px; text-align:center; box-shadow:0 8px 25px rgba(0,0,0,0.08); border-left:5px solid #3498db; transition:transform 0.3s ease; position:relative; }}
    .metric-item:hover {{ transform:translateY(-5px); box-shadow:0 12px 30px rgba(0,0,0,0.15); }}
    .metric-label {{ font-size:0.85em; color:#7f8c8d; font-weight:600; margin-bottom:8px; text-transform:uppercase; letter-spacing:0.5px; }}
    .metric-value {{ font-size:1.6em; color:#2c3e50; font-weight:700; margin-bottom:5px; }}
    .metric-trend {{ position:absolute; top:15px; right:15px; font-size:1.2em; }}
    .metric-item:nth-child(1) {{ border-left-color:#e74c3c; }}
    .metric-item:nth-child(2) {{ border-left-color:#f39c12; }}
    .metric-item:nth-child(3) {{ border-left-color:#9b59b6; }}
    .metric-item:nth-child(4) {{ border-left-color:#3498db; }}
    .metric-item:nth-child(5) {{ border-left-color:#27ae60; }}
    .metric-item:nth-child(6) {{ border-left-color:#e67e22; }}
    .metric-item:nth-child(7) {{ border-left-color:#1abc9c; }}
    .metric-item:nth-child(8) {{ border-left-color:#95a5a6; }}
    .metric-item:nth-child(9) {{ border-left-color:#34495e; }}
    .metric-item:nth-child(10) {{ border-left-color:#8e44ad; }}
    .metric-item:nth-child(11) {{ border-left-color:#d35400; }}
    /* Estilos para as tabelas regionais */
    .table-wrapper {{
      width: 100%;
      overflow-x: auto;
      -webkit-overflow-scrolling: touch; /* Smooth scrolling no iOS */
      border-radius: 8px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      margin-top: 15px;
      position: relative;
    }}
    
    /* Indicador visual de scroll horizontal */
    .table-wrapper::-webkit-scrollbar {{
      height: 6px;
    }}
    
    .table-wrapper::-webkit-scrollbar-track {{
      background: #f1f1f1;
      border-radius: 3px;
    }}
    
    .table-wrapper::-webkit-scrollbar-thumb {{
      background: #c1c1c1;
      border-radius: 3px;
    }}
    
    .table-wrapper::-webkit-scrollbar-thumb:hover {{
      background: #a1a1a1;
    }}
    
    /* Sombra para indicar mais conteúdo à direita */
    .table-wrapper::after {{
      content: '';
      position: absolute;
      top: 0;
      right: 0;
      bottom: 6px; /* Espaço para scrollbar */
      width: 20px;
      background: linear-gradient(to left, rgba(255,255,255,0.8), transparent);
      pointer-events: none;
      border-radius: 0 8px 8px 0;
    }}
    
    .region-table {{
      width: 100%;
      min-width: 600px; /* Garante largura mínima para não comprimir */
      border-collapse: collapse;
      margin: 0; /* Remove margin para ficar dentro do wrapper */
      background: white;
    }}
    
    .region-table th, .region-table td {{
      padding: 8px 12px;
      border: 1px solid #ddd;
      text-align: center;
      white-space: nowrap; /* Evita quebra de linha */
      min-width: 80px; /* Largura mínima das colunas */
    }}
    
    .region-table th {{
      background-color: #f2f2f2;
      font-weight: 600;
      position: sticky; /* Cabeçalho fixo no scroll horizontal */
      top: 0;
      z-index: 10;
    }}
    
    .region-table tr:nth-child(even) {{
      background-color: #f9f9f9;
    }}
    
    .region-table tr:hover {{
      background-color: #f5f5f5;
    }}
    
    /* Primeira coluna (Região) fixa no scroll horizontal */
    .region-table th:first-child,
    .region-table td:first-child {{
      position: sticky;
      left: 0;
      background-color: white !important; /* Força fundo branco sempre */
      z-index: 11;
      box-shadow: 2px 0 4px rgba(0,0,0,0.1);
      min-width: 120px; /* Largura adequada para nomes de região */
    }}
    
    .region-table th:first-child {{
      background-color: #f2f2f2 !important; /* Força fundo do cabeçalho */
      z-index: 12; /* Maior que as células para ficar por cima */
    }}
    
    /* Força fundo branco em linhas alternadas */
    .region-table tr:nth-child(even) td:first-child {{
      background-color: white !important;
    }}
    
    .region-table tr:hover td:first-child {{
      background-color: #f5f5f5 !important;
    }}

    /* Slide Presentation Styles */
    #presentationContainer {{
      display: none;
      position: fixed;
      top: 0;
      left: 0;
      width: 100vw;
      height: 100vh;
      background: linear-gradient(135deg,#478DDC 0%,#2c5aa0 100%);
      z-index: 9999;
      overflow-y: auto;
      padding: 40px;
    }}

    #presentationContainer .slide {{
      display: none;
      height: 100%;
      /* permitir rolagem interna quando conteúdo excede a altura */
      overflow-y: auto;
    }}

    #presentationContainer .slide.active {{
      display: block;
    }}

    /* Ajuste para tabelas em modo apresentação */
    #presentationContainer .table-wrapper {{
      max-height: 70vh;
      overflow-y: auto;
      overflow-x: auto;
      -webkit-overflow-scrolling: touch;
      border-radius: 6px;
      background: white;
    }}
    
    #presentationContainer .region-table {{
      max-height: none;
      min-width: 500px;
      font-size: 14px;
      margin: 0;
    }}
    
    #presentationContainer .region-table th,
    #presentationContainer .region-table td {{
      padding: 6px 10px;
      font-size: 13px;
    }}
    
    #presentationContainer .region-table th:first-child,
    #presentationContainer .region-table td:first-child {{
      min-width: 100px;
      background-color: white !important; /* Força fundo branco no modo apresentação */
    }}
    
    #presentationContainer .region-table th:first-child {{
      background-color: #f2f2f2 !important; /* Força fundo do cabeçalho na apresentação */
    }}
    
    #presentationContainer .region-table tr:nth-child(even) td:first-child {{
      background-color: white !important;
    }}
    
    #presentationContainer .region-table tr:hover td:first-child {{
      background-color: #f5f5f5 !important;
    }}

    #presentationControls {{
      position: fixed;
      bottom: 20px;
      right: 20px;
      z-index: 10000;
    }}
    #presentationControls button {{
      padding: 10px 20px;
      margin: 5px;
      font-size: 16px;
    }}
    
    /* Apresentação desabilitada em mobile - elementos ocultos */
  </style>
</head>
<body>
<div class="container">
  <div class="header">
    <div class="month-ref">📅 Mês Ref.: {month_ref}</div>
    <div class="header-content">
      <img src="https://raw.githubusercontent.com/aag1974/apn-ivv/main/logo.png" alt="Opinião Logo" class="logo">
      <div class="header-text">
        <h1>📊 Pesquisa IVV Residencial</h1>
        <p>Índice de Velocidade de Vendas - Análise Executiva</p>
        <small>Relatório gerado em: {report_date}</small>
      </div>
    </div>
  </div>

  <div class="nav-bar">
    <div class="nav-buttons">
      <button class="nav-btn active" onclick="showSection('resumo')">📋 {month_ref}</button>
      <button class="nav-btn" onclick="showSection('ivv')">📈 IVV</button>
      <button class="nav-btn" onclick="showSection('ofertas')">🏢 Ofertas</button>
      <button class="nav-btn" onclick="showSection('vendas')">💰 Vendas</button>
      <button class="nav-btn" onclick="showSection('lancamentos')">🚀 Lançamentos</button>
      <button class="nav-btn" onclick="showSection('precos')">💲 Preços</button>
      <button class="nav-btn" onclick="showSection('vgv')">📊 VGV</button>
      <!-- Novas seções: VGL (Valor Geral de Lançamentos) e Distratos -->
      <button class="nav-btn" onclick="showSection('vgl')">📈 VGL</button>
      <button class="nav-btn" onclick="showSection('distratos')">❌ Distratos</button>
    </div>
    <!-- Presentation Button -->
    <div style="text-align:center; margin-top:20px;">
      <button id="presentationButton" class="nav-btn">Iniciar Apresentação</button>
    </div>
  </div>

  <!-- Nova seção de resumo -->
  <div id="resumo" class="section active">
    <div class="summary-card">
      <div class="summary-title">📊 Resumo Executivo - {month_ref}</div>
      <div class="metrics-grid">
        <div class="metric-item">
          <div class="metric-trend">{summary['ivv_trend']}</div>
          <div class="metric-label">IVV</div>
          <div class="metric-value">{summary['ivv']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['ofertas_trend']}</div>
          <div class="metric-label">Unidades Ofertadas</div>
          <div class="metric-value">{summary['ofertas']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['vendas_trend']}</div>
          <div class="metric-label">Unidades Vendidas</div>
          <div class="metric-value">{summary['vendas']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['lancamentos_trend']}</div>
          <div class="metric-label">Unidades Lançadas</div>
          <div class="metric-value">{summary['lancamentos']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['oferta_m2_trend']}</div>
          <div class="metric-label">Oferta em m²</div>
          <div class="metric-value">{summary['oferta_m2']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['venda_m2_trend']}</div>
          <div class="metric-label">Venda em m²</div>
          <div class="metric-value">{summary['venda_m2']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['preco_oferta_trend']}</div>
          <div class="metric-label">Preço de Oferta</div>
          <div class="metric-value">{summary['preco_oferta']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['preco_venda_trend']}</div>
          <div class="metric-label">Preço de Venda</div>
          <div class="metric-value">{summary['preco_venda']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['vgv_trend']}</div>
          <div class="metric-label">VGV</div>
          <div class="metric-value">{summary['vgv']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['vgl_trend']}</div>
          <div class="metric-label">VGL</div>
          <div class="metric-value">{summary['vgl']}</div>
        </div>
        <div class="metric-item">
          <div class="metric-trend">{summary['distratos_trend']}</div>
          <div class="metric-label">Distratos</div>
          <div class="metric-value">{summary['distratos']}</div>
        </div>
      </div>
    </div>
  </div>

  <div id="ivv" class="section">
    <div class="chart-container">
      <div class="chart-title">IVV Mensal - Evolução 2021-2025</div>
      <div class="chart-subtitle">Variação mensal do Índice de Velocidade de Vendas (%)</div>
      <div class="chart-wrapper"><canvas id="ivvMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('IVV MoM','n/d')}</li>
          <li>{highlights.get('IVV YoY','n/d')}</li>
          <li>{highlights.get('IVV Peak','n/d')}</li>          <li>{highlights.get('IVV Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">IVV Trimestral</div>
        <div class="chart-subtitle">Performance por trimestre (%)</div>
        <div class="chart-wrapper small"><canvas id="ivvQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('IVV Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("IVV Quarterly Obs", "")}</em></p>' if 'IVV Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">IVV Anual</div>
        <div class="chart-subtitle">Performance anual consolidada (%)</div>
        <div class="chart-wrapper small"><canvas id="ivvYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('IVV Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("IVV Annual Obs", "")}</em></p>' if 'IVV Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <div id="ofertas" class="section">
    <div class="chart-container">
      <div class="chart-title">Ofertas Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de ofertas (Unidades)</div>
      <div class="chart-wrapper"><canvas id="ofertasMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('Ofertas MoM','n/d')}</li>
          <li>{highlights.get('Ofertas YoY','n/d')}</li>
          <li>{highlights.get('Ofertas Peak','n/d')}</li>          <li>{highlights.get('Ofertas Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Ofertas Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="ofertasQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Ofertas Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("Ofertas Quarterly Obs", "")}</em></p>' if 'Ofertas Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Ofertas Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="ofertasYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Ofertas Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("Ofertas Annual Obs", "")}</em></p>' if 'Ofertas Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
    <!-- Oferta em m² -->
    <div class="chart-container">
      <div class="chart-title">Oferta em m² Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de oferta (m²)</div>
      <div class="chart-wrapper"><canvas id="ofertaM2MonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('OfertaM2 MoM','n/d')}</li>
          <li>{highlights.get('OfertaM2 YoY','n/d')}</li>
          <li>{highlights.get('OfertaM2 Peak','n/d')}</li>          <li>{highlights.get('OfertaM2 Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Oferta em m² Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (m²)</div>
        <div class="chart-wrapper small"><canvas id="ofertaM2QuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('OfertaM2 Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("OfertaM2 Quarterly Obs", "")}</em></p>' if 'OfertaM2 Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Oferta em m² Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (m²)</div>
        <div class="chart-wrapper small"><canvas id="ofertaM2YearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('OfertaM2 Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("OfertaM2 Annual Obs", "")}</em></p>' if 'OfertaM2 Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <div id="vendas" class="section">
    <div class="chart-container">
      <div class="chart-title">Vendas Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de vendas (Unidades)</div>
      <div class="chart-wrapper"><canvas id="vendasMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('Vendas MoM','n/d')}</li>
          <li>{highlights.get('Vendas YoY','n/d')}</li>
          <li>{highlights.get('Vendas Peak','n/d')}</li>          <li>{highlights.get('Vendas Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Vendas Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="vendasQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Vendas Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("Vendas Quarterly Obs", "")}</em></p>' if 'Vendas Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Vendas Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="vendasYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Vendas Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("Vendas Annual Obs", "")}</em></p>' if 'Vendas Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>

    <!-- Vendas em m² -->
    <div class="chart-container">
      <div class="chart-title">Vendas em m² Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de vendas (m²)</div>
      <div class="chart-wrapper"><canvas id="vendaM2MonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('VendaM2 MoM','n/d')}</li>
          <li>{highlights.get('VendaM2 YoY','n/d')}</li>
          <li>{highlights.get('VendaM2 Peak','n/d')}</li>          <li>{highlights.get('VendaM2 Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Vendas em m² Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (m²)</div>
        <div class="chart-wrapper small"><canvas id="vendaM2QuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VendaM2 Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("VendaM2 Quarterly Obs", "")}</em></p>' if 'VendaM2 Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Vendas em m² Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (m²)</div>
        <div class="chart-wrapper small"><canvas id="vendaM2YearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VendaM2 Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("VendaM2 Annual Obs", "")}</em></p>' if 'VendaM2 Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <div id="lancamentos" class="section">
    <!-- Gráficos de empreendimentos (projetos) -->
    <div class="chart-container">
      <div class="chart-title">Empreendimentos Lançados Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Número de empreendimentos lançados por mês</div>
      <div class="chart-wrapper"><canvas id="lancProjMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques (Empreendimentos):</h4>
        <ul>
          <li>{highlights.get('LancProj MoM','n/d')}</li>
          <li>{highlights.get('LancProj YoY','n/d')}</li>
          <li>{highlights.get('LancProj Peak','n/d')}</li>
          <li>{highlights.get('LancProj Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Empreendimentos Lançados Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (Empreendimentos)</div>
        <div class="chart-wrapper small"><canvas id="lancProjQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('LancProj Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("LancProj Quarterly Obs", "")}</em></p>' if 'LancProj Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Empreendimentos Lançados Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (Empreendimentos)</div>
        <div class="chart-wrapper small"><canvas id="lancProjYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('LancProj Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("LancProj Annual Obs", "")}</em></p>' if 'LancProj Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>

    <!-- Gráficos de unidades -->
    <div class="chart-container">
      <div class="chart-title">Lançamentos Mensais (Unidades) - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de lançamentos (Unidades)</div>
      <div class="chart-wrapper"><canvas id="lancMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques (Unidades):</h4>
        <ul>
          <li>{highlights.get('Lanc MoM','n/d')}</li>
          <li>{highlights.get('Lanc YoY','n/d')}</li>
          <li>{highlights.get('Lanc Peak','n/d')}</li>
          <li>{highlights.get('Lanc Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Lançamentos Trimestrais (Unidades)</div>
        <div class="chart-subtitle">Performance por trimestre (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="lancQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Lanc Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("Lanc Quarterly Obs", "")}</em></p>' if 'Lanc Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Lançamentos Anuais (Unidades)</div>
        <div class="chart-subtitle">Performance anual consolidada (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="lancYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Lanc Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("Lanc Annual Obs", "")}</em></p>' if 'Lanc Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <div id="precos" class="section">
    <div class="chart-container">
      <div class="chart-title">Preços de Oferta - Evolução 2021-2025</div>
      <div class="chart-subtitle">Valor médio ponderado mensal (R$)</div>
      <div class="chart-wrapper"><canvas id="precosOfertaMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('PrecosOferta MoM','n/d')}</li>
          <li>{highlights.get('PrecosOferta YoY','n/d')}</li>
          <li>{highlights.get('PrecosOferta Peak','n/d')}</li>          <li>{highlights.get('PrecosOferta Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Preços de Oferta Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (R$)</div>
        <div class="chart-wrapper small"><canvas id="precosOfertaQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('PrecosOferta Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("PrecosOferta Quarterly Obs", "")}</em></p>' if 'PrecosOferta Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Preços de Oferta Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (R$)</div>
        <div class="chart-wrapper small"><canvas id="precosOfertaYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Precos Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("PrecosOferta Annual Obs", "")}</em></p>' if 'PrecosOferta Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>

    <div class="chart-container">
      <div class="chart-title">Preços de Venda - Evolução 2021-2025</div>
      <div class="chart-subtitle">Valor médio ponderado mensal (R$)</div>
      <div class="chart-wrapper"><canvas id="precosVendaMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('PrecosVenda MoM','n/d')}</li>
          <li>{highlights.get('PrecosVenda YoY','n/d')}</li>
          <li>{highlights.get('PrecosVenda Peak','n/d')}</li>          <li>{highlights.get('PrecosVenda Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Preços de Venda Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (R$)</div>
        <div class="chart-wrapper small"><canvas id="precosVendaQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('PrecosVenda Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("PrecosVenda Quarterly Obs", "")}</em></p>' if 'PrecosVenda Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Preços de Venda Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (R$)</div>
        <div class="chart-wrapper small"><canvas id="precosVendaYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('PrecosVenda Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("PrecosVenda Annual Obs", "")}</em></p>' if 'PrecosVenda Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <div id="vgv" class="section">
    <div class="chart-container">
      <div class="chart-title">VGV Mensal - Evolução 2021-2025</div>
      <div class="chart-subtitle">Valor Geral de Vendas mensal (R$ Milhões)</div>
      <div class="chart-wrapper"><canvas id="vgvMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('VGV MoM','n/d')}</li>
          <li>{highlights.get('VGV YoY','n/d')}</li>
          <li>{highlights.get('VGV Peak','n/d')}</li>          <li>{highlights.get('VGV Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">VGV Trimestral</div>
        <div class="chart-subtitle">Performance por trimestre (R$ Milhões)</div>
        <div class="chart-wrapper small"><canvas id="vgvQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VGV Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("VGV Quarterly Obs", "")}</em></p>' if 'VGV Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">VGV Anual</div>
        <div class="chart-subtitle">Performance anual consolidada (R$ Milhões)</div>
        <div class="chart-wrapper small"><canvas id="vgvYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VGV Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("VGV Annual Obs", "")}</em></p>' if 'VGV Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <!-- Seção VGL (Valor Geral de Lançamentos) -->
  <div id="vgl" class="section">
    <div class="chart-container">
      <div class="chart-title">VGL Mensal - Evolução 2021-2025</div>
      <div class="chart-subtitle">Valor geral de lançamentos mensal (R$ Milhões)</div>
      <div class="chart-wrapper"><canvas id="vglMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('VGL MoM','n/d')}</li>
          <li>{highlights.get('VGL YoY','n/d')}</li>
          <li>{highlights.get('VGL Peak','n/d')}</li>          <li>{highlights.get('VGL Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">VGL Trimestral</div>
        <div class="chart-subtitle">Performance por trimestre (R$ Milhões)</div>
        <div class="chart-wrapper small"><canvas id="vglQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VGL Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("VGL Quarterly Obs", "")}</em></p>' if 'VGL Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">VGL Anual</div>
        <div class="chart-subtitle">Performance anual consolidada (R$ Milhões)</div>
        <div class="chart-wrapper small"><canvas id="vglYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('VGL Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("VGL Annual Obs", "")}</em></p>' if 'VGL Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

  <!-- Seção Distratos -->
  <div id="distratos" class="section">
    <div class="chart-container">
      <div class="chart-title">Distratos Mensais - Evolução 2021-2025</div>
      <div class="chart-subtitle">Volume mensal de distratos (Unidades)</div>
      <div class="chart-wrapper"><canvas id="distratosMonthlyChart"></canvas></div>
      <div class="insights">
        <h4>💡 Destaques:</h4>
        <ul>
          <li>{highlights.get('Distratos MoM','n/d')}</li>
          <li>{highlights.get('Distratos YoY','n/d')}</li>
          <li>{highlights.get('Distratos Peak','n/d')}</li>          <li>{highlights.get('Distratos Yearly Avg','n/d')}</li>
        </ul>
      </div>
    </div>
    <div class="grid">
      <div class="chart-container">
        <div class="chart-title">Distratos Trimestrais</div>
        <div class="chart-subtitle">Performance por trimestre (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="distratosQuarterlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Distratos Quarterly', 'N/A')}</h3>
          <p>Destaque trimestral (melhor performance)</p>
          {f'<p class="observation"><em>{highlights.get("Distratos Quarterly Obs", "")}</em></p>' if 'Distratos Quarterly Obs' in highlights else ''}
        </div>
      </div>
      <div class="chart-container">
        <div class="chart-title">Distratos Anuais</div>
        <div class="chart-subtitle">Performance anual consolidada (Unidades)</div>
        <div class="chart-wrapper small"><canvas id="distratosYearlyChart"></canvas></div>
        <div class="highlight-box">
          <h3>{highlights.get('Distratos Annual', 'N/A')}</h3>
          <p>Performance anual</p>
          {f'<p class="observation"><em>{highlights.get("Distratos Annual Obs", "")}</em></p>' if 'Distratos Annual Obs' in highlights else ''}
        </div>
      </div>
    </div>
  </div>

</div>

<script>
/* ===== LOCALE E FORMATADORES (OBRIGATÓRIO) ===== */
Chart.defaults.locale = 'pt-BR';

function fmtIntBR(value) {{
  return value.toLocaleString('pt-BR', {{
    maximumFractionDigits: 0
  }});
}}

function fmtFloatBR(value, decimals = 1) {{
  return value.toLocaleString('pt-BR', {{
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  }});
}}

function fmtPercentBR(value, decimals = 1) {{
  return fmtFloatBR(value, decimals) + '%';
}}

function fmtCurrencyBR(value, decimals = 2) {{
  return 'R$ ' + value.toLocaleString('pt-BR', {{
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  }});
}}

/* ===== NAVEGAÇÃO ===== */
function showSection(sectionId) {{
  const sections = document.querySelectorAll('.section');
  sections.forEach(section => section.classList.remove('active'));

  const target = document.getElementById(sectionId);
  if (target) target.classList.add('active');

  const buttons = document.querySelectorAll('.nav-btn');
  buttons.forEach(btn => btn.classList.remove('active'));

  if (window.event && window.event.target) {{
    window.event.target.classList.add('active');
  }}
}}
// ---------- Datasets (injetados pelo Python)
"""

    # Inject datasets
    datasets_js = []
    if 'IVV Monthly' in data_dict:
        datasets_js.append(f"const ivvMonthlyData = {_to_js_json(data_dict['IVV Monthly'])};")
    if 'IVV Quarterly' in data_dict:
        datasets_js.append(f"const ivvQuarterlyData = {_to_js_json(data_dict['IVV Quarterly'])};")
    if 'IVV Yearly' in data_dict:
        yd, yv = data_dict['IVV Yearly']
        datasets_js.append(f"const ivvYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const ivvYearlyVar = {_to_js_json(yv)};")

    if 'Ofertas Monthly' in data_dict:
        datasets_js.append(f"const ofertasMonthlyData = {_to_js_json(data_dict['Ofertas Monthly'])};")
    if 'Ofertas Quarterly' in data_dict:
        datasets_js.append(f"const ofertasQuarterlyData = {_to_js_json(data_dict['Ofertas Quarterly'])};")
    if 'Ofertas Yearly' in data_dict:
        yd, yv = data_dict['Ofertas Yearly']
        datasets_js.append(f"const ofertasYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const ofertasYearlyVar = {_to_js_json(yv)};")

    if 'Vendas Monthly' in data_dict:
        datasets_js.append(f"const vendasMonthlyData = {_to_js_json(data_dict['Vendas Monthly'])};")
    if 'Vendas Quarterly' in data_dict:
        datasets_js.append(f"const vendasQuarterlyData = {_to_js_json(data_dict['Vendas Quarterly'])};")
    if 'Vendas Yearly' in data_dict:
        yd, yv = data_dict['Vendas Yearly']
        datasets_js.append(f"const vendasYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const vendasYearlyVar = {_to_js_json(yv)};")

    # ---------------- Oferta e Venda em m² ----------------
    # Conjuntos de dados de área (metros quadrados) para oferta
    if 'OfertaM2 Monthly' in data_dict:
        datasets_js.append(f"const ofertaM2MonthlyData = {_to_js_json(data_dict['OfertaM2 Monthly'])};")
    if 'OfertaM2 Quarterly' in data_dict:
        datasets_js.append(f"const ofertaM2QuarterlyData = {_to_js_json(data_dict['OfertaM2 Quarterly'])};")
    if 'OfertaM2 Yearly' in data_dict:
        yd, yv = data_dict['OfertaM2 Yearly']
        datasets_js.append(f"const ofertaM2YearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const ofertaM2YearlyVar = {_to_js_json(yv)};")
    # Conjuntos de dados de área (metros quadrados) para venda
    if 'VendaM2 Monthly' in data_dict:
        datasets_js.append(f"const vendaM2MonthlyData = {_to_js_json(data_dict['VendaM2 Monthly'])};")
    if 'VendaM2 Quarterly' in data_dict:
        datasets_js.append(f"const vendaM2QuarterlyData = {_to_js_json(data_dict['VendaM2 Quarterly'])};")
    if 'VendaM2 Yearly' in data_dict:
        yd, yv = data_dict['VendaM2 Yearly']
        datasets_js.append(f"const vendaM2YearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const vendaM2YearlyVar = {_to_js_json(yv)};")

    if 'Lanc Monthly' in data_dict:
        datasets_js.append(f"const lancMonthlyData = {_to_js_json(data_dict['Lanc Monthly'])};")
    if 'Lanc Quarterly' in data_dict:
        datasets_js.append(f"const lancQuarterlyData = {_to_js_json(data_dict['Lanc Quarterly'])};")
    if 'Lanc Yearly' in data_dict:
        yd, yv = data_dict['Lanc Yearly']
        datasets_js.append(f"const lancYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const lancYearlyVar = {_to_js_json(yv)};")

    # Datasets de empreendimentos (projetos) para lançamentos
    if 'LancProj Monthly' in data_dict:
        datasets_js.append(f"const lancProjMonthlyData = {_to_js_json(data_dict['LancProj Monthly'])};")
    if 'LancProj Quarterly' in data_dict:
        datasets_js.append(f"const lancProjQuarterlyData = {_to_js_json(data_dict['LancProj Quarterly'])};")
    if 'LancProj Yearly' in data_dict:
        yd, yv = data_dict['LancProj Yearly']
        datasets_js.append(f"const lancProjYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const lancProjYearlyVar = {_to_js_json(yv)};")

    if 'Precos Oferta Monthly' in data_dict:
        datasets_js.append(f"const precosOfertaMonthlyData = {_to_js_json(data_dict['Precos Oferta Monthly'])};")
    if 'Precos Oferta Quarterly' in data_dict:
        datasets_js.append(f"const precosOfertaQuarterlyData = {_to_js_json(data_dict['Precos Oferta Quarterly'])};")
    if 'Precos Yearly' in data_dict:
        yd, yv = data_dict['Precos Yearly']
        datasets_js.append(f"const precosOfertaYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const precosOfertaYearlyVar = {_to_js_json(yv)};")

    if 'Precos Venda Monthly' in data_dict:
        datasets_js.append(f"const precosVendaMonthlyData = {_to_js_json(data_dict['Precos Venda Monthly'])};")
    if 'Precos Venda Quarterly' in data_dict:
        datasets_js.append(f"const precosVendaQuarterlyData = {_to_js_json(data_dict['Precos Venda Quarterly'])};")
    if 'Precos Venda Yearly' in data_dict:
        yd, yv = data_dict['Precos Venda Yearly']
        datasets_js.append(f"const precosVendaYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const precosVendaYearlyVar = {_to_js_json(yv)};")

    if 'VGV Monthly' in data_dict:
        datasets_js.append(f"const vgvMonthlyData = {_to_js_json(data_dict['VGV Monthly'])};")
    if 'VGV Quarterly' in data_dict:
        datasets_js.append(f"const vgvQuarterlyData = {_to_js_json(data_dict['VGV Quarterly'])};")
    if 'VGV Yearly' in data_dict:
        yd, yv = data_dict['VGV Yearly']
        datasets_js.append(f"const vgvYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const vgvYearlyVar = {_to_js_json(yv)};")

    # VGL (Valor Geral de Lançamentos)
    if 'VGL Monthly' in data_dict:
        datasets_js.append(f"const vglMonthlyData = {_to_js_json(data_dict['VGL Monthly'])};")
    if 'VGL Quarterly' in data_dict:
        datasets_js.append(f"const vglQuarterlyData = {_to_js_json(data_dict['VGL Quarterly'])};")
    if 'VGL Yearly' in data_dict:
        yd, yv = data_dict['VGL Yearly']
        datasets_js.append(f"const vglYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const vglYearlyVar = {_to_js_json(yv)};")

    # Distratos
    if 'Distratos Monthly' in data_dict:
        datasets_js.append(f"const distratosMonthlyData = {_to_js_json(data_dict['Distratos Monthly'])};")
    if 'Distratos Quarterly' in data_dict:
        datasets_js.append(f"const distratosQuarterlyData = {_to_js_json(data_dict['Distratos Quarterly'])};")
    if 'Distratos Yearly' in data_dict:
        yd, yv = data_dict['Distratos Yearly']
        datasets_js.append(f"const distratosYearlyData = {_to_js_json(yd)};")
        datasets_js.append(f"const distratosYearlyVar = {_to_js_json(yv)};")

    html += "\n".join("  " + s for s in datasets_js) + "\n\n"

    html += r"""
// ---------- Config base (legenda com bolinhas)
const legendCircle = {
  position: 'top',
  labels: {
    usePointStyle: true,
    pointStyle: 'circle'
  }
};

function drawYearlyChart(ctx, data, variations, labelFormatter) {
  // Calcular valores min/max
  const values = data.datasets[0].data.filter(v => v !== null && v !== undefined);
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const range = maxValue - minValue;
  
  // Detectar tipo de dado baseado no contexto do gráfico - com verificação de segurança
  const chartId = ctx && ctx.canvas && ctx.canvas.id ? ctx.canvas.id : '';
  let yMin, yMax;
  
  if (chartId.includes('ivv')) {
    // IVV: Percentuais - SEMPRE múltiplos exatos de 1%
    yMin = Math.floor(minValue - 0.5); // Margem pequena para baixo
    yMax = Math.ceil(maxValue + 0.5);  // Margem pequena para cima
    
    // Forçar múltiplos de 1% inteiro
    yMin = Math.max(0, Math.floor(yMin));
    yMax = Math.ceil(yMax);
    
    // Garantir range mínimo para visualização
    if (yMax - yMin < 3) {
      const center = (yMin + yMax) / 2;
      yMin = Math.floor(center - 1.5);
      yMax = Math.ceil(center + 1.5);
      yMin = Math.max(0, yMin); // Não pode ser negativo
    }
  } 
  else if (chartId.includes('precos') || chartId.includes('vgv')) {
    // Preços/VGV: Múltiplos de 1000, sem decimais
    const buffer = range * 0.08;
    yMin = Math.max(0, Math.floor((minValue - buffer) / 1000) * 1000);
    yMax = Math.ceil((maxValue + buffer) / 1000) * 1000;
  }
  else {
    // Ofertas/Vendas/Lançamentos: Algoritmo conservador
    const minBuffer = range * 0.10;
    const maxBuffer = range * 0.05;
    
    const rawYMin = Math.max(0, minValue - minBuffer);
    const rawYMax = maxValue + maxBuffer;
    
    // Arredondar para centenas ou milhares conforme magnitude
    const magnitude = Math.pow(10, Math.floor(Math.log10(maxValue)) - 1);
    yMin = Math.floor(rawYMin / magnitude) * magnitude;
    yMax = Math.ceil(rawYMax / magnitude) * magnitude;
  }

  return new Chart(ctx, {
    type: 'bar',
    data: data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(context) {
              const rawValue = context.parsed.y;
              const formattedValue = labelFormatter
                ? labelFormatter(rawValue)
                : rawValue.toLocaleString('pt-BR');

              const variation =
                variations && variations[context.dataIndex]
                  ? variations[context.dataIndex]
                  : '';

              return variation
                ? `${formattedValue} (${variation})`
                : formattedValue;
            }
          }
        }
      },
      scales: {
        y: {
          min: yMin,
          max: yMax,
          ticks: {
            // Forçar ticks em números inteiros para IVV
            stepSize: chartId.includes('ivv') ? 1 : undefined,
            callback: function(value) {
              // Formatação específica por tipo
              if (chartId.includes('precos') || chartId.includes('vgv')) {
                // Preços/VGV: sem casas decimais
                return labelFormatter && labelFormatter.toString().includes('R$')
                  ? 'R$ ' + value.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
                  : value.toLocaleString('pt-BR', { maximumFractionDigits: 0 });
              } else if (chartId.includes('ivv')) {
                // IVV: sempre mostrar como percentual inteiro
                return value.toFixed(0) + '%';
              } else {
                // Outros: usar labelFormatter normal ou formatação padrão
                return labelFormatter
                  ? labelFormatter(value)
                  : value.toLocaleString('pt-BR');
              }
            }
          }
        }
      },
      animation: { duration: 2000, easing: 'easeInOutQuart' }
    }
  });
}
window.addEventListener('load', function() {

  // IVV
  if (typeof ivvMonthlyData !== 'undefined') {
    // Calcular limites para múltiplos de 1% 
    const values = ivvMonthlyData.datasets.flatMap(d => d.data.filter(v => v !== null));
    const minVal = Math.floor(Math.min(...values) - 0.5);
    const maxVal = Math.ceil(Math.max(...values) + 0.5);
    
    new Chart(document.getElementById('ivvMonthlyChart'), {
      type: 'line',
      data: ivvMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${fmtPercentBR(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          y: {
            min: Math.max(0, minVal),
            max: maxVal,
            ticks: { 
              stepSize: 1,
              callback: (v) => v.toFixed(0) + '%'
            }
          }
        },
        animation: { duration: 2000, easing: 'easeInOutQuart' }
      }
    });
  }

  if (typeof ivvQuarterlyData !== 'undefined') {
    // Calcular limites para múltiplos de 1%
    const values = ivvQuarterlyData.datasets.flatMap(d => d.data.filter(v => v !== null));
    const minVal = Math.floor(Math.min(...values) - 0.5);
    const maxVal = Math.ceil(Math.max(...values) + 0.5);
    
    new Chart(document.getElementById('ivvQuarterlyChart'), {
      type: 'bar',
      data: ivvQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${fmtPercentBR(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          y: {
            min: Math.max(0, minVal),
            max: maxVal,
            ticks: { 
              stepSize: 1,
              callback: (v) => v.toFixed(0) + '%'
            }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof ivvYearlyData !== 'undefined') {
    // Calcular limites específicos para IVV anual com múltiplos de 1%
    const values = ivvYearlyData.datasets[0].data.filter(v => v !== null && v !== undefined);
    const minVal = Math.floor(Math.min(...values) - 0.5);
    const maxVal = Math.ceil(Math.max(...values) + 0.5);
    
    new Chart(document.getElementById('ivvYearlyChart'), {
      type: 'bar',
      data: ivvYearlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: function(context) {
                const rawValue = context.parsed.y;
                const formattedValue = fmtPercentBR(rawValue);
                const variation = ivvYearlyVar && ivvYearlyVar[context.dataIndex] 
                  ? ivvYearlyVar[context.dataIndex] : '';
                return variation ? `${formattedValue} (${variation})` : formattedValue;
              }
            }
          }
        },
        scales: {
          y: {
            min: Math.max(0, minVal),
            max: maxVal,
            ticks: {
              stepSize: 1,
              callback: function(value) {
                return value.toFixed(0) + '%';
              }
            }
          }
        },
        animation: { duration: 2000, easing: 'easeInOutQuart' }
      }
    });
  }

  // Ofertas
  if (typeof ofertasMonthlyData !== 'undefined') {
    new Chart(document.getElementById('ofertasMonthlyChart'), {
      type: 'line',
      data: ofertasMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof ofertasQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('ofertasQuarterlyChart'), {
      type: 'bar',
      data: ofertasQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof ofertasYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('ofertasYearlyChart'),
      ofertasYearlyData,
      ofertasYearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Vendas
  if (typeof vendasMonthlyData !== 'undefined') {
    new Chart(document.getElementById('vendasMonthlyChart'), {
      type: 'line',
      data: vendasMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof vendasQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('vendasQuarterlyChart'), {
      type: 'bar',
      data: vendasQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof vendasYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('vendasYearlyChart'),
      vendasYearlyData,
      vendasYearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Oferta em m²
  if (typeof ofertaM2MonthlyData !== 'undefined') {
    new Chart(document.getElementById('ofertaM2MonthlyChart'), {
      type: 'line',
      data: ofertaM2MonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }
  if (typeof ofertaM2QuarterlyData !== 'undefined') {
    new Chart(document.getElementById('ofertaM2QuarterlyChart'), {
      type: 'bar',
      data: ofertaM2QuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }
  if (typeof ofertaM2YearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('ofertaM2YearlyChart'),
      ofertaM2YearlyData,
      ofertaM2YearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Venda em m²
  if (typeof vendaM2MonthlyData !== 'undefined') {
    new Chart(document.getElementById('vendaM2MonthlyChart'), {
      type: 'line',
      data: vendaM2MonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }
  if (typeof vendaM2QuarterlyData !== 'undefined') {
    new Chart(document.getElementById('vendaM2QuarterlyChart'), {
      type: 'bar',
      data: vendaM2QuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }
  if (typeof vendaM2YearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('vendaM2YearlyChart'),
      vendaM2YearlyData,
      vendaM2YearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Lançamentos
  if (typeof lancMonthlyData !== 'undefined') {
    new Chart(document.getElementById('lancMonthlyChart'), {
      type: 'line',
      data: lancMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof lancQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('lancQuarterlyChart'), {
      type: 'bar',
      data: lancQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof lancYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('lancYearlyChart'),
      lancYearlyData,
      lancYearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Lançamentos - Empreendimentos
  if (typeof lancProjMonthlyData !== 'undefined') {
    new Chart(document.getElementById('lancProjMonthlyChart'), {
      type: 'line',
      data: lancProjMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }
  if (typeof lancProjQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('lancProjQuarterlyChart'), {
      type: 'bar',
      data: lancProjQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR')}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => v.toLocaleString('pt-BR') }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }
  if (typeof lancProjYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('lancProjYearlyChart'),
      lancProjYearlyData,
      lancProjYearlyVar || [],
      (v) => v.toLocaleString('pt-BR')
    );
  }

  // Preços - Oferta
  if (typeof precosOfertaMonthlyData !== 'undefined') {
    new Chart(document.getElementById('precosOfertaMonthlyChart'), {
      type: 'line',
      data: precosOfertaMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof precosOfertaQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('precosOfertaQuarterlyChart'), {
      type: 'bar',
      data: precosOfertaQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof precosOfertaYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('precosOfertaYearlyChart'),
      precosOfertaYearlyData,
      precosOfertaYearlyVar || [],
      (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
    );
  }

  // Preços - Venda
  if (typeof precosVendaMonthlyData !== 'undefined') {
    new Chart(document.getElementById('precosVendaMonthlyChart'), {
      type: 'line',
      data: precosVendaMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof precosVendaQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('precosVendaQuarterlyChart'), {
      type: 'bar',
      data: precosVendaQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: { callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  // >>> BLOCO ANUAL (ESTAVA FALTANDO) <<<
  if (typeof precosVendaYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('precosVendaYearlyChart'),
      precosVendaYearlyData,
      precosVendaYearlyVar || [],
      (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
    );
  }

  // VGV
  if (typeof vgvMonthlyData !== 'undefined') {
    new Chart(document.getElementById('vgvMonthlyChart'), {
      type: 'line',
      data: vgvMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}M`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
            }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }

  if (typeof vgvQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('vgvQuarterlyChart'), {
      type: 'bar',
      data: vgvQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}M`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
            }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }

  if (typeof vgvYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('vgvYearlyChart'),
      vgvYearlyData,
      vgvYearlyVar || [],
      (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
    );
  }

  // VGL (Valor Geral de Lançamentos)
  if (typeof vglMonthlyData !== 'undefined') {
    new Chart(document.getElementById('vglMonthlyChart'), {
      type: 'line',
      data: vglMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}M`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
            }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }
  if (typeof vglQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('vglQuarterlyChart'), {
      type: 'bar',
      data: vglQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: R$ ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}M`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
            }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }
  if (typeof vglYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('vglYearlyChart'),
      vglYearlyData,
      vglYearlyVar || [],
      (v) => 'R$ ' + v.toLocaleString('pt-BR', { maximumFractionDigits: 0 }) + 'M'
    );
  }

  // Distratos
  if (typeof distratosMonthlyData !== 'undefined') {
    new Chart(document.getElementById('distratosMonthlyChart'), {
      type: 'line',
      data: distratosMonthlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            mode: 'index',
            intersect: false,
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
            }
          }
        },
        animation: { duration: 2000 }
      }
    });
  }
  if (typeof distratosQuarterlyData !== 'undefined') {
    new Chart(document.getElementById('distratosQuarterlyChart'), {
      type: 'bar',
      data: distratosQuarterlyData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: legendCircle,
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: false,
            ticks: {
              callback: (v) => v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
            }
          }
        },
        animation: { duration: 1500 }
      }
    });
  }
  if (typeof distratosYearlyData !== 'undefined') {
    drawYearlyChart(
      document.getElementById('distratosYearlyChart'),
      distratosYearlyData,
      distratosYearlyVar || [],
      (v) => v.toLocaleString('pt-BR', { maximumFractionDigits: 0 })
    );
  }

}); // Fechamento da função window.addEventListener

</script>

  <!-- Presentation Container for slides (Desktop only) -->
  <div id="presentationContainer"></div>

  <script>
    (function() {
      const btn = document.getElementById('presentationButton');
      const container = document.getElementById('presentationContainer');
      let slides = [];
      let current = 0;
      // Armazenar as posições originais para restaurar
      let originalInfo = [];
      function createSlides() {
        slides = [];
        originalInfo = [];
        container.innerHTML = '';
        
        // Detectar seção ativa atual para começar da view correta
        const currentActiveSection = document.querySelector('.section.active');
        
        // Primeiro slide: cabeçalho + cards de resumo. Clonamos ambos para não remover do DOM original.
        const header = document.querySelector('.header');
        const summary = document.querySelector('.summary-card');
        if (header || summary) {
          const slide0 = document.createElement('div');
          slide0.classList.add('slide');
          let combinedHTML = '';
          if (header) combinedHTML += header.outerHTML;
          if (summary) combinedHTML += summary.outerHTML;
          slide0.innerHTML = combinedHTML;
          slides.push({wrapper: slide0, element: null, sectionId: 'resumo'});
        }
        // Slides para cada gráfico e tabelas: mover o elemento real
        document.querySelectorAll('.chart-container').forEach(el => {
          const slide = document.createElement('div');
          slide.classList.add('slide');
          
          // Detectar de qual seção este container faz parte
          const parentSection = el.closest('.section');
          const sectionId = parentSection ? parentSection.id : 'unknown';
          
          slides.push({wrapper: slide, element: el, sectionId: sectionId});
        });
        
        // Encontrar índice para começar baseado na seção ativa
        if (currentActiveSection) {
          const currentSectionId = currentActiveSection.id;
          for (let i = 0; i < slides.length; i++) {
            if (slides[i].sectionId === currentSectionId) {
              current = i; // Definir onde começar
              break;
            }
          }
        } else {
          current = 0; // Começar do início se não detectar seção
        }
      }
      function attachSlides() {
        // Adicionar slides ao container, movendo os elementos reais
        slides.forEach((s) => {
          const {wrapper, element} = s;
          // se houver elemento associado, mover da página para o wrapper
          if (element) {
            originalInfo.push({
              element: element,
              parent: element.parentNode,
              nextSibling: element.nextSibling
            });
            wrapper.appendChild(element);
          }
          container.appendChild(wrapper);
        });
        // Após anexar, ativar tooltips nativas no último ponto e mantê-las visíveis
        slides.forEach(({element}) => {
          if (element && element.classList.contains('chart-container')) {
            const canvas = element.querySelector('canvas');
            if (canvas) {
              try {
                let chart = null;
                if (Chart.getChart) {
                  chart = Chart.getChart(canvas) || Chart.getChart(canvas.id);
                }
                
                if (chart && chart.data && chart.data.labels && chart.data.labels.length > 0) {
                  const lastIndex = chart.data.labels.length - 1;
                  
                  // Função para ativar tooltip nativa no último ponto permanentemente
                  const activateNativeTooltip = () => {
                    try {
                      // Configurar elementos ativos baseado no tipo de gráfico
                      let activeElements = [];
                      
                      if (chart.config && chart.config.type === 'line') {
                        // Para gráficos de linha: ativar todos os datasets no último índice
                        activeElements = chart.data.datasets.map((ds, i) => ({ 
                          datasetIndex: i, 
                          index: lastIndex 
                        }));
                        
                        // Garantir modo de interação adequado
                        if (!chart.options.interaction) {
                          chart.options.interaction = {};
                        }
                        chart.options.interaction.mode = 'index';
                        chart.options.interaction.intersect = false;
                      } else {
                        // Para gráficos de barras: ativar apenas o último dataset no último índice
                        const lastDatasetIndex = chart.data.datasets.length - 1;
                        activeElements = [{ 
                          datasetIndex: lastDatasetIndex, 
                          index: lastIndex 
                        }];
                      }
                      
                      // Ativar elementos e forçar tooltip a aparecer
                      if (chart.setActiveElements && activeElements.length > 0) {
                        chart.setActiveElements(activeElements);
                        chart.update('none'); // Atualiza sem animação
                        
                        // Forçar tooltip a permanecer visível
                        const originalTooltipPosition = chart.tooltip.getActiveElements();
                        
                        // Override do método de limpeza de tooltip para mantê-la ativa
                        const originalSetActiveElements = chart.setActiveElements.bind(chart);
                        chart.setActiveElements = function(elements) {
                          // Se tentarem limpar, manter o último ponto ativo
                          if (!elements || elements.length === 0) {
                            elements = activeElements;
                          }
                          return originalSetActiveElements(elements);
                        };
                        
                        // Simular evento de mouse no último ponto para garantir que tooltip apareça
                        setTimeout(() => {
                          try {
                            const meta = chart.getDatasetMeta(0);
                            if (meta && meta.data && meta.data[lastIndex]) {
                              const point = meta.data[lastIndex];
                              const rect = canvas.getBoundingClientRect();
                              const position = point.getCenterPoint();
                              
                              // Criar evento sintético
                              const syntheticEvent = {
                                type: 'mousemove',
                                x: position.x,
                                y: position.y,
                                native: {
                                  clientX: rect.left + position.x,
                                  clientY: rect.top + position.y
                                }
                              };
                              
                              // Ativar tooltip
                              chart._handleEvent(syntheticEvent, true);
                            }
                          } catch (e) {
                            console.log('Fallback: usando setActiveElements');
                            chart.setActiveElements(activeElements);
                            chart.update('none');
                          }
                        }, 100);
                      }
                      
                    } catch (innerErr) {
                      console.error('Erro ao ativar tooltip nativa:', innerErr);
                    }
                  };
                  
                  // Ativar tooltip após animação do gráfico
                  setTimeout(activateNativeTooltip, 2800);
                  
                  // Reativar periodicamente para garantir que permanece visível
                  const keepActive = setInterval(() => {
                    if (element.closest('.slide.active')) {
                      activateNativeTooltip();
                    }
                  }, 5000);
                  
                  // Limpar interval quando slide não estiver ativo
                  element.setAttribute('data-interval-id', keepActive);
                }
              } catch (err) {
                console.error('Erro ao configurar tooltip nativa:', err);
              }
            }
          }
        });
      }
      function detachSlides() {
        // Limpar intervals de tooltips ativas
        document.querySelectorAll('.chart-container[data-interval-id]').forEach(container => {
          const intervalId = container.getAttribute('data-interval-id');
          if (intervalId) {
            clearInterval(parseInt(intervalId));
            container.removeAttribute('data-interval-id');
          }
        });
        
        // Restaurar comportamento original dos gráficos
        document.querySelectorAll('.chart-container canvas').forEach(canvas => {
          try {
            const chart = Chart.getChart(canvas);
            if (chart && chart.setActiveElements) {
              // Limpar elementos ativos
              chart.setActiveElements([]);
              chart.update('none');
            }
          } catch (e) {
            // Ignorar erros na limpeza
          }
        });
        
        // Mover os elementos de volta para seus locais originais
        originalInfo.forEach(info => {
          const {element, parent, nextSibling} = info;
          if (parent) {
            if (nextSibling && parent.contains(nextSibling)) {
              parent.insertBefore(element, nextSibling);
            } else {
              parent.appendChild(element);
            }
          }
        });
        originalInfo = [];
      }
      function showSlide(index) {
        slides.forEach((s, i) => {
          if (i === index) {
            s.wrapper.classList.add('active');
          } else {
            s.wrapper.classList.remove('active');
          }
        });
      }
      function startPresentation() {
        // BLOQUEAR APRESENTAÇÃO EM MOBILE
        if (window.innerWidth <= 768) {
          console.log('Apresentação desabilitada em dispositivos móveis');
          return;
        }
        
        if (container.style.display === 'block') {
          // finalizar
          container.style.display = 'none';
          document.body.style.overflow = '';
          // sair de tela cheia, se ativo
          if (document.fullscreenElement) {
            const exitFull = document.exitFullscreen || document.webkitExitFullscreen || document.mozCancelFullScreen || document.msExitFullscreen;
            if (exitFull) exitFull.call(document);
          }
          if (btn) btn.textContent = 'Iniciar Apresentação';
          // remover slides e restaurar elementos
          detachSlides();
        } else {
          createSlides();
          attachSlides();
          container.style.display = 'block';
          document.body.style.overflow = 'hidden';
          // solicitar tela cheia, se suportado
          if (document.fullscreenEnabled) {
            const reqFull = container.requestFullscreen || container.webkitRequestFullscreen || container.mozRequestFullScreen || container.msRequestFullscreen;
            if (reqFull) reqFull.call(container);
          }
          // current já foi definido em createSlides()
          showSlide(current);
          if (btn) btn.textContent = 'Finalizar Apresentação';
        }
      }
      if (btn) {
        btn.addEventListener('click', (e) => {
          e.preventDefault();
          startPresentation();
        });
      }
      
      // Apresentação desabilitada em mobile - eventos touch removidos
      
      // Monitorar saída do fullscreen para finalizar apresentação automaticamente
      document.addEventListener('fullscreenchange', () => {
        if (container.style.display === 'block' && !document.fullscreenElement) {
          // Usuário saiu do fullscreen (provavelmente via ESC)
          // Finalizar apresentação automaticamente após delay
          setTimeout(() => startPresentation(), 100);
        }
      });
      
      // Compatibilidade cross-browser para fullscreenchange
      document.addEventListener('webkitfullscreenchange', () => {
        if (container.style.display === 'block' && !document.webkitFullscreenElement) {
          setTimeout(() => startPresentation(), 100);
        }
      });
      
      document.addEventListener('mozfullscreenchange', () => {
        if (container.style.display === 'block' && !document.mozFullScreenElement) {
          setTimeout(() => startPresentation(), 100);
        }
      });

      // Navegação via teclado
      document.addEventListener('keydown', (e) => {
        if (container.style.display === 'block') {
          if (e.key === 'ArrowRight' || e.key === 'PageDown') {
            if (current < slides.length - 1) {
              current += 1;
              showSlide(current);
            }
            e.preventDefault();
          } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
            if (current > 0) {
              current -= 1;
              showSlide(current);
            }
            e.preventDefault();
          } else if (e.key === 'Escape') {
            // Se não estiver em fullscreen, finalizar diretamente
            if (!document.fullscreenElement && !document.webkitFullscreenElement && !document.mozFullScreenElement) {
              startPresentation();
            }
            // Se estiver em fullscreen, o evento fullscreenchange cuidará da finalização
          }
        }
      });
    })();
  </script>

</body>
</html>
"""
    return html


# -------------------------
# Main
# -------------------------
def main():
    if len(sys.argv) != 3:
        print("Usage: python3 excel_to_html_report_final.py <input_excel.xlsx> <output_html.html>")
        sys.exit(1)

    input_excel = sys.argv[1]
    output_html = sys.argv[2]

    filename = os.path.basename(input_excel)
    month_ref = extract_month_ref(filename)
    report_date = datetime.now().strftime("%d/%m/%Y %H:%M")

    print(f"📖 Reading Excel file: {input_excel}")
    excel_file = pd.ExcelFile(input_excel)
    sheets = {}
    for sheet_name in excel_file.sheet_names:
        try:
            df = pd.read_excel(input_excel, sheet_name=sheet_name)
            sheets[sheet_name] = df
            print(f"  ✓ {sheet_name}: {df.shape[0]} rows, {df.shape[1]} cols")
        except Exception as e:
            print(f"  ✗ Error reading {sheet_name}: {e}")

    data_dict = {}
    highlights = {}

    # ---------------- IVV ----------------
    if 'IVV Mensal' in sheets:
        ivv_monthly = build_monthly_dataset(sheets['IVV Mensal'], is_percent=True)
        data_dict['IVV Monthly'] = ivv_monthly

        # Novos insights usando dados MoM/YoY do Excel
        ivv_insights = format_new_insights(sheets['IVV Mensal'], data_type='percent', month_ref=month_ref)
        
        highlights['IVV MoM'] = ivv_insights['mom']
        highlights['IVV YoY'] = ivv_insights['yoy']
        highlights['IVV Peak'] = ivv_insights['peak']
        highlights['IVV Yearly Avg'] = ivv_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if ivv_monthly['datasets']:
            cur = ivv_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['IVV Trend'] = trend

    if 'IVV Trimestral' in sheets:
        data_dict['IVV Quarterly'] = build_quarterly_dataset(sheets['IVV Trimestral'], is_percent=True)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['IVV Trimestral'], data_type='percent')
        if best_value is not None and best_quarter:
            highlights['IVV Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_percent(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['IVV Trimestral'])
        if observation:
            highlights['IVV Quarterly Obs'] = observation

    if 'IVV Anual' in sheets:
        data, var = build_yearly_dataset(sheets['IVV Anual'], is_percent=True)
        data_dict['IVV Yearly'] = (data, var)

        df_a = clean_dataframe(sheets['IVV Anual'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_percentage(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['IVV Annual'] = f"{year}: {br_percent(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['IVV Anual'])
        if observation:
            highlights['IVV Annual Obs'] = observation

    # ---------------- Ofertas ----------------
    if 'Ofertas Mensais (Unidades)' in sheets:
        ofertas_monthly = build_monthly_dataset(sheets['Ofertas Mensais (Unidades)'], is_percent=False)
        data_dict['Ofertas Monthly'] = ofertas_monthly

        # Novos insights usando dados MoM/YoY do Excel
        ofertas_insights = format_new_insights(sheets['Ofertas Mensais (Unidades)'], data_type='number', month_ref=month_ref)
        
        highlights['Ofertas MoM'] = ofertas_insights['mom']
        highlights['Ofertas YoY'] = ofertas_insights['yoy']
        highlights['Ofertas Peak'] = ofertas_insights['peak']
        highlights['Ofertas Yearly Avg'] = ofertas_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if ofertas_monthly['datasets']:
            cur = ofertas_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['Ofertas Trend'] = trend

    if 'Ofertas Trimestrais (Unidades)' in sheets:
        data_dict['Ofertas Quarterly'] = build_quarterly_dataset(sheets['Ofertas Trimestrais (Unidades)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['Ofertas Trimestrais (Unidades)'], data_type='number')
        if best_value is not None and best_quarter:
            highlights['Ofertas Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Ofertas Trimestrais (Unidades)'])
        if observation:
            highlights['Ofertas Quarterly Obs'] = observation

    if 'Ofertas Anuais (Unidades)' in sheets:
        data, var = build_yearly_dataset(sheets['Ofertas Anuais (Unidades)'], is_percent=False)
        data_dict['Ofertas Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['Ofertas Anuais (Unidades)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['Ofertas Annual'] = f"{year}: {br_int(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Ofertas Anuais (Unidades)'])
        if observation:
            highlights['Ofertas Annual Obs'] = observation

    # ---------------- Vendas ----------------
    if 'Vendas Mensais (Unidades)' in sheets:
        vendas_monthly = build_monthly_dataset(sheets['Vendas Mensais (Unidades)'], is_percent=False)
        data_dict['Vendas Monthly'] = vendas_monthly

        # Novos insights usando dados MoM/YoY do Excel
        vendas_insights = format_new_insights(sheets['Vendas Mensais (Unidades)'], data_type='number', month_ref=month_ref)
        
        highlights['Vendas MoM'] = vendas_insights['mom']
        highlights['Vendas YoY'] = vendas_insights['yoy']
        highlights['Vendas Peak'] = vendas_insights['peak']
        highlights['Vendas Yearly Avg'] = vendas_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if vendas_monthly['datasets']:
            cur = vendas_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['Vendas Trend'] = trend

    if 'Vendas Trimestrais (Unidades)' in sheets:
        data_dict['Vendas Quarterly'] = build_quarterly_dataset(sheets['Vendas Trimestrais (Unidades)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['Vendas Trimestrais (Unidades)'], data_type='number')
        if best_value is not None and best_quarter:
            highlights['Vendas Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Vendas Trimestrais (Unidades)'])
        if observation:
            highlights['Vendas Quarterly Obs'] = observation

    if 'Vendas Anuais (Unidades)' in sheets:
        data, var = build_yearly_dataset(sheets['Vendas Anuais (Unidades)'], is_percent=False)
        data_dict['Vendas Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['Vendas Anuais (Unidades)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['Vendas Annual'] = f"{year}: {br_int(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Vendas Anuais (Unidades)'])
        if observation:
            highlights['Vendas Annual Obs'] = observation

    # ---------------- Oferta em m² ----------------
    # Processa as planilhas de área para oferta (metros quadrados)
    if 'Oferta Mensal (m²)' in sheets:
        oferta_m2_monthly = build_monthly_dataset(sheets['Oferta Mensal (m²)'], is_percent=False)
        data_dict['OfertaM2 Monthly'] = oferta_m2_monthly

        # Novos insights usando dados MoM/YoY do Excel
        oferta_m2_insights = format_new_insights(sheets['Oferta Mensal (m²)'], data_type='number', month_ref=month_ref)
        
        highlights['OfertaM2 MoM'] = oferta_m2_insights['mom']
        highlights['OfertaM2 YoY'] = oferta_m2_insights['yoy']
        highlights['OfertaM2 Peak'] = oferta_m2_insights['peak']
        highlights['OfertaM2 Yearly Avg'] = oferta_m2_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if oferta_m2_monthly['datasets']:
            cur = oferta_m2_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['OfertaM2 Trend'] = trend

    if 'Oferta Trimestral (m²)' in sheets:
        data_dict['OfertaM2 Quarterly'] = build_quarterly_dataset(sheets['Oferta Trimestral (m²)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['Oferta Trimestral (m²)'], data_type='number')
        if best_value is not None and best_quarter:
            highlights['OfertaM2 Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)} m²"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Oferta Trimestral (m²)'])
        if observation:
            highlights['OfertaM2 Quarterly Obs'] = observation

    if 'Oferta Anual (m²)' in sheets:
        data, var = build_yearly_dataset(sheets['Oferta Anual (m²)'], is_percent=False)
        data_dict['OfertaM2 Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['Oferta Anual (m²)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['OfertaM2 Annual'] = f"{year}: {br_int(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Oferta Anual (m²)'])
        if observation:
            highlights['OfertaM2 Annual Obs'] = observation

    # ---------------- Venda em m² ----------------
    # Processa as planilhas de área para venda (metros quadrados)
    if 'Venda Mensal (m²)' in sheets:
        venda_m2_monthly = build_monthly_dataset(sheets['Venda Mensal (m²)'], is_percent=False)
        data_dict['VendaM2 Monthly'] = venda_m2_monthly

        # Novos insights usando dados MoM/YoY do Excel
        venda_m2_insights = format_new_insights(sheets['Venda Mensal (m²)'], data_type='number', month_ref=month_ref)
        
        highlights['VendaM2 MoM'] = venda_m2_insights['mom']
        highlights['VendaM2 YoY'] = venda_m2_insights['yoy']
        highlights['VendaM2 Peak'] = venda_m2_insights['peak']
        highlights['VendaM2 Yearly Avg'] = venda_m2_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if venda_m2_monthly['datasets']:
            cur = venda_m2_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['VendaM2 Trend'] = trend

    if 'Venda Trimestral (m²)' in sheets:
        data_dict['VendaM2 Quarterly'] = build_quarterly_dataset(sheets['Venda Trimestral (m²)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['Venda Trimestral (m²)'], data_type='number')
        if best_value is not None and best_quarter:
            highlights['VendaM2 Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)} m²"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Venda Trimestral (m²)'])
        if observation:
            highlights['VendaM2 Quarterly Obs'] = observation

    if 'Venda Anual (m²)' in sheets:
        data, var = build_yearly_dataset(sheets['Venda Anual (m²)'], is_percent=False)
        data_dict['VendaM2 Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['Venda Anual (m²)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['VendaM2 Annual'] = f"{year}: {br_int(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Venda Anual (m²)'])
        if observation:
            highlights['VendaM2 Annual Obs'] = observation

    # ---------------- Lançamentos ----------------
    lanc_month_sheet = next((n for n in sheets if n.startswith('Lançamentos Mensais')), None)
    lanc_quart_sheet = next((n for n in sheets if n.startswith('Lançamentos Trimestrais')), None)
    lanc_year_sheet = next((n for n in sheets if n.startswith('Lançamentos Anuais')), None)

    if lanc_month_sheet:
        lanc_monthly = build_monthly_dataset(sheets[lanc_month_sheet], is_percent=False)
        data_dict['Lanc Monthly'] = lanc_monthly

        # Novos insights usando dados MoM/YoY do Excel
        lanc_insights = format_new_insights(sheets[lanc_month_sheet], data_type='number', month_ref=month_ref)
        
        highlights['Lanc MoM'] = lanc_insights['mom']
        highlights['Lanc YoY'] = lanc_insights['yoy']
        highlights['Lanc Peak'] = lanc_insights['peak']
        highlights['Lanc Yearly Avg'] = lanc_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if lanc_monthly['datasets']:
            cur = lanc_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['Lanc Trend'] = trend

        # Calcular métricas para empreendimentos (número de projetos) usando números entre colchetes
        df_lanc = clean_dataframe(sheets[lanc_month_sheet])
        if not df_lanc.empty and df_lanc.shape[1] > 1:
            # Cada coluna após a primeira representa um ano; extrair última e penúltima para YoY
            years_cols = list(df_lanc.columns[1:])
            if years_cols:
                last_col = years_cols[-1]
                prev_col = years_cols[-2] if len(years_cols) >= 2 else None
                # Valores do ano corrente (empreendimentos)
                cur_vals = [parse_bracket_number(v) for v in df_lanc[last_col]]
                # Valores do ano anterior para YoY
                # Novos insights para empreendimentos usando dados MoM/YoY do Excel
                # Como empreendimentos são extraídos da mesma planilha de lançamentos,
                # usamos os mesmos dados MoM/YoY mas formatados para projetos
                lancproj_insights = format_new_insights(sheets[lanc_month_sheet], data_type='number', month_ref=month_ref)
                
                # Adaptar os insights para empreendimentos (projetos)
                highlights['LancProj MoM'] = lancproj_insights['mom'].replace('Variação MoM:', 'Variação MoM (projetos):')
                highlights['LancProj YoY'] = lancproj_insights['yoy'].replace('Variação YoY:', 'Variação YoY (projetos):')
                
                # Para pico e média, usar os dados calculados dos valores entre colchetes
                proj_peak = calc_peak(cur_vals)
                highlights['LancProj Peak'] = f"Pico: {br_int(proj_peak)} projetos" if proj_peak is not None else "Pico: n/d"
                
                # Calcular média anual dos projetos
                proj_yearly_avg = sum([v for v in cur_vals if v is not None]) / len([v for v in cur_vals if v is not None]) if cur_vals else None
                highlights['LancProj Yearly Avg'] = f"Média anual: {br_int(proj_yearly_avg)} projetos" if proj_yearly_avg is not None else "Média anual: n/d"
                
                # Manter cálculo de tendência para as setas
                proj_trend = calc_trend(cur_vals)
                highlights['LancProj Trend'] = proj_trend

        # Gerar datasets mensais de empreendimentos (nº de projetos)
        df_lanc_m = sheets[lanc_month_sheet]
        data_dict['LancProj Monthly'] = build_monthly_dataset_bracket(df_lanc_m)

    if lanc_quart_sheet:
        data_dict['Lanc Quarterly'] = build_quarterly_dataset(sheets[lanc_quart_sheet], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets[lanc_quart_sheet], data_type='number')
        if best_value is not None and best_quarter:
            highlights['Lanc Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[lanc_quart_sheet])
        if observation:
            highlights['Lanc Quarterly Obs'] = observation
            highlights['LancProj Quarterly Obs'] = observation  # Mesma observação para projetos

        # Dataset trimestral de empreendimentos
        data_dict['LancProj Quarterly'] = build_quarterly_dataset_bracket(sheets[lanc_quart_sheet])

        # Para empreendimentos (entre colchetes), também encontrar melhor trimestre
        df_q = clean_dataframe(sheets[lanc_quart_sheet])
        if not df_q.empty and df_q.shape[1] > 1:
            # Encontrar melhor trimestre para número de projetos
            best_proj_value = None
            best_proj_quarter = None
            
            for idx in range(len(df_q)):
                first_cell = str(df_q.iloc[idx, 0]).upper()
                if first_cell in ['1T', '2T', '3T', '4T']:
                    last_col_idx = len(df_q.columns) - 1
                    proj_val = parse_bracket_number(df_q.iloc[idx, last_col_idx])
                    
                    if proj_val is not None:
                        if best_proj_value is None or proj_val > best_proj_value:
                            best_proj_value = proj_val
                            best_proj_quarter = first_cell
            
            if best_proj_value is not None and best_proj_quarter:
                highlights['LancProj Quarterly'] = f"Melhor trimestre (projetos): {best_proj_quarter} - {br_int(best_proj_value)}"

    if lanc_year_sheet:
        data, var = build_yearly_dataset(sheets[lanc_year_sheet], is_percent=False)
        data_dict['Lanc Yearly'] = (data, var)
        df_a = clean_dataframe(sheets[lanc_year_sheet])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['Lanc Annual'] = f"{year}: {br_int(val)} ({var_str})"
                # Métrica anual de empreendimentos (projetos)
                proj_val = parse_bracket_number(row.iloc[1])
                if proj_val is not None:
                    highlights['LancProj Annual'] = f"{year}: {br_int(proj_val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[lanc_year_sheet])
        if observation:
            highlights['Lanc Annual Obs'] = observation
            highlights['LancProj Annual Obs'] = observation  # Mesma observação para projetos

        # Dataset anual de empreendimentos
        proj_data, proj_var = build_yearly_dataset_bracket(sheets[lanc_year_sheet])
        data_dict['LancProj Yearly'] = (proj_data, proj_var)

    # ---------------- Preços Oferta ----------------
    oferta_price_month = next((n for n in sheets if n.startswith('Oferta Valor Médio Ponderado Me')), None)
    oferta_price_quart = next((n for n in sheets if n.startswith('Oferta Valor Médio Ponderado Tr')), None)
    oferta_price_year = next((n for n in sheets if n.startswith('Oferta Valor Médio Ponderado An')), None)

    if oferta_price_month:
        po_monthly = build_monthly_dataset(sheets[oferta_price_month], is_percent=False)
        data_dict['Precos Oferta Monthly'] = po_monthly

        # Novos insights usando dados MoM/YoY do Excel
        po_insights = format_new_insights(sheets[oferta_price_month], data_type='currency', month_ref=month_ref)
        
        highlights['PrecosOferta MoM'] = po_insights['mom']
        highlights['PrecosOferta YoY'] = po_insights['yoy']
        highlights['PrecosOferta Peak'] = po_insights['peak']
        highlights['PrecosOferta Yearly Avg'] = po_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if po_monthly['datasets']:
            cur = po_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['PrecosOferta Trend'] = trend

    if oferta_price_quart:
        data_dict['Precos Oferta Quarterly'] = build_quarterly_dataset(sheets[oferta_price_quart], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets[oferta_price_quart], data_type='currency')
        if best_value is not None and best_quarter:
            highlights['PrecosOferta Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_currency(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[oferta_price_quart])
        if observation:
            highlights['PrecosOferta Quarterly Obs'] = observation

    if oferta_price_year:
        data, var = build_yearly_dataset(sheets[oferta_price_year], is_percent=False)
        data_dict['Precos Yearly'] = (data, var)
        df_a = clean_dataframe(sheets[oferta_price_year])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['Precos Annual'] = f"{year}: {br_currency(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[oferta_price_year])
        if observation:
            highlights['PrecosOferta Annual Obs'] = observation

    # ---------------- Preços Venda ----------------
    venda_price_month = next((n for n in sheets if n.startswith('Venda Valor Médio Ponderado Men')), None)
    venda_price_quart = next((n for n in sheets if n.startswith('Venda Valor Médio Ponderado Tri')), None)

    if venda_price_month:
        pv_monthly = build_monthly_dataset(sheets[venda_price_month], is_percent=False)
        data_dict['Precos Venda Monthly'] = pv_monthly

        # Novos insights usando dados MoM/YoY do Excel
        pv_insights = format_new_insights(sheets[venda_price_month], data_type='currency', month_ref=month_ref)
        
        highlights['PrecosVenda MoM'] = pv_insights['mom']
        highlights['PrecosVenda YoY'] = pv_insights['yoy']
        highlights['PrecosVenda Peak'] = pv_insights['peak']
        highlights['PrecosVenda Yearly Avg'] = pv_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if pv_monthly['datasets']:
            cur = pv_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['PrecosVenda Trend'] = trend

    if venda_price_quart:
        data_dict['Precos Venda Quarterly'] = build_quarterly_dataset(sheets[venda_price_quart], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets[venda_price_quart], data_type='currency')
        if best_value is not None and best_quarter:
            highlights['PrecosVenda Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_currency(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[venda_price_quart])
        if observation:
            highlights['PrecosVenda Quarterly Obs'] = observation

    # Buscar dados anuais de preços de venda
    venda_price_year = next((n for n in sheets if n.startswith('Venda Valor Médio Ponderado Anu')), None)
    if venda_price_year:
        data, var = build_yearly_dataset(sheets[venda_price_year], is_percent=False)
        data_dict['Precos Venda Yearly'] = (data, var)
        df_a = clean_dataframe(sheets[venda_price_year])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['PrecosVenda Annual'] = f"{year}: {br_currency(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[venda_price_year])
        if observation:
            highlights['PrecosVenda Annual Obs'] = observation

    # ---------------- VGV ----------------
    if 'VGV Mensal (R$ Milhões)' in sheets:
        vgv_monthly = build_monthly_dataset(sheets['VGV Mensal (R$ Milhões)'], is_percent=False)
        data_dict['VGV Monthly'] = vgv_monthly

        # Novos insights usando dados MoM/YoY do Excel
        vgv_insights = format_new_insights(sheets['VGV Mensal (R$ Milhões)'], data_type='currency', is_millions=True, month_ref=month_ref)
        
        highlights['VGV MoM'] = vgv_insights['mom']
        highlights['VGV YoY'] = vgv_insights['yoy']
        highlights['VGV Peak'] = vgv_insights['peak']
        highlights['VGV Yearly Avg'] = vgv_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if vgv_monthly['datasets']:
            cur = vgv_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['VGV Trend'] = trend

    if 'VGV Trimestral (R$ Milhões)' in sheets:
        data_dict['VGV Quarterly'] = build_quarterly_dataset(sheets['VGV Trimestral (R$ Milhões)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['VGV Trimestral (R$ Milhões)'], data_type='currency')
        if best_value is not None and best_quarter:
            highlights['VGV Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_currency(best_value)}M"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['VGV Trimestral (R$ Milhões)'])
        if observation:
            highlights['VGV Quarterly Obs'] = observation

    if 'VGV Anual (R$ Milhões)' in sheets:
        data, var = build_yearly_dataset(sheets['VGV Anual (R$ Milhões)'], is_percent=False)
        data_dict['VGV Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['VGV Anual (R$ Milhões)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['VGV Annual'] = f"{year}: {br_currency(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['VGV Anual (R$ Milhões)'])
        if observation:
            highlights['VGV Annual Obs'] = observation

    # ---------------- VGL (Valor Geral de Lançamentos) ----------------
    # Valores monetários de lançamentos (R$ Milhões)
    if 'VGL Mensal (R$ Milhões)' in sheets:
        vgl_monthly = build_monthly_dataset(sheets['VGL Mensal (R$ Milhões)'], is_percent=False)
        data_dict['VGL Monthly'] = vgl_monthly

        # Novos insights usando dados MoM/YoY do Excel
        vgl_insights = format_new_insights(sheets['VGL Mensal (R$ Milhões)'], data_type='currency', is_millions=True, month_ref=month_ref)
        
        highlights['VGL MoM'] = vgl_insights['mom']
        highlights['VGL YoY'] = vgl_insights['yoy']
        highlights['VGL Peak'] = vgl_insights['peak']
        highlights['VGL Yearly Avg'] = vgl_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if vgl_monthly['datasets']:
            cur = vgl_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['VGL Trend'] = trend

    if 'VGL Trimestral (R$ Milhões)' in sheets:
        data_dict['VGL Quarterly'] = build_quarterly_dataset(sheets['VGL Trimestral (R$ Milhões)'], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets['VGL Trimestral (R$ Milhões)'], data_type='currency')
        if best_value is not None and best_quarter:
            highlights['VGL Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_currency(best_value)}M"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['VGL Trimestral (R$ Milhões)'])
        if observation:
            highlights['VGL Quarterly Obs'] = observation

    if 'VGL Anual (R$ Milhões)' in sheets:
        data, var = build_yearly_dataset(sheets['VGL Anual (R$ Milhões)'], is_percent=False)
        data_dict['VGL Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['VGL Anual (R$ Milhões)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['VGL Annual'] = f"{year}: {br_currency(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['VGL Anual (R$ Milhões)'])
        if observation:
            highlights['VGL Annual Obs'] = observation

    # ---------------- Distratos ----------------
    # Unidades distratadas (cancelamentos)
    if 'Distratos Mensais (Unidades)' in sheets:
        dist_monthly = build_monthly_dataset(sheets['Distratos Mensais (Unidades)'], is_percent=False)
        data_dict['Distratos Monthly'] = dist_monthly

        # Novos insights usando dados MoM/YoY do Excel
        dist_insights = format_new_insights(sheets['Distratos Mensais (Unidades)'], data_type='number', month_ref=month_ref)
        
        highlights['Distratos MoM'] = dist_insights['mom']
        highlights['Distratos YoY'] = dist_insights['yoy']
        highlights['Distratos Peak'] = dist_insights['peak']
        highlights['Distratos Yearly Avg'] = dist_insights['yearly_avg']
        
        # Manter cálculo de tendência original para as setas
        if dist_monthly['datasets']:
            cur = dist_monthly['datasets'][-1]['data']
            trend = calc_trend(cur)
            highlights['Distratos Trend'] = trend

    # Procurar planilha trimestral de distratos (nome pode estar truncado)
    dist_quart_sheet = next((n for n in sheets if n.startswith('Distratos Trimestrais')), None)
    if dist_quart_sheet:
        data_dict['Distratos Quarterly'] = build_quarterly_dataset(sheets[dist_quart_sheet], is_percent=False)
        
        # Encontrar melhor trimestre (não sempre o último)
        best_value, best_quarter = find_best_quarter_with_performance(sheets[dist_quart_sheet], data_type='number')
        if best_value is not None and best_quarter:
            highlights['Distratos Quarterly'] = f"Melhor trimestre: {best_quarter} - {br_int(best_value)}"
        
        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets[dist_quart_sheet])
        if observation:
            highlights['Distratos Quarterly Obs'] = observation

    if 'Distratos Anuais (Unidades)' in sheets:
        data, var = build_yearly_dataset(sheets['Distratos Anuais (Unidades)'], is_percent=False)
        data_dict['Distratos Yearly'] = (data, var)
        df_a = clean_dataframe(sheets['Distratos Anuais (Unidades)'])
        if not df_a.empty:
            for idx in range(len(df_a) - 1, -1, -1):
                row = df_a.iloc[idx]
                val = parse_number(row.iloc[1])
                if val is None:
                    continue
                year = str(row.iloc[0])
                var_str = str(row.iloc[2]) if len(row) > 2 and not pd.isna(row.iloc[2]) else ''
                highlights['Distratos Annual'] = f"{year}: {br_int(val)} ({var_str})"
                break

        # Extrair observações sobre dados incompletos
        observation = extract_observation_from_sheet(sheets['Distratos Anuais (Unidades)'])
        if observation:
            highlights['Distratos Annual Obs'] = observation

    # ---------------- Tabelas de Regiões ----------------
    # Construir tabelas por região para ofertas, vendas e valores ponderados (preços)
    region_tables: dict[str, str] = {}

    # Extrair mês/ano no formato MM/AA a partir do nome do arquivo (ex.: 2025_10 -> 10/25)
    mm_aa = None
    match_fn = re.search(r"(\d{4})_(\d{2})", os.path.basename(input_excel))
    if match_fn:
        year_str, month_str = match_fn.groups()
        mm_aa = f"{int(month_str):02d}/{year_str[-2:]}"

    # Definir títulos fixos para tabelas de preços (sem mês/ano)
    precos_oferta_title = "Preços de Oferta por Região (R$)"
    precos_venda_title = "Preços de Venda por Região (R$)"

    # Mapeamento: substring da planilha -> chave -> título (expandido com mais variações)
    region_mapping = [
        # Ofertas por região
        ('ofertas por região', 'ofertas', 'Ofertas por Região'),
        ('oferta por região', 'ofertas', 'Ofertas por Região'),
        ('ofertas região', 'ofertas', 'Ofertas por Região'),
        ('oferta região', 'ofertas', 'Ofertas por Região'),
        
        # Vendas por região  
        ('vendas por região', 'vendas', 'Vendas por Região'),
        ('venda por região', 'vendas', 'Vendas por Região'),
        ('vendas região', 'vendas', 'Vendas por Região'),
        ('venda região', 'vendas', 'Vendas por Região'),
        
        # Preços de oferta
        ('oferta valor ponderado', 'precos_oferta', precos_oferta_title),
        ('ofertas valor ponderado', 'precos_oferta', precos_oferta_title),
        ('preço oferta região', 'precos_oferta', precos_oferta_title),
        ('preços oferta região', 'precos_oferta', precos_oferta_title),
        ('valor ponderado oferta', 'precos_oferta', precos_oferta_title),
        
        # Preços de venda
        ('venda valor ponderado', 'precos_venda', precos_venda_title),
        ('vendas valor ponderado', 'precos_venda', precos_venda_title),
        ('preço venda região', 'precos_venda', precos_venda_title),
        ('preços venda região', 'precos_venda', precos_venda_title),
        ('valor ponderado venda', 'precos_venda', precos_venda_title),
    ]

    # Debug: Mostrar planilhas disponíveis para tabelas regionais
    print(f"🔍 DEBUG: Procurando planilhas regionais...")
    print(f"📋 Planilhas disponíveis no Excel:")
    for sheet_name in sheets.keys():
        print(f"  - {sheet_name}")
    print()

    # Busca primária: mapeamento exato com validação de dados regionais
    found_tables = set()
    for sheet_name, df in sheets.items():
        name_lower = sheet_name.lower()
        for substr, key, title in region_mapping:
            if substr in name_lower and key not in found_tables:
                print(f"🔍 Testando planilha: '{sheet_name}' para {key}")
                
                cleaned = clean_dataframe(df)
                if cleaned.empty:
                    print(f"⚠️  Planilha '{sheet_name}' está vazia após limpeza")
                    continue
                
                # NOVA VALIDAÇÃO: Verificar se contém dados regionais
                if not is_regional_data(cleaned):
                    print(f"❌ Planilha '{sheet_name}' contém dados mensais, não regionais")
                    print(f"   Primeira coluna: {list(cleaned.iloc[:5, 0].astype(str))}")
                    continue
                    
                print(f"✅ Encontrada planilha regional VÁLIDA: '{sheet_name}' -> {key}")
                print(f"   Primeira coluna (regiões): {list(cleaned.iloc[:5, 0].astype(str))}")
                    
                parsed = parse_region_table(cleaned)
                # Se for tabela de preços, formatar valores com duas casas decimais e zero como '-'
                if key in ('precos_oferta', 'precos_venda'):
                    df_price = parsed.copy()
                    for col in df_price.columns[1:]:  # não modificar coluna Região
                        def format_val(val):
                            num = parse_number(val)
                            # Considerar None ou 0 como sem valor
                            if num is None or abs(num) < 1e-6:
                                return '-'
                            # Formatar como número com duas casas decimais (formato brasileiro)
                            return br_float(num, decimals=2)
                        df_price[col] = df_price[col].apply(format_val)
                    # Para tabelas de preço, gera HTML com os valores já formatados
                    html_table = create_region_table_html(df_price, title)
                    region_tables[key] = html_table
                else:
                    region_tables[key] = create_region_table_html(parsed, title)
                found_tables.add(key)
                break

    # Busca secundária: padrões mais flexíveis se não encontrar algumas tabelas
    if len(region_tables) < 4:  # Esperamos 4 tabelas
        print(f"⚠️  Apenas {len(region_tables)}/4 tabelas encontradas. Tentando busca flexível...")
        
        # Padrões flexíveis
        flexible_patterns = [
            # Para ofertas
            (lambda name: 'oferta' in name and 'região' in name and 'ponderado' not in name, 
             'ofertas', 'Ofertas por Região'),
            # Para vendas
            (lambda name: 'venda' in name and 'região' in name and 'ponderado' not in name, 
             'vendas', 'Vendas por Região'),
            # Para preços de oferta
            (lambda name: 'oferta' in name and ('ponderado' in name or 'preço' in name),
             'precos_oferta', precos_oferta_title),
            # Para preços de venda
            (lambda name: 'venda' in name and ('ponderado' in name or 'preço' in name),
             'precos_venda', precos_venda_title),
        ]
        
        for sheet_name, df in sheets.items():
            name_lower = sheet_name.lower()
            for pattern_func, key, title in flexible_patterns:
                if pattern_func(name_lower) and key not in found_tables:
                    print(f"🔍 Testando padrão flexível: '{sheet_name}' para {key}")
                    
                    cleaned = clean_dataframe(df)
                    if cleaned.empty:
                        continue
                    
                    # VALIDAÇÃO CRÍTICA: Verificar se são dados regionais
                    if not is_regional_data(cleaned):
                        print(f"❌ Padrão flexível rejeitado: '{sheet_name}' não contém dados regionais")
                        print(f"   Primeira coluna detectada: {list(cleaned.iloc[:3, 0].astype(str))}")
                        continue
                        
                    print(f"✅ Padrão flexível aceito: '{sheet_name}' -> {key}")
                    print(f"   Regiões encontradas: {list(cleaned.iloc[:3, 0].astype(str))}")
                        
                    parsed = parse_region_table(cleaned)
                    if key in ('precos_oferta', 'precos_venda'):
                        df_price = parsed.copy()
                        for col in df_price.columns[1:]:
                            def format_val(val):
                                num = parse_number(val)
                                if num is None or abs(num) < 1e-6:
                                    return '-'
                                return br_float(num, decimals=2)
                            df_price[col] = df_price[col].apply(format_val)
                        html_table = create_region_table_html(df_price, title)
                        region_tables[key] = html_table
                    else:
                        region_tables[key] = create_region_table_html(parsed, title)
                    found_tables.add(key)
                    break

    # Debug: Mostrar quantas tabelas regionais foram criadas
    print(f"📊 Tabelas regionais criadas: {len(region_tables)}")
    for key, html in region_tables.items():
        print(f"  ✅ {key}: {len(html)} chars de HTML")
    print()

    # Gera HTML
    html_content = generate_html(data_dict, report_date, month_ref, highlights)
    
    # Insere as tabelas de região nas seções corretas
    if region_tables:
        print(f"📋 Inserindo {len(region_tables)} tabelas regionais no HTML...")
        for key in region_tables:
            print(f"  - Inserindo tabela: {key}")
        html_content = insert_region_tables(html_content, region_tables)
        print("✅ Tabelas regionais inseridas com sucesso!")
    else:
        print("⚠️  Nenhuma tabela regional encontrada!")
        print("💡 Verificar se as planilhas têm os nomes esperados:")
        print("   - Ofertas por Região")
        print("   - Vendas por Região") 
        print("   - Oferta Valor Ponderado")
        print("   - Venda Valor Ponderado")
    print()
    # Escrever HTML final
    with open(output_html, 'w', encoding='utf-8') as f:
        f.write(html_content)

    print(f"✅ HTML report generated: {output_html}")


if __name__ == '__main__':
    main()
