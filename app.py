import io
import re
import unicodedata
from typing import Dict, List, Optional, Tuple

import streamlit as st

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
    OPENPYXL_IMPORT_ERROR = ""
except ImportError as exc:
    OPENPYXL_OK = False
    OPENPYXL_IMPORT_ERROR = str(exc)


# ============================================================
# CONFIGURAÇÃO DO STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Consolidador de Fretes",
    page_icon="📊",
    layout="wide",
)

if not OPENPYXL_OK:
    st.error(
        "A dependência openpyxl não foi instalada no ambiente. "
        "Confirme se o arquivo requirements.txt está no mesmo "
        "repositório do app.py. Detalhes: " + OPENPYXL_IMPORT_ERROR
    )
    st.stop()


# ============================================================
# CONFIGURAÇÕES DO PROCESSO
# ============================================================

# Quantas linhas do início de cada aba serão analisadas
# para descobrir onde está o cabeçalho.
MAX_HEADER_SCAN_ROWS = 50

# Para considerar que uma linha é realmente um cabeçalho,
# ela deve possuir pelo menos esta quantidade de campos reconhecidos.
MIN_HEADER_MATCHES = 3


OUTPUT_COLUMNS = [
    "EMISSÃO",
    "TRANSPORTADOR",
    "PLACA",
    "TIPOLOGIA",
    "PAGA",
    "ORIGEM",
    "DESTINO",
    "VENCIMENTO",
    "KM TOTAL",
    "ADC (KM",
    "FRETE SAÍDA",
    "PEDÁGIO",
    "FRANQUIA",
    "OUTROS CRÉDITOS",
    "DESCONTO",
    "ABASTECIMENTO",
    "FRETE SEM DESCONTO",
    "TOTAL COM DESCONTO",
    "O.S VIAG",
    "O.S FECH",
    "MOTIVO",
]


COLUMN_ALIASES = {
    "EMISSÃO": [
        "EMISSÃO",
        "EMISSAO",
    ],
    "TRANSPORTADOR": [
        "TRANSPORTADOR",
        "TRANSPORTADOR/MOTORISTA",
        "TRANSPORTADOR / MOTORISTA",
    ],
    "PLACA": [
        "PLACA",
    ],
    "TIPOLOGIA": [
        "TIPOLOGIA",
    ],
    "PAGA": [
        "PAGA",
    ],
    "ORIGEM": [
        "ORIGEM",
    ],
    "DESTINO": [
        "DESTINO",
    ],
    "VENCIMENTO": [
        "VENCIMENTO",
    ],
    "KM TOTAL": [
        "KM TOTAL",
    ],
    "ADC (KM": [
        "ADC (KM",
        "ADC KM",
        "ADC(KM",
        "ADC (KM)",
        "ADC KM)",
    ],
    "FRETE SAÍDA": [
        "FRETE SAÍDA",
        "FRETE SAIDA",
    ],
    "PEDÁGIO": [
        "PEDÁGIO",
        "PEDAGIO",
    ],
    "FRANQUIA": [
        "FRANQUIA",
        "FRANQUINA",
    ],
    "OUTROS CRÉDITOS": [
        "OUTROS CRÉDITOS",
        "OUTROS CREDITOS",
    ],
    "DESCONTO": [
        "DESCONTO",
    ],
    "ABASTECIMENTO": [
        "ABASTECIMENTO",
        "COMBUSTÍVEL",
        "COMBUSTIVEL",
    ],
    "FRETE SEM DESCONTO": [
        "FRETE SEM DESCONTO",
    ],
    "TOTAL COM DESCONTO": [
        "TOTAL COM DESCONTO",
        "TOTAL COM DIFERENÇA / DESCONTO",
        "TOTAL COM DIFERENCA / DESCONTO",
        "TOTAL COM DIFERENÇA/DESCONTO",
        "TOTAL COM DIFERENCA/DESCONTO",
    ],
    "O.S VIAG": [
        "O.S VIAG",
        "O.S. VIAG",
        "OS VIAG",
        "O.S VIAGEM",
        "O.S. VIAGEM",
        "OS VIAGEM",
    ],
    "O.S FECH": [
        "O.S FECH",
        "O.S. FECH",
        "OS FECH",
        "O.S FECHAMENTO",
        "O.S. FECHAMENTO",
        "OS FECHAMENTO",
    ],
    "MOTIVO": [
        "MOTIVO",
        "OBSERVAÇÃO",
        "OBSERVACAO",
    ],
}


# ============================================================
# NORMALIZAÇÃO
# ============================================================

def normalize_text(value) -> str:
    """
    Normaliza apenas para COMPARAR nomes de cabeçalhos.

    Não altera os dados reais da planilha.
    """

    if value is None:
        return ""

    text = str(value).replace("\xa0", " ")
    text = text.strip().lower()

    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char
        for char in text
        if not unicodedata.combining(char)
    )

    # Remove espaços, barras, pontos, hífens etc.
    # Exemplo:
    # "O.S VIAG" -> "osviag"
    # "PEDÁGIO" -> "pedagio"
    return re.sub(r"[^a-z0-9]+", "", text)


NORMALIZED_ALIASES = {
    target: {
        normalize_text(alias)
        for alias in aliases
    }
    for target, aliases in COLUMN_ALIASES.items()
}


def is_blank(value) -> bool:
    """
    Considera vazio somente None ou string sem conteúdo.
    Zero NÃO é considerado vazio.
    """
    if value is None:
        return True

    if isinstance(value, str):
        return value.strip() == ""

    return False


def is_summary_sheet(sheet_name: str) -> bool:
    """
    Ignora qualquer aba que tenha a palavra RESUMO no nome.
    """
    return "resumo" in normalize_text(sheet_name)


def targets_for_header(header_value) -> List[str]:
    """
    Retorna quais campos de saída podem corresponder ao
    texto encontrado no cabeçalho.
    """
    normalized = normalize_text(header_value)

    if not normalized:
        return []

    matches = []

    for target, aliases in NORMALIZED_ALIASES.items():
        if normalized in aliases:
            matches.append(target)

    return matches


# ============================================================
# DETECÇÃO AUTOMÁTICA DO CABEÇALHO
# ============================================================

def detect_header_row(ws) -> Tuple[Optional[int], int, List[str]]:
    """
    Procura automaticamente a linha do cabeçalho.

    A pontuação de cada linha é baseada na quantidade de
    campos DIFERENTES reconhecidos.

    Retorna:
    - número da linha do cabeçalho;
    - quantidade de campos reconhecidos;
    - lista dos campos reconhecidos.
    """

    best_row = None
    best_score = 0
    best_targets = []

    last_row_to_scan = min(
        ws.max_row,
        MAX_HEADER_SCAN_ROWS,
    )

    for row_idx in range(1, last_row_to_scan + 1):

        found_targets = set()

        for cell in ws[row_idx]:

            for target in targets_for_header(cell.value):
                found_targets.add(target)

        score = len(found_targets)

        if score > best_score:
            best_score = score
            best_row = row_idx
            best_targets = sorted(found_targets)

    if best_score < MIN_HEADER_MATCHES:
        return None, best_score, best_targets

    return best_row, best_score, best_targets


# ============================================================
# MAPEAMENTO DE COLUNAS
# ============================================================

def map_source_columns(
    ws,
    header_row: int,
) -> Tuple[
    Dict[str, dict],
    Dict[str, List[dict]],
    List[str],
    List[dict],
]:
    """
    Faz o mapeamento dos cabeçalhos da aba.

    Regra de segurança:
    - 0 candidatos -> coluna ausente;
    - 1 candidato -> coluna mapeada;
    - 2 ou mais candidatos -> AMBIGUIDADE.
      O sistema NÃO escolhe sozinho.

    Retorna:
    - mapped_columns;
    - ambiguous_columns;
    - missing_columns;
    - mapping_audit.
    """

    candidates_by_target = {
        target: []
        for target in OUTPUT_COLUMNS
    }

    # Verifica todas as células da linha do cabeçalho
    for cell in ws[header_row]:

        matched_targets = targets_for_header(
            cell.value
        )

        for target in matched_targets:

            candidates_by_target[target].append(
                {
                    "column_index": cell.column,
                    "column_letter": get_column_letter(
                        cell.column
                    ),
                    "header_original": cell.value,
                }
            )

    mapped_columns = {}
    ambiguous_columns = {}
    missing_columns = []
    mapping_audit = []

    for target in OUTPUT_COLUMNS:

        candidates = candidates_by_target[target]

        if len(candidates) == 0:

            missing_columns.append(target)

            mapping_audit.append(
                {
                    "CAMPO DESTINO": target,
                    "STATUS": "AUSENTE",
                    "COLUNA ORIGEM": "",
                    "CABEÇALHO ORIGINAL": "",
                }
            )

        elif len(candidates) == 1:

            mapped_columns[target] = candidates[0]

            mapping_audit.append(
                {
                    "CAMPO DESTINO": target,
                    "STATUS": "MAPEADA",
                    "COLUNA ORIGEM": candidates[0][
                        "column_letter"
                    ],
                    "CABEÇALHO ORIGINAL": candidates[0][
                        "header_original"
                    ],
                }
            )

        else:

            ambiguous_columns[target] = candidates

            columns_text = ", ".join(
                candidate["column_letter"]
                for candidate in candidates
            )

            headers_text = " | ".join(
                str(candidate["header_original"])
                for candidate in candidates
            )

            mapping_audit.append(
                {
                    "CAMPO DESTINO": target,
                    "STATUS": "AMBÍGUA - NÃO IMPORTADA",
                    "COLUNA ORIGEM": columns_text,
                    "CABEÇALHO ORIGINAL": headers_text,
                }
            )

    return (
        mapped_columns,
        ambiguous_columns,
        missing_columns,
        mapping_audit,
    )


# ============================================================
# LEITURA DAS LINHAS
# ============================================================

def row_has_recognized_data(
    ws,
    row_idx: int,
    recognized_column_indexes: List[int],
) -> bool:
    """
    Uma linha só entra no consolidado quando possui algum
    valor em pelo menos uma coluna reconhecida.

    Não usa formatação da célula para decidir.
    """

    for col_idx in recognized_column_indexes:

        value = ws.cell(
            row=row_idx,
            column=col_idx,
        ).value

        if not is_blank(value):
            return True

    return False


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================

def process_workbook(
    uploaded_file,
) -> Tuple[
    List[dict],
    List[dict],
    List[dict],
    List[dict],
    List[str],
    List[str],
]:
    """
    Processa o Excel com rastreabilidade.

    Retorna:
    - registros consolidados;
    - rastreabilidade por linha;
    - relatório por aba;
    - mapeamento de colunas;
    - avisos;
    - abas de resumo ignoradas.
    """

    file_bytes = uploaded_file.getvalue()

    # data_only=True:
    # se houver fórmula com valor calculado salvo no Excel,
    # importa o valor exibido, sem recalcular nada.
    workbook = load_workbook(
        io.BytesIO(file_bytes),
        data_only=True,
        read_only=False,
    )

    output_records = []
    trace_records = []
    sheet_reports = []
    mapping_records = []
    warnings = []
    ignored_sheets = []

    for ws in workbook.worksheets:

        sheet_name = ws.title

        # --------------------------------------------------------
        # 1. IGNORAR RESUMO
        # --------------------------------------------------------

        if is_summary_sheet(sheet_name):

            ignored_sheets.append(sheet_name)

            sheet_reports.append(
                {
                    "ABA": sheet_name,
                    "STATUS": "IGNORADA - RESUMO",
                    "LINHA DO CABEÇALHO": "",
                    "CAMPOS RECONHECIDOS": 0,
                    "LINHAS IMPORTADAS": 0,
                    "COLUNAS AUSENTES": "",
                    "COLUNAS AMBÍGUAS": "",
                }
            )

            continue

        # --------------------------------------------------------
        # 2. DESCOBRIR CABEÇALHO
        # --------------------------------------------------------

        (
            header_row,
            header_score,
            header_targets,
        ) = detect_header_row(ws)

        if header_row is None:

            warning = (
                f'A aba "{sheet_name}" foi ignorada porque '
                f'não foi encontrado um cabeçalho confiável '
                f'nas primeiras {MAX_HEADER_SCAN_ROWS} linhas. '
                f'Melhor tentativa: {header_score} campo(s) '
                f'reconhecido(s).'
            )

            warnings.append(warning)

            sheet_reports.append(
                {
                    "ABA": sheet_name,
                    "STATUS": "IGNORADA - CABEÇALHO NÃO IDENTIFICADO",
                    "LINHA DO CABEÇALHO": "",
                    "CAMPOS RECONHECIDOS": header_score,
                    "LINHAS IMPORTADAS": 0,
                    "COLUNAS AUSENTES": "",
                    "COLUNAS AMBÍGUAS": "",
                }
            )

            continue

        # --------------------------------------------------------
        # 3. MAPEAR COLUNAS
        # --------------------------------------------------------

        (
            mapped_columns,
            ambiguous_columns,
            missing_columns,
            mapping_audit,
        ) = map_source_columns(
            ws,
            header_row,
        )

        for item in mapping_audit:

            mapping_records.append(
                {
                    "ABA": sheet_name,
                    "LINHA DO CABEÇALHO": header_row,
                    **item,
                }
            )

        # --------------------------------------------------------
        # 4. AVISAR AMBIGUIDADES
        # --------------------------------------------------------

        if ambiguous_columns:

            for target, candidates in ambiguous_columns.items():

                candidate_text = ", ".join(
                    f'{c["column_letter"]} '
                    f'("{c["header_original"]}")'
                    for c in candidates
                )

                warnings.append(
                    f'Aba "{sheet_name}": o campo '
                    f'"{target}" apareceu em mais de uma '
                    f'coluna ({candidate_text}). '
                    f'Por segurança, "{target}" foi deixado '
                    f'em branco nesta aba.'
                )

        # --------------------------------------------------------
        # 5. IDENTIFICAR TODAS AS COLUNAS RECONHECIDAS
        # --------------------------------------------------------

        recognized_column_indexes = set()

        for info in mapped_columns.values():
            recognized_column_indexes.add(
                info["column_index"]
            )

        for candidates in ambiguous_columns.values():
            for info in candidates:
                recognized_column_indexes.add(
                    info["column_index"]
                )

        recognized_column_indexes = sorted(
            recognized_column_indexes
        )

        if not recognized_column_indexes:

            warnings.append(
                f'A aba "{sheet_name}" possui um possível '
                f'cabeçalho na linha {header_row}, mas nenhuma '
                f'coluna pôde ser utilizada.'
            )

            sheet_reports.append(
                {
                    "ABA": sheet_name,
                    "STATUS": "IGNORADA - SEM COLUNAS UTILIZÁVEIS",
                    "LINHA DO CABEÇALHO": header_row,
                    "CAMPOS RECONHECIDOS": header_score,
                    "LINHAS IMPORTADAS": 0,
                    "COLUNAS AUSENTES": ", ".join(
                        missing_columns
                    ),
                    "COLUNAS AMBÍGUAS": ", ".join(
                        ambiguous_columns.keys()
                    ),
                }
            )

            continue

        # --------------------------------------------------------
        # 6. PERCORRER LINHAS DE DADOS
        # --------------------------------------------------------

        imported_rows = 0

        for row_idx in range(
            header_row + 1,
            ws.max_row + 1,
        ):

            if not row_has_recognized_data(
                ws,
                row_idx,
                recognized_column_indexes,
            ):
                continue

            output_record = {}

            for target in OUTPUT_COLUMNS:

                source = mapped_columns.get(
                    target
                )

                if source is None:

                    # Ausente ou ambígua:
                    # não inventa dado.
                    output_record[target] = ""

                else:

                    value = ws.cell(
                        row=row_idx,
                        column=source[
                            "column_index"
                        ],
                    ).value

                    output_record[target] = (
                        ""
                        if value is None
                        else value
                    )

            output_records.append(
                output_record
            )

            output_excel_row = (
                len(output_records) + 1
            )

            # ----------------------------------------------------
            # RASTREABILIDADE DA LINHA
            # ----------------------------------------------------

            trace_record = {
                "LINHA NO RESUMO": output_excel_row,
                "ABA ORIGEM": sheet_name,
                "LINHA EXCEL ORIGEM": row_idx,
                "O.S VIAG": output_record.get(
                    "O.S VIAG",
                    "",
                ),
                "O.S FECH": output_record.get(
                    "O.S FECH",
                    "",
                ),
                "PEDÁGIO": output_record.get(
                    "PEDÁGIO",
                    "",
                ),
                "TRANSPORTADOR": output_record.get(
                    "TRANSPORTADOR",
                    "",
                ),
                "PLACA": output_record.get(
                    "PLACA",
                    "",
                ),
            }

            trace_records.append(
                trace_record
            )

            imported_rows += 1

        # --------------------------------------------------------
        # 7. RELATÓRIO DA ABA
        # --------------------------------------------------------

        sheet_reports.append(
            {
                "ABA": sheet_name,
                "STATUS": "PROCESSADA",
                "LINHA DO CABEÇALHO": header_row,
                "CAMPOS RECONHECIDOS": header_score,
                "LINHAS IMPORTADAS": imported_rows,
                "COLUNAS AUSENTES": (
                    ", ".join(missing_columns)
                    if missing_columns
                    else ""
                ),
                "COLUNAS AMBÍGUAS": (
                    ", ".join(
                        ambiguous_columns.keys()
                    )
                    if ambiguous_columns
                    else ""
                ),
            }
        )

    return (
        output_records,
        trace_records,
        sheet_reports,
        mapping_records,
        warnings,
        ignored_sheets,
    )


# ============================================================
# CRIAÇÃO DOS ARQUIVOS EXCEL
# ============================================================

def style_header(ws):
    """
    Formata apenas o cabeçalho.
    Não altera nenhum valor.
    """

    fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    font = Font(
        color="FFFFFF",
        bold=True,
    )

    for cell in ws[1]:

        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )


def auto_width(ws, max_width: int = 35):
    """
    Ajusta largura sem alterar valores.
    """

    for col_idx in range(
        1,
        ws.max_column + 1,
    ):

        max_length = 0

        # Analisa até 300 linhas para evitar lentidão
        last_row = min(
            ws.max_row,
            300,
        )

        for row_idx in range(
            1,
            last_row + 1,
        ):

            value = ws.cell(
                row=row_idx,
                column=col_idx,
            ).value

            if value is None:
                continue

            max_length = max(
                max_length,
                len(str(value)),
            )

        ws.column_dimensions[
            get_column_letter(col_idx)
        ].width = min(
            max(max_length + 2, 12),
            max_width,
        )


def write_table_to_sheet(
    ws,
    columns: List[str],
    records: List[dict],
):
    """
    Escreve uma tabela em uma aba usando openpyxl.
    """

    for col_idx, column_name in enumerate(
        columns,
        start=1,
    ):
        ws.cell(
            row=1,
            column=col_idx,
            value=column_name,
        )

    for row_idx, record in enumerate(
        records,
        start=2,
    ):

        for col_idx, column_name in enumerate(
            columns,
            start=1,
        ):

            ws.cell(
                row=row_idx,
                column=col_idx,
                value=record.get(
                    column_name,
                    "",
                ),
            )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    style_header(ws)
    auto_width(ws)


def create_summary_excel(
    output_records: List[dict],
) -> bytes:
    """
    Gera o arquivo principal com APENAS
    a aba Resumo Geral.
    """

    wb = Workbook()

    ws = wb.active
    ws.title = "Resumo Geral"

    write_table_to_sheet(
        ws,
        OUTPUT_COLUMNS,
        output_records,
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


def create_audit_excel(
    trace_records: List[dict],
    sheet_reports: List[dict],
    mapping_records: List[dict],
) -> bytes:
    """
    Gera arquivo separado de auditoria.
    """

    wb = Workbook()

    # --------------------------------------------------------
    # RASTREABILIDADE
    # --------------------------------------------------------

    ws_trace = wb.active
    ws_trace.title = "Rastreabilidade"

    trace_columns = [
        "LINHA NO RESUMO",
        "ABA ORIGEM",
        "LINHA EXCEL ORIGEM",
        "O.S VIAG",
        "O.S FECH",
        "PEDÁGIO",
        "TRANSPORTADOR",
        "PLACA",
    ]

    write_table_to_sheet(
        ws_trace,
        trace_columns,
        trace_records,
    )

    # --------------------------------------------------------
    # MAPEAMENTO
    # --------------------------------------------------------

    ws_mapping = wb.create_sheet(
        "Mapeamento"
    )

    mapping_columns = [
        "ABA",
        "LINHA DO CABEÇALHO",
        "CAMPO DESTINO",
        "STATUS",
        "COLUNA ORIGEM",
        "CABEÇALHO ORIGINAL",
    ]

    write_table_to_sheet(
        ws_mapping,
        mapping_columns,
        mapping_records,
    )

    # --------------------------------------------------------
    # RELATÓRIO DAS ABAS
    # --------------------------------------------------------

    ws_report = wb.create_sheet(
        "Relatorio"
    )

    report_columns = [
        "ABA",
        "STATUS",
        "LINHA DO CABEÇALHO",
        "CAMPOS RECONHECIDOS",
        "LINHAS IMPORTADAS",
        "COLUNAS AUSENTES",
        "COLUNAS AMBÍGUAS",
    ]

    write_table_to_sheet(
        ws_report,
        report_columns,
        sheet_reports,
    )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# ============================================================
# FUNÇÕES DE VISUALIZAÇÃO
# ============================================================

def prepare_table_records(
    records: List[dict],
    columns: List[str],
    limit: Optional[int] = None,
) -> List[dict]:
    """
    Prepara registros para exibição no Streamlit sem depender
    diretamente de pandas.
    """
    selected = records if limit is None else records[:limit]
    return [
        {column: record.get(column, "") for column in columns}
        for record in selected
    ]

def clear_previous_result():
    """
    Limpa dados do processamento anterior quando
    um novo arquivo é processado.
    """

    keys = [
        "summary_records",
        "trace_records",
        "sheet_reports",
        "mapping_records",
        "warnings",
        "ignored_sheets",
        "summary_excel",
        "audit_excel",
    ]

    for key in keys:
        st.session_state.pop(
            key,
            None,
        )


# ============================================================
# INTERFACE
# ============================================================

st.title(
    "📊 Consolidador de Planilhas de Fretes"
)

st.write(
    "O sistema identifica automaticamente onde começa "
    "o cabeçalho de cada aba, consolida os registros e "
    "mantém rastreabilidade da origem de cada linha."
)


with st.expander(
    "🔎 Regras de segurança do processamento"
):

    st.markdown(
        f"""
- O cabeçalho **não precisa estar na linha 2**.
- O sistema procura o cabeçalho nas primeiras **{MAX_HEADER_SCAN_ROWS} linhas**.
- Para aceitar uma linha como cabeçalho, precisam ser reconhecidos pelo menos **{MIN_HEADER_MATCHES} campos**.
- Abas com `Resumo` no nome são ignoradas.
- A comparação dos cabeçalhos ignora maiúsculas/minúsculas, acentos, espaços e pontuação.
- Os **valores das células não são normalizados nem recalculados**.
- `TRANSPORTADOR/MOTORISTA` = `TRANSPORTADOR`.
- `COMBUSTÍVEL` = `ABASTECIMENTO`.
- `TOTAL COM DIFERENÇA / DESCONTO` = `TOTAL COM DESCONTO`.
- `OBSERVAÇÃO` = `MOTIVO`.
- `FRANQUINA` = `FRANQUIA`.
- Se uma coluna não existir, o campo fica em branco.
- Se houver **duas colunas possíveis para o mesmo campo**, o sistema não escolhe uma delas: deixa o campo em branco e gera um aviso.
- Cada linha do resumo fica ligada à **aba e linha original do Excel** no relatório de auditoria.
        """
    )


uploaded_file = st.file_uploader(
    "Selecione a planilha original",
    type=["xlsx"],
)


if uploaded_file is not None:

    st.success(
        f"Arquivo carregado: {uploaded_file.name}"
    )

    if st.button(
        "Processar planilha",
        type="primary",
        use_container_width=True,
    ):

        clear_previous_result()

        try:

            with st.spinner(
                "Analisando cabeçalhos e consolidando dados..."
            ):

                (
                    summary_records,
                    trace_records,
                    sheet_reports,
                    mapping_records,
                    warnings,
                    ignored_sheets,
                ) = process_workbook(
                    uploaded_file
                )

                summary_excel = create_summary_excel(
                    summary_records
                )

                audit_excel = create_audit_excel(
                    trace_records,
                    sheet_reports,
                    mapping_records,
                )

            st.session_state[
                "summary_records"
            ] = summary_records

            st.session_state[
                "trace_records"
            ] = trace_records

            st.session_state[
                "sheet_reports"
            ] = sheet_reports

            st.session_state[
                "mapping_records"
            ] = mapping_records

            st.session_state[
                "warnings"
            ] = warnings

            st.session_state[
                "ignored_sheets"
            ] = ignored_sheets

            st.session_state[
                "summary_excel"
            ] = summary_excel

            st.session_state[
                "audit_excel"
            ] = audit_excel

        except Exception as exc:

            st.error(
                "Não foi possível processar o arquivo. "
                f"Detalhes: {exc}"
            )


# ============================================================
# RESULTADOS
# ============================================================

if "summary_records" in st.session_state:

    summary_records = st.session_state[
        "summary_records"
    ]

    trace_records = st.session_state[
        "trace_records"
    ]

    sheet_reports = st.session_state[
        "sheet_reports"
    ]

    mapping_records = st.session_state[
        "mapping_records"
    ]

    warnings = st.session_state[
        "warnings"
    ]

    ignored_sheets = st.session_state[
        "ignored_sheets"
    ]

    summary_excel = st.session_state[
        "summary_excel"
    ]

    audit_excel = st.session_state[
        "audit_excel"
    ]

    processed_count = sum(
        1
        for item in sheet_reports
        if item["STATUS"] == "PROCESSADA"
    )

    ambiguous_count = sum(
        1
        for item in mapping_records
        if item["STATUS"]
        == "AMBÍGUA - NÃO IMPORTADA"
    )

    st.divider()

    col1, col2, col3, col4 = st.columns(
        4
    )

    col1.metric(
        "Abas processadas",
        processed_count,
    )

    col2.metric(
        "Registros consolidados",
        len(summary_records),
    )

    col3.metric(
        "Abas de resumo ignoradas",
        len(ignored_sheets),
    )

    col4.metric(
        "Mapeamentos ambíguos",
        ambiguous_count,
    )

    # --------------------------------------------------------
    # AVISOS
    # --------------------------------------------------------

    if warnings:

        st.subheader(
            "⚠️ Pontos que precisam de atenção"
        )

        for warning in warnings:
            st.warning(warning)

    else:

        st.success(
            "Nenhuma ambiguidade de cabeçalho foi encontrada."
        )

    # --------------------------------------------------------
    # RELATÓRIO DAS ABAS
    # --------------------------------------------------------

    st.subheader(
        "📋 Relatório das abas"
    )

    report_table = prepare_table_records(
        sheet_reports,
        [
            "ABA",
            "STATUS",
            "LINHA DO CABEÇALHO",
            "CAMPOS RECONHECIDOS",
            "LINHAS IMPORTADAS",
            "COLUNAS AUSENTES",
            "COLUNAS AMBÍGUAS",
        ],
    )

    st.dataframe(
        report_table,
        use_container_width=True,
        hide_index=True,
    )

    # --------------------------------------------------------
    # PRÉVIA DO RESUMO
    # --------------------------------------------------------

    st.subheader(
        "📊 Prévia do Resumo Geral"
    )

    summary_table = prepare_table_records(
        summary_records,
        OUTPUT_COLUMNS,
        limit=100,
    )

    st.dataframe(
        summary_table,
        use_container_width=True,
        hide_index=True,
    )

    if len(summary_records) > 100:

        st.caption(
            f"Exibindo as primeiras 100 linhas de "
            f"{len(summary_records)} registros."
        )

    # --------------------------------------------------------
    # RASTREABILIDADE
    # --------------------------------------------------------

    with st.expander(
        "🔍 Consultar rastreabilidade"
    ):

        st.write(
            "Aqui você consegue verificar de qual aba "
            "e de qual linha do Excel cada registro saiu."
        )

        trace_columns = [
            "LINHA NO RESUMO",
            "ABA ORIGEM",
            "LINHA EXCEL ORIGEM",
            "O.S VIAG",
            "O.S FECH",
            "PEDÁGIO",
            "TRANSPORTADOR",
            "PLACA",
        ]

        search_os = st.text_input(
            "Pesquisar O.S VIAG ou O.S FECH",
            placeholder="Ex.: 14030",
        )

        if search_os.strip():
            query = normalize_text(search_os)
            filtered_trace_records = []

            for record in trace_records:
                os_viag = normalize_text(record.get("O.S VIAG", ""))
                os_fech = normalize_text(record.get("O.S FECH", ""))

                if query in os_viag or query in os_fech:
                    filtered_trace_records.append(record)
        else:
            filtered_trace_records = trace_records

        trace_table = prepare_table_records(
            filtered_trace_records,
            trace_columns,
            limit=300,
        )

        st.dataframe(
            trace_table,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # MAPEAMENTO
    # --------------------------------------------------------

    with st.expander(
        "🧭 Ver mapeamento dos cabeçalhos"
    ):

        mapping_table = prepare_table_records(
            mapping_records,
            [
                "ABA",
                "LINHA DO CABEÇALHO",
                "CAMPO DESTINO",
                "STATUS",
                "COLUNA ORIGEM",
                "CABEÇALHO ORIGINAL",
            ],
        )

        st.dataframe(
            mapping_table,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------
    # DOWNLOADS
    # --------------------------------------------------------

    st.subheader(
        "⬇️ Downloads"
    )

    st.download_button(
        label="Baixar Resumo Geral.xlsx",
        data=summary_excel,
        file_name="Resumo_Geral.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    st.download_button(
        label="Baixar Auditoria da Consolidação.xlsx",
        data=audit_excel,
        file_name="Auditoria_Consolidacao.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )
