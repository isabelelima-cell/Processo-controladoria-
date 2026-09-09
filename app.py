import io
import re
import unicodedata
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


st.set_page_config(
    page_title="Consolidador de Fretes",
    page_icon="📊",
    layout="wide",
)


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
    ],
    "O.S FECH": [
        "O.S FECH",
        "O.S. FECH",
        "OS FECH",
    ],
    "MOTIVO": [
        "MOTIVO",
        "OBSERVAÇÃO",
        "OBSERVACAO",
    ],
}


def normalize_text(value) -> str:
    """
    Normaliza somente para comparação de cabeçalhos.
    Não altera os valores existentes nas células.
    """
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ")
    text = text.strip().lower()

    text = unicodedata.normalize("NFKD", text)
    text = "".join(
        char for char in text
        if not unicodedata.combining(char)
    )

    return re.sub(r"[^a-z0-9]+", "", text)


NORMALIZED_ALIASES = {
    target: [normalize_text(alias) for alias in aliases]
    for target, aliases in COLUMN_ALIASES.items()
}


def is_summary_sheet(sheet_name: str) -> bool:
    """
    Ignora abas cujo nome contenha a palavra 'resumo',
    independentemente de acento, espaços ou maiúsculas.
    """
    return "resumo" in normalize_text(sheet_name)


def has_any_value(row: pd.Series) -> bool:
    """
    Identifica se a linha original possui ao menos
    uma célula preenchida.
    """
    for value in row.tolist():
        if value is None:
            continue

        if isinstance(value, str) and value.strip() == "":
            continue

        try:
            if pd.isna(value):
                continue
        except (TypeError, ValueError):
            pass

        return True

    return False


def build_header_lookup(columns) -> Dict[str, object]:
    """
    Cria um mapa:
    cabeçalho normalizado -> cabeçalho original.

    Se houver cabeçalhos repetidos, mantém a primeira ocorrência.
    """
    lookup = {}

    for column in columns:
        normalized = normalize_text(column)

        if normalized and normalized not in lookup:
            lookup[normalized] = column

    return lookup


def find_source_column(
    header_lookup: Dict[str, object],
    target_column: str,
):
    """
    Procura qual coluna da aba corresponde à coluna
    esperada no arquivo final.
    """
    for alias in NORMALIZED_ALIASES[target_column]:
        if alias in header_lookup:
            return header_lookup[alias]

    return None


def detect_header_row(
    excel_file: pd.ExcelFile,
    sheet_name: str,
    max_rows: int = 10,
):
    """
    Procura o cabeçalho nas primeiras `max_rows` linhas da aba.

    A linha escolhida é aquela que contém a maior quantidade
    de cabeçalhos reconhecidos pelas regras de equivalência.

    Retorna:
    - índice da linha para usar no parâmetro header do pandas;
    - quantidade de campos reconhecidos.

    Se nenhuma linha tiver pelo menos 3 campos reconhecidos,
    retorna None.
    """

    preview = excel_file.parse(
        sheet_name=sheet_name,
        header=None,
        nrows=max_rows,
        dtype=object,
        keep_default_na=False,
    )

    best_row = None
    best_score = 0

    for row_index, row in preview.iterrows():

        normalized_values = {
            normalize_text(value)
            for value in row.tolist()
            if normalize_text(value)
        }

        matched_targets = set()

        for target_column, aliases in NORMALIZED_ALIASES.items():
            if any(
                alias in normalized_values
                for alias in aliases
            ):
                matched_targets.add(
                    target_column
                )

        score = len(matched_targets)

        if score > best_score:
            best_score = score
            best_row = int(row_index)

    # Segurança:
    # evita considerar como cabeçalho uma linha qualquer
    # que tenha apenas uma ou duas palavras coincidentes.
    if best_score < 3:
        return None, best_score

    return best_row, best_score


def process_workbook(
    uploaded_file,
) -> Tuple[pd.DataFrame, List[dict], List[str], List[str]]:
    """
    Percorre todas as abas e consolida os dados.

    Retorna:
    - DataFrame final;
    - relatório por aba;
    - abas de resumo ignoradas;
    - avisos encontrados.
    """

    # -------------------------------------------------------------
    # LER O EXCEL SOMENTE COM OS RESULTADOS DAS FÓRMULAS
    # -------------------------------------------------------------
    #
    # data_only=True faz com que:
    #   célula normal  -> copie o valor da célula
    #   célula fórmula -> copie somente o resultado salvo da fórmula
    #
    # Exemplo:
    #   fórmula original: =K10+L10
    #   valor exibido:    1250
    #   valor importado:  1250
    #
    # A fórmula NÃO é levada para a planilha final.
    # -------------------------------------------------------------

    original_bytes = uploaded_file.getvalue()

    values_workbook = load_workbook(
        io.BytesIO(original_bytes),
        data_only=True,
    )

    # Salva uma cópia interna contendo apenas os valores já
    # calculados/salvos no arquivo original.
    values_buffer = io.BytesIO()

    values_workbook.save(
        values_buffer
    )

    values_buffer.seek(0)

    excel_file = pd.ExcelFile(
        values_buffer,
        engine="openpyxl",
    )

    consolidated_frames = []
    report = []
    ignored_sheets = []
    warnings = []

    for sheet_name in excel_file.sheet_names:

        # ---------------------------------------------------------
        # 1. IGNORAR ABAS DE RESUMO
        # ---------------------------------------------------------

        if is_summary_sheet(sheet_name):
            ignored_sheets.append(sheet_name)
            continue

        # ---------------------------------------------------------
        # 2. LOCALIZAR E LER O CABEÇALHO
        # Procura automaticamente nas primeiras 10 linhas da aba.
        # ---------------------------------------------------------

        try:
            header_row, header_score = detect_header_row(
                excel_file=excel_file,
                sheet_name=sheet_name,
                max_rows=10,
            )

            if header_row is None:
                report.append(
                    {
                        "Aba": sheet_name,
                        "Status": (
                            "Ignorada: cabeçalho não encontrado "
                            "nas 10 primeiras linhas"
                        ),
                        "Linhas importadas": 0,
                        "Colunas encontradas": header_score,
                        "Colunas ausentes": len(OUTPUT_COLUMNS),
                        "Ausentes": "-",
                    }
                )

                warnings.append(
                    f'A aba "{sheet_name}" foi ignorada porque '
                    "não foi possível identificar com segurança "
                    "o cabeçalho nas 10 primeiras linhas."
                )

                continue

            df = excel_file.parse(
                sheet_name=sheet_name,
                header=header_row,
                dtype=object,
                keep_default_na=False,
            )

        except Exception as exc:
            report.append(
                {
                    "Aba": sheet_name,
                    "Status": f"Erro ao ler: {exc}",
                    "Linhas importadas": 0,
                    "Colunas encontradas": 0,
                    "Colunas ausentes": len(OUTPUT_COLUMNS),
                    "Ausentes": "-",
                }
            )
            continue

        # ---------------------------------------------------------
        # 3. REMOVER COLUNAS SEM CABEÇALHO
        # ---------------------------------------------------------

        valid_columns = [
            column
            for column in df.columns
            if normalize_text(column)
            and not normalize_text(column).startswith("unnamed")
        ]

        df = df.loc[:, valid_columns]

        # ---------------------------------------------------------
        # 4. REMOVER SOMENTE LINHAS TOTALMENTE VAZIAS
        # ---------------------------------------------------------

        if not df.empty:
            non_empty_mask = df.apply(
                has_any_value,
                axis=1,
            )

            df = df.loc[non_empty_mask].copy()

        # ---------------------------------------------------------
        # 5. CRIAR MAPA DOS CABEÇALHOS ENCONTRADOS
        # ---------------------------------------------------------

        header_lookup = build_header_lookup(
            df.columns
        )

        column_mapping = {}
        missing_columns = []

        for target_column in OUTPUT_COLUMNS:

            source_column = find_source_column(
                header_lookup,
                target_column,
            )

            if source_column is None:
                missing_columns.append(
                    target_column
                )
            else:
                column_mapping[
                    target_column
                ] = source_column

        # ---------------------------------------------------------
        # 6. SE NÃO HOUVER NENHUMA COLUNA CONHECIDA,
        # PROVAVELMENTE NÃO É UMA ABA DE DADOS
        # ---------------------------------------------------------

        if not column_mapping:

            report.append(
                {
                    "Aba": sheet_name,
                    "Status": (
                        "Ignorada: nenhum "
                        "cabeçalho reconhecido"
                    ),
                    "Linhas importadas": 0,
                    "Colunas encontradas": 0,
                    "Colunas ausentes": len(
                        OUTPUT_COLUMNS
                    ),
                    "Ausentes": ", ".join(
                        OUTPUT_COLUMNS
                    ),
                }
            )

            warnings.append(
                f'A aba "{sheet_name}" não possui '
                "nenhum dos cabeçalhos esperados "
                "e não foi importada."
            )

            continue

        # ---------------------------------------------------------
        # 7. CRIAR A ESTRUTURA PADRÃO DA PLANILHA FINAL
        # ---------------------------------------------------------

        output_df = pd.DataFrame(
            index=df.index
        )

        for target_column in OUTPUT_COLUMNS:

            source_column = column_mapping.get(
                target_column
            )

            if source_column is None:

                # Se a coluna não existir nessa aba,
                # deixa em branco.
                output_df[target_column] = ""

            else:

                # Copia exatamente o valor encontrado.
                # Não faz conta, não preenche zero,
                # não altera nome, placa ou valor.
                output_df[target_column] = df[
                    source_column
                ]

        output_df.reset_index(
            drop=True,
            inplace=True,
        )

        consolidated_frames.append(
            output_df
        )

        # ---------------------------------------------------------
        # 8. DETECTAR POSSÍVEL "TIPOLOGIA PAGA"
        # ---------------------------------------------------------

        normalized_headers = {
            normalize_text(column)
            for column in df.columns
        }

        if (
            normalize_text("TIPOLOGIA PAGA")
            in normalized_headers
            and normalize_text("TIPOLOGIA")
            not in normalized_headers
            and normalize_text("PAGA")
            not in normalized_headers
        ):

            warnings.append(
                f'A aba "{sheet_name}" possui '
                'uma coluna literal "TIPOLOGIA PAGA". '
                "Ela não foi dividida automaticamente "
                "entre TIPOLOGIA e PAGA para evitar "
                "alteração ou interpretação indevida "
                "dos dados."
            )

        # ---------------------------------------------------------
        # 9. RELATÓRIO DA ABA
        # ---------------------------------------------------------

        report.append(
            {
                "Aba": sheet_name,
                "Status": "Processada",
                "Linhas importadas": len(
                    output_df
                ),
                "Colunas encontradas": len(
                    column_mapping
                ),
                "Colunas ausentes": len(
                    missing_columns
                ),
                "Ausentes": (
                    ", ".join(missing_columns)
                    if missing_columns
                    else "-"
                ),
            }
        )

    # -------------------------------------------------------------
    # 10. JUNTAR TODAS AS ABAS
    # -------------------------------------------------------------

    if consolidated_frames:

        final_df = pd.concat(
            consolidated_frames,
            ignore_index=True,
        )

        final_df = final_df[
            OUTPUT_COLUMNS
        ]

    else:

        final_df = pd.DataFrame(
            columns=OUTPUT_COLUMNS
        )

    return (
        final_df,
        report,
        ignored_sheets,
        warnings,
    )


def dataframe_to_excel_bytes(
    df: pd.DataFrame,
) -> bytes:
    """
    Cria o arquivo Excel final em memória.
    """

    output = io.BytesIO()

    # -------------------------------------------------------------
    # 1. ESCREVER OS DADOS
    # -------------------------------------------------------------

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        df.to_excel(
            writer,
            index=False,
            sheet_name="Resumo Geral",
        )

    output.seek(0)

    # -------------------------------------------------------------
    # 2. FORMATAR SOMENTE A APRESENTAÇÃO
    # -------------------------------------------------------------

    workbook = load_workbook(
        output
    )

    worksheet = workbook[
        "Resumo Geral"
    ]

    # Congelar cabeçalho
    worksheet.freeze_panes = "A2"

    # Filtro automático
    worksheet.auto_filter.ref = (
        worksheet.dimensions
    )

    # Formatação do cabeçalho
    header_fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78",
    )

    header_font = Font(
        color="FFFFFF",
        bold=True,
    )

    for cell in worksheet[1]:

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )

    # -------------------------------------------------------------
    # 3. AJUSTAR LARGURA DAS COLUNAS
    # Sem alterar os dados.
    # -------------------------------------------------------------

    for col_idx, column_name in enumerate(
        OUTPUT_COLUMNS,
        start=1,
    ):

        max_length = len(
            column_name
        )

        # Para desempenho, verifica no máximo
        # as primeiras 300 linhas.
        for row_idx in range(
            2,
            min(
                worksheet.max_row,
                300,
            )
            + 1,
        ):

            value = worksheet.cell(
                row=row_idx,
                column=col_idx,
            ).value

            if value is not None:

                max_length = max(
                    max_length,
                    len(str(value)),
                )

        width = min(
            max(
                max_length + 2,
                12,
            ),
            35,
        )

        worksheet.column_dimensions[
            get_column_letter(col_idx)
        ].width = width

    # -------------------------------------------------------------
    # 4. SALVAR O ARQUIVO FINAL
    # -------------------------------------------------------------

    final_output = io.BytesIO()

    workbook.save(
        final_output
    )

    final_output.seek(0)

    return final_output.getvalue()


# =================================================================
# INTERFACE STREAMLIT
# =================================================================

st.title(
    "📊 Consolidador de Planilhas de Fretes"
)

st.write(
    "Envie a planilha original. "
    "O sistema percorrerá as abas de vencimento, "
    "ignorará abas de resumo e criará uma única "
    "planilha consolidada."
)


with st.expander(
    "Regras usadas no processamento"
):

    st.markdown(
        """
- O cabeçalho é procurado automaticamente nas **10 primeiras linhas** de cada aba.
- Abas com **Resumo** no nome são ignoradas.
- Maiúsculas/minúsculas, acentos, espaços e pontuação do cabeçalho são ignorados na comparação.
- `TRANSPORTADOR/MOTORISTA` é tratado como `TRANSPORTADOR`.
- `COMBUSTÍVEL` é tratado como `ABASTECIMENTO`.
- `TOTAL COM DIFERENÇA / DESCONTO` é tratado como `TOTAL COM DESCONTO`.
- `OBSERVAÇÃO` é tratada como `MOTIVO`.
- `FRANQUINA` também é reconhecida como `FRANQUIA`.
- Se uma coluna não existir em determinada aba, ela fica **em branco**.
- Nenhum valor financeiro, placa, transportador ou outro dado é inventado ou calculado.
- Se uma célula tiver **fórmula**, somente o **resultado calculado salvo na célula** é importado; a fórmula não é copiada.
- A planilha gerada contém somente valores, como um **colar somente valores** do Excel.
- Linhas totalmente vazias são descartadas.
        """
    )


uploaded_file = st.file_uploader(
    "Selecione a planilha original",
    type=["xlsx"],
)


if uploaded_file is not None:

    st.success(
        f"Arquivo carregado: "
        f"{uploaded_file.name}"
    )

    process_button = st.button(
        "Processar planilha",
        type="primary",
        use_container_width=True,
    )

    if process_button:

        try:

            with st.spinner(
                "Processando as abas..."
            ):

                (
                    final_df,
                    report,
                    ignored_sheets,
                    warnings,
                ) = process_workbook(
                    uploaded_file
                )

                excel_bytes = (
                    dataframe_to_excel_bytes(
                        final_df
                    )
                )

            # Salva no estado da sessão para que
            # os dados não sumam quando clicar
            # no botão de download.
            st.session_state[
                "final_df"
            ] = final_df

            st.session_state[
                "report"
            ] = report

            st.session_state[
                "ignored_sheets"
            ] = ignored_sheets

            st.session_state[
                "warnings"
            ] = warnings

            st.session_state[
                "excel_bytes"
            ] = excel_bytes

        except Exception as exc:

            st.error(
                "Não foi possível processar "
                f"o arquivo: {exc}"
            )


# =================================================================
# RESULTADO
# =================================================================

if "final_df" in st.session_state:

    final_df = st.session_state[
        "final_df"
    ]

    report = st.session_state[
        "report"
    ]

    ignored_sheets = (
        st.session_state[
            "ignored_sheets"
        ]
    )

    warnings = st.session_state[
        "warnings"
    ]

    excel_bytes = st.session_state[
        "excel_bytes"
    ]

    st.divider()

    processed_count = sum(
        1
        for item in report
        if item.get("Status")
        == "Processada"
    )

    col1, col2, col3 = st.columns(
        3
    )

    col1.metric(
        "Abas processadas",
        processed_count,
    )

    col2.metric(
        "Registros consolidados",
        len(final_df),
    )

    col3.metric(
        "Abas de resumo ignoradas",
        len(ignored_sheets),
    )

    # -------------------------------------------------------------
    # ABAS IGNORADAS
    # -------------------------------------------------------------

    if ignored_sheets:

        st.info(
            "Abas ignoradas: "
            + ", ".join(
                ignored_sheets
            )
        )

    # -------------------------------------------------------------
    # AVISOS
    # -------------------------------------------------------------

    for warning in warnings:

        st.warning(
            warning
        )

    # -------------------------------------------------------------
    # RELATÓRIO
    # -------------------------------------------------------------

    st.subheader(
        "Relatório do processamento"
    )

    report_df = pd.DataFrame(
        report
    )

    st.dataframe(
        report_df,
        use_container_width=True,
        hide_index=True,
    )

    # -------------------------------------------------------------
    # PRÉVIA
    # -------------------------------------------------------------

    st.subheader(
        "Prévia do Resumo Geral"
    )

    st.dataframe(
        final_df.head(100),
        use_container_width=True,
        hide_index=True,
    )

    if len(final_df) > 100:

        st.caption(
            "Exibindo as primeiras 100 linhas "
            f"de {len(final_df)} registros. "
            "O arquivo baixado contém todos "
            "os registros."
        )

    # -------------------------------------------------------------
    # DOWNLOAD
    # -------------------------------------------------------------

    st.download_button(
        label="⬇️ Baixar Resumo Geral.xlsx",
        data=excel_bytes,
        file_name="Resumo_Geral.xlsx",
        mime=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        use_container_width=True,
    )
