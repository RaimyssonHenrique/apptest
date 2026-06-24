import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Status Pedidos Cliente", layout="wide")
st.title("📋 Atualizador de Status dos Pedidos")


# ==================================================
# FUNÇÕES
# ==================================================
@st.cache_data
def carregar_planilha_bytes(file_bytes):
    """
    Lê todas as abas de um arquivo Excel enviado pelo Streamlit.
    Compatível com Streamlit Cloud.
    """

    frames = []

    try:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Erro ao abrir arquivo: {e}")
        return pd.DataFrame()

    for aba in xls.sheet_names:
        try:
            df = pd.read_excel(
                BytesIO(file_bytes),
                sheet_name=aba,
                engine="openpyxl"
            )

            df.columns = [str(c).strip() for c in df.columns]
            frames.append(df)

        except Exception:
            continue

    if frames:
        return pd.concat(frames, ignore_index=True)

    return pd.DataFrame()


def normalizar_lote(valor):
    try:
        return str(int(float(str(valor).strip())))
    except:
        return str(valor).strip().lstrip("0")


# ==================================================
# FORMULÁRIO
# ==================================================
with st.form("form_processamento"):

    col1, col2 = st.columns(2)

    with col1:
        prog_f10 = st.file_uploader(
            "Programação F10",
            type=["xlsx", "xls"]
        )

        prog_f8 = st.file_uploader(
            "Programação Filial 8",
            type=["xlsx", "xls"]
        )

    with col2:
        prontas_f10 = st.file_uploader(
            "Prontas F10",
            type=["xlsx", "xls"]
        )

        prontas_f8 = st.file_uploader(
            "Prontas F8",
            type=["xlsx", "xls"]
        )

    cliente = st.file_uploader(
        "Pedidos do Cliente",
        type=["xlsx", "xls"]
    )

    submit = st.form_submit_button("🚀 Processar")


# ==================================================
# PROCESSAMENTO
# ==================================================
if submit:

    arquivos_obrigatorios = [
        prog_f10,
        prog_f8,
        prontas_f10,
        prontas_f8,
        cliente
    ]

    if not all(arquivos_obrigatorios):
        st.error("Envie todos os arquivos.")
        st.stop()

    with st.spinner("Processando arquivos..."):

        arquivos = [
            (prog_f10, "Programado LCL10"),
            (prog_f8, "Programado F8"),
            (prontas_f10, "Pronto na F10"),
            (prontas_f8, "Pronto na F8")
        ]

        mapas = []

        # ==========================================
        # MONTA BASE DE CONSULTA
        # ==========================================
        for arq, status in arquivos:

            if arq is None:
                continue

            try:
                file_bytes = arq.getvalue()
                df = carregar_planilha_bytes(file_bytes)

                if df.empty:
                    st.warning(
                        f"O arquivo '{arq.name}' não possui dados válidos."
                    )
                    continue

                pedido_cols = [
                    c for c in df.columns
                    if "pedido" in c.lower()
                ]

                lote_cols = [
                    c for c in df.columns
                    if "lote" in c.lower()
                ]

                if not pedido_cols or not lote_cols:
                    st.warning(
                        f"O arquivo '{arq.name}' não possui colunas Pedido/Lote."
                    )
                    continue

                aux = df[
                    [pedido_cols[0], lote_cols[0]]
                ].copy()

                aux.columns = ["Pedido", "Lote"]

                aux["Pedido"] = (
                    aux["Pedido"]
                    .astype(str)
                    .str.strip()
                )

                aux["Lote"] = (
                    aux["Lote"]
                    .apply(normalizar_lote)
                )

                aux["Status"] = status

                mapas.append(aux)

            except Exception as e:
                st.error(f"Erro no arquivo {arq.name}: {e}")

        if not mapas:
            st.error("Nenhum dado válido foi encontrado.")
            st.stop()

        base = pd.concat(mapas, ignore_index=True)

        # Remove duplicidades
        base = base.drop_duplicates()

        # ==========================================
        # CLIENTE
        # ==========================================
        try:

            cli = pd.read_excel(
                BytesIO(cliente.getvalue()),
                engine="openpyxl"
            )

        except Exception as e:
            st.error(f"Erro ao abrir arquivo do cliente: {e}")
            st.stop()

        pedido_cols = [
            c for c in cli.columns
            if "pedido" in c.lower()
        ]

        lote_cols = [
            c for c in cli.columns
            if "lote" in c.lower()
        ]

        if not pedido_cols or not lote_cols:
            st.error(
                "O arquivo do cliente não possui as colunas Pedido e Lote."
            )
            st.stop()

        pedido_col = pedido_cols[0]
        lote_col = lote_cols[0]

        cli["_Pedido"] = (
            cli[pedido_col]
            .astype(str)
            .str.strip()
        )

        cli["_Lote"] = (
            cli[lote_col]
            .apply(normalizar_lote)
        )

        # ==========================================
        # PROCESSAMENTO VETORIZADO
        # ==========================================
        merged = cli.merge(
            base,
            left_on=["_Pedido", "_Lote"],
            right_on=["Pedido", "Lote"],
            how="left"
        )

        status_agg = (
            merged.groupby(merged.index)["Status"]
            .apply(
                lambda x:
                " | ".join(sorted(set(x.dropna())))
                if x.notna().any()
                else "Não localizado"
            )
        )

        cli["Status Atual"] = status_agg.values

        cli.drop(
            columns=["_Pedido", "_Lote"],
            inplace=True
        )

        # ==========================================
        # EXPORTAÇÃO
        # ==========================================
        output = BytesIO()

        with pd.ExcelWriter(
            output,
            engine="openpyxl"
        ) as writer:

            cli.to_excel(
                writer,
                index=False,
                sheet_name="Resultado"
            )

        output.seek(0)

        st.success("✅ Processamento concluído com sucesso!")

        st.download_button(
            "📥 Baixar Excel",
            data=output,
            file_name="Pedidos_Atualizados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
