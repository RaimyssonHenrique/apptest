import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Status Pedidos Cliente", layout="wide")
st.title("📋 Atualizador de Status dos Pedidos")


# =========================
# CACHE (ACELERA LEITURA)
# =========================
@st.cache_data
def carregar_planilha(arquivo):
    xls = pd.ExcelFile(arquivo)
    frames = []

    for aba in xls.sheet_names:
        try:
            df = pd.read_excel(arquivo, sheet_name=aba)
            df.columns = [str(c).strip() for c in df.columns]
            frames.append(df)
        except:
            pass

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def normalizar_lote(valor):
    try:
        return str(int(float(str(valor).strip())))
    except:
        return str(valor).strip().lstrip("0")


# =========================
# FORM (CORRIGE ERRO FRONTEND)
# =========================
with st.form("form_processamento"):
    col1, col2 = st.columns(2)

    with col1:
        prog_f10 = st.file_uploader("Programação F10", type=["xlsx","xls"])
        prog_f8 = st.file_uploader("Programação Filial 8", type=["xlsx","xls"])

    with col2:
        prontas_f10 = st.file_uploader("Prontas F10", type=["xlsx","xls"])
        prontas_f8 = st.file_uploader("Prontas F8", type=["xlsx","xls"])

    cliente = st.file_uploader("Pedidos do Cliente", type=["xlsx","xls"])

    submit = st.form_submit_button("🚀 Processar")


# =========================
# PROCESSAMENTO
# =========================
if submit:

    if not all([prog_f10, prog_f8, prontas_f10, prontas_f8, cliente]):
        st.error("Envie todos os arquivos.")
        st.stop()

    arquivos = [
        (prog_f10, "Programado LCL10"),
        (prog_f8, "Programado F8"),
        (prontas_f10, "Pronto na F10"),
        (prontas_f8, "Pronto na F8")
    ]

    mapas = []

    for arq, status in arquivos:
        df = carregar_planilha(arq)

        pedido_cols = [c for c in df.columns if "pedido" in c.lower()]
        lote_cols = [c for c in df.columns if "lote" in c.lower()]

        if not pedido_cols or not lote_cols:
            continue

        aux = df[[pedido_cols[0], lote_cols[0]]].copy()
        aux.columns = ["Pedido", "Lote"]

        aux["Pedido"] = aux["Pedido"].astype(str).str.strip()
        aux["Lote"] = aux["Lote"].apply(normalizar_lote)
        aux["Status"] = status

        mapas.append(aux)

    if not mapas:
        st.error("Nenhum dado válido encontrado nos arquivos.")
        st.stop()

    base = pd.concat(mapas, ignore_index=True)


    # =========================
    # CLIENTE
    # =========================
    cli = pd.read_excel(cliente)

    lote_col = [c for c in cli.columns if "lote" in c.lower()][0]
    pedido_col = [c for c in cli.columns if "pedido" in c.lower()][0]

    cli["_Pedido"] = cli[pedido_col].astype(str).str.strip()
    cli["_Lote"] = cli[lote_col].apply(normalizar_lote)


    # =========================
    # 🔥 VETORIZADO (RÁPIDO)
    # =========================
    merged = cli.merge(
        base,
        left_on=["_Pedido", "_Lote"],
        right_on=["Pedido", "Lote"],
        how="left"
    )

    status_agg = (
        merged.groupby(merged.index)["Status"]
        .apply(lambda x: " | ".join(sorted(set(x.dropna()))) if x.notna().any() else "Não localizado")
    )

    cli["Status Atual"] = status_agg.values

    cli.drop(columns=["_Pedido", "_Lote"], inplace=True)


    # =========================
    # EXPORTAÇÃO
    # =========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cli.to_excel(writer, index=False, sheet_name="Resultado")

    output.seek(0)

    st.success("Processamento concluído com sucesso 🚀")

    st.download_button(
        "📥 Baixar Excel",
        output.getvalue(),
        file_name="Pedidos_Atualizados.xlsx"
    )
