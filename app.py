import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Status Pedidos Cliente", layout="wide")
st.title("📋 Atualizador de Status dos Pedidos")

def normalizar_lote(valor):
    try:
        return str(int(float(str(valor).strip())))
    except:
        return str(valor).strip().lstrip("0")

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

col1,col2 = st.columns(2)
with col1:
    prog_f10 = st.file_uploader("Programação F10", type=["xlsx","xls"])
    prog_f8 = st.file_uploader("Programação Filial 8", type=["xlsx","xls"])
with col2:
    prontas_f10 = st.file_uploader("Prontas F10", type=["xlsx","xls"])
    prontas_f8 = st.file_uploader("Prontas F8", type=["xlsx","xls"])

cliente = st.file_uploader("Pedidos do Cliente", type=["xlsx","xls"])

if st.button("Processar") and all([prog_f10, prog_f8, prontas_f10, prontas_f8, cliente]):
    mapas = []
    arquivos = [
        (prog_f10, "Programado LCL10"),
        (prog_f8, "Programado F8"),
        (prontas_f10, "Pronto na F10"),
        (prontas_f8, "Pronto na F8")
    ]

    for arq, status in arquivos:
        df = carregar_planilha(arq)
        possiveis_pedido = [c for c in df.columns if "pedido" in c.lower()]
        possiveis_lote = [c for c in df.columns if "lote" in c.lower()]
        if possiveis_pedido and possiveis_lote:
            aux = df[[possiveis_pedido[0], possiveis_lote[0]]].copy()
            aux.columns = ["Pedido","Lote"]
            aux["Pedido"] = aux["Pedido"].astype(str).str.strip()
            aux["Lote"] = aux["Lote"].apply(normalizar_lote)
            aux["Status"] = status
            mapas.append(aux)

    base = pd.concat(mapas, ignore_index=True)

    cli = pd.read_excel(cliente)
    lote_col = [c for c in cli.columns if "lote" in c.lower()][0]
    pedido_col = [c for c in cli.columns if "pedido" in c.lower()][0]

    cli["_Pedido"] = cli[pedido_col].astype(str).str.strip()
    cli["_Lote"] = cli[lote_col].apply(normalizar_lote)

    resultados = []
    for _, row in cli.iterrows():
        filtro = base[(base["Pedido"] == row["_Pedido"]) &
                       (base["Lote"] == row["_Lote"])]
        resultados.append(" | ".join(sorted(filtro["Status"].unique()))
                          if not filtro.empty else "Não localizado")

    cli["Status Atual"] = resultados
    cli.drop(columns=["_Pedido","_Lote"], inplace=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cli.to_excel(writer, index=False, sheet_name="Resultado")

    st.success("Processamento concluído.")
    st.download_button("📥 Baixar Excel", output.getvalue(),
                       file_name="Pedidos_Atualizados.xlsx")
