import io
import re
import pandas as pd
import streamlit as st

# Configuração inicial da página
st.set_page_config(
    page_title="Auditor Fiscal LC 214/25",
    page_icon="📊",
    layout="wide"
)

CCLASTRIB_INTEGRAL = "000001"

def formatar_ncm(ncm_raw):
    digits = re.sub(r'\D', '', str(ncm_raw)).zfill(8)
    if len(digits) == 8:
        return f"{digits[:4]}.{digits[4:6]}.{digits[6:]}", digits
    return str(ncm_raw), digits

def extrair_base_anexos(file_anexos):
    xls = pd.ExcelFile(file_anexos)
    mapa_anexos = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name)
        if df.empty:
            continue
        cclastrib_raw = str(df.iloc[0, 0])
        match = re.search(r'\d{6}', cclastrib_raw)
        if not match:
            continue
        cclastrib_aba = match.group(0)
        
        # Conversão segura de todas as células para texto
        texto_completo = " ".join([str(x) for x in df.values.flatten() if pd.notna(x)])
        
        matches = re.findall(r'\b\d{2,4}(?:\.\d{1,2})*(?:\.\d{1,2})*\b', texto_completo)
        for m in matches:
            ncm_clean = re.sub(r'\D', '', m).zfill(8)
            if len(ncm_clean) == 8:
                mapa_anexos[ncm_clean] = cclastrib_aba
    return mapa_anexos

def buscar_campo(row, nomes_possiveis):
    for nome in nomes_possiveis:
        for col in row.index:
            if str(col).strip().upper() == nome.upper() and pd.notna(row[col]):
                return str(row[col]).strip()
    return ""

def auditar_linha(row, mapa_anexos):
    filial = buscar_campo(row, ['FILIAL', 'L_FILIAL', 'COD_FILIAL'])
    registro = buscar_campo(row, ['REGISTRO', 'L_REGISTRO_NOTA', 'REGISTRO_NOTA', 'NOTA'])
    desc_produto = buscar_campo(row, ['DESCRIÇÃO PRODUTO', 'DESCRICAO PRODUTO', 'X_DESC_PRODUTO', 'PRODUTO', 'DESCRICAO'])
    ncm_raw = buscar_campo(row, ['NCM', 'X_NCM', 'NCM_PRODUTO'])
    cclastrib_utilizado = buscar_campo(row, ['CCLASTRIB UTILIZADO', 'CCLASTRIB', 'X_CLASSTRIB', 'CLASSTRIB'])

    ncm_fmt, ncm_digits = formatar_ncm(ncm_raw)
    desc_upper = desc_produto.upper()

    # ETAPA 1: COERÊNCIA DA NCM
    ncm_correta_sugerida = None
    if ncm_digits == "02101200":
        termos_bovinos = ["CARNE", "CONTRA FILE", "MAMINHA", "BISTECA", "PRIME RIBE", "FRALDINHA", "ALCATRA", "PICANHA", "ASSADO DE TIRAS", "RAKETE"]
        if any(term in desc_upper for term in termos_bovinos) and "BACON" not in desc_upper:
            ncm_correta_sugerida = "0201.30.00 (sem osso) ou 0201.20.90 (com osso)"
    elif ncm_digits == "21011200":
        termos_cafe = ["CAFE", "CAFÉ", "EXPRESSO", "CAPPUCCINO", "CAPUCCINO", "LATTE"]
        if any(term in desc_upper for term in ["ACHOCOLATADO", "KIT KAT", "ALPINO", "CHOCOLATE EN PO"]) and not any(term in desc_upper for term in termos_cafe):
            ncm_correta_sugerida = "1806.90.00"
        elif any(term in desc_upper for term in ["MILKSHAKE", "MILK SHAKE", "MULKSHAKE"]):
            ncm_correta_sugerida = "2202.99.00"
    elif ncm_digits == "19011090":
        if any(term in desc_upper for term in ["NUCITA", "DOCE", "CREME DE AVELA", "NUTELLA"]):
            ncm_correta_sugerida = "1806.90.00"
    elif ncm_digits == "21069090":
        if any(term in desc_upper for term in ["BALA", "PASTILHA", "DROPS", "CHICLETE"]):
            ncm_correta_sugerida = "1704.90.20 / 1704.90.90"

    if ncm_correta_sugerida:
        alerta = f"A Filial {filial}, do registro {registro}, do produto '{desc_produto}' está com a NCM {ncm_fmt} incorreta e a correta seria {ncm_correta_sugerida}."
        return {
            "FILIAL": filial, "REGISTRO": registro, "PRODUTO": desc_produto,
            "STATUS": "ERRO_NCM", "ALERTA_RETORNO": alerta,
            "NCM_INFORMADA": ncm_fmt, "NCM_SUGERIDA": ncm_correta_sugerida,
            "CCLASTRIB_UTILIZADO": cclastrib_utilizado, "CCLASTRIB_CORRETO": CCLASTRIB_INTEGRAL
        }

    # ETAPA 2: VALIDAR CCLASTRIB NOS ANEXOS DA LC 214/25
    cclastrib_correto = CCLASTRIB_INTEGRAL
    if ncm_digits in ["19012090", "19059090"]:
        if any(term in desc_upper for term in ["PAO FRANCES", "PÃO FRANCÊS", "PAO FRANC"]):
            cclastrib_correto = "200003"
    elif ncm_digits == "21011200":
        if any(term in desc_upper for term in ["EXPRESSO", "CURTO", "DUPLO", "CAPUCCINO", "CAPPUCCINO", "LATTE", "FRAPUCCINO", "MOKACCINO", "CAFE"]):
            cclastrib_correto = "200003"
    elif ncm_digits in ["02071413", "02071422", "04061010", "09012100", "09030010", "09030090", "15171000", "19021900", "25010020", "25010090"]:
        cclastrib_correto = "200003"
    elif ncm_digits == "02101200" and "BACON" in desc_upper:
        cclastrib_correto = "200003"
    else:
        if ncm_digits in mapa_anexos:
            cclastrib_correto = mapa_anexos[ncm_digits]
        elif ncm_digits[:6] in mapa_anexos:
            cclastrib_correto = mapa_anexos[ncm_digits[:6]]
        elif ncm_digits[:4] in ["0201", "0202", "0203", "0204", "0207", "0209", "0302", "0303", "0304", "0713", "0901", "0903", "1101", "1102", "1103", "1104", "1106", "1701", "1902", "2501"]:
            cclastrib_correto = "200003"

    if cclastrib_utilizado == cclastrib_correto:
        alerta = f"OK — Filial {filial}, Registro {registro}, Produto '{desc_produto}': NCM {ncm_fmt} e CCLASTRIB {cclastrib_utilizado} corretos."
        status = "OK"
    else:
        alerta = f"A Filial {filial}, do registro {registro}, do produto '{desc_produto}', da NCM {ncm_fmt} utilizou o CCLASTRIB {cclastrib_utilizado} incorreto sendo o correto {cclastrib_correto}."
        status = "ERRO_CCLASTRIB"

    return {
        "FILIAL": filial, "REGISTRO": registro, "PRODUTO": desc_produto,
        "STATUS": status, "ALERTA_RETORNO": alerta,
        "NCM_INFORMADA": ncm_fmt, "NCM_SUGERIDA": ncm_fmt,
        "CCLASTRIB_UTILIZADO": cclastrib_utilizado, "CCLASTRIB_CORRETO": cclastrib_correto
    }

# INTERFACE GRÁFICA STREAMLIT
st.title("📊 Auditor Fiscal LC 214/25")
st.markdown("Valide automaticamente a NCM e o CCLASTRIB das suas planilhas de vendas.")

st.sidebar.header("📁 Envio de Ficheiros")
file_anexos = st.sidebar.file_uploader("1. Ficheiro de Anexos LC 214/25 (.xlsx)", type=["xlsx"])
file_vendas = st.sidebar.file_uploader("2. Planilha de Vendas (.xlsx)", type=["xlsx"])

if file_anexos and file_vendas:
    if st.button("🚀 Iniciar Auditoria", type="primary"):
        with st.spinner("A processar a tabela de Anexos e a auditar as vendas..."):
            mapa_anexos = extrair_base_anexos(file_anexos)
            df_vendas = pd.read_excel(file_vendas)
            resultados = [auditar_linha(row, mapa_anexos) for _, row in df_vendas.iterrows()]
            df_resultado = pd.DataFrame(resultados)

        st.success("Auditoria concluída com sucesso!")

        # Resumo Estatístico
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Registos", len(df_resultado))
        col2.metric("Registos OK", len(df_resultado[df_resultado["STATUS"] == "OK"]))
        col3.metric("Com Inconsistência", len(df_resultado[df_resultado["STATUS"] != "OK"]))

        # Visualização da Tabela
        st.subheader("📋 Resultado da Auditoria")
        st.dataframe(df_resultado, use_container_width=True)

        # Download para Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_resultado.to_excel(writer, index=False)
        excel_bytes = buffer.getvalue()

        st.download_button(
            label="📥 Descarregar Planilha Auditada em Excel",
            data=excel_bytes,
            file_name="vendas_auditadas_lc214.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
else:
    st.info("👈 Envie os dois ficheiros (.xlsx) no menu lateral para começar.")
