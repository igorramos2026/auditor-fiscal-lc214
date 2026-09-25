import io
import os
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
ARQUIVO_ANEXOS_FIXO = "anexos_lc214.xlsx"

def formatar_cclastrib(val_raw):
    """Garante que o CCLASTRIB tenha sempre 6 dígitos com zeros à esquerda (ex: '000001')."""
    if pd.isna(val_raw) or val_raw is None:
        return ""
    s = str(val_raw).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    if digits:
        return digits.zfill(6)
    return s

def formatar_ncm(ncm_raw):
    """Garante que a NCM tenha 8 dígitos preenchidos com zeros à esquerda."""
    if pd.isna(ncm_raw) or ncm_raw is None:
        return "", ""
    s = str(ncm_raw).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    if not digits:
        return str(ncm_raw), ""
    digits = digits.zfill(8)
    if len(digits) > 8:
        digits = digits[:8]
    return f"{digits[:4]}.{digits[4:6]}.{digits[6:]}", digits

@st.cache_data(show_spinner=False)
def extrair_base_anexos(file_anexos):
    """Lê a base de anexos em segundo plano e armazena em cache."""
    xls = pd.ExcelFile(file_anexos)
    mapa_anexos = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
        if df.empty:
            continue
        cclastrib_raw = str(df.iloc[0, 0])
        match = re.search(r'\d{6}', cclastrib_raw)
        if not match:
            continue
        cclastrib_aba = match.group(0)
        
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
                return row[col]
    return ""

def auditar_linha(row, mapa_anexos):
    filial = str(buscar_campo(row, ['FILIAL', 'L_FILIAL', 'COD_FILIAL'])).strip()
    if filial.endswith('.0'):
        filial = filial[:-2]
        
    registro = str(buscar_campo(row, ['REGISTRO', 'L_REGISTRO_NOTA', 'REGISTRO_NOTA', 'NOTA'])).strip()
    if registro.endswith('.0'):
        registro = registro[:-2]

    desc_produto = str(buscar_campo(row, ['DESCRIÇÃO PRODUTO', 'DESCRICAO PRODUTO', 'X_DESC_PRODUTO', 'PRODUTO', 'DESCRICAO'])).strip()
    
    ncm_raw = buscar_campo(row, ['NCM', 'X_NCM', 'NCM_PRODUTO'])
    cclastrib_raw = buscar_campo(row, ['CCLASTRIB UTILIZADO', 'CCLASTRIB', 'X_CLASSTRIB', 'CLASSTRIB', 'CCLASTRIB_UTILIZADO'])

    ncm_fmt, ncm_digits = formatar_ncm(ncm_raw)
    cclastrib_utilizado = formatar_cclastrib(cclastrib_raw)
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

st.sidebar.header("📁 Envio de Ficheiro")

# Campo ÚNICO visível para o utilizador
file_vendas = st.sidebar.file_uploader("Selecione a Planilha de Vendas (.xlsx)", type=["xlsx"])

if file_vendas:
    if st.button("🚀 Iniciar Auditoria", type="primary"):
        if not os.path.exists(ARQUIVO_ANEXOS_FIXO):
            st.error("Erro no sistema: Base de dados da legislação não encontrada no servidor.")
        else:
            with st.spinner("A auditar as vendas..."):
                mapa_anexos = extrair_base_anexos(ARQUIVO_ANEXOS_FIXO)
                df_vendas = pd.read_excel(file_vendas, dtype=str)
                resultados = [auditar_linha(row, mapa_anexos) for _, row in df_vendas.iterrows()]
                df_resultado = pd.DataFrame(resultados)

                df_erro_ncm = df_resultado[df_resultado["STATUS"] == "ERRO_NCM"].drop(columns=["STATUS"], errors="ignore")
                df_erro_cclastrib = df_resultado[df_resultado["STATUS"] == "ERRO_CCLASTRIB"].drop(columns=["STATUS"], errors="ignore")

            st.success("Auditoria concluída com sucesso!")

            # Resumo Estatístico
            total_registos = len(df_resultado)
            qtd_erro_ncm = len(df_erro_ncm)
            qtd_erro_cclastrib = len(df_erro_cclastrib)
            qtd_ok = total_registos - (qtd_erro_ncm + qtd_erro_cclastrib)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total de Registos", total_registos)
            col2.metric("Registos OK", qtd_ok)
            col3.metric("Erros de NCM", qtd_erro_ncm)
            col4.metric("Erros de CCLASTRIB", qtd_erro_cclastrib)

            st.subheader("📋 Inconsistências Encontradas")

            if qtd_erro_ncm == 0 and qtd_erro_cclastrib == 0:
                st.balloons()
                st.success("🎉 Nenhuma inconsistência encontrada! Todos os registos estão em conformidade.")
            else:
                tab_ncm, tab_cclastrib = st.tabs([
                    f"⚠️ Inconsistências NCM ({qtd_erro_ncm})", 
                    f"⚠️ Inconsistências CCLASTRIB ({qtd_erro_cclastrib})"
                ])

                with tab_ncm:
                    if qtd_erro_ncm > 0:
                        st.dataframe(df_erro_ncm, use_container_width=True)
                    else:
                        st.info("Nenhuma inconsistência de NCM encontrada.")

                with tab_cclastrib:
                    if qtd_erro_cclastrib > 0:
                        st.dataframe(df_erro_cclastrib, use_container_width=True)
                    else:
                        st.info("Nenhuma inconsistência de CCLASTRIB encontrada.")

                # Gerar Excel com 2 abas separadas
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_erro_ncm.to_excel(writer, sheet_name="Inconsistências NCM", index=False)
                    df_erro_cclastrib.to_excel(writer, sheet_name="Inconsistências CCLASTRIB", index=False)
                excel_bytes = buffer.getvalue()

                st.download_button(
                    label="📥 Descarregar Planilha de Inconsistências (Excel)",
                    data=excel_bytes,
                    file_name="inconsistencias_auditadas_lc214.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
else:
    st.info("👈 Envie a sua planilha de vendas no menu lateral para começar.")
