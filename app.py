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

# CFOPs mapeados para regras especiais de CCLASTRIB
CFOPS_TRANSFERENCIA = ["5151", "5152", "5153", "5155", "5156", "6151", "6152", "6153", "6155", "6156", "7151", "7152"]
CFOPS_CONSERTO = ["5915", "6915", "7915"]

def formatar_cfop(cfop_raw):
    """Garante que o CFOP tenha 4 dígitos numéricos."""
    if pd.isna(cfop_raw) or cfop_raw is None:
        return ""
    s = str(cfop_raw).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    return digits.zfill(4) if digits else s

def formatar_cclastrib(val_raw):
    """Garante que o CCLASTRIB tenha sempre 6 dígitos com zeros à esquerda."""
    if pd.isna(val_raw) or val_raw is None:
        return ""
    s = str(val_raw).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    return digits.zfill(6) if digits else s

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

def auditar_linha_otimizada(row, cols_map, mapa_anexos):
    col_filial, col_reg, col_prod, col_ncm, col_cclastrib, col_cfop = cols_map

    filial = str(row[col_filial]).strip() if col_filial and pd.notna(row[col_filial]) else ""
    if filial.endswith('.0'): filial = filial[:-2]
        
    registro = str(row[col_reg]).strip() if col_reg and pd.notna(row[col_reg]) else ""
    if registro.endswith('.0'): registro = registro[:-2]

    desc_produto = str(row[col_prod]).strip() if col_prod and pd.notna(row[col_prod]) else ""
    ncm_raw = row[col_ncm] if col_ncm and pd.notna(row[col_ncm]) else ""
    cclastrib_raw = row[col_cclastrib] if col_cclastrib and pd.notna(row[col_cclastrib]) else ""
    cfop_raw = row[col_cfop] if col_cfop and pd.notna(row[col_cfop]) else ""

    ncm_fmt, ncm_digits = formatar_ncm(ncm_raw)
    cclastrib_utilizado = formatar_cclastrib(cclastrib_raw)
    cfop = formatar_cfop(cfop_raw)
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
            "FILIAL": filial, "REGISTRO": registro, "PRODUTO": desc_produto, "CFOP": cfop,
            "STATUS": "ERRO_NCM", "ALERTA_RETORNO": alerta,
            "NCM_INFORMADA": ncm_fmt, "NCM_SUGERIDA": ncm_correta_sugerida,
            "CCLASTRIB_UTILIZADO": cclastrib_utilizado, "CCLASTRIB_CORRETO": CCLASTRIB_INTEGRAL
        }

    # ETAPA 2: VALIDAR CCLASTRIB (Prioridade por CFOP)
    cclastrib_correto = None

    # Regras prioritárias por CFOP
    if cfop in CFOPS_TRANSFERENCIA:
        cclastrib_correto = "410002"
    elif cfop in CFOPS_CONSERTO:
        cclastrib_correto = "410999"
    else:
        # Regras gerais por NCM / Descrição
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

        if cclastrib_correto is None:
            if ncm_digits in mapa_anexos:
                cclastrib_correto = mapa_anexos[ncm_digits]
            elif ncm_digits[:6] in mapa_anexos:
                cclastrib_correto = mapa_anexos[ncm_digits[:6]]
            elif ncm_digits[:4] in ["0201", "0202", "0203", "0204", "0207", "0209", "0302", "0303", "0304", "0713", "0901", "0903", "1101", "1102", "1103", "1104", "1106", "1701", "1902", "2501"]:
                cclastrib_correto = "200003"
            else:
                cclastrib_correto = CCLASTRIB_INTEGRAL

    if cclastrib_utilizado == cclastrib_correto:
        alerta = f"OK — Filial {filial}, Registro {registro}, Produto '{desc_produto}': NCM {ncm_fmt}, CFOP {cfop} e CCLASTRIB {cclastrib_utilizado} corretos."
        status = "OK"
    else:
        alerta = f"A Filial {filial}, do registro {registro}, do produto '{desc_produto}' (CFOP {cfop}), da NCM {ncm_fmt} utilizou o CCLASTRIB {cclastrib_utilizado} incorreto sendo o correto {cclastrib_correto}."
        status = "ERRO_CCLASTRIB"

    return {
        "FILIAL": filial, "REGISTRO": registro, "PRODUTO": desc_produto, "CFOP": cfop,
        "STATUS": status, "ALERTA_RETORNO": alerta,
        "NCM_INFORMADA": ncm_fmt, "NCM_SUGERIDA": ncm_fmt,
        "CCLASTRIB_UTILIZADO": cclastrib_utilizado, "CCLASTRIB_CORRETO": cclastrib_correto
    }

# INTERFACE GRÁFICA STREAMLIT
st.title("📊 Auditor Fiscal LC 214/25")
st.markdown("Valide automaticamente NCM, CFOP e CCLASTRIB das suas planilhas de vendas.")

st.sidebar.header("📁 Envio de Ficheiro")
file_vendas = st.sidebar.file_uploader("Selecione a Planilha de Vendas (.xlsx)", type=["xlsx"])

if file_vendas:
    if st.button("🚀 Iniciar Auditoria", type="primary"):
        if not os.path.exists(ARQUIVO_ANEXOS_FIXO):
            st.error("Erro no sistema: Base de dados da legislação não encontrada no servidor.")
        else:
            with st.spinner("A processar e auditar as vendas..."):
                mapa_anexos = extrair_base_anexos(ARQUIVO_ANEXOS_FIXO)
                df_vendas = pd.read_excel(file_vendas, dtype=str)

                # Mapeamento rápido de colunas
                col_map_dict = {str(c).strip().upper(): c for c in df_vendas.columns}
                
                def encontrar_coluna(opcoes):
                    for opt in opcoes:
                        if opt.upper() in col_map_dict:
                            return col_map_dict[opt.upper()]
                    return None

                col_filial = encontrar_coluna(['FILIAL', 'L_FILIAL', 'COD_FILIAL'])
                col_reg = encontrar_coluna(['REGISTRO', 'L_REGISTRO_NOTA', 'REGISTRO_NOTA', 'NOTA'])
                col_prod = encontrar_coluna(['DESCRIÇÃO PRODUTO', 'DESCRICAO PRODUTO', 'X_DESC_PRODUTO', 'PRODUTO', 'DESCRICAO'])
                col_ncm = encontrar_coluna(['NCM', 'X_NCM', 'NCM_PRODUTO'])
                col_cclastrib = encontrar_coluna(['CCLASTRIB UTILIZADO', 'CCLASTRIB', 'X_CLASSTRIB', 'CLASSTRIB', 'CCLASTRIB_UTILIZADO'])
                col_cfop = encontrar_coluna(['CFOP', 'X_CFOP', 'COD_CFOP', 'CFOP_PRODUTO', 'NOP_CFOP'])

                # Validação de Colunas Obrigatórias
                faltantes = []
                if not col_prod: faltantes.append("PRODUTO / DESCRIÇÃO")
                if not col_ncm: faltantes.append("NCM")
                if not col_cfop: faltantes.append("CFOP")
                if not col_cclastrib: faltantes.append("CCLASTRIB")

                if faltantes:
                    st.error(f"⚠️ **Coluna(s) obrigatória(s) não encontrada(s):** {', '.join(faltantes)}")
                    st.warning(f"**Colunas identificadas na planilha enviada:**\n`{list(df_vendas.columns)}`")
                    st.stop()

                cols_map = (col_filial, col_reg, col_prod, col_ncm, col_cclastrib, col_cfop)

                # Processamento Otimizado
                resultados = [auditar_linha_otimizada(row, cols_map, mapa_anexos) for _, row in df_vendas.iterrows()]
                df_resultado = pd.DataFrame(resultados)

                df_erro_ncm_todos = df_resultado[df_resultado["STATUS"] == "ERRO_NCM"].copy()
                df_erro_cclastrib_todos = df_resultado[df_resultado["STATUS"] == "ERRO_CCLASTRIB"].copy()

                # Métrica de Ocorrências Afectadas por (Produto + CFOP)
                counts_ncm = df_erro_ncm_todos.groupby(["PRODUTO", "CFOP"]).size().to_dict() if not df_erro_ncm_todos.empty else {}
                counts_cclastrib = df_erro_cclastrib_todos.groupby(["PRODUTO", "CFOP"]).size().to_dict() if not df_erro_cclastrib_todos.empty else {}

                # Deduplicação por PRODUTO e CFOP para manter lista limpa
                df_erro_ncm = df_erro_ncm_todos.drop_duplicates(subset=["PRODUTO", "CFOP"]).reset_index(drop=True)
                if not df_erro_ncm.empty:
                    df_erro_ncm["NOTAS_AFETADAS"] = df_erro_ncm.apply(lambda r: counts_ncm.get((r["PRODUTO"], r["CFOP"]), 1), axis=1)

                df_erro_cclastrib = df_erro_cclastrib_todos.drop_duplicates(subset=["PRODUTO", "CFOP"]).reset_index(drop=True)
                if not df_erro_cclastrib.empty:
                    df_erro_cclastrib["NOTAS_AFETADAS"] = df_erro_cclastrib.apply(lambda r: counts_cclastrib.get((r["PRODUTO"], r["CFOP"]), 1), axis=1)

                # Organização das Colunas
                cols_ordem = ["PRODUTO", "CFOP", "NOTAS_AFETADAS", "NCM_INFORMADA", "NCM_SUGERIDA", "CCLASTRIB_UTILIZADO", "CCLASTRIB_CORRETO", "FILIAL", "REGISTRO", "ALERTA_RETORNO"]
                
                df_erro_ncm_view = df_erro_ncm[[c for c in cols_ordem if c in df_erro_ncm.columns]]
                df_erro_cclastrib_view = df_erro_cclastrib[[c for c in cols_ordem if c in df_erro_cclastrib.columns]]

            st.success("Auditoria concluída com sucesso!")

            # Resumo Estatístico
            total_linhas = len(df_resultado)
            qtd_erro_ncm = len(df_erro_ncm_view)
            qtd_erro_cclastrib = len(df_erro_cclastrib_view)
            qtd_ok = total_linhas - (len(df_erro_ncm_todos) + len(df_erro_cclastrib_todos))

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Linhas Analisadas", total_linhas)
            col2.metric("Linhas OK", qtd_ok)
            col3.metric("Erros de NCM", qtd_erro_ncm, help=f"Afeta {len(df_erro_ncm_todos)} ocorrências nas notas")
            col4.metric("Erros de CCLASTRIB", qtd_erro_cclastrib, help=f"Afeta {len(df_erro_cclastrib_todos)} ocorrências nas notas")

            st.subheader("📋 Inconsistências Encontradas (Agrupadas por Produto e CFOP)")

            if qtd_erro_ncm == 0 and qtd_erro_cclastrib == 0:
                st.balloons()
                st.success("🎉 Nenhuma inconsistência encontrada! Todos os registos estão em conformidade.")
            else:
                tab_ncm, tab_cclastrib = st.tabs([
                    f"⚠️ Inconsistências NCM ({qtd_erro_ncm} itens)", 
                    f"⚠️ Inconsistências CCLASTRIB ({qtd_erro_cclastrib} itens)"
                ])

                with tab_ncm:
                    if qtd_erro_ncm > 0:
                        st.dataframe(df_erro_ncm_view, use_container_width=True)
                    else:
                        st.info("Nenhuma inconsistência de NCM encontrada.")

                with tab_cclastrib:
                    if qtd_erro_cclastrib > 0:
                        st.dataframe(df_erro_cclastrib_view, use_container_width=True)
                    else:
                        st.info("Nenhuma inconsistência de CCLASTRIB encontrada.")

                # Excel de Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_erro_ncm_view.to_excel(writer, sheet_name="Inconsistências NCM", index=False)
                    df_erro_cclastrib_view.to_excel(writer, sheet_name="Inconsistências CCLASTRIB", index=False)
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
