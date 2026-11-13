import streamlit as st
import pandas as pd
import openpyxl
import os
import uuid
import sys
import subprocess
import zipfile
from datetime import datetime, date
from io import BytesIO

# Manipulação de Word para a Notificação
try:
    import docx
    DOCX_DISPONIVEL = True
except ImportError:
    DOCX_DISPONIVEL = False

# Importação para conversão em Windows via COM API
try:
    import win32com.client as win32
    import pythoncom
    EXCEL_DISPONIVEL = True
except ImportError:
    EXCEL_DISPONIVEL = False

st.set_page_config(
    page_title="Calculadora CNR Clandestino",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# 🎈 EASTER EGG DA SEXTA-FEIRA
# -----------------------------------------------------------------------------
agora = datetime.now()
dia_semana = agora.weekday()
hora = agora.hour
minuto = agora.minute

if dia_semana == 4 and (hora > 8 or (hora == 8 and minuto >= 30)):
    st.balloons()
    st.warning("🚨 **REGRA DE OURO DA SEXTA:** Sexta-Feira, 08:30 da manhã: **Quem fez, fez!** Quem não fez, só na SEGUNDA! ☕🎉")
# -----------------------------------------------------------------------------

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

def obter_data_extenso_pt(dt=None):
    if dt is None:
        dt = datetime.now()
    nome_mes = MESES_PT.get(dt.month, "")
    return f"{dt.day} de {nome_mes} de {dt.year}"

def fmt_br(valor, decimais=2):
    if pd.isnull(valor):
        return "0,00"
    formato = f"{{:,.{decimais}f}}"
    return formato.format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def valor_para_float_seguro(valor, padrao=0.0):
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return float(padrao)
    try:
        if isinstance(valor, str):
            texto = valor.strip()
            if texto in ('', 'N/A', 'nan', 'NaN', 'None'):
                return float(padrao)
            texto = texto.replace('R$', '').replace('$', '').strip()
            if ',' in texto and '.' in texto:
                texto = texto.replace('.', '').replace(',', '.')
            elif ',' in texto:
                texto = texto.replace(',', '.')
            return float(texto)
        return float(valor)
    except Exception:
        return float(padrao)

def aplicar_layout_a4(ws):
    try:
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        if ws.sheet_properties.pageSetUpPr is None:
            ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.page.PageSetupProperties()
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.print_area = "A1:I41"
    except Exception:
        pass

    try:
        ws.page_margins.left = 0.3
        ws.page_margins.right = 0.3
        ws.page_margins.top = 0.4
        ws.page_margins.bottom = 0.4
        ws.page_margins.header = 0.1
        ws.page_margins.footer = 0.1
    except Exception:
        pass

@st.cache_data
def carregar_arquivo_single(uploaded_file):
    if uploaded_file.name.endswith(('xlsx', 'xls')):
        return pd.read_excel(uploaded_file)
    else:
        encodings = ['latin1', 'iso-8859-1', 'cp1252', 'utf-8']
        for enc in encodings:
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, sep=None, engine='python', encoding=enc)
            except (UnicodeDecodeError, Exception):
                continue
        raise Exception(f"Não foi possível ler {uploaded_file.name}")

def converter_para_pdf_universal(caminho_entrada, caminho_saida):
    if sys.platform.startswith("win") and EXCEL_DISPONIVEL:
        try:
            pythoncom.CoInitialize()
            if caminho_entrada.endswith('.xlsx'):
                excel_app = win32.Dispatch("Excel.Application")
                try: excel_app.Visible = False
                except Exception: pass
                excel_app.DisplayAlerts = False
                
                wb_com = excel_app.Workbooks.Open(caminho_entrada)
                ws_com = wb_com.Worksheets(1)
                ws_com.PageSetup.Zoom = False
                ws_com.PageSetup.FitToPagesWide = 1
                ws_com.PageSetup.FitToPagesTall = 1
                
                wb_com.ExportAsFixedFormat(0, caminho_saida)
                wb_com.Close(False)
                excel_app.Quit()
            elif caminho_entrada.endswith('.docx'):
                word_app = win32.Dispatch("Word.Application")
                try: word_app.Visible = False
                except Exception: pass
                
                doc_com = word_app.Documents.Open(caminho_entrada)
                doc_com.SaveAs(caminho_saida, FileFormat=17)
                doc_com.Close(False)
                word_app.Quit()
            return os.path.exists(caminho_saida)
        except Exception:
            return False
        finally:
            pythoncom.CoUninitialize()
    else:
        try:
            diretorio_saida = os.path.dirname(caminho_saida)
            cmd = [
                "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", diretorio_saida,
                caminho_entrada
            ]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            pdf_gerado_temp = os.path.splitext(caminho_entrada)[0] + ".pdf"
            if os.path.exists(pdf_gerado_temp) and pdf_gerado_temp != caminho_saida:
                os.rename(pdf_gerado_temp, caminho_saida)
                
            return os.path.exists(caminho_saida)
        except Exception:
            return False

def preencher_modelo_excel(caminho_modelo, nome_cliente, uc_val, toi_val, maior_ciclo_row, consumo_diario, dias_cobranca, usou_minima, dt_ini_efetiva, dt_fim_efetiva, tarifa_vigente):
    wb = openpyxl.load_workbook(caminho_modelo)
    ws = wb['DIARIO'] if 'DIARIO' in wb.sheetnames else wb.active
    aplicar_layout_a4(ws)
    
    ws['A6'] = f"Goiânia, {obter_data_extenso_pt()}"
    ws['B8'] = str(nome_cliente)
    ws['B9'] = str(uc_val)
    ws['B10'] = str(toi_val)
    
    dt_ini_str = pd.to_datetime(dt_ini_efetiva).strftime('%d/%m/%Y') if pd.notnull(dt_ini_efetiva) else ''
    dt_fim_str = pd.to_datetime(dt_fim_efetiva).strftime('%d/%m/%Y') if pd.notnull(dt_fim_efetiva) else ''
    
    ws['D15'] = dt_ini_str
    ws['F15'] = dt_fim_str
    ws['H15'] = dias_cobranca
    
    ws['A19'] = "Consumo Base Dia (kWh/dia)"
    ws['I19'] = round(float(valor_para_float_seguro(consumo_diario, 0.0)), 4)
    ws['I20'] = f"=I19*H15"
    ws['I21'] = 0
    ws['I22'] = f"=I20-I21"
    
    dt_hoje = datetime.now()
    tarifa_valor = valor_para_float_seguro(tarifa_vigente, 0.0)
    
    ws['A25'] = f"d) Tarifa atual ref.: {dt_hoje.strftime('%m/%Y')}"
    ws['I25'] = float(tarifa_valor)
    
    criterio_str = "MÍNIMA" if usou_minima else "LIDA"
    mes_ref = int(maior_ciclo_row['MES_REF'].values[0]) if len(maior_ciclo_row) > 0 else 0
    ano_ref = int(maior_ciclo_row['ANO_REF'].values[0]) if len(maior_ciclo_row) > 0 else 0
    kwh_ref = maior_ciclo_row['CONSUMOLIDOKWH'].values[0] if len(maior_ciclo_row) > 0 else 0
    
    ws['A32'] = f"Período da base de cálculo (Maior Pós - {criterio_str}): Ref. {mes_ref:02d}/{ano_ref} com {kwh_ref:,.0f} kWh."
    ws['A35'] = f"Referência Período: {dt_ini_str} a {dt_fim_str} ({dias_cobranca} dias) - Ciclo de Referência {mes_ref:02d}/{ano_ref} - {kwh_ref:,.0f} kWh / 30 = {consumo_diario:,.4f} kWh"
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def gerar_e_converter_pdf_demanda(caminho_modelo, nome_cliente, uc_val, toi_val, maior_ciclo_row, consumo_diario, dias_cobranca, usou_minima, dt_ini_efetiva, dt_fim_efetiva, tarifa_vigente):
    uid = uuid.uuid4().hex[:8]
    temp_excel = os.path.abspath(f"temp_carta_{uc_val}_{uid}.xlsx")
    temp_pdf = os.path.abspath(f"temp_carta_{uc_val}_{uid}.pdf")
    
    wb = openpyxl.load_workbook(caminho_modelo)
    ws = wb['DIARIO'] if 'DIARIO' in wb.sheetnames else wb.active
    aplicar_layout_a4(ws)
    
    ws['A6'] = f"Goiânia, {obter_data_extenso_pt()}"
    ws['B8'] = str(nome_cliente)
    ws['B9'] = str(uc_val)
    ws['B10'] = str(toi_val)
    
    dt_ini_str = pd.to_datetime(dt_ini_efetiva).strftime('%d/%m/%Y') if pd.notnull(dt_ini_efetiva) else ''
    dt_fim_str = pd.to_datetime(dt_fim_efetiva).strftime('%d/%m/%Y') if pd.notnull(dt_fim_efetiva) else ''
    
    ws['D15'] = dt_ini_str
    ws['F15'] = dt_fim_str
    ws['H15'] = dias_cobranca
    
    ws['A19'] = "Consumo Base Dia (kWh/dia)"
    ws['I19'] = round(float(valor_para_float_seguro(consumo_diario, 0.0)), 4)
    ws['I20'] = f"=I19*H15"
    ws['I21'] = 0
    ws['I22'] = f"=I20-I21"
    
    dt_hoje = datetime.now()
    tarifa_valor = valor_para_float_seguro(tarifa_vigente, 0.0)
    
    ws['A25'] = f"d) Tarifa atual ref.: {dt_hoje.strftime('%m/%Y')}"
    ws['I25'] = float(tarifa_valor)
    
    criterio_str = "MÍNIMA" if usou_minima else "LIDA"
    mes_ref = int(maior_ciclo_row['MES_REF'].values[0]) if len(maior_ciclo_row) > 0 else 0
    ano_ref = int(maior_ciclo_row['ANO_REF'].values[0]) if len(maior_ciclo_row) > 0 else 0
    kwh_ref = maior_ciclo_row['CONSUMOLIDOKWH'].values[0] if len(maior_ciclo_row) > 0 else 0
    
    ws['A32'] = f"Período da base de cálculo (Maior Pós - {criterio_str}): Ref. {mes_ref:02d}/{ano_ref} com {kwh_ref:,.0f} kWh."
    ws['A35'] = f"Referência Período: {dt_ini_str} a {dt_fim_str} ({dias_cobranca} dias) - Ciclo de Referência {mes_ref:02d}/{ano_ref} - {kwh_ref:,.0f} kWh / 30 = {consumo_diario:,.4f} kWh"
    
    wb.save(temp_excel)
    
    pdf_bytes = None
    if converter_para_pdf_universal(temp_excel, temp_pdf):
        with open(temp_pdf, "rb") as f:
            pdf_bytes = f.read()
            
    for temp_file in [temp_excel, temp_pdf]:
        if os.path.exists(temp_file):
            try: os.remove(temp_file)
            except Exception: pass
            
    return pdf_bytes

def gerar_notificacao_pdf(caminho_docx, dic_substituicoes, uc_nome):
    uid = uuid.uuid4().hex[:8]
    temp_docx = os.path.abspath(f"temp_notif_{uc_nome}_{uid}.docx")
    temp_pdf = os.path.abspath(f"temp_notif_{uc_nome}_{uid}.pdf")
    
    doc = docx.Document(caminho_docx)
    
    def aplicar_substituicoes_em_paragrafo(p):
        texto_original = p.text
        texto_modificado = texto_original
        for k, v in dic_substituicoes.items():
            if k in texto_modificado:
                texto_modificado = texto_modificado.replace(k, str(v))
        if texto_modificado != texto_original:
            p.text = texto_modificado

    for p in doc.paragraphs:
        aplicar_substituicoes_em_paragrafo(p)
        
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    aplicar_substituicoes_em_paragrafo(p)
                    
    for section in doc.sections:
        for header_p in section.header.paragraphs:
            aplicar_substituicoes_em_paragrafo(header_p)
        for footer_p in section.footer.paragraphs:
            aplicar_substituicoes_em_paragrafo(footer_p)

    doc.save(temp_docx)
    
    pdf_bytes = None
    if converter_para_pdf_universal(temp_docx, temp_pdf):
        with open(temp_pdf, "rb") as f:
            pdf_bytes = f.read()
            
    for temp_file in [temp_docx, temp_pdf]:
        if os.path.exists(temp_file):
            try: os.remove(temp_file)
            except Exception: pass
            
    return pdf_bytes

st.title("Calculadora de CNR Clandestino")
st.write("Análise individual ou por lote de arquivos, com controle completo de datas de inspeção e períodos.")

# BARRA LATERAL
st.sidebar.header("Upload dos Arquivos")
files_historico = st.sidebar.file_uploader(
    "Upload Planilhas de Histórico UC (Selecione 1 ou Várias):",
    type=['xlsx', 'xls', 'csv'],
    accept_multiple_files=True
)

st.sidebar.markdown("---")
st.sidebar.header("Configuração Global de Tarifa")
tarifa_vigente_input = st.sidebar.number_input(
    "Tarifa Vigente (R$/kWh):",
    value=1.222451,
    format="%.6f",
    step=0.000001,
    help="Tarifa aplicável para a apuração do valor total faturado."
)

if files_historico:
    dfs_processados = []
    
    for f in files_historico:
        try:
            df_temp = carregar_arquivo_single(f)
            df_temp['UC'] = df_temp['UC'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            
            toi_extraido = "SN"
            if "TOI_" in f.name:
                parts = f.name.split("TOI_")
                if len(parts) > 1:
                    toi_extraido = parts[1].split("_")[0]
            df_temp['TOI_ARQUIVO'] = toi_extraido
            dfs_processados.append(df_temp)
        except Exception as e:
            st.sidebar.error(f"Erro ao carregar {f.name}: {e}")

    if dfs_processados:
        df_hist_full = pd.concat(dfs_processados, ignore_index=True)
        ucs_unicas = df_hist_full['UC'].dropna().unique().tolist()
        
        st.success(f"✓ **{len(files_historico)} arquivo(s)** carregado(s). **{len(ucs_unicas)} UC(s)** prontas para processamento.")

        detalhes_uc_dict = {}

        for uc_item in ucs_unicas:
            dados_uc_hist = df_hist_full[df_hist_full['UC'] == uc_item].copy()
            dados_uc_hist['DATA_FINAL'] = pd.to_datetime(dados_uc_hist['DATA_FINAL'], format='%d/%m/%Y', errors='coerce')
            dados_uc_hist['DATA_INICIAL'] = pd.to_datetime(dados_uc_hist['DATA_INICIAL'], format='%d/%m/%Y', errors='coerce')
            dados_uc_hist['ANO_REF'] = pd.to_numeric(dados_uc_hist['ANO_REF'], errors='coerce')
            dados_uc_hist['MES_REF'] = pd.to_numeric(dados_uc_hist['MES_REF'], errors='coerce')
            
            dados_uc_hist = dados_uc_hist.sort_values(by=['ANO_REF', 'MES_REF'])
            
            toi_val = dados_uc_hist['TOI_ARQUIVO'].iloc[0] if 'TOI_ARQUIVO' in dados_uc_hist.columns else "SN"
            
            dt_primeiro_ciclo = dados_uc_hist['DATA_INICIAL'].dropna().min()
            data_inspecao_padrao = dt_primeiro_ciclo.date() if (pd.notnull(dt_primeiro_ciclo) and hasattr(dt_primeiro_ciclo, 'date')) else datetime.now().date()

            nome_cliente = "N/A"
            if 'NOMECLIENTE' in dados_uc_hist.columns:
                nomes_validos = dados_uc_hist['NOMECLIENTE'].dropna().astype(str).str.strip()
                if not nomes_validos.empty:
                    nome_cliente = nomes_validos.iloc[-1]

            detalhes_uc_dict[uc_item] = {
                'nome_cliente': nome_cliente,
                'toi': toi_val,
                'data_inspecao_usada': pd.Timestamp(data_inspecao_padrao),
                'dados_uc_hist': dados_uc_hist
            }

        # -----------------------------------------------------------------------------
        # 1. SELEÇÃO E CONTROLE DE DATAS DA UC
        # -----------------------------------------------------------------------------
        st.subheader("1. Seleção da UC e Ajuste Manual de Datas")
        
        opcoes_dropdown = [f"UC: {uc_item} | TOI: {detalhes_uc_dict[uc_item]['toi']} | {detalhes_uc_dict[uc_item]['nome_cliente']}" for uc_item in ucs_unicas]
        selected_label = st.selectbox("Selecione a UC para ajustar datas e visualizar:", opcoes_dropdown)
        selected_uc = selected_label.split(" | ")[0].replace("UC: ", "").strip()

        det_sel = detalhes_uc_dict[selected_uc]
        
        col_man1, col_man2, col_man3 = st.columns(3)
        
        dt_declarada_sel = col_man1.date_input(
            "Início Declarado pelo Cliente (Opcional):",
            value=None,
            min_value=date(1900, 1, 1),
            max_value=date(2100, 12, 31),
            format="DD/MM/YYYY",
            key=f"dt_decl_{selected_uc}"
        )
        
        val_default_insp = det_sel['data_inspecao_usada'].date() if hasattr(det_sel['data_inspecao_usada'], 'date') else det_sel['data_inspecao_usada']
        
        dt_inspecao_sel = col_man2.date_input(
            "Data de Inspeção / Normalização (TR):",
            value=val_default_insp,
            min_value=date(1900, 1, 1),
            max_value=date(2100, 12, 31),
            format="DD/MM/YYYY",
            key=f"dt_insp_{selected_uc}"
        )

        nome_cliente_input = col_man3.text_input(
            "Nome do Cliente na Notificação:",
            value=det_sel['nome_cliente'],
            key=f"nome_cli_{selected_uc}"
        )

        data_inspecao_usada = pd.Timestamp(dt_inspecao_sel)
        dt_fim_efetiva = data_inspecao_usada
        
        if dt_declarada_sel is not None:
            dt_ini_efetiva = pd.Timestamp(dt_declarada_sel)
            dias_calculados = (dt_fim_efetiva - dt_ini_efetiva).days
            if dias_calculados > 180:
                st.warning(f"⚠️ Período calculado ({dias_calculados} dias) excede 180 dias. Limitando a 180 dias retroativos.")
                dias_cobranca = 180
                dt_ini_efetiva = dt_fim_efetiva - pd.Timedelta(days=180)
            else:
                dias_cobranca = max(0, dias_calculados)
        else:
            dias_cobranca = 180
            dt_ini_efetiva = dt_fim_efetiva - pd.Timedelta(days=180)

        # -----------------------------------------------------------------------------
        # 2. ANÁLISE DE HISTÓRICO, TITULARIDADE E CÁLCULO
        # -----------------------------------------------------------------------------
        dados_uc_hist = det_sel['dados_uc_hist']

        # CHECKBOX DE HISTÓRICO COMPLETO
        exibir_todo_historico = st.checkbox("🔍 Exibir histórico completo da UC (incluindo ciclos anteriores à inspeção)", value=False)
        if exibir_todo_historico:
            st.write("**Histórico Completo da UC (Ordenado por Mês/Ano):**")
            df_exibicao = dados_uc_hist.copy()
            df_exibicao = df_exibicao.sort_values(by=['ANO_REF', 'MES_REF'])
            df_exibicao['DATA_INICIAL'] = df_exibicao['DATA_INICIAL'].dt.strftime('%d/%m/%Y')
            df_exibicao['DATA_FINAL'] = df_exibicao['DATA_FINAL'].dt.strftime('%d/%m/%Y')
            cols_exib = [c for c in ['ANO_REF', 'MES_REF', 'DATA_INICIAL', 'DATA_FINAL', 'CONSUMOLIDOKWH', 'CRITERIOFATURAMENTO', 'NOMECLIENTE'] if c in df_exibicao.columns]
            st.dataframe(df_exibicao[cols_exib], use_container_width=True)

        pos_faturamento = dados_uc_hist[dados_uc_hist['DATA_FINAL'] > data_inspecao_usada].copy()
        
        if len(pos_faturamento) < 2:
            pos_faturamento = dados_uc_hist.copy()

        if len(pos_faturamento) >= 2:
            ciclos_apos_descarto = pos_faturamento.iloc[1:]
            ciclos_validos = ciclos_apos_descarto.head(3).copy()
            
            # 📌 VALIDAÇÃO DE TROCA DE TITULARIDADE NOS 3 CICLOS PÓS
            if 'NOMECLIENTE' in ciclos_validos.columns:
                nomes_ciclos = ciclos_validos['NOMECLIENTE'].dropna().astype(str).str.strip()
                nomes_ciclos = nomes_ciclos[nomes_ciclos != ""]

                if not nomes_ciclos.empty:
                    nomes_unicos = nomes_ciclos.unique()
                    if len(nomes_unicos) > 1:
                        st.warning(
                            "⚠️ **Troca de titularidade detectada:** Os 3 ciclos analisados após a inspeção têm clientes diferentes "
                            f"({', '.join(nomes_unicos)}). Isso indica alteração de titularidade e deve ser revisado antes da consolidação."
                        )
                    else:
                        st.success(f"✅ **Titularidade consistente:** Todos os 3 ciclos analisados após a inspeção pertencem ao mesmo cliente ({nomes_unicos[0]}).")

            pos_lidas = ciclos_validos[ciclos_validos['CRITERIOFATURAMENTO'].astype(str).str.upper() == 'LIDA'] if 'CRITERIOFATURAMENTO' in ciclos_validos.columns else pd.DataFrame()
            
            usou_minima = False
            if len(pos_lidas) > 0:
                base_selecionada = pos_lidas
            else:
                pos_minima = ciclos_validos[ciclos_validos['CRITERIOFATURAMENTO'].astype(str).str.upper() == 'MÍNIMA'] if 'CRITERIOFATURAMENTO' in ciclos_validos.columns else pd.DataFrame()
                if len(pos_minima) > 0:
                    base_selecionada = pos_minima
                    usou_minima = True
                else:
                    base_selecionada = ciclos_validos
            
            maior_ciclo_row = base_selecionada.nlargest(1, 'CONSUMOLIDOKWH') if len(base_selecionada) > 0 else pd.DataFrame()
            
            if len(maior_ciclo_row) > 0:
                maior_kwh = maior_ciclo_row['CONSUMOLIDOKWH'].values[0]
                
                consumo_diario = maior_kwh / 30.0
                consumo_estimado_total = consumo_diario * dias_cobranca
                valor_total_debito = consumo_estimado_total * tarifa_vigente_input
                
                st.markdown("---")
                st.subheader("2. Resumo da Memória de Cálculo")
                st.write(f"**Período Efetivo:** `{dias_cobranca} Dia(s)` | **Início:** `{dt_ini_efetiva.strftime('%d/%m/%Y')}` | **Fim:** `{dt_fim_efetiva.strftime('%d/%m/%Y')}`")
                
                st.write("**Janela dos 3 Ciclos Pós Observados:**")
                df_janela_exib = ciclos_validos.copy()
                df_janela_exib['DATA_INICIAL'] = df_janela_exib['DATA_INICIAL'].dt.strftime('%d/%m/%Y')
                df_janela_exib['DATA_FINAL'] = df_janela_exib['DATA_FINAL'].dt.strftime('%d/%m/%Y')
                cols_janela = [c for c in ['ANO_REF', 'MES_REF', 'DATA_INICIAL', 'DATA_FINAL', 'CONSUMOLIDOKWH', 'CRITERIOFATURAMENTO', 'NOMECLIENTE'] if c in df_janela_exib.columns]
                st.dataframe(df_janela_exib[cols_janela], use_container_width=True)
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Maior Consumo Pós", f"{fmt_br(maior_kwh, 2)} kWh")
                col2.metric("Consumo Diário Base", f"{fmt_br(consumo_diario, 2)} kWh/dia")
                col3.metric(f"Estimado Total ({dias_cobranca}d)", f"{fmt_br(consumo_estimado_total, 2)} kWh")
                col4.metric(f"Valor Total (Tarifa R$ {tarifa_vigente_input:.6f})", f"R$ {fmt_br(valor_total_debito, 2)}")

                # -----------------------------------------------------------------------------
                # 3. EMISSÃO DOS DOCUMENTOS DA UC SELECIONADA
                # -----------------------------------------------------------------------------
                st.markdown("---")
                st.subheader("3. Documentos Oficiais da UC Selecionada")
                
                caminho_modelo_excel = "Modelo_Carta_Lesao_Limpo.xlsx"
                caminho_modelo_word = "Modelo_Notificacao_Oficial.docx"
                toi_str = str(det_sel['toi'])

                col_doc1, col_doc2 = st.columns(2)
                
                with col_doc1:
                    st.markdown("### 📊 Detalhamento em Excel / PDF")
                    if os.path.exists(caminho_modelo_excel):
                        excel_preenchido = preencher_modelo_excel(
                            caminho_modelo_excel, nome_cliente_input, selected_uc, toi_str,
                            maior_ciclo_row, consumo_diario, dias_cobranca, usou_minima,
                            dt_ini_efetiva, dt_fim_efetiva, tarifa_vigente_input
                        )
                        st.download_button(
                            label="Baixar Detalhamento em Excel (.xlsx)",
                            data=excel_preenchido,
                            file_name=f"Detalhamento_do_Faturamento_UC_{selected_uc}_TOI_{toi_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        if st.button("Gerar Detalhamento em PDF (.pdf)", type="primary", use_container_width=True):
                            with st.spinner("Gerando PDF..."):
                                pdf_detalhamento = gerar_e_converter_pdf_demanda(
                                    caminho_modelo_excel, nome_cliente_input, selected_uc, toi_str,
                                    maior_ciclo_row, consumo_diario, dias_cobranca, usou_minima,
                                    selected_uc, dt_ini_efetiva, dt_fim_efetiva, tarifa_vigente_input
                                )
                                if pdf_detalhamento:
                                    st.download_button(
                                        label="💾 Salvar PDF do Detalhamento",
                                        data=pdf_detalhamento,
                                        file_name=f"Detalhamento_do_Faturamento_UC_{selected_uc}_TOI_{toi_str}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )

                with col_doc2:
                    st.markdown("### 📄 Notificação de Irregularidade")
                    if os.path.exists(caminho_modelo_word) and DOCX_DISPONIVEL:
                        dic_substituicoes = {
                            "{{ NUMERO_INSPECAO }}": toi_str,
                            "{{ DATA_EMISSAO }}": datetime.now().strftime('%d.%m.%Y'),
                            "{{ NOME_CLIENTE }}": nome_cliente_input,
                            "{{ ENDERECO_CLIENTE }}": "Endereço cadastrado no sistema",
                            "{{ UC_INSTALACAO }}": selected_uc,
                            "{{ DATA_INSPECAO }}": data_inspecao_usada.strftime('%d/%m/%Y'),
                            "{{ PERIODO_IRREGULARIDADE }}": dt_ini_efetiva.strftime('%d/%m/%Y'),
                            "{{ CONSUMO_REGISTRADO }}": "0",
                            "{{ CONSUMO_APURADO }}": fmt_br(consumo_estimado_total, 0),
                            "{{ VALOR_TOTAL }}": fmt_br(valor_total_debito, 2),
                            "{{ JUSTIFICATIVA_CRITERIO }}": "Não foi possível, no momento da inspeção, levantar a carga total da unidade consumidora."
                        }
                        if st.button("Gerar Notificação Oficial em PDF (.pdf)", type="primary", use_container_width=True):
                            with st.spinner("Gerando Notificação em PDF..."):
                                pdf_notificacao = gerar_notificacao_pdf(caminho_modelo_word, dic_substituicoes, selected_uc)
                                if pdf_notificacao:
                                    st.download_button(
                                        label="💾 Salvar Notificação Oficial (PDF)",
                                        data=pdf_notificacao,
                                        file_name=f"Notificacao_Irregularidade_UC_{selected_uc}_TOI_{toi_str}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )
            else:
                st.warning(f"⚠️ Nenhum ciclo faturado encontrado após a data de inspeção {data_inspecao_usada.strftime('%d/%m/%Y')}.")
        else:
            st.warning(f"⚠️ Mínimo de 2 ciclos faturados pós-inspeção não encontrados para a data {data_inspecao_usada.strftime('%d/%m/%Y')}.")

else:
    st.info("Por favor, faça o upload de um ou mais arquivos na barra lateral para iniciar.")