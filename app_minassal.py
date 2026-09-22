import streamlit as st
import pandas as pd
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta

# --- BIBLIOTECAS PARA O PDF (MODO PAISAGEM) ---
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# --- CONFIGURAÇÃO DA PÁGINA E DESIGN DARK ---
st.set_page_config(page_title="Coleta Minassal", page_icon="👑", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stApp { background-color: #0d0d0d; }
    div.stButton > button {
        height: 60px; font-size: 18px; font-weight: bold; border-radius: 8px;
        border: 2px solid #E2001A; color: #FFFFFF; background-color: #1A1A1A;
    }
    div.stButton > button:hover { background-color: #E2001A; color: white; }
    h1, h2, h3, p, span, label { color: white !important; }
    </style>
""", unsafe_allow_html=True)

st.title("📱 Portal de Auditoria - Royal Canin")
st.markdown("---")

# --- LISTA MESTRA: SELEÇÃO DE OURO ---
CODIGOS_OURO = {"97996", "98018", "98224", "98230", "98435", "97985", "98037", "98011", "98015", "98139", "98157", "98492", "98834", "99101", "98022", "97991", "98019", "97994", "98016", "98222", "98197", "98433", "97983", "98122", "98126", "98124", "98144", "98137", "98469", "98467", "98640", "98518", "98520", "98490", "99757", "99753", "99750", "98249", "99187", "98328", "98357", "98331", "98350", "98364", "98334", "98361", "98338", "98434", "98327", "98340", "98365", "98360", "98333", "98353", "98332", "98356", "98336", "98450", "98461", "98589", "98452", "98639", "98631", "98491", "98489", "98719", "98721", "98852", "98024", "97993", "98021"}

# --- FUNÇÃO PARA BUSCAR ARQUIVOS ---
def buscar_arquivo(nome_base):
    for ext in [".csv", ".xlsx"]:
        caminho = nome_base + ext
        if os.path.exists(caminho): return caminho
    return None

ARQUIVO_VENDAS = buscar_arquivo("Vendas")
ARQUIVO_MG = buscar_arquivo("Tabela_MG")

ROTAS_PROMOTORES = {
    "Pamela": ["POCOS DE CALDAS", "POÇOS DE CALDAS", "ANDRADAS", "VARGINHA", "TRES CORACOES", "TRÊS CORAÇÕES", "TRES PONTAS", "TRÊS PONTAS", "ITAJUBA", "ITAJUBÁ", "POUSO ALEGRE"],
    "Fernanda": ["JUIZ DE FORA", "JUIZ DE FORA/MG"],
    "Madalla": ["CONSELHEIRO LAFAIETE", "GUARANI", "GUIDOVAL", "MURIAE", "MURIAÉ", "PIRAUBA", "PIRAÚBA", "RIO POMBA", "TOCANTINS", "UBA", "UBÁ", "VICOSA", "VIÇOSA", "VISCONDE DO RIO BRANCO", "SAO JOAO NEPOMUCENO"]
}

# --- CARREGAR DADOS COM SUPORTE ROBUSTO A CSV E EXCEL ---
@st.cache_data
def carregar_dados(caminho):
    if not caminho: return pd.DataFrame()
    try:
        if str(caminho).endswith('.csv'):
            # Tenta ler com separador automático ou vírgula/ponto-e-vírgula
            try:
                df = pd.read_csv(caminho, sep=',', encoding='utf-8', low_memory=False)
                if len(df.columns) <= 1:
                    df = pd.read_csv(caminho, sep=';', encoding='utf-8', low_memory=False)
            except:
                df = pd.read_csv(caminho, sep=';', encoding='latin1', low_memory=False)
        else:
            df = pd.read_excel(caminho, sheet_name=0, engine='openpyxl')
            
        if len(df) > 0 and 'TOTAL GERAL' in str(df.iloc[0, 0]):
            df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Erro ao ler {caminho}: {e}")
        return pd.DataFrame()

# --- FUNÇÃO DE GERAÇÃO DE PDF ---
def gerar_pdf_relatorio(promotor, loja, cidade, df_preenchido, df_vendas_original):
    hora_brasil = datetime.now() - timedelta(hours=3)
    data_str = hora_brasil.strftime('%d/%m/%Y %H:%M')
    
    caminho_pdf = f"Auditoria_{loja[:10].replace(' ', '_')}.pdf"
    doc = SimpleDocTemplate(caminho_pdf, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
    estilos = getSampleStyleSheet()
    elementos = []
    
    elementos.append(Paragraph(f"<b>RELATÓRIO DE AUDITORIA & PRECIFICAÇÃO - ROYAL CANIN</b>", estilos['Title']))
    elementos.append(Paragraph(f"<b>LOJA:</b> {loja}", estilos['Heading2']))
    elementos.append(Paragraph(f"<b>PROMOTOR(A):</b> {promotor} | <b>CIDADE:</b> {cidade} | <b>DATA:</b> {data_str}", estilos['Normal']))
    elementos.append(Spacer(1, 10))
    
    data = [["PRODUTO", "CÓDIGO", "PREÇO SUGERIDO", "MARKUP RECOMENDADO", "PREÇO NA LOJA", "MARKUP PRATICADO", "SITUAÇÃO / DESVIO"]]
    estilo_tabela = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2001A")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),      
        ('ALIGN', (1,0), (-1,-1), 'CENTER'),    
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('FONTSIZE', (0,0), (-1,-1), 8)
    ]

    def limpar_valor(v):
        if pd.isna(v) or v == "" or str(v).lower() == "none": return 0.0
        s = str(v).replace("R$", "").replace(" ", "").strip()
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try: return float(s)
        except: return 0.0

    produtos_ausentes_detalhes = []

    for i, linha in enumerate(df_preenchido.itertuples()):
        idx = i + 1
        nome = str(linha.PRODUTO).replace("⭐ ", "")[:38]
        cod = str(linha.CÓDIGO)
        
        p_sug = limpar_valor(linha.SUGERIDO)
        markup_rec = getattr(linha, 'MARKUP_REC_VAL', 38.0)
        nao_tem = getattr(linha, "NA_LOJA", False)
        
        if nao_tem:
            sit, cor = "SEM PRODUTO", colors.gray
            p_loja_str = "AUSENTE"
            p_sug_str = f"R$ {p_sug:.2f}"
            m_rec_str = f"{markup_rec:.1f}%"
            m_prat_str = "--"
            
            df_hist = df_vendas_original[
                (df_vendas_original['CLIENTE NOME'] == loja) & 
                (df_vendas_original['PRODUTO CODIGO'].astype(str).str.replace('.0', '', regex=False).str.strip() == cod)
            ]
            
            dt_fmt = "Desconhecida"
            op_tipo = "VENDA"
            rca_resp = "Não identificado"
            
            if not df_hist.empty:
                df_hist['DATA_DT'] = pd.to_datetime(df_hist['DATA'], errors='coerce')
                df_hist = df_hist.sort_values(by='DATA_DT', ascending=False)
                ultima_linha = df_hist.iloc[0]
                
                dt_raw = ultima_linha['DATA_DT']
                if pd.notna(dt_raw):
                    dt_fmt = dt_raw.strftime('%d/%m/%Y')
                op_tipo = str(ultima_linha.get('OPERACAO', 'VENDA')).strip().upper()
                rca_resp = str(ultima_linha.get('RCA NOME', 'Não identificado')).strip().upper()
                
            produtos_ausentes_detalhes.append({
                "produto": str(linha.PRODUTO).replace("⭐ ", ""),
                "codigo": cod,
                "data": dt_fmt,
                "operacao": op_tipo,
                "rca": rca_resp
            })
        else:
            p_loja = limpar_valor(linha.PREÇO_NA_LOJA)
            p_sug_str = f"R$ {p_sug:.2f}"
            m_rec_str = f"{markup_rec:.1f}%"

            if p_loja <= 0:
                sit, cor = "OPORTUNIDADE", colors.orange
                p_loja_str = "--"
                m_prat_str = "--"
            else:
                diff = ((p_loja / p_sug) - 1) * 100
                markup_prat = markup_rec + diff
                m_prat_str = f"{markup_prat:.1f}%"
                p_loja_str = f"R$ {p_loja:.2f}"

                if p_loja <= (p_sug + 0.05):
                    sit, cor = "CORRETO (Abaixo/Igual)", colors.green
                else:
                    if diff >= 1:
                        sit, cor = f"ACIMA +{diff:.1f}%", colors.red
                    else:
                        sit, cor = "CORRETO (Abaixo/Igual)", colors.green
                
        data.append([nome, cod, p_sug_str, m_rec_str, p_loja_str, m_prat_str, sit])
        estilo_tabela.append(('TEXTCOLOR', (3, idx), (3, idx), colors.HexColor("#166534")))
        if not nao_tem and p_loja > 0:
            if p_loja <= (p_sug + 0.05):
                estilo_tabela.append(('TEXTCOLOR', (5, idx), (5, idx), colors.HexColor("#166534")))
            else:
                estilo_tabela.append(('TEXTCOLOR', (5, idx), (5, idx), colors.HexColor("#991b1b")))
        estilo_tabela.append(('TEXTCOLOR', (6, idx), (6, idx), cor))
        
        if cod in CODIGOS_OURO: 
            estilo_tabela.append(('BACKGROUND', (0, idx), (0, idx), colors.HexColor("#FEF3C7")))

    t = Table(data, colWidths=[190, 50, 95, 105, 95, 105, 140])
    t.setStyle(TableStyle(estilo_tabela))
    elementos.append(t)
    
    if produtos_ausentes_detalhes:
        elementos.append(Spacer(1, 12))
        elementos.append(Paragraph("<b>HISTÓRICO DE ITENS AUSENTES (ÚLTIMA COMPRA)</b>", estilos['Heading3']))
        elementos.append(Spacer(1, 3))
        
        for item in produtos_ausentes_detalhes:
            texto_detalhe = (
                f"• <b>Produto:</b> {item['produto']} (Cód: {item['codigo']}) | "
                f"<b>Última Compra:</b> {item['data']} | <b>Operação:</b> {item['operacao']} | "
                f"<b>RCA:</b> {item['rca']}"
            )
            elementos.append(Paragraph(texto_detalhe, estilos['Normal']))
            elementos.append(Spacer(1, 2))

    doc.build(elementos)
    return caminho_pdf

def enviar_email_coleta(promotor, loja, cidade, df_editado, feedback, df_vendas_original):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = ["benedito.bandola@minassal.com.br"]

    caminho_pdf = gerar_pdf_relatorio(promotor, loja, cidade, df_editado, df_vendas_original)
    
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, ", ".join(destino), f"✅ Auditoria PDV - {loja} ({promotor})"
    corpo = f"<html><body><h2 style='color: #E2001A;'>Auditoria Royal Canin Recebida</h2><p><b>Loja:</b> {loja}<br><b>Promotor:</b> {promotor}<br><b>Cidade:</b> {cidade}</p><hr><p><b>Observações:</b><br>{feedback if feedback.strip() else 'Sem comentários.'}</p></body></html>"
    msg.attach(MIMEText(corpo, 'html'))

    try:
        with open(caminho_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=os.path.basename(caminho_pdf))
            msg.attach(anexo)
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls(); s.login(remetente, senha); s.sendmail(remetente, destino, msg.as_string()); s.quit()
        os.remove(caminho_pdf)
        return True, "Relatório enviado com sucesso!"
    except Exception as e: return False, f"Erro: {e}"

# --- INTERFACE ---
if 'promotor_logado' not in st.session_state: st.session_state.promotor_logado = None

vendas = carregar_dados(ARQUIVO_VENDAS)
tab_mg = carregar_dados(ARQUIVO_MG)

if not vendas.empty:
    if st.session_state.promotor_logado is None:
        st.subheader("Selecione o seu Perfil")
        c1, c2, c3 = st.columns(3)
        if c1.button("👩‍💼 PAMELA", use_container_width=True): st.session_state.promotor_logado = "Pamela"; st.rerun()
        if c2.button("👩‍💼 FERNANDA", use_container_width=True): st.session_state.promotor_logado = "Fernanda"; st.rerun()
        if c3.button("👩‍💼 MADALLA", use_container_width=True): st.session_state.promotor_logado = "Madalla"; st.rerun()
    else:
        promotor = st.session_state.promotor_logado
        st.sidebar.info(f"👤 {promotor}")
        if st.sidebar.button("Sair"): st.session_state.promotor_logado = None; st.rerun()

        cidades_autorizadas = ROTAS_PROMOTORES[promotor]
        vendas['CIDADE_LIMPA'] = vendas['CIDADE'].astype(str).str.upper().str.strip()
        df_f = vendas[vendas['CIDADE_LIMPA'].isin(cidades_autorizadas)]
        
        loja_sel = st.selectbox("🏪 Selecione a Loja:", ["-- Selecione --"] + sorted(df_f['CLIENTE NOME'].dropna().unique()))

        mapa_precos = {}
        if not tab_mg.empty:
            col_cod_tab = next((c for c in tab_mg.columns if "COD" in c), None)
            col_preco = next((c for c in tab_mg.columns if "SUGEST" in c or "RECOMEN" in c or "PRECO" in c), None)
            
            if col_cod_tab and col_preco:
                for _, row in tab_mg.iterrows():
                    c_val = str(row[col_cod_tab])
                    if c_val.endswith('.0'): c_val = c_val[:-2]
                    c_val = c_val.strip()
                    try:
                        mapa_precos[c_val] = float(row[col_preco])
                    except:
                        pass

        if loja_sel != "-- Selecione --":
            df_loja = df_f[df_f['CLIENTE NOME'] == loja_sel].drop_duplicates(subset=['PRODUTO CODIGO'])
            dados_tabela = []
            for _, r in df_loja.iterrows():
                cod = str(r['PRODUTO CODIGO'])
                if cod.endswith('.0'): cod = cod[:-2]
                cod = cod.strip()
                
                p_sug = mapa_precos.get(cod, 0.0)
                if p_sug > 0:
                    dados_tabela.append({
                        "NA_LOJA": False,  
                        "CÓDIGO": cod, 
                        "PRODUTO": ("⭐ " if cod in CODIGOS_OURO else "") + str(r['PRODUTO NOME']), 
                        "SUGERIDO": f"R$ {float(p_sug):.2f}", 
                        "MARKUP_REC_VAL": 38.0, 
                        "PREÇO_NA_LOJA": 0.0
                    })

            if dados_tabela:
                st.info("💡 **Legenda de Auditoria:** Informe o preço praticado na loja. O sistema exibe o Markup Recomendado (Verde) e o Markup Praticado (Verde se igual/abaixo do recomendado, Vermelho se estiver acima).")
                
                df_editor = st.data_editor(
                    pd.DataFrame(dados_tabela), 
                    use_container_width=True, 
                    hide_index=True, 
                    disabled=["CÓDIGO", "PRODUTO", "SUGERIDO", "MARKUP_REC_VAL"],
                    column_config={
                        "NA_LOJA": st.column_config.CheckboxColumn(
                            "NÃO TEM NA LOJA?",
                            help="Marque caso o produto esteja ausente.",
                            default=False,
                        ),
                        "PREÇO_NA_LOJA": st.column_config.NumberColumn(
                            "PREÇO NA LOJA (R$)",
                            min_value=0.0,
                            format="R$ %.2f"
                        ),
                        "SUGERIDO": st.column_config.TextColumn("PREÇO SUGERIDO"),
                        "MARKUP_REC_VAL": st.column_config.NumberColumn("MARKUP RECOMENDADO (%)", format="%.1f%%")
                    }
                )
                
                obs = st.text_area("Inteligência de Campo:")
                
                if st.button("🚀 ENVIAR AUDITORIA", use_container_width=True):
                    tem_erro = False
                    for _, row_ed in df_editor.iterrows():
                        nao_tem = row_ed["NA_LOJA"]
                        p_loja = row_ed["PREÇO_NA_LOJA"]
                        
                        if not nao_tem:
                            try:
                                if float(p_loja) <= 0.0:
                                    tem_erro = True
                                    break
                            except:
                                tem_erro = True
                                break

                    if tem_erro:
                        st.error("⚠️ Atenção: Há produtos sem o preço preenchido na loja! Se o produto estiver ausente, marque a caixinha 'NÃO TEM NA LOJA?'.")
                    else:
                        cid_final = df_f[df_f['CLIENTE NOME'] == loja_sel]['CIDADE'].iloc[0]
                        ok, res = enviar_email_coleta(promotor, loja_sel, cid_final, df_editor, obs, vendas)
                        if ok: st.success(res); st.balloons()
                        else: st.error(res)
            else:
                st.warning("Nenhum produto com preço sugerido mapeado foi encontrado para esta loja.")
else:
    st.error("Arquivo de vendas não encontrado ou vazio na raiz do repositório.")
