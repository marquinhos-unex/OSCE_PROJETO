import sys
import subprocess

# Instalação automática das dependências caso não estejam no ambiente
for pkg, import_name in [("python-docx", "docx"), ("pandas", "pandas")]:
    try:
        __import__(import_name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

import os
import re
import io
import sqlite3
import hashlib
import pandas as pd
import streamlit as st

DB_FILE = "banco_osce.db"

# ==========================================
# 1. BANCO DE DADOS (SQLite + Criptografia)
# ==========================================

def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS utilizadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            email TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS estacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            professor_id INTEGER NOT NULL,
            semestre TEXT,
            data TEXT,
            estacao TEXT,
            componente TEXT,
            professor_nome TEXT,
            cenario_texto TEXT,
            tarefa_1 TEXT, tarefa_2 TEXT, tarefa_3 TEXT,
            item_1 TEXT, item_2 TEXT, item_3 TEXT, item_4 TEXT, item_5 TEXT,
            item_6 TEXT, item_7 TEXT, item_8 TEXT, item_9 TEXT, item_10 TEXT,
            FOREIGN KEY (professor_id) REFERENCES utilizadores (id)
        )
    ''')
    conn.commit()
    conn.close()

def gerar_hash_senha(senha):
    return hashlib.sha256(senha.encode()).hexdigest()

def cadastrar_utilizador(nome, email, senha):
    try:
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("INSERT INTO utilizadores (nome, email, senha_hash) VALUES (?, ?, ?)",
                  (nome, email.lower().strip(), gerar_hash_senha(senha)))
        conn.commit()
        conn.close()
        return True, "Registo efetuado com sucesso!"
    except sqlite3.IntegrityError:
        return False, "Este e-mail já está registado."
    except Exception as e:
        return False, f"Erro ao cadastrar: {e}"

def autenticar_utilizador(email, senha):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, nome, email FROM utilizadores WHERE email = ? AND senha_hash = ?",
              (email.lower().strip(), gerar_hash_senha(senha)))
    user = c.fetchone()
    conn.close()
    return user

def salvar_estacao_db(prof_id, dados):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        INSERT INTO estacoes (
            professor_id, semestre, data, estacao, componente, professor_nome,
            cenario_texto, tarefa_1, tarefa_2, tarefa_3,
            item_1, item_2, item_3, item_4, item_5,
            item_6, item_7, item_8, item_9, item_10
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        prof_id, dados.get('semestre'), dados.get('data'), dados.get('ESTAÇÃO'),
        dados.get('Componente'), dados.get('Professor'), dados.get('cenario_texto'),
        dados.get('tarefa_1'), dados.get('tarefa_2'), dados.get('tarefa_3'),
        dados.get('item_1'), dados.get('item_2'), dados.get('item_3'), dados.get('item_4'), dados.get('item_5'),
        dados.get('item_6'), dados.get('item_7'), dados.get('item_8'), dados.get('item_9'), dados.get('item_10')
    ))
    conn.commit()
    conn.close()

def carregar_estacoes_professor(prof_id):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM estacoes WHERE professor_id = ?", conn, params=(prof_id,))
    conn.close()
    return df

def excluir_estacao_db(estacao_id, prof_id):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("DELETE FROM estacoes WHERE id = ? AND professor_id = ?", (estacao_id, prof_id))
    linhas_afetadas = c.rowcount
    conn.commit()
    conn.close()
    return linhas_afetadas > 0

# ==========================================
# 2. GERADOR DE DOCX (WORD)
# ==========================================

def set_cell_background(cell, fill_hex):
    """Aplica cor de fundo a uma célula da tabela no Word."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)

def formatar_paragrafo_item(paragraph, item_num, texto_raw):
    """Formata o item dividindo rubrica/estudante em vermelho/itálico e fala do ator em texto normal."""
    p_run = paragraph.add_run(f"{item_num}. ")
    p_run.bold = True
    
    if not isinstance(texto_raw, str) or pd.isna(texto_raw) or texto_raw.strip() == "":
        return

    texto = texto_raw.strip()
    
    if '|' in texto:
        partes = texto.split('|', 1)
        rubrica = partes[0].strip()
        fala_ator = partes[1].strip()
        if fala_ator and not (fala_ator.startswith('"') or fala_ator.startswith('“')):
            fala_ator = f'"{fala_ator}"'
        
        run_rub = paragraph.add_run(rubrica + " ")
        run_rub.italic = True
        run_rub.font.color.rgb = RGBColor(208, 0, 0)
        
        run_fala = paragraph.add_run(fala_ator)
        run_fala.font.color.rgb = RGBColor(0, 0, 0)
    elif 'Estudante:' in texto:
        partes = texto.split('Estudante:', 1)
        if partes[0]:
            paragraph.add_run(partes[0])
        run_est = paragraph.add_run("Estudante:" + partes[1])
        run_est.italic = True
        run_est.font.color.rgb = RGBColor(208, 0, 0)
    else:
        paragraph.add_run(texto)

def gerar_bytes_docx_osce(dados_estacao):
    doc = docx.Document()

    # Configuração de Página: A4 em Orientação Paisagem (Landscape)
    section = doc.sections[0]
    section.orientation = docx.enum.section.WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.0)
    section.right_margin = Cm(1.0)

    # Dados da Estação
    semestre = str(dados_estacao.get('semestre', ''))
    estacao = str(dados_estacao.get('ESTAÇÃO', dados_estacao.get('estacao', '')))
    componente = str(dados_estacao.get('Componente', dados_estacao.get('componente', '')))
    titulo_estacao = f"OSCE {semestre} – {estacao} ({componente})" if semestre else str(dados_estacao.get('titulo_estacao', 'OSCE - ESTAÇÃO'))
    data_prova = str(dados_estacao.get('data', ''))
    prof_cadastrado = str(dados_estacao.get('Professor', dados_estacao.get('professor_nome', '')))
    if prof_cadastrado == 'nan': prof_cadastrado = ""

    # Título Principal
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    run_t = p_title.add_run(titulo_estacao)
    run_t.bold = True
    run_t.font.size = Pt(12)

    # Subtítulo (Data)
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(6)
    run_s = p_sub.add_run(data_prova)
    run_s.font.size = Pt(9.5)

    # Tabela de Identificação (Estudante e Professor)
    tab_id = doc.add_table(rows=1, cols=2)
    tab_id.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_id.autofit = False
    
    cell_est, cell_prof = tab_id.rows[0].cells
    cell_est.width = Cm(17.0)
    cell_prof.width = Cm(10.7)

    p_est = cell_est.paragraphs[0]
    p_est.paragraph_format.space_after = Pt(4)
    run_est_label = p_est.add_run("Nome do(a) estudante: ")
    run_est_label.font.size = Pt(9.5)
    p_est.add_run("____________________________________________________________________").font.size = Pt(9.5)

    p_prof = cell_prof.paragraphs[0]
    p_prof.paragraph_format.space_after = Pt(4)
    run_prof_label = p_prof.add_run("Professor(a): ")
    run_prof_label.font.size = Pt(9.5)
    if prof_cadastrado:
        run_prof_val = p_prof.add_run(f" {prof_cadastrado} ")
        run_prof_val.underline = True
        run_prof_val.font.size = Pt(9.5)
    else:
        p_prof.add_run("__________________________________").font.size = Pt(9.5)

    # Quadro de Cenário e Tarefas
    cenario_texto = str(dados_estacao.get('cenario_texto', ''))
    tarefa_1 = str(dados_estacao.get('tarefa_1', ''))
    tarefa_2 = str(dados_estacao.get('tarefa_2', ''))
    tarefa_3 = str(dados_estacao.get('tarefa_3', ''))

    tab_cenario = doc.add_table(rows=3, cols=1)
    tab_cenario.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_cenario.style = 'Table Grid'

    # Cabeçalho do Cenário
    cell_c1 = tab_cenario.rows[0].cells[0]
    set_cell_background(cell_c1, "E6E6E6")
    p_c1 = cell_c1.paragraphs[0]
    p_c1.paragraph_format.space_before = Pt(3)
    p_c1.paragraph_format.space_after = Pt(3)
    r_c1 = p_c1.add_run("CENÁRIO")
    r_c1.bold = True
    r_c1.font.size = Pt(9.5)

    # Texto do Cenário
    cell_c2 = tab_cenario.rows[1].cells[0]
    p_c2 = cell_c2.paragraphs[0]
    p_c2.paragraph_format.space_before = Pt(3)
    p_c2.paragraph_format.space_after = Pt(3)
    r_c2 = p_c2.add_run(cenario_texto)
    r_c2.font.size = Pt(8.5)

    # Texto das Tarefas
    cell_c3 = tab_cenario.rows[2].cells[0]
    p_c3 = cell_c3.paragraphs[0]
    p_c3.paragraph_format.space_before = Pt(3)
    p_c3.paragraph_format.space_after = Pt(3)
    
    r_t_title = p_c3.add_run("Nos próximos 4 minutos, deverão ser realizadas as seguintes TAREFAS:\n")
    r_t_title.bold = True
    r_t_title.font.size = Pt(8.5)

    p_c3.add_run(f"1. {tarefa_1}\n2. {tarefa_2}\n3. {tarefa_3}").font.size = Pt(8.5)

    # Espaçador
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Tabela de Checklist (Item de desempenho e Pontuação)
    tab_check = doc.add_table(rows=15, cols=4)
    tab_check.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab_check.style = 'Table Grid'

    col_widths = [Cm(18.0), Cm(2.8), Cm(3.9), Cm(2.8)]
    for row in tab_check.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = width

    # Linha 0: Cabeçalho Checklist (Mesclada)
    c_head = tab_check.rows[0].cells[0]
    for i in range(1, 4):
        c_head.merge(tab_check.rows[0].cells[i])
    set_cell_background(c_head, "E6E6E6")
    p_ch = c_head.paragraphs[0]
    r_ch = p_ch.add_run("CHECKLIST DA ESTAÇÃO PARA AVALIADORES E PACIENTES SIMULADOS")
    r_ch.bold = True
    r_ch.font.size = Pt(9.5)

    # Linha 1: Legenda (Mesclada)
    c_leg = tab_check.rows[1].cells[0]
    for i in range(1, 4):
        c_leg.merge(tab_check.rows[1].cells[i])
    set_cell_background(c_leg, "E6E6E6")
    p_lg = c_leg.paragraphs[0]
    
    r_lg1 = p_lg.add_run("Instruções do ator: ")
    r_lg1.bold = True
    r_lg1.font.size = Pt(8.5)
    
    r_lg2 = p_lg.add_run("negrito                ")
    r_lg2.font.size = Pt(8.5)

    r_lg3 = p_lg.add_run("Instruções esperadas para o(a) estudante: ")
    r_lg3.bold = True
    r_lg3.font.size = Pt(8.5)

    r_lg4 = p_lg.add_run("vermelho e itálico")
    r_lg4.italic = True
    r_lg4.font.color.rgb = RGBColor(208, 0, 0)
    r_lg4.font.size = Pt(8.5)

    # Linha 2 e 3: Títulos das Colunas
    cell_item_h = tab_check.rows[2].cells[0]
    cell_item_h.merge(tab_check.rows[3].cells[0])
    p_ith = cell_item_h.paragraphs[0]
    p_ith.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ith = p_ith.add_run("Itens de desempenho avaliados")
    r_ith.bold = True
    r_ith.font.size = Pt(8.5)

    cell_des_h = tab_check.rows[2].cells[1]
    cell_des_h.merge(tab_check.rows[2].cells[2])
    cell_des_h.merge(tab_check.rows[2].cells[3])
    p_dh = cell_des_h.paragraphs[0]
    p_dh.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_dh = p_dh.add_run("Desempenho observado")
    r_dh.bold = True
    r_dh.font.size = Pt(8.5)

    sub_headers = ["Inadequado", "Parcialmente adequado", "Adequado"]
    for idx, sh_text in enumerate(sub_headers, start=1):
        cell_sh = tab_check.rows[3].cells[idx]
        p_sh = cell_sh.paragraphs[0]
        p_sh.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_sh = p_sh.add_run(sh_text)
        r_sh.bold = True
        r_sh.font.size = Pt(8.0)

    # Linhas 4 a 13: Itens 1 a 10
    for i in range(1, 11):
        row_idx = i + 3
        cell_item = tab_check.rows[row_idx].cells[0]
        p_item = cell_item.paragraphs[0]
        p_item.paragraph_format.space_before = Pt(2)
        p_item.paragraph_format.space_after = Pt(2)
        
        texto_item_raw = str(dados_estacao.get(f'item_{i}', ''))
        formatar_paragrafo_item(p_item, i, texto_item_raw)

    # Linha 14: Rodapé de Regras e Acertos
    cell_regra = tab_check.rows[14].cells[0]
    p_reg = cell_regra.paragraphs[0]
    r_reg = p_reg.add_run("Quando o total de acertos findar em acerto parcialmente adequado, considerar a quantidade de acertos imediatamente superior.")
    r_reg.font.size = Pt(8.0)

    cell_acerto = tab_check.rows[14].cells[1]
    for i in range(2, 4):
        cell_acerto.merge(tab_check.rows[14].cells[i])
    set_cell_background(cell_acerto, "F2F2F2")
    p_ac = cell_acerto.paragraphs[0]
    p_ac.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ac = p_ac.add_run("Nº acertos______ X (2 itens parcialmente adequados = 1 item adequado)    Nota PONTOS (máximo 20 pontos): ________")
    r_ac.font.size = Pt(8.0)

    # Espaçador
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # Avaliação Global
    p_glob_t = doc.add_paragraph()
    r_gt = p_glob_t.add_run("Avaliação global do professor sobre a postura do estudante:")
    r_gt.bold = True
    r_gt.font.size = Pt(9.5)

    p_glob_o = doc.add_paragraph()
    p_glob_o.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_glob_o.paragraph_format.space_before = Pt(2)
    r_go = p_glob_o.add_run("(  ) Insatisfatória         (  ) Aceitável         (  ) Boa         (  ) Muito boa         (  ) Extraordinária")
    r_go.font.size = Pt(8.5)

    # Gravar em memória
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 3. INTERFACE STREAMLIT
# ==========================================

st.set_page_config(page_title="Sistema OSCE", layout="wide", initial_sidebar_state="collapsed")
init_db()

# Estilização
st.markdown("""
    <style>
    .header-bar {
        background-color: #E0E0E0;
        padding: 8px 20px;
        border-radius: 4px;
        display: flex;
        justify-content: flex-end;
        align-items: center;
        margin-bottom: 10px;
    }
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

if "utilizador" not in st.session_state:
    st.session_state["utilizador"] = None

# --- TELA DE LOGIN / REGISTO ---
if st.session_state["utilizador"] is None:
    st.title("🏥 Sistema OSCE - Acesso Restrito")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        aba_login, aba_registo = st.tabs(["🔑 Entrar", "📝 Registar Conta"])
        with aba_login:
            email_login = st.text_input("E-mail", key="login_email")
            senha_login = st.text_input("Palavra-passe", type="password", key="login_senha")
            if st.button("Entrar", type="primary", use_container_width=True):
                user = autenticar_utilizador(email_login, senha_login)
                if user:
                    st.session_state["utilizador"] = {"id": user[0], "nome": user[1], "email": user[2]}
                    st.rerun()
                else:
                    st.error("Credenciais inválidas.")

        with aba_registo:
            nome_reg = st.text_input("Nome Completo", key="reg_nome")
            email_reg = st.text_input("E-mail Institucional", key="reg_email")
            senha_reg = st.text_input("Definir Palavra-passe", type="password", key="reg_senha")
            if st.button("Criar Conta", use_container_width=True):
                if nome_reg and email_reg and senha_reg:
                    sucesso, msg = cadastrar_utilizador(nome_reg, email_reg, senha_reg)
                    if sucesso:
                        st.success(msg)
                    else:
                        st.error(msg)
                else:
                    st.warning("Preencha todos os campos.")

# --- TELA PRINCIPAL ---
else:
    usuario_atual = st.session_state["utilizador"]
    
    # BARRA SUPERIOR
    c_top_1, c_top_2 = st.columns([3, 1])
    with c_top_2:
        with st.popover(f"👤 {usuario_atual['nome']}", use_container_width=True):
            st.markdown(f"**Sessão Ativa**")
            st.caption(f"E-mail: {usuario_atual['email']}")
            st.divider()
            
            if st.button("🚪 Terminar Sessão", type="primary", use_container_width=True):
                st.session_state["utilizador"] = None
                st.rerun()

    # Layout de 2 Colunas
    col_esquerda, col_direita = st.columns([1.1, 1.3], gap="medium")

    # --------------------------------------------------
    # COLUNA ESQUERDA: FORMULÁRIO COM ROLAGEM PRÓPRIA
    # --------------------------------------------------
    with col_esquerda:
        with st.container(height=600, border=True):
            st.subheader("Cabeçalho")
            lista_componentes = [
                "Habilidades Médicas I",
                "Habilidades Médicas II",
                "Habilidades Médicas III",
                "Habilidades Médicas IV",
                "Clínica Médica",
                "Cirurgia Geral",
                "Pediatria",
                "Ginecologia e Obstetrícia",
                "Saúde da Família e Comunidade",
                "Outro (Digitar manualmente...)"
            ]
            
            componente_selecionado = st.selectbox(
                "Componente Curricular",
                options=lista_componentes,
                key="select_componente"
            )
            
            if componente_selecionado == "Outro (Digitar manualmente...)":
                componente = st.text_input("Especifique o Componente", key="custom_componente")
            else:
                componente = componente_selecionado
            
            c_sem, c_etapa = st.columns(2)
            semestre = c_sem.text_input("Semestre", value="6º", key="semestre")
            etapa = c_etapa.text_input("Etapa", value="Estação 1", key="etapa")
            
            professor_nome = st.text_input("Professor", value=f"Prof. {usuario_atual['nome']}", key="prof_nome")
            
            c_data, c_empty = st.columns(2)
            data = c_data.text_input("Data da Prova", value="20/10/2026", key="data")

            if st.button("Inserir / Salvar Identificação", type="secondary", use_container_width=True):
                st.toast("Dados do cabeçalho confirmados!", icon="✅")

            st.divider()

            st.subheader("Inicial (Cenário e Tarefas)")
            cenario_texto = st.text_area("Cenário Clínico", value="Paciente de 45 anos com dor precordial...", height=80)
            
            proc1 = st.text_input("Procedimento 01 (Tarefa 1)", value="Anamnese direcionada.")
            proc2 = st.text_input("Procedimento 02 (Tarefa 2)", value="Exame físico cardiovascular.")
            proc3 = st.text_input("Procedimento 03 (Tarefa 3)", value="Solicitação de exames.")

            st.divider()

            st.subheader("Itens de Desempenho (Checklist)")
            itens = {}
            for i in range(1, 11):
                itens[f'item_{i}'] = st.text_input(f"Item {i:02d}", key=f"inp_item_{i}")

            st.divider()

            dados_atual = {
                'semestre': semestre,
                'data': data,
                'ESTAÇÃO': etapa,
                'Componente': componente,
                'Professor': professor_nome,
                'cenario_texto': cenario_texto,
                'tarefa_1': proc1,
                'tarefa_2': proc2,
                'tarefa_3': proc3,
                **itens
            }

            if st.button("🚀 Gravar Estação no Banco de Dados", type="primary", use_container_width=True):
                salvar_estacao_db(usuario_atual["id"], dados_atual)
                st.success("Estação gravada com sucesso!")

    # --------------------------------------------------
    # COLUNA DIREITA: EXPORTAÇÃO EM WORD E GERENCIAMENTO
    # --------------------------------------------------
    with col_direita:
        st.markdown("**Exportação do Documento Word (.docx)**")
        try:
            docx_bytes = gerar_bytes_docx_osce(dados_atual)
            
            st.success("Documento Word gerado com sucesso!")
            
            st.download_button(
                label="📄 Baixar Arquivo Word (.docx)",
                data=docx_bytes,
                file_name=f"osce_{etapa}.docx".replace(" ", "_").lower(),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"Erro na geração do Word: {e}")

        st.markdown("<br>", unsafe_allow_html=True)

        st.info(f"**Resumo:** {componente} | {semestre} | {etapa}")

        with st.expander("📚 Gerir minhas OSCEs Guardadas / Excluir", expanded=True):
            df_minhas = carregar_estacoes_professor(usuario_atual["id"])
            if not df_minhas.empty:
                opcoes_exclusao = {f"ID #{row['id']} - {row['estacao']} ({row['componente']})": row['id'] for _, row in df_minhas.iterrows()}
                
                col_sel, col_btn = st.columns([2.5, 1])
                escolha = col_sel.selectbox("Selecionar OSCE:", list(opcoes_exclusao.keys()))
                
                with col_btn:
                    st.write(" ")
                    st.write(" ")
                    if st.button("❌ Eliminar", type="primary"):
                        id_del = opcoes_exclusao[escolha]
                        if excluir_estacao_db(id_del, usuario_atual["id"]):
                            st.success("Eliminada!")
                            st.rerun()
            else:
                st.caption("Nenhuma OSCE salva no seu perfil.")
