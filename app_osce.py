import sys
import os
import re
import io
import sqlite3
import hashlib
import pandas as pd
import streamlit as st

import fitz  # PyMuPDF

from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER

import docx
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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
# 2. GERADOR DE PDF (E PARA PRÉ-VISUALIZAÇÃO)
# ==========================================

def processar_item_texto(texto):
    if not isinstance(texto, str) or pd.isna(texto) or texto.strip() == "":
        return ""
    if '|' in texto:
        partes = texto.split('|', 1)
        rubrica = partes[0].strip()
        fala_ator = partes[1].strip()
        if fala_ator and not fala_ator.startswith('"') and not fala_ator.startswith('“'):
            fala_ator = f'"{fala_ator}"'
        return f'<font color="#D00000"><i>{rubrica}</i></font> <font color="#000000">{fala_ator}</font>'

    match_aspa = re.search(r'["“]', texto)
    if match_aspa:
        idx = match_aspa.start()
        parte_anterior = texto[:idx]
        fala_ator = texto[idx:]
        if 'Estudante:' in parte_anterior:
            partes_rubrica = parte_anterior.split('Estudante:', 1)
            resultado = f'{partes_rubrica[0]}<font color="#D00000"><i>Estudante:{partes_rubrica[1]}</i></font> <font color="#000000">{fala_ator}</font>'
        else:
            resultado = f'{parte_anterior}<font color="#000000">{fala_ator}</font>'
    else:
        if 'Estudante:' in texto:
            partes_rubrica = texto.split('Estudante:', 1)
            resultado = f'{partes_rubrica[0]}<font color="#D00000"><i>Estudante:{partes_rubrica[1]}</i></font>'
        else:
            resultado = texto
    return resultado

def gerar_bytes_pdf_osce(dados_estacao):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        leftMargin=1*cm, rightMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm
    )
    styles = getSampleStyleSheet()
    LARGURA_UTIL = 785.19

    style_header_title = ParagraphStyle('HeaderTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, alignment=TA_CENTER)
    style_header_sub = ParagraphStyle('HeaderSub', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, alignment=TA_CENTER)
    style_campos = ParagraphStyle('Campos', parent=styles['Normal'], fontName='Helvetica', fontSize=9)
    style_titulos_secao = ParagraphStyle('TitulosSecao', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9.5)
    style_corpo = ParagraphStyle('Corpo', parent=styles['Normal'], fontName='Helvetica', fontSize=8.5, leading=11)
    style_tabela_head = ParagraphStyle('TabelaHead', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8.5, alignment=TA_CENTER)
    style_tabela_item = ParagraphStyle('TabelaItem', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10)

    story = []

    semestre = str(dados_estacao.get('semestre', ''))
    estacao = str(dados_estacao.get('ESTAÇÃO', dados_estacao.get('estacao', '')))
    componente = str(dados_estacao.get('Componente', dados_estacao.get('componente', '')))
    titulo_estacao = f"OSCE {semestre} – {estacao} ({componente})" if semestre else str(dados_estacao.get('titulo_estacao', 'OSCE - ESTAÇÃO'))
    data_prova = str(dados_estacao.get('data', ''))
    prof_cadastrado = str(dados_estacao.get('Professor', dados_estacao.get('professor_nome', '')))
    if prof_cadastrado == 'nan': prof_cadastrado = ""

    story.append(Paragraph(f"<b>{titulo_estacao}</b>", style_header_title))
    story.append(Paragraph(data_prova, style_header_sub))
    story.append(Spacer(1, 4))

    tabela_identificacao = Table(
        [[
            Paragraph("Nome do(a) estudante: ____________________________________________________________________", style_campos),
            Paragraph(f"Professor(a): <u>{prof_cadastrado:<35}</u>" if prof_cadastrado else "Professor(a): __________________________________", style_campos)
        ]], colWidths=[480, 305.19]
    )
    tabela_identificacao.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(tabela_identificacao)
    story.append(Spacer(1, 4))

    cenario_texto = str(dados_estacao.get('cenario_texto', ''))
    tarefa_1 = str(dados_estacao.get('tarefa_1', ''))
    tarefa_2 = str(dados_estacao.get('tarefa_2', ''))
    tarefa_3 = str(dados_estacao.get('tarefa_3', ''))

    texto_tarefas = (
        "<b>Nos próximos 4 minutos, deverão ser realizadas as seguintes TAREFAS:</b><br/>"
        f"1. {tarefa_1}<br/>"
        f"2. {tarefa_2}<br/>"
        f"3. {tarefa_3}"
    )

    conteudo_cenario = [
        [Paragraph("<b>CENÁRIO</b>", style_titulos_secao)],
        [Paragraph(cenario_texto, style_corpo)],
        [Paragraph(texto_tarefas, style_corpo)]
    ]

    tabela_cenario = Table(conteudo_cenario, colWidths=[LARGURA_UTIL])
    tabela_cenario.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('LINEBELOW', (0,0), (-1,0), 0.5, colors.black),
        ('LINEBELOW', (0,1), (-1,1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#E6E6E6')),
        ('TOPPADDING', (0,0), (-1,-1), 3), ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5), ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(tabela_cenario)
    story.append(
