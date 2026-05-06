import streamlit as st
import pandas as pd
import json
import os
import plotly.express as px
from datetime import datetime
from docx import Document
import io
import re
import fitz  # PyMuPDF

# Configuração da página
st.set_page_config(page_title="Aura - Diagnóstico Pedagógico", page_icon="✨", layout="wide")

# Senhas
SENHA_ESCOLA = "Aura2024"
SENHA_ADM = "adm2024aura"
VERSAO_APP = "Aura 1.0 - Oficial"

# Caminhos dos arquivos
QUESTIONS_FILE = 'questoes.json'
RESPONSES_FILE = 'respostas.csv'
PROVAS_DIR = 'provas'
MEDIA_DIR = 'media'
for d in [PROVAS_DIR, MEDIA_DIR]: os.makedirs(d, exist_ok=True)

# Tenta carregar conexão com Google Sheets (apenas se configurado nos Secrets)
try:
    from streamlit_gsheets import GSheetsConnection
    conn = st.connection("gsheets", type=GSheetsConnection)
    USE_CLOUD = True
except:
    USE_CLOUD = False

def load_questions():
    if os.path.exists(QUESTIONS_FILE):
        with open(QUESTIONS_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    return {"Português": [], "Matemática": []}

def save_questions(data):
    with open(QUESTIONS_FILE, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=4)

def save_response(nome, turma, disciplina, descritor, resultado):
    data = {"Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Nome": nome, "Turma": turma, "Disciplina": disciplina, "Descritor": descritor, "Resultado": resultado}
    df_new = pd.DataFrame([data])
    
    # Salva localmente
    if os.path.exists(RESPONSES_FILE): df_new.to_csv(RESPONSES_FILE, mode='a', header=False, index=False)
    else: df_new.to_csv(RESPONSES_FILE, index=False)
    
    # Se estiver na nuvem e configurado, tenta salvar no Google Sheets também
    if USE_CLOUD:
        try:
            existing_data = conn.read(worksheet="Respostas")
            updated_df = pd.concat([existing_data, df_new], ignore_index=True)
            conn.update(worksheet="Respostas", data=updated_df)
        except: pass

def extract_text_and_images_pdf(file_path):
    doc = fitz.open(file_path)
    text = ""
    images = []
    
    base_name = os.path.basename(file_path).split('.')[0]
    file_media_dir = os.path.join(MEDIA_DIR, base_name)
    if not os.path.exists(file_media_dir): os.makedirs(file_media_dir)

    for i, page in enumerate(doc):
        # Usar blocos ajuda a manter a ordem de leitura correta (especialmente em colunas)
        blocks = page.get_text("blocks")
        for b in blocks:
            text += b[4] + "\n"
            
    # Removendo cabeçalho (primeiras linhas de metadados da escola)
    lines = text.split('\n')
    if len(lines) > 2:
        # Se as primeiras linhas contiverem palavras-chave de cabeçalho, removemos
        if any(kw in lines[0].upper() for kw in ["ESCOLA", "EREM", "DISCIPLINA", "PROFESSOR", "DATA"]):
            text = '\n'.join(lines[1:])
        if any(kw in lines[0].upper() for kw in ["ESCOLA", "EREM", "DISCIPLINA", "PROFESSOR", "DATA"]):
             text = '\n'.join(lines[1:]) # Faz de novo para garantir
             
    for i, page in enumerate(doc):
        for img_index, img in enumerate(page.get_images(full=True)):
            xref = img[0]
            try:
                base_image = doc.extract_image(xref)
                image_bytes = base_image["image"]
                ext = base_image["ext"]
                img_name = f"page{i+1}_img{img_index+1}.{ext}"
                img_path = os.path.join(file_media_dir, img_name)
                with open(img_path, "wb") as f:
                    f.write(image_bytes)
                images.append(img_path)
            except: continue
    return text, images

def extract_text_and_images_docx(file_path):
    doc = Document(file_path)
    text = ""
    images = []
    
    base_name = os.path.basename(file_path).split('.')[0]
    file_media_dir = os.path.join(MEDIA_DIR, base_name)
    if not os.path.exists(file_media_dir): os.makedirs(file_media_dir)

    for para in doc.paragraphs:
        text += para.text + "\n"
    
    # Removendo cabeçalho
    lines = text.split('\n')
    if len(lines) > 2:
        if any(kw in lines[0].upper() for kw in ["ESCOLA", "EREM", "DISCIPLINA", "PROFESSOR", "DATA"]):
            text = '\n'.join(lines[1:])

    # Extração de imagens em DOCX
    count = 1
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_data = rel.target_part.blob
            img_name = f"img_{count}.png"
            img_path = os.path.join(file_media_dir, img_name)
            with open(img_path, "wb") as f:
                f.write(img_data)
            images.append(img_path)
            count += 1
            
    return text, images

def get_pedagogical_advice(descritor):
    # Banco de dados com ações diretas e objetivas
    advices = {
        # PORTUGUÊS
        "D1": {"diag": "Dificuldade em localizar informações explícitas no texto.", "acao": "Pedir que o aluno sublinhe a resposta diretamente no texto antes de marcar."},
        "D2": {"diag": "Dificuldade em identificar a quem se referem pronomes como 'ele', 'ela' ou 'isso'.", "acao": "Realizar exercícios de ligar pronomes aos seus substantivos correspondentes no texto."},
        "D3": {"diag": "Dificuldade em descobrir o sentido de palavras pelo contexto.", "acao": "Praticar a substituição de palavras difíceis por sinônimos que mantenham o sentido."},
        "D4": {"diag": "Dificuldade em entender informações que não estão escritas, mas sugeridas.", "acao": "Interpretar tirinhas e charges focando no que a imagem sugere além das palavras."},
        "D5": {"diag": "Dificuldade em ler textos com imagens e palavras juntas.", "acao": "Analisar como a imagem ajuda a explicar a frase principal em anúncios."},
        "D6": {"diag": "Dificuldade em identificar o assunto principal do texto.", "acao": "Solicitar que o aluno crie um novo título que resuma o texto todo."},
        "D14": {"diag": "Dificuldade em separar o que aconteceu (fato) do que alguém pensa (opinião).", "acao": "Listar frases do texto e classificar em duas colunas: Fato ou Opinião."},
        "D18": {"diag": "Dificuldade em entender o motivo de se usar certas pontuações ou palavras.", "acao": "Trocar palavras do texto por outras e observar como o sentimento da frase muda."},
        
        # MATEMÁTICA
        "D13": {"diag": "Dificuldade em resolver problemas de somar, subtrair, multiplicar ou dividir.", "acao": "Simular situações de compras em mercado usando preços reais e troco."},
        "D19": {"diag": "Dificuldade em calcular a área (espaço interno) de figuras.", "acao": "Usar papel quadriculado para contar os quadradinhos dentro de cada figura plana."},
        "D21": {"diag": "Dificuldade em identificar retas paralelas ou que se cruzam.", "acao": "Identificar no mapa do bairro ruas que nunca se cruzam e ruas que se cruzam."},
        "D32": {"diag": "Dificuldade em descobrir a regra de uma sequência numérica.", "acao": "Apresentar sequências de desenhos e pedir que o aluno descubra o próximo elemento."}
    }
    
    code = descritor.split(' - ')[0].upper().strip()
    return advices.get(code, {"diag": "Necessita de revisão do conceito base deste descritor.", "acao": "Revisar o conceito em sala com um exemplo prático do dia a dia."})

def parse_questions_comprehensive(text, available_images=[]):
    gabarito_map = {}
    gabarito_match = re.search(r'(?:GABARITO|RESPOSTAS).*?\n(.*)', text, re.IGNORECASE | re.DOTALL)
    if gabarito_match:
        g_text = gabarito_match.group(1)
        g_entries = re.findall(r'(\d+)\s*[\-\.\s]\s*([A-E])', g_text, re.IGNORECASE)
        for q_num, ans in g_entries: gabarito_map[q_num] = ans.upper()

    # Regex para início de questão (mais criterioso)
    q_start_regex = r'(?:\n|^)[\s]*[\(]?(\d+|[IVXLC]+)[\.\)]\s+|(?:\n|^)[\s]*(Questão\s*\d+)[:\.\s\-]*'
    parts = re.split(q_start_regex, text)
    
    raw_blocks = []
    # Ignoramos parts[0] (cabeçalho) por solicitação do usuário
    
    # Organizamos em blocos de [número, conteúdo]
    for i in range(1, len(parts), 3):
        num = parts[i] if parts[i] else parts[i+1]
        content = parts[i+2] if i+2 < len(parts) else ""
        raw_blocks.append({"num": num, "content": content})

    if not raw_blocks: return []

    final_questions = []
    for block in raw_blocks:
        content = block["content"]
        
        # Procura por alternativas A, B, C, D, E
        # Exigimos que a alternativa esteja no início de uma linha ou após espaço, para evitar confusão no meio de frases
        opt_pattern = re.compile(r'(?:\n|^)[\s]*([A-Ea-e])[\s]*[\)\.\-\]]\s+', re.MULTILINE)
        all_matches = list(opt_pattern.finditer(content))
        
        # Só começamos a contar alternativas a partir da primeira letra 'A' ou 'a' encontrada
        opt_matches = []
        found_a = False
        for m in all_matches:
            if m.group(1).upper() == 'A': found_a = True
            if found_a: opt_matches.append(m)
        
        # Se NÃO encontrou alternativas, anexa este bloco inteiro ao enunciado da questão anterior
        if not opt_matches:
            if final_questions:
                final_questions[-1]["pergunta"] += f"\n{block['num']}. {content}"
            continue
            
        # Se encontrou alternativas:
        # O enunciado é TUDO desde o início do bloco até a posição da PRIMEIRA alternativa 'A'
        first_opt_pos = opt_matches[0].start()
        pergunta = content[:first_opt_pos].strip()
        
        # Identificação de descritor (sem apagar o texto importante)
        desc_match = re.search(r'(D\d+|Descritor\s+\d+)', content, re.IGNORECASE)
        desc = desc_match.group(1).upper() if desc_match else "D? - Identificar"
        
        # Extração das opções
        opts_list = []
        for j in range(len(opt_matches)):
            start = opt_matches[j].start()
            end = opt_matches[j+1].start() if j+1 < len(opt_matches) else len(content)
            opt_text = content[start:end].strip()
            if opt_text: opts_list.append(opt_text)
            
        q_num_clean = re.sub(r'\D', '', block["num"])
        correta_sugestao = opts_list[0] if opts_list else ""
        # Tentativa de usar gabarito se existir
        if q_num_clean in gabarito_map:
            letra_g = gabarito_map[q_num_clean]
            for o in opts_list:
                if o.upper().startswith(letra_g):
                    correta_sugestao = o
                    break
        
        final_questions.append({
            "num": block["num"],
            "descritor": desc,
            "pergunta": pergunta,
            "opcoes": opts_list,
            "correta": correta_sugestao,
            "explicacao": "Análise automática.",
            "imagem": None
        })
        
    return final_questions

# CSS
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;} header {visibility: hidden;} footer {visibility: hidden;}
    .stAppDeployButton {display:none;}
    .block-container { padding-top: 1rem !important; }
    .question-card { background-color: #ffffff; padding: 20px 0; border-bottom: 2px solid #e2e8f0; margin-bottom: 35px; text-align: justify; text-justify: inter-word; color: #1a202c; }
    .descriptor-text { color: #2563eb; font-weight: bold; font-size: 1rem; margin-bottom: 10px; display: block; }
    .enunciado-text { font-size: 1.2rem; line-height: 1.7; color: #1a202c; font-weight: 500; display: inline; }
    .question-title { font-weight: 900; font-size: 1.3rem; display: inline; margin-right: 12px; color: #000000; }
    .stImage > img { border-radius: 12px; margin: 20px auto; border: 1px solid #e2e8f0; display: block; max-width: 450px !important; width: 100%; height: auto; }
    hr { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

# Estados
if 'auth_escola' not in st.session_state: st.session_state.auth_escola = False
if 'perfil' not in st.session_state: st.session_state.perfil = None
if 'respostas_aluno' not in st.session_state: st.session_state.respostas_aluno = {}

# --- LOGIN ---
if not st.session_state.auth_escola:
    st.title("✨ Aura - Diagnóstico Pedagógico")
    if st.text_input("Senha de Acesso:", type="password") == SENHA_ESCOLA:
        if st.button("Entrar"): st.session_state.auth_escola = True; st.rerun()
    st.stop()

if st.session_state.perfil is None:
    st.title("Acesso:")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🙋‍♂️ ALUNO", use_container_width=True): st.session_state.perfil = 'aluno'; st.rerun()
    with c2:
        if st.button("🔑 PROFESSOR", use_container_width=True): st.session_state.perfil = 'adm'; st.rerun()
    st.stop()

# --- ALUNO ---
if st.session_state.perfil == 'aluno':
    if 'aluno_nome' not in st.session_state:
        st.title("✍️ Identificação")
        nome = st.text_input("Nome:")
        turma = st.selectbox("Turma:", ["3º Ano A", "3º Ano B", "9º Ano A", "9º Ano B"])
        if st.button("Começar"):
            if nome: st.session_state.aluno_nome = nome; st.session_state.aluno_turma = turma; st.rerun()
        if st.button("Voltar"): st.session_state.perfil = None; st.rerun()
        st.stop()
    
    st.sidebar.write(f"Aluno: {st.session_state.aluno_nome}")
    if st.sidebar.button("Sair"): st.session_state.perfil = None; st.rerun()
    questoes = load_questions()
    t1, t2 = st.tabs(["📚 Português", "📐 Matemática"])

    def quiz(disc):
        if disc in questoes:
            for i, q in enumerate(questoes[disc]):
                key = f"{disc}_{i}"
                # Limpa quebras de linha simples para permitir justificativa, mas mantém parágrafos
                p_limpa = q["pergunta"].replace("\n\n", "[[PAR]]").replace("\n", " ").replace("[[PAR]]", "<br><br>")
                st.markdown(f'<div class="question-card"><span class="descriptor-text">Descritor: {q["descritor"]}</span> <span class="question-title">Questão {i+1}:</span> <span class="enunciado-text">{p_limpa}</span>', unsafe_allow_html=True)
                if q.get('imagem'):
                    st.image(q['imagem'])
                st.markdown('</div>', unsafe_allow_html=True)
                resp = key in st.session_state.respostas_aluno
                escolha = st.radio("Alt:", q['opcoes'], index=None, key=f"r_{key}", disabled=resp, label_visibility="collapsed")
                if not resp:
                    if st.button(f"Confirmar Resposta {i+1}", key=f"b_{key}"):
                        if escolha:
                            res = "Acertou" if escolha == q['correta'] else "Errou"
                            st.session_state.respostas_aluno[key] = res
                            save_response(st.session_state.aluno_nome, st.session_state.aluno_turma, disc, q['descritor'], res)
                            st.rerun()
                else:
                    st.info("✔️ Resposta registrada com sucesso.")
                st.markdown('</div>', unsafe_allow_html=True)
    with t1: quiz("Português")
    with t2: quiz("Matemática")

# --- ADMIN ---
elif st.session_state.perfil == 'adm':
    if 'adm_auth' not in st.session_state:
        st.title("🔑 Painel ADM")
        if st.text_input("Senha ADM:", type="password") == SENHA_ADM:
            if st.button("Acessar"): st.session_state.adm_auth = True; st.rerun()
        st.stop()
    st.title("📊 Painel Administrativo")
    if st.sidebar.button("Sair"): st.session_state.perfil = None; st.rerun()
    
    tab_g, tab_p, tab_m, tab_est, tab_quest = st.tabs(["🌎 Visão Geral", "📚 Português", "📐 Matemática", "👤 Alunos Detalhado", "⚙️ Gestão de Provas"])
    
    df = pd.DataFrame()
    if os.path.exists(RESPONSES_FILE):
        df = pd.read_csv(RESPONSES_FILE)
    
    with tab_g:
        if not df.empty:
            df['Descritor'] = df['Descritor'].str.strip()
            df['Cod_Descritor'] = df['Descritor'].str.split(' - ').str[0]
            turma_sel = st.selectbox("Filtrar Turma:", ["Todas"] + sorted(df['Turma'].unique().tolist()), key="sel_turma_g")
            df_b = df if turma_sel == "Todas" else df[df['Turma'] == turma_sel]
            stats_t = df.groupby('Turma').agg(Acertos=('Resultado', lambda x: (x == 'Acertou').sum()), Total=('Resultado', 'count')).reset_index()
            stats_t['%'] = (stats_t['Acertos'] / stats_t['Total'] * 100).round(1)
            st.plotly_chart(px.bar(stats_t, x='Turma', y='%', text='%', color='Turma', title="Desempenho por Turma"), use_container_width=True)
            st.markdown("---")
            st.subheader("🌍 Diagnóstico Global e Plano de Ação Coletivo")
            erros_globais = df_b[df_b['Resultado'] == 'Errou']
            if not erros_globais.empty:
                top_erros = erros_globais.groupby('Descritor').size().sort_values(ascending=False).head(3)
                for desc, count in top_erros.items():
                    advice = get_pedagogical_advice(desc)
                    with st.expander(f"🚩 {desc} - {count} erros detectados"):
                        st.write(f"**Diagnóstico:** {advice['diag']}")
                        st.info(f"**💡 Ação Sugerida:** {advice['acao']}")
            else: st.success("Desempenho Global Excelente!")
            st.markdown("---")
            aluno_res = df_b.groupby(['Nome', 'Turma']).agg(Total=('Resultado', 'count'), Acertos=('Resultado', lambda x: (x == 'Acertou').sum())).reset_index()
            aluno_res['%'] = (aluno_res['Acertos'] / aluno_res['Total'] * 100).round(1)
            st.subheader("Ranking de Alunos")
            st.dataframe(aluno_res.sort_values(by='%', ascending=False), use_container_width=True)
        else:
            st.info("Aguardando as primeiras respostas dos alunos para gerar estatísticas.")

    def render_m(mat, tab_obj):
        with tab_obj:
            if not df.empty:
                df['Descritor'] = df['Descritor'].str.strip()
                df['Cod_Descritor'] = df['Descritor'].str.split(' - ').str[0]
                df_m = df[df['Disciplina'] == mat]
                if not df_m.empty:
                    desc_df = df_m.groupby(['Cod_Descritor', 'Resultado']).size().reset_index(name='Qtd')
                    fig = px.bar(desc_df, x='Cod_Descritor', y='Qtd', color='Resultado', barmode='group', color_discrete_map={'Acertou': '#10b981', 'Errou': '#ef4444'}, title=f"Desempenho por Descritor ({mat})")
                    st.plotly_chart(fig, use_container_width=True)
                else: st.info(f"Sem dados de {mat}.")
            else: st.info("Sem dados para exibir.")

    render_m("Português", tab_p)
    render_m("Matemática", tab_m)

    with tab_est:
        if not df.empty:
            st.subheader("Análise Individual por Estudante")
            alunos_lista = sorted(df['Nome'].unique().tolist())
            aluno_sel = st.selectbox("Selecione o Estudante:", alunos_lista)
            if aluno_sel:
                df_aluno = df[df['Nome'] == aluno_sel]
                col1, col2 = st.columns(2)
                with col1:
                    total_q = len(df_aluno); acertos_q = (df_aluno['Resultado'] == 'Acertou').sum()
                    st.metric("Total de Questões", total_q); st.metric("Acertos", acertos_q, f"{(acertos_q/total_q*100):.1f}%")
                with col2:
                    erros = df_aluno[df_aluno['Resultado'] == 'Errou']
                    if not erros.empty:
                        st.warning("⚠️ Maiores Dificuldades:"); dif_count = erros.groupby('Descritor').size().sort_values(ascending=False)
                        for desc, count in dif_count.items(): st.write(f"- **{desc}**: {count} erro(s)")
                    else: st.success("100% de Aproveitamento!")
                st.markdown("---")
                st.subheader("📋 Diagnóstico e Plano de Ação")
                if not erros.empty:
                    pior_desc = erros.groupby('Descritor').size().idxmax(); advice = get_pedagogical_advice(pior_desc)
                    st.error(f"**Diagnóstico:** {advice['diag']}"); st.info(f"**💡 Ação Sugerida:** {advice['acao']}")
                st.write("**Desempenho por Descritor:**")
                df_aluno['Cod_Descritor'] = df_aluno['Descritor'].str.split(' - ').str[0]
                indiv_stats = df_aluno.groupby(['Cod_Descritor', 'Resultado']).size().reset_index(name='Qtd')
                st.plotly_chart(px.bar(indiv_stats, x='Cod_Descritor', y='Qtd', color='Resultado', barmode='group', color_discrete_map={'Acertou': '#10b981', 'Errou': '#ef4444'}), use_container_width=True)
        else: st.info("Nenhum aluno respondeu ainda.")

        st.markdown("---")
        st.subheader("💾 Backup e Segurança")
        c_b1, c_b2 = st.columns(2)
        with c_b1:
            with open(QUESTIONS_FILE, "rb") as f:
                st.download_button("📥 Baixar Banco de Questões", f, file_name="questoes_backup.json")
        with c_b2:
            if os.path.exists(RESPONSES_FILE):
                with open(RESPONSES_FILE, "rb") as f:
                    st.download_button("📥 Baixar Resultados (CSV)", f, file_name="respostas_backup.csv")
        
        st.write("---")
        st.subheader("📁 Gestão de Arquivos de Provas")
        uploaded_file = st.file_uploader("Anexar Prova (PDF ou DOCX)", type=["pdf", "docx"])
        if uploaded_file:
            save_path = os.path.join(PROVAS_DIR, uploaded_file.name)
            with open(save_path, "wb") as f: f.write(uploaded_file.getbuffer())
            st.success(f"Arquivo '{uploaded_file.name}' salvo!")
        
        arquivos = os.listdir(PROVAS_DIR)
        if arquivos:
            for arq in arquivos:
                col_a, col_b, col_c = st.columns([0.6, 0.2, 0.2])
                col_a.write(f"📄 {arq}")
                if col_b.button("Ler e Analisar (Completo)", key=f"read_{arq}"):
                    file_path = os.path.join(PROVAS_DIR, arq)
                    ext = arq.split('.')[-1].lower()
                    if ext == 'pdf': text, imgs = extract_text_and_images_pdf(file_path)
                    else: text, imgs = extract_text_and_images_docx(file_path)
                    st.session_state.extracted_text_draft = text; st.session_state.extracted_images = imgs
                    st.session_state.parsed_drafts = parse_questions_comprehensive(text, imgs)
                    st.info(f"Detectadas {len(st.session_state.parsed_drafts)} questões e {len(imgs)} imagens.")
                if col_c.button("Excluir", key=f"del_file_{arq}"): os.remove(os.path.join(PROVAS_DIR, arq)); st.rerun()
        
        if 'parsed_drafts' in st.session_state and st.session_state.parsed_drafts:
            st.markdown("---")
            st.subheader("📋 Revisão de Importação Automática")
            disc_import = st.selectbox("Importar para qual disciplina?", ["Português", "Matemática"])
            for i, q in enumerate(st.session_state.parsed_drafts):
                with st.expander(f"Questão {q['num']}: {q['pergunta'][:50]}..."):
                    q['descritor'] = st.text_input(f"Descritor Q{i}", q['descritor'], key=f"d_imp_{i}")
                    q['pergunta'] = st.text_area(f"Enunciado Q{i}", q['pergunta'], key=f"p_imp_{i}")
                    if st.session_state.get('extracted_images'):
                        img_options = ["Nenhuma"] + st.session_state.extracted_images
                        sel_img = st.selectbox(f"Anexar Imagem Q{i}", img_options, key=f"img_sel_{i}")
                        q['imagem'] = None if sel_img == "Nenhuma" else sel_img
                        if q['imagem']: st.image(q['imagem'], width=200)
                    q['opcoes'] = st.text_area(f"Opções Q{i}", "\n".join(q['opcoes']), key=f"o_imp_{i}").split('\n')
                    q['correta'] = st.selectbox(f"Resposta Correta Q{i}", q['opcoes'], index=q['opcoes'].index(q['correta']) if q['correta'] in q['opcoes'] else 0, key=f"c_imp_{i}")
            if st.button("✅ Importar Todas as Questões"):
                questoes = load_questions()
                if disc_import not in questoes: questoes[disc_import] = []
                for q in st.session_state.parsed_drafts:
                    questoes[disc_import].append({"descritor": q['descritor'], "pergunta": q['pergunta'], "opcoes": [o.strip() for o in q['opcoes'] if o.strip()], "correta": q['correta'], "explicacao": q['explicacao'], "imagem": q.get('imagem')})
                save_questions(questoes); st.session_state.parsed_drafts = []; st.success("Questões importadas!"); st.rerun()

        st.markdown("---")
        st.subheader("📝 Gestão de Questões")
        questoes = load_questions()
        if 'edit_q' not in st.session_state: st.session_state.edit_q = None
        st.write("### " + ("Editar Questão" if st.session_state.edit_q else "Cadastrar Nova Questão"))
        default_disc = "Português"; default_desc = ""; default_pergunta = ""; default_opcoes = ""; default_correta = ""; default_explica = ""
        if st.session_state.edit_q:
            ed = st.session_state.edit_q; default_disc = ed['disc']; default_desc = ed['q']['descritor']; default_pergunta = ed['q']['pergunta']; default_opcoes = "\n".join(ed['q']['opcoes']); default_correta = ed['q']['correta']; default_explica = ed['q']['explicacao']
        with st.form("form_questao"):
            f_disc = st.selectbox("Disciplina", ["Português", "Matemática"], index=0 if default_disc=="Português" else 1)
            f_desc = st.text_input("Descritor", value=default_desc)
            f_pergunta = st.text_area("Pergunta", value=default_pergunta)
            f_opcoes = st.text_area("Opções (uma por linha)", value=default_opcoes)
            f_correta = st.text_input("Resposta Correta", value=default_correta)
            f_explica = st.text_area("Explicação Pedagógica", value=default_explica)
            if st.form_submit_button("Salvar"):
                nova_q = {"descritor": f_desc, "pergunta": f_pergunta, "opcoes": [o.strip() for o in f_opcoes.split('\n') if o.strip()], "correta": f_correta, "explicacao": f_explica}
                if st.session_state.edit_q:
                    old = st.session_state.edit_q; questoes[old['disc']].pop(old['idx'])
                if f_disc not in questoes: questoes[f_disc] = []
                questoes[f_disc].append(nova_q); st.session_state.edit_q = None; save_questions(questoes); st.success("Sucesso!"); st.rerun()
        if st.checkbox("Ver/Remover Questões Existentes"):
            for d in ["Português", "Matemática"]:
                if d in questoes:
                    for idx, q in enumerate(questoes[d]):
                        col_q, col_edit, col_del = st.columns([0.7, 0.15, 0.15])
                        col_q.write(f"{idx+1}. {q['descritor']} - {q['pergunta'][:50]}...")
                        if col_edit.button("✏️", key=f"edit_{d}_{idx}"): st.session_state.edit_q = {'disc': d, 'idx': idx, 'q': q}; st.rerun()
                        if col_del.button("🗑️", key=f"del_{d}_{idx}"): questoes[d].pop(idx); save_questions(questoes); st.rerun()
st.sidebar.title("✨ Aura")
st.sidebar.caption(f"Tecnologia Educacional | {VERSAO_APP}")
