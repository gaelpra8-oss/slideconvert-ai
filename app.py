import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from PIL import Image
from google import genai
import os
import re
import tempfile
import time

# ==============================================================================
# 1. CONFIGURATION DU DECK & INJECTION CSS (MASQUAGE DEPLOY & GLASSMORPHISM)
# ==============================================================================
st.set_page_config(
    page_title="SlideConvert AI - Numériseur Pédagogique Intelligent",
    page_icon="✨",
    layout="wide"
)
# Protection par mot de passe
mot_de_passe = st.text_input("🔑 Mot de passe", type="password")

if mot_de_passe != st.secrets.get("APP_PASSWORD", "Gpereira8989@&"):
    st.warning("🔒 Accès restreint. Veuillez entrer le mot de passe.")
    st.stop()

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

/* Masquage strict du bouton "Deploy" et du header Streamlit */
header, [data-testid="stHeader"], [data-testid="stAppDeployButton"], .stDeployButton, #MainMenu, footer {
    display: none !important;
    visibility: hidden !important;
}

/* Style global du fond */
html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #070913 0%, #0c0e27 45%, #140b2e 100%) !important;
    color: #e2e8f0 !important;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
}

/* Cartes Glassmorphism */
.glass-card {
    background: rgba(255, 255, 255, 0.04);
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 20px;
    padding: 24px;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5), inset 0 1px 0 rgba(255, 255, 255, 0.1);
    margin-bottom: 20px;
}

/* Style du volet dépliable (st.expander) */
div[data-testid="stExpander"] {
    background: rgba(255, 255, 255, 0.03) !important;
    backdrop-filter: blur(15px) !important;
    border: 1px solid rgba(255, 255, 255, 0.12) !important;
    border-radius: 16px !important;
    margin-bottom: 20px !important;
}
div[data-testid="stExpander"] summary {
    font-weight: 700 !important;
    color: #a5b4fc !important;
}

/* Zone Drag and Drop */
.dropzone-container {
    border: 2px dashed rgba(129, 140, 248, 0.5);
    border-radius: 20px;
    padding: 35px 20px;
    text-align: center;
    background: rgba(15, 23, 42, 0.3);
    transition: all 0.3s ease;
    box-shadow: inset 0 0 20px rgba(99, 102, 241, 0.05);
}
.dropzone-container:hover {
    border-color: #818cf8;
    background: rgba(30, 27, 75, 0.3);
    box-shadow: 0 0 25px rgba(99, 102, 241, 0.25);
}

/* Card fichier chargé */
.file-card {
    display: flex;
    align-items: center;
    justify-content: space-between;
    background: rgba(30, 41, 59, 0.6);
    border: 1px solid rgba(148, 163, 184, 0.2);
    border-radius: 14px;
    padding: 14px 20px;
    margin-top: 15px;
}

/* Bouton CTA violet éclatant */
.stButton > button {
    width: 100% !important;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #d946ef 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 16px !important;
    padding: 18px 28px !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    letter-spacing: 0.5px !important;
    box-shadow: 0 0 25px rgba(139, 92, 246, 0.5) !important;
    transition: all 0.3s ease !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 0 35px rgba(217, 70, 239, 0.7) !important;
}

/* Carte Succès Vert Émeraude */
.success-card {
    background: rgba(6, 78, 59, 0.4);
    border: 1px solid rgba(52, 211, 153, 0.6);
    border-radius: 20px;
    padding: 24px;
    text-align: center;
    box-shadow: 0 0 30px rgba(16, 185, 129, 0.25);
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

API_KEY = os.environ.get("GEMINI_API_KEY", "")

# ==============================================================================
# 2. EN-TÊTE DE L'APPLICATION (TITRE SLIDECONVERT AI)
# ==============================================================================
st.markdown("""
    <div style="display: flex; align-items: center; gap: 14px; margin-bottom: 24px;">
        <div style="background: linear-gradient(135deg, #6366f1, #ec4899); padding: 10px 14px; border-radius: 14px; font-size: 24px; box-shadow: 0 0 20px rgba(236, 72, 153, 0.4);">
            ✨
        </div>
        <div>
            <h1 style="margin: 0; font-size: 32px; font-weight: 800; background: linear-gradient(135deg, #ffffff 0%, #a5b4fc 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">
                SlideConvert <span style="background: linear-gradient(135deg, #34d399, #10b981); -webkit-background-clip: text; -webkit-text-fill-color: transparent;">AI</span>
            </h1>
            <p style="margin: 4px 0 0 0; color: #94a3b8; font-size: 14px; font-weight: 500;">
                Numériseur pédagogique intelligent : Vos diapositives PDF en cours Word rédigés
            </p>
        </div>
    </div>
""", unsafe_allow_html=True)

# ==============================================================================
# 3. MENU PARAMÈTRES (VOLET DÉPLIABLE)
# ==============================================================================
with st.expander("⚙️ PARAMÈTRES & STYLE DE MISE EN FORME", expanded=False):
    col_p1, col_p2, col_p3 = st.columns(3)

    with col_p1:
        st.markdown("<p style='font-size:13px; font-weight:600; color:#cbd5e1; margin-bottom:8px;'>Modèle d'IA</p>", unsafe_allow_html=True)
        modele_option = st.radio(
            "Modèle IA",
            options=["Gemini 1.5 Pro (Haute précision)", "Gemini 1.5 Flash (Rapide)"],
            index=0,
            label_visibility="collapsed"
        )
        selected_model_code = "gemini-1.5-pro" if "Pro" in modele_option else "gemini-3.5-flash-lite"

    with col_p2:
        st.markdown("<p style='font-size:13px; font-weight:600; color:#cbd5e1; margin-bottom:8px;'>Style du document Word</p>", unsafe_allow_html=True)
        style_option = st.selectbox(
            "Style de document",
            options=[
                "Bleu Marine & Encadrés Pédagogiques",
                "Émeraude & Modern Minimalist",
                "Sombre Élégant & Minimal"
            ],
            index=0,
            label_visibility="collapsed"
        )

    with col_p3:
        st.markdown("<p style='font-size:13px; font-weight:600; color:#cbd5e1; margin-bottom:8px;'>Options de traitement</p>", unsafe_allow_html=True)
        opt_latex = st.toggle("Conversion LaTeX (ex: 1 m³ = 330 kg)", value=True)
        opt_decor = st.toggle("Suppression du décor (logos, en-têtes)", value=True)

# Palettes de couleurs appliquées au document Word
PALETTES = {
    "Bleu Marine & Encadrés Pédagogiques": {
        "H2": RGBColor(0, 51, 102),
        "H3": RGBColor(0, 102, 153),
        "H4": RGBColor(51, 102, 153)
    },
    "Émeraude & Modern Minimalist": {
        "H2": RGBColor(6, 78, 59),
        "H3": RGBColor(16, 185, 129),
        "H4": RGBColor(52, 211, 153)
    },
    "Sombre Élégant & Minimal": {
        "H2": RGBColor(30, 41, 59),
        "H3": RGBColor(71, 85, 105),
        "H4": RGBColor(100, 116, 139)
    }
}

# ==============================================================================
# 4. FONCTIONS DU MOTEUR (ROGNAGE & FORMATAGE DOCX)
# ==============================================================================

def rogner_schema(image_path, coordonnees_norm):
    try:
        img = Image.open(image_path)
        width, height = img.size
        ymin, xmin, ymax, xmax = coordonnees_norm

        left = int((xmin / 1000.0) * width)
        top = int((ymin / 1000.0) * height)
        right = int((xmax / 1000.0) * width)
        bottom = int((ymax / 1000.0) * height)

        pad_w = int(width * 0.03)
        pad_h = int(height * 0.03)
        left = max(0, left - pad_w)
        top = max(0, top - pad_h)
        right = min(width, right + pad_w)
        bottom = min(height, bottom + pad_h)

        img_croppee = img.crop((left, top, right, bottom))
        path_crop = image_path.replace(".png", "_crop.png")
        img_croppee.save(path_crop)
        return path_crop
    except Exception:
        return image_path

def configurer_page(doc, top=0.8, bottom=0.8, left=0.8, right=0.8):
    for section in doc.sections:
        section.top_margin = Inches(top)
        section.bottom_margin = Inches(bottom)
        section.left_margin = Inches(left)
        section.right_margin = Inches(right)

def parse_and_add_formatted_runs(paragraph, text):
    pattern = r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|`.*?`)'
    tokens = re.split(pattern, text)

    for token in tokens:
        if not token:
            continue
        if token.startswith('***') and token.endswith('***') and len(token) > 6:
            run = paragraph.add_run(token[3:-3])
            run.bold = True
            run.italic = True
        elif token.startswith('**') and token.endswith('**') and len(token) > 4:
            run = paragraph.add_run(token[2:-2])
            run.bold = True
        elif token.startswith('*') and token.endswith('*') and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
        elif token.startswith('`') and token.endswith('`') and len(token) > 2:
            run = paragraph.add_run(token[1:-1])
            run.font.name = 'Consolas'
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(180, 40, 40)
        else:
            paragraph.add_run(token)

def ajouter_titre(doc, text, level, color_rgb):
    h = doc.add_heading(level=level)
    h.paragraph_format.space_before = Pt(14 if level <= 2 else 10)
    h.paragraph_format.space_after = Pt(4)
    h.paragraph_format.keep_with_next = True
    text_clean = re.sub(r'^#+\s*', '', text).strip()
    parse_and_add_formatted_runs(h, text_clean)
    for run in h.runs:
        run.font.color.rgb = color_rgb

def ajouter_encadre_callout(doc, text):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Inches(6.5)
    cell = table.cell(0, 0)

    shading_xml = parse_xml(r'<w:shd {} w:fill="F0F4F8"/>'.format(nsdecls('w')))
    cell._tc.get_or_add_tcPr().append(shading_xml)

    borders_xml = parse_xml(
        r'<w:tcBorders {}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="003366"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>'.format(nsdecls('w'))
    )
    cell._tc.get_or_add_tcPr().append(borders_xml)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.15)
    p.paragraph_format.right_indent = Inches(0.15)

    parse_and_add_formatted_runs(p, text)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(0)
    p_spacer.paragraph_format.space_after = Pt(4)

PROMPT_EXPERT = """
Tu es un ingénieur pédagogique d'élite, professeur agrégé et expert reconnu dans la numérisation, la restructuration et l'optimisation didactique de supports de cours universitaires.
Ton rôle est d'analyser l'image de cette diapositive de cours et de la convertir en un chapitre de cours rédigé, parfaitement structuré, hautement synthétique, professionnel et prêt à être converti en un document Microsoft Word (.docx).

======================================================================
1. RÈGLES D'ÉLIMINATION STRICTE DU BRUIT ET DES MÉTADONNÉES
======================================================================
- SUPPRESSION DU DÉCOR INSTITUTIONNEL : Ignore et élimine TOUTES les mentions récurrentes, filigranes, logos, en-têtes, pieds de page et masques de diapositives.
- ZÉRO VERBOSITÉ SYSTÈME : Ne rédige AUCUNE phrase d'introduction, AUCUNE conclusion et AUCUN bloc de code Markdown global.
- AUCUN SÉPARATEUR PARASITE : Ne génère aucune ligne horizontale de séparation de type `---` dans le corps du texte.

======================================================================
2. DIAGNOSTIC ET DÉTECTION VISION (OBLIGATOIRE EN PREMIÈRE LIGNE)
======================================================================
- Si la diapositive contient un schéma synoptique, un organigramme, un graphique ou une illustration :
  --> Première ligne : ---VISUEL_DETECTE: OUI---
  --> Deuxième ligne : CADRE_VISUEL: [ymin, xmin, ymax, xmax]
- Sinon :
  --> Première ligne : ---VISUEL_DETECTE: NON---

======================================================================
3. STRATÉGIE RÉDACTIONNELLE
======================================================================
--- CAS 1 : SI VISUEL_DETECTE EST OUI ---
- Rédige une ANALYSE PÉDAGOGIQUE ET SYNTHÉTIQUE sous le visuel. Titre : `#### Explication & Synthèse du schéma`.
--- CAS 2 : SI VISUEL_DETECTE EST NON ---
- Restructure et rédige L'INTÉGRALITÉ du contenu sous forme de cours rédigé.

======================================================================
4. FORMULES MATHÉMATIQUES ET UNITÉS (SANS SYNTAXE LATEX)
======================================================================
- INTERDICTION ABSOLUE DU LATEX : N'utilise JAMAIS de symboles $,$$, \\text{}, \\times, \\frac.
- ÉCRITURE UNICODE LISIBLE : `1 m³ = 330 kg`, `1 ml = 1 750 kg`, `Poids Volumétrique (PV) = Volume (m³) × 330 kg`.

======================================================================
5. NORMES DE FORMATAGE MARKDOWN STRICTES
======================================================================
- Titres : `### ` pour sujets principaux, `#### ` pour sous-sections.
- Mots-clés en gras : `**terme**`.
- Listes à puces (`- `) et encadrés (`> `).
"""

def appeler_gemini(client_gemini, img_path, model_choice):
    img_pil = Image.open(img_path)
    models = [model_choice, "gemini-3.5-flash-lite", "gemini-flash-lite-latest"]
    
    for model_name in models:
        try:
            res = client_gemini.models.generate_content(
                model=model_name,
                contents=[img_pil, PROMPT_EXPERT]
            )
            if res and res.text:
                return res.text.strip()
        except Exception:
            continue
    return None

# ==============================================================================
# 5. ZONE PRINCIPALE DE DÉPÔT ET DE TRAITEMENT
# ==============================================================================
st.markdown('<div class="glass-card">', unsafe_allow_html=True)

st.markdown("""
    <div class="dropzone-container">
        <div style="font-size: 48px; margin-bottom: 8px;">☁️</div>
        <p style="margin: 0; font-size: 16px; font-weight: 600; color: #cbd5e1;">
            Glissez votre cours PDF ici ou <span style="color: #818cf8; text-decoration: underline; cursor: pointer;">Parcourir</span>
        </p>
    </div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Téléverser le PDF", type=["pdf"], label_visibility="collapsed")

st.markdown('</div>', unsafe_allow_html=True)

if uploaded_file:
    file_size_mb = round(len(uploaded_file.getvalue()) / (1024 * 1024), 1)
    st.markdown(f"""
        <div class="file-card">
            <div style="display:flex; align-items:center; gap:12px;">
                <span style="font-size:24px;">📄</span>
                <div>
                    <div style="font-weight:700; color:#f8fafc; font-size:14px;">{uploaded_file.name}</div>
                    <div style="font-size:12px; color:#94a3b8;">({file_size_mb} Mo)</div>
                </div>
            </div>
            <div style="background:rgba(239, 68, 68, 0.2); border:1px solid rgba(248, 113, 113, 0.4); padding:6px 10px; border-radius:8px; cursor:pointer; color:#fca5a5;">
                🗑️
            </div>
        </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:16px;'></div>", unsafe_allow_html=True)

    btn_start = st.button("🚀 LANCER LA NUMÉRISATION PÉDAGOGIQUE", use_container_width=True)

    col_monitor, col_success = st.columns([1.2, 1])

    if btn_start:
        client_gemini = genai.Client(api_key=API_KEY) if API_KEY else None

        with col_monitor:
            st.markdown('<div class="glass-card">', unsafe_allow_html=True)
            status_header = st.empty()
            progress_bar = st.progress(0)
            step_1 = st.empty()
            step_2 = st.empty()
            step_3 = st.empty()
            st.markdown('</div>', unsafe_allow_html=True)

        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = os.path.join(temp_dir, uploaded_file.name)
            with open(pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            pdf_doc = fitz.open(pdf_path)
            total_pages = len(pdf_doc)

            doc = Document()
            configurer_page(doc)

            # Application de la palette dynamique sélectionnée
            palette_active = PALETTES.get(style_option, PALETTES["Bleu Marine & Encadrés Pédagogiques"])
            COLOR_H2 = palette_active["H2"]
            COLOR_H3 = palette_active["H3"]
            COLOR_H4 = palette_active["H4"]

            # Titre principal dans le document Word
            titre_p = doc.add_heading(level=0)
            titre_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_titre = titre_p.add_run(f"COURS COMPLET : {uploaded_file.name.replace('.pdf', '').upper()}")
            run_titre.font.color.rgb = COLOR_H2
            run_titre.bold = True

            step_1.markdown("✓ **PDF converti en images HD**")

            for page_num in range(total_pages):
                pct = int(((page_num + 1) / total_pages) * 100)
                status_header.markdown(f"⏳ **Analyse de la diapositive {page_num + 1} / {total_pages}** `[{pct}%]`")
                progress_bar.progress((page_num + 1) / total_pages)

                step_2.markdown("⟳ *Nettoyage et rognage des schémas...*")

                page = pdf_doc[page_num]
                pix = page.get_pixmap(dpi=200)
                img_path = os.path.join(temp_dir, f"page_{page_num + 1}.png")
                pix.save(img_path)

                ajouter_titre(doc, f"Partie {page_num + 1}", level=2, color_rgb=COLOR_H2)

                texte_ia = appeler_gemini(client_gemini, img_path, selected_model_code) if client_gemini else "---VISUEL_DETECTE: NON---\n### Contenu synthétisé\n- Notion de cours."

                if texte_ia:
                    contient_visuel = "---VISUEL_DETECTE: OUI---" in texte_ia
                    path_a_inserer = img_path

                    if contient_visuel:
                        match_cadre = re.search(r'CADRE_VISUEL:\s*\[(\d+),\s*(\d+),\s*(\d+),\s*(\d+)\]', texte_ia)
                        if match_cadre:
                            coords = [int(g) for g in match_cadre.groups()]
                            path_a_inserer = rogner_schema(img_path, coords)

                    lignes_raw = texte_ia.split("\n")
                    lignes_clean = []
                    for l in lignes_raw:
                        l_str = l.strip()
                        if opt_decor and (l_str.startswith("---VISUEL_DETECTE:") or l_str.startswith("VISUEL:") or l_str.startswith("CADRE_VISUEL:") or l_str == "---"):
                            continue
                        if opt_latex:
                            l = re.sub(r'\\text\{([^}]+)\}', r'\1', l)
                            l = re.sub(r'\$\$(.*?)\$\$', r'\1', l)
                            l = re.sub(r'\$(.*?)\$', r'\1', l)
                            l = l.replace(r'\times', '×').replace(r'\euro', '€')
                        lignes_clean.append(l)

                    texte_formate = "\n".join(lignes_clean).strip()

                    if contient_visuel:
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run_img = p_img.add_run()
                        run_img.add_picture(path_a_inserer, width=Inches(5.5))

                    for ligne in texte_formate.split("\n"):
                        l = ligne.strip()
                        if not l:
                            continue
                        if l.startswith("#### "):
                            ajouter_titre(doc, l.replace("#### ", ""), level=4, color_rgb=COLOR_H4)
                        elif l.startswith("### "):
                            ajouter_titre(doc, l.replace("### ", ""), level=3, color_rgb=COLOR_H3)
                        elif l.startswith("> "):
                            ajouter_encadre_callout(doc, l.replace("> ", ""))
                        elif l.startswith("- ") or l.startswith("* "):
                            p = doc.add_paragraph(style='List Bullet')
                            parse_and_add_formatted_runs(p, l[2:])
                        elif re.match(r'^\d+\.\s', l):
                            p = doc.add_paragraph(style='List Number')
                            parse_and_add_formatted_runs(p, re.sub(r'^\d+\.\s', '', l))
                        else:
                            p = doc.add_paragraph()
                            parse_and_add_formatted_runs(p, l)

                step_3.markdown("✓ **Rédaction des puces et conversion Unicode**")
                time.sleep(0.5)

            output_docx_path = os.path.join(temp_dir, "Cours_Formate.docx")
            doc.save(output_docx_path)

            with col_success:
                st.markdown("""
                    <div class="success-card">
                        <div style="font-size:28px; margin-bottom:8px;">🎉</div>
                        <h4 style="margin:0 0 12px 0; color:#34d399; font-size:16px;">Document Word généré avec succès !</h4>
                    </div>
                """, unsafe_allow_html=True)
                
                st.markdown("<div style='height:10px;'></div>", unsafe_allow_html=True)

                with open(output_docx_path, "rb") as file_docx:
                    st.download_button(
                        label="📥 Télécharger le document Word (.docx)",
                        data=file_docx,
                        file_name=f"{uploaded_file.name.replace('.pdf', '')}_Cours_Ultime.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
