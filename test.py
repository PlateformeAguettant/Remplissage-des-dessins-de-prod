# -*- coding: utf-8 -*-
import io, os, glob, zipfile, tempfile, shutil, base64
from datetime import datetime
import pandas as pd
import streamlit as st
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color

# ====== Config watermark par défaut ======
TARGET_PAGES = "1"
FONT_NAME = "Helvetica-Bold"
FONT_SIZE = 9
OPACITY    = 0.70

COL_ART_CANDIDATES = ["N° article","N°article","No article","N° Article"]
COL_LOT_CANDIDATES = ["N° lot","N°lot","No lot","N° Lot"]
# >>> NEW: candidats pour la colonne Date d’expiration
COL_EXP_CANDIDATES = [
    "Date d'expiration","Date D'expiration","Date d’expiration",
    "Date Expiration","Expiration","Date d`expiration","Date d expiration"
]

# ====== Comptes démo ======
USERS = {
    "safae": {"password": "1234"},
    "benjamin": {"password": "BnSzny2023v6"},
    "sandra": {"password": "sasoare"},
    "geoffrey": {"password": "Aguettant2025"},
}

# ====== Icônes locales ======
ICON_PATH  = "icons8-ouvrir-50.png"      # cadenas / ouvrir
ICON_PATH1 = "icons8-utilisateur-32.png" # utilisateur

def inline_icon(path: str, height: int = 32) -> str:
    """Retourne une balise <img> en base64 si le fichier existe, sinon chaîne vide."""
    if os.path.exists(path):
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        return f'<img src="data:image/png;base64,{b64}" style="height:{height}px; vertical-align:middle; margin-right:10px;" />'
    return ""

# ==================== CONFIG & CSS ====================
st.set_page_config(page_title="Remplissage des dessins de prod", page_icon="", layout="wide")

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] { height: 100%; }
.main .block-container{
    max-width: 100% !important;
    padding-left: 3vw; padding-right: 3vw;
    padding-top: 1.2rem;
}
section.main > div { min-height: calc(100vh - 140px); }

.header-box, div.stForm { max-width: 100%; margin: 0 0 24px 0; }

.header-box{
    width: 100% !important;
    height: 150px;
    padding: 16px 24px;
    background:#1E3361;
    border-radius:12px;
    display:flex; align-items:center; justify-content:space-between;
    text-align:center; box-shadow: 0 6px 18px rgba(0,0,0,0.08);
}
.header-box h2{ color:white; margin:0; font-size:23px !important; font-weight:700; }

div[data-testid="stForm"]{
    width: 100% !important;
    max-width: 100% !important;
    margin: 0 auto 30px !important;
    border: 1px solid rgba(0,0,0,0.08);
    border-radius: 12px;
    padding: 28px 24px 36px;
    background: white;
    min-height: 420px;
    display: flex; flex-direction: column; gap: 18px;
}
div[data-testid="stForm"] > div{ display:flex; flex-direction:column; gap:16px; }

div.stForm input[type="text"], div.stForm input[type="password"]{
    width: 100% !important;
    height: 54px !important;
    font-size: 16px !important;
    line-height: 54px !important;
    padding: 0 16px !important;
    border-radius: 10px !important;
}

[data-testid="stSidebar"]{
    background-color:#1E3361 !important; color:white !important;
    border-right: 1px solid rgba(255,255,255,0.08);
}
[data-testid="stSidebar"] *{ color:white !important; }

.user-box{
    background:#324A86; padding:12px 14px; border-radius:12px; font-weight:700;
    box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin: 8px 6px 14px 6px;
}

div.stButton > button, .stDownloadButton > button, [data-testid="stSidebar"] button{
    background-color:#1E3361 !important; color:white !important; font-weight:700 !important;
    border:none !important; border-radius:10px !important; padding:10px 24px !important;
    box-shadow:0 4px 12px rgba(0,0,0,0.15) !important; display:block; margin:0 auto;
}
div.stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stSidebar"] button:hover{
    background-color:#43c3b4 !important; color:white !important;
}

button[data-testid*="FormSubmit"],
div.stForm button[type="submit"],
div[data-testid="stForm"] button {
    background-color:#1E3361 !important;
    color:#ffffff !important;
    font-weight:700 !important;
    border:none !important;
    border-radius:10px !important;
    padding:12px 26px !important;
    box-shadow:0 4px 12px rgba(0,0,0,0.15) !important;
}
button[data-testid*="FormSubmit"]:hover,
div.stForm button[type="submit"]:hover,
div[data-testid="stForm"] button:hover {
    background-color:#43c3b4 !important; color:#ffffff !important;
}
button[data-testid*="FormSubmit"] *,
div.stForm button[type="submit"] *,
div[data-testid="stForm"] button * {
    color:#ffffff !important; fill:#ffffff !important; stroke:#ffffff !important;
}

[data-testid="collapsedControl"]{
    opacity:1 !important; visibility:visible !important;
    background-color:#324A86 !important; border-radius:10px !important;
    border:2px solid #ffffff !important; width:44px !important; height:44px !important;
    display:flex !important; align-items:center !important; justify-content:center !important;
    box-shadow:0 4px 12px rgba(0,0,0,0.25) !important;
}
[data-testid="collapsedControl"] svg{ fill:white !important; stroke:white !important; color:white !important; }
[data-testid="collapsedControl"]:hover{ background-color:#43c3b4 !important; border-color:#ffffff !important; }
</style>
""", unsafe_allow_html=True)

# ==================== EN-TÊTE ====================
def app_header(title="Application"):
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        logo_base64 = base64.b64encode(open(logo_path, "rb").read()).decode()
        logo_html = f'<img src="data:image/jpg;base64,{logo_base64}" style="height:60px;"/>'
    else:
        logo_html = ""
    st.markdown(
        f"""
        <div class="header-box">
            {logo_html}
            <h2>{title}</h2>
            {logo_html}
        </div>
        """,
        unsafe_allow_html=True
    )

# ============ Utilitaires ============
def read_excel_filelike(filelike, **kwargs) -> pd.DataFrame:
    data = filelike.read() if hasattr(filelike, "read") else filelike
    bio = io.BytesIO(data if isinstance(data,(bytes,bytearray)) else data.getvalue())
    return pd.read_excel(bio, engine="openpyxl", **kwargs)

def first_existing_col(df: pd.DataFrame, candidates):
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    return None

def draw_layer(w, h, elements):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(w, h))
    for el in elements:
        text  = str(el["text"])
        x     = float(el.get("x", 0)); y = float(el.get("y", 0))
        ang   = float(el.get("angle", 0))
        size  = float(el.get("size", FONT_SIZE))
        opac  = float(el.get("opacity", OPACITY))
        font  = el.get("font", FONT_NAME)
        align = el.get("align", "left")
        grey  = max(0.0, min(1.0, 1.0 - opac))
        c.setFillColor(Color(grey, grey, grey))
        c.setFont(font, size)
        c.saveState(); c.translate(x, y); c.rotate(ang)
        (c.drawCentredString if align=="center" else c.drawString)(0,0,text)
        c.restoreState()
    c.showPage(); c.save(); buf.seek(0)
    return PdfReader(buf).pages[0]

def _map_coords_for_rotation(xv, yv, w, h, rot):
    rot = (rot or 0) % 360
    if rot==0: return xv, yv
    if rot==90: return yv, w-xv
    if rot==180: return w-xv, h-yv
    if rot==270: return h-yv, xv
    return xv, yv

def parse_pages(spec, total_pages):
    s = str(spec).strip().lower()
    if s=="all": return list(range(total_pages))
    res=set()
    for part in s.split(","):
        part=part.strip()
        if not part: continue
        if "-" in part:
            a,b = part.split("-",1); a,b = int(a),int(b)
            for i in range(a,b+1):
                if 1<=i<=total_pages: res.add(i-1)
        else:
            i=int(part); 
            if 1<=i<=total_pages: res.add(i-1)
    return sorted(res)

def watermark_pdf_multi(src_bytes, out_buf, elements):
    reader = PdfReader(io.BytesIO(src_bytes))
    first  = reader.pages[0]
    w,h    = float(first.mediabox.width), float(first.mediabox.height)
    rot    = (getattr(first,"rotation",0) or 0) % 360

    elems_adj=[]
    for e in elements:
        x0,y0=_map_coords_for_rotation(float(e["x"]), float(e["y"]), w, h, rot)
        e2=dict(e); e2["x"],e2["y"]=x0,y0; elems_adj.append(e2)

    layer  = draw_layer(w,h,elems_adj)
    targets= parse_pages(TARGET_PAGES, len(reader.pages))
    writer = PdfWriter()

    for idx,p in enumerate(reader.pages):
        base=PageObject.create_blank_page(width=w,height=h)
        base.merge_page(p)
        if idx in targets: base.merge_page(layer)
        if getattr(p,"rotation",0): base.rotate(p.rotation)
        writer.add_page(base)

    writer.write(out_buf); out_buf.seek(0)

# ==================== NAVIGATION / ÉTATS ====================
for k,v in dict(step="login", username="", nom="", prenom="").items():
    st.session_state.setdefault(k, v)

# ---------------- PAGE 1 : LOGIN ----------------
def login_view():
    app_header("Remplissage des dessins de production")
    icon_html = inline_icon(ICON_PATH, height=32)
    st.markdown(f"<h2 style='font-size:32px; font-weight:800;'>{icon_html}Connexion</h2>", unsafe_allow_html=True)

    with st.form("login"):
        u = st.text_input("Identifiant")
        p = st.text_input("Mot de passe", type="password")
        ok = st.form_submit_button("Se connecter")
    if ok:
        rec = USERS.get(u.strip().lower())
        if rec and p == rec["password"]:
            st.session_state.username = u.strip()
            st.session_state.step = "profile"
            st.success("Connexion réussie ✅"); st.rerun()
        else:
            st.error("Identifiants invalides")

# ---------------- PAGE 2 : PROFIL ----------------
def profile_view():
    app_header("Remplissage des dessins de production")
    icon_html1 = inline_icon(ICON_PATH1, height=32)
    st.markdown(f"<h2 style='font-size:28px; font-weight:700;'>{icon_html1}Profil utilisateur</h2>", unsafe_allow_html=True)
    st.caption(f"Connecté en tant que **{st.session_state.username}**")
    with st.form("profil"):
        col1,col2 = st.columns(2)
        with col1:
            prenom = st.text_input("Prénom", value=st.session_state.prenom)
        with col2:
            nom    = st.text_input("Nom", value=st.session_state.nom)
        ok = st.form_submit_button("Continuer")
    if ok:
        if not prenom or not nom:
            st.warning("Merci de renseigner **Nom** et **Prénom**.")
        else:
            st.session_state.prenom = prenom.strip()
            st.session_state.nom    = nom.strip()
            st.session_state.step   = "app"
            st.success("Profil enregistré ✅"); st.rerun()

# ---------------- PAGE 3 : APP ----------------
def app_view():
    app_header("Remplissage des dessins de production")
    st.markdown("<h2 style='font-size:28px; font-weight:700;'>Application</h2>", unsafe_allow_html=True)
    full_name = f"{st.session_state.prenom} {st.session_state.nom}".strip()

    # --- Sidebar : user box stylée ---
    st.sidebar.markdown(
        f"<div class='user-box'>Utilisateur : {full_name}</div>",
        unsafe_allow_html=True
    )
    if st.sidebar.button("Se déconnecter"):
        st.session_state.clear(); st.session_state.step="login"; st.rerun()

    c1,c2 = st.columns(2)
    with c1: excel_rempl = st.file_uploader("Excel — Remplissage", type=["xlsx"])
    with c2: excel_refs  = st.file_uploader("Excel — Références",  type=["xlsx"])
    uploaded = st.file_uploader("PDFs vierges (plusieurs) ou un ZIP", type=["pdf","zip"], accept_multiple_files=True)

    go = st.button("Remplir")
    if not go: return

    if not excel_rempl or not excel_refs:
        st.error("Merci d’uploader les **deux Excel**."); return
    if not uploaded:
        st.error("Merci d’uploader des **PDF** ou un **ZIP**."); return

    # Charger Excel
    df_rempl = read_excel_filelike(excel_rempl)
    df_refs  = read_excel_filelike(excel_refs)
    col_art  = first_existing_col(df_rempl, COL_ART_CANDIDATES)
    col_lot  = first_existing_col(df_rempl, COL_LOT_CANDIDATES)
    col_ref  = first_existing_col(df_refs,  COL_ART_CANDIDATES)
    # >>> NEW: localiser la colonne Date d'expiration (si présente)
    col_exp  = first_existing_col(df_rempl, COL_EXP_CANDIDATES)

    if not col_art or not col_ref:
        st.error("Colonnes 'N° article' manquantes."); return

    refs = set(df_refs[col_ref].dropna().map(lambda x: str(x).strip().lower()))

    # Construire l’index PDF
    pdf_bytes = {}
    tmp=None
    zips=[f for f in uploaded if f.name.lower().endswith(".zip")]
    pdfs=[f for f in uploaded if f.name.lower().endswith(".pdf")]
    if len(zips)==1 and len(pdfs)==0:
        tmp=tempfile.mkdtemp(prefix="pdf_src_")
        with zipfile.ZipFile(io.BytesIO(zips[0].read())) as zf: zf.extractall(tmp)
        for p in glob.glob(os.path.join(tmp,"**","*.pdf"), recursive=True):
            with open(p,"rb") as f: pdf_bytes[os.path.basename(p)] = f.read()
    else:
        for up in pdfs: pdf_bytes[up.name]=up.read()

    processed=skipped_no_match=skipped_not_ref=0
    zip_buf=io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
        for _,row in df_rempl.iterrows():
            art_full=str(row.get(col_art,"") or "").strip().lower()
            if not art_full: continue
            if art_full not in refs: skipped_not_ref+=1; continue

            k8 = (art_full[:8]).lower()
            candidates=[fn for fn in pdf_bytes if os.path.splitext(fn)[0].split("-",1)[0][:8].lower()==k8]
            if not candidates: skipped_no_match+=1; continue

            lot_raw=str(row.get(col_lot,"") or "").strip()
            lot8   ="Valeur manquante" if (not lot_raw or len(lot_raw)<8) else lot_raw[:8]
            ordre2 = "Valeur manquante" if (not lot_raw or len(lot_raw) < 2) else lot_raw[-2:]
            date_txt = datetime.now().strftime("%d/%m/%Y")

            # >>> NEW: calcul de la Date d'expiration au format MM/AAAA
            exp_txt = "Valeur manquante"
            if col_exp:
                raw_exp = row.get(col_exp, "")
                exp_dt = pd.to_datetime(raw_exp, dayfirst=True, errors="coerce")
                if pd.notna(exp_dt):
                    if exp_dt.day in (1, 2):
                        exp_dt = exp_dt - pd.DateOffset(months=1)
                    exp_txt = exp_dt.strftime("%m/%Y")

            for fn in candidates:
                base_full = os.path.splitext(fn)[0]
                out_name  = f"{base_full}-{lot8.replace(' ','_')}-{ordre2.replace(' ','_')}.pdf"

                elements = [
                    {"text": full_name, "x": 56,  "y": 688, "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                    {"text": date_txt,  "x": 78,  "y": 670, "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                    {"text": lot8,      "x": 115, "y": 80,  "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                    {"text": ordre2,    "x": 113, "y": 295, "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                    {"text": "Voir planning", "x": 113, "y": 378, "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                    # >>> NEW: Date d'expiration placée à x=96, y=575
                    {"text": exp_txt,   "x": 96,  "y": 575, "angle": 90, "size": FONT_SIZE, "opacity": OPACITY, "font": FONT_NAME},
                ]

                outb=io.BytesIO()
                watermark_pdf_multi(pdf_bytes[fn], outb, elements)
                z.writestr(out_name, outb.getvalue())
                processed+=1

    zip_buf.seek(0)
    st.success(f"Terminé ✅ — Traités: {processed} | Sans PDF: {skipped_no_match} | Hors références: {skipped_not_ref}")
    st.download_button("⬇️ Télécharger le ZIP", data=zip_buf,
        file_name=f"dessins_remplis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip", use_container_width=True)

    if tmp and os.path.isdir(tmp): shutil.rmtree(tmp, ignore_errors=True)

# ================= ROUTEUR =================
step = st.session_state.get("step", "login")
if step == "login":
    login_view()
elif step == "profile":
    profile_view()
elif step == "app":
    app_view()
else:
    # fallback au cas où l'état est corrompu
    st.session_state.step = "login"
    st.rerun()


