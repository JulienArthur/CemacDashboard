"""
generate_cemac_dashboard.py
============================
Script de génération du Dashboard mensuel CEMAC.

USAGE :
    1. Déposer les fichiers PDF du mois dans ce dossier :
         - "Point 3_RPM [mois] [année]_vf_.pdf"
         - "Point 3_annexe_Tableau_de_bord_CPM_[mois]_[année]_vN.pdf"
    2. Lancer : python generate_cemac_dashboard.py
    3. Le fichier Excel Dashboard_[mois]_[année].xlsx est généré dans ce dossier.

CARTOGRAPHIE DES SOURCES (vérifiée sur Dec 2025) :
    Annexe PDF — pages réelles (index 0-based entre parenthèses) :
        p26 (idx 25) : T18 CEMAC Agrégats monnaie/crédit (TCEM)
        p22 (idx 21) : T16 Balance des paiements
        p33 (idx 32) : T25 Taux directeurs BEAC/BCE
        p39 (idx 38) : T32 CEMAC indicateurs globaux
        p40 (idx 39) : T33 Cameroun
        p41 (idx 40) : T34 Centrafrique
        p42 (idx 41) : T35 Congo
        p43 (idx 42) : T36 Gabon
        p44 (idx 43) : T37 Guinée Équatoriale
        p45 (idx 44) : T38 Tchad
"""

import os
import re
import glob
import unicodedata
import pdfplumber
import pandas as pd
from collections import defaultdict
from datetime import datetime


# ==============================================================================
# BLOC 1 : CONFIGURATION
# ==============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MOIS_FR = {
    'janvier': '01', 'fevrier': '02', 'mars': '03', 'avril': '04',
    'mai': '05', 'juin': '06', 'juillet': '07', 'aout': '08',
    'septembre': '09', 'octobre': '10', 'novembre': '11',
    'decembre': '12', 'dec': '12'
}


def normalise(text):
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii').lower()


def find_latest_pdfs():
    rpm_files    = sorted(glob.glob(os.path.join(BASE_DIR, "Point 3_RPM*.pdf")))
    annexe_files = sorted(glob.glob(os.path.join(BASE_DIR, "Point 3_annexe*.pdf")))
    rpm_path    = rpm_files[-1]    if rpm_files    else None
    annexe_path = annexe_files[-1] if annexe_files else None
    source_name = normalise(os.path.basename(rpm_path) if rpm_path else "")
    month_rx = r'(' + '|'.join(MOIS_FR.keys()) + r')\s*(\d{4})'
    m = re.search(month_rx, source_name)
    period_str = f"{m.group(1)}_{m.group(2)}" if m else datetime.now().strftime("%B_%Y")
    return rpm_path, annexe_path, period_str


# ==============================================================================
# BLOC 2 : EXTRACTION — Fonctions de parsing par type de tableau
# ==============================================================================

def group_words_by_row(words, y_tolerance=4):
    """Groupe les mots par ligne (proximité verticale — bucketing fixe)."""
    rows = defaultdict(list)
    for w in words:
        key = round(w['top'] / y_tolerance) * y_tolerance
        rows[key].append(w)
    return {k: sorted(rows[k], key=lambda w: w['x0']) for k in sorted(rows.keys())}


def group_words_by_proximity(words, gap=5.0):
    """
    Groupe les mots par lignes en fusionnant les mots dont la distance
    verticale est < gap. Retourne {key: [words_sorted_by_x]}.
    Robuste aux légères variations de top au sein d'une même ligne logique.
    """
    if not words:
        return {}
    sorted_words = sorted(words, key=lambda w: w['top'])
    groups = {}
    current_key = sorted_words[0]['top']
    groups[current_key] = [sorted_words[0]]

    for w in sorted_words[1:]:
        if w['top'] - current_key > gap:
            current_key = w['top']
            groups[current_key] = []
        groups[current_key].append(w)

    return {k: sorted(v, key=lambda w: w['x0']) for k, v in sorted(groups.items())}


def find_column_centers(rows_dict, month_pattern=r'\w{3,4}-\d{2}'):
    """
    Trouve la ligne d'en-tête de colonnes (mois/années) et retourne
    (header_labels, col_centers_x, header_row_y).
    """
    month_rx = re.compile(month_pattern)
    for y, row_words in rows_dict.items():
        texts = [w['text'] for w in row_words]
        matches = [t for t in texts if month_rx.match(t)]
        if len(matches) >= 5:
            col_centers = [(w['x0'] + w['x1']) / 2 for w in row_words if month_rx.match(w['text'])]
            col_labels  = [w['text'] for w in row_words if month_rx.match(w['text'])]
            return col_labels, col_centers, y
    return None, None, None


def assign_to_column(x0, col_centers):
    """Retourne l'index de la colonne la plus proche pour une position x donnée."""
    if not col_centers:
        return 0
    return min(range(len(col_centers)), key=lambda i: abs(x0 - col_centers[i]))


def extract_tcem(pdf_path, page_idx):
    """
    T18 CEMAC : Extraction par positions des mots.
    Gère les nombres avec espace comme séparateur de milliers (ex: 2 485 304).
    Retourne (df_niveaux, df_variations, meta).
    """
    meta = {
        "source_pdf": os.path.basename(pdf_path),
        "page_pdf": f"p.{page_idx + 1} de l'Annexe",
        "table_label": "T18 — CEMAC : Agrégats de monnaie et de crédit",
        "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M")
    }

    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return None, None, meta
        words = pdf.pages[page_idx].extract_words(x_tolerance=2, y_tolerance=3)

    # group_words_by_proximity(gap=5) : fusionne mots à < 5pt verticalement.
    # Robuste aux décalages label/valeurs de ~0.35pt tout en séparant les lignes (~15pt d'écart)
    rows_dict = group_words_by_proximity(words, gap=5.0)

    # Trouver les deux tableaux : niveaux et variations
    # On identifie les deux lignes d'en-têtes de mois
    month_rx = re.compile(r'\w{3,4}-\d{2}')

    header_rows = []
    for y, row_words in rows_dict.items():
        if sum(1 for w in row_words if month_rx.match(w['text'])) >= 5:
            header_rows.append(y)

    if len(header_rows) < 1:
        print(f"  ⚠ Aucune ligne d'en-tête trouvée dans T18")
        return None, None, meta

    # Trouver la ligne "Source:" pour délimiter la fin du tableau niveaux
    source_rows = []
    for y, row_words in rows_dict.items():
        if any(w['text'].lower().startswith('source') for w in row_words):
            source_rows.append(y)

    def extract_one_table(header_y, end_y, label_xmax=280):
        """
        Extrait un tableau entre header_y et end_y.
        Gère les cas : label+valeurs sur la même ligne, label seul puis valeurs,
        et label multi-ligne.
        """
        # En-têtes de colonnes
        header_words = rows_dict[header_y]
        col_labels  = [w['text'] for w in header_words if month_rx.match(w['text'])]
        col_centers = [(w['x0'] + w['x1']) / 2 for w in header_words if month_rx.match(w['text'])]
        n_cols = len(col_labels)

        # Lignes de données (après header_y, avant end_y)
        data_rows_y = [y for y in rows_dict
                       if y > header_y and (end_y is None or y < end_y)]

        result_rows = []
        pending_label_text = ""
        pending_values = defaultdict(list)

        def flush_pending():
            nonlocal pending_label_text, pending_values
            if pending_label_text and any(pending_values.values()):
                vals = [' '.join(pending_values.get(ci, [])) for ci in range(n_cols)]
                result_rows.append([pending_label_text] + vals)
            # Dans tous les cas, réinitialiser
            pending_label_text = ""
            pending_values = defaultdict(list)

        for y in sorted(data_rows_y):
            row_words = rows_dict[y]
            label_words = [w for w in row_words if w['x0'] < label_xmax]
            value_words = [w for w in row_words if w['x0'] >= label_xmax]

            # Ignorer : numéro de page, "Source:", lignes vides
            if not row_words:
                continue
            if len(row_words) == 1 and re.fullmatch(r'\d{1,3}', row_words[0]['text']):
                continue
            if any(w['text'].lower().startswith('source') for w in label_words):
                break  # Fin du tableau

            # Ignorer les lignes inter-section (seulement des années ou "CEMAC" dans la zone valeurs)
            if not label_words and value_words:
                val_texts = [w['text'] for w in value_words]
                if all(re.fullmatch(r'\d{4}', t) for t in val_texts):
                    continue  # Ligne "2024 2025" inter-section, ignorer

            has_label  = bool(label_words)
            has_values = bool(value_words)

            if has_label and has_values:
                # Ligne complète (label + valeurs) : flush le pending, émettre immédiatement
                flush_pending()
                label_text = ' '.join(w['text'] for w in label_words)
                row_vals = defaultdict(list)
                for vw in value_words:
                    ci = assign_to_column((vw['x0'] + vw['x1']) / 2, col_centers)
                    row_vals[ci].append(vw['text'])
                vals = [' '.join(row_vals.get(ci, [])) for ci in range(n_cols)]
                result_rows.append([label_text] + vals)

            elif has_label and not has_values:
                # Ligne de label seul
                label_text = ' '.join(w['text'] for w in label_words)
                if pending_label_text and any(pending_values.values()):
                    # Le pending précédent est complet (label + valeurs) → flush
                    # puis commencer un nouveau pending pour cet indicateur
                    flush_pending()
                    pending_label_text = label_text
                elif pending_label_text and not any(pending_values.values()):
                    # Continuation d'un label multi-lignes (ex: "la monnaie (en %)")
                    pending_label_text += ' ' + label_text
                else:
                    # Nouveau label sans valeurs encore
                    pending_label_text = label_text

            elif not has_label and has_values:
                # Ligne de valeurs seules (les valeurs de la ligne label-only précédente)
                for vw in value_words:
                    ci = assign_to_column((vw['x0'] + vw['x1']) / 2, col_centers)
                    pending_values[ci].append(vw['text'])
                # Flush immédiatement
                flush_pending()

        flush_pending()  # Dernière ligne en attente

        if not result_rows:
            return None
        df = pd.DataFrame(result_rows, columns=['Indicateur'] + col_labels)
        return df

    # Délimiter le tableau niveaux : du premier header jusqu'à la première ligne "Source:"
    h0 = header_rows[0]
    # Fin du tableau niveaux = première "Source:" après h0
    niv_end = next((y for y in source_rows if y > h0), None)
    df_niv = extract_one_table(h0, niv_end, label_xmax=280)

    # Tableau variations : du deuxième header jusqu'à la deuxième "Source:" (ou fin de page)
    df_var = None
    if len(header_rows) >= 2:
        h1 = header_rows[1]
        var_end = next((y for y in source_rows if y > h1), None)
        df_var = extract_one_table(h1, var_end, label_xmax=280)

    # Post-traitement : reconstruire le label "Taux de couverture extérieure de la monnaie (en %)"
    # qui apparaît fragmenté dans les deux tableaux
    def fix_tcem_labels(df):
        if df is None:
            return None
        TAUX_LABEL = "Taux de couverture extérieure de la monnaie (en %)"
        rows_out = []
        i = 0
        data = df.values.tolist()
        while i < len(data):
            row = data[i]
            lbl = str(row[0]) if row[0] else ""
            # Si le label se termine par "de" ou contient "couverture extérieure de"
            # et que la ligne SUIVANTE commence par "la monnaie", fusionner les labels
            if "couverture" in lbl.lower() and lbl.strip().endswith("de"):
                # Remplacer par le label complet
                row[0] = TAUX_LABEL
            elif lbl.startswith("la monnaie"):
                # C'est un artefact du label fragmenté — ignorer cette ligne
                i += 1
                continue
            rows_out.append(row)
            i += 1
        return pd.DataFrame(rows_out, columns=df.columns)

    df_niv = fix_tcem_labels(df_niv)
    df_var = fix_tcem_labels(df_var)

    if df_niv is not None:
        print(f"  ✓ T18 Niveaux : {len(df_niv)} lignes × {len(df_niv.columns)} colonnes")
    if df_var is not None:
        print(f"  ✓ T18 Variations : {len(df_var)} lignes × {len(df_var.columns)} colonnes")

    return df_niv, df_var, meta


def extract_bop(pdf_path, page_idx):
    """
    T16 Balance des paiements — parsing texte ligne par ligne.
    Colonnes : Pays | STC_2023 | STC_2024 | STC_2025 | SCKF_2023 | SCKF_2024 | SCKF_2025 | SG_2023 | SG_2024 | SG_2025
    """
    meta = {
        "source_pdf": os.path.basename(pdf_path),
        "page_pdf": f"p.{page_idx + 1} de l'Annexe",
        "table_label": "T16 — Soldes de la balance des paiements (% PIB)",
        "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return None, meta
        text = pdf.pages[page_idx].extract_text() or ""

    # Trouver la ligne d'en-tête des colonnes (années)
    # Format: "2023 20242025 Prev. 2023 20242025 Prev. 2023 20242025 Prev."
    # ou "2023 2024 2025 Prev."
    cols_fixed = ['Pays',
                  'Solde TC 2023', 'Solde TC 2024', 'Solde TC 2025 Prév.',
                  'Solde CKF 2023', 'Solde CKF 2024', 'Solde CKF 2025 Prév.',
                  'Solde Global 2023', 'Solde Global 2024', 'Solde Global 2025 Prév.']

    pays_list = ['Cameroun', 'Centrafrique', 'Congo', 'Gabon', 'Guinée', 'Tchad', 'CEMAC',
                 'République', 'Equatoriale']
    rows = []

    for line in text.split('\n'):
        line = line.strip()
        # Détecter lignes avec pays + données numériques
        nums = re.findall(r'-?\d+[,\.]\d+', line)
        is_pays = any(p.lower() in line.lower() for p in pays_list)
        if not (is_pays and len(nums) >= 3):
            continue
        # Extraire nom du pays (tout avant le premier chiffre ou signe -)
        m = re.search(r'(?<!\d)(-?\d)', line)
        if m:
            pays_name = line[:m.start()].strip()
            vals = nums[:9]
        else:
            pays_name = line
            vals = []
        while len(vals) < 9:
            vals.append('')
        rows.append([pays_name] + vals)

    if not rows:
        print(f"  ⚠ T16 BOP : aucune donnée trouvée")
        return None, meta

    df = pd.DataFrame(rows, columns=cols_fixed)
    print(f"  ✓ T16 BOP : {len(df)} pays × {len(df.columns)} colonnes")
    return df, meta


def extract_taux_directeurs(pdf_path, page_idx):
    """
    T25 Taux directeurs BEAC/BCE — parsing texte.
    Indicateur | oct-24 | nov-24 | ... | oct-25
    """
    meta = {
        "source_pdf": os.path.basename(pdf_path),
        "page_pdf": f"p.{page_idx + 1} de l'Annexe",
        "table_label": "T25 — Évolution des principaux taux directeurs BEAC et BCE",
        "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return None, meta
        text = pdf.pages[page_idx].extract_text() or ""

    lines = text.split('\n')
    month_rx = re.compile(r'\w{3,4}-\d{2}')
    num_rx   = re.compile(r'-?\d+[,\.]\d+')

    header_months = None
    rows = []

    # Regex pour détecter du bruit de graphique (axes G29/G30 extraits comme texte)
    # Motif : ligne composée quasi-uniquement de chiffres/lettres isolés séparés par des espaces
    noise_rx = re.compile(r'^[\d\s\-]{20,}$')

    data_found = False  # Passé True dès qu'on trouve la première ligne de données

    for line in lines:
        line = line.strip()
        # Trouver la ligne d'en-tête des mois
        months = month_rx.findall(line)
        if len(months) >= 5 and header_months is None:
            header_months = months
            continue

        if header_months is None:
            continue

        # STOP si on atteint les notes de bas de page ou les légendes de graphiques
        # (évite de capturer G29/G30 dont les axes s'affichent en texte fragmenté)
        if line.startswith('*') or line.startswith('G2') or line.startswith('G3'):
            break
        if noise_rx.match(line):
            break  # Séquence de chiffres isolés = axe de graphique inversé

        # Ignorer les lignes courtes non-significatives
        if not line or re.fullmatch(r'\d{1,3}', line):
            continue
        if line.lower().startswith('source'):
            break

        nums = num_rx.findall(line)
        if not nums:
            # Titre de section (aucun chiffre)
            if len(line) > 3:
                rows.append([line] + [''] * len(header_months))
            continue

        data_found = True

        # Séparer libellé et valeurs
        m = re.search(r'(?<=[A-Za-zÀ-ÿ%\)])\s+(-?\d)', line)
        if m:
            label = line[:m.start()].strip()
            vals  = num_rx.findall(line[m.start():])
        else:
            m2 = re.search(r'(-?\d)', line)
            label = line[:m2.start()].strip() if m2 else line
            vals  = nums

        # Compléter ou tronquer à n_cols valeurs
        n = len(header_months)
        vals = (vals + [''] * n)[:n]
        rows.append([label] + vals)

    if not header_months or not rows:
        print(f"  ⚠ T25 Taux directeurs : aucune donnée trouvée")
        return None, meta

    cols = ['Indicateur'] + header_months
    df = pd.DataFrame(rows, columns=cols)
    print(f"  ✓ T25 Taux directeurs : {len(df)} lignes × {len(df.columns)} colonnes")
    return df, meta


def extract_indicators_table(pdf_path, page_idx, table_label="", page_label=""):
    """
    Extraction générique pour T32 (CEMAC) et T33-T38 (pays).
    Ces tableaux ont des années en colonnes et des indicateurs en lignes.
    Les valeurs sont des nombres décimaux sans espace comme séparateur.
    """
    meta = {
        "source_pdf": os.path.basename(pdf_path),
        "page_pdf": page_label or f"p.{page_idx + 1} de l'Annexe",
        "table_label": table_label,
        "extraction_date": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    with pdfplumber.open(pdf_path) as pdf:
        if page_idx >= len(pdf.pages):
            return None, meta
        text = pdf.pages[page_idx].extract_text() or ""

    lines = text.split('\n')

    # Trouver les lignes d'en-tête contenant les années (19xx ou 20xx uniquement)
    year_rx  = re.compile(r'\b(19\d{2}|20[0-3]\d)\b')
    num_rx   = re.compile(r'-?\d+[,\.]\d+|-?\d+\s')

    # On cherche la ligne "Estim" ou une ligne d'années multiples
    header_lines  = []
    header_y_idx  = None

    for i, line in enumerate(lines):
        years = year_rx.findall(line)
        if len(years) >= 4:
            header_lines.append((i, line))
            if header_y_idx is None:
                header_y_idx = i

    if header_y_idx is None:
        print(f"  ⚠ {table_label[:40]} : en-tête années introuvable")
        return None, meta

    # Construire les colonnes à partir des lignes d'en-tête
    # Pour T32 : "2020 2021 2022 2023 2024 Scénario de base* ..."
    # → colonnes : 2020, 2021, 2022, 2023, 2024, Estim_2025, Base_2026, Base_2027, Base_2028, Pess_2025, Pess_2026, ...
    # On garde les années brutes + Estim + les scénarios
    all_header_text = ' '.join(line for _, line in header_lines)
    year_tokens = year_rx.findall(all_header_text)

    # Détecter si "Estim" présent
    has_estim = 'estim' in all_header_text.lower()

    # Construire des noms de colonnes uniques
    col_names = []
    seen_years = {}
    for yr in year_tokens:
        if yr not in seen_years:
            seen_years[yr] = 0
            col_names.append(yr)
        else:
            seen_years[yr] += 1
            suffix = 'Pess.' if seen_years[yr] >= 1 else 'Base'
            col_names.append(f"{yr}_{suffix}")

    # Parser les lignes de données (après la dernière ligne d'en-tête)
    last_header_idx = max(i for i, _ in header_lines)

    # Regex pour détecter une ligne de données numériques
    # Une ligne de données a : label + une suite de valeurs (nombres décimaux, …, -)
    val_rx = re.compile(r'[-−]?\d+[,\.]\d+|…|\.\.\.|#DIV/0!')

    rows = []
    for line in lines[last_header_idx + 1:]:
        line = line.strip()
        if not line:
            continue
        if re.fullmatch(r'\d{1,3}', line):
            continue  # numéro de page
        if line.startswith('Source') or line.startswith('*') or line.startswith('Note'):
            continue

        vals = val_rx.findall(line)
        if not vals:
            # Titre de section
            rows.append([line] + [''] * len(col_names))
            continue

        # Séparer le libellé des valeurs
        # Chercher le début des valeurs : premier match de val_rx
        m = val_rx.search(line)
        label = line[:m.start()].strip() if m else line
        # Supprimer trailing parenthèses/astérisques du label
        label = re.sub(r'\s*[*†‡]+$', '', label).strip()

        # Aligner les valeurs sur les colonnes
        n = len(col_names)
        vals_padded = (vals + [''] * n)[:n]
        rows.append([label] + vals_padded)

    if not rows:
        print(f"  ⚠ {table_label[:40]} : aucune ligne de données")
        return None, meta

    cols = ['Indicateur'] + col_names
    # Normaliser la longueur des lignes
    max_len = max(len(r) for r in rows)
    for r in rows:
        while len(r) < max_len:
            r.append('')
    rows = [r[:max_len] for r in rows]

    df = pd.DataFrame(rows, columns=cols[:max_len])
    print(f"  ✓ {table_label[:40]} : {len(df)} lignes × {len(df.columns)} colonnes")
    return df, meta


# ==============================================================================
# BLOC 3 : ÉCRITURE EXCEL
# ==============================================================================

def build_formats(wb):
    """Définit tous les formats xlsxwriter du classeur."""
    base = {'font_name': 'Arial Narrow', 'font_size': 10}
    return {
        'main_title':  wb.add_format({**base, 'bold': True, 'font_size': 16,
                         'bg_color': '#1F4E79', 'font_color': '#FFFFFF',
                         'align': 'center', 'valign': 'vcenter', 'border': 1}),
        'bloc_title':  wb.add_format({**base, 'bold': True, 'font_size': 11,
                         'bg_color': '#2E75B6', 'font_color': '#FFFFFF',
                         'border': 1, 'valign': 'vcenter'}),
        'sub_header':  wb.add_format({**base, 'bold': True,
                         'bg_color': '#BDD7EE', 'border': 1, 'text_wrap': True}),
        'section':     wb.add_format({**base, 'bold': True, 'italic': True,
                         'bg_color': '#F2F2F2', 'border': 1}),
        'label_cell':  wb.add_format({**base, 'bg_color': '#DEEAF1',
                         'border': 1, 'bold': True, 'text_wrap': True}),
        'data_cell':   wb.add_format({**base, 'border': 1, 'align': 'right'}),
        # Cellule mise à jour (orange léger) — utilisée quand la valeur a changé par rapport au mois précédent
        'data_updated':wb.add_format({**base, 'border': 1, 'align': 'right',
                         'bg_color': '#FFE0B2'}),  # orange pastel
        'src_updated': wb.add_format({**base, 'border': 1, 'bg_color': '#FFE0B2'}),
        'placeholder': wb.add_format({**base, 'bg_color': '#F2F2F2',
                         'font_color': '#7F7F7F', 'border': 1,
                         'align': 'center', 'valign': 'vcenter', 'text_wrap': True}),
        'src_title':   wb.add_format({**base, 'bold': True, 'font_size': 12,
                         'bg_color': '#1F4E79', 'font_color': '#FFFFFF', 'border': 1}),
        'src_meta':    wb.add_format({**base, 'bg_color': '#DEEAF1',
                         'italic': True, 'border': 1}),
        'src_header':  wb.add_format({**base, 'bold': True,
                         'bg_color': '#BDD7EE', 'border': 1, 'text_wrap': True}),
        'src_cell':    wb.add_format({**base, 'border': 1}),
        'src_section': wb.add_format({**base, 'bold': True, 'italic': True,
                         'bg_color': '#F2F2F2', 'border': 1}),
    }


# ==============================================================================
# BLOC 3b : COMPARAISON AVEC L'EXCEL PRÉCÉDENT (surlignage orange)
# ==============================================================================

def load_previous_values(prev_xlsx_path):
    """
    Charge toutes les valeurs d'un Excel précédent sous forme de dict :
    { sheet_name: { (row_idx, col_idx): value_str } }
    Ne charge que les onglets source (pas Dashboard ni README).
    """
    if not prev_xlsx_path or not os.path.exists(prev_xlsx_path):
        return {}
    try:
        xl = pd.ExcelFile(prev_xlsx_path)
        result = {}
        for sheet in xl.sheet_names:
            if sheet.startswith('Source_'):
                df = xl.parse(sheet, header=None)
                sheet_dict = {}
                for ri, row in df.iterrows():
                    for ci, val in enumerate(row):
                        if pd.notna(val) and str(val).strip():
                            sheet_dict[(ri, ci)] = str(val).strip()
                result[sheet] = sheet_dict
        return result
    except Exception as e:
        print(f"  ⚠ Impossible de charger l'Excel précédent : {e}")
        return {}


def is_updated(sheet_name, row_idx, col_idx, new_val, prev_values):
    """
    Retourne True si new_val est différent de la valeur précédente
    pour cette cellule (sheet, row, col).
    """
    if not prev_values or sheet_name not in prev_values:
        return False
    prev = prev_values[sheet_name].get((row_idx, col_idx))
    if prev is None:
        return True  # Cellule nouvelle → surligner
    new_str = str(new_val).strip() if new_val is not None else ''
    return new_str != prev


def write_source_sheet(wb, sheet_name, df, meta, fmt,
                       extra_df=None, extra_title="", prev_values=None):
    """
    Crée un onglet source avec traçabilité + données.
    prev_values : dict {(row_idx, col_idx): value_str} de l'Excel précédent.
    Les cellules nouvelles ou modifiées sont surlignées en orange.
    """
    ws = wb.add_worksheet(sheet_name[:31])
    ws.set_column('A:A', 45)

    # En-tête de traçabilité
    ws.merge_range('A1:N1', f"SOURCE : {meta['table_label']}", fmt['src_title'])
    ws.write('A2', 'Fichier PDF :',        fmt['src_meta'])
    ws.merge_range('B2:N2', meta['source_pdf'],    fmt['src_meta'])
    ws.write('A3', 'Page(s) :',            fmt['src_meta'])
    ws.merge_range('B3:N3', meta['page_pdf'],      fmt['src_meta'])
    ws.write('A4', 'Date d\'extraction :', fmt['src_meta'])
    ws.merge_range('B4:N4', meta['extraction_date'], fmt['src_meta'])
    ws.write('A5', 'Note :',               fmt['src_meta'])
    ws.merge_range('B5:N5',
        'Données extraites automatiquement. Vérifier la cohérence avec le PDF source.',
        fmt['src_meta'])
    # Légende surlignage
    ws.write('A6', '🟠 = Valeur nouvelle ou modifiée par rapport au mois précédent',
             fmt['src_meta'])

    current_row = 7  # Données à partir de la ligne 8 (0-indexed = 7)

    # Récupérer les valeurs précédentes pour cet onglet
    prev = prev_values.get(sheet_name, {}) if prev_values else {}

    def write_df(ws, df, start_row):
        """Écrit un DataFrame, en orange les cellules qui ont changé."""
        if df is None or df.empty:
            ws.merge_range(start_row, 0, start_row, 13,
                           '⚠ Données non extraites.', fmt['src_meta'])
            return start_row + 1

        # Largeurs de colonnes automatiques
        for ci, col in enumerate(df.columns):
            max_w = max(len(str(col)),
                        max((len(str(v)) for v in df[col]), default=0))
            ws.set_column(ci, ci, min(max_w + 2, 50))

        # En-têtes
        for ci, col in enumerate(df.columns):
            ws.write(start_row, ci, col, fmt['src_header'])
        start_row += 1

        # Données
        for _, row in df.iterrows():
            vals = list(row)
            first_val = str(vals[0]) if vals else ''
            all_empty = all(str(v).strip() == '' or str(v) == 'nan' for v in vals[1:])

            if all_empty and first_val.strip() and first_val != 'nan':
                ws.merge_range(start_row, 0, start_row, len(df.columns) - 1,
                               first_val, fmt['src_section'])
            else:
                ws.write(start_row, 0, first_val, fmt['src_cell'])
                for ci, v in enumerate(vals[1:], start=1):
                    v_str = str(v) if (v != '' and str(v) != 'nan') else ''
                    # Choisir le format selon si la valeur a changé
                    changed = is_updated(sheet_name, start_row, ci, v_str, prev_values)
                    cell_fmt = fmt['src_updated'] if changed else fmt['src_cell']
                    ws.write(start_row, ci, v_str, cell_fmt)
            start_row += 1

        return start_row

    current_row = write_df(ws, df, current_row)

    if extra_df is not None:
        current_row += 1
        ws.merge_range(current_row, 0, current_row, 13, extra_title, fmt['bloc_title'])
        current_row += 1
        current_row = write_df(ws, extra_df, current_row)

    return ws


# ==============================================================================
# BLOC 4 : DASHBOARD
# ==============================================================================

def build_dashboard(wb, ex, period_str, fmt):
    """Construit l'onglet 01_Dashboard avec toutes les données extraites."""
    ws = wb.add_worksheet('01_Dashboard')
    ws.set_paper(9)        # A4
    ws.set_landscape()
    ws.fit_to_pages(1, 0)
    ws.set_margins(left=0.5, right=0.5, top=0.7, bottom=0.7)

    # Largeurs colonnes (A=libellé large, B-N=valeurs)
    ws.set_column(0, 0, 38)   # Libellé
    for ci in range(1, 14):
        ws.set_column(ci, ci, 11)

    NCOLS = 14
    row = 0

    def merge_title(r, text, fk, height=30):
        ws.merge_range(r, 0, r, NCOLS - 1, text, fmt[fk])
        ws.set_row(r, height)

    def write_df_block(ws, df, row, max_rows=None):
        """Écrit les lignes d'un DataFrame dans le dashboard."""
        if df is None or df.empty:
            return row
        n_data_cols = min(len(df.columns) - 1, NCOLS - 1)
        # En-têtes
        ws.write(row, 0, df.columns[0], fmt['sub_header'])
        for ci, col in enumerate(df.columns[1:n_data_cols + 1], start=1):
            ws.write(row, ci, col, fmt['sub_header'])
        row += 1

        limit = max_rows if max_rows else len(df)
        for idx, r_data in enumerate(df.itertuples(index=False)):
            if idx >= limit:
                break
            vals = list(r_data)
            first = str(vals[0]) if vals else ''
            rest  = [str(v) if v != '' else '' for v in vals[1:n_data_cols + 1]]
            all_empty = all(v == '' for v in rest)

            if all_empty and first:
                ws.merge_range(row, 0, row, NCOLS - 1, first, fmt['section'])
            else:
                ws.write(row, 0, first, fmt['label_cell'])
                for ci, v in enumerate(rest, start=1):
                    ws.write(row, ci, v, fmt['data_cell'])
            row += 1
        return row

    # ── TITRE ──────────────────────────────────────────────────
    merge_title(row, f"TABLEAU DE BORD CEMAC — {period_str.upper().replace('_', ' ')}", 'main_title', 35)
    row += 1
    merge_title(row, f"Dernière mise à jour : {datetime.now().strftime('%d/%m/%Y')}", 'sub_header', 16)
    row += 2

    # ── BLOC FIXE : MINISTRES ──────────────────────────────────
    merge_title(row, "BLOC FIXE — MINISTRES DE L'ÉCONOMIE ET DES FINANCES (ZONE CEMAC)", 'bloc_title', 20)
    row += 1
    ws.write(row, 0, "Pays", fmt['sub_header'])
    ws.merge_range(row, 1, row, 4, "Ministre", fmt['sub_header'])
    ws.merge_range(row, 5, row, NCOLS - 1, "Titre officiel", fmt['sub_header'])
    row += 1
    ministres = [
        ("🇨🇲 Cameroun",           "Alamine Ousmane MEY",    "Ministre de l'Économie, de la Planification et de l'Aménagement du territoire (MINEPAT)"),
        ("🇨🇫 Centrafrique (RCA)",  "Richard FILAKOTA",       "Ministre de l'Économie, du Plan et de la Coopération internationale"),
        ("🇨🇬 Congo (Brazzaville)", "Ludovic NGATSÉ",         "Ministre de l'Économie, du Plan, de la Statistique et de l'Intégration régionale"),
        ("🇬🇦 Gabon",               "Henri-Claude OYIMA",     "Ministre d'État, Ministre de l'Économie, des Finances, de la Dette et des Participations"),
        ("🇬🇶 Guinée Équatoriale",  "Ivan BACALE EBE MOLINA", "Ministre des Finances, de la Planification et du Développement économique"),
        ("🇹🇩 Tchad",               "Tahir Hamid NGUILIN",    "Ministre d'État, Ministre des Finances, du Budget, de l'Économie et du Plan"),
    ]
    for pays, nom, titre in ministres:
        ws.write(row, 0, pays,  fmt['data_cell'])
        ws.merge_range(row, 1, row, 4, nom,   fmt['data_cell'])
        ws.merge_range(row, 5, row, NCOLS - 1, titre, fmt['data_cell'])
        row += 1
    row += 1

    # ── BLOC MONÉTAIRE ─────────────────────────────────────────
    merge_title(row, "BLOC MONÉTAIRE", 'bloc_title', 20)
    row += 1

    # Réserves — placeholder image
    ws.write(row, 0, "Évolution des réserves (G8 RPM p.18)", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1,
        "[PLACEHOLDER IMAGE] Graphique 8 p.18 du RPM — Insérer capture d'écran", fmt['placeholder'])
    row += 1

    ws.write(row, 0, "Réserves par pays G28 (Annexe p.24)", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1,
        "[PLACEHOLDER IMAGE] Graphique 28 Annexe p.24 — Insérer capture d'écran", fmt['placeholder'])
    row += 1

    # TCEM T18 — Niveaux
    merge_title(row, "Agrégats monnaie & crédit CEMAC — T18 Niveaux (Annexe p.26, en millions FCFA)", 'sub_header', 18)
    row += 1
    df_tcem_niv = ex.get('tcem_niv')
    row = write_df_block(ws, df_tcem_niv, row, max_rows=None)

    row += 1
    merge_title(row, "Agrégats monnaie & crédit CEMAC — T18 Variations annuelles (Annexe p.26)", 'sub_header', 18)
    row += 1
    df_tcem_var = ex.get('tcem_var')
    row = write_df_block(ws, df_tcem_var, row, max_rows=None)

    # Prix pétrole — externe
    row += 1
    ws.write(row, 0, "Prix du pétrole (Brent USD/baril)", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1, "[SOURCE EXTERNE — Bloomberg] Renseigner manuellement", fmt['placeholder'])
    row += 1

    # TCER — externe
    ws.write(row, 0, "TCER", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1, "[SOURCE EXTERNE — SADEV/FMI] Assemblées Annuelles / Printemps", fmt['placeholder'])
    row += 1

    # Taux directeurs T25
    row += 1
    merge_title(row, "Taux directeurs BEAC/BCE — T25 (Annexe p.33)", 'sub_header', 18)
    row += 1
    ws.write(row, 0, "[PLACEHOLDER IMAGE]", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1,
        "Graphique 15 p.23 du RPM — Insérer capture d'écran | Données tabulaires ci-dessous", fmt['placeholder'])
    row += 1
    df_td = ex.get('taux_directeurs')
    row = write_df_block(ws, df_td, row, max_rows=None)

    row += 1

    # ── BLOC BALANCE DES PAIEMENTS ─────────────────────────────
    merge_title(row, "BLOC BALANCE DES PAIEMENTS", 'bloc_title', 20)
    row += 1
    merge_title(row, "T16 Soldes Balance des paiements (% du PIB) — Annexe p.22", 'sub_header', 18)
    row += 1
    df_bop = ex.get('bop')
    row = write_df_block(ws, df_bop, row, max_rows=None)
    row += 1
    ws.write(row, 0, "→ Données pays détaillées", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1, "Voir onglets Source_Cameroun, Source_Centrafrique, etc.", fmt['data_cell'])
    row += 2

    # ── BLOC BUDGÉTAIRE ────────────────────────────────────────
    merge_title(row, "BLOC BUDGÉTAIRE", 'bloc_title', 20)
    row += 1
    merge_title(row, "T32 CEMAC Principaux indicateurs économiques, financiers et sociaux — Annexe p.39", 'sub_header', 18)
    row += 1
    df_budg = ex.get('budg')
    row = write_df_block(ws, df_budg, row, max_rows=None)
    row += 1
    ws.write(row, 0, "→ Données pays détaillées", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1, "Voir onglets Source_Cameroun, Source_Centrafrique, etc.", fmt['data_cell'])
    row += 2

    # ── BLOC BANCAIRE ──────────────────────────────────────────
    merge_title(row, "BLOC BANCAIRE", 'bloc_title', 20)
    row += 1
    ws.write(row, 0, "Nb établissements en dépendance refinancement", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1,
        "[SOURCE EXTERNE — Note CPM] Renseigner manuellement", fmt['placeholder'])
    row += 1
    ws.write(row, 0, "Part financement des États dans le bilan bancaire", fmt['label_cell'])
    ws.merge_range(row, 1, row, NCOLS - 1,
        "[SOURCE EXTERNE — Dossier COBAC trimestriel] Renseigner manuellement", fmt['placeholder'])
    row += 1

    return ws


# ==============================================================================
# BLOC 5 : README
# ==============================================================================

def build_readme(wb, fmt):
    ws = wb.add_worksheet('README_Mise_a_jour')
    ws.set_column('A:B', 80)

    ws.merge_range('A1:B1', "INSTRUCTIONS DE MISE À JOUR MENSUELLE DU DASHBOARD CEMAC", fmt['src_title'])
    ws.set_row(0, 25)

    instructions = [
        ("ÉTAPE 1 — Préparer les fichiers PDF",
         "Placer dans ce dossier les deux PDFs du nouveau mois :\n"
         "  • 'Point 3_RPM [mois] [année]_vf_.pdf'\n"
         "  • 'Point 3_annexe_Tableau_de_bord_CPM_[mois]_[année]_vN.pdf'\n"
         "Les anciens fichiers peuvent être archivés (le script prend le dernier en date)."),
        ("ÉTAPE 2 — Lancer le script",
         "Ouvrir un Terminal dans ce dossier et exécuter :\n"
         "  python generate_cemac_dashboard.py\n"
         "→ Un nouveau fichier 'Dashboard_[mois]_[année].xlsx' est créé."),
        ("ÉTAPE 3 — Compléter manuellement",
         "Dans '01_Dashboard', renseigner les zones grises [PLACEHOLDER] :\n"
         "  • Graphique 8 (réserves), G28 (réserves pays), Graphique 15 (taux directeurs)\n"
         "  • Prix du pétrole (Bloomberg — Brent USD/baril)\n"
         "  • TCER (SADEV/FMI — après Assemblées Annuelles ou de Printemps)\n"
         "  • Données COBAC (Note CPM + Dossier COBAC trimestriel)"),
        ("ÉTAPE 4 — Exporter en PDF",
         "Dans Excel : Fichier → Exporter → Créer un PDF/XPS\n"
         "Sélectionner uniquement l'onglet '01_Dashboard'. Format A4 paysage pré-configuré."),
        ("CARTOGRAPHIE DES SOURCES",
         "T18 TCEM Niveaux & Variations → Annexe p.26\n"
         "T16 Balance des paiements     → Annexe p.22\n"
         "T25 Taux directeurs            → Annexe p.33\n"
         "T32 Indicateurs CEMAC          → Annexe p.39\n"
         "T33 Cameroun                   → Annexe p.40\n"
         "T34 Centrafrique               → Annexe p.41\n"
         "T35 Congo                      → Annexe p.42\n"
         "T36 Gabon                      → Annexe p.43\n"
         "T37 Guinée Équatoriale         → Annexe p.44\n"
         "T38 Tchad                      → Annexe p.45"),
    ]

    r = 2
    for title, body in instructions:
        ws.write(r, 0, title, fmt['src_header'])
        ws.write(r, 1, "", fmt['src_cell'])
        ws.set_row(r, 18)
        r += 1
        ws.write(r, 0, "", fmt['src_cell'])
        fmt_body = wb.add_format({'font_name': 'Arial Narrow', 'font_size': 10,
                                   'border': 1, 'text_wrap': True, 'valign': 'top'})
        ws.merge_range(r, 0, r, 1, body, fmt_body)
        ws.set_row(r, 80)
        r += 2


# ==============================================================================
# BLOC 6 : MAIN
# ==============================================================================

if __name__ == "__main__":
    print("=" * 60)
    print("  GÉNÉRATION DASHBOARD CEMAC")
    print("=" * 60)

    rpm, annexe, period = find_latest_pdfs()
    print(f"\n📁 Période détectée : {period}")
    print(f"   RPM    : {os.path.basename(rpm) if rpm else '⚠ NON TROUVÉ'}")
    print(f"   Annexe : {os.path.basename(annexe) if annexe else '⚠ NON TROUVÉ'}")

    if not annexe:
        print("\n❌ Fichier Annexe introuvable. Vérifier le dossier.")
        exit(1)

    print("\n📊 Extraction des tableaux...")
    ex = {}

    # T18 — TCEM — p.26 (idx 25)
    df_niv, df_var, meta_tcem = extract_tcem(annexe, 25)
    ex['tcem_niv']  = df_niv
    ex['tcem_var']  = df_var
    ex['meta_tcem'] = meta_tcem

    # T16 — Balance des paiements — p.22 (idx 21)
    df_bop, meta_bop = extract_bop(annexe, 21)
    ex['bop']      = df_bop
    ex['meta_bop'] = meta_bop

    # T25 — Taux directeurs — p.33 (idx 32)
    df_td, meta_td = extract_taux_directeurs(annexe, 32)
    ex['taux_directeurs'] = df_td
    ex['meta_td']         = meta_td

    # T32 — CEMAC global — p.39 (idx 38)
    df_budg, meta_budg = extract_indicators_table(
        annexe, 38,
        table_label="T32 — CEMAC : Principaux indicateurs économiques, financiers et sociaux",
        page_label="p.39 de l'Annexe")
    ex['budg']      = df_budg
    ex['meta_budg'] = meta_budg

    # T33-T38 — Données pays — p.40-45 (idx 39-44)
    pays_config = [
        ('cameroun',    39, "T33 — CAMEROUN",           "p.40 de l'Annexe"),
        ('centrafrique',40, "T34 — CENTRAFRIQUE",        "p.41 de l'Annexe"),
        ('congo',       41, "T35 — CONGO",               "p.42 de l'Annexe"),
        ('gabon',       42, "T36 — GABON",               "p.43 de l'Annexe"),
        ('guinee',      43, "T37 — GUINÉE ÉQUATORIALE",  "p.44 de l'Annexe"),
        ('tchad',       44, "T38 — TCHAD",               "p.45 de l'Annexe"),
    ]
    for key, idx, label, page_label in pays_config:
        df_p, meta_p = extract_indicators_table(annexe, idx,
            table_label=f"{label} : Principaux indicateurs économiques, financiers et sociaux",
            page_label=page_label)
        ex[f'pays_{key}']      = df_p
        ex[f'meta_pays_{key}'] = meta_p

    # Création du fichier Excel
    out_path = os.path.join(BASE_DIR, f"Dashboard_{period}.xlsx")
    print(f"\n📝 Création : {os.path.basename(out_path)}")

    # ── Comparaison avec le mois précédent (surlignage orange) ──
    # On cherche tous les Dashboard_*.xlsx existants SAUF celui qu'on va créer
    existing_dashboards = sorted(
        f for f in glob.glob(os.path.join(BASE_DIR, "Dashboard_*.xlsx"))
        if os.path.abspath(f) != os.path.abspath(out_path)
    )
    prev_xlsx = existing_dashboards[-1] if existing_dashboards else None
    if prev_xlsx:
        print(f"   📋 Comparaison avec : {os.path.basename(prev_xlsx)}")
    else:
        print("   ℹ️  Aucun fichier précédent trouvé — pas de surlignage orange.")
    prev_values = load_previous_values(prev_xlsx)

    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt = build_formats(wb)

        # Dashboard principal
        build_dashboard(wb, ex, period, fmt)

        # Onglets sources (avec surlignage des cellules modifiées)
        write_source_sheet(wb, "Source_TCEM", ex['tcem_niv'], ex['meta_tcem'], fmt,
                           extra_df=ex['tcem_var'],
                           extra_title="T18 — Variations annuelles (%)",
                           prev_values=prev_values)

        write_source_sheet(wb, "Source_BOP",  ex['bop'],      ex['meta_bop'],  fmt,
                           prev_values=prev_values)
        write_source_sheet(wb, "Source_TauxDir", ex['taux_directeurs'], ex['meta_td'], fmt,
                           prev_values=prev_values)
        write_source_sheet(wb, "Source_Budg_CEMAC", ex['budg'], ex['meta_budg'], fmt,
                           prev_values=prev_values)

        for key, idx, label, page_label in pays_config:
            write_source_sheet(wb, f"Source_{key.capitalize()}",
                               ex[f'pays_{key}'], ex[f'meta_pays_{key}'], fmt,
                               prev_values=prev_values)

        build_readme(wb, fmt)

    print(f"\n✅ Dashboard généré : {out_path}")
    print("=" * 60)
