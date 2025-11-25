# app.py
# ------------------------------------------------------
# Streamlit – Générateur d'Excel récapitulatif de paramétrage Hopia
# Entrées possibles :
#   - Fichier brut (TXT / CSV / XLSX)
#   - Copier-coller de texte brut (y compris format vertical type Urgentistes)
# ------------------------------------------------------

import io
import re
from typing import Tuple

import pandas as pd
import streamlit as st

# ------------------------------------------------------
# Configuration Streamlit
# ------------------------------------------------------
st.set_page_config(page_title="Hopia – Récap Paramétrage", layout="wide")

st.title("📊 Hopia – Générateur d’Excel récapitulatif de paramétrage à partir du Back-Office")

# ------------------------------------------------------
# Couleurs pour l'export Excel
# ------------------------------------------------------
COLOR_DURE = "#ffcccc"
COLOR_MOY = "#ffe5b4"
COLOR_SOFT = "#ccffcc"
COLOR_HEADER = "#003366"
COLOR_HEADER_TXT = "#FFFFFF"

# ------------------------------------------------------
# Parseur pour format vertical
# ------------------------------------------------------
def parse_vertical_blocks(content: str) -> pd.DataFrame:
    lines = [l.strip() for l in content.splitlines()]
    lines = [l for l in lines if l]

    i = 0
    while i < len(lines) and not lines[i].isdigit():
        i += 1

    def looks_like_pk(s: str): return s.isdigit()

    def looks_like_priority(s: str):
        return bool(re.search(r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1|<|>|≤)", s))

    records = []

    while i < len(lines):
        if not looks_like_pk(lines[i]):
            i += 1
            continue

        pk = lines[i]
        i += 1
        if i >= len(lines):
            break

        line2 = lines[i]
        i += 1

        parts = re.split(r"\t+| {2,}", line2, maxsplit=1)
        intitule = parts[0].strip()
        type_val = parts[1].strip() if len(parts) == 2 else ""

        prio_lines = []
        equipe = ""

        while i < len(lines) and not looks_like_pk(lines[i]):
            l = lines[i]

            if not looks_like_priority(l) and prio_lines:
                equipe = l
                i += 1
                break

            if not looks_like_priority(l) and not prio_lines:
                equipe = l
                i += 1
                break

            prio_lines.append(l)
            i += 1

        priorite = " ".join(prio_lines).strip()

        records.append({
            "PK": pk,
            "Intitulé": intitule,
            "Type": type_val,
            "Priorités": priorite,
            "Équipes": equipe
        })

    return pd.DataFrame(records)

# ------------------------------------------------------
# Lecture texte collé
# ------------------------------------------------------
def read_text_content(content: str):
    content = content.strip()
    if not content:
        return None

    df_vert = parse_vertical_blocks(content)
    if not df_vert.empty:
        return df_vert

    for sep in [";", "\t", ","]:
        try:
            df = pd.read_csv(io.StringIO(content), sep=sep, engine="python")
            if df.shape[1] >= 2:
                return df
        except:
            pass

    try:
        return pd.read_csv(io.StringIO(content), delim_whitespace=True, engine="python")
    except:
        st.error("Impossible d'interpréter le texte collé.")
        return None

# ------------------------------------------------------
# Lecture fichier uploadé
# ------------------------------------------------------
def read_uploaded_file(uploaded_file):
    name = uploaded_file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep=None, engine="python")

    if name.endswith(".txt"):
        try:
            content = uploaded_file.read().decode("utf-8")
        except:
            content = uploaded_file.read().decode("latin-1")
        return read_text_content(content)

    st.error("Format non supporté.")
    return None

# ------------------------------------------------------
# Normalisation colonnes
# ------------------------------------------------------
def normalize_cols(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["PK", "Type", "Priorités", "Équipes"]:
        if col not in df.columns:
            df[col] = None

    if "Intitulé" not in df.columns:
        df["Intitulé"] = df["Type"]

    front = ["PK", "Intitulé", "Type", "Priorités", "Équipes"]
    return df[[c for c in front if c in df.columns] +
              [c for c in df.columns if c not in front]]

# ------------------------------------------------------
# Extraction tokens priorités
# ------------------------------------------------------
def token_set(priorites: str):
    if pd.isna(priorites):
        return set()
    txt = str(priorites).upper()
    return set(re.findall(r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1)", txt))

# ------------------------------------------------------
# Mapping vers Niveau
# ------------------------------------------------------
def map_level(row):
    type_txt = str(row["Type"]).lower()
    toks = token_set(row["Priorités"])

    is_remplissage = "remplissage des postes" in type_txt

    if is_remplissage:
        if {"PRIORITY_1", "HARD", "STRONG_1"} & toks:
            return "DURE", ""
        if {"PRIORITY_2", "STRONG_2", "STRONG_3"} & toks:
            return "MOYENNE", ""
        if {"PRIORITY_3", "DEFAULT_PENALTY"} & toks or any(t.startswith("SOFT_") for t in toks):
            return "SOUPLE", ""
        return "SOUPLE", ""

    if "HARD" in toks or "HARD_LOWER" in toks:
        return "DURE", ""
    if any(t.startswith("STRONG_") for t in toks) or any(t.startswith("PRIORITY_") for t in toks) \
       or "DEFAULT_PENALTY" in toks or "PRIVATE_ALGO_1" in toks:
        return "MOYENNE", ""
    if any(t.startswith("SOFT_") for t in toks):
        return "SOUPLE", ""

    return "SOUPLE", ""

# ------------------------------------------------------
# Couleur par Niveau
# ------------------------------------------------------
def color_for_level(level: str):
    l = level.upper()
    if l == "DURE":
        return COLOR_DURE
    if l == "MOYENNE":
        return COLOR_MOY
    return COLOR_SOFT

# ------------------------------------------------------
# Construction Excel
# (Résumé en DERNIÈRE FEUILLE et renommé "Résumé")
# ------------------------------------------------------
def to_excel_bytes(df_autres, df_remp, df_summary):
    import xlsxwriter

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({
            "bold": True,
            "bg_color": COLOR_HEADER,
            "font_color": COLOR_HEADER_TXT,
            "align": "center"
        })

        # === Feuille 1 – Autres contraintes ===
        df_autres.to_excel(writer, sheet_name="Paramétrage – Autres", index=False)
        ws2 = writer.sheets["Paramétrage – Autres"]

        for col_idx, col in enumerate(df_autres.columns):
            ws2.write(0, col_idx, col, fmt_header)
            ws2.set_column(col_idx, col_idx, 28)

        # Coloration seulement colonne Niveau
        niveau_col = df_autres.columns.get_loc("Niveau")
        for row_idx in range(1, len(df_autres) + 1):
            val = str(df_autres.iloc[row_idx - 1]["Niveau"])
            ws2.write(row_idx, niveau_col, val,
                      wb.add_format({"bg_color": color_for_level(val)}))

        # === Feuille 2 – Remplissage des postes ===
        df_remp.to_excel(writer, sheet_name="Paramétrage – Remplissage", index=False)
        ws3 = writer.sheets["Paramétrage – Remplissage"]

        for col_idx, col in enumerate(df_remp.columns):
            ws3.write(0, col_idx, col, fmt_header)
            ws3.set_column(col_idx, col_idx, 28)

        niveau_col = df_remp.columns.get_loc("Niveau")
        for row_idx in range(1, len(df_remp) + 1):
            val = str(df_remp.iloc[row_idx - 1]["Niveau"])
            ws3.write(row_idx, niveau_col, val,
                      wb.add_format({"bg_color": color_for_level(val)}))

        # === Feuille 3 – Résumé (dernière feuille) ===
        df_summary.to_excel(writer, sheet_name="Résumé", index=False)
        ws = writer.sheets["Résumé"]

        for col_idx, col in enumerate(df_summary.columns):
            ws.write(0, col_idx, col, fmt_header)
            ws.set_column(col_idx, col_idx, 25)

        # Coloration colonnes par Niveau
        col_map = {"DURE": COLOR_DURE, "MOYENNE": COLOR_MOY, "SOUPLE": COLOR_SOFT}
        for col_name, bg in col_map.items():
            if col_name in df_summary.columns:
                col_idx = df_summary.columns.get_loc(col_name)
                ws.conditional_format(
                    1, col_idx, len(df_summary) + 50, col_idx,
                    {"type": "no_errors", "format": wb.add_format({"bg_color": bg})}
                )

    return output.getvalue()

# ------------------------------------------------------
# Interface utilisateur
# ------------------------------------------------------
uploaded = st.file_uploader(
    "📁 Importer un fichier de paramétrage",
    type=["txt", "csv", "xlsx", "xls"],
)

text_pasted = st.text_area(
    "✂️ Ou collez ici l'export brut du back-office ( copier-coller) :",
    height=200,
)

df_raw = None
if uploaded:
    df_raw = read_uploaded_file(uploaded)
elif text_pasted.strip():
    df_raw = read_text_content(text_pasted)

if df_raw is not None:
    try:
        df_norm = normalize_cols(df_raw)

        # Calcul des niveaux
        levels = df_norm.apply(map_level, axis=1, result_type="expand")
        df_norm["Niveau"] = levels[0]

        # Filtrage : suppression des demandes
        df_filtered = df_norm[
            ~df_norm["Type"].astype(str).isin(["Demandes d'absence", "Demandes de travail"])
        ].copy()

        # Niveau catégorisé
        niveau_order = pd.CategoricalDtype(
            ["DURE", "MOYENNE", "SOUPLE"], ordered=True
        )
        df_filtered["Niveau"] = df_filtered["Niveau"].astype(niveau_order)

        # Résumé
        df_summary = (
            df_filtered
            .pivot_table(index="Type", columns="Niveau", values="PK", aggfunc="count", fill_value=0)
            .reset_index()
            .rename(columns={"Type": "Rubrique"})
        )

        df_filtered["Equipe"] = df_filtered["Équipes"]
        is_rem = df_filtered["Type"].str.lower() == "remplissage des postes"

        df_autres = df_filtered[~is_rem][["Intitulé", "Type", "Equipe", "Niveau"]].sort_values(["Type", "Niveau", "Intitulé"])
        df_remp = df_filtered[is_rem][["Intitulé", "Type", "Equipe", "Niveau"]].sort_values(["Type", "Niveau", "Intitulé"])

        excel_bytes = to_excel_bytes(df_autres, df_remp, df_summary)

        st.download_button(
            "⬇️ Télécharger l’Excel récapitulatif",
            data=excel_bytes,
            file_name="Parametrage_Harmonise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Erreur : {e}")
else:
    st.info("Charge un fichier ou colle ton texte.")
