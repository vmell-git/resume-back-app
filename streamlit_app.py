# app.py
# ------------------------------------------------------
# Streamlit – Générateur d'Excel récapitulatif de paramétrage
# Entrée : CSV/XLSX avec colonnes (idéales) :
#   PK | Intitulé | Type | Priorités | Équipes
# - 'Intitulé' peut manquer : on le créera vide.
# - Le mapping DURE/MOYENNE/SOUPLE suit les règles Hopia :
#   • Hors "Remplissage des postes" :
#       - DURE     : contient HARD ou HARD_LOWER
#       - MOYENNE  : contient STRONG_* ou PRIORITY_* ou DEFAULT_PENALTY
#       - SOUPLE   : contient SOFT_*
#   • Pour "Remplissage des postes" :
#       - DURE     : PRIORITY_1 ou HARD ou STRONG_1
#       - MOYENNE  : PRIORITY_2 ou STRONG_2/3
#       - SOUPLE   : PRIORITY_3 ou DEFAULT_PENALTY ou SOFT_*
# Sortie : Excel téléchargeable avec mise en forme/couleurs.
# ------------------------------------------------------

import io
import re
from typing import Tuple

import pandas as pd
import streamlit as st

# ------------------------------------------------------
# UI – Sidebar
# ------------------------------------------------------
st.set_page_config(page_title="Hopia – Récap Paramétrage", layout="wide")
st.sidebar.title("⚙️ Options")
st.sidebar.markdown(
    "- Format attendu : CSV ou XLSX\n"
    "- Colonnes recommandées : **PK, Intitulé, Type, Priorités, Équipes**\n"
    "- Les autres colonnes seront conservées."
)

# Couleurs export Excel (hex)
COLOR_DURE = "#ffcccc"     # rouge clair
COLOR_MOY = "#ffe5b4"      # orange clair
COLOR_SOFT = "#ccffcc"     # vert clair
COLOR_HEADER = "#003366"   # bleu foncé
COLOR_HEADER_TXT = "#FFFFFF"

# ------------------------------------------------------
# Helpers
# ------------------------------------------------------
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    # Strip / harmonise basic column names
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # Auto-add missing known columns
    needed = ["PK", "Intitulé", "Type", "Priorités", "Équipes"]
    for col in needed:
        if col not in df.columns:
            if col == "Intitulé":
                df[col] = ""
            else:
                df[col] = None
    # Reorder (keep extras after)
    front = [c for c in ["PK", "Intitulé", "Type", "Priorités", "Équipes"] if c in df.columns]
    others = [c for c in df.columns if c not in front]
    return df[front + others]

def token_set(priorites: str) -> set:
    if pd.isna(priorites):
        return set()
    txt = str(priorites).upper()
    # Extract common tokens
    parts = re.findall(r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY)", txt)
    return set(parts)

def map_level(row) -> Tuple[str, str]:
    """Return (Niveau, Règle utilisée) based on Type & Priorités."""
    type_txt = str(row.get("Type", "")).strip().lower()
    prio_raw = str(row.get("Priorités", ""))
    toks = token_set(prio_raw)

    is_remplissage = "remplissage des postes" in type_txt

    # Default fallback
    niveau = "SOU PLE"
    rule = "SOFT_* → SOUPLE (fallback)"

    if is_remplissage:
        # Remplissage des postes : DURE > MOYENNE > SOUPLE
        if {"PRIORITY_1", "HARD", "STRONG_1"} & toks:
            return "DURE", "Remplissage : PRIORITY_1/HARD/STRONG_1 → DURE"
        if {"PRIORITY_2", "STRONG_2", "STRONG_3"} & toks:
            return "MOYENNE", "Remplissage : PRIORITY_2/STRONG_2/STRONG_3 → MOYENNE"
        if {"PRIORITY_3", "DEFAULT_PENALTY"} & toks or any(t.startswith("SOFT_") for t in toks):
            return "SOU PLE", "Remplissage : PRIORITY_3/DEFAULT_PENALTY/SOFT_* → SOUPLE"
        # If no known token, leave as SOUPLE
        return niveau, rule
    else:
        # Hors remplissage : HARD = DURE, STRONG/PRIORITY/DEFAULT = MOYENNE, SOFT = SOUPLE
        if "HARD" in toks or "HARD_LOWER" in toks:
            return "DURE", "Hors remplissage : HARD/HARD_LOWER → DURE"
        if any(t.startswith("STRONG_") for t in toks) or any(t.startswith("PRIORITY_") for t in toks) or "DEFAULT_PENALTY" in toks:
            return "MOYENNE", "Hors remplissage : STRONG_*/PRIORITY_*/DEFAULT → MOYENNE"
        if any(t.startswith("SOFT_") for t in toks):
            return "SOU PLE", "Hors remplissage : SOFT_* → SOUPLE"
        # If no known token, keep SOUPLE
        return niveau, rule

def color_for_level(level: str) -> str:
    l = (level or "").upper().replace("É", "E")
    if "DURE" in l:
        return COLOR_DURE
    if "MOY" in l:  # MOYENNE
        return COLOR_MOY
    return COLOR_SOFT

def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    # Count by rubrique (Type) and niveau
    tmp = df.copy()
    tmp["Niveau"] = tmp["Niveau"].fillna("SOU PLE")
    piv = pd.pivot_table(
        tmp,
        index="Type",
        columns="Niveau",
        values="PK",
        aggfunc="count",
        fill_value=0
    )
    piv = piv.reindex(columns=["DURE", "MOYENNE", "SOU PLE"], fill_value=0)
    piv["Total"] = piv.sum(axis=1)
    piv = piv.reset_index().rename(columns={"Type": "Rubrique"})
    return piv

def to_excel_bytes(df_summary: pd.DataFrame, df_full: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Sheet 1 – Résumé
        df_summary.to_excel(writer, sheet_name="Résumé par rubrique", index=False)
        ws = writer.sheets["Résumé par rubrique"]
        # Header format
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": COLOR_HEADER, "font_color": COLOR_HEADER_TXT, "align": "center"})
        fmt_center = wb.add_format({"align": "center"})
        # Apply header style
        for col_idx, _ in enumerate(df_summary.columns):
            ws.write(0, col_idx, df_summary.columns[col_idx], fmt_header)
            ws.set_column(col_idx, col_idx, 22)
        # Shade columns by level
        col_map = { "DURE": COLOR_DURE, "MOYENNE": COLOR_MOY, "SOU PLE": COLOR_SOFT }
        for col_name, bg in col_map.items():
            if col_name in df_summary.columns:
                col_idx = df_summary.columns.get_loc(col_name)
                fmt_lvl = wb.add_format({"bg_color": bg})
                ws.conditional_format(1, col_idx, len(df_summary)+50, col_idx, {"type": "no_errors", "format": fmt_lvl})

        # Sheet 2 – Paramétrage harmonisé
        df_full.to_excel(writer, sheet_name="Paramétrage harmonisé", index=False)
        ws2 = writer.sheets["Paramétrage harmonisé"]
        # Header
        for col_idx, _ in enumerate(df_full.columns):
            ws2.write(0, col_idx, df_full.columns[col_idx], fmt_header)
            ws2.set_column(col_idx, col_idx, 28)
        # Row background by Niveau
        if "Niveau" in df_full.columns:
            lvl_col = df_full.columns.get_loc("Niveau")
            for row_idx in range(1, len(df_full) + 1):
                level = str(df_full.iloc[row_idx-1]["Niveau"])
                bg = color_for_level(level)
                fmt_row = wb.add_format({"bg_color": bg})
                ws2.set_row(row_idx, None, fmt_row)
    return output.getvalue()

# ------------------------------------------------------
# UI – Body
# ------------------------------------------------------
st.title("📊 Hopia – Récapitulatif de Paramétrage (Excel)")
st.markdown(
    "Charge un **paramétrage brut** (CSV/XLSX) et récupère un **Excel récapitulatif harmonisé** "
    "avec les niveaux **DURE / MOYENNE / SOUPLE** et un **résumé par rubrique**."
)

uploaded = st.file_uploader("Dépose ton fichier de paramétrage (CSV ou XLSX)", type=["csv", "xlsx"])

example_expander = st.expander("Voir un exemple de structure attendue")
example_expander.dataframe(
    pd.DataFrame({
        "PK": [1536, 1692],
        "Intitulé": ["OS4 - Jeudi MAT CS", "Pas d'affectation sur Jour OFF"],
        "Type": ["Enchaînement de postes", "Demandes d'absence"],
        "Priorités": ["HARD", "HARD"],
        "Équipes": ["ARE", "ARE"]
    })
)

if uploaded:
    try:
        if uploaded.name.lower().endswith(".csv"):
            raw = pd.read_csv(uploaded)
        else:
            raw = pd.read_excel(uploaded)

        df = normalize_cols(raw)

        # Calcul du niveau & trace de règle
        levels = df.apply(map_level, axis=1, result_type="expand")
        df["Niveau"] = levels[0]
        df["Règle de mapping"] = levels[1]

        # Résumé
        df_summary = build_summary(df)

        st.success("Fichier chargé et interprété ✅")
        with st.expander("Aperçu – Paramétrage harmonisé"):
            st.dataframe(df.head(50))
        with st.expander("Aperçu – Résumé par rubrique"):
            st.dataframe(df_summary)

        # Export Excel
        excel_bytes = to_excel_bytes(df_summary, df)
        st.download_button(
            "⬇️ Télécharger l'Excel récapitulatif",
            data=excel_bytes,
            file_name="Parametrage_Harmonise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Infos mapping
        st.markdown("### 🔎 Rappels de mapping")
        st.markdown(
            "- **Hors Remplissage des postes** : `HARD/HARD_LOWER → DURE` ; `STRONG_* / PRIORITY_* / DEFAULT_PENALTY → MOYENNE` ; `SOFT_* → SOUPLE`  \n"
            "- **Remplissage des postes** : `PRIORITY_1 / HARD / STRONG_1 → DURE` ; `PRIORITY_2 / STRONG_2/3 → MOYENNE` ; `PRIORITY_3 / DEFAULT_PENALTY / SOFT_* → SOUPLE`"
        )

    except Exception as e:
        st.error(f"Erreur lors de la lecture ou de la transformation du fichier : {e}")
        st.stop()
else:
    st.info("Charge un fichier pour générer l’Excel.")
