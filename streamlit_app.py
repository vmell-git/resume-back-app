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

st.title("📊 Hopia – Générateur d’Excel récapitulatif de paramétrage")

# ------------------------------------------------------
# Couleurs pour l'export Excel
# ------------------------------------------------------
COLOR_DURE = "#ffcccc"
COLOR_MOY = "#ffe5b4"
COLOR_SOFT = "#ccffcc"
COLOR_HEADER = "#003366"
COLOR_HEADER_TXT = "#FFFFFF"


# ------------------------------------------------------
# Parseur pour le format vertical (exemple Urgentistes)
# ------------------------------------------------------
def parse_vertical_blocks(content: str) -> pd.DataFrame:
    """
    Format attendu :
    PK
    Intitulé [TAB ou 2+ espaces] Type
    (une ou plusieurs lignes de Priorités)
    Équipes
    """
    lines = [l.strip() for l in content.splitlines()]
    lines = [l for l in lines if l]

    i = 0
    while i < len(lines) and not lines[i].isdigit():
        i += 1

    def looks_like_pk(s: str) -> bool:
        return s.isdigit()

    def looks_like_priority(s: str) -> bool:
        return bool(
            re.search(
                r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1|<|>|≤|>=|≥)",
                s,
            )
        )

    records: list[dict] = []

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
        if len(parts) == 2:
            intitule = parts[0].strip()
            type_val = parts[1].strip()
        else:
            intitule = parts[0].strip()
            type_val = ""

        prio_lines: list[str] = []
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

        records.append(
            {
                "PK": pk,
                "Intitulé": intitule,
                "Type": type_val,
                "Priorités": priorite,
                "Équipes": equipe,
            }
        )

    return pd.DataFrame(records)


# ------------------------------------------------------
# Lecture de contenu texte brut (copier-coller)
# ------------------------------------------------------
def read_text_content(content: str) -> pd.DataFrame | None:
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
        except Exception:
            pass

    try:
        return pd.read_csv(io.StringIO(content), delim_whitespace=True, engine="python")
    except Exception:
        st.error("Impossible d'interpréter le texte collé. Vérifie le format.")
        return None


# ------------------------------------------------------
# Lecture du fichier uploadé (TXT / CSV / XLSX)
# ------------------------------------------------------
def read_uploaded_file(uploaded_file) -> pd.DataFrame | None:
    name = uploaded_file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)

    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file, sep=None, engine="python")

    if name.endswith(".txt"):
        raw_bytes = uploaded_file.read()
        try:
            content = raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            content = raw_bytes.decode("latin-1")
        return read_text_content(content)

    st.error("Format de fichier non supporté. Utilise un .txt, .csv ou .xlsx.")
    return None


# ------------------------------------------------------
# Normalisation des colonnes
# ------------------------------------------------------
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    for col in ["PK", "Type", "Priorités", "Équipes"]:
        if col not in df.columns:
            df[col] = None

    if "Intitulé" not in df.columns:
        df["Intitulé"] = df["Type"]

    front = ["PK", "Intitulé", "Type", "Priorités", "Équipes"]
    front_existing = [c for c in front if c in df.columns]
    others = [c for c in df.columns if c not in front_existing]

    return df[front_existing + others]


# ------------------------------------------------------
# Analyse des tokens de priorité
# ------------------------------------------------------
def token_set(priorites: str) -> set:
    if pd.isna(priorites):
        return set()
    txt = str(priorites).upper()
    parts = re.findall(
        r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1)", txt
    )
    return set(parts)


# ------------------------------------------------------
# Mapping vers DURE / MOYENNE / SOUPLE
# ------------------------------------------------------
def map_level(row) -> Tuple[str, str]:
    type_txt = str(row.get("Type", "")).strip().lower()
    toks = token_set(row.get("Priorités", ""))

    is_remplissage = "remplissage des postes" in type_txt

    niveau = "SOUPLE"
    rule = "SOFT_* → SOUPLE (fallback)"

    if is_remplissage:
        if {"HARD_LOWER", "HARD"} & toks:
            return "DURE", "Remplissage : HARD/HARD_LOWER → DURE"
        if {"PRIORITY_1","PRIORITY_2","PRIORITY_3","DEFAULT_PENALTY","STRONG_1","STRONG_2", "STRONG_3"} & toks:
            return "MOYENNE", "Remplissage : PRIORITY_1/PRIORITY_2/PRIORITY_3/DEFAULT_PENALTY/STRONG_1/STRONG_2/STRONG_3 → MOYENNE"
        if (
            {"PRIVATE_ALGO_1","PRIVATE_ALGO_2","PRIVATE_ALGO_3","SOFT_1","SOFT_2","SOFT_3"} & toks
            or any(t.startswith("SOFT_") for t in toks)):
            return "SOUPLE", "Remplissage : PRIVATE_ALGO_*/SOFT_* → SOUPLE"
        return niveau, rule

    if "HARD" in toks or "HARD_LOWER" in toks:
        return "DURE", "Hors remplissage : HARD/HARD_LOWER → DURE"
    if (
        any(t.startswith("STRONG_") for t in toks)
        or any(t.startswith("PRIORITY_") for t in toks)
        or "DEFAULT_PENALTY" in toks
    ):
        return "MOYENNE", "Hors remplissage : STRONG_*/PRIORITY_*/DEFAULT/PRIVATE → MOYENNE"
    if any(t.startswith("SOFT_") for t in toks)
    or any(t.startswith("PRIVATE_ALGO_") for t in toks):
        return "SOUPLE", "Hors remplissage : SOFT_*/PRIVATE_ALGO_* → SOUPLE"

    return niveau, rule


def color_for_level(level: str) -> str:
    l = (level or "").upper()
    if l == "DURE":
        return COLOR_DURE
    if l == "MOYENNE":
        return COLOR_MOY
    return COLOR_SOFT


# ------------------------------------------------------
# Résumé par rubrique (Type)
# ------------------------------------------------------
def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    tmp = df.copy()
    tmp["Niveau"] = tmp["Niveau"].fillna("SOUPLE")

    piv = pd.pivot_table(
        tmp,
        index="Type",
        columns="Niveau",
        values="PK",
        aggfunc="count",
        fill_value=0,
    )

    piv = piv.reindex(columns=["DURE", "MOYENNE", "SOUPLE"], fill_value=0)
    piv["Total"] = piv.sum(axis=1)
    piv = piv.reset_index().rename(columns={"Type": "Rubrique"})

    return piv


# ------------------------------------------------------
# Construction du fichier Excel
# (Résumé en dernière feuille, nommée "Résumé")
# ------------------------------------------------------
def to_excel_bytes(df_autres: pd.DataFrame,
                   df_remp: pd.DataFrame,
                   df_summary: pd.DataFrame) -> bytes:
    import xlsxwriter

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format(
            {
                "bold": True,
                "bg_color": COLOR_HEADER,
                "font_color": COLOR_HEADER_TXT,
                "align": "center",
            }
        )

        # === Feuille 1 – Paramétrage – Autres ===
        df_autres.to_excel(writer, sheet_name="Paramétrage – Autres", index=False)
        ws2 = writer.sheets["Paramétrage – Autres"]

        for col_idx, col in enumerate(df_autres.columns):
            ws2.write(0, col_idx, col, fmt_header)
            ws2.set_column(col_idx, col_idx, 28)

        if not df_autres.empty and "Niveau" in df_autres.columns:
            niveau_col = df_autres.columns.get_loc("Niveau")
            for row_idx in range(1, len(df_autres) + 1):
                val = str(df_autres.iloc[row_idx - 1]["Niveau"])
                ws2.write(
                    row_idx,
                    niveau_col,
                    val,
                    wb.add_format({"bg_color": color_for_level(val)}),
                )

        # === Feuille 2 – Paramétrage – Remplissage ===
        df_remp.to_excel(writer, sheet_name="Paramétrage – Remplissage", index=False)
        ws3 = writer.sheets["Paramétrage – Remplissage"]

        for col_idx, col in enumerate(df_remp.columns):
            ws3.write(0, col_idx, col, fmt_header)
            ws3.set_column(col_idx, col_idx, 28)

        if not df_remp.empty and "Niveau" in df_remp.columns:
            niveau_col = df_remp.columns.get_loc("Niveau")
            for row_idx in range(1, len(df_remp) + 1):
                val = str(df_remp.iloc[row_idx - 1]["Niveau"])
                ws3.write(
                    row_idx,
                    niveau_col,
                    val,
                    wb.add_format({"bg_color": color_for_level(val)}),
                )

        # === Feuille 3 – Résumé (dernière feuille) ===
        df_summary.to_excel(writer, sheet_name="Résumé", index=False)
        ws = writer.sheets["Résumé"]

        for col_idx, col in enumerate(df_summary.columns):
            ws.write(0, col_idx, col, fmt_header)
            ws.set_column(col_idx, col_idx, 25)

        col_map = {"DURE": COLOR_DURE, "MOYENNE": COLOR_MOY, "SOUPLE": COLOR_SOFT}
        for col_name, bg in col_map.items():
            if col_name in df_summary.columns:
                col_idx = df_summary.columns.get_loc(col_name)
                ws.conditional_format(
                    1,
                    col_idx,
                    len(df_summary) + 50,
                    col_idx,
                    {"type": "no_errors", "format": wb.add_format({"bg_color": bg})},
                )

    return output.getvalue()


# ------------------------------------------------------
# Interface – Upload OU copier-coller
# ------------------------------------------------------
uploaded = st.file_uploader(
    "📁 Importer un fichier texte ou Excel de paramétrage",
    type=["txt", "csv", "xlsx", "xls"],
)

text_pasted = st.text_area(
    "✂️ Ou collez directement ici le contenu de votre export :",
    placeholder="PK\tType\tPriorités\tÉquipes\n549\tPas de MAO...\n...",
    height=200,
)

df_raw = None
if uploaded is not None:
    df_raw = read_uploaded_file(uploaded)
elif text_pasted.strip():
    df_raw = read_text_content(text_pasted)

if df_raw is not None:
    try:
        df_norm = normalize_cols(df_raw)

        levels = df_norm.apply(map_level, axis=1, result_type="expand")
        df_norm["Niveau"] = levels[0]
        df_norm["Règle de mapping"] = levels[1]

        # Filtrage : suppressions Demandes d'absence / Demandes de travail
        df_filtered = df_norm[
            ~df_norm["Type"].astype(str).str.strip().isin(
                ["Demandes d'absence", "Demandes de travail"]
            )
        ].copy()

        # Niveau avec ordre DURE > MOYENNE > SOUPLE
        niveau_order = pd.CategoricalDtype(
            categories=["DURE", "MOYENNE", "SOUPLE"],
            ordered=True,
        )
        df_filtered["Niveau"] = (
            df_filtered["Niveau"].astype(str).str.upper().astype(niveau_order)
        )

        # Résumé global
        df_summary = build_summary(df_filtered)

        # Colonne Equipe pour sortie
        df_filtered["Equipe"] = df_filtered["Équipes"]

        # >>> CORRECTION ICI : détection des Remplissages par CONTAINS <<<
        type_series = df_filtered["Type"].fillna("").astype(str).str.lower()
        is_rem = type_series.str.contains("remplissage des postes")

        df_autres = df_filtered.loc[~is_rem, ["Intitulé", "Type", "Equipe", "Niveau"]].copy()
        df_remp = df_filtered.loc[is_rem, ["Intitulé", "Type", "Equipe", "Niveau"]].copy()

        # Tri par Type, Niveau, Intitulé
        df_autres = df_autres.sort_values(by=["Type", "Niveau", "Intitulé"])
        df_remp = df_remp.sort_values(by=["Type", "Niveau", "Intitulé"])

        st.success("✅ Données chargées, filtrées et interprétées avec succès.")

        with st.expander("Aperçu – Paramétrage – Autres"):
            st.dataframe(df_autres, use_container_width=True)

        with st.expander("Aperçu – Paramétrage – Remplissage"):
            st.dataframe(df_remp, use_container_width=True)

        with st.expander("Aperçu – Résumé"):
            st.dataframe(df_summary, use_container_width=True)

        excel_bytes = to_excel_bytes(df_autres, df_remp, df_summary)
        st.download_button(
            "⬇️ Télécharger l'Excel récapitulatif harmonisé",
            data=excel_bytes,
            file_name="Parametrage_Harmonise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Erreur lors du traitement des données : {e}")
else:
    st.info("Importe un fichier **ou** colle le contenu de ton export pour générer l’Excel harmonisé.")
