# app.py
# ------------------------------------------------------
# Streamlit – Générateur d'Excel récapitulatif de paramétrage Hopia
# Entrées possibles :
#   - Fichier brut (TXT / CSV / XLSX) de type "export Excel"
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
COLOR_DURE = "#ffcccc"        # rouge clair
COLOR_MOY = "#ffe5b4"         # orange clair
COLOR_SOFT = "#ccffcc"        # vert clair
COLOR_HEADER = "#003366"      # bleu foncé
COLOR_HEADER_TXT = "#FFFFFF"  # texte header blanc


# ------------------------------------------------------
# Parseur pour le format vertical (exemple Urgentistes)
# ------------------------------------------------------
def parse_vertical_blocks(content: str) -> pd.DataFrame:
    """
    Parse un texte copié-collé au format :
    PK
    Intitulé [TAB ou 2+ espaces] Type
    (une ou plusieurs lignes de Priorités)
    Équipes

    Exemple concret : le gros bloc 'Urgentistes' fourni.
    """
    lines = [l.strip() for l in content.splitlines()]
    # On enlève les lignes vides
    lines = [l for l in lines if l]

    # On avance jusqu'à la première ligne qui ressemble à un PK numérique
    i = 0
    while i < len(lines) and not lines[i].isdigit():
        i += 1

    def looks_like_pk(s: str) -> bool:
        return s.isdigit()

    def looks_like_priority(s: str) -> bool:
        # Tout ce qui ressemble à une inégalité ou un token de pénalité
        return bool(
            re.search(
                r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1|<|>|≤|>=|≥)",
                s
            )
        )

    records = []

    while i < len(lines):
        if not looks_like_pk(lines[i]):
            # Si on tombe sur autre chose qu'un PK, on avance
            i += 1
            continue

        pk = lines[i]
        i += 1
        if i >= len(lines):
            break

        # Ligne suivante : Intitulé + Type séparés par tab ou au moins 2 espaces
        line2 = lines[i]
        i += 1

        parts = re.split(r"\t+| {2,}", line2, maxsplit=1)
        if len(parts) == 2:
            intitule = parts[0].strip()
            type_val = parts[1].strip()
        else:
            intitule = parts[0].strip()
            type_val = ""

        # On accumule les lignes de Priorités, puis l'Équipe
        prio_lines: list[str] = []
        equipe = ""

        while i < len(lines) and not looks_like_pk(lines[i]):
            l = lines[i]

            # Si on a déjà des priorités et que la ligne ne ressemble plus à une priorité → équipe
            if not looks_like_priority(l) and prio_lines:
                equipe = l
                i += 1
                break

            # Si aucune priorité lue et la ligne ne ressemble pas à une priorité :
            # on considère que c'est une ligne "Équipes" sans priorités détaillées.
            if not looks_like_priority(l) and not prio_lines:
                equipe = l
                i += 1
                break

            # Sinon, c'est une ligne de priorités
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

    if not records:
        return pd.DataFrame()

    return pd.DataFrame(records)


# ------------------------------------------------------
# Lecture de contenu texte brut (copier-coller)
# ------------------------------------------------------
def read_text_content(content: str) -> pd.DataFrame | None:
    """
    Lit un contenu texte brut :
    1) On tente le format vertical (type Urgentistes)
    2) Sinon on tente des formats CSV/TSV classiques
    """
    content = content.strip()
    if not content:
        return None

    # 1) Tentative : format vertical
    df_vert = parse_vertical_blocks(content)
    if not df_vert.empty:
        return df_vert

    # 2) Tentative : CSV classique (séparateurs ; tab ,)
    possible_seps = [";", "\t", ","]
    for sep in possible_seps:
        try:
            df = pd.read_csv(io.StringIO(content), sep=sep, engine="python")
            if df.shape[1] >= 2:
                return df
        except Exception:
            pass

    # 3) Dernier recours : séparation par espaces
    try:
        df = pd.read_csv(io.StringIO(content), delim_whitespace=True, engine="python")
        return df
    except Exception:
        st.error("Impossible d'interpréter le texte collé. Vérifie le format.")
        return None


# ------------------------------------------------------
# Lecture du fichier uploadé (TXT / CSV / XLSX)
# ------------------------------------------------------
def read_uploaded_file(uploaded_file) -> pd.DataFrame | None:
    """Lit le fichier uploadé, en gérant TXT, CSV, XLSX, avec détection auto du séparateur pour le texte."""
    name = uploaded_file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)

    # CSV classique
    if name.endswith(".csv"):
        # sep=None + engine="python" → détection automatique du séparateur
        return pd.read_csv(uploaded_file, sep=None, engine="python")

    # Fichier texte brut
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
    """
    - Harmonise les noms de colonnes (strip)
    - S'assure que PK, Type, Priorités, Équipes existent
    - Crée 'Intitulé' si absent (copie de Type pour la lisibilité)
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Colonnes obligatoires minimales
    needed = ["PK", "Type", "Priorités", "Équipes"]
    for col in needed:
        if col not in df.columns:
            df[col] = None

    # Intitulé : si absent, on duplique Type pour avoir un libellé lisible
    if "Intitulé" not in df.columns:
        df["Intitulé"] = df["Type"]

    # Ordre des colonnes (les autres suivront)
    front = ["PK", "Intitulé", "Type", "Priorités", "Équipes"]
    front_existing = [c for c in front if c in df.columns]
    others = [c for c in df.columns if c not in front_existing]

    return df[front_existing + others]


# ------------------------------------------------------
# Analyse des tokens de priorité
# ------------------------------------------------------
def token_set(priorites: str) -> set:
    """Extrait les tokens de pénalités/forces (HARD, SOFT_1, PRIORITY_1, etc.) d'une chaîne."""
    if pd.isna(priorites):
        return set()
    txt = str(priorites).upper()

    # On capture les mots-clés même dans des expressions du type
    # "2 < SOFT_2 ≤ 3 3 < PRIORITY_1 ≤ 4"
    parts = re.findall(
        r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_1)",
        txt
    )
    return set(parts)


# ------------------------------------------------------
# Mapping vers DURE / MOYENNE / SOUPLE
# ------------------------------------------------------
def map_level(row) -> Tuple[str, str]:
    """
    Retourne (Niveau, Règle utilisée) en fonction de :
    - Type de contrainte (Remplissage des postes ou non)
    - Priorités (HARD, SOFT_x, PRIORITY_x, STRONG_x, HARD_LOWER, DEFAULT_PENALTY)
    """
    type_txt = str(row.get("Type", "")).strip().lower()
    prio_raw = str(row.get("Priorités", ""))
    toks = token_set(prio_raw)

    is_remplissage = "remplissage des postes" in type_txt

    # Valeur par défaut si on ne comprend rien → SOUPLE
    niveau = "SOUPLE"
    rule = "SOFT_* → SOUPLE (fallback)"

    if is_remplissage:
        # Cas particulier des postes à remplir
        if {"PRIORITY_1", "HARD", "STRONG_1"} & toks:
            return "DURE", "Remplissage : PRIORITY_1/HARD/STRONG_1 → DURE"
        if {"PRIORITY_2", "STRONG_2", "STRONG_3"} & toks:
            return "MOYENNE", "Remplissage : PRIORITY_2/STRONG_2/STRONG_3 → MOYENNE"
        if (
            {"PRIORITY_3", "DEFAULT_PENALTY"} & toks
            or any(t.startswith("SOFT_") for t in toks)
        ):
            return "SOUPLE", "Remplissage : PRIORITY_3/DEFAULT_PENALTY/SOFT_* → SOUPLE"
        # Aucun token connu : on laisse SOUPLE
        return niveau, rule

    # Cas général (hors Remplissage des postes)
    if "HARD" in toks or "HARD_LOWER" in toks:
        return "DURE", "Hors remplissage : HARD/HARD_LOWER → DURE"
    if (
        any(t.startswith("STRONG_") for t in toks)
        or any(t.startswith("PRIORITY_") for t in toks)
        or "DEFAULT_PENALTY" in toks
        or "PRIVATE_ALGO_1" in toks
    ):
        return "MOYENNE", "Hors remplissage : STRONG_*/PRIORITY_*/DEFAULT/PRIVATE → MOYENNE"
    if any(t.startswith("SOFT_") for t in toks):
        return "SOUPLE", "Hors remplissage : SOFT_* → SOUPLE"

    # Aucun mot-clé reconnu : on laisse SOUPLE
    return niveau, rule


def color_for_level(level: str) -> str:
    """Renvoie la couleur hex correspondante au niveau (DURE/MOYENNE/SOUPLE)."""
    l = (level or "").upper().replace("É", "E")
    if "DURE" in l:
        return COLOR_DURE
    if "MOY" in l:  # MOYENNE
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
        fill_value=0
    )

    # On force l’ordre des colonnes
    piv = piv.reindex(columns=["DURE", "MOYENNE", "SOUPLE"], fill_value=0)
    piv["Total"] = piv.sum(axis=1)
    piv = piv.reset_index().rename(columns={"Type": "Rubrique"})

    return piv


# ------------------------------------------------------
# Construction du fichier Excel
# (df_autres / df_remp = seulement Intitulé / Type / Equipe / Niveau)
# ------------------------------------------------------
def to_excel_bytes(df_summary: pd.DataFrame,
                   df_autres: pd.DataFrame,
                   df_remp: pd.DataFrame) -> bytes:
    import xlsxwriter  # utilisé par pandas.ExcelWriter(engine='xlsxwriter')

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        wb = writer.book

        # -------- Feuille 1 – Résumé par rubrique --------
        df_summary.to_excel(writer, sheet_name="Résumé par rubrique", index=False)
        ws = writer.sheets["Résumé par rubrique"]

        fmt_header = wb.add_format(
            {
                "bold": True,
                "bg_color": COLOR_HEADER,
                "font_color": COLOR_HEADER_TXT,
                "align": "center",
            }
        )

        # Largeur + header
        for col_idx, col_name in enumerate(df_summary.columns):
            ws.write(0, col_idx, col_name, fmt_header)
            ws.set_column(col_idx, col_idx, 25)

        # Coloration des colonnes par niveau
        col_map = {"DURE": COLOR_DURE, "MOYENNE": COLOR_MOY, "SOUPLE": COLOR_SOFT}
        for col_name, bg in col_map.items():
            if col_name in df_summary.columns:
                col_idx = df_summary.columns.get_loc(col_name)
                fmt_lvl = wb.add_format({"bg_color": bg})
                ws.conditional_format(
                    1,
                    col_idx,
                    len(df_summary) + 50,
                    col_idx,
                    {"type": "no_errors", "format": fmt_lvl},
                )

        # -------- Feuille 2 – Autres contraintes --------
        df_autres.to_excel(writer, sheet_name="Paramétrage – Autres", index=False)
        ws2 = writer.sheets["Paramétrage – Autres"]

        for col_idx, col_name in enumerate(df_autres.columns):
            ws2.write(0, col_idx, col_name, fmt_header)
            ws2.set_column(col_idx, col_idx, 28)

        if "Niveau" in df_autres.columns:
            for row_idx in range(1, len(df_autres) + 1):
                level = str(df_autres.iloc[row_idx - 1]["Niveau"])
                bg = color_for_level(level)
                fmt_row = wb.add_format({"bg_color": bg})
                ws2.set_row(row_idx, None, fmt_row)

        # -------- Feuille 3 – Remplissage des postes --------
        df_remp.to_excel(writer, sheet_name="Paramétrage – Remplissage", index=False)
        ws3 = writer.sheets["Paramétrage – Remplissage"]

        for col_idx, col_name in enumerate(df_remp.columns):
            ws3.write(0, col_idx, col_name, fmt_header)
            ws3.set_column(col_idx, col_idx, 28)

        if "Niveau" in df_remp.columns:
            for row_idx in range(1, len(df_remp) + 1):
                level = str(df_remp.iloc[row_idx - 1]["Niveau"])
                bg = color_for_level(level)
                fmt_row = wb.add_format({"bg_color": bg})
                ws3.set_row(row_idx, None, fmt_row)

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

with st.expander("🔍 Exemple de structure attendue (mini)"):
    example_df = pd.DataFrame(
        {
            "PK": [1536, 1692, 1767],
            "Type": [
                "OS4 - Jeudi MAT CS",
                "Pas d'affectation sur Jour OFF",
                "Remplissage des postes Bloc",
            ],
            "Priorités": ["HARD", "HARD", "PRIORITY_2"],
            "Équipes": ["ARE", "ARE", "ARE"],
        }
    )
    st.dataframe(example_df, use_container_width=True)

if df_raw is not None:
    try:
        df_norm = normalize_cols(df_raw)

        # Calcul Niveau + Règle
        levels = df_norm.apply(map_level, axis=1, result_type="expand")
        df_norm["Niveau"] = levels[0]
        df_norm["Règle de mapping"] = levels[1]

        # 🔹 Filtrage : on enlève Demandes d'absence / Demandes de travail
        mask = ~df_norm["Type"].astype(str).str.strip().isin(
            ["Demandes d'absence", "Demandes de travail"]
        )
        df_filtered = df_norm[mask].copy()

        # Résumé (sur le df filtré, on garde PK pour le pivot)
        df_summary = build_summary(df_filtered)

        # 🔹 DataFrames pour export : seulement Intitulé / Type / Equipe / Niveau
        df_filtered["Equipe"] = df_filtered["Équipes"]

        is_remplissage = df_filtered["Type"].astype(str).str.strip().str.lower() == "remplissage des postes"

        df_autres = df_filtered[~is_remplissage][["Intitulé", "Type", "Equipe", "Niveau"]].copy()
        df_remp = df_filtered[is_remplissage][["Intitulé", "Type", "Equipe", "Niveau"]].copy()

        st.success("✅ Données chargées, filtrées et interprétées avec succès.")

        with st.expander("Aperçu – Autres contraintes (exportées)"):
            st.dataframe(df_autres.head(50), use_container_width=True)

        with st.expander("Aperçu – Remplissage des postes (exportées)"):
            st.dataframe(df_remp.head(50), use_container_width=True)

        with st.expander("Aperçu – Résumé par rubrique"):
            st.dataframe(df_summary, use_container_width=True)

        # Export Excel
        excel_bytes = to_excel_bytes(df_summary, df_autres, df_remp)
        st.download_button(
            "⬇️ Télécharger l'Excel récapitulatif harmonisé",
            data=excel_bytes,
            file_name="Parametrage_Harmonise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("### 🔎 Rappels de mapping (mémo)")
        st.markdown(
            """
**Hors _Remplissage des postes_ :**

- `HARD` / `HARD_LOWER` → **DURE**
- `STRONG_*` / `PRIORITY_*` / `DEFAULT_PENALTY` / `PRIVATE_ALGO_1` → **MOYENNE**
- `SOFT_*` → **SOUPLE**

**Pour _Remplissage des postes_ :**

- `PRIORITY_1` / `HARD` / `STRONG_1` → **DURE** (poste à remplir en priorité extrême)
- `PRIORITY_2` / `STRONG_2` / `STRONG_3` → **MOYENNE** (poste à remplir normalement)
- `PRIORITY_3` / `DEFAULT_PENALTY` / `SOFT_*` → **SOUPLE** (poste à remplir en dernier)
"""
        )

    except Exception as e:
        st.error(f"Erreur lors du traitement des données : {e}")
else:
    st.info("Importe un fichier **ou** colle le contenu de ton export pour générer l’Excel harmonisé.")
