# app.py
# ------------------------------------------------------
# Streamlit – Générateur d'Excel récapitulatif de paramétrage Hopia
# Entrées possibles :
#   - Paramétrage brut (TXT / CSV / XLSX) OU copier-coller
#   - Permissions des membres (copier-coller Back-Office)
# Export :
#   - Si seulement Permissions -> Excel avec feuille "Permissions"
#   - Si Paramétrage (+ éventuellement Permissions) -> Excel complet
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
st.title("📊 Hopia – Générateur d’Excel récapitulatif harmonisé")

# ------------------------------------------------------
# Couleurs pour l'export Excel
# ------------------------------------------------------
COLOR_DURE = "#ffcccc"
COLOR_MOY = "#ffe5b4"
COLOR_SOFT = "#ccffcc"
COLOR_HEADER = "#003366"
COLOR_HEADER_TXT = "#FFFFFF"

# ------------------------------------------------------
# Ordre hiérarchique des tokens de priorité (Remplissage)
# ------------------------------------------------------
PRIORITY_ORDER = [
    "HARD",
    "HARD_LOWER",
    "PRIORITY_1",
    "PRIORITY_2",
    "PRIORITY_3",
    "DEFAULT_PENALTY",
    "STRONG_1",
    "STRONG_2",
    "STRONG_3",
    "PRIVATE_ALGO_1",
    "PRIVATE_ALGO_2",
    "PRIVATE_ALGO_3",
    "SOFT_1",
    "SOFT_2",
    "SOFT_3",
]

# ------------------------------------------------------
# Mapping permissions (token -> (colonne, valeur))
# - Priorité: Gestionnaire > Modifications > Lecture seule > X
# ------------------------------------------------------
PERMISSIONS_SPECS = [
    ("SwapsWrite", "Echanges et Reprises", "Modifications"),
    ("DemandsRead", "Desiderata", "Lecture seule"),
    ("DemandsWrite", "Desiderata", "Modifications"),
    ("PlanningRead", "Planning Personnel", "Lecture seule"),
    ("PlanningWrite", "Planning Personnel", "Modifications"),
    ("DashboardRead", "Dashboard Equipe", "Lecture seule"),
    ("TeamManageRead", "Gestion d'équipe", "Lecture seule"),
    ("TeamManageWrite", "Gestion d'équipe", "Modifications"),
    ("SwapsManageRead", "Gestion Echanges et Reprises", "Lecture seule"),
    ("SwapsManageWrite", "Echanges et Reprises", "Gestionnaire"),
    ("TaskCommentsRead", "Commentaires", "Lecture seule"),
    ("TeamPlanningRead", "Planning d'équipe", "Lecture seule"),
    ("TeamPlanningWrite", "Planning d'équipe", "Modifications"),
    ("DemandsManageRead", "Gestion des desiderata", "Lecture seule"),
    ("DemandsManageWrite", "Gestion des desiderata", "Modifications"),
    ("PlanningManageRead", "Gestion de Planning", "Lecture seule"),
    ("PlanningManageWrite", "Gestion de Planning", "Modifications"),
]

PERM_TOKEN_TO_COLVAL = {t: (col, val) for (t, col, val) in PERMISSIONS_SPECS}
PERM_COLUMNS = [col for (_, col, _) in PERMISSIONS_SPECS]
_seen = set()
PERM_COLUMNS = [c for c in PERM_COLUMNS if not (c in _seen or _seen.add(c))]

PERM_RANK = {"X": 0, "Lecture seule": 1, "Modifications": 2, "Gestionnaire": 3}


# ------------------------------------------------------
# Parseur pour le format vertical (exemple Urgentistes)
# ------------------------------------------------------
def parse_vertical_blocks(content: str) -> pd.DataFrame:
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
                r"(HARD(?:_LOWER)?|SOFT_\d|STRONG_\d|PRIORITY_\d|DEFAULT_PENALTY|PRIVATE_ALGO_\d|<|>|≤|>=|≥)",
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
# Lecture de contenu texte brut (copier-coller paramétrage)
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
# Normalisation des colonnes (paramétrage)
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
        r"(HARD(?:_LOWER)?"
        r"|SOFT_\d"
        r"|STRONG_\d"
        r"|PRIORITY_\d"
        r"|DEFAULT_PENALTY"
        r"|PRIVATE_ALGO_\d)",
        txt,
    )
    return set(parts)


def main_priority_token(priorites: str) -> str | None:
    toks = token_set(priorites)
    if not toks:
        return None
    for t in PRIORITY_ORDER:
        if t in toks:
            return t
    return None


# ------------------------------------------------------
# Mapping vers DURE / MOYENNE / SOUPLE
# ------------------------------------------------------
def map_level(row) -> Tuple[str, str]:
    type_txt = str(row.get("Type", "")).strip().lower()
    toks = token_set(row.get("Priorités", ""))

    is_remplissage = "remplissage des postes" in type_txt

    niveau = "SOUPLE"
    rule = "Aucune priorité détectée → SOUPLE (fallback)"

    if is_remplissage:
        if {"HARD_LOWER", "HARD"} & toks:
            return "DURE", "Remplissage : HARD/HARD_LOWER → DURE"
        if {
            "PRIORITY_1",
            "PRIORITY_2",
            "PRIORITY_3",
            "DEFAULT_PENALTY",
            "STRONG_1",
            "STRONG_2",
            "STRONG_3",
        } & toks:
            return (
                "MOYENNE",
                "Remplissage : PRIORITY_1/PRIORITY_2/PRIORITY_3/DEFAULT_PENALTY/"
                "STRONG_1/STRONG_2/STRONG_3 → MOYENNE",
            )
        if (
            {"PRIVATE_ALGO_1", "PRIVATE_ALGO_2", "PRIVATE_ALGO_3", "SOFT_1", "SOFT_2", "SOFT_3"}
            & toks
            or any(t.startswith("SOFT_") for t in toks)
        ):
            return "SOUPLE", "Remplissage : PRIVATE_ALGO_*/SOFT_* → SOUPLE"
        return niveau, rule

    if "HARD" in toks or "HARD_LOWER" in toks:
        return "DURE", "Hors remplissage : HARD/HARD_LOWER → DURE"
    if (
        any(t.startswith("STRONG_") for t in toks)
        or any(t.startswith("PRIORITY_") for t in toks)
        or "DEFAULT_PENALTY" in toks
    ):
        return "MOYENNE", "Hors remplissage : STRONG_*/PRIORITY_*/DEFAULT_PENALTY → MOYENNE"
    if any(t.startswith("SOFT_") for t in toks) or any(t.startswith("PRIVATE_ALGO_") for t in toks):
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
# Parseur Permissions (copier-coller Back-Office)
# -> table "longue" : Membre, Email, Permissions_raw (list)
# ------------------------------------------------------
def parse_permissions_text(content: str) -> pd.DataFrame | None:
    content = (content or "").strip()
    if not content:
        return None

    lines = [l.strip() for l in content.splitlines()]
    lines = [l for l in lines if l]

    candidate_lines = [l for l in lines if ("\t" in l) or re.search(r" {2,}", l)]
    if not candidate_lines:
        return None

    header_idx = None
    for idx, l in enumerate(candidate_lines):
        cols = re.split(r"\t+| {2,}", l.strip())
        cols = [c.strip() for c in cols if c.strip()]
        if len(cols) >= 3 and cols[0].lower() == "membre" and cols[1].lower() == "email":
            header_idx = idx
            break

    if header_idx is None:
        return None

    records = []
    for l in candidate_lines[header_idx + 1 :]:
        cols = re.split(r"\t+| {2,}", l.strip(), maxsplit=2)
        cols = [c.strip() for c in cols]
        if len(cols) < 3:
            continue

        membre, email, perms = cols[0], cols[1], cols[2]
        if "@" not in email:
            continue

        perm_list = [p.strip() for p in perms.split(",") if p.strip()]
        records.append({"Membre": membre, "Email": email, "Permissions_raw": perm_list})

    if not records:
        return None

    return pd.DataFrame(records)


def build_permissions_matrix(df_perm_long: pd.DataFrame) -> pd.DataFrame:
    df = df_perm_long.copy()
    for col in PERM_COLUMNS:
        df[col] = "X"

    def apply_token(current: str, new_val: str) -> str:
        return new_val if PERM_RANK.get(new_val, 0) > PERM_RANK.get(current, 0) else current

    for idx, row in df.iterrows():
        tokens = row.get("Permissions_raw") or []
        for tok in tokens:
            tok = tok.strip()
            if tok in PERM_TOKEN_TO_COLVAL:
                col, val = PERM_TOKEN_TO_COLVAL[tok]
                df.at[idx, col] = apply_token(df.at[idx, col], val)

    out = df[["Membre", "Email"] + PERM_COLUMNS].copy()
    out = out.sort_values(by=["Membre", "Email"])
    return out


# ------------------------------------------------------
# Construction du fichier Excel
# -> écrit seulement les feuilles disponibles
# ------------------------------------------------------
def to_excel_bytes(
    df_autres: pd.DataFrame | None = None,
    df_remp: pd.DataFrame | None = None,
    df_summary: pd.DataFrame | None = None,
    df_permissions_matrix: pd.DataFrame | None = None,
) -> bytes:
    import xlsxwriter
    from xlsxwriter.utility import xl_col_to_name

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

        def add_type_borders(ws, df: pd.DataFrame, col_name: str):
            if df is None or df.empty or col_name not in df.columns:
                return
            type_col_idx = df.columns.get_loc(col_name)
            col_letter = xl_col_to_name(type_col_idx)
            border_fmt = wb.add_format({"top": 2})
            nrows, ncols = df.shape
            ws.conditional_format(
                1,
                0,
                nrows,
                ncols - 1,
                {
                    "type": "formula",
                    "criteria": f"=AND(ROW()>2, ${col_letter}2<>${col_letter}1)",
                    "format": border_fmt,
                },
            )

        # --- Paramétrage – Autres (si présent) ---
        if df_autres is not None and not df_autres.empty:
            df_autres.to_excel(writer, sheet_name="Paramétrage – Autres", index=False)
            ws2 = writer.sheets["Paramétrage – Autres"]
            for col_idx, col in enumerate(df_autres.columns):
                ws2.write(0, col_idx, col, fmt_header)
                ws2.set_column(col_idx, col_idx, 28)

            if "Niveau" in df_autres.columns:
                niveau_col = df_autres.columns.get_loc("Niveau")
                for row_idx in range(1, len(df_autres) + 1):
                    val = str(df_autres.iloc[row_idx - 1]["Niveau"])
                    ws2.write(row_idx, niveau_col, val, wb.add_format({"bg_color": color_for_level(val)}))

            add_type_borders(ws2, df_autres, "Type")

        # --- Paramétrage – Remplissage (si présent) ---
        if df_remp is not None and not df_remp.empty:
            df_remp.to_excel(writer, sheet_name="Paramétrage – Remplissage", index=False)
            ws3 = writer.sheets["Paramétrage – Remplissage"]
            for col_idx, col in enumerate(df_remp.columns):
                ws3.write(0, col_idx, col, fmt_header)
                ws3.set_column(col_idx, col_idx, 28)

            add_type_borders(ws3, df_remp, "Type")

        # --- Permissions (si présent) ---
        if df_permissions_matrix is not None and not df_permissions_matrix.empty:
            df_permissions_matrix.to_excel(writer, sheet_name="Permissions", index=False)
            wsp = writer.sheets["Permissions"]
            for col_idx, col in enumerate(df_permissions_matrix.columns):
                wsp.write(0, col_idx, col, fmt_header)
                if col == "Membre":
                    wsp.set_column(col_idx, col_idx, 28)
                elif col == "Email":
                    wsp.set_column(col_idx, col_idx, 30)
                else:
                    wsp.set_column(col_idx, col_idx, 26)

        # --- Résumé (si présent) ---
        if df_summary is not None and not df_summary.empty:
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

            add_type_borders(ws, df_summary, "Rubrique")

    return output.getvalue()


# ------------------------------------------------------
# UI
# ------------------------------------------------------
uploaded = st.file_uploader(
    "📁 Importer un Excel de paramétrage (optionnel)",
    type=["txt", "csv", "xlsx", "xls"],
)

text_pasted = st.text_area(
    "✂️ Collez ici le contenu du Back-Office (Paramétrage) (optionnel) :",
    placeholder="PK\tType\tPriorités\tÉquipes\n549\tPas de MAO...\n...",
    height=200,
)

permissions_pasted = st.text_area(
    "🔐 Collez ici le contenu du Back-Office (Permissions des membres) (optionnel) :",
    placeholder="Permissions des membres\n...\nMembre\tEmail\tPédiatre\nAlice\talice@...\tPlanningRead, PlanningWrite\n...",
    height=200,
)

# --- Permissions ---
df_permissions_matrix = None
if permissions_pasted.strip():
    df_perm_long = parse_permissions_text(permissions_pasted)
    if df_perm_long is None or df_perm_long.empty:
        st.warning("⚠️ Permissions : format non reconnu (vérifie l'entête 'Membre  Email  <Rôle>').")
    else:
        df_permissions_matrix = build_permissions_matrix(df_perm_long)
        st.success(f"✅ Permissions : {len(df_permissions_matrix)} membres détectés.")
        with st.expander("Aperçu – Permissions (matrice)"):
            st.dataframe(df_permissions_matrix, use_container_width=True)

# --- Paramétrage ---
df_autres = None
df_remp = None
df_summary = None

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

        df_filtered = df_norm[
            ~df_norm["Type"].astype(str).str.strip().isin(
                ["Demandes d'absence", "Demandes de travail"]
            )
        ].copy()

        niveau_order = pd.CategoricalDtype(
            categories=["DURE", "MOYENNE", "SOUPLE"],
            ordered=True,
        )
        df_filtered["Niveau"] = (
            df_filtered["Niveau"].astype(str).str.upper().astype(niveau_order)
        )

        df_summary = build_summary(df_filtered)
        df_filtered["Equipe"] = df_filtered["Équipes"]

        type_series = df_filtered["Type"].fillna("").astype(str).str.lower()
        is_rem = type_series.str.contains("remplissage des postes")

        df_autres = df_filtered.loc[
            ~is_rem, ["Intitulé", "Type", "Equipe", "Niveau"]
        ].copy().sort_values(by=["Type", "Niveau", "Intitulé"])

        df_remp_raw = df_filtered.loc[
            is_rem, ["Intitulé", "Type", "Equipe", "Priorités"]
        ].copy()
        df_remp_raw["Token principal"] = df_remp_raw["Priorités"].apply(main_priority_token)

        tokens_present = [
            t
            for t in PRIORITY_ORDER
            if t in df_remp_raw["Token principal"].dropna().unique()
        ]
        priority_rank_map = {t: i + 1 for i, t in enumerate(tokens_present)}
        df_remp_raw["Ordre de priorité"] = df_remp_raw["Token principal"].map(priority_rank_map)

        df_remp = df_remp_raw[["Intitulé", "Type", "Equipe", "Ordre de priorité"]].copy()
        df_remp = df_remp.sort_values(
            by=["Type", "Ordre de priorité", "Intitulé"], na_position="last"
        )

        st.success("✅ Données paramétrage chargées, filtrées et interprétées.")
        with st.expander("Aperçu – Paramétrage – Autres"):
            st.dataframe(df_autres, use_container_width=True)
        with st.expander("Aperçu – Paramétrage – Remplissage"):
            st.dataframe(df_remp, use_container_width=True)
        with st.expander("Aperçu – Résumé"):
            st.dataframe(df_summary, use_container_width=True)

    except Exception as e:
        st.error(f"Erreur lors du traitement des données paramétrage : {e}")

# --- Export : autorisé si au moins permissions OU paramétrage ---
has_anything = (
    (df_permissions_matrix is not None and not df_permissions_matrix.empty)
    or (df_autres is not None and not df_autres.empty)
    or (df_remp is not None and not df_remp.empty)
    or (df_summary is not None and not df_summary.empty)
)

if has_anything:
    excel_bytes = to_excel_bytes(
        df_autres=df_autres,
        df_remp=df_remp,
        df_summary=df_summary,
        df_permissions_matrix=df_permissions_matrix,
    )
    st.download_button(
        "⬇️ Télécharger l'Excel (Permissions et/ou Paramétrage)",
        data=excel_bytes,
        file_name="Export_Hopia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Colle au moins **le Paramétrage** ou **les Permissions** pour générer un Excel.")
