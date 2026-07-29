# -*- coding: utf-8 -*-
"""
Slangeprogram - Streamlit Version (refactored)

Behavior-preserving refactor of the original app:
- All 4 modes (Quick / Full / Certificate paste / Excel batch) work exactly
  as before, with the same fields and the same order of steps.
- Repeated UI blocks (lager/antall, pos/delnr, pressure-test details,
  AgGrid selection tables) are pulled into small helper functions.
- Shared business-logic bits that were duplicated in the app (adjust_length,
  the Prod.no normalizer, the MONT number list, the coupling-sheet lookup)
  now live in core.py and are reused from there.
- A lighter, cleaner visual theme replaces the old dark-photo-background CSS.
"""

import io
from datetime import datetime

import pandas as pd
import openpyxl
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder

import core

# =====================================================================
# CONFIG
# =====================================================================

st.set_page_config(page_title="Slangeprogram", layout="wide", page_icon="assets/HP_icon.ico")

FIRST_FILE = "Slanger_hylser.xlsx"
SECOND_FILE = "kuplinger_316.xlsx"
CERT_TEMPLATE = "Mal Trykktest Sertikat.xlsx"
SLUTT_TEMPLATE = "Mal sluttkontroll slanger.xlsx"
FLER_SLANGE_MAL = "MAL_slangebeskrivelse_flere_rader.xlsx"
SERTIFIKAT_MAL = "MAL_Lim_inn_rader_for_Sertifikat.xlsx"

MODE_LABELS = {
    "quick": "⌨️ Skriv inn Slangebeskrivelse",
    "full": "🖱 Velg Slange og Kuplinger",
    "certificate": "📋 Lim inn rader for Sertifikat",
    "excel_batch": "📂 Excel – flere slanger",
}
LABEL_TO_MODE = {v: k for k, v in MODE_LABELS.items()}


# =====================================================================
# THEME
# =====================================================================

def inject_theme():
    """Light, modern theme. Replaces the old full-page photo + dark overlay
    with a clean neutral background, a single accent color, and softer,
    rounded controls. No layout/flow changes — purely visual."""
    st.markdown(
        """
        <style>
        :root {
            --hp-accent: #0F6E5B;
            --hp-accent-dark: #0B5445;
            --hp-accent-light: #E7F3EF;
            --hp-bg: #F4F7F9;
            --hp-card: #FFFFFF;
            --hp-border: #E2E8F0;
            --hp-text: #1F2937;
        }

        .stApp {
            background: linear-gradient(180deg, #F6F9FA 0%, #EEF2F5 100%);
        }

        h1, h2, h3, h4 {
            color: var(--hp-text) !important;
            font-weight: 700 !important;
        }

        hr {
            border-top: 1px solid var(--hp-border) !important;
            margin: 1.1rem 0 !important;
        }

        /* Card look for bordered containers (used to group each section) */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--hp-card);
            border-radius: 14px !important;
            border: 1px solid var(--hp-border) !important;
            box-shadow: 0 1px 3px rgba(15, 23, 42, 0.05);
            padding: 0.25rem 0.25rem;
        }

        /* Buttons */
        .stButton > button, .stDownloadButton > button {
            background-color: var(--hp-accent) !important;
            color: white !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 0.6rem 1.2rem !important;
            font-weight: 600 !important;
            transition: background-color 0.15s ease, transform 0.05s ease;
        }
        .stButton > button:hover, .stDownloadButton > button:hover {
            background-color: var(--hp-accent-dark) !important;
            color: white !important;
        }
        .stButton > button:active, .stDownloadButton > button:active {
            transform: translateY(1px);
        }

        /* Inputs */
        .stTextInput input, .stNumberInput input, .stTextArea textarea {
            border-radius: 8px !important;
            border: 1px solid var(--hp-border) !important;
        }

        /* Alerts */
        .stAlert {
            border-radius: 10px !important;
        }

        /* Radio (mode selector) rendered as pill-like segments */
        div[role="radiogroup"] label {
            background: var(--hp-card);
            border: 1px solid var(--hp-border);
            border-radius: 999px;
            padding: 0.3rem 0.9rem;
            margin-right: 0.4rem;
        }

        /* AG-Grid */
        .ag-root-wrapper {
            border-radius: 10px !important;
            overflow: hidden;
            border: 1px solid var(--hp-border) !important;
        }
        .ag-header, .ag-header-row {
            background-color: var(--hp-accent-light) !important;
        }
        .ag-theme-streamlit .center-header .ag-header-cell-label {
            justify-content: center !important;
        }
        .ag-theme-streamlit .center-header .ag-header-cell-text {
            text-align: center !important;
            width: 100%;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =====================================================================
# DATA LOADING
# =====================================================================

@st.cache_data
def load_all():
    try:
        df1, df2_all = core.load_main_data(FIRST_FILE, SECOND_FILE)
        mont_df, trykktest_df, prikling_df = core.load_support_sheets(FIRST_FILE)
        return df1, df2_all, mont_df, trykktest_df, prikling_df
    except Exception as e:
        st.error(f"Feil ved lasting av data: {e}")
        st.info("Sørg for at Excel-filene er i samme mappe som appen")
        st.stop()


def make_cert_row_lookup(abs_sert_df):
    """Return a function that looks up a row in the ABS Sert. sheet by Prod.no."""
    def get_cert_row(prod_no):
        col_a = abs_sert_df.columns[0]  # Prod.no (column A)
        matches = abs_sert_df[abs_sert_df[col_a].astype(str).str.strip() == str(prod_no)]
        return matches.iloc[0] if not matches.empty else None
    return get_cert_row


# =====================================================================
# SESSION STATE
# =====================================================================

def init_session_state():
    defaults = {
        "abs_selected_any": False,
        "output_rows": [],
        "certificate_data_list": [],
        "pos_counter": 1,
        "input_mode": "quick",
        "selected_hose_row": None,
        "selected_c1_row": None,
        "selected_c2_row": None,
        "full_df2": None,
        "output_batches": [],
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


# =====================================================================
# SHARED UI HELPERS
# =====================================================================

def render_lager_select(key):
    return st.selectbox(
        "Lager",
        options=list(core.LAGER_OPTIONS.keys()),
        format_func=lambda x: core.LAGER_OPTIONS[x],
        key=key,
    )


def render_common_settings(prefix):
    """Lager, antall, POS.nr and kundes delnummer — identical block used by
    both Quick and Full mode."""
    settings = {}

    c1, c2 = st.columns(2)
    with c1:
        settings["lager"] = render_lager_select(f"{prefix}_lager")
    with c2:
        settings["antall_slanger"] = st.number_input(
            "Antall slanger", min_value=1, value=1, key=f"{prefix}_antall"
        )

    c1, c2 = st.columns(2)
    with c1:
        settings["pos_mark"] = st.checkbox("Merke med POS.nr?", key=f"{prefix}_pos_mark")
        settings["posnr"] = (
            st.text_input(
                "POS.nr", value=str(st.session_state.pos_counter), key=f"{prefix}_posnr"
            )
            if settings["pos_mark"]
            else ""
        )
    with c2:
        settings["input_linje"] = st.checkbox(
            "Merke med kundes delnummer?", key=f"{prefix}_input_linje"
        )
        settings["inputlinje"] = (
            st.text_input("Kundes delnummer:", key=f"{prefix}_inputlinje")
            if settings["input_linje"]
            else ""
        )

    return settings


def render_pressure_test_toggle(prefix, force_on):
    if force_on:
        st.checkbox(
            "🚰 Skal slangen trykktestes?",
            value=True,
            disabled=True,
            key=f"{prefix}_pressure_test",
        )
        return True
    return st.checkbox("🚰 Skal slangen trykktestes?", key=f"{prefix}_pressure_test")


def render_pressure_details(prefix, antall_slanger, input_linje, inputlinje, angle=""):
    st.subheader("📋 Trykktest Detaljer")
    details = {"antall_slanger": antall_slanger, "angle": angle}

    c1, c2 = st.columns(2)
    with c1:
        details["kunde"] = st.text_input("Kunde", key=f"{prefix}_kunde")
        details["kundens_best_nr"] = st.text_input("Kundens best. Nr.", key=f"{prefix}_best_nr")
    with c2:
        details["hydra_ordre_nr"] = st.text_input(
            "Hydra Pipe ordre nr.", key=f"{prefix}_hydra_ordre"
        )
        if input_linje and inputlinje:
            details["kundes_del_nr"] = inputlinje
        else:
            details["kundes_del_nr"] = st.text_input("Kundes del nr.", key=f"{prefix}_del_nr")

    return details


def render_type_approval_info(type_approval, type_approval1):
    if not (type_approval or type_approval1):
        return
    c1, c2 = st.columns(2)
    with c1:
        if type_approval:
            st.markdown(
                "Krav til DNV Type Approval:  \nStål:  \nVed bruk av Gates slanger, må Gates "
                "kuplinger brukes(M-kuplinger, eller GS/GSM-kuplinger)  \nVed bruk av Vitillo "
                "slange, må Vitillo kuplinger brukes  \nSyrefast:  \nHP kuplinger brukes på både "
                "Gates of Vitillo slanger  \n  \nEr du usikker på hvike slanger som har DNV Type "
                "Approval, gå til Velg Slange og Kuplinger"
            )
    with c2:
        if type_approval1:
            st.markdown(
                "Krav til ABS Type Approval:  \nStål:  \nVed bruk av Gates slanger, må Gates "
                "kuplinger brukes(M-kuplinger, eller GS/GSM-kuplinger)  \nSyrefast:  \nHP kuplinger "
                "brukes på både Gates of Vitillo slanger  \n  \nEr du usikker på hvike slanger som "
                "har ABS Type Approval, gå til Velg Slange og Kuplinger"
            )


def render_selection_table(df, visible_cols, key, hidden_cols=None, header_map=None):
    """Render a single-selection AgGrid table. Returns the selected row as a
    dict, or None if nothing is selected."""
    hidden_cols = hidden_cols or []
    header_map = header_map or {}

    display_df = df[visible_cols + hidden_cols] if hidden_cols else df[visible_cols]

    gb = GridOptionsBuilder.from_dataframe(display_df)
    for col in hidden_cols:
        gb.configure_column(col, hide=True)
    for col, label in header_map.items():
        gb.configure_column(col, headerName=label)

    gb.configure_default_column(
        headerClass="center-header",
        cellStyle={
            "display": "flex",
            "justifyContent": "center",
            "alignItems": "center",
            "textAlign": "center",
        },
    )
    gb.configure_selection(selection_mode="single", use_checkbox=False)

    custom_css = {
        ".ag-header-cell-label": {"justify-content": "center"},
        ".ag-header-cell-text": {"text-align": "center", "width": "100%"},
    }

    grid_response = AgGrid(
        display_df.copy(),
        gridOptions=gb.build(),
        custom_css=custom_css,
        update_on=["selectionChanged"],
        fit_columns_on_grid_load=True,
        theme="streamlit",
        key=key,
    )

    selected = grid_response["selected_rows"]
    if selected is not None and not selected.empty:
        return selected.iloc[0].to_dict()
    return None


# =====================================================================
# ORDER-BUILDING ENGINE
# =====================================================================

def process_and_add_hose(
    selected_row, second_row1, second_row2, sheet_name_found, size_str,
    length_int, material, lager, pos_mark, posnr, input_linje, inputlinje,
    pressure_test, pressure_details, antall_slanger, mont_df, trykktest_df,
    prikling_df, get_cert_row, prikling=False, first_line="", angle="", dnv=False,
):
    """Build the Visma output rows for one hose assembly and register it in
    session state. Used by Quick mode and Full mode alike."""
    rows = []
    start_len = len(st.session_state.output_rows)

    if pos_mark and posnr:
        rows.append(["1", f"POS: {posnr}", int(lager), 1])
        try:
            st.session_state.pos_counter = int(posnr) + 1
        except Exception:
            pass

    if input_linje and inputlinje:
        rows.append(["1", f"{inputlinje}", int(lager), 1])

    if first_line:
        # Quick mode - the summary line is used as-is
        rows.append(["1", first_line, int(lager), 1])
    else:
        # Full mode - build the summary line from the selected components
        part1 = str(selected_row["Beskrivelse"])[:7] if selected_row is not None else ""
        part2 = str(length_int if length_int else "")
        part3 = (
            core.adjust_length(str(second_row1["Beskrivelse"]), material)
            if second_row1 is not None else ""
        )
        part4 = (
            core.adjust_length(str(second_row2["Beskrivelse"]), material)
            if second_row2 is not None else ""
        )
        if angle and angle.strip():
            first_line_display = f"{part1}/{part2}/{part3}/{part4}/{angle}°"
        else:
            first_line_display = f"{part1}/{part2}/{part3}/{part4}"
        rows.append(["1", first_line_display, int(lager), 1])

    if selected_row is not None:
        try:
            qty = round((length_int or 1000) / 1000, 3)
            rows.append([selected_row["Prod.no"], selected_row["Beskrivelse"], int(lager), qty])
        except Exception:
            rows.append(
                [selected_row.get("Prod.no", ""), selected_row.get("Beskrivelse", ""), int(lager), 1]
            )
    else:
        rows.append(["", "Fant ikke første produkt", int(lager), 1])

    # Kupling 2 missing -> treat it as the same as Kupling 1.
    if second_row1 is not None and second_row2 is None:
        second_row2 = second_row1

    same_coupling = (
        second_row1 is not None
        and second_row2 is not None
        and str(second_row1.get("Prod.no", "")).strip() == str(second_row2.get("Prod.no", "")).strip()
    )

    if same_coupling:
        # Kupling 1 and Kupling 2 are the same product (whether because
        # Kupling 2 was left empty, or the same coupling was picked/typed
        # for both ends) -> one line, Antall doubled, instead of two lines.
        rows.append([second_row1["Prod.no"], second_row1["Beskrivelse"], int(lager), 2])
    else:
        if second_row1 is not None:
            rows.append([second_row1["Prod.no"], second_row1["Beskrivelse"], int(lager), 1])
        else:
            rows.append(["", "Fant ikke første kupling", int(lager), 1])

        if second_row2 is not None:
            rows.append([second_row2["Prod.no"], second_row2["Beskrivelse"], int(lager), 1])
        else:
            rows.append(["", "Fant ikke andre kupling", int(lager), 1])

    gsm_count = 0
    if second_row1 is not None and str(second_row1.get("Beskrivelse", "")).startswith("GSM"):
        gsm_count += 1
    if second_row2 is not None and str(second_row2.get("Beskrivelse", "")).startswith("GSM"):
        gsm_count += 1

    if material.lower() == "stål" and selected_row is not None:
        mat_prod = selected_row.get("Stål hylse(Posd.no)", "")
        mat_desc = selected_row.get("Stål hylse(beskrivelse)", "")
    elif selected_row is not None:
        mat_prod = selected_row.get("316 hylse(Posd.no)", "")
        mat_desc = selected_row.get("316 hylse(beskrivelse)", "")
    else:
        mat_prod = ""
        mat_desc = ""

    sheet_key = (
        core._extract_sheet_key_from_sheetname(sheet_name_found)
        if sheet_name_found else ("(st)" if material == "stål" else "(316)")
    )
    skip_staal_hylse = "(M-st)" in sheet_key or "(GSM)" in sheet_key

    if gsm_count < 2 and not skip_staal_hylse and mat_prod:
        stahl_value = 2 if gsm_count == 0 else 1
        rows.append([mat_prod, mat_desc, int(lager), stahl_value])

    mont_row = core.get_mont_row(size_str, sheet_key, mont_df)
    if mont_row is not None:
        rows.append([mont_row["Prod.no"], mont_row["Beskrivelse"], int(lager), 1])

    if prikling and size_str:
        prikling_row = core.get_prikling_row(size_str, prikling_df)
        if prikling_row is not None:
            rows.append([prikling_row["Prod.no"], prikling_row["Beskrivelse"], int(lager), 1])

    if dnv:
        dnv_cert_row = get_cert_row("90003")
        if dnv_cert_row is not None:
            rows.append(
                [dnv_cert_row.get("Prod.no", ""), dnv_cert_row.get("Beskrivelse", ""), int(lager), 1]
            )

    if pressure_test:
        trykktest_row = core.get_trykktest_prodno(size_str, length_int or 1000, trykktest_df)
        if trykktest_row is not None:
            rows.append(
                [trykktest_row["Prod.no"], trykktest_row["Beskrivelse"], int(lager), 1]
            )
        else:
            rows.append(["", "Trykktest: Ja", int(lager), 1])

    rows.append(["1", "", int(lager), ""])

    if antall_slanger and antall_slanger != 1:
        for r in rows:
            core._multiply_row_quantity(r, antall_slanger)

    st.session_state.output_rows.extend(rows)

    if pressure_test:
        st.session_state.certificate_data_list.append({
            "selected_row": selected_row,
            "second_rows": [second_row1, second_row2],
            "size_str": size_str,
            "length_int": length_int,
            "material": material,
            "pressure_details": pressure_details,
        })

    end_len = len(st.session_state.output_rows)
    st.session_state.output_batches.append(end_len - start_len)


def generate_excel():
    rows_for_excel = [r.copy() for r in st.session_state.output_rows]

    # Add ABS cert row (only once, always at the bottom)
    if st.session_state.abs_selected_any:
        lager_value = rows_for_excel[-1][2] if rows_for_excel else 3
        abs_row = st.session_state.get_cert_row("90478")
        if abs_row is not None:
            rows_for_excel.append(["1", "", lager_value, ""])
            rows_for_excel.append(
                [abs_row.get("Prod.no", ""), abs_row.get("Beskrivelse", ""), lager_value, 1]
            )

    output_wb = core.create_output_workbook(
        [[r[0], r[1], r[2], r[3]] for r in rows_for_excel]
    )

    if st.session_state.certificate_data_list:
        for idx, cert_info in enumerate(st.session_state.certificate_data_list, 1):
            try:
                cert_data = core.fill_pressure_test_certificate_data(
                    cert_info["pressure_details"],
                    cert_info["selected_row"],
                    cert_info["second_rows"],
                    cert_info["size_str"],
                    cert_info["length_int"],
                    cert_info["material"],
                )
                if cert_data:
                    sheet_name = (
                        f"Sertifikat {idx}"
                        if len(st.session_state.certificate_data_list) > 1
                        else "Trykktest Sertifikat"
                    )
                    output_wb = core.add_certificate_sheet(
                        output_wb, CERT_TEMPLATE, cert_data, sheet_name
                    )
            except Exception as e:
                st.warning(f"Kunne ikke legge til sertifikat {idx}: {e}")

    try:
        kunde = ""
        hydra_ordre_nr = ""
        if st.session_state.certificate_data_list:
            kunde = st.session_state.certificate_data_list[0]["pressure_details"].get("kunde", "")
            hydra_ordre_nr = st.session_state.certificate_data_list[0]["pressure_details"].get(
                "hydra_ordre_nr", ""
            )
        output_wb = core.add_sluttkontroll_sheet(
            output_wb, SLUTT_TEMPLATE, kunde=kunde, hydra_ordre_nr=hydra_ordre_nr
        )
    except Exception as e:
        st.warning(f"Kunne ikke legge til sluttkontroll: {e}")

    output_buffer = io.BytesIO()
    output_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# =====================================================================
# QUICK MODE
# =====================================================================

def render_quick_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row):
    st.header("➕ Skriv in Slangebeskrivelse")

    c1, c2 = st.columns(2)
    with c1:
        type_approval = st.checkbox("Type Approval (DNV)?", key="quick_type_approval")
    with c2:
        type_approval1 = st.checkbox("Type Approval (ABS)?", key="quick_type_approval1")

    render_type_approval_info(type_approval, type_approval1)

    c1, c2 = st.columns([2, 1])
    with c1:
        first_line = st.text_input(
            "Slangebeskrivelse",
            placeholder="Slange/Lengde/Kupling 1/Kupling 2",
            key="quick_first_line",
        )
    with c2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="quick_material")

    settings = render_common_settings("quick")

    st.divider()
    prikling = st.checkbox("🪛 Skal slangen prikles?", key="quick_prikling")
    pressure_test = render_pressure_test_toggle("quick", type_approval or type_approval1)

    pressure_details = {
        "kunde": "",
        "kundens_best_nr": "",
        "hydra_ordre_nr": "",
        "kundes_del_nr": "",
        "antall_slanger": settings["antall_slanger"],
    }
    if pressure_test:
        pressure_details = render_pressure_details(
            "quick", settings["antall_slanger"], settings["input_linje"], settings["inputlinje"]
        )

    if st.button("✅ Legg til slange", use_container_width=True, key="quick_add_btn"):
        if not first_line:
            st.error("Første utdata-linje må oppgis!")
        else:
            try:
                result = core.find_matches_from_summary(
                    first_line, df1, df2_all, material_pref=material
                )
                if result and result[0] is not None:
                    (
                        selected_row, second_row1, second_row2,
                        sheet_name_found, size_str, length_int,
                    ) = result

                    if settings["input_linje"] and settings["inputlinje"]:
                        pressure_details["kundes_del_nr"] = settings["inputlinje"]

                    if type_approval1:
                        st.session_state.type_approval1 = True

                    process_and_add_hose(
                        selected_row, second_row1, second_row2, sheet_name_found, size_str,
                        length_int, material, settings["lager"], settings["pos_mark"],
                        settings["posnr"], settings["input_linje"], settings["inputlinje"],
                        pressure_test, pressure_details, settings["antall_slanger"],
                        mont_df, trykktest_df, prikling_df, get_cert_row,
                        prikling=prikling, first_line=first_line, dnv=type_approval,
                    )

                    if type_approval1:
                        st.session_state.abs_selected_any = True

                    st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")
                else:
                    st.error(
                        "❌ Kunne ikke tolke slangebeskrivelsen. "
                        "Sjekk at formatet er riktig (Slange-Lengde-Kupling-Kupling)."
                    )
            except Exception as e:
                st.error(f"⚠️ En feil oppstod under tolking: {e}")


# =====================================================================
# FULL MODE
# =====================================================================

def render_full_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row):
    st.header("📝 Velg Slange og Kuplinger")
    st.subheader("1️⃣ Velg slange")

    c1, c2 = st.columns([2, 1])
    with c2:
        type_approval1 = st.checkbox("Type Approval (ABS)?", key="full_type_approval1")
    with c1:
        type_approval = st.checkbox("Type Approval (DNV)?", key="full_type_approval")

    search = st.text_input("Søk etter slange", key="full_search")

    filtered_df = df1.copy()
    dnv_col, abs_col = "Type Approval", "Type Approval1"

    if type_approval and type_approval1:
        filtered_df = filtered_df[
            filtered_df[dnv_col].fillna("").astype(str).str.strip().ne("")
            & filtered_df[abs_col].fillna("").astype(str).str.strip().ne("")
        ]
    elif type_approval:
        filtered_df = filtered_df[filtered_df[dnv_col].fillna("").astype(str).str.strip().ne("")]
    elif type_approval1:
        filtered_df = filtered_df[filtered_df[abs_col].fillna("").astype(str).str.strip().ne("")]

    if search:
        st.session_state.selected_hose_row = None
        filtered_df = filtered_df[
            filtered_df["Beskrivelse_2"].astype(str).str.contains(search, case=False, na=False)
        ]

    st.write("**Velg slange fra tabellen under:**")

    hose_visible_cols = ["Prod.no", "Beskrivelse_2", "Dimensjon", "Trykk(bar)"]
    hose_hidden_cols = [
        "Beskrivelse", "Stål hylse(Posd.no)", "Stål hylse(beskrivelse)",
        "316 hylse(Posd.no)", "316 hylse(beskrivelse)",
    ]
    hose_header_map = {
        "Prod.no": "Artikkel nummer",
        "Beskrivelse_2": "Beskrivelse",
        "Trykk(bar)": "Arbeidstrykk (Bar)",
    }

    selected = render_selection_table(
        filtered_df, hose_visible_cols, key="hose_grid",
        hidden_cols=hose_hidden_cols, header_map=hose_header_map,
    )
    if selected is not None:
        st.session_state.selected_hose_row = selected

    if st.session_state.selected_hose_row is not None:
        selected_row = st.session_state.selected_hose_row
        st.success(f"✅ Valgt: {selected_row['Beskrivelse_2']}")
    else:
        selected_row = None
        st.warning("⚠️ Du må velge slange fra tabellen.")

    c1, c2, c3 = st.columns(3)
    with c1:
        length = st.number_input("Lengde (mm)", value=1000, key="full_length")
    with c2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="full_material")
    with c3:
        st.write("")

    if selected_row is None:
        return

    size = str(selected_row["Dimensjon"]).zfill(2)
    sheet_name = core.determine_coupling_sheet_name(
        selected_row, material, type_approval, FIRST_FILE
    )

    if sheet_name not in df2_all:
        st.error(f"Fant ikke ark: {sheet_name}")
        return

    df2 = df2_all[sheet_name]
    st.session_state.full_df2 = df2

    st.divider()
    st.subheader("2️⃣ Velg kuplinger")

    c1, c2 = st.columns(2)
    with c1:
        st.write("**Kupling 1**")
        st.write("Velg kupling fra tabellen:")
        sel1 = render_selection_table(df2, ["Prod.no", "Beskrivelse"], key="coupling1_grid")
        if sel1 is not None:
            st.session_state.selected_c1_row = sel1
        if st.session_state.selected_c1_row is not None:
            st.write(f"✅ Valgt: *{st.session_state.selected_c1_row['Beskrivelse']}*")
        else:
            st.info("Velg kupling fra tabellen")

    with c2:
        st.write("**Kupling 2**")
        st.write("Velg kupling fra tabellen:")
        sel2 = render_selection_table(df2, ["Prod.no", "Beskrivelse"], key="coupling2_grid")
        if sel2 is not None:
            st.session_state.selected_c2_row = sel2
        if st.session_state.selected_c2_row is not None:
            st.write(f"✅ Valgt: *{st.session_state.selected_c2_row['Beskrivelse']}*")
        else:
            st.info("Velg kupling fra tabellen")

    if st.session_state.selected_c1_row is None or st.session_state.selected_c2_row is None:
        st.warning("⚠️ Du må velge kuplinger i begge ender")
        return

    row_c1 = st.session_state.selected_c1_row
    row_c2 = st.session_state.selected_c2_row

    st.divider()
    st.subheader("3️⃣ Innstillinger")
    settings = render_common_settings("full")

    has_angle_c1 = "45" in str(row_c1["Beskrivelse"]) or "90" in str(row_c1["Beskrivelse"])
    has_angle_c2 = "45" in str(row_c2["Beskrivelse"]) or "90" in str(row_c2["Beskrivelse"])
    angle = ""
    if has_angle_c1 and has_angle_c2:
        st.divider()
        st.subheader("📐 Vinkel")
        angle = st.text_input("Skriv inn vinkel", key="full_angle")

    st.divider()
    prikling = st.checkbox("🪛 Skal slangen prikles?", key="full_prikling")
    pressure_test = render_pressure_test_toggle("full", type_approval or type_approval1)

    pressure_details = {
        "kunde": "",
        "kundens_best_nr": "",
        "hydra_ordre_nr": "",
        "kundes_del_nr": "",
        "antall_slanger": settings["antall_slanger"],
        "angle": angle,
    }
    if pressure_test:
        pressure_details = render_pressure_details(
            "full", settings["antall_slanger"], settings["input_linje"],
            settings["inputlinje"], angle=angle,
        )

    if st.button("✅ Legg til slange", use_container_width=True, key="full_add_btn"):
        pressure_details["angle"] = angle
        if settings["input_linje"] and settings["inputlinje"]:
            pressure_details["kundes_del_nr"] = settings["inputlinje"]

        process_and_add_hose(
            selected_row, row_c1, row_c2, sheet_name, size, length, material,
            settings["lager"], settings["pos_mark"], settings["posnr"],
            settings["input_linje"], settings["inputlinje"], pressure_test,
            pressure_details, settings["antall_slanger"], mont_df, trykktest_df,
            prikling_df, get_cert_row, prikling=prikling, first_line="",
            angle=angle, dnv=type_approval,
        )

        st.session_state.selected_hose_row = None
        st.session_state.selected_c1_row = None
        st.session_state.selected_c2_row = None

        if type_approval1:
            st.session_state.abs_selected_any = True

        st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")


# =====================================================================
# CERTIFICATE PASTE MODE
# =====================================================================

def render_certificate_mode(df1, df2_all, get_cert_row):
    st.header("📋 Lim inn rader for Sertifikat")

    with open(SERTIFIKAT_MAL, "rb") as file:
        st.download_button(
            label="Last ned MAL",
            data=file,
            file_name="MAL_Lim_inn_rader_for_Sertifikat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    uploaded_cert_file = st.file_uploader(
        "Last opp utfylt MAL_Lim_inn_rader_for_Sertifikat.xlsx",
        type=["xlsx"],
        key="cert_file_uploader",
    )

    if uploaded_cert_file is not None:
        try:
            st.session_state.cert_df = pd.read_excel(uploaded_cert_file)
        except Exception as e:
            st.error(f"Kunne ikke lese Excel: {e}")
            return

    if "cert_df" not in st.session_state:
        return

    df_editor = st.session_state.cert_df

    st.subheader("Importerte rader")
    st.dataframe(df_editor, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("📋 Trykktest Detaljer")

    c1, c2 = st.columns(2)
    with c1:
        kunde = st.text_input("Kunde")
        kundens_best_nr = st.text_input("Kundens best. Nr.")
    with c2:
        hydra_ordre_nr = st.text_input("Hydra Pipe ordre nr.")
        material = st.selectbox("Materiale", ["stål", "syrefast"])

    if not st.button("📄 Generer Sertifikater", use_container_width=True):
        return

    df_clean = df_editor.dropna(subset=["Prod.no"])
    df_clean = df_clean[df_clean["Prod.no"].astype(str).str.strip() != ""]

    if df_clean.empty:
        st.warning("Tabellen er tom.")
        return

    # 1. Group rows into hose/component blocks
    assemblies = []
    current_hose_row = None
    current_components = []

    for _, row in df_clean.iterrows():
        p_no = core.normalize_prod_no(row["Prod.no"])

        if p_no == "1":
            if current_hose_row is not None:
                assemblies.append({"hose": current_hose_row, "components": current_components})
            current_hose_row = None
            current_components = []
            continue

        if current_hose_row is None:
            current_hose_row = row
            current_components = []
        else:
            current_components.append(row)

    if current_hose_row is not None:
        assemblies.append({"hose": current_hose_row, "components": current_components})

    # 2. Generate one certificate sheet per assembly
    output_wb = openpyxl.Workbook()
    success_count = 0

    for idx, asm in enumerate(assemblies):
        h_pno = core.normalize_prod_no(asm["hose"]["Prod.no"])
        h_match = df1[df1["Prod.no"].astype(str).str.strip() == h_pno]

        if h_match.empty:
            continue

        # Number of physical hoses, taken from the MONT row's Antall
        real_antall = 1
        for comp in asm["components"]:
            if core.normalize_prod_no(comp["Prod.no"]) in core.MONT_NUMBERS:
                try:
                    val_str = str(comp["Antall"]).replace(",", ".")
                    real_antall = int(float(val_str))
                    break
                except Exception:
                    real_antall = 1

        # Length per hose = (total quantity / number of hoses) * 1000
        try:
            hose_qty_str = str(asm["hose"]["Antall"]).replace(",", ".")
            total_qty = float(hose_qty_str)
            length_mm = int((total_qty / real_antall) * 1000)
        except Exception:
            length_mm = 1000

        # Find the (up to 2) coupling technical rows for the certificate
        c_tech_data = []
        for comp in asm["components"]:
            c_pno = core.normalize_prod_no(comp["Prod.no"])
            if c_pno in core.MONT_NUMBERS or c_pno.startswith("900"):
                continue
            for sheet in df2_all.values():
                m = sheet[sheet["Prod.no"].astype(str).str.strip() == c_pno]
                if not m.empty:
                    c_tech_data.append(m.iloc[0].to_dict())
                    break
            if len(c_tech_data) >= 2:
                break

        if len(c_tech_data) == 1:
            c_tech_data.append(None)

        cert_data = core.fill_pressure_test_certificate_data(
            {
                "kunde": kunde,
                "kundens_best_nr": kundens_best_nr,
                "hydra_ordre_nr": hydra_ordre_nr,
                "antall_slanger": real_antall,
            },
            h_match.iloc[0].to_dict(),
            c_tech_data,
            str(h_match.iloc[0].get("Dimensjon", "00")).zfill(2),
            length_mm,
            material,
        )

        sheet_name = f"Cert_{idx + 1}_{h_pno}"[:31]
        output_wb = core.add_certificate_sheet(output_wb, CERT_TEMPLATE, cert_data, sheet_name)
        success_count += 1

    if success_count > 0:
        if "Sheet" in output_wb.sheetnames:
            del output_wb["Sheet"]
        output_wb.active = 0
        buf = io.BytesIO()
        output_wb.save(buf)
        st.success(f"✅ Generert {success_count} sertifikater med korrekt antall/lengde!")
        st.download_button(
            "⬇️ Last ned",
            buf.getvalue(),
            file_name=f"sertifikater_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# =====================================================================
# EXCEL BATCH MODE
# =====================================================================

def render_excel_batch_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row):
    st.header("📂 Excel – flere slanger")

    with open(FLER_SLANGE_MAL, "rb") as file:
        st.download_button(
            label="Last ned MAL",
            data=file,
            file_name="MAL_slangebeskrivelse_flere_rader.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    uploaded_file = st.file_uploader(
        "Last opp utfylt MAL_slangebeskrivelse_flere_rader.xlsx", type=["xlsx"]
    )

    if uploaded_file is None:
        return

    try:
        import_df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Kunne ikke lese Excel: {e}")
        return

    st.subheader("Importerte rader")
    st.dataframe(import_df, use_container_width=True)

    st.divider()

    c1, c2 = st.columns(2)
    with c1:
        add_trykktest = st.checkbox("Legg til Trykktest")
    with c2:
        add_prikling = st.checkbox("Legg til Prikling")

    c3, c4 = st.columns(2)
    with c3:
        add_abs = st.checkbox("Type Approval (ABS)?")
    with c4:
        add_dnv = st.checkbox("Type Approval (DNV)?")

    pressure_details = {}
    if add_trykktest:
        st.subheader("Trykktest Detaljer")
        c1, c2 = st.columns(2)
        with c1:
            pressure_details["kunde"] = st.text_input("Kunde")
            pressure_details["kundens_best_nr"] = st.text_input("Kundens Best.nr")
        with c2:
            pressure_details["hydra_ordre_nr"] = st.text_input("Hydra Ordre.nr")

    st.divider()

    if not st.button("⚙️ Generer Output", use_container_width=True):
        return

    output_rows = []
    certificate_data_list = []

    for _, row in import_df.iterrows():
        summary_line = str(row.get("Slangebeskrivelse", "")).strip()
        if summary_line == "":
            continue

        antall = row.get("Antall", 1)
        material = str(row.get("Materiale", "")).strip().lower()
        try:
            antall = int(antall)
        except Exception:
            antall = 1

        pos_nr = row.get("POS.nr", "")
        kundes_del_nr = row.get("Kundes delnummer", "")
        lager_nr = row.get("Lager", "")

        selected_row, second_row1, second_row2, sheet_name, size_str, length_int = (
            core.find_matches_from_summary(summary_line, df1, df2_all, material_pref=material)
        )

        if selected_row is None:
            st.warning(f"Fant ikke slange: {summary_line}")
            continue

        second_rows = [second_row1, second_row2]

        if pos_nr:
            output_rows.append(["1", pos_nr, lager_nr, ""])
        if kundes_del_nr:
            output_rows.append(["1", kundes_del_nr, lager_nr, ""])

        output_rows.append(["1", summary_line, lager_nr, 1])

        hose_qty = length_int / 1000 if length_int else 1
        output_rows.append([selected_row["Prod.no"], selected_row["Beskrivelse"], lager_nr, hose_qty])

        # Kupling 2 missing -> treat it as the same as Kupling 1.
        if second_row1 is not None and second_row2 is None:
            second_row2 = second_row1
            second_rows = [second_row1, second_row2]

        same_coupling = (
            second_row1 is not None
            and second_row2 is not None
            and str(second_row1.get("Prod.no", "")).strip() == str(second_row2.get("Prod.no", "")).strip()
        )

        if same_coupling:
            # Kupling 1 and Kupling 2 are the same product -> one line,
            # Antall doubled, instead of two separate lines.
            output_rows.append([second_row1["Prod.no"], second_row1["Beskrivelse"], lager_nr, 2 * antall])
        else:
            for r in second_rows:
                if r is None:
                    continue
                output_rows.append([r["Prod.no"], r["Beskrivelse"], lager_nr, antall])

        gsm_count = sum(
            1 for r in second_rows if r is not None and str(r.get("Beskrivelse", "")).startswith("GSM")
        )

        if material == "stål":
            mat_prod = selected_row.get("Stål hylse(Posd.no)", "")
            mat_desc = selected_row.get("Stål hylse(beskrivelse)", "")
        else:
            mat_prod = selected_row.get("316 hylse(Posd.no)", "")
            mat_desc = selected_row.get("316 hylse(beskrivelse)", "")

        sheet_key = core._extract_sheet_key_from_sheetname(sheet_name)
        skip_staal_hylse = "(M-st)" in sheet_key or "(GSM)" in sheet_key

        if gsm_count < 2 and not skip_staal_hylse and mat_prod:
            hylse_qty = 2 if gsm_count == 0 else 1
            output_rows.append([mat_prod, mat_desc, lager_nr, hylse_qty * antall])

        mont_row = core.get_mont_row(size_str, sheet_name, mont_df)
        if mont_row is not None:
            output_rows.append([mont_row["Prod.no"], mont_row["Beskrivelse"], lager_nr, antall])

        if add_trykktest:
            trykk_row = core.get_trykktest_prodno(size_str, length_int, trykktest_df)
            if trykk_row is not None:
                output_rows.append([trykk_row["Prod.no"], trykk_row["Beskrivelse"], lager_nr, antall])

        if add_prikling:
            prikling_row = core.get_prikling_row(size_str, prikling_df)
            if prikling_row is not None:
                output_rows.append([prikling_row["Prod.no"], prikling_row["Beskrivelse"], lager_nr, antall])

        if add_dnv:
            dnv_cert_row = get_cert_row("90003")
            if dnv_cert_row is not None:
                output_rows.append(
                    [dnv_cert_row.get("Prod.no", ""), dnv_cert_row.get("Beskrivelse", ""), lager_nr, antall]
                )

        output_rows.append([1, "", lager_nr, ""])

        if add_trykktest:
            row_pressure_details = pressure_details.copy()
            row_pressure_details["antall_slanger"] = antall
            row_pressure_details["kundes_del_nr"] = kundes_del_nr

            certificate_data = core.fill_pressure_test_certificate_data(
                row_pressure_details, selected_row, second_rows, size_str, length_int, ""
            )
            certificate_data_list.append(certificate_data)

    if not output_rows:
        st.warning("Ingen rader generert.")
        return

    last_lager = output_rows[-1][2] if output_rows else ""

    if add_abs:
        abs_cert_row = get_cert_row("90478")
        if abs_cert_row is not None:
            output_rows.append(["1", "", last_lager, ""])
            output_rows.append(
                [abs_cert_row.get("Prod.no", ""), abs_cert_row.get("Beskrivelse", ""), last_lager, 1]
            )

    if add_dnv:
        dnv_cert_row = get_cert_row("90003")
        if dnv_cert_row is not None:
            output_rows.append(["1", "", last_lager, ""])
            output_rows.append(
                [dnv_cert_row.get("Prod.no", ""), dnv_cert_row.get("Beskrivelse", ""), last_lager, 1]
            )

    wb = core.create_output_workbook(output_rows)

    if add_trykktest:
        for i, cert_data in enumerate(certificate_data_list, start=1):
            wb = core.add_certificate_sheet(wb, CERT_TEMPLATE, cert_data, f"Sertifikat {i}")

    wb = core.add_sluttkontroll_sheet(
        wb, SLUTT_TEMPLATE,
        kunde=pressure_details.get("kunde", ""),
        hydra_ordre_nr=pressure_details.get("hydra_ordre_nr", ""),
    )

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    st.success(f"✅ {len(import_df)} slanger prosessert.")
    st.download_button(
        "📥 Last ned Output.xlsx",
        buffer,
        file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# =====================================================================
# ORDER PREVIEW (common to Quick / Full / Excel batch)
# =====================================================================

def render_output_preview():
    st.divider()
    st.header("📊 Foreløpig slangestruktur i Visma")

    if not st.session_state.output_rows:
        return

    output_df = pd.DataFrame(
        st.session_state.output_rows, columns=["Prod.no", "Beskrivelse", "Lager", "Antall"]
    )
    st.dataframe(output_df, use_container_width=True, hide_index=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        if st.button("🗑️ Slett siste", use_container_width=True):
            if st.session_state.output_rows and st.session_state.output_batches:
                last_batch_size = st.session_state.output_batches.pop()
                if last_batch_size > 0:
                    st.session_state.output_rows = st.session_state.output_rows[:-last_batch_size]
            st.rerun()

    with c2:
        if st.button("🧹 Tøm alt", use_container_width=True):
            st.session_state.output_rows = []
            st.session_state.certificate_data_list = []
            st.session_state.abs_selected_any = False
            st.rerun()

    with c3:
        excel_buffer = generate_excel()
        st.download_button(
            label="⬇️ Last ned Excel",
            data=excel_buffer,
            file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# =====================================================================
# HEADER
# =====================================================================

def render_header():
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.image("assets/logo.png", width='stretch')
    st.title("🔎 Eivinds Slangeprogram")


# =====================================================================
# MAIN
# =====================================================================

def main():
    inject_theme()

    try:
        df1, df2_all, mont_df, trykktest_df, prikling_df = load_all()
    except Exception as e:
        st.error(f"❌ Kunne ikke laste data: {str(e)}")
        st.stop()

    abs_sert_df = core.clean_columns(pd.read_excel(FIRST_FILE, sheet_name="ABS Sert."))
    get_cert_row = make_cert_row_lookup(abs_sert_df)

    init_session_state()
    # generate_excel() needs the cert lookup but takes no arguments (kept for
    # a stable, easy-to-call signature) - stash it in session state.
    st.session_state.get_cert_row = get_cert_row

    if st.session_state.get("full_abs", False):
        st.session_state.abs_selected_any = True

    render_header()
    st.divider()

    mode_choice = st.radio(
        "Velg funksjon:",
        options=list(MODE_LABELS.values()),
        index=0,
        key="mode_radio",
        horizontal=True,
    )
    st.session_state.input_mode = LABEL_TO_MODE[mode_choice]
    if st.session_state.input_mode != "certificate":
        st.session_state.pop("cert_df", None)

    st.divider()

    mode = st.session_state.input_mode

    if mode == "certificate":
        render_certificate_mode(df1, df2_all, get_cert_row)
        return  # certificate mode has its own download flow; no order preview

    if mode == "quick":
        render_quick_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row)
    elif mode == "full":
        render_full_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row)
    elif mode == "excel_batch":
        render_excel_batch_mode(df1, df2_all, mont_df, trykktest_df, prikling_df, get_cert_row)

    render_output_preview()


if __name__ == "__main__":
    main()
