# -*- coding: utf-8 -*-
"""
Slangeprogram - Streamlit Version
Created on Tue Feb 24 11:33:34 2026

@author: eivind
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import io
import os
import base64
from st_aggrid import AgGrid, GridOptionsBuilder

import core



# -------------------------------------------------
# CONFIG
# -------------------------------------------------

st.set_page_config(page_title="Slangeprogram", layout="wide", page_icon="assets/HP_icon.ico")

# -------------------------------------------------
# CUSTOM STYLING
# -------------------------------------------------


def set_background(image_path):
    with open(image_path, "rb") as img_file:
        encoded = base64.b64encode(img_file.read()).decode()

    st.markdown(
    f"""
    <style>

        
        
        /* Add/Update this in your st.markdown block */
        .ag-theme-streamlit .center-header .ag-header-cell-label {{
            justify-content: center !important;
        }}
        
        /* This ensures the text itself is centered if it wraps */
        .ag-theme-streamlit .center-header .ag-header-cell-text {{
            text-align: center !important;
            width: 100%;
        }}


        

    
    
        /* ============================================================
           === BAKGRUNN OG GLOBALT UTSEENDE ===
           ============================================================ */
    
        .stApp {{
            background-image: url("data:image/jpg;base64,{encoded}");
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
    
        .stApp::before {{
            content: "";
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.65);
            z-index: 0;
        }}
    
        /* Global hvit tekst */
        .stMarkdown, .stText, .stHeader, .stSubheader,
        label, h1, h2, h3, h4, h5, h6 {{
            color: white !important;
        }}
    
        /* ============================================================
           === INPUTFELT ===
           ============================================================ */
    
        .stTextInput input,
        .stNumberInput input,
        .stTextArea textarea {{
            color: black !important;
            background-color: rgba(255,255,255,0.9) !important;
        }}
    
        ::placeholder {{
            color: #444 !important;
            opacity: 1 !important;
        }}
    
        /* Selectbox */
        .stSelectbox div[data-baseweb="select"] * {{
            color: black !important;
        }}
    
        ul[role="listbox"] li {{
            color: black !important;
        }}
    
        /* ============================================================
           === KNAPPER ===
           ============================================================ */
    
        .stButton > button {{
            background-color: white !important;
            color: black !important;
            border: 1px solid #ccc !important;
            padding: 0.6rem 1.2rem !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
        }}
    
        .stButton > button:hover {{
            background-color: #f2f2f2 !important;
            border: 1px solid #999 !important;
            color: black !important;
        }}
    
        /* ============================================================
           === DATAFRAME (ikke AG‑Grid) ===
           ============================================================ */
    
        .stDataFrame tbody tr td {{
            color: white !important;
            background-color: rgba(0,0,0,0.6) !important;
        }}
    
        .stDataFrame thead tr th {{
            color: white !important;
            background-color: rgba(0,0,0,0.8) !important;
        }}
    
        .stDataFrame tbody tr {{
            border-bottom: 1px solid rgba(255,255,255,0.2) !important;
        }}
    
        .stDataFrame tbody tr:hover td {{
            background-color: rgba(255,255,255,0.1) !important;
        }}
    
        /* ============================================================
           === RADIO, CHECKBOX, INFOBOKSER ===
           ============================================================ */
    
        .stRadio div[role="radiogroup"] p {{
            color: white !important;
        }}
    
        .stCheckbox label > div > div {{
            color: white !important;
        }}
    
        .stAlert [data-testid="stMarkdownContainer"] {{
            color: white !important;
        }}
    
        /* Fjern hvit toppstripe */
        header[data-testid="stHeader"],
        header[data-testid="stHeader"]::before {{
            background: transparent !important;
        }}
    
        /* ============================================================
           === AG‑GRID BORDERS ===
           ============================================================ */
    
        .ag-root-wrapper {{
            border: 2px solid black !important;
        }}
    
        .ag-cell {{
            border-right: 2px solid black !important;
            border-bottom: 2px solid black !important;
        }}
    
        .ag-header-cell {{
            border-right: 2px solid black !important;
            border-bottom: 2px solid black !important;
        }}
    
    </style>
    """,
    unsafe_allow_html=True
)

set_background("assets/background.png")

FIRST_FILE = "Slanger_hylser.xlsx"
SECOND_FILE = "kuplinger_316.xlsx"
CERT_TEMPLATE = "Mal Trykktest Sertikat.xlsx"
SLUTT_TEMPLATE = "Mal sluttkontroll slanger.xlsx"


# -------------------------------------------------
# LOAD DATA - WITH ERROR HANDLING
# -------------------------------------------------

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


try:
    df1, df2_all, mont_df, trykktest_df, prikling_df = load_all()
except Exception as e:
    st.error(f"❌ Kunne ikke laste data: {str(e)}")
    st.stop()
abs_sert_df = core.clean_columns(
    pd.read_excel(FIRST_FILE, sheet_name="ABS Sert.")
)

# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------

if "abs_selected_any" not in st.session_state:
    st.session_state.abs_selected_any = False

if "output_rows" not in st.session_state:
    st.session_state.output_rows = []

if "certificate_data_list" not in st.session_state:
    st.session_state.certificate_data_list = []

if "pos_counter" not in st.session_state:
    st.session_state.pos_counter = 1

if "input_mode" not in st.session_state:
    st.session_state.input_mode = "quick"

if "selected_hose_row" not in st.session_state:
    st.session_state.selected_hose_row = None

if "selected_c1_row" not in st.session_state:
    st.session_state.selected_c1_row = None

if "selected_c2_row" not in st.session_state:
    st.session_state.selected_c2_row = None

if "full_df2" not in st.session_state:
    st.session_state.full_df2 = None


# -------------------------------------------------
# HELPER FUNCTIONS
# -------------------------------------------------
if st.session_state.get("full_abs", False):
    st.session_state.abs_selected_any = True
    
def process_and_add_hose(selected_row, second_row1, second_row2, sheet_name_found, size_str, 
                        length_int, material, lager, pos_mark, posnr, input_linje, inputlinje, pressure_test, 
                        pressure_details, antall_slanger,prikling=False, first_line="", angle=""):
    """Process hose data and add to output rows"""
    rows = []

    if pos_mark and posnr:
        rows.append(["1", f"POS: {posnr}", int(lager), 1])
        try:
            st.session_state.pos_counter = int(posnr) + 1
        except:
            pass
    
    if input_linje and inputlinje:
        rows.append(["1", f"{inputlinje}", int(lager), 1])
        
        
    

    if first_line:
        # Quick mode - just use the first line as-is
        rows.append(["1", first_line, int(lager), 1])
    else:
        # Full mode - build first line from components with angle if provided
        part1 = str(selected_row["Beskrivelse"])[:7] if selected_row is not None else ""
        part2 = str(length_int if length_int else "")
        part3 = str(second_row1["Beskrivelse"])[:9 if material == "stål" else 15] if second_row1 is not None else ""
        part4 = str(second_row2["Beskrivelse"])[:9 if material == "stål" else 15] if second_row2 is not None else ""
        
        if angle and angle.strip():
            first_line_display = f"{part1}/{part2}/{part3}/{part4}/{angle}°"
        else:
            first_line_display = f"{part1}/{part2}/{part3}/{part4}"
        rows.append(["1", first_line_display, int(lager), 1])

    # Add products
    if selected_row is not None:
        try:
            qty = round((length_int or 1000) / 1000, 3)
            rows.append([selected_row["Prod.no"], selected_row["Beskrivelse"], int(lager), qty])
        except Exception:
            rows.append([selected_row.get("Prod.no", ""), selected_row.get("Beskrivelse", ""), int(lager), 1])
    else:
        rows.append(["", "Fant ikke første produkt", int(lager), 1])

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

    sheet_key = core._extract_sheet_key_from_sheetname(sheet_name_found) if sheet_name_found else "(st)" if material == "stål" else "(316)"
    skip_staal_hylse = "(M-st)" in sheet_key or "(GSM)" in sheet_key

    if gsm_count < 2 and not skip_staal_hylse and mat_prod:
        stahl_value = 2 if gsm_count == 0 else 1
        rows.append([mat_prod, mat_desc, int(lager), stahl_value])

    mont_row = core.get_mont_row(size_str, sheet_key, mont_df)
    if mont_row is not None:
        rows.append([mont_row["Prod.no"], mont_row["Beskrivelse"], int(lager), 1])
    # --- Add Prikling if selected ---
    if prikling and size_str:
        prikling_row = core.get_prikling_row(size_str, prikling_df)
        if prikling_row is not None:
            rows.append([
                prikling_row["Prod.no"],
                prikling_row["Beskrivelse"],
                int(lager),
                1
            ])

    if pressure_test:
        trykktest_row = core.get_trykktest_prodno(size_str, length_int or 1000, trykktest_df)
        if trykktest_row is not None:
            rows.append([trykktest_row["Prod.no"], trykktest_row["Beskrivelse"], int(lager), 1])
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
            "pressure_details": pressure_details
        })
def generate_excel():
    rows_for_excel = [r.copy() for r in st.session_state.output_rows]

    # -------------------------------------------------
    # ADD ABS CERT ROW (ONLY ONCE, ALWAYS AT BOTTOM)
    # -------------------------------------------------
    
    if st.session_state.abs_selected_any and not abs_sert_df.empty:
    
        lager_value = rows_for_excel[-1][2] if rows_for_excel else 3
    
        # spacer line
        rows_for_excel.append(["1", "", lager_value, ""])
    
        abs_row = abs_sert_df.iloc[0]
    
        rows_for_excel.append([
            abs_row.get("Prod.no", ""),
            abs_row.get("Beskrivelse", ""),
            lager_value,
            1
        ])
    
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
                    cert_info["material"]
                )

                if cert_data:
                    sheet_name = f"Sertifikat {idx}" if len(st.session_state.certificate_data_list) > 1 else "Trykktest Sertifikat"
                    output_wb = core.add_certificate_sheet(
                        output_wb,
                        CERT_TEMPLATE,
                        cert_data,
                        sheet_name
                    )
            except Exception as e:
                st.warning(f"Kunne ikke legge til sertifikat {idx}: {e}")

    try:
        kunde = ""
        hydra_ordre_nr = ""
        if st.session_state.certificate_data_list:
            kunde = st.session_state.certificate_data_list[0]["pressure_details"].get("kunde", "")
            hydra_ordre_nr = st.session_state.certificate_data_list[0]["pressure_details"].get("hydra_ordre_nr", "")

        output_wb = core.add_sluttkontroll_sheet(
            output_wb,
            SLUTT_TEMPLATE,
            kunde=kunde,
            hydra_ordre_nr=hydra_ordre_nr
        )
    except Exception as e:
        st.warning(f"Kunne ikke legge til sluttkontroll: {e}")

    output_buffer = io.BytesIO()
    output_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# -------------------------------------------------
# MAIN UI
# -------------------------------------------------




# Image at top
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("assets/logo.png", use_column_width=True)

st.title("🔎 Eivinds Slangeprogram")


st.divider()
# Mode selection
col1, col2 = st.columns(2)
with col1:
    mode_choice = st.radio(
        "Vil du skrive inn slangebeskrivelse eller velge slanger og kuplinger?",
        options=["⌨️ Skriv inn Slangebeskrivelse", "🖱 Velg Slange og Kuplinger"],
        index=0,
        key="mode_radio"
    )
# Update session state based on selection
if mode_choice == "⌨️ Skriv inn Slangebeskrivelse":
    st.session_state.input_mode = "quick"
else:
    st.session_state.input_mode = "full"


# -------------------------------------------------
# QUICK MODE
# -------------------------------------------------

if st.session_state.input_mode == "quick":
    st.header("➕ Skriv in Slangebeskrivelse")

    col1, col2 = st.columns([2, 1])

    with col1:
        first_line = st.text_input("Slangebeskrivelse (Bindestreker må være med 😒)", placeholder="Slange/Lengde/Kupling 1/Kupling 2", key="quick_first_line")

    with col2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="quick_material")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        lager = st.selectbox("Lager",
                             options=["3", "1", "5"],
                             format_func=lambda x: {"3": "Lillestrøm", "1": "Ålesund", "5": "Trondheim"}[x],
                             key="quick_lager")

    with col2:
        antall_slanger = st.number_input("Antall slanger", min_value=1, value=1, key="quick_antall")

    with col3:
        type_approval = st.checkbox("Type Approval (DNV)?", key="quick_type_approval")
        
    with col4:
        type_approval1 = st.checkbox("Type Approval (ABS)?", key="quick_type_approval1")

    col1, col2 = st.columns([1, 2])
    with col1:
        pos_mark = st.checkbox("Merke med POS.nr?", key="quick_pos_mark")
    with col2:
        if pos_mark:
            posnr = st.text_input("POS.nr", value=str(st.session_state.pos_counter), key="quick_posnr")
        else:
            posnr = ""

    col3, col4 = st.columns([1, 2])
    with col1:
        input_linje = st.checkbox("Merke med kundes delnummer? ", key="quick_input_linje")
    with col2:
        if input_linje:
            inputlinje = st.text_input("Kundes delnummer:", key="quick_inputlinje")
        else:
            inputlinje = ""

    st.divider()
    # --- Prikling ---
    prikling = st.checkbox("🪛 Skal slangen prikles?", key="full_prikling")
    
    # --- Trykktest ---
    if type_approval or type_approval1:
        pressure_test = True
        st.checkbox(
            "🚰 Skal slangen trykktestes?",
            value=True,
            disabled=True,
            key="quick_pressure_test"
        )
    else:
        pressure_test = st.checkbox(
            "🚰 Skal slangen trykktestes?",
            key="quick_pressure_test"
        )

    pressure_details = {
        "kunde": "",
        "kundens_best_nr": "",
        "hydra_ordre_nr": "",
        "kundes_del_nr": "",
        "antall_slanger": antall_slanger,
        "angle": ""
    }

    if pressure_test:
        st.subheader("📋 Trykktest Detaljer")
        col1, col2 = st.columns(2)
        with col1:
            pressure_details["kunde"] = st.text_input("Kunde", key="quick_kunde")
            pressure_details["kundens_best_nr"] = st.text_input("Kundens best. Nr.", key="quick_best_nr")
        with col2:
            pressure_details["hydra_ordre_nr"] = st.text_input("Hydra Pipe ordre nr.", key="quick_hydra_ordre")
            # Hvis input_linje er valgt → bruk inputlinje som kundes_del_nr og ikke vis feltet
            if input_linje and inputlinje:
                pressure_details["kundes_del_nr"] = inputlinje
            else:
                # Vis feltet for kundes_del_nr kun hvis input_linje IKKE er valgt
                pressure_details["kundes_del_nr"] = st.text_input("Kundes del nr.", key="quick_del_nr")

    if st.button("✅ Legg til slange", use_container_width=True, key="quick_add_btn"):
        if not first_line:
            st.error("Første utdata-linje må oppgis!")
        else:
            selected_row, second_row1, second_row2, sheet_name_found, size_str, length_int = core.find_matches_from_summary(
                first_line, df1, df2_all, material_pref=material
            )
    
        # Sett kundes_del_nr riktig
        if input_linje and inputlinje:
            pressure_details["kundes_del_nr"] = inputlinje
            
        if type_approval1:
            st.session_state.type_approval1 = True
    
        # 🚀 Denne må ALLTID kjøres, uansett input_linje
        process_and_add_hose(
            selected_row, second_row1, second_row2, sheet_name_found, size_str,
            length_int, material, lager, pos_mark, posnr, input_linje, inputlinje, pressure_test,
            pressure_details, antall_slanger, prikling=prikling, first_line=first_line
        )
    
        if type_approval1:
            st.session_state.abs_selected_any = True
        st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")


# -------------------------------------------------
# FULL MODE
# -------------------------------------------------

elif st.session_state.input_mode == "full":

    st.header("📝 Velg Slange og Kuplinger")

    st.subheader("1️⃣ Velg slange")

    # Type Approval FIRST - before search and table
    col1, col2 = st.columns([2, 1])
    with col2:
        type_approval1 = st.checkbox("Type Approval (ABS)?", key="full_type_approval1")
    with col1:
        type_approval = st.checkbox("Type Approval (DNV)?", key="full_type_approval")

    # Search hose
    search = st.text_input("Søk etter slange", key="full_search")

    # -------------------------------------------------
    # TYPE APPROVAL FILTERING (DNV + ABS)
    # -------------------------------------------------
    
    filtered_df = df1.copy()
    
    dnv_col = "Type Approval"
    abs_col = "Type Approval1"
    
    if type_approval and type_approval1:
        # BOTH required
        filtered_df = filtered_df[
            filtered_df[dnv_col].fillna("").astype(str).str.strip().ne("") &
            filtered_df[abs_col].fillna("").astype(str).str.strip().ne("")
        ]
    
    elif type_approval:
        # Only DNV
        filtered_df = filtered_df[
            filtered_df[dnv_col].fillna("").astype(str).str.strip().ne("")
        ]
    
    elif type_approval1:
        # Only ABS
        filtered_df = filtered_df[
            filtered_df[abs_col].fillna("").astype(str).str.strip().ne("")
        ]
    
    # else: no filtering


    # -------------------------------------------------
    # APPLY SEARCH FILTER
    # -------------------------------------------------
    
    if search:
        st.session_state.selected_hose_row = None
        filtered_df = filtered_df[
            filtered_df["Beskrivelse_2"]
            .astype(str)
            .str.contains(search, case=False, na=False)
        ]
    
    st.write("**Velg slange fra tabellen under:**")
    
    event = None
    selected_row = None
    
    # -------------------------------------------------
    # TABLE DISPLAY (only if rows exist)
    # -------------------------------------------------
    
   
    
    # -------------------------------------------------
    # TABLE SELECTION (SAFE)
    # -------------------------------------------------
    
    

    

    # Velg kolonner
    df_view = filtered_df[["Prod.no", "Beskrivelse", "Beskrivelse_2", "Dimensjon", "Trykk(bar)"]]
    
    # Bygg grid options
    gb = GridOptionsBuilder.from_dataframe(df_view)
    
    gb.configure_column("Beskrivelse", hide=True)
    
    gb.configure_column("Prod.no", headerName="Artikkel nummer")
    gb.configure_column("Beskrivelse_2", headerName="Beskrivelse")
    gb.configure_column("Dimensjon", headerName="Dimensjon")
    gb.configure_column("Trykk(bar)", headerName="Arbeidstrykk (Bar)")
    
    # Midtstill celler
    gb.configure_default_column(
        
        headerClass="center-header", # Make sure this matches your CSS
        cellStyle={
            "display": "flex",
            "justifyContent": "center",
            "alignItems": "center",
            "textAlign": "center"
        }
    )
    
    # Aktiver radvalg
    gb.configure_selection(
        selection_mode="single",
        use_checkbox=False
    )
    
    # Bygg gridOptions
    grid_options = gb.build()
    
    custom_css = {
        ".ag-header-cell-label": {"justify-content": "center"},
        ".ag-header-cell-text": {"text-align": "center", "width": "100%"}
    }

    
    # Kjør AG‑Grid med ny API
    grid_response = AgGrid(
        df_view,
        gridOptions=grid_options,
        custom_css=custom_css,
        update_on=["selectionChanged"],   # ← NY API
        fit_columns_on_grid_load=True,
        theme="streamlit"                 # ← sikrer riktig CSS‑tema
    )
    
    # Hent valgt rad
    selected_df = grid_response["selected_rows"]
    

    if selected_df is not None and not selected_df.empty:
        selected_row = selected_df.iloc[0].to_dict()
        st.session_state.selected_hose_row = selected_row
            
        
    

    # -------------------------------------------------
    # FINAL STATUS
    # -------------------------------------------------
    
    if st.session_state.selected_hose_row is not None:
        selected_row = st.session_state.selected_hose_row
        st.success(f"✅ Valgt: {selected_row['Beskrivelse_2']}")
    else:
        st.warning("⚠️ Du må velge slange fra tabellen.")
    
    
    
    
    # Options (moved AFTER selection)
    col1, col2, col3 = st.columns(3)

    with col1:
        length = st.number_input("Lengde (mm)", value=1000, key="full_length")

    with col2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="full_material")

    with col3:
        st.write("")  # spacer

    if selected_row is not None:

        size = str(selected_row["Dimensjon"]).zfill(2)
    
        # Determine sheet_name based on type approval and material
        if material == "syrefast":
            try:
                slange_hylse_df = core.clean_columns(pd.read_excel(FIRST_FILE, sheet_name="Slange+Hylse"))
                prod_no = selected_row.get("Prod.no")
                match = slange_hylse_df.loc[slange_hylse_df["Prod.no"] == prod_no]
                if not match.empty and len(slange_hylse_df.columns) > 11:
                    col_l_val = str(match.iloc[0, 11])
                    if "5" in col_l_val:
                        sheet_name = f"Kuplinger {size}(5-316)"
                    else:
                        sheet_name = f"Kuplinger {size}(316)"
                else:
                    sheet_name = f"Kuplinger {size}(316)"
            except:
                sheet_name = f"Kuplinger {size}(316)"
        else:  # stål
            type_approval_val = type_approval
            gates_in_k = False
            
            # Check for Type Approval with Gates in column K
            if type_approval_val:
                try:
                    slange_hylse_df = core.clean_columns(pd.read_excel(FIRST_FILE, sheet_name="Slange+Hylse"))
                    prod_no = selected_row.get("Prod.no")
                    match = slange_hylse_df.loc[slange_hylse_df["Prod.no"] == prod_no]
                    if not match.empty and len(slange_hylse_df.columns) > 10:
                        col_k_val = str(match.iloc[0, 10])
                        if "Gates" in col_k_val:
                            gates_in_k = True
                except:
                    pass
            
            # Determine sheet key
            if type_approval_val and gates_in_k:
                sheet_key = "(M-st)"
                sheet_name = f"Kuplinger {size}(M-st)"
            else:
                desc = str(selected_row.get("Beskrivelse", ""))
                if len(desc) > 2 and desc[0] == "G" and desc[2] == "K":
                    if desc.startswith("G5K-24") or desc.startswith("G6K-24"):
                        sheet_name = f"Kuplinger {size}(GSM)"
                    else:
                        sheet_name = f"Kuplinger {size}(GS)"
                else:
                    sheet_name = f"Kuplinger {size}(st)"

        if sheet_name not in df2_all:
            st.error(f"Fant ikke ark: {sheet_name}")
            st.stop()
    
        df2 = df2_all[sheet_name]
        st.session_state.full_df2 = df2
    
        

        # -------------------------------------------------
        # COUPLINGS
        # -------------------------------------------------
        
        st.divider()
        st.subheader("2️⃣ Velg kuplinger")
        
        col1, col2 = st.columns(2)
        
        # -------------------------
        # Kupling 1
        # -------------------------
        
        with col1:
            st.write("**Kupling 1**")
            st.write("Velg kupling fra tabellen:")
        
            

            gb1 = GridOptionsBuilder.from_dataframe(
                df2[["Prod.no", "Beskrivelse"]]
            )
            
            gb1.configure_default_column(
                headerClass="center-header", 
                cellStyle={"display": "flex", "justifyContent": "center", "alignItems": "center"}
            )
            
            # Single row selection without checkbox
            gb1.configure_selection(
                selection_mode="single",
                use_checkbox=False
            )
            
            custom_css = {
                ".ag-header-cell-label": {"justify-content": "center"},
                ".ag-header-cell-text": {"text-align": "center", "width": "100%"}
            }
            
            grid_response1 = AgGrid(
                df2[["Prod.no", "Beskrivelse"]],
                gridOptions=gb1.build(),
                custom_css=custom_css,
                update_mode="SELECTION_CHANGED",
                fit_columns_on_grid_load=True,
                key="coupling1_grid"
            )
        
            selected_df1 = grid_response1["selected_rows"]
        
            if selected_df1 is not None and not selected_df1.empty:
                st.session_state.selected_c1_row = selected_df1.iloc[0].to_dict()
        
            if st.session_state.selected_c1_row is not None:
                st.write(f"✅ Valgt: *{st.session_state.selected_c1_row['Beskrivelse']}*")
            else:
                st.info("Velg kupling fra tabellen")
        
        
        # -------------------------
        # Kupling 2
        # -------------------------
        
        with col2:
            st.write("**Kupling 2**")
            st.write("Velg kupling fra tabellen:")
        
            gb2 = GridOptionsBuilder.from_dataframe(
                df2[["Prod.no", "Beskrivelse"]]
            )
            
            # Center all cells
            gb2.configure_default_column(
                headerClass="center-header", 
                cellStyle={"display": "flex", "justifyContent": "center", "alignItems": "center"}
            )
            gb2.configure_selection(
                selection_mode="single",
                use_checkbox=False
            )
        
            custom_css = {
                ".ag-header-cell-label": {"justify-content": "center"},
                ".ag-header-cell-text": {"text-align": "center", "width": "100%"}
            }
            
            grid_response2 = AgGrid(
                df2[["Prod.no", "Beskrivelse"]],
                gridOptions=gb2.build(),
                custom_css=custom_css,
                update_mode="SELECTION_CHANGED",
                fit_columns_on_grid_load=True,
                key="coupling2_grid"
            )
        
            selected_df2 = grid_response2["selected_rows"]
        
            if selected_df2 is not None and not selected_df2.empty:
                st.session_state.selected_c2_row = selected_df2.iloc[0].to_dict()
        
            if st.session_state.selected_c2_row is not None:
                st.write(f"✅ Valgt: *{st.session_state.selected_c2_row['Beskrivelse']}*")
            else:
                st.info("Velg kupling fra tabellen")
        
        
        # -------------------------
        # VALIDATION
        # -------------------------
        
        if (
            st.session_state.selected_c1_row is None
            or st.session_state.selected_c2_row is None
        ):
            st.warning("⚠️ Du må velge kuplinger i begge ender")
            st.stop()
        
        row_c1 = st.session_state.selected_c1_row
        row_c2 = st.session_state.selected_c2_row
    
        # -------------------------------------------------
        # ADDITIONAL OPTIONS
        # -------------------------------------------------
    
        st.divider()
        st.subheader("3️⃣ Innstillinger")
    
        col1, col2, col3, col4 = st.columns(4)
    
        with col1:
            lager = st.selectbox("Lager",
                                 options=["3", "1", "5"],
                                 format_func=lambda x: {"3": "Lillestrøm", "1": "Ålesund", "5": "Trondheim"}[x],
                                 key="full_lager")
    
        with col2:
            antall_slanger = st.number_input("Antall slanger", min_value=1, value=1, key="full_antall")
    
        with col3:
            pos_mark = st.checkbox("Merke med POS.nr?", key="full_pos_mark")
    
        if pos_mark:
            posnr = st.text_input("POS.nr", value=str(st.session_state.pos_counter), key="full_posnr")
        else:
            posnr = ""
    
        with col4:
            input_linje = st.checkbox("Merke med kundes delnummer?", key="full_input_linje")
    
        if input_linje:
            inputlinje = st.text_input("Kundes delnummer: ",  key="full_inputlinje")
        else:
            inputlinje = ""    
    
        # Check if either coupling has angle (45 or 90)
        has_angle_c1 = "45" in str(row_c1["Beskrivelse"]) or "90" in str(row_c1["Beskrivelse"])
        has_angle_c2 = "45" in str(row_c2["Beskrivelse"]) or "90" in str(row_c2["Beskrivelse"])
        
        # Show angle input only if one of the couplings has angle
        angle = ""
        if has_angle_c1 and has_angle_c2:
            st.divider()
            st.subheader("📐 Vinkel")
            angle = st.text_input("Skriv inn vinkel", key="full_angle")
    
        # Pressure test
        st.divider()
        # --- Prikling ---
        prikling = st.checkbox("🪛 Skal slangen prikles?", key="full_prikling")
        
        # --- Trykktest ---
        # --- Trykktest ---
        if type_approval or type_approval1:
            pressure_test = True
            st.checkbox(
                "🚰 Skal slangen trykktestes?",
                value=True,
                disabled=True,
                key="full_pressure_test"
            )
        else:
            pressure_test = st.checkbox(
                "🚰 Skal slangen trykktestes?",
                key="full_pressure_test"
            )
    
        pressure_details = {
            "kunde": "",
            "kundens_best_nr": "",
            "hydra_ordre_nr": "",
            "kundes_del_nr": "",
            "antall_slanger": antall_slanger,
            "angle": angle
        }
    
        if pressure_test:
            st.subheader("📋 Trykktest Detaljer")
            col1, col2= st.columns(2)
            
            with col1:
                pressure_details["kunde"] = st.text_input("Kunde", key="full_kunde")
                pressure_details["kundens_best_nr"] = st.text_input("Kundens best. Nr.", key="full_best_nr")
            with col2:
                pressure_details["hydra_ordre_nr"] = st.text_input("Hydra Pipe ordre nr.", key="full_hydra_ordre")
                # Hvis input_linje er valgt, skal kundes_del_nr IKKE vises som inputfelt
                if input_linje and inputlinje:
                    pressure_details["kundes_del_nr"] = inputlinje
                else:
                    pressure_details["kundes_del_nr"] = st.text_input("Kundes del nr.", key="full_del_nr")
    
            # Add to order
        if st.button("✅ Legg til slange", use_container_width=True, key="full_add_btn"):
            # Update pressure_details with angle for certificate
            pressure_details["angle"] = angle
            if input_linje and inputlinje:
                pressure_details["kundes_del_nr"] = inputlinje
            process_and_add_hose(
                selected_row, row_c1, row_c2, sheet_name, size,
                length, material, lager, pos_mark, posnr, input_linje, inputlinje, pressure_test,
                pressure_details, antall_slanger, prikling=prikling, first_line="", angle=angle
            )
    
            # Reset selections
            st.session_state.selected_hose_row = None
            st.session_state.selected_c1_row = None
            st.session_state.selected_c2_row = None
            
            if type_approval1:
                st.session_state.abs_selected_any = True
            
            st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")
        
        
# -------------------------------------------------
# ORDER PREVIEW (Common to both modes)
# -------------------------------------------------

st.divider()
st.header("📊 Foreløpig slangestruktur i Visma")

if type_approval1 is True:
    st.info("For ABS gokjenning trengs det bevitnelse av trykktesting ")
    st.info("Du finner Type Approval fra ABS her på Felles")
    
if type_approval is True:
    st.info("Du finner Type Approval fra DNV her på Felles")

if st.session_state.output_rows:
    output_df = pd.DataFrame(st.session_state.output_rows, columns=["Prod.no", "Beskrivelse", "Lager", "Antall"])
    st.dataframe(output_df, use_container_width=True, hide_index=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("🗑️ Slett siste", use_container_width=True):
            if len(st.session_state.output_rows) > 0:
                st.session_state.output_rows.pop()
            st.rerun()

    with col2:
        if st.button("🧹 Tøm alt", use_container_width=True):
            st.session_state.output_rows = []
            st.session_state.certificate_data_list = []
            st.session_state.abs_selected_any = False
            st.rerun()

    with col3:
        excel_buffer = generate_excel()
    
        st.download_button(
            label="⬇️ Last ned Excel",
            data=excel_buffer,
            file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
   
