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

import core


# -------------------------------------------------
# CONFIG
# -------------------------------------------------

st.set_page_config(page_title="Slangeprogram", layout="wide")

FIRST_FILE = "Slanger_hylser.xlsx"
SECOND_FILE = "kuplinger_316.xlsx"
CERT_TEMPLATE = "Mal Trykktest Sertikat.xlsx"
SLUTT_TEMPLATE = "Mal sluttkontroll slanger.xlsx"


# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------

@st.cache_data
def load_all():
    df1, df2_all = core.load_main_data(FIRST_FILE, SECOND_FILE)
    mont_df, trykktest_df, prikling_df = core.load_support_sheets(FIRST_FILE)
    return df1, df2_all, mont_df, trykktest_df, prikling_df


df1, df2_all, mont_df, trykktest_df, prikling_df = load_all()


# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------

if "output_rows" not in st.session_state:
    st.session_state.output_rows = []

if "certificate_data_list" not in st.session_state:
    st.session_state.certificate_data_list = []

if "pos_counter" not in st.session_state:
    st.session_state.pos_counter = 1

if "input_mode" not in st.session_state:
    st.session_state.input_mode = "quick"  # "quick" or "full"

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

def process_and_add_hose(selected_row, second_row1, second_row2, sheet_name_found, size_str, 
                        length_int, material, lager, pos_mark, posnr, pressure_test, 
                        pressure_details, antall_slanger, first_line=""):
    """Process hose data and add to output rows"""
    rows = []

    # Add POS marking if enabled
    if pos_mark and posnr:
        rows.append(["1", f"POS: {posnr}", int(lager), 1])
        try:
            st.session_state.pos_counter = int(posnr) + 1
        except:
            pass

    # Add first line
    if first_line:
        rows.append(["1", first_line, int(lager), 1])
    else:
        # Build first line from components
        part1 = str(selected_row["Beskrivelse"])[:7] if selected_row is not None else ""
        part2 = str(length_int if length_int else "")
        part3 = str(second_row1["Beskrivelse"])[:9 if material == "stål" else 15] if second_row1 is not None else ""
        part4 = str(second_row2["Beskrivelse"])[:9 if material == "stål" else 15] if second_row2 is not None else ""
        first_line = f"{part1}/{part2}/{part3}/{part4}"
        rows.append(["1", first_line, int(lager), 1])

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

    # Check for GSM
    gsm_count = 0
    if second_row1 is not None and str(second_row1.get("Beskrivelse", "")).startswith("GSM"):
        gsm_count += 1
    if second_row2 is not None and str(second_row2.get("Beskrivelse", "")).startswith("GSM"):
        gsm_count += 1

    # Add material hylse
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

    # Add mont row
    mont_row = core.get_mont_row(size_str, sheet_key, mont_df)
    if mont_row is not None:
        rows.append([mont_row["Prod.no"], mont_row["Beskrivelse"], int(lager), 1])

    # Add pressure test if selected
    if pressure_test:
        trykktest_row = core.get_trykktest_prodno(size_str, length_int or 1000, trykktest_df)
        if trykktest_row is not None:
            rows.append([trykktest_row["Prod.no"], trykktest_row["Beskrivelse"], int(lager), 1])
        else:
            rows.append(["", "Trykktest: Ja", int(lager), 1])

    rows.append(["1", "", int(lager), ""])

    # Apply antall_slanger multiplier
    if antall_slanger and antall_slanger != 1:
        for r in rows:
            core._multiply_row_quantity(r, antall_slanger)

    st.session_state.output_rows.extend(rows)

    # Store certificate data
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
    """Generate Excel file with all sheets"""
    output_wb = core.create_output_workbook(
        [[r[0], r[1], r[2], r[3]] for r in st.session_state.output_rows]
    )

    # Add Trykktest certificate sheets
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

    # Add Sluttkontroll sheet (always)
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

    # Save to buffer
    output_buffer = io.BytesIO()
    output_wb.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


# -------------------------------------------------
# MAIN UI
# -------------------------------------------------

st.title("🔧 Slangeprogram")

# Mode selection
col1, col2 = st.columns(2)
with col1:
    mode = st.radio("Innføringmodus", ["Rask innføring", "Full dialog"], 
                    index=0 if st.session_state.input_mode == "quick" else 1)
    st.session_state.input_mode = "quick" if mode == "Rask innføring" else "full"


# -------------------------------------------------
# QUICK MODE
# -------------------------------------------------

if st.session_state.input_mode == "quick":
    st.header("➕ Rask innføring")

    col1, col2 = st.columns([2, 1])

    with col1:
        first_line = st.text_input("Slangebeskrivelse", placeholder="Del1/Lengde/Del2/Del3[/Vinkel°]")

    with col2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="quick_material")

    col1, col2, col3 = st.columns(3)

    with col1:
        lager = st.selectbox("Lager",
                             options=["3", "1", "5"],
                             format_func=lambda x: {"3": "Lillestrøm", "1": "Ålesund", "5": "Trondheim"}[x],
                             key="quick_lager")

    with col2:
        antall_slanger = st.number_input("Antall slanger", min_value=1, value=1, key="quick_antall")

    with col3:
        type_approval = st.checkbox("Type Approval (DNV)?", key="quick_type_approval")

    # POS marking
    col1, col2 = st.columns([1, 2])
    with col1:
        pos_mark = st.checkbox("Merke med POS.nr?", key="quick_pos_mark")
    with col2:
        if pos_mark:
            posnr = st.text_input("POS.nr", value=str(st.session_state.pos_counter), key="quick_posnr")
        else:
            posnr = ""

    # Pressure test section
    st.divider()
    pressure_test = st.checkbox("Skal slangen trykkteststes?", key="quick_pressure_test")

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
            pressure_details["kundes_del_nr"] = st.text_input("Kundes del nr.", key="quick_del_nr")

    # Process button
    if st.button("✅ Legg til slange", use_container_width=True, key="quick_add_btn"):
        if not first_line:
            st.error("Slangebeskrivelse må oppgis!")
        else:
            # Find matches from summary
            selected_row, second_row1, second_row2, sheet_name_found, size_str, length_int = core.find_matches_from_summary(
                first_line, df1, df2_all, material_pref=material
            )

            process_and_add_hose(
                selected_row, second_row1, second_row2, sheet_name_found, size_str,
                length_int, material, lager, pos_mark, posnr, pressure_test,
                pressure_details, antall_slanger, first_line
            )

            st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")


# -------------------------------------------------
# FULL MODE
# -------------------------------------------------

else:
    st.header("📝 Full dialog")

    st.subheader("1️⃣ Velg slange")

    # Search hose
    search = st.text_input("Søk i slange beskrivelse", key="full_search")

    filtered_df = df1.copy()
    if search:
        filtered_df = filtered_df[
            filtered_df["Beskrivelse_2"]
            .astype(str)
            .str.contains(search, case=False, na=False)
        ]

    # Display table with selection
    st.write("**Velg slange fra tabellen under:**")
    event = st.dataframe(
        filtered_df[["Prod.no", "Beskrivelse_2", "Dimensjon", "Trykk(bar)"]],
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="hose_table"
    )

    # Check if row was selected
    if event.selection and event.selection["rows"]:
        selected_idx = event.selection["rows"][0]
        selected_prod_no = filtered_df.iloc[selected_idx]["Prod.no"]
        st.session_state.selected_hose_row = df1[df1["Prod.no"] == selected_prod_no].iloc[0]

    # Manual selection fallback
    if st.session_state.selected_hose_row is None:
        selected_prod_no = st.selectbox(
            "Eller velg slange (Prod.no)",
            options=filtered_df["Prod.no"].unique(),
            key="full_hose_select"
        )
        st.session_state.selected_hose_row = df1[df1["Prod.no"] == selected_prod_no].iloc[0]

    selected_row = st.session_state.selected_hose_row
    st.success(f"✅ Valgt: {selected_row['Beskrivelse_2']}")

    # Options
    col1, col2, col3 = st.columns(3)

    with col1:
        length = st.number_input("Lengde (mm)", value=1000, key="full_length")

    with col2:
        material = st.selectbox("Materiale", ["stål", "syrefast"], key="full_material")

    with col3:
        type_approval = st.checkbox("Type Approval (DNV)?", key="full_type_approval")

    size = str(selected_row["Dimensjon"]).zfill(2)

    # Determine sheet_name based on type approval and material
    if material == "syrefast":
        # Check for type approval with column L
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

    with col1:
        st.write("**Kupling 1**")
        st.write("Velg kupling fra tabellen:")
        event1 = st.dataframe(
            df2[["Prod.no", "Beskrivelse"]],
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="coupling1_table"
        )
        
        if event1.selection and event1.selection["rows"]:
            selected_idx1 = event1.selection["rows"][0]
            selected_c1_prod = df2.iloc[selected_idx1]["Prod.no"]
            st.session_state.selected_c1_row = df2[df2["Prod.no"] == selected_c1_prod].iloc[0]
        
        if st.session_state.selected_c1_row is not None:
            st.write(f"✅ Valgt: *{st.session_state.selected_c1_row['Beskrivelse']}*")
        else:
            st.info("Velg kupling fra tabellen")

    with col2:
        st.write("**Kupling 2**")
        st.write("Velg kupling fra tabellen:")
        event2 = st.dataframe(
            df2[["Prod.no", "Beskrivelse"]],
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="coupling2_table"
        )
        
        if event2.selection and event2.selection["rows"]:
            selected_idx2 = event2.selection["rows"][0]
            selected_c2_prod = df2.iloc[selected_idx2]["Prod.no"]
            st.session_state.selected_c2_row = df2[df2["Prod.no"] == selected_c2_prod].iloc[0]
        
        if st.session_state.selected_c2_row is not None:
            st.write(f"✅ Valgt: *{st.session_state.selected_c2_row['Beskrivelse']}*")
        else:
            st.info("Velg kupling fra tabellen")

    if st.session_state.selected_c1_row is None or st.session_state.selected_c2_row is None:
        st.warning("⚠️ Du må velge begge kuplinger")
        st.stop()

    row_c1 = st.session_state.selected_c1_row
    row_c2 = st.session_state.selected_c2_row

    # -------------------------------------------------
    # ADDITIONAL OPTIONS
    # -------------------------------------------------

    st.divider()
    st.subheader("3️⃣ Innstillinger")

    col1, col2, col3 = st.columns(3)

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

    # Pressure test
    st.divider()
    pressure_test = st.checkbox("Skal slangen trykkteststes?", key="full_pressure_test")

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
            pressure_details["kunde"] = st.text_input("Kunde", key="full_kunde")
            pressure_details["kundens_best_nr"] = st.text_input("Kundens best. Nr.", key="full_best_nr")
        with col2:
            pressure_details["hydra_ordre_nr"] = st.text_input("Hydra Pipe ordre nr.", key="full_hydra_ordre")
            pressure_details["kundes_del_nr"] = st.text_input("Kundes del nr.", key="full_del_nr")

    # Add to order
    if st.button("✅ Legg til slange", use_container_width=True, key="full_add_btn"):
        process_and_add_hose(
            selected_row, row_c1, row_c2, sheet_name, size,
            length, material, lager, pos_mark, posnr, pressure_test,
            pressure_details, antall_slanger
        )

        # Reset selections
        st.session_state.selected_hose_row = None
        st.session_state.selected_c1_row = None
        st.session_state.selected_c2_row = None

        st.success(f"✅ Slange lagt til! ({len(st.session_state.output_rows)} rader)")


# -------------------------------------------------
# ORDER PREVIEW (Common to both modes)
# -------------------------------------------------

st.divider()
st.header("📊 Nåværende utdata")

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
            st.rerun()

    with col3:
        if st.button("⬇️ Last ned Excel", use_container_width=True):
            excel_buffer = generate_excel()
            st.download_button(
                label="⬇️ Last ned Excel",
                data=excel_buffer,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
else:

    st.info("👈 Ingen slanger lagt til ennå. Velg innføringmodus og fyll inn feltene")

