import tempfile
from pathlib import Path

import streamlit as st
from frame_trace import render_traced_tables, render_formula_table
import roi_excel_calculator_v2_dynamic_premises_fixed_v8_scalar as calc

from roi_excel_calculator_v2_dynamic_premises_fixed_v8_scalar import (
    create_excel_download,
    process_file,
)
from frame_chat import render_chat_panel


# =========================================================
# PAGE CONFIGURATION
# =========================================================

st.set_page_config(
    page_title="Frame ROI Calculator",
    page_icon="📊",
    layout="wide",
)


# =========================================================
# HEADER
# =========================================================

st.title("Frame Services Calculator")

st.write(
    "Upload the Services Project Excel file. "
    "The application calculates Current Services and Future Services "
    "row by row using the X/Y/Z-specific cost and value assumptions from the Premises tab."
)

# st.info(
#     "No ROI percentage is calculated. "
#     "The output contains FTE, cost, value, yearly margin contribution, "
#     "and monthly margin contribution."
# )


# =========================================================
# INPUT
# =========================================================

uploaded_file = st.file_uploader(
    "Upload Excel file",
    type=["xlsx", "xls"],
    help="Upload the Excel workbook containing Services Project and Premises tabs.",
)


# =========================================================
# PROCESS
# =========================================================

if uploaded_file is not None:

    st.success(f"Uploaded: {uploaded_file.name}")

    upload_key = f"{uploaded_file.name}:{uploaded_file.size}"
    if st.session_state.get("_planned_roi_upload_key") != upload_key:
        st.session_state["_planned_roi_upload_key"] = upload_key
        st.session_state["planned_roi_submitted"] = False
        st.session_state["actual_roi_submitted"] = False
        st.session_state["roi_inputs_submitted"] = False
        st.session_state["roi_more_details"] = False
        for key in ("planned_talent_months", "planned_ai_months", "actual_talent_hours", "actual_ai_hours"):
            st.session_state.pop(key, None)

    with tempfile.NamedTemporaryFile(
        delete=False,
        suffix=Path(uploaded_file.name).suffix,
    ) as temp_file:
        temp_file.write(uploaded_file.getbuffer())
        input_path = temp_file.name

    try:
        with st.spinner("Calculating Current and Future Services..."):
            (
                current_df,
                future_df,
                combined_df,
                summary_df,
                future_schedule_df,
                future_current_margin_df,
            ) = process_file(input_path)

        # -----------------------------------------------------
        # SUMMARY
        # -----------------------------------------------------

        # st.subheader("Calculation Summary")

        # if not summary_df.empty:
        #     st.dataframe(
        #         summary_df,
        #         use_container_width=True,
        #         hide_index=True,
        #     )

        # -----------------------------------------------------
        # RESULTS WITH HOVER FORMULAS AND DEPENDENCY HIGHLIGHTING
        # -----------------------------------------------------
        premises = calc.read_premises_values(input_path)

        # -----------------------------------------------------
        # INPUTS: Planned and Actual, submitted together
        # -----------------------------------------------------
        premises = calc.read_premises_values(input_path)
        service_rows = future_df[
            future_df["Service Improvement"].astype(str).str.strip().str.casefold() != "total"
        ].reset_index(drop=True)
        if "roi_inputs_submitted" not in st.session_state:
            st.session_state["roi_inputs_submitted"] = False
        if "roi_more_details" not in st.session_state:
            st.session_state["roi_more_details"] = False

        if not st.session_state["roi_inputs_submitted"]:
            st.subheader("Project Cost Inputs")
            st.write("Enter the Planned completion months and Actual hours for each Future Services row, then submit once.")
            with st.form("planned_actual_project_cost_inputs"):
                with st.expander("Planned · Project Cost Worksheet and ROI Model — inputs", expanded=False):
                    planned_month_inputs = {}
                    for i, sr in service_rows.iterrows():
                        name = str(sr.get("Service Improvement", f"Row {i+1}"))
                        st.markdown(f"**{name}**")
                        c1,c2 = st.columns(2)
                        with c1:
                            tm = st.number_input(f"Talent months to complete — {name}", min_value=0.0, value=0.0, step=1.0, key=f"planned_talent_{i}")
                        with c2:
                            am = st.number_input(f"AI months to complete — {name}", min_value=0.0, value=0.0, step=1.0, key=f"planned_ai_{i}")
                        planned_month_inputs[name] = {"talent":float(tm),"ai":float(am)}
                with st.expander("Actual · Project Cost Worksheet and ROI Model — inputs", expanded=False):
                    actual_hours_inputs = {}
                    for i, sr in service_rows.iterrows():
                        name = str(sr.get("Service Improvement", f"Row {i+1}"))
                        st.markdown(f"**{name}**")
                        c1,c2 = st.columns(2)
                        with c1:
                            th = st.number_input(f"Talent hours from cards + todos — {name}", min_value=0.0, value=0.0, step=1.0, key=f"actual_talent_{i}")
                        with c2:
                            ah = st.number_input(f"AI hours from cards + todos — {name}", min_value=0.0, value=0.0, step=1.0, key=f"actual_ai_{i}")
                        actual_hours_inputs[name] = {"talent":float(th),"ai":float(ah)}
                submitted = st.form_submit_button("Submit Planned and Actual Inputs", type="primary")
            if submitted:
                st.session_state["planned_talent_months"] = {n:v["talent"] for n,v in planned_month_inputs.items()}
                st.session_state["planned_ai_months"] = {n:v["ai"] for n,v in planned_month_inputs.items()}
                st.session_state["actual_talent_hours"] = {n:v["talent"] for n,v in actual_hours_inputs.items()}
                st.session_state["actual_ai_hours"] = {n:v["ai"] for n,v in actual_hours_inputs.items()}
                st.session_state["planned_roi_submitted"] = True
                st.session_state["actual_roi_submitted"] = True
                st.session_state["roi_inputs_submitted"] = True
                st.session_state["roi_more_details"] = False
                st.rerun()

        planned_project_cost_df = None
        actual_project_cost_df = None
        comparison_df = None
        if st.session_state["roi_inputs_submitted"]:
            planned_project_cost_df = calc.build_planned_project_cost_worksheet(
                future_df=future_df, premises_values=premises,
                talent_months_to_complete=st.session_state.get("planned_talent_months", {}),
                ai_months_to_complete=st.session_state.get("planned_ai_months", {}),
            )
            actual_project_cost_df = calc.build_actual_project_cost_worksheet(
                future_df=future_df, premises_values=premises,
                talent_hours_by_service=st.session_state.get("actual_talent_hours", {}),
                ai_hours_by_service=st.session_state.get("actual_ai_hours", {}),
            )
            comparison_df = calc.build_planned_vs_actual_comparison(planned_project_cost_df, actual_project_cost_df)
            if comparison_df is not None and not comparison_df.empty:
                st.markdown("## Planned vs Actual")
                def comparison_formula(row_idx, row, col, value):
                    if col in ("Talent months — Actual","AI months — Actual"):
                        return ("Corresponding Actual hours from cards + todos ÷ 160", [])
                    if col.endswith(" — Δ"):
                        base = col.rsplit(" — ",1)[0]
                        return (f"{base}: Actual − Planned", [f"{base} — Planned",f"{base} — Actual"])
                    if col.endswith(" — Planned"):
                        return ("Value taken from Planned Project Cost Worksheet and ROI Model", [])
                    if col.endswith(" — Actual"):
                        return ("Value taken from Actual Project Cost Worksheet and ROI Model", [])
                    return None
                render_formula_table(comparison_df,"Planned vs Actual",comparison_formula)
                if not st.session_state["roi_more_details"]:
                    _, mid, _ = st.columns([2,1,2])
                    with mid:
                        if st.button("More details", type="primary", use_container_width=True, key="more_details_button"):
                            st.session_state["roi_more_details"] = True
                            st.rerun()

                if st.session_state["roi_more_details"]:
                    # st.markdown("## Detailed Results")
                    st.markdown("### Planned · Project Cost Worksheet and ROI Model")
                    def planned_formula(row_idx,row,col,value):
                        deps = {
                            "Talent cost (K$/month)":["Talent Tier"], "AI cost (K$/month)":["AI Tier"],
                            "Project non-recurring cost (K$)":["Talent months to complete","AI months to complete","Talent Tier","AI Tier","Talent cost (K$/month)","AI cost (K$/month)"],
                            "Project monthly recurring cost (K$)":["Talent cost (K$/month)","AI cost (K$/month)"],
                            "ROI in 12 months":["Talent+AI monthly margin (K$)","Project non-recurring cost (K$)","Project monthly recurring cost (K$)"],
                            "Months to breakeven":["Project monthly recurring cost (K$)","Project non-recurring cost (K$)","Talent+AI monthly margin (K$)"]}
                        formulas = {
                            "Talent cost (K$/month)":"Current Talent FTE Cost for selected X/Y/Z tier ÷ 12",
                            "AI cost (K$/month)":"Current AI FTE Cost for selected X/Y/Z tier ÷ 12",
                            "External NRE incl. AI (K$)":"Not entered; placeholder '-'",
                            "Project non-recurring cost (K$)":f"({row.get('Talent months to complete',0)} × {row.get('Talent cost (K$/month)',0)}) + ({row.get('AI months to complete',0)} × {row.get('AI cost (K$/month)',0)})",
                            "Project monthly recurring cost (K$)":"Talent cost (K$/month) + AI cost (K$/month)",
                            "Talent+AI monthly margin (K$)":"Matched Future Services row's Talent+AI Monthly Margin Contribution (K$)",
                            "ROI in 12 months":"Talent+AI monthly margin × 12 ÷ (project non-recurring cost + project monthly recurring cost × 12)",
                            "Months to breakeven":"(project monthly recurring cost × 12 + project non-recurring cost) ÷ Talent+AI monthly margin"}
                        return (formulas[col],deps.get(col,[])) if col in formulas else None
                    render_formula_table(planned_project_cost_df,"Planned · Project Cost Worksheet and ROI Model",planned_formula,show_legend=False)

                    st.markdown("### Actual · Project Cost Worksheet and ROI Model")
                    def actual_formula(row_idx,row,col,value):
                        deps = {
                            "Talent cost (K$/month)":["Talent Tier"], "AI cost (K$/month)":["AI Tier"],
                            "Project non-recurring cost (K$)":["Talent hours from cards + todos","AI hours from cards + todos","Talent Tier","AI Tier","Talent cost (K$/month)","AI cost (K$/month)"],
                            "Project monthly recurring cost (K$)":["Talent cost (K$/month)","AI cost (K$/month)"],
                            "ROI in 12 months":["Talent+AI monthly margin (K$)","Project non-recurring cost (K$)","Project monthly recurring cost (K$)"],
                            "Months to breakeven":["Project monthly recurring cost (K$)","Project non-recurring cost (K$)","Talent+AI monthly margin (K$)"]}
                        formulas = {
                            "Talent cost (K$/month)":"Current Talent FTE Cost for selected X/Y/Z tier ÷ 12",
                            "AI cost (K$/month)":"Current AI FTE Cost for selected X/Y/Z tier ÷ 12",
                            "External NRE incl. AI (K$)":"Not entered; placeholder '-'",
                            "Project non-recurring cost (K$)":f"({row.get('Talent hours from cards + todos',0)} ÷ 160 × {row.get('Talent cost (K$/month)',0)}) + ({row.get('AI hours from cards + todos',0)} ÷ 160 × {row.get('AI cost (K$/month)',0)})",
                            "Project monthly recurring cost (K$)":"Talent cost (K$/month) + AI cost (K$/month)",
                            "Talent+AI monthly margin (K$)":"Matched Future Services row's Talent+AI Monthly Margin Contribution (K$)",
                            "ROI in 12 months":"Talent+AI monthly margin × 12 ÷ (project non-recurring cost + project monthly recurring cost × 12)",
                            "Months to breakeven":"(project monthly recurring cost × 12 + project non-recurring cost) ÷ Talent+AI monthly margin"}
                        return (formulas[col],deps.get(col,[])) if col in formulas else None
                    render_formula_table(actual_project_cost_df,"Actual · Project Cost Worksheet and ROI Model",actual_formula,show_legend=False)

                    render_traced_tables(current_df=current_df,future_df=future_df,schedule_df=future_schedule_df,margin_df=future_current_margin_df,premises=premises,show_legend=False)
                    _, bottom_mid, _ = st.columns([2,1,2])
                    with bottom_mid:
                        if st.button("Hide details", use_container_width=True, key="hide_details_bottom"):
                            st.session_state["roi_more_details"] = False
                            st.rerun()
        else:
            planned_project_cost_df = None
            actual_project_cost_df = None

        if st.session_state.get("roi_more_details", False):
            # -----------------------------------------------------
            # DOWNLOAD
            # -----------------------------------------------------

            excel_data = create_excel_download(
                current_df=current_df,
                future_df=future_df,
                combined_df=combined_df,
                summary_df=summary_df,
                future_schedule_df=future_schedule_df,
                future_current_margin_df=future_current_margin_df,
                planned_project_cost_df=planned_project_cost_df,
            )

            st.download_button(
                label="Download Calculated Excel",
                data=excel_data,
                file_name="frame_roi_calculated_output.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
            )

            st.caption(
                "The downloaded workbook includes the calculated service tables and, after submission, "
                "the Planned Project Cost Worksheet and ROI Model."
            )

        if st.session_state.get("roi_more_details", False):
            # -----------------------------------------------------
            # CHAT WITH RESULTS
            # -----------------------------------------------------

            render_chat_panel(
                input_path=input_path,
                file_key=f"{uploaded_file.name}:{uploaded_file.size}",
                current_df=current_df,
                future_df=future_df,
                schedule_df=future_schedule_df,
                margin_df=future_current_margin_df,
            )

    except Exception as exc:
        st.error(f"Error processing Excel file: {exc}")

    finally:
        try:
            Path(input_path).unlink(missing_ok=True)
        except Exception:
            pass