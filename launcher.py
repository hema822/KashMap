import streamlit as st

st.set_page_config(
    page_title="KashMap Migration",
    page_icon="🔄",
    layout="wide"
)

st.title("KashMap Migration Workspace")
st.write("Select the source and destination platform to start the migration, or launch a direct workflow below.")

st.divider()

col_wf, col_cg = st.columns(2, gap="medium")

with col_wf:
    with st.container(border=True):
        st.subheader("WebFOCUS → Power BI")
        st.markdown(
            """
            **Input format:** WebFOCUS procedure files (`.fex`) or ZIP folder  
            **Optional:** Metadata ZIP (`.mas` / `.acx`) & Resource Analyzer files  
            **Capabilities:** FEX logic extraction, DEFINE/COMPUTE mapping, Hold file lineage, duplicate analysis, and Power BI migration plan.
            """
        )
        if st.button("Launch WebFOCUS → Power BI 🚀", key="quick_wf_pbi", type="primary", use_container_width=True):
            st.switch_page("pages/webfocus_powerbi.py")

with col_cg:
    with st.container(border=True):
        st.subheader("Cognos → Power BI")
        st.markdown(
            """
            **Input format:** IBM Cognos Report specifications (`.xml`)  
            **Optional:** Data Module JSON & Framework Manager model XML  
            **Capabilities:** Deterministic DAX & Power Query M translator, Lists & Crosstabs layout extraction, extended prompt/macro tokenization, and multi-sheet Excel export.
            """
        )
        if st.button("Launch Cognos → Power BI 🚀", key="quick_cg_pbi", type="primary", use_container_width=True):
            st.switch_page("pages/cognos_powerbi.py")

st.divider()
st.subheader("All Migration Routes")

# Source selection
source = st.selectbox(
    "Select Source",
    [
        "Select Source",
        "WebFOCUS",
        "Cognos",
        "Tableau",
        "SSRS"
    ]
)

# Destination selection
destination = st.selectbox(
    "Select Destination",
    [
        "Select Destination",
        "Cognos",
        "Power BI",
        "Pyramid Analytics",
        "Tableau"
    ]
)

st.divider()

if st.button("Start Migration", type="secondary", use_container_width=True):

    if source == "Select Source":
        st.warning("Please select a source platform.")

    elif destination == "Select Destination":
        st.warning("Please select a destination platform.")

    else:
        # WebFOCUS → Cognos
        if source == "WebFOCUS" and destination == "Cognos":
            st.switch_page("pages/webfocus_cognos.py")

        # WebFOCUS → Power BI
        elif source == "WebFOCUS" and destination == "Power BI":
            st.switch_page("pages/webfocus_powerbi.py")

        # WebFOCUS → Pyramid Analytics
        elif source == "WebFOCUS" and destination == "Pyramid Analytics":
            st.switch_page("pages/webfocus_pyramid.py")

        # WebFOCUS → Tableau
        elif source == "WebFOCUS" and destination == "Tableau":
            st.switch_page("pages/webfocus_tableau.py")

        # Cognos → Power BI
        elif source == "Cognos" and destination == "Power BI":
            st.switch_page("pages/cognos_powerbi.py")

        # Tableau → Power BI
        elif source == "Tableau" and destination == "Power BI":
            st.switch_page("pages/tableau_powerbi.py")

        # SSRS → Pyramid Analytics
        elif source == "SSRS" and destination == "Pyramid Analytics":
            st.switch_page("pages/ssrs_pyramid.py")

        # Other combinations
        else:
            st.warning(
                f"{source} → {destination} migration is not configured yet."
            )