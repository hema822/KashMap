import streamlit as st

st.set_page_config(
    page_title="KashMap Migration",
    page_icon="🔄",
    layout="wide"
)

st.title("KashMap Migration Workspace")
st.write("Select the source and destination platform to start the migration.")

st.divider()

# Source selection
source = st.selectbox(
    "Select Source",
    [
        "Select Source",
        "WebFOCUS",
        "Cognos",
        "Tableau"
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

if st.button("Start Migration", type="primary"):

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

        # Other combinations
        else:
            st.warning(
                f"{source} → {destination} migration is not configured yet."
            )