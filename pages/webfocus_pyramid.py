import sys
from pathlib import Path
import runpy
import streamlit as st

try:
    st.set_page_config(
        page_title="WebFOCUS to Pyramid Analytics - KashMap",
        page_icon="🔄",
        layout="wide"
    )
except Exception:
    pass

ROOT_DIR = Path(__file__).resolve().parent.parent
TARGET_MODULE_DIR = ROOT_DIR / "migrations" / "webfocus_pyramid"
TARGET_APP = TARGET_MODULE_DIR / "app.py"

if str(TARGET_MODULE_DIR) not in sys.path:
    sys.path.insert(0, str(TARGET_MODULE_DIR))

if TARGET_APP.exists():
    runpy.run_path(str(TARGET_APP), run_name="__main__")
else:
    st.title("WebFOCUS to Pyramid Analytics")
    st.warning(f"Migration script not found at: {TARGET_APP}")
