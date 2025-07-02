import streamlit as st
import os
import shutil
import uuid
from datetime import datetime
import atexit
from concurrent.futures import ThreadPoolExecutor
from main import PowerPointExtractor, convert_ppt_to_pptx, run_captioning_threaded

# --- Constants ---
BASE_SESSION_DIR = "session_data"
CAPTIONED_FILE_KEY = "output_pptx_path"
CAPTION_DONE_KEY = "caption_done"
SESSION_ID_KEY = "session_id"

# --- Page Config ---
st.set_page_config(page_title="AutoCapPPT", layout="centered")
st.title("🧠 AutoCapPPT - AI-Powered Presentation Captioning")

# --- Cleanup Function ---
def cleanup():
    if SESSION_ID_KEY in st.session_state:
        session_dir = os.path.join(BASE_SESSION_DIR, st.session_state[SESSION_ID_KEY])
        if os.path.exists(session_dir):
            try:
                shutil.rmtree(session_dir)
                print(f"🧹 Deleted session directory: {session_dir}")
            except Exception as e:
                print(f"⚠️ Error while deleting session_data: {e}")
    st.session_state.clear()

# --- Register Cleanup on Exit ---
atexit.register(cleanup)

# --- Initialize Session State ---
if SESSION_ID_KEY not in st.session_state:
    st.session_state[SESSION_ID_KEY] = str(uuid.uuid4())
    print(f"🧾 New session ID: {st.session_state[SESSION_ID_KEY]}")
if CAPTION_DONE_KEY not in st.session_state:
    st.session_state[CAPTION_DONE_KEY] = False

# --- Session Directory Path ---
SESSION_DIR = os.path.join(BASE_SESSION_DIR, st.session_state[SESSION_ID_KEY])

# --- Download Page ---
if st.session_state[CAPTION_DONE_KEY]:
    st.success("✅ Your presentation has been captioned.")

    if CAPTIONED_FILE_KEY in st.session_state and os.path.exists(st.session_state[CAPTIONED_FILE_KEY]):
        with open(st.session_state[CAPTIONED_FILE_KEY], "rb") as f:
            st.download_button(
                label="⬇ Download Captioned PPTX",
                data=f,
                file_name=os.path.basename(st.session_state[CAPTIONED_FILE_KEY]),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("❌ Captioned file not found.")

    if st.button("🔄 Start Over"):
        cleanup()
        st.rerun()

# --- Upload Page ---
else:
    st.markdown("Upload a `.ppt` or `.pptx` file to automatically caption images using contextual slide information.")
    uploaded_file = st.file_uploader("📤 Upload PowerPoint", type=["ppt", "pptx"], key="uploaded_file")

    if uploaded_file:
        cleanup()
        os.makedirs(SESSION_DIR, exist_ok=True)

        # Save uploaded file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        input_filename = f"input_{timestamp}_{uploaded_file.name}"
        input_path = os.path.join(SESSION_DIR, input_filename)

        with open(input_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        if input_path.lower().endswith(".ppt"):
            convert_ppt_to_pptx(input_path)
            input_path = input_path.rsplit(".", 1)[0] + ".pptx"

        with st.spinner("⚙️ Generating image captions using AI... Please wait."):
            future = run_captioning_threaded(input_path, SESSION_DIR)
            output_pptx_path = future.result()

        if output_pptx_path and os.path.exists(output_pptx_path):
            st.session_state[CAPTION_DONE_KEY] = True
            st.session_state[CAPTIONED_FILE_KEY] = output_pptx_path
            st.rerun()
        else:
            st.error("❌ Failed to generate captioned presentation.")
