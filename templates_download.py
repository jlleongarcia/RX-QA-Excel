import streamlit as st
import os

# --- Configuration ---
EXCEL_DIR = os.path.join(os.path.dirname(__file__), "Excel_templates") # Directory relative to the script location
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# --- Helper Function ---
def get_excel_files(directory):
    """Scans the directory for files with allowed Excel extensions."""
    files = []
    if os.path.isdir(directory):
        for item in os.listdir(directory):
            if os.path.isfile(os.path.join(directory, item)):
                # Check if the file extension is in the allowed set
                if os.path.splitext(item)[1].lower() in ALLOWED_EXTENSIONS:
                    files.append(item)
    return files

# --- Streamlit App ---
st.set_page_config(page_title="Excel Template Downloader", layout="centered")
st.title("📊 Excel Template Selector & Downloader")

st.write(f"Searching for Excel files in: `{EXCEL_DIR}`")

# Create the directory if it doesn't exist
if not os.path.exists(EXCEL_DIR):
    st.warning(f"Directory '{EXCEL_DIR}' not found. Creating it for you.")
    try:
        os.makedirs(EXCEL_DIR)
        st.success(f"Directory '{EXCEL_DIR}' created. Please add your Excel files there.")
    except OSError as e:
        st.error(f"Failed to create directory '{EXCEL_DIR}': {e}")
        st.stop() # Stop execution if directory creation fails

excel_files = get_excel_files(EXCEL_DIR)

if not excel_files:
    st.warning(f"No Excel files found in '{EXCEL_DIR}'. Please add some `.xlsx` or `.xls` files.")
else:
    st.success(f"Found {len(excel_files)} Excel file(s).")

    # --- File Selection Dropdown ---
    selected_file = st.selectbox(
        "Select an Excel file to download:",
        options=excel_files,
        index=None, # No default selection
        placeholder="Choose an option"
    )

    # --- Download Button ---
    if selected_file:
        file_path = os.path.join(EXCEL_DIR, selected_file)
        try:
            with open(file_path, "rb") as fp:
                # Determine the correct MIME type
                if selected_file.lower().endswith(".xlsx"):
                    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                elif selected_file.lower().endswith(".xls"):
                    mime_type = "application/vnd.ms-excel"
                else:
                    mime_type = "application/octet-stream" # Fallback

                st.download_button(
                    label=f"Download '{selected_file}'",
                    data=fp,
                    file_name=selected_file,
                    mime=mime_type,
                )

        except FileNotFoundError:
            st.error(f"Error: File '{selected_file}' not found at path '{file_path}'. Please check the directory.")
        except Exception as e:
            st.error(f"An error occurred while preparing the download: {e}")
