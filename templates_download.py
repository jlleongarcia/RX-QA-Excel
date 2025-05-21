import streamlit as st
import os

# --- Configuration ---
EXCEL_DIR = os.path.join(os.path.dirname(__file__), "Excel_templates") # Directory relative to the script location
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# --- Helper Function ---
def get_excel_files_categorized(directory):
    """
    Scans the directory and its immediate subdirectories for files with allowed Excel extensions.
    Returns a dictionary where keys are category names (subfolder names or "Root")
    and values are lists of file names within that category.
    """
    categorized_files = {}
    if not os.path.isdir(directory):
        return categorized_files # Return empty if directory doesn't exist

    # First, check for files directly in the root directory
    root_files = []
    try:
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            if os.path.isfile(item_path):
                if os.path.splitext(item)[1].lower() in ALLOWED_EXTENSIONS:
                    root_files.append(item)
    except OSError as e:
        st.error(f"Error listing files in root directory '{directory}': {e}")

    if root_files:
        categorized_files["Root"] = root_files # Use "Root" for files not in subfolders

    # Then, check subdirectories
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        if os.path.isdir(item_path):
            # Recursively call or list files in this subdirectory
            category_files = []
            try:
                for sub_item in os.listdir(item_path):
                    sub_item_path = os.path.join(item_path, sub_item)
                    if os.path.isfile(sub_item_path):
                        if os.path.splitext(sub_item)[1].lower() in ALLOWED_EXTENSIONS:
                            category_files.append(sub_item)
            except OSError as e:
                 st.warning(f"Could not list files in directory '{item_path}': {e}")
                 continue # Skip this directory if listing fails

            if category_files:
                categorized_files[item] = category_files # Use subfolder name as category
    return categorized_files

# --- Streamlit App ---
st.set_page_config(page_title="Excel Template Downloader", layout="centered")
st.title("📊 Excel Template Selector & Downloader")

# st.write(f"Searching for Excel files in: `{EXCEL_DIR}`")

# Create the directory if it doesn't exist
if not os.path.exists(EXCEL_DIR):
    st.warning(f"Directory '{EXCEL_DIR}' not found. Creating it for you.")
    try:
        os.makedirs(EXCEL_DIR)
        st.success(f"Directory '{EXCEL_DIR}' created. Please add your Excel files there.")
    except OSError as e:
        st.error(f"Failed to create directory '{EXCEL_DIR}': {e}")
        st.stop() # Stop execution if directory creation fails

categorized_excel_files = get_excel_files_categorized(EXCEL_DIR)

total_files = sum(len(files) for files in categorized_excel_files.values())

st.write("Browse and download files by category:")

# Sort categories alphabetically, putting "Root" first if it exists
sorted_categories = sorted(categorized_excel_files.keys())
if "Root" in sorted_categories:
    sorted_categories.remove("Root")
    sorted_categories.insert(0, "Root")

# Display files and download buttons for each category within an expander
for category in sorted_categories:
    files_in_category = categorized_excel_files[category]

    # Determine the label for the expander
    if category == "Root":
        expander_label = f"**Files in Root Directory** ({len(files_in_category)} file(s))"
    else:
        expander_label = f"**{category}** ({len(files_in_category)} file(s))"

    with st.expander(expander_label, expanded=False): # Set expanded=True if you want them open by default
        if not files_in_category:
            st.caption("No files in this category.")
            continue

        for file_name in files_in_category:
            # Construct the full relative path for the file
            if category == "Root":
                file_relative_path = file_name
            else:
                file_relative_path = os.path.join(category, file_name)

            file_full_path = os.path.join(EXCEL_DIR, file_relative_path)

            try:
                with open(file_full_path, "rb") as fp:
                    # Determine the correct MIME type
                    if file_name.lower().endswith(".xlsx"):
                        mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    elif file_name.lower().endswith(".xls"):
                        mime_type = "application/vnd.ms-excel"
                    else:
                        mime_type = "application/octet-stream" # Fallback

                    # Use columns to place filename and button side-by-side
                    col1, col2 = st.columns([0.7, 0.3]) # Adjust column width ratio as needed

                    with col1:
                        st.write(f"- {file_name}") # Display filename

                    with col2:
                        st.download_button(
                            label="Download", # Label for the button
                            data=fp,
                            file_name=file_name, # The downloaded file name will just be the filename, not the full path
                            mime=mime_type,
                            key=f"download_{category}_{file_name}" # Unique key for each button
                        )

            except FileNotFoundError:
                st.error(f"Error: File '{file_relative_path}' not found.")
            except Exception as e:
                st.error(f"An error occurred while preparing '{file_relative_path}' for download: {e}")
