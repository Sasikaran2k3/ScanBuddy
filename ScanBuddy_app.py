import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
import re
import os
import io
from typing import List, Any

# Page configuration
st.set_page_config(
    page_title="ScanBuddy: PDF-Excel Comparator",
    page_icon="📊",
    layout="wide"
)

# Initialize session state
if 'current_page' not in st.session_state:
    st.session_state.current_page = 1
if 'pdf_uploaded' not in st.session_state:
    st.session_state.pdf_uploaded = False
if 'excel_uploaded' not in st.session_state:
    st.session_state.excel_uploaded = False
if 'pdf_pages' not in st.session_state:
    st.session_state.pdf_pages = 0
if 'excel_columns' not in st.session_state:
    st.session_state.excel_columns = 0

# Comparison conditions - easily extensible
COMPARISON_CONDITIONS = {
    "/YYYY (4-digit year)": lambda text: bool(re.search(r"/\d{4}", text)),
    # Add more conditions here in the future
    # "Email pattern": lambda text: bool(re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)),
    # "Phone pattern": lambda text: bool(re.search(r'\b\d{3}-\d{3}-\d{4}\b', text)),
}

def save_uploaded_file(uploaded_file, filename: str) -> bool:
    """Save uploaded file with consistent naming"""
    try:
        with open(filename, "wb") as f:
            f.write(uploaded_file.getbuffer())
        return True
    except Exception as e:
        st.error(f"Error saving file: {str(e)}")
        return False

def get_pdf_page_count(pdf_path: str) -> int:
    """Get total number of pages in PDF"""
    try:
        doc = fitz.open(pdf_path)
        page_count = len(doc)
        doc.close()
        return page_count
    except Exception as e:
        st.error(f"Error reading PDF: {str(e)}")
        return 0

def get_excel_column_count(excel_path: str) -> int:
    """Get total number of columns in Excel"""
    try:
        df = pd.read_excel(excel_path, nrows=1)
        return len(df.columns)
    except Exception as e:
        st.error(f"Error reading Excel: {str(e)}")
        return 0

def create_short_pdf(start_page: int, end_page: int) -> bool:
    """Create short PDF from selected page range"""
    try:
        src = fitz.open("raw.pdf")
        short = fitz.open()
        
        for i in range(start_page - 1, end_page):  # Convert to 0-based indexing
            if i < len(src):
                short.insert_pdf(src, from_page=i, to_page=i)
        
        short.save("short.pdf")
        src.close()
        short.close()
        return True
    except Exception as e:
        st.error(f"Error creating short PDF: {str(e)}")
        return False

def extract_and_filter_pdf_text(condition_func) -> List[str]:
    """Extract text from PDF and filter based on condition"""
    try:
        doc = fitz.open("short.pdf")
        pdf_filtered_lines = []
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            text = page.get_text()
            lines = text.split('\n')
            
            for line in lines:
                line = line.strip()
                if line and condition_func(line):
                    pdf_filtered_lines.append(line)
        
        doc.close()
        return pdf_filtered_lines
    except Exception as e:
        st.error(f"Error extracting PDF text: {str(e)}")
        return []

def read_excel_column(col_index: int) -> List[str]:
    """Read specific column from Excel file"""
    try:
        df = pd.read_excel("raw.xlsx")
        if col_index >= len(df.columns):
            st.error(f"Column {col_index + 1} does not exist in the Excel file")
            return []
        
        excel_values = df.iloc[:, col_index].dropna().astype(str).tolist()
        return excel_values
    except Exception as e:
        st.error(f"Error reading Excel column: {str(e)}")
        return []

def find_matches(pdf_lines: List[str], excel_values: List[str]) -> List[str]:
    """Find matches between PDF lines and Excel values"""
    matches = []
    # Write excel_values and pdf_lines to txt files
    try:
        with open("checkExcel.txt", "w", encoding="utf-8") as f_excel:
            for val in excel_values:
                f_excel.write(f"{val}\n")
        with open("checkPDF.txt", "w", encoding="utf-8") as f_pdf:
            for line in pdf_lines:
                f_pdf.write(f"{line}\n")
    except Exception as e:
        st.error(f"Error writing check files: {str(e)}")
    for val in excel_values:
        if any(val in line for line in pdf_lines):
            matches.append(val)
    return matches

def create_output_file(matches: List[str]) -> bool:
    """Create Excel file with matched results"""
    try:
        output_df = pd.DataFrame(matches, columns=["Matched Rows"])
        output_df.to_excel("matched_output.xlsx", index=False)
        return True
    except Exception as e:
        st.error(f"Error creating output file: {str(e)}")
        return False

# PAGE 1: Welcome and File Upload
def page_1_welcome():
    st.title("🔍 ScanBuddy: PDF-Excel Comparator")
    st.markdown("### Upload your files to get started")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload PDF File")
        pdf_file = st.file_uploader("Choose PDF file", type=['pdf'], key="pdf_uploader")
        
        if pdf_file is not None:
            if save_uploaded_file(pdf_file, "raw.pdf"):
                st.session_state.pdf_uploaded = True
                st.session_state.pdf_pages = get_pdf_page_count("raw.pdf")
                st.success(f"✅ PDF uploaded successfully! ({st.session_state.pdf_pages} pages)")
            else:
                st.session_state.pdf_uploaded = False
    
    with col2:
        st.subheader("📊 Upload Excel File")
        excel_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key="excel_uploader")
        
        if excel_file is not None:
            if save_uploaded_file(excel_file, "raw.xlsx"):
                st.session_state.excel_uploaded = True
                st.session_state.excel_columns = get_excel_column_count("raw.xlsx")
                st.success(f"✅ Excel uploaded successfully! ({st.session_state.excel_columns} columns)")
            else:
                st.session_state.excel_uploaded = False
    
    # Next button
    if st.session_state.pdf_uploaded and st.session_state.excel_uploaded:
        if st.button("Next ➡️", type="primary"):
            st.session_state.current_page = 2
            st.rerun()
    else:
        st.info("📋 Please upload both PDF and Excel files to continue")

# PAGE 2: Page Range & Column Selection
def page_2_selection():
    st.title("📋 Select Page Range and Column")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 PDF Page Range")
        st.info(f"Total pages in PDF: {st.session_state.pdf_pages}")
        
        start_page = st.number_input(
            "Start Page", 
            min_value=1, 
            max_value=st.session_state.pdf_pages,
            value=1,
            key="start_page_input"
        )
        
        end_page = st.number_input(
            "End Page", 
            min_value=start_page, 
            max_value=st.session_state.pdf_pages,
            value=min(start_page, st.session_state.pdf_pages),
            key="end_page_input"
        )
    
    with col2:
        st.subheader("📊 Excel Column")
        st.info(f"Total columns in Excel: {st.session_state.excel_columns}")
        
        column_number = st.number_input(
            "Column Number", 
            min_value=1, 
            max_value=st.session_state.excel_columns,
            value=1,
            key="column_number_input"
        )
    
    # Navigation buttons
    col_back, col_next = st.columns(2)
    
    with col_back:
        if st.button("⬅️ Back", type="secondary"):
            st.session_state.current_page = 1
            st.rerun()
    
    with col_next:
        if st.button("Next ➡️", type="primary"):
            st.session_state.start_page = start_page
            st.session_state.end_page = end_page
            st.session_state.column_number = column_number
            st.session_state.current_page = 3
            st.rerun()

# PAGE 3: Comparison Condition
def page_3_condition():
    st.title("⚙️ Compare Condition")
    
    st.subheader("Select Comparison Condition")
    condition_name = st.selectbox(
        "Choose condition to filter PDF text:",
        list(COMPARISON_CONDITIONS.keys()),
        key="condition_input"
    )
    
    st.info(f"Selected condition: **{condition_name}**")
    
    # Navigation buttons
    col_back, col_next = st.columns(2)
    
    with col_back:
        if st.button("⬅️ Back", type="secondary"):
            st.session_state.current_page = 2
            st.rerun()
    
    with col_next:
        if st.button("Start Comparison ➡️", type="primary"):
            st.session_state.condition_name = condition_name
            st.session_state.current_page = 4
            st.rerun()

# PAGE 4: Results
def page_4_results():
    st.title("📊 Comparison Results")
    
    # Show processing status
    with st.spinner("Processing files..."):
        # Create short PDF
        if not create_short_pdf(st.session_state.start_page, st.session_state.end_page):
            st.error("Failed to create short PDF")
            return
        
        # Get condition function
        condition_func = COMPARISON_CONDITIONS[st.session_state.condition_name]
        
        # Extract and filter PDF text
        pdf_filtered_lines = extract_and_filter_pdf_text(condition_func)
        
        # Read Excel column
        excel_values = read_excel_column(st.session_state.column_number - 1)  # Convert to 0-based
        
        # Find matches
        matches = find_matches(pdf_filtered_lines, excel_values)
        
        # Create output file
        create_output_file(matches)
    
    # Display results
    st.success("✅ Processing completed!")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.metric("Number of matching rows", len(matches))
        
    with col2:
        st.metric("Total PDF lines filtered", len(pdf_filtered_lines))
    
    # Show sample matches
    if matches:
        st.subheader("📋 Sample Matches (First 10)")
        sample_matches = matches[:10]
        for i, match in enumerate(sample_matches, 1):
            st.write(f"{i}. {match}")
        
        if len(matches) > 10:
            st.info(f"... and {len(matches) - 10} more matches")
    else:
        st.warning("No matches found")
    
    # Download button
    if matches:
        try:
            with open("matched_output.xlsx", "rb") as file:
                st.download_button(
                    label="📥 Download Results",
                    data=file,
                    file_name="matched_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
        except Exception as e:
            st.error(f"Error preparing download: {str(e)}")
    
    # Navigation button
    if st.button("🔄 Start New Comparison", type="secondary"):
        # Reset session state
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.session_state.current_page = 1
        st.rerun()

# Main App Logic
def main():
    # Sidebar for navigation
    with st.sidebar:
        st.header("📍 Navigation")
        page_names = [
            "1️⃣ Upload Files",
            "2️⃣ Select Range & Column", 
            "3️⃣ Choose Condition",
            "4️⃣ View Results"
        ]
        
        current_page_name = page_names[st.session_state.current_page - 1]
        st.write(f"Current: **{current_page_name}**")
        
        st.markdown("---")
        st.markdown("### 📋 Status")
        st.write(f"PDF: {'✅' if st.session_state.pdf_uploaded else '❌'}")
        st.write(f"Excel: {'✅' if st.session_state.excel_uploaded else '❌'}")
    
    # Route to appropriate page
    if st.session_state.current_page == 1:
        page_1_welcome()
    elif st.session_state.current_page == 2:
        page_2_selection()
    elif st.session_state.current_page == 3:
        page_3_condition()
    elif st.session_state.current_page == 4:
        page_4_results()

if __name__ == "__main__":
    main()

# Auto-run command (uncomment to use)
# import os
# os.system("streamlit run ScanBuddy_app.py")