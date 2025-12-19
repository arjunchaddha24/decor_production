import streamlit as st
import pandas as pd
import os
import sys
from io import StringIO
import traceback
from datetime import datetime

# Import all the production functions
from production_functions import (
    get_style_numbers_from_plan,
    get_row_wise_data_from_plan,
    get_row_wise_data_from_daily_prod,
    convert_cumulative_to_daywise_quantities_for_daily_prod,
    match_plan_with_actual,
    delete_empty_rows,
    add_cumulative_columns_to_matched_dict,
    write_production_report_to_excel,
    do_everything
)

# Page configuration
st.set_page_config(
    page_title="Production Report Generator",
    page_icon="📊",
    layout="centered"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        font-size: 1.1rem;
        padding: 0.75rem;
        border-radius: 5px;
        border: none;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #155a8a;
    }
    </style>
""", unsafe_allow_html=True)

# Header
st.markdown('<div class="main-header">📊 Production Report Generator</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Upload your plan and daily production files to generate a comprehensive report</div>', unsafe_allow_html=True)

st.markdown("---")

# File upload section
st.markdown("### 📁 Upload Files")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**Plan File**")
    plan_file = st.file_uploader(
        "Choose plan Excel file",
        type=['xlsx'],
        key="plan",
        help="Upload the production plan Excel file"
    )
    if plan_file:
        st.success(f"✅ {plan_file.name}")

with col2:
    st.markdown("**Daily Production File**")
    daily_file = st.file_uploader(
        "Choose daily production Excel file",
        type=['xlsx'],
        key="daily",
        help="Upload the daily production Excel file"
    )
    if daily_file:
        st.success(f"✅ {daily_file.name}")

st.markdown("---")

# Generate report button
if st.button("🚀 Generate Report", disabled=(not plan_file or not daily_file)):
    
    # Create temporary directory for processing
    temp_dir = "/tmp/production_report"
    os.makedirs(temp_dir, exist_ok=True)
    
    # File paths
    plan_path = os.path.join(temp_dir, "plan.xlsx")
    daily_path = os.path.join(temp_dir, "daily.xlsx")
    output_path = os.path.join(temp_dir, "production_report.xlsx")
    
    try:
        # Save uploaded files
        with st.spinner("📤 Uploading files..."):
            with open(plan_path, "wb") as f:
                f.write(plan_file.getbuffer())
            with open(daily_path, "wb") as f:
                f.write(daily_file.getbuffer())
        
        st.success("✅ Files uploaded successfully")
        
        # Validate files
        with st.spinner("🔍 Validating files..."):
            try:
                # Try to read the files to check if they're valid Excel files
                pd.ExcelFile(plan_path)
                pd.ExcelFile(daily_path)
                st.success("✅ Files validated successfully")
            except Exception as e:
                st.error(f"❌ Invalid Excel file(s): {str(e)}")
                st.stop()
        
        # Process files
        st.markdown("### 📊 Processing Report")
        
        # Capture console output
        output_capture = StringIO()
        
        # Redirect stdout to capture print statements
        old_stdout = sys.stdout
        sys.stdout = output_capture
        
        try:
            # Run the main processing function
            with st.spinner("⚙️ Generating report... This may take a few moments."):
                do_everything(plan_path, daily_path, output_path)
            
            # Restore stdout
            sys.stdout = old_stdout
            
            # Get captured output
            console_output = output_capture.getvalue()
            
            # Display ALL console output without filtering
            if console_output.strip():
                st.markdown("### 📋 Processing Log")
                
                # Split into lines and display each one
                lines = console_output.split('\n')
                
                for line in lines:
                    line = line.strip()
                    if not line:
                        continue
                    
                    # Display every line as an info message to preserve all details
                    st.info(line)
                
                st.markdown("---")
            
            # Check if output file was created
            if os.path.exists(output_path):
                st.success("✅ Report generated successfully!")
                
                # Read the file for download
                with open(output_path, "rb") as f:
                    excel_data = f.read()
                
                # Generate filename with timestamp
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                download_filename = f"production_report_{timestamp}.xlsx"
                
                # Download button
                st.markdown("### 📥 Download Report")
                st.download_button(
                    label="⬇️ Download Excel Report",
                    data=excel_data,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Show file info
                file_size = len(excel_data) / 1024  # KB
                st.info(f"📄 File size: {file_size:.1f} KB")
                
            else:
                st.error("❌ Report generation failed. Please check the errors and warnings above.")
        
        except Exception as e:
            # Restore stdout
            sys.stdout = old_stdout
            
            # Get captured output
            console_output = output_capture.getvalue()
            
            # Display error
            st.error(f"❌ An error occurred during processing:")
            st.code(str(e))
            
            # Show traceback in expander
            with st.expander("🔍 Technical Details"):
                st.code(traceback.format_exc())
            
            # Show console output if any
            if console_output.strip():
                st.markdown("### Console Output")
                st.text(console_output)
    
    except Exception as e:
        st.error(f"❌ An unexpected error occurred:")
        st.code(str(e))
        with st.expander("🔍 Technical Details"):
            st.code(traceback.format_exc())

else:
    if not plan_file or not daily_file:
        st.info("👆 Please upload both files to continue")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9rem;'>
    <p>Production Report Generator v1.0</p>
</div>
""", unsafe_allow_html=True)
