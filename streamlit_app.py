import streamlit as st
import os
import sys
import json
import shutil
import tempfile

st.set_page_config(page_title="Tata Compliance Generator", layout="wide")

st.title("Tata Compliance Generator")
st.write("Upload your files below to generate the compliance report and dashboard. This works on any device (Windows, Mac, Phone).")

# 1. Inputs
profile = st.selectbox("Select Profile", ["742 CRCA", "743&744 PPGL"])
prod_file = st.file_uploader("1. Upload Production Data File (Excel)", type=["xlsx", "xls"])
comm_files = st.file_uploader("2. Upload Monthly Commitment Files", type=["xlsx", "xls"], accept_multiple_files=True)

if st.button("Generate Dashboard", type="primary"):
    if not prod_file:
        st.error("Please upload the Production Data file.")
    elif not comm_files:
        st.error("Please upload at least one Monthly Commitment file.")
    else:
        original_cwd = os.getcwd()
        with st.spinner("Analyzing and Generating Dashboard... Please wait."):
            try:
                # Create a temp directory
                temp_dir = tempfile.mkdtemp()
                
                # Save Production file
                prod_filename = "742.XLSX" if profile == "742 CRCA" else "743&744.XLSX"
                prod_path = os.path.join(temp_dir, prod_filename)
                with open(prod_path, "wb") as f:
                    f.write(prod_file.getbuffer())
                    
                # Save Commitment files
                comm_dir = os.path.join(temp_dir, "monthly_commitments")
                os.makedirs(comm_dir, exist_ok=True)
                for c_file in comm_files:
                    with open(os.path.join(comm_dir, c_file.name), "wb") as f:
                        f.write(c_file.getbuffer())
                        
                # Dynamically load ALL P Codes from product_mapping.json
                mapping_path = os.path.join(original_cwd, "Codes", "product_mapping.json")
                with open(mapping_path, "r") as f:
                    pcode_map = json.load(f)
                    
                if profile == "742 CRCA":
                    # For CRCA, find all P Codes that map to 'CRCA'
                    crca_codes = [k for k, v in pcode_map.items() if str(v).strip().upper() == "CRCA"]
                    config = {
                        "profile_name": profile,
                        "production_file": prod_filename,
                        "plant": ["742"],
                        "plant_codes": ["742"],
                        "product_category": "CRCA",
                        "p_codes": crca_codes
                    }
                else:
                    # For PPGL/GL, find all P Codes that map to PPGL, GL, GP, PPGI, Galv
                    target_cats = ["PPGL", "GL", "GP", "PPGI", "GALV.", "GALV.-GC", "GALV", "PPGI-EMBO"]
                    ppgl_codes = [k for k, v in pcode_map.items() if str(v).strip().upper() in target_cats]
                    config = {
                        "profile_name": profile,
                        "production_file": prod_filename,
                        "plant": ["743", "744", 743, 744],
                        "plant_codes": ["743", "744"],
                        "product_category": "PPGL,GL,GI,PPGI",
                        "p_codes": ppgl_codes
                    }
                    
                with open(os.path.join(temp_dir, "config.json"), "w") as f:
                    json.dump(config, f, indent=4)
                    
                # Run the pipeline just like app.py does
                codes_dir = os.path.join(original_cwd, "Codes")
                os.chdir(temp_dir)
                
                # Copy product mapping
                mapping_path = os.path.join(codes_dir, "product_mapping.json")
                if os.path.exists(mapping_path):
                    shutil.copy2(mapping_path, temp_dir)
                
                # Execute scripts dynamically
                scripts = ["clean_production.py", "clean_commitments.py", "merge_master.py", "generate_dashboard.py"]
                for script in scripts:
                    script_path = os.path.join(codes_dir, script)
                    with open(script_path, "r", encoding="utf-8") as s:
                        code = s.read()
                        # Execute in fresh namespace
                        exec(code, {"__name__": "__main__"})
                
                # Verify intermediate outputs to provide better error messages
                if not os.path.exists("Cleaned_Production_Data.xlsx"):
                    raise Exception("Failed during Phase 1: Production data could not be processed. Please check if the uploaded production file is correct.")
                if not os.path.exists("Cleaned_Commitments_Data.xlsx"):
                    raise Exception("Failed during Phase 2: Commitment data could not be processed. Please check if the uploaded commitment files are correct Excel files.")
                
                # Verify final outputs
                if not os.path.exists("Master_Compliance_Report.xlsx"):
                    raise Exception("Failed during Phase 3: Pipeline finished but output Excel was not created.")
                
                # Read outputs into memory for download
                with open("Master_Compliance_Report.xlsx", "rb") as f:
                    excel_data = f.read()
                    
                html_data = ""
                if os.path.exists("Compliance_Dashboard.html"):
                    with open("Compliance_Dashboard.html", "r", encoding="utf-8") as f:
                        html_data = f.read()
                        
                # Cleanup
                os.chdir(original_cwd)
                shutil.rmtree(temp_dir, ignore_errors=True)
                
                st.success("Successfully generated!")
                
                # Provide download button
                st.download_button(
                    label="Download Master Compliance Report (Excel)",
                    data=excel_data,
                    file_name="Master_Compliance_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Display Dashboard inline
                if html_data:
                    st.components.v1.html(html_data, height=900, scrolling=True)

            except Exception as e:
                import traceback
                os.chdir(original_cwd)
                st.error(f"An error occurred:\\n{str(e)}")
                with st.expander("Show details"):
                    st.code(traceback.format_exc())
