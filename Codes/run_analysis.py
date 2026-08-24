import os
import json
import subprocess
import sys

def prompt_with_default(prompt, default):
    user_input = input(f"{prompt} (Default: {default}): ").strip()
    if not user_input:
        return default
    return user_input

def main():
    print("=========================================")
    print("   TATA STEEL COMPLIANCE CONTROL CENTER")
    print("=========================================\n")
    print("This tool will configure the filters for your analysis and run the entire pipeline.\n")

    plant_input = prompt_with_default("Enter Plant Code(s) separated by comma", "742")
    plant_codes = [p.strip() for p in plant_input.split(",")]

    product_cat = prompt_with_default("Enter Product Category", "CRCA")

    # Load mapping to automatically suggest P Codes
    suggested_pcodes = "C01, C06"
    try:
        with open("product_mapping.json", "r") as f:
            mapping = json.load(f)
            codes = [k for k, v in mapping.items() if str(v).upper() == product_cat.upper()]
            if codes:
                suggested_pcodes = ", ".join(codes)
    except Exception:
        pass

    pcode_input = prompt_with_default(f"Enter P Codes for {product_cat} (separated by comma)", suggested_pcodes)
    p_codes = [p.strip() for p in pcode_input.split(",")]

    prod_file = prompt_with_default("Enter Production Data Excel file", "Production_Data.XLSX")

    config = {
        "plant_codes": plant_codes,
        "product_category": product_cat,
        "p_codes": p_codes,
        "production_file": prod_file
    }

    print("\nSaving configuration...")
    with open("config.json", "w") as f:
        json.dump(config, f, indent=4)
        
    print(json.dumps(config, indent=4))
    print("\n--- BEGINNING PIPELINE ---\n")
    
    scripts = [
        "clean_production.py",
        "clean_commitments.py",
        "merge_master.py"
    ]
    
    for script in scripts:
        print(f"\n>> Running {script}...")
        result = subprocess.run([sys.executable, script])
        if result.returncode != 0:
            print(f"\nCRITICAL ERROR: {script} failed! Pipeline stopped.")
            sys.exit(1)
            
    print("\n=========================================")
    print("PIPELINE COMPLETE! OPEN Master_Compliance_Report.xlsx")
    print("=========================================")

if __name__ == "__main__":
    main()
