import os
import subprocess
import xlwings as xw
import time

# =============================================================================
# 1. CONFIGURATION & SCRIPT DEFINITIONS
# =============================================================================

# Define the name for our new project and its files
PROJECT_NAME = "exitpath_interactive_model"
PYTHON_FILENAME = f"{PROJECT_NAME}.py"
EXCEL_FILENAME = f"{PROJECT_NAME}.xlsm"

# Define the default assumptions to populate the Excel sheet
DEFAULT_ASSUMPTIONS = {
    # Funnel & CAC
    "Operator_CAC": 1000,
    "Investor_CAC": 200,
    "MF_Churn_Rate": 0.25,
    "Conversion_MF_CF": 0.75,
    "Conversion_CF_Ready": 0.3,
    "Conversion_Ready_Go": 0.1,

    # Pricing
    "Price_MF": 0,
    "Price_CF": 0,
    "Price_Ready": 50000,
    "Go_Deal_Size": 75000000,
    "Go_Fee": 0.015,

    # Payroll & Hiring
    "Salary_SDR": 8000,
    "Salary_CS": 10000,
    "Salary_Eng": 12000,
    "Salary_AE": 12000,
    "Salary_GA": 10000,
    "Customers_Per_CS": 20,
    "Customers_Per_Eng": 40,

    # Marketing & Funding
    "Marketing_Start": 10000,
    "Marketing_End": 30000,
    "Funding_Months": "7, 18",
    "Funding_Amounts": "750000, 1250000",
    "Collection_Upfront": 0.7,

    # COGS %
    "COGS_Pct": 0.15
}

# This is the full Python code that will be written into the project's .py file
# It's the interactive model logic that will be called by the Excel button
INTERACTIVE_MODEL_CODE = """
import numpy as np
import pandas as pd
import xlwings as xw

def main():
    try:
        # ============================================
        # 1. CONNECT TO THE WORKBOOK & READ INPUTS
        # ============================================
        wb = xw.Book.caller()
        ws_inputs = wb.sheets['Inputs']
        ws_inputs.range('F4').value = "Running..." # Status update

        assumptions = ws_inputs.range('A1').expand().options(pd.DataFrame, header=1).value.set_index('Variable')['Value'].to_dict()

        # Convert comma-separated string inputs to lists of numbers
        assumptions["Funding_Months"] = [int(x.strip()) for x in str(assumptions["Funding_Months"]).split(',')]
        assumptions["Funding_Amounts"] = [float(x.strip()) for x in str(assumptions["Funding_Amounts"]).split(',')]

        # ============================================
        # 2. CORE LOGIC (Your proven model)
        # ============================================
        MONTHS = 24
        QUARTERS_OUTYEARS = 12
        TOTAL_PERIODS = MONTHS + QUARTERS_OUTYEARS
        PERIOD_LABELS = [f"M{i+1}" for i in range(MONTHS)] + [f"Q{i+1}" for i in range(QUARTERS_OUTYEARS)]

        marketing_spend = np.concatenate([np.linspace(assumptions["Marketing_Start"], assumptions["Marketing_End"], MONTHS), np.repeat(assumptions["Marketing_End"], QUARTERS_OUTYEARS)])
        funding = np.zeros(TOTAL_PERIODS)
        for i, m in enumerate(assumptions["Funding_Months"]):
            if m - 1 < TOTAL_PERIODS:
                funding[m-1] = assumptions["Funding_Amounts"][i]

        operator_free = marketing_spend / assumptions["Operator_CAC"]
        operator_mf = np.zeros(TOTAL_PERIODS); operator_cf = np.zeros(TOTAL_PERIODS); operator_ready = np.zeros(TOTAL_PERIODS); operator_go = np.zeros(TOTAL_PERIODS)

        for t in range(1, TOTAL_PERIODS):
            operator_mf[t] = operator_free[t-1] * (1 - assumptions["MF_Churn_Rate"])
            operator_cf[t] = operator_mf[t-1] * assumptions["Conversion_MF_CF"]
            operator_ready[t] = operator_cf[t-1] * assumptions["Conversion_CF_Ready"]
            operator_go[t] = operator_ready[t-1] * assumptions["Conversion_Ready_Go"]

        arr_ready = operator_ready * assumptions["Price_Ready"]
        go_revenue = operator_go * assumptions["Go_Deal_Size"] * assumptions["Go_Fee"]
        total_revenue = arr_ready + go_revenue

        customers = operator_ready
        cs_headcount = np.ceil(customers / assumptions["Customers_Per_CS"]); eng_headcount = np.ceil(customers / assumptions["Customers_Per_Eng"]); sdr_headcount = np.maximum(1, np.ceil(customers / 10)); ae_headcount = np.where(customers >= 20, 1, 0); ga_headcount = np.ones(TOTAL_PERIODS) * 0.5
        total_headcount = cs_headcount + eng_headcount + sdr_headcount + ae_headcount + ga_headcount
        payroll = (cs_headcount * assumptions["Salary_CS"] + eng_headcount * assumptions["Salary_Eng"] + sdr_headcount * assumptions["Salary_SDR"] + ae_headcount * assumptions["Salary_AE"] + ga_headcount * assumptions["Salary_GA"])

        cogs = total_revenue * assumptions["COGS_Pct"]
        opex = payroll + marketing_spend
        gross_margin = total_revenue - cogs
        ebitda = gross_margin - opex

        collections = total_revenue * assumptions["Collection_Upfront"]
        cash = np.zeros(TOTAL_PERIODS)
        for t in range(TOTAL_PERIODS):
            net_cash_flow = (collections[t] + funding[t]) - (opex[t] + cogs[t])
            cash[t] = (cash[t-1] if t > 0 else 0) + net_cash_flow
        
        # ============================================
        # 3. CREATE DATA FRAMES & WRITE TO EXCEL
        # ============================================
        df_rev = pd.DataFrame({"Period": PERIOD_LABELS, "Free": operator_free, "MF": operator_mf, "CF": operator_cf, "Ready": operator_ready, "Go": operator_go, "ARR_Ready": arr_ready, "Go_Revenue": go_revenue, "Total_Revenue": total_revenue}).set_index("Period")
        df_hc = pd.DataFrame({"Period": PERIOD_LABELS, "CS": cs_headcount, "Eng": eng_headcount, "SDR": sdr_headcount, "AE": ae_headcount, "G&A": ga_headcount, "Total_Headcount": total_headcount, "Payroll": payroll}).set_index("Period")
        df_stmt = pd.DataFrame({"Period": PERIOD_LABELS, "Revenue": total_revenue, "COGS": cogs, "Gross_Margin": gross_margin, "Opex": opex, "EBITDA": ebitda, "Funding": funding, "Ending_Cash": cash}).set_index("Period")

        for ws_name, df in [('Revenue_Cohorts', df_rev), ('Headcount_Payroll', df_hc), ('3_Statement', df_stmt)]:
            if ws_name not in [s.name for s in wb.sheets]: wb.sheets.add(ws_name)
            ws = wb.sheets[ws_name]
            ws.clear()
            ws.range('A1').value = df
            ws.autofit()

        ws_inputs.range('F5').value = cash[-1] # Write out a key KPI
        ws_inputs.range('F4').value = "Success!"
        
    except Exception as e:
        # If an error occurs, write it to a cell in Excel for easy debugging
        wb = xw.Book.caller()
        wb.sheets['Inputs'].range('F4').value = f"ERROR: {e}"

if __name__ == "__main__":
    # This allows you to test the script from a Python environment
    # You must have the Excel file open to run this.
    xw.Book(f'{xw.App().books[0].name}').set_mock_caller()
    main()
"""


# =============================================================================
# 2. MAIN EXECUTION BLOCK
# =============================================================================

def create_model():
    """Main function to generate the entire interactive model project."""
    print(f"--- Creating project: {PROJECT_NAME} ---")

    # Step 1: Run 'xlwings quickstart' to generate the project structure
    # This creates the .py file and the .xlsm file with the VBA module embedded.
    try:
        # Use --standalone to embed all VBA code, making the Excel file portable
        subprocess.run(["xlwings", "quickstart", PROJECT_NAME, "--standalone"], check=True, capture_output=True, text=True)
        print(f"[SUCCESS] Created project folder and files for '{PROJECT_NAME}'.")
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] Could not run 'xlwings quickstart'. Is xlwings installed?")
        print(f"Details: {e.stderr}")
        return
    except FileNotFoundError:
        print("[ERROR] 'xlwings' command not found. Is it installed and in your system's PATH?")
        return

    # Step 2: Overwrite the boilerplate Python file with our real model logic
    project_path = os.path.join(os.getcwd(), PROJECT_NAME)
    python_file_path = os.path.join(project_path, PYTHON_FILENAME)
    try:
        with open(python_file_path, "w") as f:
            f.write(INTERACTIVE_MODEL_CODE)
        print(f"[SUCCESS] Wrote interactive model logic to '{PYTHON_FILENAME}'.")
    except IOError as e:
        print(f"[ERROR] Could not write to Python file: {e}")
        return

    # Step 3: Programmatically configure the Excel workbook
    excel_file_path = os.path.join(project_path, EXCEL_FILENAME)
    print(f"--- Configuring Excel Interface: {EXCEL_FILENAME} ---")
    
    # Use a visible app instance to ensure everything loads correctly
    with xw.App(visible=False) as app:
        try:
            wb = app.books.open(excel_file_path)
            
            # Prepare the 'Inputs' sheet
            ws = wb.sheets[0]
            ws.name = "Inputs"
            ws.clear() # Clear any boilerplate content

            # Write headers
            ws.range('A1').value = ["Variable", "Value", "Description"]
            ws.range('A1:C1').font.bold = True
            ws.range('A1:C1').color = (20, 55, 90) # Dark blue
            ws.range('A1:C1').font.color = (255, 255, 255) # White

            # Write the assumptions
            row = 2
            for key, value in DEFAULT_ASSUMPTIONS.items():
                ws.range(f'A{row}').value = key
                ws.range(f'B{row}').value = value
                row += 1

            # Add a results section
            ws.range('E1').value = "Key Metrics"
            ws.range('E1').font.bold = True
            ws.range('E2').value = [["Status:", ""], ["Ending Cash:", ""]]

            # Format the sheet
            ws.autofit()
            ws.range('B:B').number_format = '#,##0.00'
            ws.range('B1').clear() # Clear number format from header
            
            # The button is already created by quickstart, just save and close
            print("[SUCCESS] Configured 'Inputs' sheet.")
            
            wb.save()
            time.sleep(1) # Give Excel a moment to save
            wb.close()

        except Exception as e:
            print(f"[ERROR] Failed while configuring the Excel file: {e}")
        finally:
            # Ensure the Excel process is closed
            if app.pid:
                app.quit()

    print("\n--- SETUP COMPLETE ---")
    print(f"Your interactive pro forma model is ready.")
    print(f"1. Navigate to the '{PROJECT_NAME}' folder.")
    print(f"2. Open '{EXCEL_FILENAME}'.")
    print("3. Change any assumption in column B and click the 'Run main' button.")


if __name__ == "__main__":
    create_model()
