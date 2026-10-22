'''from google.adk.agents import Agent

root_agent = Agent(
    name="novaretail_agent",
    model="gemini-3.8-flash",
    description="NovaRetail quarterly business review agent",
    instruction="""
You are a business intelligent analyst for NovaRetail.
Analyze quarterly business performance using the available revenue, customer experience , and operational risk data

Identify:
-Overall business performance and quarter-on-quarter trends
-Regional strengths and weaknesses
-Customer satisfaction and delivery trends
-Operational risks and recommended actions

Use actualdata and calculations. Do not invent metrics.
Provide concise, executive- ready insights for a quarterly business review presentation
"""
)'''

from pathlib import Path
import pandas as pd
from google.adk.agents import Agent

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_FILE=BASE_DIR/"novaretail_dataset.xlsx"

def analyze_novaretail_data()-> dict:
    """Analyze actual quarterly performance from the Excel dataset."""

    if not DATA_FILE.exists():
        return{"error":f"Dataset not found:{DATA_FILE}"}

    try:
        workbook= pd.ExcelFile(DATA_FILE)

        required_sheets =[
            "Regional_Performance",
            "Customer_Experience",
            "Operational_Risk",
        ]

        missing =[
            sheet for sheet in required_sheets
            if sheet not in workbook.sheet_names

        ]

        if missing:
            return{"error":f"Missing worksheets:{missing}"}

        results= {}
        
        for sheet in required_sheets:
            df=pd.read_excel(DATA_FILE, sheet_name=sheet)

            #Preserve the actual column names and available data

            results[sheet]={"columns":df.columns.tolist(),
                            "row_count":len(df),
                            "quarterly_data":df.to_dict(orient="records"),
                            }
            return results
    except Exception as exc:
        return{"error":str(exc)}


root_agent = Agent(
    name="novaretail_agent",
    model="gemini-3.8-flash",
    description="NovaRetail quarterly business review agent",
    instruction="""
You are a business intelligent analyst for NovaRetail.
Analyze quarterly business performance using the available revenue, customer experience , and operational risk data

Identify:
-Overall business performance and quarter-on-quarter trends
-Regional strengths and weaknesses
-Customer satisfaction and delivery trends
-Operational risks and recommended actions

Use actualdata and calculations. Do not invent metrics.
Provide concise, executive- ready insights for a quarterly business review presentation
  """,
  tools=[analyze_novaretail_data] ,
) 

"""============================================================================"""




