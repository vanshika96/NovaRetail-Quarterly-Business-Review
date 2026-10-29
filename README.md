# NovaRetail Quarterly Business Review Automation

## Project Overview

This capstone project automates the creation of a five-slide executive PowerPoint presentation for NovaRetail, a multi-region consumer electronics retailer.

The Python script analyzes quarterly business data from an Excel workbook and updates a fixed PowerPoint template with:

- Calculated KPI values
- Insight-led slide headlines
- Executive summary content
- Regional performance charts
- Customer experience analysis
- Operational risk analysis
- Recommended business actions
- Relevant icons from the provided icon repository

The final presentation is generated programmatically using Python, Pandas, and python-pptx.

## Project Objective

The objective of this project is to demonstrate an end-to-end presentation automation workflow.

Instead of manually updating a quarterly business review deck, the solution:

1. Reads structured business data from Excel.
2. Calculates quarterly KPIs.
3. Identifies the latest and previous quarters.
4. Generates charts and business insights.
5. Locates named shapes in the PowerPoint template.
6. Replaces template placeholders with generated content.
7. Inserts relevant icons into designated positions.
8. Saves the completed executive presentation as a new PowerPoint file.

## Technology Used

- Python
- Pandas
- openpyxl
- python-pptx
- Microsoft Excel
- Microsoft PowerPoint
- Visual Studio Code
- Git
- GitHub

## Repository Structure

```text
NovaRetail_Project/
│
├── generate_deck.py
├── analyze_data.py
├── inspect_data.py
├── inspect_template.py
├── test_template_edit.py
├── novaretail_dataset.xlsx
├── novaretail_template.pptx
├── novaretail_quarterly_review.pptx
├── template_test_output.pptx
├── requirements.txt
├── README.md
├── .gitignore
│
└── novaretail_icon_repo/
    ├── customer_star.png
    ├── delivery_truck.png
    ├── globe_region.png
    ├── growth_chart.png
    ├── lightbulb_action.png
    ├── return_arrow.png
    ├── risk_alert.png
    └── shield_security.png
```

## Input Files

### 1. `novaretail_dataset.xlsx`

This Excel workbook contains the business data used to generate the presentation.

The workbook includes the following worksheets:

- `Regional_Performance`
- `Customer_Experience`
- `Operational_Risk`

The worksheets contain quarterly and regional business metrics such as:

- Revenue
- Units sold
- Marketing spend
- Return rate
- Net Promoter Score
- Customer satisfaction
- Delivery time
- Support tickets
- Stockout rate
- Return-processing time
- Warehouse capacity utilization
- Fraud incidents

### 2. `novaretail_template.pptx`

This is the fixed five-slide PowerPoint template used by the automation script.

The script identifies PowerPoint objects through their assigned shape names and places the generated content into the appropriate locations.

Static design elements such as backgrounds, logos, page numbers, footers, and separators are preserved.

### 3. `novaretail_icon_repo`

This folder contains the icons used in the presentation.

The script selects and places icons according to the topic of each slide.

## Generated Output

Running the script generates the final PowerPoint presentation:

```text
novaretail_quarterly_review.pptx
```

The final presentation contains the following five slides:

1. Executive Summary
2. Regional Performance
3. Customer Experience
4. Operational Risk
5. Recommended Actions and Next Steps

## Key Features

### Automated Data Analysis

The script reads the Excel workbook and calculates the business KPIs required for the executive presentation.

### Quarter Identification

Quarter values are arranged chronologically so that the latest quarter can be compared with the previous quarter.

### Dynamic KPI Calculation

The executive summary is populated with calculated business metrics, including:

- Revenue
- Quarter-over-quarter revenue movement
- Customer experience indicators
- Operational performance indicators

### Insight-Led Headlines

The presentation uses insight-led headlines instead of generic slide titles. This helps communicate the key business message of each slide.

### Automated Chart Generation

Charts are created using data from the Excel workbook and placed into the designated chart areas of the PowerPoint template.

### Named-Shape Processing

The script uses PowerPoint shape names to identify and update presentation elements instead of relying only on manually estimated coordinates.

### Automated Icon Placement

Relevant icons are selected from the icon repository and inserted into the corresponding placeholders.

### Template Preservation

The original five-slide structure and visual design are retained. Only the dynamic content areas are updated.

### Input Validation

The script validates:

- Required input files
- Expected Excel worksheets
- Empty datasets
- PowerPoint slide count
- Required template shapes
- Icon availability

## How to Run the Project

### Step 1: Clone the repository

```bash
git clone <your-github-repository-url>
cd NovaRetail_Project
```

Replace `<your-github-repository-url>` with the actual URL of the GitHub repository.

### Step 2: Create a virtual environment

```bash
python -m venv .venv
```

### Step 3: Activate the virtual environment

For Windows Command Prompt:

```bash
.venv\Scripts\activate
```

For Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

### Step 4: Install the required Python libraries

```bash
pip install -r requirements.txt
```

If a `requirements.txt` file is not available, install the required libraries directly:

```bash
pip install pandas openpyxl python-pptx
```

### Step 5: Run the presentation-generation script

```bash
python generate_deck.py
```

After successful execution, the completed PowerPoint presentation will be generated in the project folder.

## Python Requirements

The project requires:

- Python 3.10 or later
- pandas
- openpyxl
- python-pptx

The `requirements.txt` file contains:

```text
pandas
openpyxl
python-pptx
```

## Business Value

This project demonstrates how Python can automate a recurring business-reporting workflow.

The solution helps:

- Reduce manual effort in presentation creation
- Improve reporting consistency
- Minimize copy-and-paste errors
- Maintain a standardized presentation format
- Convert structured data into executive-ready insights
- Create a repeatable quarterly reporting process

The same approach can be adapted for:

- Quarterly business reviews
- Sales performance presentations
- Operational dashboards
- Customer experience reports
- Management reporting
- Executive committee presentations

## Important Notes

- Keep the Python script, dataset, PowerPoint template, and icon folder inside the main project directory.
- Do not rename the Excel worksheets unless the corresponding worksheet names are also updated in the Python script.
- Do not rename the PowerPoint template shapes unless the shape references are updated in the script.
- Close the generated PowerPoint file before running the Python script again.
- The generated PowerPoint file is included in the repository as part of the final submission.
- The `.venv` folder is excluded from GitHub through the `.gitignore` file.

## Project Workflow

```text
Excel Dataset
      ↓
Python Data Analysis
      ↓
KPI and Insight Generation
      ↓
PowerPoint Template Update
      ↓
Chart and Icon Placement
      ↓
Final Executive Presentation
```

## Author

**Vanshika Sood**

Capstone Project: AI-Assisted Executive PowerPoint Generation