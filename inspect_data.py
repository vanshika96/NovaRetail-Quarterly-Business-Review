import pandas as pd

file_path= "novaretail_dataset.xlsx"

#Read the excel workbook
excel = pd.ExcelFile(file_path)
print("Available Sheets",excel.sheet_names)

#Inspect every sheet
for sheet in excel.sheet_names:
    df=pd.read_excel(file_path,sheet_name=sheet)

    print(f"\n{'='*50}")
    print(f"\n{sheet}")
    print(f"{'='*50}")

    print("\nColumn names:")
    print(df.columns.tolist())

    print("\nFirst 5 rows:")
    print(df.head().to_string(index = False))

    print("\nDataset size:",df.shape)

    print("\nMissingvalues:")
    print(df.isnull().sum().to_string())