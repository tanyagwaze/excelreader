import os
import pandas as pd

def summarize_excel_files(input_folder, output_file):
    summary_data = []
    categories = ["Adobe Sign", "Carahsoft", "DocuSign", "DigiCert", "High Emprise"]
    
    
    summary_dict = {category: [0, 0, 0, 0, 0] for category in categories}
    
    
    for file in os.listdir(input_folder):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            file_path = os.path.join(input_folder, file)
            try:
                xls = pd.ExcelFile(file_path)
                
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name)
                    
                    for i, category in enumerate(categories):
                        summary_dict[category][i] += df.iloc[:, i].sum() if i < df.shape[1] else 0
            except Exception as e:
                print(f"Error processing {file}: {e}")
    
    
    for category, totals in summary_dict.items():
        summary_data.append([category, "", *totals, sum(totals)])
    
    
    summary_df = pd.DataFrame(summary_data, columns=["Category", "Criteria 1", "Criteria 2", "Criteria 3", "Criteria 4", "Total"])
    
    
    summary_df.to_excel(output_file, index=False)
    print(f"Summary saved to {output_file}")


input_folder = "path_to_excel_files"  
output_file = "summary.xlsx"
summarize_excel_files(input_folder, output_file)
