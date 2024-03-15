import pandas as pd
import json

def json_to_excel(input_files, output_file):
    with pd.ExcelWriter(output_file) as writer:  # Create Excel writer
        for input_file in input_files:
            with open(input_file, 'r') as f:
                data = json.load(f)
            
            df = pd.json_normalize(data)  # Convert JSON to DataFrame
            sheet_name = input_file.split('.')[0]  # Use file name as sheet name
            
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Write DataFrame to Excel

def main():
    input_files = ['security_rule_output.json', 'addresslist_output.json', 
                   'servicelist_output.json', 'service_output.json', 
                   'application_output.json', 'applicationlist_output.json']
    output_file = 'output.xlsx'
    
    json_to_excel(input_files, output_file)

if __name__ == "__main__":
    main()
