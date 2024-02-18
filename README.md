# OCI Network Firewall Policy Configuration Tool

## Description
The OCI Network Firewall Policy Configuration Tool is a Python script designed to simplify the process of creating and managing OCI (Oracle Cloud Infrastructure) Network Firewall policies. It automates the conversion of firewall policy configurations stored in Excel sheets into JSON format, which can be easily imported into OCI Network Firewall policies.

## Code Explanation
The tool consists of several functions:

1. **write_json_to_file(data, filename):**
   - This function writes JSON data to a file.
   - It takes JSON data and a filename as input and writes the data to the specified file in JSON format.

2. **excel_to_json_iplist(excel_file):**
   - Converts the 'iplist' sheet in the Excel file to `iplist.json`.
   - Reads IP address data from the 'iplist' sheet, constructs JSON objects for each IP address, and writes them to `iplist.json`.

3. **excel_to_json_service_list(excel_file):**
   - Converts the 'service' sheet in the Excel file to `services.json`, `servicelist.json`, `application.json`, and `applicationlist.json`.
   - Reads service data from the 'service' sheet, constructs JSON objects for different service types, and writes them to respective JSON files.

4. **replace_with_names(df_security_rules, df_iplist):**
   - Replaces IP addresses in security rule configurations with corresponding names from the 'iplist' sheet.
   - Reads security rule configurations and IP address data, replaces IP addresses with names, and returns the updated DataFrame.

5. **convert_to_json(df_security_rules):**
   - Converts security rule configurations to JSON format.
   - Reads security rule configurations from a DataFrame, constructs JSON objects for each rule, and returns a list of JSON objects.

6. **excel_to_json(excel_file):**
   - Converts the 'security-rules' sheet in the Excel file to `securityrules.json`.
   - Reads security rule configurations and IP address data, replaces IP addresses with names, converts configurations to JSON, and writes them to `securityrules.json`.

7. **main():**
   - Entry point of the script.
   - Parses command-line arguments, calls appropriate functions based on input, and handles exceptions.

## Usage
1. Clone the repository and navigate to the project directory.
2. Install dependencies using `pip install -r Requirements.txt`.
3. Run the script using `python3 Firewall-import.py -i input.xlsx`, where `input_file.xlsx` is the path to your input Excel file containing firewall policy configurations.

## Configuration
- Ensure that your input Excel file follows the specified format with sheets named 'iplist', 'service', and 'security-rules'.
- Customize the Excel sheets according to your network firewall policy configurations.

## Troubleshooting
- If you encounter any errors during execution, ensure that the input Excel file is correctly formatted and contains the required sheets.
- Check for any missing dependencies by running `pip install -r requirements.txt`.

## Contributing
Contributions to the project are welcome. If you encounter any bugs or have suggestions for improvements, please open an issue on GitHub or submit a pull request with your changes.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Credits
- This tool was developed by Karthik Mani.
- Special thanks to the contributors and users for their feedback and support.

## References
- [OCI Network Firewall Documentation](https://docs.oracle.com/en-us/iaas/Content/Network/Concepts/networksecurity.htm)
