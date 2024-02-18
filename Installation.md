**Installation instructions:**
1. You should have Python 3 installed on your system.
Refer this link for installing python on your workstation.
2. Required packages: The code requires several Python packages to run. Make sure you have them installed by running the following command in your terminal:
        pip install pandas, xlrd
3. To convert excel file, to Json file using the script, follow the steps below:
4. Open the terminal or command prompt on your Mac or Windows computer.
5. Navigate to the directory where the “Firewall-Import-v6.py” script is located using the “cd” command.
6. Download and move input.xlsx file to the same folder as the script. Then in your cmd line / terminal , navigate to that folder using the “cd” command.
7. Type “python3 Firewall-Import-v6.py -i input.xlsx” in the terminal or command prompt and press Enter.
8. Once the script runs successfully, you must find json files (iplist.json, services.json, servicelist.json, application.json, applicationlist.json, securityrules.json) in the location the script was run.

**Troubleshooting:**
If you encounter any errors during the execution of the script, ensure that the input Excel file is correctly formatted and contains the required sheets.
Check for any missing dependencies for python.
