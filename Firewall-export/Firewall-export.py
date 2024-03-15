import json
import subprocess
import pandas as pd

def run_oci_command(command):
    subprocess.run(command, shell=True)

def fetch_output_json_files(network_firewall_policy_id):
    print("Fetching output JSON files...")
    oci_commands = [
        "oci network-firewall security-rule list --network-firewall-policy-id {network_firewall_policy_id} --all > security_rule.json",
        "oci network-firewall address-list list --network-firewall-policy-id {network_firewall_policy_id} --all > addresslist.json",
        "oci network-firewall service list --network-firewall-policy-id {network_firewall_policy_id} --all > service.json",
        "oci network-firewall service-list list --network-firewall-policy-id {network_firewall_policy_id} --all > servicelist.json",
        "oci network-firewall application list --network-firewall-policy-id {network_firewall_policy_id} --all > application.json",
        "oci network-firewall application-group list --network-firewall-policy-id {network_firewall_policy_id} --all > applicationlist.json"
    ]
    
    for i, command in enumerate(oci_commands):
        oci_commands[i] = command.format(network_firewall_policy_id="--id " + network_firewall_policy_id)
    
    for command in oci_commands:
        run_oci_command(command)
    
    print("Output JSON files fetched successfully.")

def json_to_excel(input_files, output_file):
    with pd.ExcelWriter(output_file) as writer:  # Create Excel writer
        for input_file in input_files:
            with open(input_file, 'r') as f:
                data = json.load(f)
            
            df = pd.json_normalize(data)  # Convert JSON to DataFrame
            sheet_name = input_file.split('.')[0]  # Use file name as sheet name
            
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Write DataFrame to Excel

def convert_json_to_excel():
    print("Converting JSON files to Excel...")
    input_files = ['security_rule_output.json', 'addresslist_output.json', 
                   'servicelist_output.json', 'service_output.json', 
                   'application_output.json', 'applicationlist_output.json']
    output_file = 'output.xlsx'
    json_to_excel(input_files, output_file)
    print("Conversion to Excel completed. Excel file saved as output.xlsx.")

def main():
    print("Welcome to the OCI Network Firewall Tool!")
    print("1. Fetch Output JSON Files")
    print("2. Convert JSON Files to Excel")
    choice = input("Enter your choice (1 or 2): ")
    
    if choice == '1':
        network_firewall_policy_id = input("Enter Network Firewall Policy ID: ")
        fetch_output_json_files(network_firewall_policy_id)
    elif choice == '2':
        convert_json_to_excel()
    else:
        print("Invalid choice. Please enter either '1' or '2'.")

if __name__ == "__main__":
    main()
