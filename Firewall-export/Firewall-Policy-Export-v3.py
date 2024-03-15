import json
import subprocess
import argparse

def run_oci_command(command):
    subprocess.run(command, shell=True)

def generate_json_files(network_firewall_policy_id):
    # Define OCI CLI commands to generate JSON files
    oci_commands = [
        "oci network-firewall security-rule list --network-firewall-policy-id {network_firewall_policy_id} --all > security_rule.json",
        "oci network-firewall address-list list --network-firewall-policy-id {network_firewall_policy_id} --all > addresslist.json",
        "oci network-firewall service list --network-firewall-policy-id {network_firewall_policy_id} --all > service.json",
        "oci network-firewall service-list list --network-firewall-policy-id {network_firewall_policy_id} --all > servicelist.json",
        "oci network-firewall application list --network-firewall-policy-id {network_firewall_policy_id} --all > application.json",
        "oci network-firewall application-group list --network-firewall-policy-id {network_firewall_policy_id} --all > applicationlist.json"
    ]

    # Replace {network_firewall_policy_id} with actual value
    for i, command in enumerate(oci_commands):
        oci_commands[i] = command.format(network_firewall_policy_id=network_firewall_policy_id)

    # Run OCI CLI commands to generate JSON files
    for command in oci_commands:
        run_oci_command(command)

def export_items(input_file, output_file, command_template):
    with open(input_file, 'r') as f:
        data = json.load(f)

    print(f"Wait while the items from {input_file} are getting exported...")

    items = data.get('data', {}).get('items', [])
    output_data = []

    # Loop through each entry in the JSON file
    for entry in items:
        name = entry.get('name')
        parent_resource_id = entry.get('parent-resource-id')

        # Construct OCI CLI command
        command = command_template.format(parent_resource_id=parent_resource_id, name=name)

        # Execute OCI CLI command and capture output
        output = subprocess.run(command, shell=True, capture_output=True, text=True)

        # Append output to a list
        output_data.append(json.loads(output.stdout))

    print(f"Export successful for items from {input_file}.")

    # Save output to a JSON file
    with open(output_file, 'w') as outfile:
        json.dump(output_data, outfile)

def main():
    parser = argparse.ArgumentParser(description="Export Network Firewall Policy items using OCI CLI commands")
    parser.add_argument("--id", type=str, help="Network Firewall Policy ID")
    args = parser.parse_args()

    if not args.id:
        parser.error("Please provide the Network Firewall Policy ID using --id")

    network_firewall_policy_id = args.id

    # Generate JSON files
    generate_json_files(network_firewall_policy_id)

    # Run export functions
    export_items('security_rule.json', 'security_rule_output.json', 'oci network-firewall security-rule get --network-firewall-policy-id {parent_resource_id} --security-rule-name {name}')
    export_items('addresslist.json', 'addresslist_output.json', 'oci network-firewall address-list get --network-firewall-policy-id {parent_resource_id} --address-list-name {name}')
    export_items('servicelist.json', 'servicelist_output.json', 'oci network-firewall service-list get --network-firewall-policy-id {parent_resource_id} --service-list-name {name}')
    export_items('service.json', 'service_output.json', 'oci network-firewall service get --network-firewall-policy-id {parent_resource_id} --service-name {name}')
    export_items('application.json', 'application_output.json', 'oci network-firewall application get --network-firewall-policy-id {parent_resource_id} --application-name {name}')
    export_items('applicationlist.json', 'applicationlist_output.json', 'oci network-firewall application-group get --network-firewall-policy-id {parent_resource_id} --application-group-name {name}')

if __name__ == "__main__":
    main()
