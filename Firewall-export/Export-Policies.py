import json
import subprocess
import re
import argparse
import logging


def run_oci_command(command, verbose, logger):
    """
    Runs the given OCI CLI command.
    """
    subprocess.run(command, shell=True)
    print()

def execute_oci_command(command, verbose, logger):
    """
    Executes the given OCI CLI command and captures the output.
    """
    output = subprocess.run(command, shell=True, capture_output=True, text=True)
    if verbose:
        logger.debug(f"Verbose - Command: {command}")
        logger.debug(json.loads(output.stdout))
    return json.loads(output.stdout)

def generate_json_files(network_firewall_policy_id, tokenstring, verbose, logger):
    """
    Generates JSON files containing information about network firewall policy items.
    """
    # Define OCI CLI commands to generate JSON files
    oci_commands = [
        "oci network-firewall security-rule list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > security_rule.json",
        "oci network-firewall address-list list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > addresslist.json",
        "oci network-firewall service list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > service.json",
        "oci network-firewall service-list list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > servicelist.json",
        "oci network-firewall application list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > application.json",
        "oci network-firewall application-group list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > applicationlist.json",
        "oci network-firewall url-list list --network-firewall-policy-id {network_firewall_policy_id} --all {tokenstring} > url_list.json"
    ]

    # Replace {network_firewall_policy_id} with actual value
    for i, command in enumerate(oci_commands):
        oci_commands[i] = command.format(network_firewall_policy_id=network_firewall_policy_id, tokenstring=tokenstring)


    # Run OCI CLI commands to generate JSON files
    for command in oci_commands:
        logger.info(f"Running: {command}.")
        run_oci_command(command, verbose, logger)
    print()

def export_items(input_file, output_file, command_template, tokenstring, verbose, logger):
    """
    Export items from input JSON file using OCI CLI commands based on the command template.
    """
    with open(input_file, 'r') as f:
        data = json.load(f)

    logger.info(f"Wait while the items from {input_file} are getting exported...")
    print()

    items = data.get('data', {}).get('items', [])
    output_data = []

    # Loop through each entry in the JSON file
    for entry in items:
        name = entry.get('name')
        parent_resource_id = entry.get('parent-resource-id')

        # Construct OCI CLI command
        command = command_template.format(parent_resource_id=parent_resource_id, name=name, tokenstring=tokenstring)

        # Execute OCI CLI command using the helper function
        logger.info(f"Running: {command}.")
        output = execute_oci_command(command, verbose, logger)

        # Append output to a list
        output_data.append(output)

    print(f"Export successful for items from {input_file}.")
    print()

    # Save output to a JSON file
    with open(output_file, 'w') as outfile:
        json.dump(output_data, outfile)


def setup_logging(log_filename, logger_name, console_level, file_level):
    """
    setup_logging
    
    Takes in
    log_filename
    logger_name
    console_level
    file_level
    
    Returns logger
    """
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler(log_filename)
    fh.setLevel(file_level)
    ch = logging.StreamHandler()
    ch.setLevel(console_level)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger

def parse_arguments(logger):
    """
    parse_arguments

    Takes in
    logger

    Returns args
    """
    logger.debug("Parsing args")
    parser = argparse.ArgumentParser()
    parser.add_argument('-t', '--token', action='store_true', help='./%(prog)s -t')
    parser.add_argument('-v', '--verbose', action='store_true', help='./%(prog)s -v')
    parser.set_defaults(verbose=False)
    parser.set_defaults(token=False)
    args = parser.parse_args()
    logger.debug("Done parsing args")
    logger.debug(f"args = {args}")
    return args 


def main():

    try: 
        # set up logging stuff
        logger = setup_logging(
            log_filename="export_policies.log", 
            logger_name="oci-access-check", 
            console_level=logging.INFO, 
            file_level=logging.DEBUG
        )
        args = parse_arguments(logger)

        if args.token:
            tokenstring = "--auth security_token"
        else: 
            tokenstring = ""

        verbose = False
        if args.verbose:
            verbose = True

        logger.info(f"STARTING EXPORT-POLICIES.PY")

        print("Please provide the Network Firewall Policy ID:")
        network_firewall_policy_id = input().strip()
        print()

        # Validate the format of the Network Firewall Policy ID
        if not re.match(r'^ocid1\.networkfirewallpolicy\.oc1\..*$', network_firewall_policy_id):
            print("Invalid Network Firewall Policy ID. It should start with 'ocid1.networkfirewallpolicy.oc1.'. Exiting.")
            return

        # Generate JSON files
        generate_json_files(network_firewall_policy_id, tokenstring, verbose, logger)

        # Run export functions
        export_items('security_rule.json', 'security_rule_output.json', 'oci network-firewall security-rule get --network-firewall-policy-id {parent_resource_id} --security-rule-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('addresslist.json', 'addresslist_output.json', 'oci network-firewall address-list get --network-firewall-policy-id {parent_resource_id} --address-list-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('servicelist.json', 'servicelist_output.json', 'oci network-firewall service-list get --network-firewall-policy-id {parent_resource_id} --service-list-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('service.json', 'service_output.json', 'oci network-firewall service get --network-firewall-policy-id {parent_resource_id} --service-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('application.json', 'application_output.json', 'oci network-firewall application get --network-firewall-policy-id {parent_resource_id} --application-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('applicationlist.json', 'applicationlist_output.json', 'oci network-firewall application-group get --network-firewall-policy-id {parent_resource_id} --application-group-name {name} {tokenstring}', tokenstring, verbose, logger)
        export_items('url_list.json', 'url_list_output.json', 'oci network-firewall url-list get --network-firewall-policy-id {parent_resource_id} --url-list-name {name}  {tokenstring}', tokenstring, verbose, logger)

        logger.info(f"FINISHED EXPORT-POLICIES.PY")
    except SystemExit:
        pass # nothing wrong here, just exit if we catch this

    except BaseException as e:
        logger.exception(f"Something went wrong:  {e}")


if __name__ == "__main__":
    main()
