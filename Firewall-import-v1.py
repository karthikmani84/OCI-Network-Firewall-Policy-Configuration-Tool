import argparse
import pandas as pd
import json
import sys
import logging
from typing import List, Dict, Any

def write_json_to_file(data: Any, filename: str, logger: logging.Logger):
    """Write JSON data to a file."""
    try:
        with open(filename, 'w') as json_file:
            json.dump(data, json_file, indent=4)
        logger.info(f"JSON written to '{filename}'")
    except IOError as e:
        logger.error(f"Error writing to file '{filename}': {e}")
        sys.exit(1)

def excel_to_json_iplist(excel_file, logger):
    """Convert Excel sheet 'iplist' to JSON."""
    try:
        df = pd.read_excel(excel_file, sheet_name='iplist')
    except FileNotFoundError:
        print(f"File '{excel_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
    
    json_list = []
    for index, row in df.iterrows():
        name = row['name']
        addresses = [address.strip() for address in str(row['addresses']).split(',') if address.strip()]
        json_obj = {"name": name, "type": "IP", "addresses": addresses}
        json_list.append(json_obj)
    
    write_json_to_file(json_list, 'iplist.json', logger)

def excel_to_json_url_lists(excel_file: str, logger: logging.Logger) -> List[Dict[str, Any]]:
    """
    Convert Excel sheet 'url_lists' to JSON format matching the expected structure.

    Returns:
        List[Dict[str, Any]]: List of dictionaries representing URL lists.
    """
    try:
        df = pd.read_excel(excel_file, sheet_name='url_lists')
        expected_columns = ['name', 'pattern']
        if not all(col in df.columns for col in expected_columns):
            raise ValueError(f"Expected columns {expected_columns} not found. Found: {list(df.columns)}")

        result = []
        for name in df['name'].dropna().unique():
            filtered_df = df[df['name'] == name]
            urls = [{"pattern": pattern, "type": "SIMPLE"} for pattern in filtered_df['pattern'].dropna()]
            result.append({
                "name": name,
                "urls": urls
            })

        write_json_to_file(result, 'url_lists.json', logger)
        return result

    except FileNotFoundError:
        logger.error(f"File '{excel_file}' not found.")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error processing Excel file: {e}")
        sys.exit(1)
    
def excel_to_json_service_list(excel_file, logger):
    """
    Convert Excel sheet 'service' to JSON and write to 'services.json',
    and create 'servicelist.json' based on service types.
    """
    df = pd.read_excel(excel_file, sheet_name='service')
    json_list_service = []
    json_list_servicelist = []
    json_list_application = []
    json_list_applicationlist = []

    service_items_list = []
    for index, row in df.iterrows():
        name = row['name']
        service_type = row['type']
        if service_type in ["TCP_SERVICE", "UDP_SERVICE"]:
            service_items_list.append(name)        

    for index, row in df.iterrows():
        name = row['name']
        service_type = row['type']

        # Handle multiple ports
        portRanges = []
        if not pd.isna(row['minimumPort']) and not pd.isna(row['maximumPort']):
            min_ports = [p.strip() for p in str(row['minimumPort']).split(',')]
            max_ports = [p.strip() for p in str(row['maximumPort']).split(',')]

            if len(min_ports) != len(max_ports):
                logger.error(f"Port count mismatch for service '{name}': min={min_ports}, max={max_ports}")
            else:
                for min_p, max_p in zip(min_ports, max_ports):
                    try:
                        portRanges.append({
                            "minimumPort": int(min_p),
                            "maximumPort": int(max_p)
                        })
                    except ValueError:
                        logger.error(f"Invalid port number in service '{name}': {min_p}-{max_p}")

        # Handle ICMP type
        icmp_type = int(row['icmpType']) if not pd.isna(row['icmpType']) else None

        # services.json
        json_obj_service = {"name": name, "type": service_type}
        if service_type in ["TCP_SERVICE", "UDP_SERVICE"]:
            json_obj_service["portRanges"] = portRanges
            json_list_service.append(json_obj_service)

        json_obj_servicelist = None
        json_obj_applicationlist = None

        if service_type == "TCP_SERVICE" or service_type == "UDP_SERVICE":
            json_obj_servicelist = {"name": name, "services": [name]}
        elif service_type == "SERVICE_GROUP":
            services = row['services']
            if isinstance(services, str) and services.strip():
                matching_list = []
                for service in services.split(','):
                    if service in service_items_list:
                        matching_list.append(service.strip())
                    else:
                        logger.debug(f'Error: {name} : {service} - service not in available services.')
                json_obj_servicelist = {"name": name, "services":[matched_service.strip() for matched_service in matching_list]}

        if json_obj_servicelist is not None:
            json_list_servicelist.append(json_obj_servicelist)

        if service_type in ["ICMP_TYPE"]:
            json_obj_application = {
                "name": name,
                "type": "ICMP",
                "icmpType": icmp_type,
                "icmpCode": None
            }
            json_list_application.append(json_obj_application)

        if service_type == "ICMP_TYPE":
            json_obj_applicationlist = {"name": name, "apps": [name]}
        elif service_type == "ICMP_GROUP":
            services = row['services']
            if isinstance(services, str) and services.strip():
                json_obj_applicationlist = {"name": name, "apps":[service.strip() for service in services.split(',')]}

        if json_obj_applicationlist is not None:
            json_list_applicationlist.append(json_obj_applicationlist)

    write_json_to_file(json_list_service, 'service_input.json', logger)
    write_json_to_file(json_list_servicelist, 'service_list_input.json', logger)
    write_json_to_file(json_list_application, 'application_input.json', logger)
    write_json_to_file(json_list_applicationlist, 'application_list_input.json', logger)

def replace_with_names(df_security_rules, df_iplist, logger):
    """Replace IP addresses with corresponding names, preserving existing names."""
    df_security_rules.columns = df_security_rules.columns.str.strip()
    address_to_name = dict(zip(df_iplist['addresses'], df_iplist['name']))
    
    def replace_with_names_single(value):
        if isinstance(value, str):
            elements = value.split(',')
            replaced_elements = []
            for element in elements:
                element = element.strip()
                if element in address_to_name:
                    replaced_elements.append(address_to_name[element])
                elif element:
                    replaced_elements.append(element)
            return ', '.join(replaced_elements)
        else:
            return value

    df_security_rules['Source Address Lists'] = df_security_rules['Source Address Lists'].apply(replace_with_names_single)
    df_security_rules['Destination Address Lists'] = df_security_rules['Destination Address Lists'].apply(replace_with_names_single)
    
    return df_security_rules

def convert_to_json(df_security_rule, logger):
    """
    Convert security rules DataFrame to a list of JSON objects.

    Parameters:
    df_security_rules (pandas.DataFrame): DataFrame containing security rules data.

    Returns:
    list: List of JSON objects representing security rules.
    """
    json_list = []
    prev_rule_name = None
    
    for index, row in df_security_rule.iterrows():
        source_addresses = [address.strip() for address in str(row['Source Address Lists']).split(',') if address.strip()] if not pd.isna(row['Source Address Lists']) and row['Source Address Lists'] else []
        destination_addresses = [address.strip() for address in str(row['Destination Address Lists']).split(',') if address.strip()] if not pd.isna(row['Destination Address Lists']) and row['Destination Address Lists'] else []
        
        # Handling empty or NaN values for action
        action = row['Action'] if not pd.isna(row['Action']) and row['Action'] else "ALLOW"
        
        # Handling empty or NaN values for ports
        service_ports = [port.strip() for port in str(row['Service Lists']).split(',') if port.strip()] if not pd.isna(row['Service Lists']) and row['Service Lists'] else []
        
        url_lists = [url.strip() for url in str(row['Url Lists']).split(',') if url.strip()] if not pd.isna(row['Url Lists']) and row['Url Lists'] else []

        # Handling empty or NaN values for application
        icmp_code = [code.strip() for code in str(row['Application Lists']).split(',') if code.strip()] if not pd.isna(row['Application Lists']) and row['Application Lists'] else []
        
        # Create the JSON object excluding empty source and destination fields
        json_obj = {
            "name": row['name'],
            "condition": {
                "sourceAddress": source_addresses,
                "destinationAddress": destination_addresses,
                "service": service_ports,
                "url": url_lists,
                "application": icmp_code
            },
            "position": {"afterRule": prev_rule_name} if prev_rule_name else {},
            "action": action
        }
        json_list.append(json_obj)
        prev_rule_name = row['name']
    
    return json_list

def excel_to_json(excel_file, logger):
    """Convert Excel sheet 'security-rules' to JSON."""
    try:
        df_security_rules = pd.read_excel(excel_file, sheet_name='security-rules')
        df_iplist = pd.read_excel(excel_file, sheet_name='iplist')
    except FileNotFoundError:
        print(f"File '{excel_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
    
    df_security_rules = replace_with_names(df_security_rules, df_iplist, logger)
    json_output = convert_to_json(df_security_rules, logger)
    write_json_to_file(json_output, 'securityrules.json', logger)


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
    parser = argparse.ArgumentParser(description='Convert Excel file to JSON')
    parser.add_argument('-i', '--input', type=str, help='Input Excel file name')
    args = parser.parse_args()
    logger.debug("Done parsing args")
    logger.debug(f"args = {args}")
    return args 


def main():
    try: 
        # set up logging stuff
        logger = setup_logging(
            log_filename="import_policies.log", 
            logger_name="import_policies", 
            console_level=logging.INFO, 
            file_level=logging.DEBUG
        )
        args = parse_arguments(logger)
        logger.info(f"STARTING IMPORT-POLICIES.PY")

        if args.input:
            excel_to_json(args.input, logger)
            excel_to_json_iplist(args.input, logger)
            excel_to_json_service_list(args.input, logger)
            excel_to_json_url_lists(args.input, logger)
        else:
            print("Please provide the input Excel file name using -i or --input option.")
            sys.exit(1)

        logger.info(f"FINISHED IMPORT-POLICIES.PY")
    except SystemExit:
        pass # nothing wrong here, just exit if we catch this

    except BaseException as e:
        logger.exception(f"Something went wrong:  {e}")


if __name__ == "__main__":
    main()
