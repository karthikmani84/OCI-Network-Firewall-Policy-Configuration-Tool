import argparse
import pandas as pd
import json
import sys

def write_json_to_file(data, filename):
    """Write JSON data to a file."""
    try:
        with open(filename, 'w') as json_file:
            json.dump(data, json_file, indent=4)
    except IOError as e:
        print(f"Error writing to file: {e}")
        sys.exit(1)

def excel_to_json_iplist(excel_file):
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
    
    write_json_to_file(json_list, 'iplist.json')


def excel_to_json_url_lists(excel_file):
    """Convert Excel sheet 'url_lists' to JSON."""
    try:
        df = pd.read_excel(excel_file, sheet_name='url_lists')

        #else:
            #print('columns present')
        
        json_list = []
        # Iterate over the unique names
        for name in df['name'].unique():
            #print(name)
            # Filter the DataFrame by the current name
            filtered_df = df[df['name'] == name]
            
            # Create the 'urls' list
            urls = [{"pattern": pattern, "type": "SIMPLE"} for pattern in filtered_df['pattern']]
            
            # Create the data dictionary for the current name
            data_dict = {
                "data": {
                    "name": name,
                    "urls": urls
                }
            }
        json_list.append(data_dict)
        write_json_to_file(json_list, 'url_lists.json')

    except FileNotFoundError:
        print(f"File '{excel_file}' not found.")
        sys.exit(1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        sys.exit(1)
    # Make sure the column names are correct
    expected_columns = ['name', 'pattern']
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"Expected columns {expected_columns} not found in the DataFrame. Found columns: {df.columns}")
    
def excel_to_json_service_list(excel_file):
    """
    Convert Excel sheet 'service' to JSON and write to 'services.json',
    and create 'servicelist.json' based on service types.

    Parameters:
    excel_file (str): Path to the Excel file.

    Returns:
    None
    """
    # Read data from 'service' sheet
    df = pd.read_excel(excel_file, sheet_name='service')
    json_list_service = []
    json_list_servicelist = []
    json_list_application = []
    json_list_applicationlist = []

    for index, row in df.iterrows():
        name = row['name']
        service_type = row['type']
        
        # Check if 'minimumPort' and 'maximumPort' are NaN, if so, set them to None
        min_port = int(row['minimumPort']) if not pd.isna(row['minimumPort']) else None
        max_port = int(row['maximumPort']) if not pd.isna(row['maximumPort']) else None
        icmp_type = int(row['icmpType']) if not pd.isna(row['icmpType']) else None

        # Prepare JSON object for services.json
        json_obj_service = {"name": name, "type": service_type, "portRanges": [{"minimumPort": min_port, "maximumPort": max_port}]}

        # Append to services.json if service_type is TCP_SERVICE or UDP_SERVICE
        if service_type in ["TCP_SERVICE", "UDP_SERVICE"]:
            json_list_service.append(json_obj_service)
        
        # Initialize JSON object for servicelist.json and applicationlist.json
        json_obj_servicelist = None
        json_obj_applicationlist = None

        # If service_type is TCP_SERVICE or UDP_SERVICE, create JSON object for servicelist.json
        if service_type == "TCP_SERVICE" or service_type == "UDP_SERVICE":
            json_obj_servicelist = {"name": name, "services": [name]}
        elif service_type == "SERVICE_GROUP":
            services = row['services']  # Assuming 'services' is a column in your Excel sheet
            if isinstance(services, str) and services.strip():  # Check if 'services' is non-empty string
                json_obj_servicelist = {"name": name, "services":[service.strip() for service in services.split(',')]}

        # Append non-empty JSON objects to the list
        if json_obj_servicelist is not None:
            json_list_servicelist.append(json_obj_servicelist)

        # Prepare JSON object for application.json
        if service_type in ["ICMP_TYPE"]:
            json_obj_application = {"name": name, "type": "ICMP", "icmpType": icmp_type, "icmpCode": None}
            json_list_application.append(json_obj_application)

        # If service_type is ICMP_TYPE , create JSON object for applicationlist.json
        if service_type == "ICMP_TYPE":
            json_obj_applicationlist = {"name": name, "apps": [name]}
        elif service_type == "ICMP_GROUP":
            services = row['services']
            if isinstance(services, str) and services.strip():  # Check if 'services' is non-empty string
                json_obj_applicationlist = {"name": name, "apps":[service.strip() for service in services.split(',')]}

        # Append non-empty JSON objects to the list
        if json_obj_applicationlist is not None:
            json_list_applicationlist.append(json_obj_applicationlist)




    # Write data to services.json
    write_json_to_file(json_list_service, 'service_input.json')
    # Write data to servicelist.json
    write_json_to_file(json_list_servicelist, 'service_list_input.json')
    # Write data to application.json
    write_json_to_file(json_list_application, 'application_input.json')
    # Write data to applicationlist.json
    write_json_to_file(json_list_applicationlist, 'application_list_input.json')



def replace_with_names(df_security_rules, df_iplist):
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

def convert_to_json(df_security_rules):
    """
    Convert security rules DataFrame to a list of JSON objects.

    Parameters:
    df_security_rules (pandas.DataFrame): DataFrame containing security rules data.

    Returns:
    list: List of JSON objects representing security rules.
    """
    json_list = []
    prev_rule_name = None
    
    for index, row in df_security_rules.iterrows():
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

def excel_to_json(excel_file):
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
    
    df_security_rules = replace_with_names(df_security_rules, df_iplist)
    json_output = convert_to_json(df_security_rules)
    write_json_to_file(json_output, 'securityrules.json')

def main():
    parser = argparse.ArgumentParser(description='Convert Excel file to JSON')
    parser.add_argument('-i', '--input', type=str, help='Input Excel file name')
    args = parser.parse_args()

    if args.input:
        excel_to_json(args.input)
        excel_to_json_iplist(args.input)
        excel_to_json_service_list(args.input)
        excel_to_json_url_lists(args.input)
    else:
        print("Please provide the input Excel file name using -i or --input option.")
        sys.exit(1)

if __name__ == "__main__":
    main()
