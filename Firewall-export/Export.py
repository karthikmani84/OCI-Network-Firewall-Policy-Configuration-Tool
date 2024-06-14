import subprocess

def main():
    print(f"Export the OCI Network firewall policy using this tool!")
    print(f"Choose an option:")
    print(f"1. Export Policies to Json format")
    print(f"2. Convert Policies from Json to Excel Format")

    choice = input(f"Enter your choice (1 or 2): ")

    if choice == "1":
        sectoken = input(f"You've chosen to Export. Will you be using a security token? Y/N?")
        if sectoken.upper() == 'Y':
            print(f"Run \'oci network-firewall network-firewall-policy get --network-firewall-policy-id <ocid>  --auth security_token\' prior to running this script to authenticate. \nAbort with CTRL+C if you have not authenticated.")
            print(f"Calling Export-Policies.py with token...")
            subprocess.run(["python3", "Export-Policies.py", "-t"])
        else:
            print(f"Calling Export-Policies.py...")
            subprocess.run(["python3", "Export-Policies.py"])            
    elif choice == "2":
        print(f"Calling Convert-Policies.py...")
        subprocess.run(["python3", "Convert-Policies.py"])
    else:
        print(f"Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()
