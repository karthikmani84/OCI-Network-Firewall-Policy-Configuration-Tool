import subprocess

def main():
    print("Export the OCI Network firewall policy using this tool!")
    print("Choose an option:")
    print("1. Export Policies to Json format")
    print("2. Convert Policies from Json to Excel Format")

    choice = input("Enter your choice (1 or 2): ")

    if choice == "1":
        sectoken = input(f"You've chosen to Export. Will you be using a security token? Y/N?")
        if sectoken == 'Y':
            print("Calling Export-Policies.py with token...")
            subprocess.run(["python3", "Export-Policies.py", "-t"])
        else:
            print("Calling Export-Policies.py...")
            subprocess.run(["python3", "Export-Policies.py"])            
    elif choice == "2":
        print("Calling Convert-Policies.py...")
        subprocess.run(["python3", "Convert-Policies.py"])
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()
