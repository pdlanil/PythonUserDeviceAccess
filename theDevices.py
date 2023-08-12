import openpyxl  # Import the openpyxl library for Excel operations

# Function to read and return devices from a file
def read_devices():
    with open('devices.txt', 'r') as file:
        devices = file.read().splitlines()
    return devices

# Function to write devices to a file
def write_devices(devices):
    with open('devices.txt', 'w') as file:
        file.write("\n".join(devices))

# Function to view all devices
def view_devices(devices):
    print("\n".join(devices))

# Function to add a device to the list and file
def add_device(devices, device):
    devices.append(device)
    write_devices(devices)  # Using built-in function to write to the file
    print(f"{device} is added.")

# Function to delete a device from the list and file
def delete_device(devices, device):
    if device in devices:
        devices.remove(device)
        write_devices(devices)  # Using built-in function to write to the file
        print(f"{device} is deleted.")
    else:
        print(f"{device} not found.")

# Function to update a device in the list and file
def update_device(devices, old_device, new_device):
    if old_device in devices:
        devices[devices.index(old_device)] = new_device
        write_devices(devices)  # Using built-in function to write to the file
        print(f"{old_device} is updated to {new_device}.")
    else:
        print(f"{old_device} not found.")

# Function to search for devices by keyword and print results
def search_device(devices, keyword):
    found_devices = [device for device in devices if keyword in device]
    if found_devices:
        print("Devices matching the keyword:")
        print("\n".join(found_devices))
    else:
        print("No devices found matching the keyword.")

# Function to validate user login
def login():
    # Function to read credentials from an Excel file
    def read_credentials(filename):
        credentials = {}
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active
        for row in sheet.iter_rows(min_row=2, values_only=True):
            username = row[0]
            password = row[1]
            credentials[username] = password
        return credentials

    filename = "credentials.xlsx"
    credentials = read_credentials(filename)
    
    max_attempts = 3
    attempt = 1
    
    while attempt <= max_attempts:
        username = input("Enter your username: ")
        password = input("Enter your password: ")
        
        if username in credentials and credentials[username] == password:
            print("Login successful!")
            return True
        else:
            print(f"Invalid username or password. {max_attempts - attempt} attempts remaining.")
            attempt += 1
    else:
        print("Maximum number of attempts reached. Login failed.")
    return False

def main():
    if not login():
        print("Login to start your operation")
        return

    devices = read_devices()  # Calling read_devices function to read devices from the file

    while True:
        # Display options for device management
        print("\nWelcome to the Device Management System")
        print("1. View all devices")
        print("2. Add a device")
        print("3. Delete a device")
        print("4. Update a device")
        print("5. Search for a device")
        print("6. Exit the program")

        try:
            option = int(input("Select one option from the list (1, 2, 3, 4, or 5): "))
        except ValueError:
            print("Invalid selection. Please try again.")
            continue

        # Based on the selected option, perform relevant device management tasks
        if option == 1:
            view_devices(devices)
        elif option == 2:
            device = input("Add the device: ")
            add_device(devices, device)
        elif option == 3:
            device = input("Enter the device to delete: ")
            delete_device(devices, device)
        elif option == 4:
            old_device = input("Enter the device to update: ")
            new_device = input("Enter the updated device: ")
            update_device(devices, old_device, new_device)
        elif option == 5:
            keyword = input("Enter the keyword to search for: ")
            search_device(devices, keyword)
        elif option == 6:
            print("Thanks for using the application!")
            break
        else:
            print("Invalid selection. Please try again.")

if __name__ == "__main__":
    main()
