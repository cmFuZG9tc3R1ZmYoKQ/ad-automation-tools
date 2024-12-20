import pyodbc
import subprocess
import re

# Database connection settings
#SERVER = '\\127.0.0.1\\pipe\\MICROSOFT##WID\\tsql\\query'
SERVER = '\\\\.\\pipe\\MICROSOFT##WID\\tsql\\query'
DATABASE = 'SUSDB'

# Sanitize function to remove unsafe characters - possible injections
def sanitize_input(input_string):
    return re.sub(r'[()\'`"\[\]]', '', input_string)

# Helper function to execute PowerShell commands
def execute_powershell_command(command, params):
    try:
        sanitized_params = [sanitize_input(param) for param in params]
        completed_command = command.format(*sanitized_params)
        subprocess.run(["powershell", "-Command", completed_command], check=True)
    except subprocess.CalledProcessError as e:
        print("Failed to execute command: {}".format(e))

# Functions to update AD attributes
def set_computer_attribute(pc, attribute, value):
    command = (
        "Set-ADComputer -Identity '{0}' -Replace @{{{1} = \"{2}\"}}"
    )
    execute_powershell_command(command, [pc, attribute, value])

# Retrieve rows from the database
def get_rows():
    try:
        connection = pyodbc.connect(
            "Driver={{ODBC Driver 17 for SQL Server}};Server={};Database={};Trusted_Connection=yes;".format(SERVER, DATABASE)
        )
        cursor = connection.cursor()
        cursor.execute(
            """
            SELECT 
                ct.FullDomainName, 
                ctd.ComputerMake,
                ctd.ComputerModel,
                ctd.BiosVersion,
                ctd.OSLocale 
            FROM tbComputerTargetDetail ctd 
            LEFT JOIN tbComputerTarget ct 
            ON ctd.TargetID = ct.TargetID;
            """
        )
        return cursor.fetchall()
    except pyodbc.Error as e:
        print("Database connection failed: {}".format(e))
        return []

# Update AD attributes for each row
def update_AD(rows):
    for row in rows:
        pc = sanitize_input(row[0].split(".")[0])
        make = row[1]
        model = row[2]
        bios = row[3]
        oslocale = row[4]

        print("Updating: PC={}, Make={}, Model={}, BIOS={}, Locale={}".format(pc, make, model, bios, oslocale))

        set_computer_attribute(pc, "computerMake", make)
        set_computer_attribute(pc, "computerModel", model)
        set_computer_attribute(pc, "biosVersion", bios)
        set_computer_attribute(pc, "oSLocale", oslocale)

# Main execution
if __name__ == "__main__":
    rows = get_rows()
    if rows:
        update_AD(rows)
    else:
        print("No data retrieved from the database.")

