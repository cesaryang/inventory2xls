import re
import glob
import argparse
from openpyxl import Workbook

# Set up the argument parser
parser = argparse.ArgumentParser(description="Convert Cisco 'show inventory' output to Excel")
parser.add_argument("input_pattern", help="The pattern to match input files (supports wildcards)")
args = parser.parse_args()

# Use the provided input pattern to find matching files
input_files = glob.glob(args.input_pattern)

# Process each matching input file
for input_filename in input_files:
    # Generate the output filename based on the input filename
    if input_filename.lower().endswith('.txt'):
        output_filename = input_filename[:-4] + '.xlsx'
    else:
        output_filename = input_filename + '.xlsx'

    # Initialize a workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Cisco Inventory"

    # Define the header row
    headers = ["Device Name", "Name", "Description", "Product ID", "Version ID", "Serial Number"]
    ws.append(headers)

    # Initialize device name variable
    device_name = None

    # Compile regex patterns to match the relevant lines
    device_name_pattern = re.compile(r'RP/\d+/RP\d+/CPU\d+:(\S+)#')
    name_descr_pattern = re.compile(r'NAME: "(.+?)", DESCR: "(.+?)"')
    pid_vid_sn_pattern = re.compile(r'PID: (\S+) *, VID: (\S+), SN: (\S+)')

    # Process the input file
    with open(input_filename, 'r', encoding='utf-8') as infile:
        # Read the file line by line
        for line in infile:
            # Try to match the Device Name line
            device_name_match = device_name_pattern.search(line)
            if device_name_match:
                device_name = device_name_match.group(1)
                continue  # We have the device name, continue to process inventory items
            
            # Try to match the NAME and DESCR line
            name_descr_match = name_descr_pattern.search(line)
            if name_descr_match:
                name, descr = name_descr_match.groups()
                continue  # Move to the next line to find the PID, VID, and SN
            
            # Try to match the PID, VID, and SN line
            pid_vid_sn_match = pid_vid_sn_pattern.search(line)
            if pid_vid_sn_match:
                pid, vid, sn = pid_vid_sn_match.groups()
                # Append the row to the worksheet, including the device name
                ws.append([device_name, name, descr, pid, vid, sn])

    # Save the workbook to an XLSX file
    wb.save(output_filename)
    print(f"Inventory has been saved to '{output_filename}'")