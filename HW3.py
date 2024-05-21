import argparse
import csv

parser = argparse.ArgumentParser(description='Process Baselight and Xytech files')
parser.add_argument('--baselight', type=str, help='Path to the Baselight file')
parser.add_argument('--xytech', type=str, help='Path to the Xytech file')
args = parser.parse_args()

BL_File = open(args.baselight, "r")  # Open Baselight file
with open(args.xytech) as f:
    XY_File = f.read().splitlines()
isVerbose = 0  # Set to True for verbose output

output_data = []  # List to store data for CSV export

for currentReadLine in BL_File:
    parseline = currentReadLine.split()  # Split the line
    if parseline:  # Check if the list is not empty
        currentFolder = parseline.pop(0)  # Get the current folder
        parseFolder = currentFolder.split("/")  # Split current folder by "/"
        if len(parseFolder) > 1:
            parseFolder.pop(1)  # Remove the second element if exists
        newFolder = "/".join(parseFolder)  # Reconstruct the folder path
        for techfile in XY_File:
            if newFolder in techfile:
                currentFolder = techfile.strip()
        tempStart = None
        tempLast = None
        for number in parseline:
            if not number.isdigit():
                continue
            number = int(number)
            if tempStart is None:
                tempStart = number
                tempLast = number
            elif number == tempLast + 1:
                tempLast = number
            else:
                if tempStart == tempLast:
                    output_data.append([currentFolder, tempStart])
                else:
                    output_data.append([currentFolder, f"{tempStart}-{tempLast}"])  # Concatenate range
                tempStart = number
                tempLast = number
        if tempStart is not None:
            if tempStart == tempLast:
                output_data.append([currentFolder, tempStart])
            else:
                output_data.append([currentFolder, f"{tempStart}-{tempLast}"])  # Concatenate range
    else:
        print("Empty line, skipping...")

BL_File.close()  # Close Baselight file

# Write data to CSV file
with open("output.csv", "w", newline="") as csv_file:
    writer = csv.writer(csv_file)
    writer.writerow(["Producer", "Operator", "Job", "Notes"])  # Write header
    for _ in range(2):  # Skip 4 rows
        writer.writerow([])
    writer.writerow(["Location", "frames to fix"])
    for data in output_data:
        writer.writerow(data)  # Write data

print("CSV file exported successfully!")
