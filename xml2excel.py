import csv
import xml.etree.ElementTree as ET
import pandas as pd
import exexcel

def xml_to_csv(xml_file, csv_file):
    
    tree = ET.parse(xml_file)
    root = tree.getroot()
    rows = []
    for element in root:
        row = []
        for child in element:
            row.append(child.text)                      
        rows.append(row)

    with open(csv_file, 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(['LOCATION_ID', 'DESTINATION_ID', 'CAPTURE_DATE', 'JOURNEY_TYPE', 'JOURNEY_DATA', 'COLOUR_ID'])
        writer.writerows(rows)

# XML file path
xml_file = 'traffic.xml'

# Convert XML to CSV
csv_file = 'traffic.csv'
xml_to_csv(xml_file, csv_file)

# Read the CSV file
csv_file = 'traffic.csv'
df = pd.read_csv(csv_file)

#run vba
exexcel.autovba()

print(f"XML data converted and saved as {csv_file}.")