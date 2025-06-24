from vehicle_transport_parser import JobParser
import csv
import os

# Create a parser with the AC01 job type and collection/delivery dates
parser = JobParser('AC01', '22/04/2025', '22/04/2025')

# Read the test input file
with open('test_input.txt', 'r') as f:
    text = f.read()

# Parse the jobs
jobs = parser.parse_jobs(text)

# Generate a test CSV file
test_filename = "test_empty_fields.csv"
parser.jobs = jobs  # Make sure the jobs are set in the parser
with open(test_filename, 'w', newline='') as csvfile:
    # Define fields in the EXACT order shown in the Excel spreadsheet image
    fieldnames = [
        'REG NUMBER',          # Column A
        'VIN',                 # Column B
        'MAKE',               # Column C
        'MODEL',              # Column D
        'COLLECTION DATE',    # Column E
        'YOUR REF NO',        # Column F
        'COLLECTION ADDR1',   # Column G
        'COLLECTION ADDR2',   # Column H
        'COLLECTION ADDR3',   # Column I
        'COLLECTION ADDR4',   # Column J
        'COLLECTION POSTCODE', # Column K
        'COLLECTION CONTACT NAME', # Column L
        'COLLECTION PHONE',   # Column M
        'DELIVERY DATE',      # Column N
        'DELIVERY ADDR1',     # Column O
        'DELIVERY ADDR2',     # Column P
        'DELIVERY ADDR3',     # Column Q
        'DELIVERY ADDR4',     # Column R
        'DELIVERY POSTCODE',  # Column S
        'DELIVERY CONTACT NAME', # Column T
        'DELIVERY CONTACT PHONE', # Column U
        'SPECIAL INSTRUCTIONS', # Column V
        'PRICE',              # Column W
        'CUSTOMER REF',       # Column X
        'TRANSPORT TYPE'      # Column Y
    ]
    
    # Map our internal field names to the Excel header names
    field_mapping = {
        'REG_NUMBER': 'REG NUMBER',
        'COLOR': 'COLOUR'
    }
    
    # Fields to leave empty in the output
    empty_fields = ['SPECIAL_INSTRUCTIONS', 'PRICE']
    
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    for job in jobs:
        # Transform the job data to match the Excel headers
        row = {}
        for field in fieldnames:
            # Leave specified fields empty
            if field in empty_fields:
                row[field] = ''
                continue
                
            # Check if this field needs to be mapped from a different name
            internal_field = field
            for k, v in field_mapping.items():
                if field == v:
                    internal_field = k
                    break
            
            # Get the value from the job dictionary
            row[field] = job.get(internal_field, '')
        
        writer.writerow(row)

print(f"Created CSV file: {test_filename}")
print(f"Number of jobs written: {len(jobs)}")

# Read the CSV back to verify that the fields are empty
with open(test_filename, 'r', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    rows = list(reader)
    
    # Check specific fields
    for i, row in enumerate(rows):
        print(f"Row {i+1}:")
        print(f"  Registration: {row['REG NUMBER']}")
        print(f"  Make: {row['MAKE']}")
        print(f"  Model: {row['MODEL']}")
        print(f"  Special Instructions: '{row['SPECIAL INSTRUCTIONS']}' (should be empty)")
        print(f"  Price: '{row['PRICE']}' (should be empty)")
        print()

print("Test complete.") 