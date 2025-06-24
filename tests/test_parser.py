from vehicle_transport_parser import JobParser

# Create a parser with the AC01 job type and collection/delivery dates
parser = JobParser('AC01', '22/04/2025', '22/04/2025')

# Read the test input file
with open('test_input.txt', 'r') as f:
    text = f.read()

# Parse the jobs
jobs = parser.parse_jobs(text)

# Print the results
print(f"Parsed {len(jobs)} jobs:")
print("-" * 50)

for i, job in enumerate(jobs, 1):
    print(f"Job #{i}:")
    print(f"  Make: {job['MAKE']}")
    print(f"  Model: {job['MODEL']}")
    print(f"  Color: {job['COLOR']}")
    print(f"  Registration: {job['REG_NUMBER']}")
    print(f"  VIN/Chassis: {job['VIN']}")
    print(f"  Job Number: {job['YOUR_REF_NO']}")
    print(f"  Price: {job['PRICE']}")
    print(f"  Collection: {job['COLLECTION_ADDR1']}, {job['COLLECTION_POSTCODE']}")
    print(f"  Delivery: {job['DELIVERY_ADDR1']}, {job['DELIVERY_POSTCODE']}")
    print("-" * 50) 