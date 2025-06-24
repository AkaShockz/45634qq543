from vehicle_transport_parser import JobParser

def test_empty_vin():
    # Test cases with and without VIN/CHASSIS
    test_cases = [
        # Case 1: No CHASSIS field in header or data
        """FROM
Stoneacre Renault Chesterfield
Brimington Road North
Chesterfield
S41 9AJ
Tel: 01246450450
TO
Leyland Motorstore
Goldenhill Lane
Farrington Ind Est
Leyland
PR25 3GG
Tel: 01772 419400
JOB NO KEY LOC BARCODE
2965424 N/A 0
MAKE MODEL COLOUR REGISTRATION
RENAULT CLIO Orange YM74BGE
COMMENTS
ORIGIN WARNINGS DESTINATION WARNINGS
VALUE
69.26""",

        # Case 2: CHASSIS field in header but empty in data
        """FROM
Eden Hyundai Basingstoke
35 London Road On A30)
BASINGSTOKE
RG24 7JD
Tel: 01256355221
TO
Wolverhampton Motorstore
Pantheon Park
Wednesfield Way
Wolverhampton
WV11 3DR
Tel: 01902 937977
JOB NO KEY LOC BARCODE
2965520 N/A 0
MAKE MODEL COLOUR REGISTRATION CHASSIS
NISSAN QASHQAI Blue RO22HOH
COMMENTS
ORIGIN WARNINGS DESTINATION WARNINGS
VALUE
115.43"""
    ]

    # Test each case
    parser = JobParser('AC01', '22/04/2025', '22/04/2025')
    
    for i, test_data in enumerate(test_cases, 1):
        print(f"\nTest Case {i}:")
        print("-" * 40)
        
        jobs = parser.parse_jobs(test_data)
        if jobs:
            job = jobs[0]
            print("Job parsed successfully:")
            print(f"Registration: {job['REG_NUMBER']}")
            print(f"Make: {job['MAKE']}")
            print(f"Model: {job['MODEL']}")
            print(f"Color: {job['COLOR']}")
            print(f"VIN: '{job['VIN']}' (should be empty)")
            print(f"Collection Phone: {job['COLLECTION_PHONE']}")
            print(f"Delivery Phone: {job['DELIVERY_CONTACT_PHONE']}")
        else:
            print("Failed to parse job")

if __name__ == "__main__":
    test_empty_vin() 