from vehicle_transport_parser import JobParser

def test_phone_formatting():
    # Test AC01 job type
    parser = JobParser('AC01', '22/04/2025', '22/04/2025')
    
    # Test various phone number formats
    test_numbers = [
        'Tel: 01246 450450',
        'Tel: 01772 419400',
        'Phone: +441782968610',
        'T: 441902937977',
        'Tel: 01253 376904'
    ]
    
    print("Testing AC01 phone number formatting:")
    for number in test_numbers:
        cleaned = parser.clean_phone_number(number)
        print(f"Original: {number}")
        print(f"Cleaned:  {cleaned}")
        print()

if __name__ == "__main__":
    test_phone_formatting() 