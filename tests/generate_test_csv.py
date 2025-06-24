from vehicle_transport_parser import JobParser
import sys

def generate_csv(input_file, job_type='AC01'):
    """Generate a CSV file from the specified input file."""
    # Create a parser with the specified job type
    parser = JobParser(job_type, '22/04/2025', '22/04/2025')
    
    # Read the input file
    with open(input_file, 'r') as f:
        text = f.read()
    
    # Parse the jobs
    jobs = parser.parse_jobs(text)
    
    # Generate CSV file
    output_file = f"{job_type.lower()}_test_output.csv"
    parser.save_to_csv(jobs, output_file)
    
    print(f"Successfully processed {len(jobs)} jobs from {input_file}")
    print(f"Output saved to {output_file}")
    print("\nJob details:")
    for i, job in enumerate(jobs, 1):
        print(f"{i}. Make: {job['MAKE']}, Model: {job['MODEL']}, Reg: {job['REG_NUMBER']}")

if __name__ == "__main__":
    # Check command line arguments
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = "test_input.txt"  # Default to test input
    
    # Check for job type argument
    job_type = 'AC01'  # Default
    if len(sys.argv) > 2:
        job_type = sys.argv[2]
    
    generate_csv(input_file, job_type) 