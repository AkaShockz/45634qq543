import re
from datetime import datetime, timedelta
import holidays

class JobParser:
    def __init__(self, collection_date, delivery_date=None):
        self.jobs = []
        self.collection_date = collection_date
        self.delivery_date = delivery_date if delivery_date else collection_date
        
    def calculate_delivery_date(self, collection_date):
        """Calculate delivery date as 3 business days from collection date."""
        if isinstance(collection_date, str):
            collection_date = datetime.strptime(collection_date, "%d/%m/%Y")
        uk_holidays = holidays.UK()
        current_date = collection_date
        business_days = 0
        while business_days < 3:
            current_date += timedelta(days=1)
            if current_date.weekday() < 5 and current_date not in uk_holidays:
                business_days += 1
        return current_date.strftime("%d/%m/%Y")
    
    def fix_location_name(self, name):
        name = name.replace('18 AC Stoke Logistics Hub', '18 Arnold Clark Stoke Logistics Hub')
        name = name.replace('4 AC Accrington Logistics Hub', '4 Arnold Clark Accrington Logistics Hub')
        if "Unit 1 Calder Park Services" in name:
            return "Wakefield Motorstore"
        if name.startswith("Unit ") and len(name.split()) >= 3:
            pass
        return name
    
    def clean_phone_number(self, phone):
        if not phone:
            return ""
        digits = ''.join(c for c in phone if c.isdigit())
        if digits.startswith('44'):
            digits = digits[2:]
        elif digits.startswith('0044'):
            digits = digits[4:]
        elif digits.startswith('0'):
            digits = digits[1:]
        if len(digits) > 10:
            digits = digits[:10]
        elif len(digits) < 10:
            digits = digits.ljust(10, '0')
        return digits

    def is_postcode(self, line):
        postcode_patterns = [
            r'^[A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2}$',
            r'^[A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2}$',
            r'^(?:Postcode|Post Code|P/Code|PC)[\s:]+[A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2}$'
        ]
        line = line.strip()
        for pattern in postcode_patterns:
            if re.match(pattern, line):
                postcode_match = re.search(r'([A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2})$', line)
                if postcode_match:
                    return postcode_match.group(1).upper()
        return None

    def parse_tabular_format(self, text):
        """Parse a tabular format where each row is a job and columns are fields."""
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        # Find the header row
        header_row = None
        for i, line in enumerate(lines):
            if 'REG' in line and 'NUMBER' in line:
                header_row = i
                break
        
        if header_row is None:
            return []
        
        # Process data rows
        jobs = []
        for i in range(header_row + 1, len(lines)):
            line = lines[i]
            if not line.strip():
                continue
            
            # Create a job with default values
            job = {}
            job['REG NUMBER'] = ''
            job['VIN'] = ''
            job['MAKE'] = ''
            job['MODEL'] = ''
            job['COLOR'] = ''
            job['COLLECTION DATE'] = self.collection_date
            job['YOUR REF NO'] = ''
            job['COLLECTION ADDR1'] = ''
            job['COLLECTION ADDR2'] = ''
            job['COLLECTION ADDR3'] = ''
            job['COLLECTION ADDR4'] = ''
            job['COLLECTION POSTCODE'] = ''
            job['COLLECTION CONTACT NAME'] = ''
            job['COLLECTION PHONE'] = ''
            job['DELIVERY DATE'] = self.delivery_date
            job['DELIVERY ADDR1'] = ''
            job['DELIVERY ADDR2'] = ''
            job['DELIVERY ADDR3'] = ''
            job['DELIVERY ADDR4'] = ''
            job['DELIVERY POSTCODE'] = ''
            job['DELIVERY CONTACT NAME'] = ''
            job['DELIVERY CONTACT PHONE'] = ''
            job['SPECIAL INSTRUCTIONS'] = 'Please call 1hour before collection'
            job['PRICE'] = ''
            job['CUSTOMER REF'] = 'AC01'
            job['TRANSPORT TYPE'] = ''
            
            # Split the line by tabs or multiple spaces
            fields = re.split(r'\t+|\s{2,}', line)
            fields = [f.strip() for f in fields if f.strip()]
            
            # Try to identify fields by position and pattern matching
            for idx, field in enumerate(fields):
                # Registration number (typically first column)
                if idx == 0 and re.match(r'^[A-Z0-9]+$', field):
                    job['REG NUMBER'] = field
                
                # Sheffield, Furrows F Benbow etc. (location/model)
                elif idx in [1, 2, 3] and not re.match(r'^[0-9/]+$', field):
                    if job['MAKE'] == '':
                        job['MAKE'] = field
                    elif job['MODEL'] == '':
                        job['MODEL'] = field
                    elif job['COLOR'] == '':
                        job['COLOR'] = field
                
                # Collection address (typically has postcode format)
                elif re.match(r'^[A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2}$', field):
                    if job['COLLECTION POSTCODE'] == '':
                        job['COLLECTION POSTCODE'] = field
                    elif job['DELIVERY POSTCODE'] == '':
                        job['DELIVERY POSTCODE'] = field
                
                # Date format (e.g., 2E-09)
                elif re.match(r'^[0-9]{1,2}[A-Z]-[0-9]{2}$', field):
                    if job['COLLECTION DATE'] == self.collection_date:
                        job['COLLECTION DATE'] = field
                    else:
                        job['DELIVERY DATE'] = field
                
                # Phone number format (e.g., 361Arms)
                elif re.match(r'^[0-9]{3}[A-Za-z]+', field):
                    if job['COLLECTION PHONE'] == '':
                        job['COLLECTION PHONE'] = field
                    elif job['DELIVERY CONTACT PHONE'] == '':
                        job['DELIVERY CONTACT PHONE'] = field
                
                # Postcode format (e.g., BB5 4EQ)
                elif re.match(r'^[A-Z]{1,2}[0-9]{1,2}\s*[0-9][A-Z]{2}$', field):
                    if job['COLLECTION POSTCODE'] == '':
                        job['COLLECTION POSTCODE'] = field
                    elif job['DELIVERY POSTCODE'] == '':
                        job['DELIVERY POSTCODE'] = field
                
                # Address (e.g., Johnson, 38 Derby Liverpool)
                elif re.match(r'^[A-Za-z]+,?\s+[0-9]+\s+[A-Za-z]+', field):
                    if job['COLLECTION ADDR1'] == '':
                        job['COLLECTION ADDR1'] = field
                    elif job['DELIVERY ADDR1'] == '':
                        job['DELIVERY ADDR1'] = field
                
                # Special instructions (e.g., Please call 1hour)
                elif "call" in field.lower() and "hour" in field.lower():
                    job['SPECIAL INSTRUCTIONS'] = field
                
                # Customer reference (e.g., AC01)
                elif re.match(r'^[A-Z]{2}[0-9]{2}$', field):
                    job['CUSTOMER REF'] = field
            
            # If we have a valid registration number, add the job
            if job['REG NUMBER'] and job['REG NUMBER'] != '######':
                jobs.append(job)
        
        return jobs

    def parse_jobs(self, text):
        # First try the traditional format with FROM/TO markers
        job_texts = re.split(r'\nFROM\n', text)
        if len(job_texts) > 1:  # If we found FROM markers
            job_texts = [t for t in job_texts if t.strip()]
            for job_text in job_texts:
                if not job_text.startswith('FROM'):
                    job_text = 'FROM\n' + job_text
                if not re.search(r'TO\n', job_text):
                    continue
                job = self.parse_single_job(job_text)
                if job and job.get('REG NUMBER'):
                    if 'SPECIAL INSTRUCTIONS' not in job or not job['SPECIAL INSTRUCTIONS']:
                        job['SPECIAL INSTRUCTIONS'] = 'Please call 1 hour before collection'
                    self.jobs.append(job)
            return self.jobs
        
        # If no FROM markers found, try parsing as tabular format (like in the image)
        tabular_jobs = self.parse_tabular_format(text)
        if tabular_jobs:
            self.jobs.extend(tabular_jobs)
            return self.jobs
        
        # If no jobs found yet, try parsing as simple format
        lines = text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        # Check if we have a header row
        header_row = None
        for i, line in enumerate(lines):
            if 'REG' in line and 'NUMBER' in line and 'MAKE' in line and 'MODEL' in line:
                header_row = i
                break
        
        if header_row is not None:
            # Skip header row
            data_rows = lines[header_row+1:]
            
            # Process each row as a job
            for row in data_rows:
                if not row.strip():
                    continue
                    
                # Create a job with default values
                job = {}
                job['REG NUMBER'] = ''
                job['VIN'] = ''
                job['MAKE'] = ''
                job['MODEL'] = ''
                job['COLOR'] = ''
                job['COLLECTION DATE'] = self.collection_date
                job['YOUR REF NO'] = ''
                job['COLLECTION ADDR1'] = ''
                job['COLLECTION ADDR2'] = ''
                job['COLLECTION ADDR3'] = ''
                job['COLLECTION ADDR4'] = ''
                job['COLLECTION POSTCODE'] = ''
                job['COLLECTION CONTACT NAME'] = ''
                job['COLLECTION PHONE'] = ''
                job['DELIVERY DATE'] = self.delivery_date
                job['DELIVERY ADDR1'] = ''
                job['DELIVERY ADDR2'] = ''
                job['DELIVERY ADDR3'] = ''
                job['DELIVERY ADDR4'] = ''
                job['DELIVERY POSTCODE'] = ''
                job['DELIVERY CONTACT NAME'] = ''
                job['DELIVERY CONTACT PHONE'] = ''
                job['SPECIAL INSTRUCTIONS'] = 'Please call 1hour before collection'
                job['PRICE'] = ''
                job['CUSTOMER REF'] = 'AC01'
                job['TRANSPORT TYPE'] = ''
                
                # Split the row by tabs or multiple spaces
                fields = re.split(r'\t+|\s{2,}', row)
                fields = [f.strip() for f in fields if f.strip()]
                
                # Try to extract fields based on position
                for i, field in enumerate(fields):
                    if i == 0 and re.match(r'^[A-Z]{2}\d{2}[A-Z]{3}$', field):
                        job['REG NUMBER'] = field
                    elif i == 1:
                        job['MAKE'] = field
                    elif i == 2:
                        job['MODEL'] = field
                    elif i == 3:
                        job['COLOR'] = field
                    elif i == 4:
                        job['COLLECTION DATE'] = field
                    elif i == 5:
                        job['YOUR REF NO'] = field
                    elif i == 6:
                        job['COLLECTION ADDR1'] = field
                    elif i == 7:
                        job['COLLECTION POSTCODE'] = field
                    elif i == 8:
                        job['COLLECTION PHONE'] = field
                    elif i == 9:
                        job['DELIVERY DATE'] = field
                    elif i == 10:
                        job['DELIVERY ADDR1'] = field
                    elif i == 11:
                        job['DELIVERY POSTCODE'] = field
                    elif i == 12:
                        job['DELIVERY CONTACT PHONE'] = field
                    elif i == 13:
                        job['SPECIAL INSTRUCTIONS'] = field
                    elif i == 14:
                        job['PRICE'] = field
                    elif i == 15:
                        job['CUSTOMER REF'] = field
                    elif i == 16:
                        job['TRANSPORT TYPE'] = field
                
                # If we have a valid registration number, add the job
                if job['REG NUMBER']:
                    self.jobs.append(job)
        
        return self.jobs
    
    def parse_address_lines(self, lines):
        preserved_patterns = [
            (r'St\.\s+[A-Z][a-z]+', lambda m: m.group().replace('.', '@')),
            (r'St\s+[A-Z][a-z]+', lambda m: m.group().replace(' ', '#')),
            (r'D\.\s*M\.\s*Keith', lambda m: m.group().replace('.', '@')),
            (r'[A-Z]\.\s+[A-Z]\.\s+\w+', lambda m: m.group().replace('.', '@')),
        ]
        processed_lines = []
        for line in lines:
            if not line.strip():
                continue
            processed_line = line
            for pattern, replacement in preserved_patterns:
                processed_line = re.sub(pattern, replacement, processed_line)
            processed_line = processed_line.replace('@', '.').replace('#', ' ')
            processed_lines.append(processed_line.strip())
        return processed_lines

    def clean_duplicate_towns(self, lines):
        if not lines:
            return lines
        cleaned_lines = []
        i = 0
        while i < len(lines):
            if i == len(lines) - 1 or lines[i].strip().upper() != lines[i + 1].strip().upper():
                cleaned_lines.append(lines[i])
                i += 1
            else:
                i += 1
        return cleaned_lines

    def parse_single_job(self, job_text):
        job = {}
        job['REG NUMBER'] = ''
        job['VIN'] = ''
        job['MAKE'] = ''
        job['MODEL'] = ''
        job['COLOR'] = ''
        job['COLLECTION DATE'] = self.collection_date
        job['YOUR REF NO'] = ''
        job['COLLECTION ADDR1'] = ''
        job['COLLECTION ADDR2'] = ''
        job['COLLECTION ADDR3'] = ''
        job['COLLECTION ADDR4'] = ''
        job['COLLECTION POSTCODE'] = ''
        job['COLLECTION CONTACT NAME'] = ''
        job['COLLECTION PHONE'] = ''
        job['DELIVERY DATE'] = self.delivery_date
        job['DELIVERY ADDR1'] = ''
        job['DELIVERY ADDR2'] = ''
        job['DELIVERY ADDR3'] = ''
        job['DELIVERY ADDR4'] = ''
        job['DELIVERY POSTCODE'] = ''
        job['DELIVERY CONTACT NAME'] = ''
        job['DELIVERY CONTACT PHONE'] = ''
        job['SPECIAL INSTRUCTIONS'] = 'Please call 1hour before collection'
        job['PRICE'] = ''
        job['CUSTOMER REF'] = 'AC01'
        job['TRANSPORT TYPE'] = ''
        
        # Simple parsing for the format shown in the image
        lines = job_text.split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        # Look for registration number pattern (e.g., AB12CDE)
        for line in lines:
            reg_match = re.search(r'([A-Z]{2}\d{2}\s*[A-Z]{3})', line)
            if reg_match:
                job['REG NUMBER'] = reg_match.group(1).replace(' ', '')
                break
        
        # Look for collection address
        collection_section = False
        collection_lines = []
        for i, line in enumerate(lines):
            if "FROM" in line:
                collection_section = True
                continue
            if collection_section and "TO" in line:
                collection_section = False
                break
            if collection_section:
                collection_lines.append(line)
        
        # Process collection address
        if collection_lines:
            # Check for postcode
            postcode_idx = None
            for idx, line in enumerate(collection_lines):
                postcode = self.is_postcode(line)
                if postcode:
                    job['COLLECTION POSTCODE'] = postcode
                    postcode_idx = idx
                    break
            
            # Extract address
            if postcode_idx is not None:
                addr_lines = collection_lines[:postcode_idx]
            else:
                addr_lines = collection_lines
            
            # Clean and assign address lines
            addr_lines = self.parse_address_lines(addr_lines)
            addr_lines = self.clean_duplicate_towns(addr_lines)
            
            for idx, line in enumerate(addr_lines[:4]):
                job[f'COLLECTION ADDR{idx+1}'] = line
            
            # Look for phone number
            for line in collection_lines:
                phone_match = re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+([+\d()\s-]+)', line, re.IGNORECASE)
                if phone_match:
                    job['COLLECTION PHONE'] = self.clean_phone_number(phone_match.group(1))
                    break
        
        # Look for delivery address
        delivery_section = False
        delivery_lines = []
        for i, line in enumerate(lines):
            if "TO" in line:
                delivery_section = True
                continue
            if delivery_section and i < len(lines) - 1 and any(keyword in lines[i+1] for keyword in ["Vehicle", "Special", "FROM"]):
                delivery_section = False
                break
            if delivery_section:
                delivery_lines.append(line)
        
        # Process delivery address
        if delivery_lines:
            # Check for postcode
            postcode_idx = None
            for idx, line in enumerate(delivery_lines):
                postcode = self.is_postcode(line)
                if postcode:
                    job['DELIVERY POSTCODE'] = postcode
                    postcode_idx = idx
                    break
            
            # Extract address
            if postcode_idx is not None:
                addr_lines = delivery_lines[:postcode_idx]
            else:
                addr_lines = delivery_lines
            
            # Clean and assign address lines
            addr_lines = self.parse_address_lines(addr_lines)
            addr_lines = self.clean_duplicate_towns(addr_lines)
            
            for idx, line in enumerate(addr_lines[:4]):
                job[f'DELIVERY ADDR{idx+1}'] = line
            
            # Look for phone number
            for line in delivery_lines:
                phone_match = re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+([+\d()\s-]+)', line, re.IGNORECASE)
                if phone_match:
                    job['DELIVERY CONTACT PHONE'] = self.clean_phone_number(phone_match.group(1))
                    break
        
        # Look for special instructions
        special_section = False
        special_lines = []
        for i, line in enumerate(lines):
            if "Special" in line and "Instructions" in line:
                special_section = True
                continue
            if special_section and i < len(lines) - 1 and "FROM" in lines[i+1]:
                special_section = False
                break
            if special_section:
                special_lines.append(line)
        
        if special_lines:
            job['SPECIAL INSTRUCTIONS'] = ' '.join(special_lines)
        
        # If no registration found in the text, try to extract from the first line
        if not job['REG NUMBER'] and lines:
            reg_match = re.search(r'([A-Z]{2}\d{2}\s*[A-Z]{3})', lines[0])
            if reg_match:
                job['REG NUMBER'] = reg_match.group(1).replace(' ', '')
        
        # If we have a simple format like in the image, try direct field extraction
        # This handles formats where each line is a field
        if len(lines) >= 10 and not job['REG NUMBER']:
            # Try to find the registration number in any line
            for line in lines:
                if re.match(r'^[A-Z]{2}\d{2}[A-Z]{3}$', line.strip()):
                    job['REG NUMBER'] = line.strip()
                    break
        
        return job

class BC04Parser:
    def __init__(self, collection_date, delivery_date=None):
        self.jobs = []
        self.collection_date = collection_date
        self.delivery_date = delivery_date if delivery_date else collection_date
        self.bc04_special_instructions = (
            "MUST GET A FULL NAME AND SIGNATURE ON COLLECTION CALL OFFICE AND Non Conformance Motability on 0121 788 6940 option 1 IF THEY REFUSE ** - PHOTO'S MUST BE CLEAR PLEASE. COLL AND DEL 09:00-17:00 ONLY"
        )
    def calculate_delivery_date(self, collection_date):
        try:
            import holidays
            uk_holidays = holidays.UnitedKingdom()
        except ImportError:
            class DummyHolidays:
                def __init__(self, *args, **kwargs):
                    pass
                def __contains__(self, date):
                    return False
            uk_holidays = DummyHolidays()
        delivery = collection_date
        delivery += timedelta(days=1)
        while delivery.weekday() >= 5 or delivery in uk_holidays:
            delivery += timedelta(days=1)
        return delivery
    def clean_phone_number(self, phone):
        if not phone:
            return ''
        phone = phone.strip()
        phone = re.sub(r'^(Tel|Phone|T|Telephone)[\s:.]*', '', phone, flags=re.IGNORECASE)
        phone = re.sub(r'[^\d+\s()-]', '', phone)
        phone = re.sub(r'\s+', ' ', phone).strip()
        return phone
    def is_postcode(self, line):
        postcode_pattern = r'\b([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b'
        match = re.search(postcode_pattern, line.upper())
        if match:
            postcode = match.group(1)
            postcode = re.sub(r'([A-Z]\d+[A-Z]?)(\d[A-Z]{2})', r'\1 \2', postcode)
            return postcode
        return None
    def parse_jobs(self, text):
        self.jobs = []
        job_sections = re.split(r'Job Sheet\s*\n', text)
        job_sections = [section.strip() for section in job_sections if section.strip()]
        for section in job_sections:
            if section.strip():
                job = self.parse_single_job(section)
                if job and job.get('REG NUMBER'):
                    self.jobs.append(job)
        return self.jobs
    def parse_single_job(self, job_text):
        job = {}
        job['REG NUMBER'] = ''
        job['VIN'] = ''
        job['MAKE'] = ''
        job['MODEL'] = ''
        job['COLOR'] = ''
        job['COLLECTION DATE'] = self.collection_date
        job['YOUR REF NO'] = ''
        job['COLLECTION ADDR1'] = ''
        job['COLLECTION ADDR2'] = ''
        job['COLLECTION ADDR3'] = ''
        job['COLLECTION ADDR4'] = ''
        job['COLLECTION POSTCODE'] = ''
        job['COLLECTION CONTACT NAME'] = ''
        job['COLLECTION PHONE'] = ''
        job['DELIVERY DATE'] = self.delivery_date
        job['DELIVERY ADDR1'] = ''
        job['DELIVERY ADDR2'] = ''
        job['DELIVERY ADDR3'] = ''
        job['DELIVERY ADDR4'] = ''
        job['DELIVERY POSTCODE'] = ''
        job['DELIVERY CONTACT NAME'] = ''
        job['DELIVERY CONTACT PHONE'] = ''
        job['SPECIAL INSTRUCTIONS'] = self.bc04_special_instructions
        job['PRICE'] = ''
        job['CUSTOMER REF'] = 'BC04'
        job['TRANSPORT TYPE'] = ''
        job_number_match = re.search(r'Job Number.*?(\d+/\d+)', job_text, re.DOTALL)
        if job_number_match:
            job['YOUR REF NO'] = job_number_match.group(1)
        reg_match = re.search(r'([A-Z]{2}\d{2}[A-Z]{3})', job_text)
        if reg_match:
            job['REG NUMBER'] = reg_match.group(1)
        vin_match = re.search(rf'{job["REG NUMBER"]}\s+(\d{{9,}})', job_text) if job['REG NUMBER'] else None
        if vin_match:
            job['VIN'] = vin_match.group(1)
        job['MAKE'] = ''
        job['MODEL'] = ''
        price_matches = re.findall(r'£?\s*(\d+\.\d{2})', job_text)
        if len(price_matches) >= 2:
            job['PRICE'] = price_matches[1]
        elif price_matches:
            job['PRICE'] = price_matches[0]
        lines = [line.strip() for line in job_text.split('\n')]
        addr_start = None
        reg_line_idx = None
        reg_pattern = r'^[A-Z]{2}\d{2}[A-Z]{3}\s+\d{9,}'
        for i, line in enumerate(lines):
            if line.strip().lower().startswith('special instructions'):
                addr_start = i + 1
            if re.match(reg_pattern, line.strip()):
                reg_line_idx = i
                break
        if addr_start is not None and reg_line_idx is not None and addr_start < reg_line_idx:
            address_lines = [l for l in lines[addr_start:reg_line_idx] if l.strip()]
            postcode_indices = [i for i, l in enumerate(address_lines) if self.is_postcode(l)]
            if len(postcode_indices) == 2:
                split_idx = postcode_indices[0] + 1
            else:
                split_idx = len(address_lines) // 2
            collection_lines = address_lines[:split_idx]
            delivery_lines = address_lines[split_idx:]
            if collection_lines:
                c_postcode_idx = None
                for idx, l in enumerate(collection_lines):
                    if self.is_postcode(l):
                        c_postcode_idx = idx
                        break
                if c_postcode_idx is not None and c_postcode_idx > 0:
                    c_addr = collection_lines[:c_postcode_idx]
                    c_town = c_addr[-1] if len(c_addr) >= 1 else ''
                    for i in range(3):
                        job[f'COLLECTION ADDR{i+1}'] = c_addr[i] if i < len(c_addr)-1 else ''
                    job['COLLECTION ADDR4'] = c_town
                    job['COLLECTION POSTCODE'] = collection_lines[c_postcode_idx]
                else:
                    for idx, val in enumerate(collection_lines):
                        if idx < 4:
                            job[f'COLLECTION ADDR{idx+1}'] = val
            if delivery_lines:
                d_postcode_idx = None
                for idx, l in enumerate(delivery_lines):
                    if self.is_postcode(l):
                        d_postcode_idx = idx
                        break
                if d_postcode_idx is not None and d_postcode_idx > 0:
                    d_addr = delivery_lines[:d_postcode_idx]
                    d_town = d_addr[-1] if len(d_addr) >= 1 else ''
                    for i in range(3):
                        job[f'DELIVERY ADDR{i+1}'] = d_addr[i] if i < len(d_addr)-1 else ''
                    job['DELIVERY ADDR4'] = d_town
                    job['DELIVERY POSTCODE'] = delivery_lines[d_postcode_idx]
                else:
                    for idx, val in enumerate(delivery_lines):
                        if idx < 4:
                            job[f'DELIVERY ADDR{idx+1}'] = val
        phone_line = ''
        found_dates = False
        for i, line in enumerate(lines):
            phones = re.findall(r'\d{8,}', line)
            if len(phones) >= 2:
                if i+1 < len(lines) and re.match(r'\d{2}/\d{2}/\d{4}', lines[i+1]):
                    job['COLLECTION PHONE'] = phones[0]
                    job['DELIVERY CONTACT PHONE'] = phones[1]
                    date_matches = []
                    for l in lines[i+1:i+5]:
                        date_matches += re.findall(r'\d{2}/\d{2}/\d{4}', l)
                        if len(date_matches) >= 2:
                            break
        return job 