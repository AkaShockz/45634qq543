import re
import csv
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import holidays
import os
from tkinter import font as tkfont
import pandas as pd
import tempfile
import threading
import time
import schedule
import json

# Gmail API imports for email automation
try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    import base64
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    GMAIL_AVAILABLE = True
except ImportError:
    GMAIL_AVAILABLE = False

# Load environment variables for email automation
try:
    from dotenv import load_dotenv
    load_dotenv()
    ENV_AVAILABLE = True
except ImportError:
    ENV_AVAILABLE = False

# Set up debug logging
DEBUG_LOG = os.path.join(tempfile.gettempdir(), "debug_log.txt")

def log_debug(message):
    """Write debug message to file in a safe way"""
    try:
        with open(DEBUG_LOG, 'a') as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        # If logging fails, silently ignore to avoid crashing the app
        pass

class RoundedButton(tk.Frame):
    """A simpler button implementation that works with standard tkinter"""
    def __init__(self, parent, width, height, cornerradius, padding, color, text, command=None, fg="white", hover_color=None):
        tk.Frame.__init__(self, parent, width=width, height=height)
        
        self.command = command
        self.fg = fg
        self.color = color
        self.hover_color = hover_color if hover_color else self._get_darker_color(color, 20)
        
        # Don't resize frame to match button
        self.pack_propagate(False)
        
        # Create the actual button
        self.button = tk.Button(
            self, 
            text=text, 
            font=('Segoe UI', 10, 'bold'),
            bg=color,
            fg=fg,
            activebackground=self.hover_color,
            activeforeground=fg,
            relief=tk.FLAT,
            bd=0,
            padx=padding,
            pady=padding,
            command=self.on_click,
            cursor="hand2"
        )
        self.button.pack(fill=tk.BOTH, expand=True)
        
        # Bind hover events
        self.button.bind("<Enter>", self._on_enter)
        self.button.bind("<Leave>", self._on_leave)
    
    def _get_darker_color(self, hex_color, percent):
        """Return a darker shade of the given color in hex format"""
        # Remove the '#' if present
        hex_color = hex_color.lstrip('#')
        
        # Convert hex to RGB
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        
        # Reduce each component by the given percentage
        factor = 1 - percent/100
        r = max(0, int(r * factor))
        g = max(0, int(g * factor))
        b = max(0, int(b * factor))
        
        # Convert back to hex
        return f'#{r:02x}{g:02x}{b:02x}'
    
    def _on_enter(self, event):
        """Handle mouse enter event"""
        self.button.config(bg=self.hover_color)
    
    def _on_leave(self, event):
        """Handle mouse leave event"""
        self.button.config(bg=self.color)
    
    def on_click(self):
        """Handle button click"""
        if self.command:
            self.command()


class ModernFrame(tk.Frame):
    """A simplified frame with a border for better compatibility"""
    def __init__(self, parent, bordercolor="#e0e0e0", bgcolor="#ffffff", borderwidth=1, **kwargs):
        tk.Frame.__init__(self, parent, bg=bordercolor, highlightthickness=0, bd=0, **kwargs)
        
        # Create the primary content frame with given background color
        self.content_frame = tk.Frame(self, bg=bgcolor, bd=0, highlightthickness=0)
        self.content_frame.pack(fill="both", expand=True, padx=borderwidth, pady=borderwidth)


class TabButton(tk.Frame):
    """A simpler tab button implementation for better compatibility"""
    def __init__(self, parent, text, width=120, height=40, command=None, selected=False, 
                 bg="#FFFFFF", fg="#333333", selected_fg="#4361EE", hover_bg="#F5F5F5", 
                 selected_bg="#FFFFFF"):
        tk.Frame.__init__(self, parent, width=width, height=height, bg=bg,
                         highlightthickness=0, bd=0)
        self.pack_propagate(False)  # Don't shrink to fit content
        
        self.text = text
        self.command = command
        self.width = width
        self.height = height
        self.bg = bg
        self.fg = fg
        self.selected_fg = selected_fg
        self.hover_bg = hover_bg
        self.selected_bg = selected_bg
        self.selected = selected
        
        # Create the actual elements
        self.container = tk.Frame(self, bg=bg)
        self.container.pack(fill=tk.BOTH, expand=True)
        
        # Create label for text
        self.label = tk.Label(
            self.container, 
            text=text,
            bg=bg,
            fg=selected_fg if selected else fg,
            font=('Segoe UI', 10, 'bold' if selected else 'normal')
        )
        self.label.pack(fill=tk.BOTH, expand=True)
        
        # Create indicator
        self.indicator = tk.Frame(
            self, 
            height=3, 
            bg=selected_fg if selected else bg
        )
        self.indicator.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Bind events
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", self._on_click)
        self.label.bind("<Enter>", self._on_enter)
        self.label.bind("<Leave>", self._on_leave)
        self.label.bind("<Button-1>", self._on_click)
        
    def set_selected(self, selected):
        """Set the selected state of the tab button"""
        self.selected = selected
        if selected:
            self.label.config(fg=self.selected_fg, font=('Segoe UI', 10, 'bold'))
            self.indicator.config(bg=self.selected_fg)
            self.config(bg=self.selected_bg)
            self.container.config(bg=self.selected_bg)
            self.label.config(bg=self.selected_bg)
        else:
            self.label.config(fg=self.fg, font=('Segoe UI', 10, 'normal'))
            self.indicator.config(bg=self.bg)
            current_bg = self.hover_bg if hasattr(self, 'hovering') and self.hovering else self.bg
            self.config(bg=current_bg)
            self.container.config(bg=current_bg)
            self.label.config(bg=current_bg)
    
    def _on_enter(self, event):
        """Handle mouse enter event"""
        self.hovering = True
        if not self.selected:
            self.config(bg=self.hover_bg)
            self.container.config(bg=self.hover_bg)
            self.label.config(bg=self.hover_bg)
    
    def _on_leave(self, event):
        """Handle mouse leave event"""
        self.hovering = False
        if not self.selected:
            self.config(bg=self.bg)
            self.container.config(bg=self.bg)
            self.label.config(bg=self.bg)
    
    def _on_click(self, event):
        """Handle mouse click event"""
        if self.command is not None:
            self.command(self.text)


class JobParser:
    def __init__(self, collection_date, delivery_date=None):
        self.jobs = []
        self.collection_date = collection_date
        self.delivery_date = delivery_date if delivery_date else collection_date
        
    def calculate_delivery_date(self, collection_date):
        """Calculate delivery date as 3 business days from collection date."""
        # Convert string date to datetime if needed
        if isinstance(collection_date, str):
            collection_date = datetime.strptime(collection_date, "%d/%m/%Y")
            
        # Get UK holidays
        uk_holidays = holidays.UK()
        
        # Start with collection date
        current_date = collection_date
        business_days = 0
        
        # Keep adding days until we have 3 business days
        while business_days < 3:
            current_date += timedelta(days=1)
            # Skip weekends and holidays
            if current_date.weekday() < 5 and current_date not in uk_holidays:
                business_days += 1
                
        # Format the date back to string
        return current_date.strftime("%d/%m/%Y")
        
    def fix_location_name(self, name):
        # Fix common location names
        name = name.replace('18 AC Stoke Logistics Hub', '18 Arnold Clark Stoke Logistics Hub')
        name = name.replace('4 AC Accrington Logistics Hub', '4 Arnold Clark Accrington Logistics Hub')
        
        # Handle "Unit" in address lines - fix for Wakefield Motorstore
        if "Unit 1 Calder Park Services" in name:
            return "Wakefield Motorstore"
        
        # Handle cases where Unit appears in other locations
        if name.startswith("Unit ") and len(name.split()) >= 3:
            # This is likely a unit number that should be in ADDR2
            # Don't modify here - the whole address structure would need to change
            pass
            
        return name
        
    def clean_phone_number(self, phone):
        """Clean and format phone number."""
        if not phone:
            return ""
            
        # Remove any non-digit characters
        digits = ''.join(c for c in phone if c.isdigit())
        
        # Handle various formats
        if digits.startswith('44'):
            # Remove 44 prefix
            digits = digits[2:]
        elif digits.startswith('0044'):
            # Remove 0044 prefix
            digits = digits[4:]
        elif digits.startswith('0'):
            # Remove leading 0
            digits = digits[1:]
            
        # Ensure proper length (10 digits)
        if len(digits) > 10:
            digits = digits[:10]
        elif len(digits) < 10:
            # If less than 10 digits, pad with zeros at the end
            digits = digits.ljust(10, '0')
                
        return digits

    def is_postcode(self, line):
        # More flexible UK postcode regex that handles common variations
        postcode_patterns = [
            # Standard format: AA9A 9AA, AA99 9AA, A9A 9AA, A99 9AA, AA9 9AA
            r'^[A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2}$',
            # With optional spaces and lowercase
            r'^[A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2}$',
            # With "Postcode:" prefix
            r'^(?:Postcode|Post Code|P/Code|PC)[\s:]+[A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2}$'
        ]
        
        line = line.strip()
        for pattern in postcode_patterns:
            if re.match(pattern, line):
                # Extract just the postcode part if it has a prefix
                postcode_match = re.search(r'([A-Za-z]{1,2}[0-9][0-9A-Za-z]?\s*[0-9][A-Za-z]{2})$', line)
                if postcode_match:
                    return postcode_match.group(1).upper()
        return None

    def parse_jobs(self, text):
        # Split by FROM sections
        job_texts = re.split(r'\nFROM\n', text)
        job_texts = [t for t in job_texts if t.strip()]
        
        for job_text in job_texts:
            if not job_text.startswith('FROM'):
                job_text = 'FROM\n' + job_text
            
            if not re.search(r'TO\n', job_text):
                continue
                
            job = self.parse_single_job(job_text)
            if job:
                # Ensure special instructions are set for each job
                if 'SPECIAL INSTRUCTIONS' not in job or not job['SPECIAL INSTRUCTIONS']:
                    job['SPECIAL INSTRUCTIONS'] = 'Please call 1 hour before collection'
                self.jobs.append(job)
            
        return self.jobs
    
    def parse_address_lines(self, lines):
        """Helper method to parse address lines while preserving St. and similar abbreviations"""
        # Preserve common street name patterns before splitting
        preserved_patterns = [
            (r'St\.\s+[A-Z][a-z]+', lambda m: m.group().replace('.', '@')),  # St. Andrews, St. Mary's etc
            (r'St\s+[A-Z][a-z]+', lambda m: m.group().replace(' ', '#')),    # St Andrews, St Mary's etc
            (r'D\.\s*M\.\s*Keith', lambda m: m.group().replace('.', '@')),   # D.M.Keith
            (r'[A-Z]\.\s+[A-Z]\.\s+\w+', lambda m: m.group().replace('.', '@')),  # Any X.Y. format names
        ]
        
        processed_lines = []
        for line in lines:
            if not line.strip():
                continue
                
            # Preserve special patterns
            processed_line = line
            for pattern, replacement in preserved_patterns:
                processed_line = re.sub(pattern, replacement, processed_line)
            
            # Now restore the preserved patterns
            processed_line = processed_line.replace('@', '.').replace('#', ' ')
            processed_lines.append(processed_line.strip())
            
        return processed_lines

    def clean_duplicate_towns(self, lines):
        """Remove duplicate consecutive town names while preserving the last occurrence"""
        if not lines:
            return lines
        
        cleaned_lines = []
        i = 0
        while i < len(lines):
            # If we're at the last line or current line is different from next line
            if i == len(lines) - 1 or lines[i].strip().upper() != lines[i + 1].strip().upper():
                cleaned_lines.append(lines[i])
                i += 1
            else:
                # Skip the first occurrence of duplicate town
                i += 1
        return cleaned_lines

    def parse_single_job(self, job_text):
        job = {}
        
        # Initialize all fields to empty strings
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
        job['SPECIAL INSTRUCTIONS'] = 'Must call 1hour before collection and get a name'
        job['PRICE'] = ''
        job['CUSTOMER REF'] = 'AC01'
        job['TRANSPORT TYPE'] = ''
        
        # Extract FROM section
        from_match = re.search(r'FROM\n(.*?)(?=\nTO|$)', job_text, re.DOTALL)
        if from_match:
            from_text = from_match.group(1).strip()
            from_lines = [line.strip() for line in from_text.split('\n') if line.strip()]
            
            # Process phone number first - Updated pattern to better match phone numbers
            phone_lines = [line for line in from_lines if re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+[+\d()\s-]+', line, re.IGNORECASE)]
            if phone_lines:
                phone_match = re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+([+\d()\s-]+)', phone_lines[0], re.IGNORECASE)
                if phone_match:
                    phone_number = phone_match.group(1).strip()
                    # Clean and store the phone number
                    job['COLLECTION PHONE'] = self.clean_phone_number(phone_number)
                # Remove phone lines from address processing
                from_lines = [line for line in from_lines if line not in phone_lines]
            
            # Process remaining lines for address
            address_lines = []
            postcode = None
            
            for line in from_lines:
                # Check if line is a postcode
                postcode_value = self.is_postcode(line)
                if postcode_value:
                    postcode = postcode_value
                    continue
                # Add to address lines if not a postcode
                address_lines.append(line)
            
            # Process address lines with special handling for St. names
            address_lines = self.parse_address_lines(address_lines)
            
            # Remove duplicate town names
            address_lines = self.clean_duplicate_towns(address_lines)
            
            # Process address lines
            if address_lines:
                # First line is always ADDR1
                job['COLLECTION ADDR1'] = self.fix_location_name(address_lines[0])
                
                # Handle remaining address lines
                remaining_lines = address_lines[1:]
                
                # Assign remaining lines
                if len(remaining_lines) > 0:
                    # Always assign the second line to ADDR2 if it exists
                    if len(remaining_lines) >= 1:
                        job['COLLECTION ADDR2'] = remaining_lines[0]
                    
                    # For the rest of the lines
                    if len(remaining_lines) == 2:
                        job['COLLECTION ADDR3'] = ''  # Ensure ADDR3 is empty
                        job['COLLECTION ADDR4'] = remaining_lines[1]  # Town
                    elif len(remaining_lines) >= 3:
                        job['COLLECTION ADDR3'] = remaining_lines[1]
                        job['COLLECTION ADDR4'] = remaining_lines[-1]  # Town
                    
                    # Ensure ADDR4 is not empty if we have remaining lines
                    if not job['COLLECTION ADDR4'] and remaining_lines:
                        job['COLLECTION ADDR4'] = remaining_lines[-1]
            
            if postcode:
                job['COLLECTION POSTCODE'] = postcode
        
        # Extract TO section
        to_match = re.search(r'TO\n(.*?)(?=\nJOB NO|$)', job_text, re.DOTALL)
        if to_match:
            to_text = to_match.group(1).strip()
            to_lines = [line.strip() for line in to_text.split('\n') if line.strip()]
            
            # Process phone number first - Updated pattern to better match phone numbers
            phone_lines = [line for line in to_lines if re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+[+\d()\s-]+', line, re.IGNORECASE)]
            if phone_lines:
                phone_match = re.search(r'(?:Tel|Phone|T|Telephone)[\s:.]+([+\d()\s-]+)', phone_lines[0], re.IGNORECASE)
                if phone_match:
                    phone_number = phone_match.group(1).strip()
                    # Clean and store the phone number
                    job['DELIVERY CONTACT PHONE'] = self.clean_phone_number(phone_number)
                # Remove phone lines from address processing
                to_lines = [line for line in to_lines if line not in phone_lines]
            
            # Process remaining lines for address
            address_lines = []
            postcode = None
            
            for line in to_lines:
                # Check if line is a postcode
                postcode_value = self.is_postcode(line)
                if postcode_value:
                    postcode = postcode_value
                    continue
                # Add to address lines if not a postcode
                address_lines.append(line)
            
            # Process address lines with special handling for St. names
            address_lines = self.parse_address_lines(address_lines)
            
            # Remove duplicate town names
            address_lines = self.clean_duplicate_towns(address_lines)
            
            # Process address lines
            if address_lines:
                # First line is always ADDR1
                job['DELIVERY ADDR1'] = self.fix_location_name(address_lines[0])
                
                # Special handling for Wakefield Motorstore
                if job['DELIVERY ADDR1'] == "Wakefield Motorstore":
                    # Check if "Peel Avenue" exists in the address lines
                    has_peel_avenue = False
                    for line in address_lines:
                        if "Peel Avenue" in line:
                            has_peel_avenue = True
                            # Insert Unit 1 Calder Park Services as ADDR2
                            job['DELIVERY ADDR2'] = "Unit 1 Calder Park Services"
                            job['DELIVERY ADDR3'] = "Peel Avenue"
                            job['DELIVERY ADDR4'] = "Wakefield"
                            # Set the correct phone number for Wakefield Motorstore
                            job['DELIVERY CONTACT PHONE'] = '01924975790'
                            # Don't return early - continue processing to handle other fields
                            break
                
                # Handle remaining address lines
                remaining_lines = address_lines[1:]
                
                # Only process remaining lines if we haven't already set them for Wakefield
                if not (job['DELIVERY ADDR1'] == "Wakefield Motorstore" and 
                        job['DELIVERY ADDR2'] == "Unit 1 Calder Park Services"):
                    # Assign remaining lines
                    if len(remaining_lines) > 0:
                        # Always assign the second line to ADDR2 if it exists
                        if len(remaining_lines) >= 1:
                            job['DELIVERY ADDR2'] = remaining_lines[0]
                        
                        # For the rest of the lines
                        if len(remaining_lines) == 2:
                            job['DELIVERY ADDR3'] = ''  # Ensure ADDR3 is empty
                            job['DELIVERY ADDR4'] = remaining_lines[1]  # Town
                        elif len(remaining_lines) >= 3:
                            job['DELIVERY ADDR3'] = remaining_lines[1]
                            job['DELIVERY ADDR4'] = remaining_lines[-1]  # Town
                        
                        # Ensure ADDR4 is not empty if we have remaining lines
                        if not job['DELIVERY ADDR4'] and remaining_lines:
                            job['DELIVERY ADDR4'] = remaining_lines[-1]
            
            if postcode:
                job['DELIVERY POSTCODE'] = postcode
        
        # Extract vehicle details - better pattern for the format in the provided example
        vehicle_section_match = re.search(r'MAKE\s+MODEL\s+COLOU?R\s+REGISTRATION(?:\s+CHASSIS)?\s*\n(.*?)(?=\n\s*COMMENTS|\n\s*ORIGIN|\n\s*VALUE|$)', 
                                         job_text, re.DOTALL | re.IGNORECASE)
        
        if vehicle_section_match:
            vehicle_line = vehicle_section_match.group(1).strip()
            
            # Try to match with a more structured pattern first - chassis is optional
            vehicle_pattern = r'(\w+(?:-\w+)?)\s+([\w\s]+?)\s+(Blue|Grey|Black|White|Red|Silver|Green|Yellow|Orange|Purple|Brown|Gold)\s+([A-Z0-9]+)(?:\s+([A-Z0-9]*))?'
            structured_match = re.match(vehicle_pattern, vehicle_line, re.IGNORECASE)
            
            if structured_match:
                make = structured_match.group(1).strip()
                model = structured_match.group(2).strip()
                color = structured_match.group(3).strip()
                registration = structured_match.group(4).strip()
                # Chassis might be empty, that's okay
                chassis = structured_match.group(5).strip() if structured_match.group(5) else ''
                
                job['MAKE'] = make
                job['MODEL'] = model
                job['COLOR'] = color
                if self.is_valid_uk_registration(registration):
                    job['REG NUMBER'] = registration
                else:
                    job['REG NUMBER'] = ''
                job['VIN'] = chassis
            else:
                # Split the vehicle line into parts - fallback method
                parts = vehicle_line.split()
                
                if len(parts) >= 4:  # We need at least make, model, color, and registration
                    # Start from the end and work backwards
                    if len(parts) >= 5 and re.match(r'^[A-Z0-9]+$', parts[-1]):
                        # Last part looks like a chassis number
                        chassis = parts[-1]
                        registration = parts[-2]
                        color_index = -3
                    else:
                        # No chassis number
                        chassis = ''
                        registration = parts[-1]
                        color_index = -2
                    
                    # Color is typically one word
                    color = parts[color_index] if abs(color_index) < len(parts) else ''
                    
                    # Make is usually one word, but can be hyphenated
                    common_makes = ["AUDI", "BMW", "CUPRA", "DACIA", "FORD", "HYUNDAI", "KIA", "MG", 
                                    "NISSAN", "PEUGEOT", "RENAULT", "SEAT", "SKODA", "TOYOTA", 
                                    "VAUXHALL", "VOLKSWAGEN", "VOLVO", "MERCEDES-BENZ"]
                    
                    # Try to identify the make
                    make = parts[0]
                    model_start = 1
                    
                    # Handle hyphenated makes like MERCEDES-BENZ
                    if len(parts) > 4 and parts[0] + "-" + parts[1] in [m.upper() for m in common_makes]:
                        make = parts[0] + "-" + parts[1]
                        model_start = 2
                    
                    # The model is everything between make and color
                    model = " ".join(parts[model_start:color_index])
                    
                    job['MAKE'] = make
                    job['MODEL'] = model
                    job['COLOR'] = color
                    if self.is_valid_uk_registration(registration):
                        job['REG NUMBER'] = registration
                    else:
                        job['REG NUMBER'] = ''
                    job['VIN'] = chassis
        
        # Look for individual fields if not found above
        if not job['MAKE']:
            make_match = re.search(r'MAKE\s*:\s*([A-Z-]+(?:-[A-Z]+)?)', job_text, re.IGNORECASE)
            if make_match:
                job['MAKE'] = make_match.group(1).strip()
                
        if not job['MODEL']:
            model_match = re.search(r'MODEL\s*:\s*([A-Z0-9\s]+?)(?:\n|$)', job_text, re.IGNORECASE)
            if model_match:
                job['MODEL'] = model_match.group(1).strip()
                
        if not job['COLOR']:
            color_match = re.search(r'COLOU?R\s*:\s*([A-Za-z]+)', job_text, re.IGNORECASE)
            if color_match:
                job['COLOR'] = color_match.group(1).strip()
                
        if not job['REG NUMBER']:
            reg_match = re.search(r'REG(?:ISTRATION)?\s*:?\s*([A-Z0-9]+)', job_text, re.IGNORECASE)
            if reg_match:
                reg_candidate = reg_match.group(1).strip()
                if self.is_valid_uk_registration(reg_candidate):
                    job['REG NUMBER'] = reg_candidate
                else:
                    job['REG NUMBER'] = ''
        
        if not job['VIN']:
            vin_match = re.search(r'CHASSIS\s*:?\s*([A-Z0-9]+)', job_text, re.IGNORECASE)
            if vin_match:
                job['VIN'] = vin_match.group(1).strip()
        
        # Extract job number - specific pattern for "JOB NO KEY LOC BARCODE" format
        job_no_match = re.search(r'JOB\s+NO(?:\s+KEY\s+LOC\s+BARCODE)?\s*\n(\d+)', job_text, re.IGNORECASE)
        if job_no_match:
            job['YOUR REF NO'] = job_no_match.group(1).strip()
        else:
            # Fallback to other patterns if the above doesn't match
            job_no_patterns = [
                r'JOB\s+NO[\s:.]+(\d+)',
                r'YOUR_REF.*?\n(\d+)',
                r'REF(?:ERENCE)?(?:\s+NO)?[\s:.]+(\d+)',
                r'REFERENCE[\s:.]+(\d+)'
            ]
            
            for pattern in job_no_patterns:
                job_no_match = re.search(pattern, job_text, re.IGNORECASE)
                if job_no_match:
                    job['YOUR REF NO'] = job_no_match.group(1).strip()
                    break
        
        # Extract value/price - specific pattern for the example format
        value_match = re.search(r'VALUE\s*\n([\d.]+)', job_text, re.IGNORECASE)
        if value_match:
            # We don't store the price as requested
            job['PRICE'] = ''
        
        return job

    def is_valid_uk_registration(self, reg):
        reg = reg.upper().replace(" ", "")
        if reg.isdigit():
            return False  # Never just numbers
        patterns = [
            r"^[A-Z]{2}[0-9]{2}[A-Z]{3}$",         # Current style: AB12CDE
            r"^[A-Z]{1}[0-9]{1,3}[A-Z]{3}$",       # Prefix: A123BCD
            r"^[A-Z]{3}[0-9]{1,3}[A-Z]{1}$",       # Suffix: ABC123A
            r"^[A-Z]{3}[0-9]{4}$",                 # NI: BYZ3210
            r"^[0-9]{1,4}[A-Z]{1,3}$",             # Dateless: 1–4 numbers + 1–3 letters
            r"^[A-Z]{1,3}[0-9]{1,4}$",             # Dateless: 1–3 letters + 1–4 numbers
        ]
        return any(re.match(p, reg) for p in patterns)

class EmailAutomation:
    """Email automation functionality integrated into the parser"""
    
    def __init__(self, app_instance):
        self.app = app_instance
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
                      'https://www.googleapis.com/auth/gmail.send']
        self.creds = None
        self.service = None
        self.sender_email = os.getenv('SENDER_EMAIL') if ENV_AVAILABLE else None
        self.sender_password = os.getenv('SENDER_PASSWORD') if ENV_AVAILABLE else None
        self.is_running = False
        self.automation_thread = None
        
        # Job type mappings
        self.job_types = {
            'ac01': 'AC01',
            'bc04': 'BC04', 
            'gr11': 'GR11',
            'cw09': 'CW09',
            'eu01': 'EU01'
        }
        
    def authenticate_gmail(self):
        """Authenticate with Gmail API, robust error handling"""
        print(f"DEBUG: Looking for credentials.json at: {CREDENTIALS_PATH}")
        print(f"DEBUG: Looking for token.json at: {TOKEN_PATH}")
        if not os.path.exists(CREDENTIALS_PATH):
            print("DEBUG: credentials.json does not exist.")
            return False, f"credentials.json not found at {CREDENTIALS_PATH}"
        try:
            with open(CREDENTIALS_PATH, 'r') as f:
                try:
                    json.load(f)
                except json.JSONDecodeError:
                    print("DEBUG: credentials.json is not valid JSON.")
                    return False, f"credentials.json is not valid JSON at {CREDENTIALS_PATH}"
        except PermissionError:
            print("DEBUG: Permission denied when opening credentials.json.")
            return False, f"Permission denied for credentials.json at {CREDENTIALS_PATH}"
        except Exception as e:
            print(f"DEBUG: Unexpected error opening credentials.json: {e}")
            return False, f"Unexpected error opening credentials.json: {e}"
        try:
            if os.path.exists(TOKEN_PATH):
                self.creds = Credentials.from_authorized_user_file(TOKEN_PATH, self.SCOPES)
            if not self.creds or not self.creds.valid:
                if self.creds and self.creds.expired and self.creds.refresh_token:
                    self.creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, self.SCOPES)
                    self.creds = flow.run_local_server(port=0)
                with open(TOKEN_PATH, 'w') as token:
                    token.write(self.creds.to_json())
            self.service = build('gmail', 'v1', credentials=self.creds)
            print("DEBUG: Gmail authentication successful!")
            return True, "Gmail authentication successful!"
        except Exception as e:
            print(f"DEBUG: Gmail authentication failed: {e}")
            return False, f"Gmail authentication failed: {e}"
    
    def get_unread_emails(self, query: str = "is:unread"):
        """Get unread emails matching the query"""
        if not self.service:
            return []
            
        try:
            results = self.service.users().messages().list(
                userId='me', q=query).execute()
            messages = results.get('messages', [])
            
            emails = []
            for message in messages:
                msg = self.service.users().messages().get(
                    userId='me', id=message['id']).execute()
                emails.append(msg)
            
            return emails
        except HttpError as error:
            log_debug(f"Error getting emails: {error}")
            return []
    
    def parse_email_content(self, email_data):
        """Parse email content to extract job information"""
        try:
            headers = email_data['payload']['headers']
            subject = next((h['value'] for h in headers if h['name'] == 'Subject'), '')
            from_email = next((h['value'] for h in headers if h['name'] == 'From'), '')
            
            # Extract email body
            body = self._get_email_body(email_data['payload'])
            
            # Parse job type from subject or body
            job_type = self._extract_job_type(subject, body)
            
            # Extract dates
            collection_date, delivery_date = self._extract_dates(body)
            
            # Get attachments
            attachments = self._get_attachments(email_data['payload'])
            
            return {
                'id': email_data['id'],
                'subject': subject,
                'from_email': from_email,
                'body': body,
                'job_type': job_type,
                'collection_date': collection_date,
                'delivery_date': delivery_date,
                'attachments': attachments
            }
            
        except Exception as e:
            log_debug(f"Error parsing email: {e}")
            return {}
    
    def _get_email_body(self, payload):
        """Extract email body from payload"""
        if 'body' in payload and payload['body'].get('data'):
            return base64.urlsafe_b64decode(
                payload['body']['data']).decode('utf-8')
        
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    if 'data' in part['body']:
                        return base64.urlsafe_b64decode(
                            part['body']['data']).decode('utf-8')
        
        return ""
    
    def _extract_job_type(self, subject, body):
        """Extract job type from email subject or body"""
        text = f"{subject} {body}".lower()
        
        for key, value in self.job_types.items():
            if key in text:
                return value
        
        # Default to AC01 if no specific type found
        return 'AC01'
    
    def _extract_dates(self, body):
        """Extract collection and delivery dates from email body"""
        # Date patterns
        date_patterns = [
            r'(\d{1,2}/\d{1,2}/\d{4})',
            r'(\d{1,2}-\d{1,2}-\d{4})',
            r'(\d{4}-\d{1,2}-\d{1,2})'
        ]
        
        dates = []
        for pattern in date_patterns:
            dates.extend(re.findall(pattern, body))
        
        if len(dates) >= 2:
            return dates[0], dates[1]
        elif len(dates) == 1:
            return dates[0], None
        else:
            return None, None
    
    def _get_attachments(self, payload):
        """Extract file attachments from email"""
        attachments = []
        
        def extract_from_parts(parts):
            for part in parts:
                if part.get('filename'):
                    attachment = {
                        'id': part['body']['attachmentId'],
                        'filename': part['filename'],
                        'mimeType': part['mimeType']
                    }
                    attachments.append(attachment)
                elif 'parts' in part:
                    extract_from_parts(part['parts'])
        
        if 'parts' in payload:
            extract_from_parts(payload['parts'])
        
        return attachments
    
    def download_attachment(self, attachment_id, filename):
        """Download attachment from Gmail"""
        try:
            attachment = self.service.users().messages().attachments().get(
                userId='me', messageId=attachment_id, id=attachment_id).execute()
            
            file_data = base64.urlsafe_b64decode(attachment['data'])
            file_path = os.path.join(tempfile.gettempdir(), filename)
            
            with open(file_path, 'wb') as f:
                f.write(file_data)
            
            return file_path
            
        except Exception as e:
            log_debug(f"Error downloading attachment: {e}")
            return None
    
    def process_job(self, job_info):
        """Process the job using existing parser logic"""
        try:
            # Download attachments
            file_paths = []
            for attachment in job_info['attachments']:
                file_path = self.download_attachment(
                    attachment['id'], attachment['filename'])
                if file_path:
                    file_paths.append(file_path)
            
            if not file_paths:
                return None
            
            # Process based on job type
            if job_info['job_type'] == 'AC01':
                return self._process_ac01_job(job_info, file_paths)
            elif job_info['job_type'] == 'BC04':
                return self._process_bc04_job(job_info, file_paths)
            elif job_info['job_type'] == 'GR11':
                return self._process_gr11_job(job_info, file_paths)
            else:
                return self._process_generic_job(job_info, file_paths)
                
        except Exception as e:
            log_debug(f"Error processing job: {e}")
            return None
    
    def _process_ac01_job(self, job_info, file_paths):
        """Process AC01 job type"""
        # Read file content
        with open(file_paths[0], 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse jobs using existing parser
        parser = JobParser(
            collection_date=job_info['collection_date'],
            delivery_date=job_info['delivery_date']
        )
        jobs = parser.parse_jobs(content)
        
        # Save to CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"ac01_jobs_{timestamp}.csv"
        output_path = os.path.join('jobs', 'AC01', output_filename)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Job Number', 'Collection Date', 'Delivery Date', 'From', 'To', 'Vehicle', 'Phone']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for job in jobs:
                writer.writerow(job)
        
        return output_path
    
    def _process_bc04_job(self, job_info, file_paths):
        """Process BC04 job type"""
        # Read file content
        with open(file_paths[0], 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse jobs using existing parser
        parser = BC04Parser(
            collection_date=job_info['collection_date'],
            delivery_date=job_info['delivery_date']
        )
        jobs = parser.parse_jobs(content)
        
        # Save to CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"bc04_jobs_{timestamp}.csv"
        output_path = os.path.join('jobs', 'BC04', output_filename)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Job Number', 'Collection Date', 'Delivery Date', 'From', 'To', 'Vehicle', 'Phone']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for job in jobs:
                writer.writerow(job)
        
        return output_path
    
    def _process_gr11_job(self, job_info, file_paths):
        """Process GR11 job type (Excel files)"""
        # Read Excel file
        df = pd.read_excel(file_paths[0])
        
        # Process Excel data (simplified - you may need to adapt this)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"gr11_jobs_{timestamp}.csv"
        output_path = os.path.join('jobs', 'GR11', output_filename)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Convert to CSV
        df.to_csv(output_path, index=False)
        
        return output_path
    
    def _process_generic_job(self, job_info, file_paths):
        """Process generic job type"""
        # Default processing for unknown job types
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"processed_{timestamp}.csv"
        output_path = os.path.join('jobs', 'processed', output_filename)
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Simple conversion to CSV
        with open(file_paths[0], 'r', encoding='utf-8') as f:
            content = f.read()
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        return output_path
    
    def send_result_email(self, recipient_email, result_file, job_info):
        """Send processed result back to the requester"""
        if not self.sender_email or not self.sender_password:
            log_debug("Email credentials not configured")
            return False
            
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email
            msg['To'] = recipient_email
            msg['Subject'] = f"Processed Job Results - {job_info['job_type']}"
            
            body = f"""
            Hello,
            
            Your job request has been processed successfully.
            
            Job Type: {job_info['job_type']}
            Collection Date: {job_info['collection_date']}
            Delivery Date: {job_info['delivery_date']}
            
            Please find the processed results attached.
            
            Best regards,
            Vehicle Transport Parser
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # Attach the result file
            with open(result_file, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(result_file)}'
            )
            msg.attach(part)
            
            # Send email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.sender_email, self.sender_password)
            text = msg.as_string()
            server.sendmail(self.sender_email, recipient_email, text)
            server.quit()
            
            return True
            
        except Exception as e:
            log_debug(f"Error sending result email: {e}")
            return False
    
    def mark_email_as_read(self, email_id):
        """Mark email as read"""
        if not self.service:
            return
            
        try:
            self.service.users().messages().modify(
                userId='me', id=email_id, body={'removeLabelIds': ['UNREAD']}).execute()
        except Exception as e:
            log_debug(f"Error marking email as read: {e}")
    
    def process_pending_emails(self):
        """Main function to process all pending emails"""
        if not self.service:
            return
            
        log_debug(f"Checking for new emails at {datetime.now()}")
        
        # Get unread emails
        emails = self.get_unread_emails()
        
        for email_data in emails:
            try:
                # Parse email content
                job_info = self.parse_email_content(email_data)
                
                if not job_info:
                    continue
                
                log_debug(f"Processing job from {job_info['from_email']} - Type: {job_info['job_type']}")
                
                # Process the job
                result_file = self.process_job(job_info)
                
                if result_file:
                    # Send result back
                    if self.send_result_email(
                        job_info['from_email'], 
                        result_file, 
                        job_info
                    ):
                        # Mark email as read
                        self.mark_email_as_read(email_data['id'])
                        log_debug(f"Successfully processed job from {job_info['from_email']}")
                    else:
                        log_debug(f"Failed to send result to {job_info['from_email']}")
                else:
                    log_debug(f"Failed to process job from {job_info['from_email']}")
                    
            except Exception as e:
                log_debug(f"Error processing email {email_data['id']}: {e}")
    
    def start_automation(self, check_interval_minutes=5):
        """Start the automation in a separate thread"""
        if self.is_running:
            return False, "Automation is already running"
        
        # Authenticate first
        success, message = self.authenticate_gmail()
        if not success:
            return False, message
        
        self.is_running = True
        
        def automation_loop():
            schedule.every(check_interval_minutes).minutes.do(self.process_pending_emails)
            
            while self.is_running:
                schedule.run_pending()
                time.sleep(60)  # Check every minute
        
        self.automation_thread = threading.Thread(target=automation_loop, daemon=True)
        self.automation_thread.start()
        
        return True, f"Automation started (checking every {check_interval_minutes} minutes)"
    
    def stop_automation(self):
        """Stop the automation"""
        self.is_running = False
        if self.automation_thread:
            self.automation_thread.join(timeout=5)
        return True, "Automation stopped"
    
    def run_once(self):
        """Run the automation once"""
        success, message = self.authenticate_gmail()
        if not success:
            return False, message
        
        self.process_pending_emails()
        return True, "Processed pending emails"


class VehicleTransportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Vehicle Transport Batch Processor")
        self.root.geometry("1100x800")
        
        # Price tracking for AC01 jobs
        self.total_ac01_price = 0.0
        self.ac01_price_var = tk.StringVar()
        self.ac01_price_var.set("Total Price: £0.00")
        
        # Set modern color scheme
        self.primary_color = "#4361EE"          # Blue accent color
        self.secondary_color = "#3A0CA3"        # Darker blue for hover
        self.bg_color = "#F7F7F9"               # Light background
        self.card_bg = "#FFFFFF"                # White card background
        self.text_color = "#333333"             # Dark text
        self.light_text = "#666666"             # Light text for secondary content
        self.border_color = "#E0E0E0"           # Light border color
        self.success_color = "#4CAF50"          # Green for success messages
        self.warning_color = "#FF9800"          # Orange for warnings
        self.error_color = "#F44336"            # Red for errors
        
        # Tab colors
        self.tab_bg = "#F5F7FA"                # Background of tabs area
        self.tab_hover = "#EAEEF2"             # Hover color for tabs
        self.tab_active = "#FFFFFF"            # Active tab background
        
        # Configure app-wide fonts
        self.title_font = ('Segoe UI', 18, 'bold')
        self.header_font = ('Segoe UI', 14)
        self.normal_font = ('Segoe UI', 10)
        self.small_font = ('Segoe UI', 9)
        
        # Configure the root window
        self.root.configure(bg=self.bg_color)
        
        # Create main container
        self.main_frame = tk.Frame(root, bg=self.bg_color, padx=20, pady=20)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title section with logo-like styling
        title_frame = tk.Frame(self.main_frame, bg=self.bg_color)
        title_frame.pack(fill=tk.X, pady=20)
        
        # Logo/icon representation
        logo_frame = tk.Canvas(title_frame, width=40, height=40, bg=self.bg_color, highlightthickness=0)
        logo_frame.pack(side=tk.LEFT, padx=10)
        logo_frame.create_oval(5, 5, 35, 35, fill=self.primary_color, outline="")
        logo_frame.create_text(20, 20, text="VT", fill="white", font=('Segoe UI', 14, 'bold'))
        
        # App title
        title_label = tk.Label(
            title_frame, 
            text="Vehicle Transport Batch Processor",
            font=self.title_font,
            fg=self.text_color,
            bg=self.bg_color
        )
        title_label.pack(side=tk.LEFT)
        
        # Main content card
        self.content_card = ModernFrame(
            self.main_frame, 
            bordercolor=self.border_color,
            bgcolor=self.card_bg,
            borderwidth=1
        )
        self.content_card.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Content inside the card
        self.content = tk.Frame(self.content_card.content_frame, bg=self.card_bg)
        self.content.pack(fill=tk.BOTH, expand=True)
        
        # Job type tabs container
        self.tabs_container = tk.Frame(self.content, bg=self.card_bg, padx=20, pady=20)
        self.tabs_container.pack(fill=tk.X)
        
        # Create tab buttons
        self.tabs_frame = tk.Frame(self.tabs_container, bg=self.card_bg)
        self.tabs_frame.pack(side=tk.LEFT, anchor=tk.W)
        
        self.tab_buttons = {}
        self.job_types = ["AC01", "BC04", "GR11", "CO03", "CW08/09"]  # Removed Email Automation
        
        # Create each tab button
        for i, job_type in enumerate(self.job_types):
            tab = TabButton(
                self.tabs_frame,
                text=job_type,
                command=self.switch_tab,
                selected=(i == 0),  # Set AC01 as default selected
                bg=self.tab_bg,
                selected_bg=self.tab_active,
                hover_bg=self.tab_hover,
                fg=self.text_color,
                selected_fg=self.primary_color
            )
            tab.grid(row=0, column=i, padx=5 if i > 0 else 0)
            self.tab_buttons[job_type] = tab
            
        # Tab content separator
        separator = tk.Frame(self.content, height=1, bg=self.border_color)
        separator.pack(fill=tk.X, padx=20)
        
        # Container for tab content
        self.tab_content = tk.Frame(self.content, bg=self.card_bg, padx=20, pady=20)
        self.tab_content.pack(fill=tk.BOTH, expand=True)
        
        # Create frames for each tab content 
        self.tab_frames = {}
        # Create AC01 content (current functionality)
        self.tab_frames["AC01"] = self.create_ac01_tab()
        # Create BC04 content (similar to AC01 but with BC04 parser)
        self.tab_frames["BC04"] = self.create_bc04_tab()
        # Create GR11 tab with functionality
        self.tab_frames["GR11"] = self.create_gr11_tab()
        # Create CO03 as coming soon
        self.tab_frames["CO03"] = self.create_coming_soon_tab("CO03")
        # Create CW08/09 as a real functional tab
        self.tab_frames["CW08/09"] = self.create_cw08_09_tab()

        # Show default tab (AC01)
        self.current_tab = "AC01"
        self.show_tab("AC01")
        
        # Footer
        footer = tk.Frame(self.main_frame, bg=self.bg_color, pady=5)
        footer.pack(fill=tk.X)
        
        footer_text = tk.Label(
            footer,
            text="Vehicle Transport Parser v1.1",
            font=self.small_font,
            fg=self.light_text,
            bg=self.bg_color
        )
        footer_text.pack(side=tk.RIGHT)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            footer,
            textvariable=self.status_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.bg_color,
            anchor=tk.W
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Add clickable file link button (initially hidden)
        self.file_link_var = tk.StringVar()
        self.file_link_button = tk.Button(
            footer,
            textvariable=self.file_link_var,
            font=self.small_font,
            fg=self.primary_color,
            bg=self.card_bg,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            command=self.open_last_saved_file
        )
        self.file_link_button.pack(side=tk.LEFT, padx=10)
        self.file_link_button.pack_forget()  # Hide initially
    
    def switch_tab(self, tab_name):
        """Switch between tabs"""
        # Deselect current tab button
        self.tab_buttons[self.current_tab].set_selected(False)
        
        # Select new tab button
        self.tab_buttons[tab_name].set_selected(True)
        
        # Hide current tab content
        self.tab_frames[self.current_tab].pack_forget()
        
        # Show new tab content
        self.show_tab(tab_name)
        
        # Update current tab
        self.current_tab = tab_name
    
    def show_tab(self, tab_name):
        """Show a specific tab content"""
        self.tab_frames[tab_name].pack(fill=tk.BOTH, expand=True)
    
    def create_ac01_tab(self):
        """Create the AC01 tab content with current functionality"""
        ac01_frame = tk.Frame(self.tab_content, bg=self.card_bg)
        
        # Date selection section with modern styling
        date_section = tk.Frame(ac01_frame, bg=self.card_bg)
        date_section.pack(fill=tk.X, pady=20)
        
        # Section title
        date_title = tk.Label(
            date_section,
            text="Job Dates",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        date_title.pack(anchor=tk.W, pady=10)
        
        # Date inputs in their own card
        date_card = ModernFrame(
            date_section,
            bordercolor=self.border_color,
            bgcolor="#F9F9FC",
            borderwidth=1
        )
        date_card.pack(fill=tk.X, padx=0, pady=0)
        
        date_container = tk.Frame(date_card.content_frame, bg="#F9F9FC", padx=15, pady=15)
        date_container.pack(fill=tk.X)
        
        # Collection date
        collection_frame = tk.Frame(date_container, bg="#F9F9FC")
        collection_frame.pack(side=tk.LEFT, padx=30)
        
        collection_label = tk.Label(
            collection_frame,
            text="Collection Date",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        collection_label.pack(anchor=tk.W, pady=5)
        
        # Get today's date
        today = datetime.now()
        delivery = self.calculate_delivery_date(today)
        
        # Modern styled date entry
        self.collection_date = DateEntry(
            collection_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=0,
            date_pattern='dd/mm/yyyy',
            font=self.normal_font
        )
        self.collection_date.set_date(today)
        self.collection_date.pack(anchor=tk.W)
        self.collection_date.bind("<<DateEntrySelected>>", self.update_delivery_date)
        
        # Delivery date
        delivery_frame = tk.Frame(date_container, bg="#F9F9FC")
        delivery_frame.pack(side=tk.LEFT)
        
        delivery_label = tk.Label(
            delivery_frame,
            text="Delivery Date",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        delivery_label.pack(anchor=tk.W, pady=5)
        
        self.delivery_date = DateEntry(
            delivery_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=0,
            date_pattern='dd/mm/yyyy',
            font=self.normal_font
        )
        self.delivery_date.set_date(delivery)
        self.delivery_date.pack(anchor=tk.W)
        
        # Text input section
        text_section = tk.Frame(ac01_frame, bg=self.card_bg)
        text_section.pack(fill=tk.BOTH, expand=True, pady=15)
        
        # Section title with helper text
        text_header = tk.Frame(text_section, bg=self.card_bg)
        text_header.pack(fill=tk.X, pady=10)
        
        text_title = tk.Label(
            text_header,
            text="Job Data",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        text_title.pack(side=tk.LEFT)
        
        help_text = tk.Label(
            text_header,
            text="Paste the job text below",
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        help_text.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Job counter on the right side
        self.job_count_var = tk.StringVar()
        self.job_count_var.set("0 jobs detected")
        job_count_label = tk.Label(
            text_header,
            textvariable=self.job_count_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        job_count_label.pack(side=tk.RIGHT)
        
        # Modern text area with subtle border
        text_container = ModernFrame(
            text_section,
            bordercolor=self.border_color,
            bgcolor=self.card_bg,
            borderwidth=1
        )
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.text_input = scrolledtext.ScrolledText(
            text_container.content_frame,
            font=('Consolas', 10),
            wrap=tk.WORD,
            bd=0,
            highlightthickness=0,
            padx=10,
            pady=10,
            background="white",
            foreground=self.text_color
        )
        self.text_input.pack(fill=tk.BOTH, expand=True)
        
        # Bind text change to update job count
        self.text_input.bind("<KeyRelease>", self.update_job_count)
        
        # Price tracking section
        price_section = tk.Frame(ac01_frame, bg=self.card_bg)
        price_section.pack(fill=tk.X, pady=(5, 10))
        
        # Total price display
        price_label = tk.Label(
            price_section,
            textvariable=self.ac01_price_var,
            font=("Segoe UI", 12, "bold"),
            fg=self.success_color,
            bg=self.card_bg
        )
        price_label.pack(side=tk.RIGHT)
        
        # Reset price button
        reset_price_button = tk.Button(
            price_section,
            text="Reset Price",
            font=self.small_font,
            fg=self.text_color,
            bg=self.card_bg,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            command=self.reset_ac01_price
        )
        reset_price_button.pack(side=tk.RIGHT, padx=10)
        
        # Action section
        action_section = tk.Frame(ac01_frame, bg=self.card_bg)
        action_section.pack(fill=tk.X, pady=0)
        
        # Status message
        self.status_var = tk.StringVar()
        self.status_label = tk.Label(
            action_section,
            textvariable=self.status_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg,
            anchor=tk.W
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Process button
        self.process_button = RoundedButton(
            action_section,
            width=150,
            height=36,
            cornerradius=18,
            padding=8,
            color=self.primary_color,
            hover_color=self.secondary_color,
            text="Process Jobs",
            command=self.process_jobs,
            fg="white"
        )
        self.process_button.pack(side=tk.RIGHT)
        
        return ac01_frame
    
    def create_bc04_tab(self):
        """Create the BC04 tab content similar to AC01 but with BC04 parser"""
        bc04_frame = tk.Frame(self.tab_content, bg=self.card_bg)
        
        # Date selection section with modern styling
        date_section = tk.Frame(bc04_frame, bg=self.card_bg)
        date_section.pack(fill=tk.X, pady=20)
        
        # Section title
        date_title = tk.Label(
            date_section,
            text="Job Dates",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        date_title.pack(anchor=tk.W, pady=10)
        
        # Date inputs in their own card
        date_card = ModernFrame(
            date_section,
            bordercolor=self.border_color,
            bgcolor="#F9F9FC",
            borderwidth=1
        )
        date_card.pack(fill=tk.X, padx=0, pady=0)
        
        date_container = tk.Frame(date_card.content_frame, bg="#F9F9FC", padx=15, pady=15)
        date_container.pack(fill=tk.X)
        
        # Collection date
        collection_frame = tk.Frame(date_container, bg="#F9F9FC")
        collection_frame.pack(side=tk.LEFT, padx=30)
        
        collection_label = tk.Label(
            collection_frame,
            text="Collection Date",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        collection_label.pack(anchor=tk.W, pady=5)
        
        # Get today's date
        today = datetime.now()
        delivery = self.calculate_delivery_date(today)
        
        # Modern styled date entry for BC04
        self.bc04_collection_date = DateEntry(
            collection_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=0,
            date_pattern='dd/mm/yyyy',
            font=self.normal_font
        )
        self.bc04_collection_date.set_date(today)
        self.bc04_collection_date.pack(anchor=tk.W)
        self.bc04_collection_date.bind("<<DateEntrySelected>>", self.update_bc04_delivery_date)
        
        # Delivery date
        delivery_frame = tk.Frame(date_container, bg="#F9F9FC")
        delivery_frame.pack(side=tk.LEFT)
        
        delivery_label = tk.Label(
            delivery_frame,
            text="Delivery Date",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        delivery_label.pack(anchor=tk.W, pady=5)
        
        self.bc04_delivery_date = DateEntry(
            delivery_frame,
            width=12,
            background=self.primary_color,
            foreground='white',
            borderwidth=0,
            date_pattern='dd/mm/yyyy',
            font=self.normal_font
        )
        self.bc04_delivery_date.set_date(delivery)
        self.bc04_delivery_date.pack(anchor=tk.W)
        
        # Text input section
        text_section = tk.Frame(bc04_frame, bg=self.card_bg)
        text_section.pack(fill=tk.BOTH, expand=True, pady=15)
        
        # Section title with helper text
        text_header = tk.Frame(text_section, bg=self.card_bg)
        text_header.pack(fill=tk.X, pady=10)
        
        text_title = tk.Label(
            text_header,
            text="Job Data",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        text_title.pack(side=tk.LEFT)
        
        help_text = tk.Label(
            text_header,
            text="Paste the BC04 job sheet text below",
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        help_text.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Job counter on the right side
        self.bc04_job_count_var = tk.StringVar()
        self.bc04_job_count_var.set("0 jobs detected")
        job_count_label = tk.Label(
            text_header,
            textvariable=self.bc04_job_count_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        job_count_label.pack(side=tk.RIGHT)
        
        # Modern text area with subtle border
        text_container = ModernFrame(
            text_section,
            bordercolor=self.border_color,
            bgcolor=self.card_bg,
            borderwidth=1
        )
        text_container.pack(fill=tk.BOTH, expand=True)
        
        self.bc04_text_input = scrolledtext.ScrolledText(
            text_container.content_frame,
            font=('Consolas', 10),
            wrap=tk.WORD,
            bd=0,
            highlightthickness=0,
            padx=10,
            pady=10,
            background="white",
            foreground=self.text_color
        )
        self.bc04_text_input.pack(fill=tk.BOTH, expand=True)
        
        # Bind text change to update job count
        self.bc04_text_input.bind("<KeyRelease>", self.update_bc04_job_count)
        
        # Price tracking section
        price_section = tk.Frame(bc04_frame, bg=self.card_bg)
        price_section.pack(fill=tk.X, pady=(5, 10))
        
        # Initialize BC04 price tracking
        if not hasattr(self, 'total_bc04_price'):
            self.total_bc04_price = 0.0
            self.bc04_price_var = tk.StringVar()
            self.bc04_price_var.set("Total Price: £0.00")
        
        # Total price display
        price_label = tk.Label(
            price_section,
            textvariable=self.bc04_price_var,
            font=("Segoe UI", 12, "bold"),
            fg=self.success_color,
            bg=self.card_bg
        )
        price_label.pack(side=tk.RIGHT)
        
        # Reset price button
        reset_price_button = tk.Button(
            price_section,
            text="Reset Price",
            font=self.small_font,
            fg=self.text_color,
            bg=self.card_bg,
            cursor="hand2",
            relief=tk.FLAT,
            bd=0,
            command=self.reset_bc04_price
        )
        reset_price_button.pack(side=tk.RIGHT, padx=10)
        
        # Action section
        action_section = tk.Frame(bc04_frame, bg=self.card_bg)
        action_section.pack(fill=tk.X, pady=0)
        
        # Status message
        self.bc04_status_var = tk.StringVar()
        self.bc04_status_label = tk.Label(
            action_section,
            textvariable=self.bc04_status_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg,
            anchor=tk.W
        )
        self.bc04_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Process button
        self.bc04_process_button = RoundedButton(
            action_section,
            width=150,
            height=36,
            cornerradius=18,
            padding=8,
            color=self.primary_color,
            hover_color=self.secondary_color,
            text="Process Jobs",
            command=self.process_bc04_jobs,
            fg="white"
        )
        self.bc04_process_button.pack(side=tk.RIGHT)
        
        return bc04_frame
    
    def create_gr11_tab(self):
        """Create the GR11 tab content for Excel file processing"""
        gr11_frame = tk.Frame(self.tab_content, bg=self.card_bg)
        
        # Date selection section with modern styling
        date_section = tk.Frame(gr11_frame, bg=self.card_bg)
        date_section.pack(fill=tk.X, pady=20)
        
        # Section title
        date_title = tk.Label(
            date_section,
            text="Excel Import",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        date_title.pack(anchor=tk.W, pady=10)
        
        # Excel import section
        import_card = ModernFrame(
            date_section,
            bordercolor=self.border_color,
            bgcolor="#F9F9FC",
            borderwidth=1
        )
        import_card.pack(fill=tk.X, padx=0, pady=0)
        
        import_container = tk.Frame(import_card.content_frame, bg="#F9F9FC", padx=15, pady=15)
        import_container.pack(fill=tk.X)
        
        # File selection
        file_frame = tk.Frame(import_container, bg="#F9F9FC")
        file_frame.pack(fill=tk.X, pady=10)
        
        # Selected file display
        self.selected_file_var = tk.StringVar()
        self.selected_file_var.set("No file selected")
        
        file_label = tk.Label(
            file_frame,
            text="Excel File:",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        file_label.pack(side=tk.LEFT, padx=(0, 10))
        
        file_path_label = tk.Label(
            file_frame,
            textvariable=self.selected_file_var,
            font=self.normal_font,
            fg=self.light_text,
            bg="#F9F9FC"
        )
        file_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_button = tk.Button(
            file_frame,
            text="Browse",
            font=self.normal_font,
            bg=self.primary_color,
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            command=self.browse_excel_file
        )
        browse_button.pack(side=tk.RIGHT)
        
        # Job preview section
        preview_section = tk.Frame(gr11_frame, bg=self.card_bg)
        preview_section.pack(fill=tk.BOTH, expand=True, pady=15)
        
        # Section title with job counter
        preview_header = tk.Frame(preview_section, bg=self.card_bg)
        preview_header.pack(fill=tk.X, pady=10)
        
        preview_title = tk.Label(
            preview_header,
            text="Job Preview",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        preview_title.pack(side=tk.LEFT)
        
        # Job counter on the right side
        self.gr_job_count_var = tk.StringVar()
        self.gr_job_count_var.set("0 jobs detected")
        job_count_label = tk.Label(
            preview_header,
            textvariable=self.gr_job_count_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        job_count_label.pack(side=tk.RIGHT)
        
        # Preview area
        preview_container = ModernFrame(
            preview_section,
            bordercolor=self.border_color,
            bgcolor=self.card_bg,
            borderwidth=1
        )
        preview_container.pack(fill=tk.BOTH, expand=True)
        
        # Create a frame to hold the treeview and scrollbar
        tree_frame = tk.Frame(preview_container.content_frame, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create scrollbar
        tree_scroll = tk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create the treeview
        columns = ("reg_no", "customer_ref", "vin", "make", "model", "collection_date", "collection_addr", "delivery_addr", "pdi_centre")
        self.job_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set)
        
        # Configure column headings
        self.job_tree.heading("reg_no", text="REG NUMBER")
        self.job_tree.heading("customer_ref", text="CUSTOMER REF")
        self.job_tree.heading("vin", text="VIN")
        self.job_tree.heading("make", text="MAKE")
        self.job_tree.heading("model", text="MODEL")
        self.job_tree.heading("collection_date", text="COLLECTION DATE")
        self.job_tree.heading("collection_addr", text="COLLECTION ADDRESS")
        self.job_tree.heading("delivery_addr", text="DELIVERY ADDRESS")
        self.job_tree.heading("pdi_centre", text="PDI CENTRE")
        
        # Configure column widths
        self.job_tree.column("reg_no", width=80)
        self.job_tree.column("customer_ref", width=100)
        self.job_tree.column("vin", width=140)
        self.job_tree.column("make", width=80)
        self.job_tree.column("model", width=100)
        self.job_tree.column("collection_date", width=100)
        self.job_tree.column("collection_addr", width=200)
        self.job_tree.column("delivery_addr", width=200)
        self.job_tree.column("pdi_centre", width=150)
        
        # Pack the treeview
        self.job_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure the scrollbar
        tree_scroll.config(command=self.job_tree.yview)
        
        # Action section
        action_section = tk.Frame(gr11_frame, bg=self.card_bg)
        action_section.pack(fill=tk.X, pady=10)
        
        # Status message
        self.gr_status_var = tk.StringVar()
        self.gr_status_label = tk.Label(
            action_section,
            textvariable=self.gr_status_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg,
            anchor=tk.W
        )
        self.gr_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Process button
        self.gr_process_button = RoundedButton(
            action_section,
            width=150,
            height=36,
            cornerradius=18,
            padding=8,
            color=self.primary_color,
            hover_color=self.secondary_color,
            text="Process Jobs",
            command=self.process_gr_jobs,
            fg="white"
        )
        self.gr_process_button.pack(side=tk.RIGHT)
        
        return gr11_frame
    
    def create_coming_soon_tab(self, job_type):
        """Create a 'Coming Soon' tab content for future job types"""
        frame = tk.Frame(self.tab_content, bg=self.card_bg)
        
        # Center the coming soon content
        coming_soon_frame = tk.Frame(frame, bg=self.card_bg)
        coming_soon_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Icon for coming soon
        icon_canvas = tk.Canvas(coming_soon_frame, width=80, height=80, 
                             bg=self.card_bg, highlightthickness=0)
        icon_canvas.pack(pady=20)
        
        # Draw clock or construction icon
        icon_canvas.create_oval(10, 10, 70, 70, fill="#F5F7FA", outline=self.primary_color, width=2)
        # Clock hands
        icon_canvas.create_line(40, 40, 40, 20, fill=self.primary_color, width=2)
        icon_canvas.create_line(40, 40, 55, 50, fill=self.primary_color, width=2)
        
        # Coming soon text
        title = tk.Label(
            coming_soon_frame,
            text=f"{job_type} Transport Processing",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        title.pack(pady=10)
        
        message = tk.Label(
            coming_soon_frame,
            text="Coming Soon",
            font=('Segoe UI', 16, 'bold'),
            fg=self.primary_color,
            bg=self.card_bg
        )
        message.pack(pady=20)
        
        description = tk.Label(
            coming_soon_frame,
            text=f"Support for {job_type} transport jobs is currently under development.\nCheck back soon for updates.",
            font=self.normal_font,
            fg=self.light_text,
            bg=self.card_bg,
            justify=tk.CENTER
        )
        description.pack()
        
        return frame
        
    def calculate_delivery_date(self, collection_date):
        """Calculate delivery date as 3 business days from collection date."""
        # Get UK holidays
        uk_holidays = holidays.UK()
        
        # Start with collection date
        current_date = collection_date
        business_days = 0
        
        # Keep adding days until we have 3 business days
        while business_days < 3:
            current_date += timedelta(days=1)
            # Skip weekends and holidays
            if current_date.weekday() < 5 and current_date not in uk_holidays:
                business_days += 1
                
        return current_date
    
    def update_delivery_date(self, event=None):
        """Update delivery date when collection date changes"""
        collection = self.collection_date.get_date()
        delivery = self.calculate_delivery_date(collection)
        self.delivery_date.set_date(delivery)
    
    def update_job_count(self, event=None):
        """Count the number of jobs in the text input and update the counter label"""
        text = self.text_input.get("1.0", tk.END)
        
        # Count jobs by counting "FROM" sections
        job_count = text.count("\nFROM\n")
        
        # Add one more if the text starts with "FROM"
        if text.lstrip().startswith("FROM"):
            job_count += 1
        
        # Update the counter
        if job_count == 1:
            self.job_count_var.set("1 job detected")
        else:
            self.job_count_var.set(f"{job_count} jobs detected")
    
    def process_jobs(self):
        try:
            # Update status
            self.status_var.set("Processing jobs...")
            self.status_label.config(fg=self.primary_color)
            self.root.update_idletasks()
            
            # Get the input text
            text = self.text_input.get("1.0", tk.END)
            
            # Get dates directly from the date pickers
            collection_date = self.collection_date.get_date().strftime("%d/%m/%Y")
            delivery_date = self.delivery_date.get_date().strftime("%d/%m/%Y")
            
            # Parse the jobs
            parser = JobParser(collection_date, delivery_date)
            jobs = parser.parse_jobs(text)

            # Debug: Log each job's REG NUMBER
            for i, job in enumerate(jobs, 1):
                reg_val = job.get('REG NUMBER')
                print(f"Job {i} REG NUMBER: '{reg_val}'")
                log_debug(f"Job {i} REG NUMBER: '{reg_val}'")

            if not jobs:
                self.status_var.set("No valid jobs found in the input text")
                self.status_label.config(fg=self.warning_color)
                messagebox.showerror("Error", "No valid jobs found in the input text")
                return

            # Count jobs and registrations (robust)
            job_count = len(jobs)
            reg_count = sum(1 for job in jobs if isinstance(job.get('REG NUMBER'), str) and job.get('REG NUMBER').strip())

            # Update the job count label to show both
            if job_count == 1:
                self.job_count_var.set(f"1 job detected, {reg_count} registrations detected")
            else:
                self.job_count_var.set(f"{job_count} jobs detected, {reg_count} registrations detected")
            
            # For AC01 jobs, validate registrations
            if self.current_tab == "AC01":
                missing_regs = []
                for i, job in enumerate(jobs, 1):
                    if not job.get('REG NUMBER') or job['REG NUMBER'].strip() == '':
                        missing_regs.append(f"Job #{i}")
                
                if missing_regs:
                    missing_msg = f"The following jobs are missing registration numbers:\n{', '.join(missing_regs)}"
                    self.status_var.set("Warning: Some jobs missing registration numbers")
                    self.status_label.config(fg=self.warning_color)
                    messagebox.showwarning("Missing Registrations", missing_msg)
            
            # Get the current job type (tab)
            job_type = self.current_tab
            
            # Generate CSV file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{job_type.lower()}_jobs_{timestamp}.csv"
            
            try:
                self.save_to_csv(jobs, filename, job_type)
                
                # Calculate total price for this batch of jobs
                batch_total = 0.0
                for job in jobs:
                    if 'PRICE' in job and job['PRICE']:
                        try:
                            # Handle various price formats
                            price_str = job['PRICE'].strip()
                            price_str = price_str.replace('£', '').replace(',', '')
                            price = float(price_str)
                            batch_total += price
                        except (ValueError, TypeError):
                            log_debug(f"Could not parse price: {job.get('PRICE', 'N/A')}")
                
                # Update the total price if this is AC01
                if job_type == "AC01":
                    self.total_ac01_price += batch_total
                    self.ac01_price_var.set(f"Total Price: £{self.total_ac01_price:.2f}")
                
                self.status_var.set(f"Successfully processed {len(jobs)} jobs")
                self.status_label.config(fg=self.success_color)
                
                # Include price in success message
                if batch_total > 0:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} {job_type} jobs\nBatch Total: £{batch_total:.2f}\nRunning Total: £{self.total_ac01_price:.2f}")
                else:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} {job_type} jobs")
                
            except Exception as e:
                self.status_var.set("Error: Failed to save CSV file")
                self.status_label.config(fg=self.error_color)
                messagebox.showerror("Error", f"Failed to save CSV file: {str(e)}")
            
        except Exception as e:
            self.status_var.set("Error occurred during processing")
            self.status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_to_csv(self, jobs, filename, job_type):
        if not jobs:
            return

        try:
            # Clear debug log before saving
            log_debug("\n--- START SAVING JOBS TO CSV ---")
            
            # Get the current script directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            log_debug(f"Current directory: {current_dir}")
            
            # Go up one level to get the workspace root directory
            workspace_root = os.path.dirname(current_dir)
            log_debug(f"Workspace root: {workspace_root}")
            
            # Create main 'jobs' directory if it doesn't exist
            output_dir = os.path.join(workspace_root, 'jobs')
            os.makedirs(output_dir, exist_ok=True)
            log_debug(f"Main jobs directory: {output_dir}")
            
            # Create job type specific directory if it doesn't exist
            job_type_dir = os.path.join(output_dir, job_type)
            os.makedirs(job_type_dir, exist_ok=True)
            log_debug(f"Job type directory: {job_type_dir}")
            
            # Full path for the output file
            output_path = os.path.join(job_type_dir, filename)
            log_debug(f"Output file path: {output_path}")
            
            # Debug: Count original job types
            gr11_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR11')
            gr15_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR15')
            log_debug(f"Job counts before saving - GR11: {gr11_count}, GR15: {gr15_count}")
            
            # Log all customer references before saving
            log_debug("Customer references before saving:")
            for i, job in enumerate(jobs):
                log_debug(f"Job {i} - Reg: {job.get('REG NUMBER', '')}, Customer Ref: {job.get('CUSTOMER REF', '')}")
            
            # IMPORTANT: For Greenhous jobs NEVER override the customer reference
            # For other job types, set all jobs to the same customer ref
            if job_type != 'GR11' and not job_type.startswith('GR'):
                # Set all jobs to the same customer ref for non-Greenhous jobs
                for job in jobs:
                    job['CUSTOMER REF'] = job_type
            
            # Debug: Check final job types
            gr11_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR11')
            gr15_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR15')
            log_debug(f"After customer ref setup - GR11: {gr11_count}, GR15: {gr15_count}")
            
            # Check if the output file already exists
            if os.path.exists(output_path):
                log_debug(f"Warning: Output file already exists, will try to overwrite: {output_path}")
                # Check if we can write to it
                try:
                    with open(output_path, 'a') as f:
                        pass  # Just testing if we can open it for append
                except PermissionError:
                    log_debug(f"Permission denied when trying to open existing file: {output_path}")
                    raise PermissionError(f"The file {filename} is open in another program. Please close it and try again.")
            
            # Test folder permissions
            test_file = os.path.join(job_type_dir, "_test_write_permission.tmp")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                log_debug(f"Write permission test successful on directory: {job_type_dir}")
            except Exception as e:
                log_debug(f"Failed write permission test: {str(e)}")
                raise PermissionError(f"Cannot write to the output directory ({job_type_dir}). Please check your permissions.")
            
            # Try to save the file
            log_debug(f"Attempting to write CSV to: {output_path}")
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = [
                    'REG NUMBER',
                    'VIN',
                    'MAKE',
                    'MODEL',
                    'COLLECTION DATE',
                    'YOUR REF NO',
                    'COLLECTION ADDR1',
                    'COLLECTION ADDR2',
                    'COLLECTION ADDR3',
                    'COLLECTION ADDR4',
                    'COLLECTION POSTCODE',
                    'COLLECTION CONTACT NAME',
                    'COLLECTION PHONE',
                    'DELIVERY DATE',
                    'DELIVERY ADDR1',
                    'DELIVERY ADDR2',
                    'DELIVERY ADDR3',
                    'DELIVERY ADDR4',
                    'DELIVERY POSTCODE',
                    'DELIVERY CONTACT NAME',
                    'DELIVERY CONTACT PHONE',
                    'SPECIAL INSTRUCTIONS',
                    'PRICE',
                    'CUSTOMER REF',
                    'TRANSPORT TYPE'
                ]

                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                log_debug("Writing to CSV file:")
                for job in jobs:
                    # Create a new dict with only the fields we want
                    row = {field: job.get(field, '') for field in fieldnames}
                    # Debug: Check individual job customer reference
                    log_debug(f"Writing job {row['REG NUMBER']} with customer ref: {row['CUSTOMER REF']}")
                    writer.writerow(row)
                
            log_debug(f"CSV file saved to: {output_path}")
            log_debug("--- END SAVING JOBS TO CSV ---")
            
            self.status_var.set(f"File saved: {filename}")
            self.last_saved_file = output_path
            self.file_link_var.set(f"View file: {filename}")
            self.file_link_button.pack(side=tk.LEFT, padx=10)
            print("File link button packed and visible")
            messagebox.showinfo("Success", f"File saved successfully at:\n{output_path}")

        except PermissionError as pe:
            log_debug(f"Permission error during save: {str(pe)}")
            error_msg = f"Permission denied. Please ensure:\n1. No other program has the file open\n2. You have permission to write to the folder: {os.path.dirname(output_path) if 'output_path' in locals() else 'unknown'}\n3. Close Excel or any CSV viewers before trying again."
            messagebox.showerror("Permission Error", error_msg)
            raise pe
        except Exception as e:
            log_debug(f"Error during save: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{str(e)}")
            raise e
    
    def process_gr_jobs(self):
        """Process the GR11/GR15 jobs from the Excel data"""
        try:
            # Get all items from treeview
            items = self.job_tree.get_children()
            
            if not items:
                messagebox.showerror("Error", "No jobs to process. Please load an Excel file first.")
                return
            
            # Collect job data
            jobs = []
            for item in items:
                try:
                    values = self.job_tree.item(item, "values")
                    
                    # Check if we have the minimum required data
                    if not values[0]:  # Reg No
                        continue
                    
                    # Get the customer ref and PDI Centre values
                    # values[1] is the customer ref from the preview (GR11 or GR15)
                    customer_ref = values[1] if len(values) > 1 and values[1] else "GR11"
                    
                    # values[8] is the PDI Centre column
                    pdi_centre = values[8] if len(values) > 8 and values[8] else ""
                    
                    # Debug: Print the PDI Centre value
                    print(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    log_debug(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    
                    # Collection address details
                    collection_addr1 = ""
                    collection_addr2 = ""
                    collection_addr3 = ""
                    collection_addr4 = ""
                    collection_postcode = ""
                    
                    # Use the customer reference already set in the preview 
                    # This ensures consistency with what's shown in the UI
                    
                    # Set address based on the customer reference
                    if customer_ref == "GR15":
                        collection_addr1 = "Greenhous Upper Heyford"
                        collection_addr2 = "Heyford Park, Bicester"
                        collection_addr3 = "Bicester"
                        collection_addr4 = "UPPER HEYFORD"
                        collection_postcode = "OX25 5HA"
                        print(f"  => Using Upper Heyford address")
                        log_debug(f"  => Using Upper Heyford address")
                    else:
                        # Default to High Ercall (GR11)
                        collection_addr1 = "Greenhous High Ercall"
                        collection_addr2 = "Greenhous Village Osbaston"
                        collection_addr3 = "High Ercall"
                        collection_addr4 = ""
                        collection_postcode = "TF6 6RA"
                        print(f"  => Using High Ercall address")
                        log_debug(f"  => Using High Ercall address")
                    
                    # Create job dictionary
                    job = {
                        'REG NUMBER': values[0],
                        'VIN': values[2] if len(values) > 2 and values[2] else "",  # Chassis
                        'MAKE': values[3] if len(values) > 3 and values[3] else "TRANSIT",  # Make
                        'MODEL': values[4] if len(values) > 4 and values[4] else "VAN",  # Model
                        'COLOR': "",
                        'COLLECTION DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Collection date
                        'YOUR REF NO': values[0],  # Use reg number as reference
                        'COLLECTION ADDR1': collection_addr1,
                        'COLLECTION ADDR2': collection_addr2,
                        'COLLECTION ADDR3': collection_addr3,
                        'COLLECTION ADDR4': collection_addr4,
                        'COLLECTION POSTCODE': collection_postcode,
                        'COLLECTION CONTACT NAME': "",
                        'COLLECTION PHONE': "",
                        'DELIVERY DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Same as collection date
                        'DELIVERY ADDR1': "",
                        'DELIVERY ADDR2': "",
                        'DELIVERY ADDR3': "",
                        'DELIVERY ADDR4': "",
                        'DELIVERY POSTCODE': "",
                        'DELIVERY CONTACT NAME': "",
                        'DELIVERY CONTACT PHONE': "",
                        'SPECIAL INSTRUCTIONS': f"VIN: {values[2] if len(values) > 2 and values[2] else 'Unknown'}",
                        'PRICE': "",
                        'CUSTOMER REF': customer_ref,
                        'TRANSPORT TYPE': ""
                    }
                    
                    # Parse delivery address
                    if len(values) > 7 and values[7]:
                        delivery_address = values[7]  # Correct index for delivery address
                        
                        # Pre-process common patterns that should be kept together
                        # Examples: "FLEX E REN 84-90", "FLEX E RE WEST LON"
                        patterns_to_fix = [
                            (r'(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)[\s,]+(\d+[\-\d]*)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(WEST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(UNIT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(MARCH)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(IVATT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(LANDGATI)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ASCOT)', r'\1 \2'),
                            # New patterns to keep road numbers with road names
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:Road|Street|Avenue|Lane|Drive|Close|Way|ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY))', r'\1 \2'),
                            # Handle specific cases like 84-90 BRADES ROAD
                            (r'(\d+[\-\d]*)\s*,\s*(BRADES[\s]*ROAD)', r'\1 \2'),
                            # More general pattern for any number + street pattern that got split
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:\s(?:Rd|St|Ave|Ln|Dr|Cl|RD|ST|AVE|LN|DR|CL)))', r'\1 \2')
                        ]
                        
                        # Apply each pattern fix
                        processed_address = delivery_address
                        for pattern, replacement in patterns_to_fix:
                            processed_address = re.sub(pattern, replacement, processed_address, flags=re.IGNORECASE)
                        
                        # Log any changes made during pre-processing
                        if processed_address != delivery_address:
                            log_debug(f"Pre-processed address: '{delivery_address}' -> '{processed_address}'")
                            delivery_address = processed_address
                        
                        # Try to extract postcode from address (UK format)
                        postcode_pattern = r'([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})'
                        postcode_match = re.search(postcode_pattern, delivery_address, re.IGNORECASE)
                        
                        # Clean the postcode if found
                        extracted_postcode = ""
                        if postcode_match:
                            extracted_postcode = postcode_match.group(1).upper().strip()
                            # Clean up spacing in the postcode
                            if len(extracted_postcode) > 4 and ' ' not in extracted_postcode:
                                # Add space in the correct position for UK postcodes (before last 3 chars)
                                extracted_postcode = extracted_postcode[:-3] + ' ' + extracted_postcode[-3:]
                            job['DELIVERY POSTCODE'] = extracted_postcode
                        
                        # Remove the postcode from the address for cleaner parsing
                        if extracted_postcode:
                            clean_address = delivery_address.replace(postcode_match.group(0), "")
                        else:
                            clean_address = delivery_address
                        
                        # Enhanced address parsing to keep road numbers with street names
                        # First split by commas
                        address_parts = [part.strip() for part in clean_address.split(',') if part.strip()]
                        
                        # First pass: Check for patterns like "FLEX-E-RE 84-90" being split across parts
                        # or combine prefix with following parts
                        combined_parts = []
                        i = 0
                        while i < len(address_parts):
                            current_part = address_parts[i]
                            
                            # Special patterns to check for
                            flex_prefix_pattern = r'^(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)$'
                            
                            # Check if this part ends with a prefix that should be combined with the next part
                            if i < len(address_parts) - 1 and re.match(flex_prefix_pattern, current_part, re.IGNORECASE):
                                # Look ahead to see if we need to combine more parts (prefix + number + road name)
                                if i < len(address_parts) - 2 and re.match(r'^\d+[\-\d]*$', address_parts[i+1]) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+2], re.IGNORECASE):
                                    # We have a pattern like: [FLEX-E-RE], [84-90], [BRADES ROAD]
                                    combined_parts.append(f"{current_part} {address_parts[i+1]} {address_parts[i+2]}")
                                    i += 3  # Skip all three parts
                                else:
                                    # Just combine with next part
                                    combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                    i += 2  # Skip both parts
                            # Check for standalone number that should be combined with street name
                            elif i < len(address_parts) - 1 and re.match(r'^\d+[\-\d]*$', current_part) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+1], re.IGNORECASE):
                                # We have a pattern like: [84-90], [BRADES ROAD]
                                combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                i += 2  # Skip both parts
                            else:
                                combined_parts.append(current_part)
                                i += 1
                        
                        # Replace original parts with combined parts
                        address_parts = combined_parts
                        
                        # Log the address parts after combining
                        log_debug(f"Address parts after combining prefixes and road numbers: {address_parts}")
                        
                        # Initialize address fields
                        addr1 = ""
                        addr2 = ""
                        addr3 = ""
                        addr4 = ""
                        
                        # Process address parts based on their content and position
                        if len(address_parts) > 0:
                            addr1 = address_parts[0]  # First part is always ADDR1 (usually business name)
                        
                        if len(address_parts) > 1:
                            # Check if the second part looks like a road/street with numbers
                            road_pattern = r'^\d+[\-\d]*\s+[A-Za-z\s]+'  # Matches patterns like "84-90 BRADES ROAD"
                            if len(address_parts) > 2 and (re.match(road_pattern, address_parts[1]) or 
                                                          re.search(r'\b(Road|Street|Avenue|Lane|Drive|Close)\b', address_parts[1], re.IGNORECASE)):
                                # This is a street address, keep it in ADDR2
                                addr2 = address_parts[1]
                                
                                # The rest goes into ADDR3 and ADDR4
                                if len(address_parts) > 3:
                                    addr3 = address_parts[2]
                                    addr4 = address_parts[3]
                                elif len(address_parts) > 2:
                                    addr3 = address_parts[2]
                            else:
                                # Standard processing
                                if len(address_parts) > 2:
                                    addr2 = address_parts[1]
                                    addr3 = address_parts[2]
                                    if len(address_parts) > 3:
                                        addr4 = ' '.join(address_parts[3:])
                                else:
                                    addr2 = address_parts[1]
                        
                        # Special handling for when ADDR4 is empty but we have a town in ADDR3
                        if not addr4 and addr3 and len(addr3.split()) <= 2:
                            addr4 = addr3
                            addr3 = ""
                        
                        # Set the address fields
                        job['DELIVERY ADDR1'] = addr1
                        job['DELIVERY ADDR2'] = addr2
                        job['DELIVERY ADDR3'] = addr3
                        job['DELIVERY ADDR4'] = addr4
                    
                    jobs.append(job)
                except Exception as item_error:
                    # Log individual item parsing errors but continue with others
                    print(f"Error processing item: {str(item_error)}")
                    continue
            
            if not jobs:
                messagebox.showerror("Error", "No valid jobs to process.")
                return
            
            # Count jobs by type for reporting purposes
            gr11_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR11')
            gr15_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR15')
            
            # Generate a single CSV file for all jobs
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"greenhous_jobs_{timestamp}.csv"
            
            # Save to the GR11 folder for consistency
            self.save_to_csv(jobs, filename, 'GR11')
            
            # Success message
            total_jobs = len(jobs)
            
            self.gr_status_var.set(f"Loaded {job_count} jobs from file")
            self.gr_status_label.config(fg=self.success_color)
            
            # More detailed success message with counts of each job type
            message = f"Processed {total_jobs} Greenhous jobs in total:\n"
            
            # Count by job type
            job_types = {}
            for job in jobs:
                customer_ref = job.get('CUSTOMER REF', 'Unknown')
                job_types[customer_ref] = job_types.get(customer_ref, 0) + 1
            
            # Add counts to message
            for ref, count in job_types.items():
                message += f"- {count} {ref} jobs\n"
                
            message += f"\nAll jobs saved to: jobs/GR11/{filename}"
                
            messagebox.showinfo("Success", message)
            
        except Exception as e:
            self.gr_status_var.set(f"Error occurred: {str(e)}")
            self.gr_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
    def browse_excel_file(self):
        """Open file dialog to select Excel file"""
        filetypes = [("Excel files", "*.xlsx *.xls")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        
        if file_path:
            self.selected_file_var.set(file_path)
            self.load_excel_data(file_path)
    
    def load_excel_data(self, file_path):
        """Load data from Excel file and display in treeview"""
        try:
            # Clear existing items
            for item in self.job_tree.get_children():
                self.job_tree.delete(item)
            
            # Load Excel file
            df = pd.read_excel(file_path)
            
            # Check if required columns exist
            required_columns = ['Reg No', 'PDI Centre', 'Model', 'Chassis', 'Delivery Due Date', 'Delivery Address']
            
            # Check column names (case insensitive)
            actual_columns = list(df.columns)
            found_columns = []
            
            for req_col in required_columns:
                found = False
                for act_col in actual_columns:
                    if req_col.lower() == act_col.lower():
                        found_columns.append(act_col)
                        found = True
                        break
                if not found:
                    found_columns.append(None)
            
            # Use the found column names or provide feedback if any are missing
            missing_columns = [req_col for i, req_col in enumerate(required_columns) if found_columns[i] is None]
            
            if missing_columns:
                messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_columns)}")
                self.gr_job_count_var.set("0 jobs detected")
                return
            
            # Get the actual column names from the file
            reg_col = found_columns[0]
            pdi_col = found_columns[1]
            model_col = found_columns[2]
            chassis_col = found_columns[3]
            date_col = found_columns[4]
            address_col = found_columns[5]
            
            # Debug: Print found column names
            print(f"Found columns:")
            print(f"  Reg No: {reg_col}")
            print(f"  PDI Centre: {pdi_col}")
            print(f"  Model: {model_col}")
            print(f"  Chassis: {chassis_col}")
            print(f"  Delivery Due Date: {date_col}")
            print(f"  Delivery Address: {address_col}")
            
            # Log to debug file
            log_debug("Excel columns found:")
            log_debug(f"  Reg No: {reg_col}")
            log_debug(f"  PDI Centre: {pdi_col}")
            log_debug(f"  Model: {model_col}")
            log_debug(f"  Chassis: {chassis_col}")
            log_debug(f"  Delivery Due Date: {date_col}")
            log_debug(f"  Delivery Address: {address_col}")
            log_debug("\nExcel data preview:")
            
            # Log a sample of the data
            for i, (_, row) in enumerate(df.head(5).iterrows()):
                log_debug(f"Row {i}: {row[pdi_col]} - {row[reg_col]}")
                
            # Add data to treeview
            job_count = 0
            for _, row in df.iterrows():
                # Skip rows with empty registration numbers
                if pd.isna(row[reg_col]) or str(row[reg_col]).strip() == "":
                    continue
                
                # Format the date if it exists
                delivery_date = ""
                if not pd.isna(row[date_col]):
                    if isinstance(row[date_col], str):
                        delivery_date = row[date_col]
                    else:
                        try:
                            delivery_date = row[date_col].strftime("%d/%m/%Y")
                        except:
                            delivery_date = str(row[date_col])
                
                # Process the data to match the CSV output format
                pdi_centre = str(row[pdi_col]) if not pd.isna(row[pdi_col]) else ""
                
                # Debug: Print PDI Centre values
                print(f"PDI Centre: '{pdi_centre}'")
                log_debug(f"PDI Centre detected: '{pdi_centre}' for reg {row[reg_col]}")
                
                # Determine customer reference - improved detection logic
                customer_ref = "GR11"  # Default
                
                # First look for "Upper" Heyford in the PDI Centre column
                if "UPPER" in pdi_centre.upper():
                    customer_ref = "GR15"
                    log_debug(f"  -> Detected as GR15 (UPPER keyword)")
                # Then check for Heyford
                elif "HEYFORD" in pdi_centre.upper():
                    customer_ref = "GR15"
                    log_debug(f"  -> Detected as GR15 (HEYFORD keyword)")
                else:
                    log_debug(f"  -> Detected as GR11 (default)")
                
                # Set collection address based on customer_ref
                if customer_ref == "GR15":
                    collection_addr = "Greenhous Upper Heyford, Heyford Park, Bicester, OX25 5HA"
                else:
                    collection_addr = "Greenhous High Ercall, Greenhous Village Osbaston, TF6 6RA"
                
                # Try to extract Make from Model
                make = ""
                model = str(row[model_col]) if not pd.isna(row[model_col]) else ""
                
                if model:
                    # Common car makes to check
                    makes = ["FORD", "VAUXHALL", "VOLKSWAGEN", "VW", "BMW", "MERCEDES", "AUDI", 
                            "TOYOTA", "HONDA", "NISSAN", "HYUNDAI", "KIA", "SKODA", "SEAT", 
                            "RENAULT", "PEUGEOT", "CITROEN", "FIAT", "MAZDA", "VOLVO"]
                    
                    for m in makes:
                        if m.lower() in model.lower():
                            make = m
                            # Remove make from model if it's at the beginning
                            if model.lower().startswith(m.lower()):
                                model = model[len(m):].strip()
                            break
                
                # Get delivery address
                delivery_addr = str(row[address_col]) if not pd.isna(row[address_col]) else ""
                
                # Add to treeview with the processed data
                self.job_tree.insert("", tk.END, values=(
                    str(row[reg_col]),
                    customer_ref,
                    str(row[chassis_col]) if not pd.isna(row[chassis_col]) else "",
                    make,
                    model,
                    delivery_date,
                    collection_addr,
                    delivery_addr,
                    pdi_centre
                ))
                job_count += 1
            
            # Update job count
            if job_count == 1:
                self.gr_job_count_var.set("1 job detected")
            else:
                self.gr_job_count_var.set(f"{job_count} jobs detected")
                
            self.gr_status_var.set(f"Loaded {job_count} jobs from file")
            self.gr_status_label.config(fg=self.success_color)
            
        except Exception as e:
            self.gr_status_var.set(f"Error loading file: {str(e)}")
            self.gr_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    
    def process_gr_jobs(self):
        """Process the GR11/GR15 jobs from the Excel data"""
        try:
            # Get all items from treeview
            items = self.job_tree.get_children()
            
            if not items:
                messagebox.showerror("Error", "No jobs to process. Please load an Excel file first.")
                return
            
            # Collect job data
            jobs = []
            for item in items:
                try:
                    values = self.job_tree.item(item, "values")
                    
                    # Check if we have the minimum required data
                    if not values[0]:  # Reg No
                        continue
                    
                    # Get the customer ref and PDI Centre values
                    # values[1] is the customer ref from the preview (GR11 or GR15)
                    customer_ref = values[1] if len(values) > 1 and values[1] else "GR11"
                    
                    # values[8] is the PDI Centre column
                    pdi_centre = values[8] if len(values) > 8 and values[8] else ""
                    
                    # Debug: Print the PDI Centre value
                    print(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    log_debug(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    
                    # Collection address details
                    collection_addr1 = ""
                    collection_addr2 = ""
                    collection_addr3 = ""
                    collection_addr4 = ""
                    collection_postcode = ""
                    
                    # Use the customer reference already set in the preview 
                    # This ensures consistency with what's shown in the UI
                    
                    # Set address based on the customer reference
                    if customer_ref == "GR15":
                        collection_addr1 = "Greenhous Upper Heyford"
                        collection_addr2 = "Heyford Park, Bicester"
                        collection_addr3 = "Bicester"
                        collection_addr4 = "UPPER HEYFORD"
                        collection_postcode = "OX25 5HA"
                        print(f"  => Using Upper Heyford address")
                        log_debug(f"  => Using Upper Heyford address")
                    else:
                        # Default to High Ercall (GR11)
                        collection_addr1 = "Greenhous High Ercall"
                        collection_addr2 = "Greenhous Village Osbaston"
                        collection_addr3 = "High Ercall"
                        collection_addr4 = ""
                        collection_postcode = "TF6 6RA"
                        print(f"  => Using High Ercall address")
                        log_debug(f"  => Using High Ercall address")
                    
                    # Create job dictionary
                    job = {
                        'REG NUMBER': values[0],
                        'VIN': values[2] if len(values) > 2 and values[2] else "",  # Chassis
                        'MAKE': values[3] if len(values) > 3 and values[3] else "TRANSIT",  # Make
                        'MODEL': values[4] if len(values) > 4 and values[4] else "VAN",  # Model
                        'COLOR': "",
                        'COLLECTION DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Collection date
                        'YOUR REF NO': values[0],  # Use reg number as reference
                        'COLLECTION ADDR1': collection_addr1,
                        'COLLECTION ADDR2': collection_addr2,
                        'COLLECTION ADDR3': collection_addr3,
                        'COLLECTION ADDR4': collection_addr4,
                        'COLLECTION POSTCODE': collection_postcode,
                        'COLLECTION CONTACT NAME': "",
                        'COLLECTION PHONE': "",
                        'DELIVERY DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Same as collection date
                        'DELIVERY ADDR1': "",
                        'DELIVERY ADDR2': "",
                        'DELIVERY ADDR3': "",
                        'DELIVERY ADDR4': "",
                        'DELIVERY POSTCODE': "",
                        'DELIVERY CONTACT NAME': "",
                        'DELIVERY CONTACT PHONE': "",
                        'SPECIAL INSTRUCTIONS': f"VIN: {values[2] if len(values) > 2 and values[2] else 'Unknown'}",
                        'PRICE': "",
                        'CUSTOMER REF': customer_ref,
                        'TRANSPORT TYPE': ""
                    }
                    
                    # Parse delivery address
                    if len(values) > 7 and values[7]:
                        delivery_address = values[7]  # Correct index for delivery address
                        
                        # Pre-process common patterns that should be kept together
                        # Examples: "FLEX E REN 84-90", "FLEX E RE WEST LON"
                        patterns_to_fix = [
                            (r'(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)[\s,]+(\d+[\-\d]*)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(WEST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(UNIT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(MARCH)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(IVATT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(LANDGATI)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ASCOT)', r'\1 \2'),
                            # New patterns to keep road numbers with road names
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:Road|Street|Avenue|Lane|Drive|Close|Way|ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY))', r'\1 \2'),
                            # Handle specific cases like 84-90 BRADES ROAD
                            (r'(\d+[\-\d]*)\s*,\s*(BRADES[\s]*ROAD)', r'\1 \2'),
                            # More general pattern for any number + street pattern that got split
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:\s(?:Rd|St|Ave|Ln|Dr|Cl|RD|ST|AVE|LN|DR|CL)))', r'\1 \2')
                        ]
                        
                        # Apply each pattern fix
                        processed_address = delivery_address
                        for pattern, replacement in patterns_to_fix:
                            processed_address = re.sub(pattern, replacement, processed_address, flags=re.IGNORECASE)
                        
                        # Log any changes made during pre-processing
                        if processed_address != delivery_address:
                            log_debug(f"Pre-processed address: '{delivery_address}' -> '{processed_address}'")
                            delivery_address = processed_address
                        
                        # Try to extract postcode from address (UK format)
                        postcode_pattern = r'([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})'
                        postcode_match = re.search(postcode_pattern, delivery_address, re.IGNORECASE)
                        
                        # Clean the postcode if found
                        extracted_postcode = ""
                        if postcode_match:
                            extracted_postcode = postcode_match.group(1).upper().strip()
                            # Clean up spacing in the postcode
                            if len(extracted_postcode) > 4 and ' ' not in extracted_postcode:
                                # Add space in the correct position for UK postcodes (before last 3 chars)
                                extracted_postcode = extracted_postcode[:-3] + ' ' + extracted_postcode[-3:]
                            job['DELIVERY POSTCODE'] = extracted_postcode
                        
                        # Remove the postcode from the address for cleaner parsing
                        if extracted_postcode:
                            clean_address = delivery_address.replace(postcode_match.group(0), "")
                        else:
                            clean_address = delivery_address
                        
                        # Enhanced address parsing to keep road numbers with street names
                        # First split by commas
                        address_parts = [part.strip() for part in clean_address.split(',') if part.strip()]
                        
                        # First pass: Check for patterns like "FLEX-E-RE 84-90" being split across parts
                        # or combine prefix with following parts
                        combined_parts = []
                        i = 0
                        while i < len(address_parts):
                            current_part = address_parts[i]
                            
                            # Special patterns to check for
                            flex_prefix_pattern = r'^(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)$'
                            
                            # Check if this part ends with a prefix that should be combined with the next part
                            if i < len(address_parts) - 1 and re.match(flex_prefix_pattern, current_part, re.IGNORECASE):
                                # Look ahead to see if we need to combine more parts (prefix + number + road name)
                                if i < len(address_parts) - 2 and re.match(r'^\d+[\-\d]*$', address_parts[i+1]) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+2], re.IGNORECASE):
                                    # We have a pattern like: [FLEX-E-RE], [84-90], [BRADES ROAD]
                                    combined_parts.append(f"{current_part} {address_parts[i+1]} {address_parts[i+2]}")
                                    i += 3  # Skip all three parts
                                else:
                                    # Just combine with next part
                                    combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                    i += 2  # Skip both parts
                            # Check for standalone number that should be combined with street name
                            elif i < len(address_parts) - 1 and re.match(r'^\d+[\-\d]*$', current_part) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+1], re.IGNORECASE):
                                # We have a pattern like: [84-90], [BRADES ROAD]
                                combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                i += 2  # Skip both parts
                            else:
                                combined_parts.append(current_part)
                                i += 1
                        
                        # Replace original parts with combined parts
                        address_parts = combined_parts
                        
                        # Log the address parts after combining
                        log_debug(f"Address parts after combining prefixes and road numbers: {address_parts}")
                        
                        # Initialize address fields
                        addr1 = ""
                        addr2 = ""
                        addr3 = ""
                        addr4 = ""
                        
                        # Process address parts based on their content and position
                        if len(address_parts) > 0:
                            addr1 = address_parts[0]  # First part is always ADDR1 (usually business name)
                        
                        if len(address_parts) > 1:
                            # Check if the second part looks like a road/street with numbers
                            road_pattern = r'^\d+[\-\d]*\s+[A-Za-z\s]+'  # Matches patterns like "84-90 BRADES ROAD"
                            if len(address_parts) > 2 and (re.match(road_pattern, address_parts[1]) or 
                                                          re.search(r'\b(Road|Street|Avenue|Lane|Drive|Close)\b', address_parts[1], re.IGNORECASE)):
                                # This is a street address, keep it in ADDR2
                                addr2 = address_parts[1]
                                
                                # The rest goes into ADDR3 and ADDR4
                                if len(address_parts) > 3:
                                    addr3 = address_parts[2]
                                    addr4 = address_parts[3]
                                elif len(address_parts) > 2:
                                    addr3 = address_parts[2]
                            else:
                                # Standard processing
                                if len(address_parts) > 2:
                                    addr2 = address_parts[1]
                                    addr3 = address_parts[2]
                                    if len(address_parts) > 3:
                                        addr4 = ' '.join(address_parts[3:])
                                else:
                                    addr2 = address_parts[1]
                        
                        # Special handling for when ADDR4 is empty but we have a town in ADDR3
                        if not addr4 and addr3 and len(addr3.split()) <= 2:
                            addr4 = addr3
                            addr3 = ""
                        
                        # Set the address fields
                        job['DELIVERY ADDR1'] = addr1
                        job['DELIVERY ADDR2'] = addr2
                        job['DELIVERY ADDR3'] = addr3
                        job['DELIVERY ADDR4'] = addr4
                    
                    jobs.append(job)
                except Exception as item_error:
                    # Log individual item parsing errors but continue with others
                    print(f"Error processing item: {str(item_error)}")
                    continue
            
            if not jobs:
                messagebox.showerror("Error", "No valid jobs to process.")
                return
            
            # Count jobs by type for reporting purposes
            gr11_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR11')
            gr15_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR15')
            
            # Generate a single CSV file for all jobs
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"greenhous_jobs_{timestamp}.csv"
            
            # Save to the GR11 folder for consistency
            self.save_to_csv(jobs, filename, 'GR11')
            
            # Success message
            total_jobs = len(jobs)
            
            self.gr_status_var.set(f"Loaded {job_count} jobs from file")
            self.gr_status_label.config(fg=self.success_color)
            
            # More detailed success message with counts of each job type
            message = f"Processed {total_jobs} Greenhous jobs in total:\n"
            
            # Count by job type
            job_types = {}
            for job in jobs:
                customer_ref = job.get('CUSTOMER REF', 'Unknown')
                job_types[customer_ref] = job_types.get(customer_ref, 0) + 1
            
            # Add counts to message
            for ref, count in job_types.items():
                message += f"- {count} {ref} jobs\n"
                
            message += f"\nAll jobs saved to: jobs/GR11/{filename}"
                
            messagebox.showinfo("Success", message)
            
        except Exception as e:
            self.gr_status_var.set(f"Error occurred: {str(e)}")
            self.gr_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
    def create_coming_soon_tab(self, job_type):
        """Create a 'Coming Soon' tab content for future job types"""
        frame = tk.Frame(self.tab_content, bg=self.card_bg)
        
        # Center the coming soon content
        coming_soon_frame = tk.Frame(frame, bg=self.card_bg)
        coming_soon_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # Icon for coming soon
        icon_canvas = tk.Canvas(coming_soon_frame, width=80, height=80, 
                             bg=self.card_bg, highlightthickness=0)
        icon_canvas.pack(pady=20)
        
        # Draw clock or construction icon
        icon_canvas.create_oval(10, 10, 70, 70, fill="#F5F7FA", outline=self.primary_color, width=2)
        # Clock hands
        icon_canvas.create_line(40, 40, 40, 20, fill=self.primary_color, width=2)
        icon_canvas.create_line(40, 40, 55, 50, fill=self.primary_color, width=2)
        
        # Coming soon text
        title = tk.Label(
            coming_soon_frame,
            text=f"{job_type} Transport Processing",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        title.pack(pady=10)
        
        message = tk.Label(
            coming_soon_frame,
            text="Coming Soon",
            font=('Segoe UI', 16, 'bold'),
            fg=self.primary_color,
            bg=self.card_bg
        )
        message.pack(pady=20)
        
        description = tk.Label(
            coming_soon_frame,
            text=f"Support for {job_type} transport jobs is currently under development.\nCheck back soon for updates.",
            font=self.normal_font,
            fg=self.light_text,
            bg=self.card_bg,
            justify=tk.CENTER
        )
        description.pack()
        
        return frame
        
    def calculate_delivery_date(self, collection_date):
        """Calculate delivery date as 3 business days from collection date."""
        # Get UK holidays
        uk_holidays = holidays.UK()
        
        # Start with collection date
        current_date = collection_date
        business_days = 0
        
        # Keep adding days until we have 3 business days
        while business_days < 3:
            current_date += timedelta(days=1)
            # Skip weekends and holidays
            if current_date.weekday() < 5 and current_date not in uk_holidays:
                business_days += 1
                
        return current_date
    
    def update_delivery_date(self, event=None):
        """Update delivery date when collection date changes"""
        collection = self.collection_date.get_date()
        delivery = self.calculate_delivery_date(collection)
        self.delivery_date.set_date(delivery)
    
    def update_job_count(self, event=None):
        """Count the number of jobs in the text input and update the counter label"""
        text = self.text_input.get("1.0", tk.END)
        
        # Count jobs by counting "FROM" sections
        job_count = text.count("\nFROM\n")
        
        # Add one more if the text starts with "FROM"
        if text.lstrip().startswith("FROM"):
            job_count += 1
        
        # Update the counter
        if job_count == 1:
            self.job_count_var.set("1 job detected")
        else:
            self.job_count_var.set(f"{job_count} jobs detected")
    
    def process_jobs(self):
        try:
            # Update status
            self.status_var.set("Processing jobs...")
            self.status_label.config(fg=self.primary_color)
            self.root.update_idletasks()
            
            # Get the input text
            text = self.text_input.get("1.0", tk.END)
            
            # Get dates directly from the date pickers
            collection_date = self.collection_date.get_date().strftime("%d/%m/%Y")
            delivery_date = self.delivery_date.get_date().strftime("%d/%m/%Y")
            
            # Parse the jobs
            parser = JobParser(collection_date, delivery_date)
            jobs = parser.parse_jobs(text)

            # Debug: Log each job's REG NUMBER
            for i, job in enumerate(jobs, 1):
                reg_val = job.get('REG NUMBER')
                print(f"Job {i} REG NUMBER: '{reg_val}'")
                log_debug(f"Job {i} REG NUMBER: '{reg_val}'")

            if not jobs:
                self.status_var.set("No valid jobs found in the input text")
                self.status_label.config(fg=self.warning_color)
                messagebox.showerror("Error", "No valid jobs found in the input text")
                return

            # Count jobs and registrations (robust)
            job_count = len(jobs)
            reg_count = sum(1 for job in jobs if isinstance(job.get('REG NUMBER'), str) and job.get('REG NUMBER').strip())

            # Update the job count label to show both
            if job_count == 1:
                self.job_count_var.set(f"1 job detected, {reg_count} registrations detected")
            else:
                self.job_count_var.set(f"{job_count} jobs detected, {reg_count} registrations detected")
            
            # For AC01 jobs, validate registrations
            if self.current_tab == "AC01":
                missing_regs = []
                for i, job in enumerate(jobs, 1):
                    if not job.get('REG NUMBER') or job['REG NUMBER'].strip() == '':
                        missing_regs.append(f"Job #{i}")
                
                if missing_regs:
                    missing_msg = f"The following jobs are missing registration numbers:\n{', '.join(missing_regs)}"
                    self.status_var.set("Warning: Some jobs missing registration numbers")
                    self.status_label.config(fg=self.warning_color)
                    messagebox.showwarning("Missing Registrations", missing_msg)
            
            # Get the current job type (tab)
            job_type = self.current_tab
            
            # Generate CSV file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{job_type.lower()}_jobs_{timestamp}.csv"
            
            try:
                self.save_to_csv(jobs, filename, job_type)
                
                # Calculate total price for this batch of jobs
                batch_total = 0.0
                for job in jobs:
                    if 'PRICE' in job and job['PRICE']:
                        try:
                            # Handle various price formats
                            price_str = job['PRICE'].strip()
                            price_str = price_str.replace('£', '').replace(',', '')
                            price = float(price_str)
                            batch_total += price
                        except (ValueError, TypeError):
                            log_debug(f"Could not parse price: {job.get('PRICE', 'N/A')}")
                
                # Update the total price if this is AC01
                if job_type == "AC01":
                    self.total_ac01_price += batch_total
                    self.ac01_price_var.set(f"Total Price: £{self.total_ac01_price:.2f}")
                
                self.status_var.set(f"Successfully processed {len(jobs)} jobs")
                self.status_label.config(fg=self.success_color)
                
                # Include price in success message
                if batch_total > 0:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} {job_type} jobs\nBatch Total: £{batch_total:.2f}\nRunning Total: £{self.total_ac01_price:.2f}")
                else:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} {job_type} jobs")
                
            except Exception as e:
                self.status_var.set("Error: Failed to save CSV file")
                self.status_label.config(fg=self.error_color)
                messagebox.showerror("Error", f"Failed to save CSV file: {str(e)}")
            
        except Exception as e:
            self.status_var.set("Error occurred during processing")
            self.status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_to_csv(self, jobs, filename, job_type):
        if not jobs:
            return

        try:
            # Clear debug log before saving
            log_debug("\n--- START SAVING JOBS TO CSV ---")
            
            # Get the current script directory
            current_dir = os.path.dirname(os.path.abspath(__file__))
            log_debug(f"Current directory: {current_dir}")
            
            # Go up one level to get the workspace root directory
            workspace_root = os.path.dirname(current_dir)
            log_debug(f"Workspace root: {workspace_root}")
            
            # Create main 'jobs' directory if it doesn't exist
            output_dir = os.path.join(workspace_root, 'jobs')
            os.makedirs(output_dir, exist_ok=True)
            log_debug(f"Main jobs directory: {output_dir}")
            
            # Create job type specific directory if it doesn't exist
            job_type_dir = os.path.join(output_dir, job_type)
            os.makedirs(job_type_dir, exist_ok=True)
            log_debug(f"Job type directory: {job_type_dir}")
            
            # Full path for the output file
            output_path = os.path.join(job_type_dir, filename)
            log_debug(f"Output file path: {output_path}")
            
            # Debug: Count original job types
            gr11_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR11')
            gr15_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR15')
            log_debug(f"Job counts before saving - GR11: {gr11_count}, GR15: {gr15_count}")
            
            # Log all customer references before saving
            log_debug("Customer references before saving:")
            for i, job in enumerate(jobs):
                log_debug(f"Job {i} - Reg: {job.get('REG NUMBER', '')}, Customer Ref: {job.get('CUSTOMER REF', '')}")
            
            # IMPORTANT: For Greenhous jobs NEVER override the customer reference
            # For other job types, set all jobs to the same customer ref
            if job_type != 'GR11' and not job_type.startswith('GR'):
                # Set all jobs to the same customer ref for non-Greenhous jobs
                for job in jobs:
                    job['CUSTOMER REF'] = job_type
            
            # Debug: Check final job types
            gr11_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR11')
            gr15_count = sum(1 for job in jobs if job.get('CUSTOMER REF', '') == 'GR15')
            log_debug(f"After customer ref setup - GR11: {gr11_count}, GR15: {gr15_count}")
            
            # Check if the output file already exists
            if os.path.exists(output_path):
                log_debug(f"Warning: Output file already exists, will try to overwrite: {output_path}")
                # Check if we can write to it
                try:
                    with open(output_path, 'a') as f:
                        pass  # Just testing if we can open it for append
                except PermissionError:
                    log_debug(f"Permission denied when trying to open existing file: {output_path}")
                    raise PermissionError(f"The file {filename} is open in another program. Please close it and try again.")
            
            # Test folder permissions
            test_file = os.path.join(job_type_dir, "_test_write_permission.tmp")
            try:
                with open(test_file, 'w') as f:
                    f.write("test")
                os.remove(test_file)
                log_debug(f"Write permission test successful on directory: {job_type_dir}")
            except Exception as e:
                log_debug(f"Failed write permission test: {str(e)}")
                raise PermissionError(f"Cannot write to the output directory ({job_type_dir}). Please check your permissions.")
            
            # Try to save the file
            log_debug(f"Attempting to write CSV to: {output_path}")
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = [
                    'REG NUMBER',
                    'VIN',
                    'MAKE',
                    'MODEL',
                    'COLLECTION DATE',
                    'YOUR REF NO',
                    'COLLECTION ADDR1',
                    'COLLECTION ADDR2',
                    'COLLECTION ADDR3',
                    'COLLECTION ADDR4',
                    'COLLECTION POSTCODE',
                    'COLLECTION CONTACT NAME',
                    'COLLECTION PHONE',
                    'DELIVERY DATE',
                    'DELIVERY ADDR1',
                    'DELIVERY ADDR2',
                    'DELIVERY ADDR3',
                    'DELIVERY ADDR4',
                    'DELIVERY POSTCODE',
                    'DELIVERY CONTACT NAME',
                    'DELIVERY CONTACT PHONE',
                    'SPECIAL INSTRUCTIONS',
                    'PRICE',
                    'CUSTOMER REF',
                    'TRANSPORT TYPE'
                ]

                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                log_debug("Writing to CSV file:")
                for job in jobs:
                    # Create a new dict with only the fields we want
                    row = {field: job.get(field, '') for field in fieldnames}
                    # Debug: Check individual job customer reference
                    log_debug(f"Writing job {row['REG NUMBER']} with customer ref: {row['CUSTOMER REF']}")
                    writer.writerow(row)
                
            log_debug(f"CSV file saved to: {output_path}")
            log_debug("--- END SAVING JOBS TO CSV ---")
            
            self.status_var.set(f"File saved: {filename}")
            self.last_saved_file = output_path
            self.file_link_var.set(f"View file: {filename}")
            self.file_link_button.pack(side=tk.LEFT, padx=10)
            print("File link button packed and visible")
            messagebox.showinfo("Success", f"File saved successfully at:\n{output_path}")

        except PermissionError as pe:
            log_debug(f"Permission error during save: {str(pe)}")
            error_msg = f"Permission denied. Please ensure:\n1. No other program has the file open\n2. You have permission to write to the folder: {os.path.dirname(output_path) if 'output_path' in locals() else 'unknown'}\n3. Close Excel or any CSV viewers before trying again."
            messagebox.showerror("Permission Error", error_msg)
            raise pe
        except Exception as e:
            log_debug(f"Error during save: {str(e)}")
            messagebox.showerror("Error", f"An error occurred while saving the file:\n{str(e)}")
            raise e
    
    def process_gr_jobs(self):
        """Process the GR11/GR15 jobs from the Excel data"""
        try:
            # Get all items from treeview
            items = self.job_tree.get_children()
            
            if not items:
                messagebox.showerror("Error", "No jobs to process. Please load an Excel file first.")
                return
            
            # Collect job data
            jobs = []
            for item in items:
                try:
                    values = self.job_tree.item(item, "values")
                    
                    # Check if we have the minimum required data
                    if not values[0]:  # Reg No
                        continue
                    
                    # Get the customer ref and PDI Centre values
                    # values[1] is the customer ref from the preview (GR11 or GR15)
                    customer_ref = values[1] if len(values) > 1 and values[1] else "GR11"
                    
                    # values[8] is the PDI Centre column
                    pdi_centre = values[8] if len(values) > 8 and values[8] else ""
                    
                    # Debug: Print the PDI Centre value
                    print(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    log_debug(f"Processing job - Customer Ref: '{customer_ref}', PDI Centre: '{pdi_centre}'")
                    
                    # Collection address details
                    collection_addr1 = ""
                    collection_addr2 = ""
                    collection_addr3 = ""
                    collection_addr4 = ""
                    collection_postcode = ""
                    
                    # Use the customer reference already set in the preview 
                    # This ensures consistency with what's shown in the UI
                    
                    # Set address based on the customer reference
                    if customer_ref == "GR15":
                        collection_addr1 = "Greenhous Upper Heyford"
                        collection_addr2 = "Heyford Park, Bicester"
                        collection_addr3 = "Bicester"
                        collection_addr4 = "UPPER HEYFORD"
                        collection_postcode = "OX25 5HA"
                        print(f"  => Using Upper Heyford address")
                        log_debug(f"  => Using Upper Heyford address")
                    else:
                        # Default to High Ercall (GR11)
                        collection_addr1 = "Greenhous High Ercall"
                        collection_addr2 = "Greenhous Village Osbaston"
                        collection_addr3 = "High Ercall"
                        collection_addr4 = ""
                        collection_postcode = "TF6 6RA"
                        print(f"  => Using High Ercall address")
                        log_debug(f"  => Using High Ercall address")
                    
                    # Create job dictionary
                    job = {
                        'REG NUMBER': values[0],
                        'VIN': values[2] if len(values) > 2 and values[2] else "",  # Chassis
                        'MAKE': values[3] if len(values) > 3 and values[3] else "TRANSIT",  # Make
                        'MODEL': values[4] if len(values) > 4 and values[4] else "VAN",  # Model
                        'COLOR': "",
                        'COLLECTION DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Collection date
                        'YOUR REF NO': values[0],  # Use reg number as reference
                        'COLLECTION ADDR1': collection_addr1,
                        'COLLECTION ADDR2': collection_addr2,
                        'COLLECTION ADDR3': collection_addr3,
                        'COLLECTION ADDR4': collection_addr4,
                        'COLLECTION POSTCODE': collection_postcode,
                        'COLLECTION CONTACT NAME': "",
                        'COLLECTION PHONE': "",
                        'DELIVERY DATE': values[5] if len(values) > 5 and values[5] else datetime.now().strftime("%d/%m/%Y"),  # Same as collection date
                        'DELIVERY ADDR1': "",
                        'DELIVERY ADDR2': "",
                        'DELIVERY ADDR3': "",
                        'DELIVERY ADDR4': "",
                        'DELIVERY POSTCODE': "",
                        'DELIVERY CONTACT NAME': "",
                        'DELIVERY CONTACT PHONE': "",
                        'SPECIAL INSTRUCTIONS': f"VIN: {values[2] if len(values) > 2 and values[2] else 'Unknown'}",
                        'PRICE': "",
                        'CUSTOMER REF': customer_ref,
                        'TRANSPORT TYPE': ""
                    }
                    
                    # Parse delivery address
                    if len(values) > 7 and values[7]:
                        delivery_address = values[7]  # Correct index for delivery address
                        
                        # Pre-process common patterns that should be kept together
                        # Examples: "FLEX E REN 84-90", "FLEX E RE WEST LON"
                        patterns_to_fix = [
                            (r'(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)[\s,]+(\d+[\-\d]*)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(WEST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ST)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(UNIT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(MARCH)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(IVATT)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(LANDGATI)', r'\1 \2'),
                            (r'(FLEX[\-\s]E[\-\s]RE)[\s,]+(ASCOT)', r'\1 \2'),
                            # New patterns to keep road numbers with road names
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:Road|Street|Avenue|Lane|Drive|Close|Way|ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY))', r'\1 \2'),
                            # Handle specific cases like 84-90 BRADES ROAD
                            (r'(\d+[\-\d]*)\s*,\s*(BRADES[\s]*ROAD)', r'\1 \2'),
                            # More general pattern for any number + street pattern that got split
                            (r'(\d+[\-\d]*)\s*,\s*([A-Za-z\s]+(?:\s(?:Rd|St|Ave|Ln|Dr|Cl|RD|ST|AVE|LN|DR|CL)))', r'\1 \2')
                        ]
                        
                        # Apply each pattern fix
                        processed_address = delivery_address
                        for pattern, replacement in patterns_to_fix:
                            processed_address = re.sub(pattern, replacement, processed_address, flags=re.IGNORECASE)
                        
                        # Log any changes made during pre-processing
                        if processed_address != delivery_address:
                            log_debug(f"Pre-processed address: '{delivery_address}' -> '{processed_address}'")
                            delivery_address = processed_address
                        
                        # Try to extract postcode from address (UK format)
                        postcode_pattern = r'([A-Z]{1,2}[0-9][0-9A-Z]?\s*[0-9][A-Z]{2})'
                        postcode_match = re.search(postcode_pattern, delivery_address, re.IGNORECASE)
                        
                        # Clean the postcode if found
                        extracted_postcode = ""
                        if postcode_match:
                            extracted_postcode = postcode_match.group(1).upper().strip()
                            # Clean up spacing in the postcode
                            if len(extracted_postcode) > 4 and ' ' not in extracted_postcode:
                                # Add space in the correct position for UK postcodes (before last 3 chars)
                                extracted_postcode = extracted_postcode[:-3] + ' ' + extracted_postcode[-3:]
                            job['DELIVERY POSTCODE'] = extracted_postcode
                        
                        # Remove the postcode from the address for cleaner parsing
                        if extracted_postcode:
                            clean_address = delivery_address.replace(postcode_match.group(0), "")
                        else:
                            clean_address = delivery_address
                        
                        # Enhanced address parsing to keep road numbers with street names
                        # First split by commas
                        address_parts = [part.strip() for part in clean_address.split(',') if part.strip()]
                        
                        # First pass: Check for patterns like "FLEX-E-RE 84-90" being split across parts
                        # or combine prefix with following parts
                        combined_parts = []
                        i = 0
                        while i < len(address_parts):
                            current_part = address_parts[i]
                            
                            # Special patterns to check for
                            flex_prefix_pattern = r'^(FLEX[\-\s]E[\-\s]RE|FLEX[\-\s]RENT|FLEX[\-\s]E[\-\s]REN|FLEX[\-\s]E[\-\s]RE[\-\s]ST|ENTERPRI[\-\s]ROSS)$'
                            
                            # Check if this part ends with a prefix that should be combined with the next part
                            if i < len(address_parts) - 1 and re.match(flex_prefix_pattern, current_part, re.IGNORECASE):
                                # Look ahead to see if we need to combine more parts (prefix + number + road name)
                                if i < len(address_parts) - 2 and re.match(r'^\d+[\-\d]*$', address_parts[i+1]) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+2], re.IGNORECASE):
                                    # We have a pattern like: [FLEX-E-RE], [84-90], [BRADES ROAD]
                                    combined_parts.append(f"{current_part} {address_parts[i+1]} {address_parts[i+2]}")
                                    i += 3  # Skip all three parts
                                else:
                                    # Just combine with next part
                                    combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                    i += 2  # Skip both parts
                            # Check for standalone number that should be combined with street name
                            elif i < len(address_parts) - 1 and re.match(r'^\d+[\-\d]*$', current_part) and re.match(r'^[A-Za-z\s]+(?:ROAD|STREET|AVENUE|LANE|DRIVE|CLOSE|WAY|Road|Street|Avenue|Lane|Drive|Close|Way)$', address_parts[i+1], re.IGNORECASE):
                                # We have a pattern like: [84-90], [BRADES ROAD]
                                combined_parts.append(f"{current_part} {address_parts[i+1]}")
                                i += 2  # Skip both parts
                            else:
                                combined_parts.append(current_part)
                                i += 1
                        
                        # Replace original parts with combined parts
                        address_parts = combined_parts
                        
                        # Log the address parts after combining
                        log_debug(f"Address parts after combining prefixes and road numbers: {address_parts}")
                        
                        # Initialize address fields
                        addr1 = ""
                        addr2 = ""
                        addr3 = ""
                        addr4 = ""
                        
                        # Process address parts based on their content and position
                        if len(address_parts) > 0:
                            addr1 = address_parts[0]  # First part is always ADDR1 (usually business name)
                        
                        if len(address_parts) > 1:
                            # Check if the second part looks like a road/street with numbers
                            road_pattern = r'^\d+[\-\d]*\s+[A-Za-z\s]+'  # Matches patterns like "84-90 BRADES ROAD"
                            if len(address_parts) > 2 and (re.match(road_pattern, address_parts[1]) or 
                                                          re.search(r'\b(Road|Street|Avenue|Lane|Drive|Close)\b', address_parts[1], re.IGNORECASE)):
                                # This is a street address, keep it in ADDR2
                                addr2 = address_parts[1]
                                
                                # The rest goes into ADDR3 and ADDR4
                                if len(address_parts) > 3:
                                    addr3 = address_parts[2]
                                    addr4 = address_parts[3]
                                elif len(address_parts) > 2:
                                    addr3 = address_parts[2]
                            else:
                                # Standard processing
                                if len(address_parts) > 2:
                                    addr2 = address_parts[1]
                                    addr3 = address_parts[2]
                                    if len(address_parts) > 3:
                                        addr4 = ' '.join(address_parts[3:])
                                else:
                                    addr2 = address_parts[1]
                        
                        # Special handling for when ADDR4 is empty but we have a town in ADDR3
                        if not addr4 and addr3 and len(addr3.split()) <= 2:
                            addr4 = addr3
                            addr3 = ""
                        
                        # Set the address fields
                        job['DELIVERY ADDR1'] = addr1
                        job['DELIVERY ADDR2'] = addr2
                        job['DELIVERY ADDR3'] = addr3
                        job['DELIVERY ADDR4'] = addr4
                    
                    jobs.append(job)
                except Exception as item_error:
                    # Log individual item parsing errors but continue with others
                    print(f"Error processing item: {str(item_error)}")
                    continue
            
            if not jobs:
                messagebox.showerror("Error", "No valid jobs to process.")
                return
            
            # Count jobs by type for reporting purposes
            gr11_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR11')
            gr15_count = sum(1 for job in jobs if job['CUSTOMER REF'] == 'GR15')
            
            # Generate a single CSV file for all jobs
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"greenhous_jobs_{timestamp}.csv"
            
            # Save to the GR11 folder for consistency
            self.save_to_csv(jobs, filename, 'GR11')
            
            # Success message
            total_jobs = len(jobs)
            
            self.gr_status_var.set(f"Loaded {job_count} jobs from file")
            self.gr_status_label.config(fg=self.success_color)
            
            # More detailed success message with counts of each job type
            message = f"Processed {total_jobs} Greenhous jobs in total:\n"
            
            # Count by job type
            job_types = {}
            for job in jobs:
                customer_ref = job.get('CUSTOMER REF', 'Unknown')
                job_types[customer_ref] = job_types.get(customer_ref, 0) + 1
            
            # Add counts to message
            for ref, count in job_types.items():
                message += f"- {count} {ref} jobs\n"
                
            message += f"\nAll jobs saved to: jobs/GR11/{filename}"
                
            messagebox.showinfo("Success", message)
            
        except Exception as e:
            self.gr_status_var.set(f"Error occurred: {str(e)}")
            self.gr_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            
    def open_last_saved_file(self, event=None):
        print("open_last_saved_file called")
        import subprocess
        import os
        print(f"Trying to open: {self.last_saved_file}")
        if self.last_saved_file and os.path.exists(self.last_saved_file):
            subprocess.Popen(['explorer', f'/select,{self.last_saved_file}'])
        else:
            print("File does not exist or path is not set.")

    def create_cw08_09_tab(self):
        """Create the CW08/09 tab content for Excel file processing"""
        cw_frame = tk.Frame(self.tab_content, bg=self.card_bg)
        # Excel import section
        import_card = ModernFrame(
            cw_frame,
            bordercolor=self.border_color,
            bgcolor="#F9F9FC",
            borderwidth=1
        )
        import_card.pack(fill=tk.X, padx=0, pady=20)
        import_container = tk.Frame(import_card.content_frame, bg="#F9F9FC", padx=15, pady=15)
        import_container.pack(fill=tk.X)
        # File selection
        file_frame = tk.Frame(import_container, bg="#F9F9FC")
        file_frame.pack(fill=tk.X, pady=10)
        self.cw_selected_file_var = tk.StringVar()
        self.cw_selected_file_var.set("No file selected")
        file_label = tk.Label(
            file_frame,
            text="Excel File:",
            font=self.normal_font,
            fg=self.text_color,
            bg="#F9F9FC"
        )
        file_label.pack(side=tk.LEFT, padx=(0, 10))
        file_path_label = tk.Label(
            file_frame,
            textvariable=self.cw_selected_file_var,
            font=self.normal_font,
            fg=self.light_text,
            bg="#F9F9FC"
        )
        file_path_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        browse_button = tk.Button(
            file_frame,
            text="Browse",
            font=self.normal_font,
            bg=self.primary_color,
            fg="white",
            relief=tk.FLAT,
            padx=15,
            pady=5,
            command=self.browse_cw_excel_file
        )
        browse_button.pack(side=tk.RIGHT)
        # Job preview section
        preview_section = tk.Frame(cw_frame, bg=self.card_bg)
        preview_section.pack(fill=tk.BOTH, expand=True, pady=15)
        preview_header = tk.Frame(preview_section, bg=self.card_bg)
        preview_header.pack(fill=tk.X, pady=10)
        preview_title = tk.Label(
            preview_header,
            text="Job Preview",
            font=self.header_font,
            fg=self.text_color,
            bg=self.card_bg
        )
        preview_title.pack(side=tk.LEFT)
        self.cw_job_count_var = tk.StringVar()
        self.cw_job_count_var.set("0 jobs detected")
        job_count_label = tk.Label(
            preview_header,
            textvariable=self.cw_job_count_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg
        )
        job_count_label.pack(side=tk.RIGHT)
        preview_container = ModernFrame(
            preview_section,
            bordercolor=self.border_color,
            bgcolor=self.card_bg,
            borderwidth=1
        )
        preview_container.pack(fill=tk.BOTH, expand=True)
        tree_frame = tk.Frame(preview_container.content_frame, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True)
        tree_scroll = tk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        columns = (
            "reg_no", "vin", "make", "model", "collection_date", "your_ref", "collection_addr1", "collection_addr2", "collection_addr3", "collection_addr4", "collection_postcode", "delivery_addr1", "delivery_addr2", "delivery_addr3", "delivery_addr4", "delivery_postcode", "special_instructions", "price"
        )
        self.cw_job_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", yscrollcommand=tree_scroll.set)
        for col in columns:
            self.cw_job_tree.heading(col, text=col.replace('_', ' ').upper())
            self.cw_job_tree.column(col, width=100)
        self.cw_job_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.cw_job_tree.yview)
        action_section = tk.Frame(cw_frame, bg=self.card_bg)
        action_section.pack(fill=tk.X, pady=10)
        self.cw_status_var = tk.StringVar()
        self.cw_status_label = tk.Label(
            action_section,
            textvariable=self.cw_status_var,
            font=self.small_font,
            fg=self.light_text,
            bg=self.card_bg,
            anchor=tk.W
        )
        self.cw_status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cw_process_button = RoundedButton(
            action_section,
            width=150,
            height=36,
            cornerradius=18,
            padding=8,
            color=self.primary_color,
            hover_color=self.secondary_color,
            text="Process Jobs",
            command=self.process_cw_jobs,
            fg="white"
        )
        self.cw_process_button.pack(side=tk.RIGHT)
        return cw_frame

    def browse_cw_excel_file(self):
        filetypes = [("Excel/CSV files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.cw_selected_file_var.set(file_path)
            self.load_cw_excel_data(file_path)

    def load_cw_excel_data(self, file_path):
        import pandas as pd
        import os
        try:
            for item in self.cw_job_tree.get_children():
                self.cw_job_tree.delete(item)
            ext = os.path.splitext(file_path)[1].lower()
            if ext == '.csv':
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            job_count = 0
            for _, row in df.iterrows():
                # Replace nan with empty string
                row = row.fillna("") if hasattr(row, 'fillna') else row
                # Slide delivery address fields and put contact name in addr1 and delivery contact name
                delivery_contact_name = row.get('Delivery Contact Name', '')
                delivery_addr1 = delivery_contact_name
                delivery_addr2 = row.get('Delivery Address1', '')
                delivery_addr3 = row.get('Delivery Address2', '')
                delivery_addr4 = row.get('Delivery Address3', '')
                delivery_contact_phone = row.get('Delivery Contact Phone', '')
                # Insert into treeview, add delivery contact name as last column for later use
                values = [
                    row.get('Reg', ''),
                    row.get('VIN', ''),
                    row.get('Make', ''),
                    row.get('Model', ''),
                    row.get('Collection Date', ''),
                    row.get('Your Ref No', ''),
                    row.get('Collection Address1', ''),
                    row.get('Collection Address2', ''),
                    row.get('Collection Address3', ''),
                    row.get('Collection Address4', ''),
                    row.get('Collection Postcode', ''),
                    delivery_addr1,
                    delivery_addr2,
                    delivery_addr3,
                    delivery_addr4,
                    row.get('Delivery Postcode', ''),
                    row.get('SpecialInstructions', ''),
                    row.get('Price', ''),
                    delivery_contact_phone,
                    delivery_contact_name  # for delivery contact name field
                ]
                # Replace any 'nan' string values with ''
                values = [v if (v != 'nan' and str(v).lower() != 'nan') else '' for v in values]
                self.cw_job_tree.insert("", tk.END, values=values)
                job_count += 1
            if job_count == 1:
                self.cw_job_count_var.set("1 job detected")
            else:
                self.cw_job_count_var.set(f"{job_count} jobs detected")
            self.cw_status_var.set(f"Loaded {job_count} jobs from file")
            self.cw_status_label.config(fg=self.success_color)
        except Exception as e:
            self.cw_status_var.set(f"Error loading file: {str(e)}")
            self.cw_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")

    def process_cw_jobs(self):
        items = self.cw_job_tree.get_children()
        if not items:
            messagebox.showerror("Error", "No jobs to process. Please load an Excel file first.")
            return
        jobs = []
        for item in items:
            values = list(self.cw_job_tree.item(item, "values"))
            # Replace any 'nan' string values with ''
            values = [v if (v != 'nan' and str(v).lower() != 'nan') else '' for v in values]
            jobs.append({
                'REG NUMBER': values[0],
                'VIN': values[1],
                'MAKE': values[2],
                'MODEL': values[3],
                'COLLECTION DATE': values[4],
                'YOUR REF NO': values[5],
                'COLLECTION ADDR1': values[6],
                'COLLECTION ADDR2': values[7],
                'COLLECTION ADDR3': values[8],
                'COLLECTION ADDR4': values[9],
                'COLLECTION POSTCODE': values[10],
                'DELIVERY ADDR1': values[11],  # Delivery Contact Name
                'DELIVERY ADDR2': values[12],  # Delivery Address1
                'DELIVERY ADDR3': values[13],  # Delivery Address2
                'DELIVERY ADDR4': values[14],  # Delivery Address3
                'DELIVERY POSTCODE': values[15],
                'SPECIAL INSTRUCTIONS': values[16],
                'PRICE': values[17],
                'DELIVERY CONTACT PHONE': values[18],
                'DELIVERY CONTACT NAME': values[19],
                'CUSTOMER REF': 'CW09',
                'TRANSPORT TYPE': ''
            })
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"cw09_jobs_{timestamp}.csv"
        self.save_to_csv(jobs, filename, 'CW09')
        self.cw_status_var.set(f"Successfully processed {len(jobs)} jobs")
        self.cw_status_label.config(fg=self.success_color)

    def reset_ac01_price(self):
        """Reset the total AC01 price counter."""
        self.total_ac01_price = 0.0
        self.ac01_price_var.set("Total Price: £0.00")
        messagebox.showinfo("Price Reset", "The total price counter has been reset to £0.00")

    def update_bc04_delivery_date(self, event=None):
        """Update BC04 delivery date when collection date changes"""
        collection_date = self.bc04_collection_date.get_date()
        delivery_date = self.calculate_delivery_date(collection_date)
        self.bc04_delivery_date.set_date(delivery_date)

    def update_bc04_job_count(self, event=None):
        """Update BC04 job count as user types"""
        try:
            text = self.bc04_text_input.get("1.0", tk.END)
            # Count "Job Sheet" occurrences to estimate job count
            job_count = len(re.findall(r'Job Sheet', text, re.IGNORECASE))
            if job_count == 1:
                self.bc04_job_count_var.set("1 job detected")
            else:
                self.bc04_job_count_var.set(f"{job_count} jobs detected")
        except:
            self.bc04_job_count_var.set("0 jobs detected")

    def process_bc04_jobs(self):
        """Process BC04 jobs using the BC04Parser"""
        try:
            # Update status
            self.bc04_status_var.set("Processing BC04 jobs...")
            self.bc04_status_label.config(fg=self.primary_color)
            self.root.update_idletasks()
            
            # Get the input text
            text = self.bc04_text_input.get("1.0", tk.END)
            
            # Get dates directly from the date pickers
            collection_date = self.bc04_collection_date.get_date().strftime("%d/%m/%Y")
            delivery_date = self.bc04_delivery_date.get_date().strftime("%d/%m/%Y")
            
            # Parse the jobs using BC04Parser
            parser = BC04Parser(collection_date, delivery_date)
            jobs = parser.parse_jobs(text)
            
            if not jobs:
                self.bc04_status_var.set("No valid BC04 jobs found in the input text")
                self.bc04_status_label.config(fg=self.warning_color)
                messagebox.showerror("Error", "No valid BC04 jobs found in the input text")
                return
            
            # Update the job count to match the actual parsed jobs
            job_count = len(jobs)
            if job_count == 1:
                self.bc04_job_count_var.set("1 job detected")
            else:
                self.bc04_job_count_var.set(f"{job_count} jobs detected")
            
            # Generate CSV file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"bc04_jobs_{timestamp}.csv"
            
            try:
                self.save_to_csv(jobs, filename, "BC04")
                
                # Calculate total price for this batch of jobs
                batch_total = 0.0
                for job in jobs:
                    if 'PRICE' in job and job['PRICE']:
                        try:
                            # Handle various price formats
                            price_str = job['PRICE'].strip()
                            price_str = price_str.replace('£', '').replace(',', '')
                            price = float(price_str)
                            batch_total += price
                        except (ValueError, TypeError):
                            log_debug(f"Could not parse BC04 price: {job.get('PRICE', 'N/A')}")
                
                # Update the total price for BC04
                if not hasattr(self, 'total_bc04_price'):
                    self.total_bc04_price = 0.0
                    self.bc04_price_var = tk.StringVar()
                    self.bc04_price_var.set("Total Price: £0.00")
                
                self.total_bc04_price += batch_total
                self.bc04_price_var.set(f"Total Price: £{self.total_bc04_price:.2f}")
                
                self.bc04_status_var.set(f"Successfully processed {len(jobs)} BC04 jobs")
                self.bc04_status_label.config(fg=self.success_color)
                
                # Include price in success message
                if batch_total > 0:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} BC04 jobs\nBatch Total: £{batch_total:.2f}\nRunning Total: £{self.total_bc04_price:.2f}")
                else:
                    messagebox.showinfo("Success", f"Processed {len(jobs)} BC04 jobs")
                
            except Exception as e:
                self.bc04_status_var.set("Error: Failed to save CSV file")
                self.bc04_status_label.config(fg=self.error_color)
                messagebox.showerror("Error", f"Failed to save CSV file: {str(e)}")
            
        except Exception as e:
            self.bc04_status_var.set("Error occurred during BC04 processing")
            self.bc04_status_label.config(fg=self.error_color)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def reset_bc04_price(self):
        """Reset the total BC04 price counter."""
        if not hasattr(self, 'total_bc04_price'):
            self.total_bc04_price = 0.0
            self.bc04_price_var = tk.StringVar()
            self.bc04_price_var.set("Total Price: £0.00")
        
        self.total_bc04_price = 0.0
        self.bc04_price_var.set("Total Price: £0.00")
        messagebox.showinfo("Price Reset", "The BC04 total price counter has been reset to £0.00")

    def is_valid_uk_registration(self, reg):
        reg = reg.upper().replace(" ", "")
        if reg.isdigit():
            return False  # Never just numbers
        patterns = [
            r"^[A-Z]{2}[0-9]{2}[A-Z]{3}$",         # Current style: AB12CDE
            r"^[A-Z]{1}[0-9]{1,3}[A-Z]{3}$",       # Prefix: A123BCD
            r"^[A-Z]{3}[0-9]{1,3}[A-Z]{1}$",       # Suffix: ABC123A
            r"^[A-Z]{3}[0-9]{4}$",                 # NI: BYZ3210
            r"^[0-9]{1,4}[A-Z]{1,3}$",             # Dateless: 1–4 numbers + 1–3 letters
            r"^[A-Z]{1,3}[0-9]{1,4}$",             # Dateless: 1–3 letters + 1–4 numbers
        ]
        return any(re.match(p, reg) for p in patterns)

    def parse_single_job(self, job_text):
        job = {}
        # ... existing code ...
        # After extracting registration (structured_match or fallback)
        # Replace:
        #     job['REG NUMBER'] = registration
        # With:
        if 'registration' in locals() and self.is_valid_uk_registration(registration):
            job['REG NUMBER'] = registration
        else:
            job['REG NUMBER'] = ''
        # ... existing code ...
        # Also update the fallback regex assignment:
        # if not job['REG NUMBER']:
        #     reg_match = re.search(r'REG(?:ISTRATION)?\s*:?
        #     if reg_match:
        #         job['REG NUMBER'] = reg_match.group(1).strip()
        # Should become:
        if not job['REG NUMBER']:
            reg_match = re.search(r'REG(?:ISTRATION)?\s*:?\s*([A-Z0-9]+)', job_text, re.IGNORECASE)
            if reg_match:
                reg_candidate = reg_match.group(1).strip()
                if self.is_valid_uk_registration(reg_candidate):
                    job['REG NUMBER'] = reg_candidate
                else:
                    job['REG NUMBER'] = ''

class BC04Parser:
    def __init__(self, collection_date, delivery_date=None):
        self.jobs = []
        self.collection_date = collection_date
        self.delivery_date = delivery_date if delivery_date else collection_date
        self.bc04_special_instructions = (
            "MUST GET A FULL NAME AND SIGNATURE ON COLLECTION CALL OFFICE AND Non Conformance Motability on 0121 788 6940 option 1 IF THEY REFUSE ** - PHOTO'S MUST BE CLEAR PLEASE. COLL AND DEL 09:00-17:00 ONLY"
        )
    def calculate_delivery_date(self, collection_date):
        """Calculate delivery date based on collection date (next business day)"""
        try:
            # Import holidays here to avoid issues if not available
            import holidays
            uk_holidays = holidays.UnitedKingdom()
        except ImportError:
            # Fallback if holidays package is not available
            class DummyHolidays:
                def __init__(self, *args, **kwargs):
                    pass
                def __contains__(self, date):
                    return False
            uk_holidays = DummyHolidays()
        
        # Start with the collection date
        delivery = collection_date
        
        # Add one day
        delivery += timedelta(days=1)
        
        # Skip weekends and holidays
        while delivery.weekday() >= 5 or delivery in uk_holidays:  # 5=Saturday, 6=Sunday
            delivery += timedelta(days=1)
        
        return delivery
    
    def clean_phone_number(self, phone):
        """Clean and format phone number"""
        if not phone:
            return ''
        
        # Remove common prefixes and clean
        phone = phone.strip()
        phone = re.sub(r'^(Tel|Phone|T|Telephone)[\s:.]*', '', phone, flags=re.IGNORECASE)
        phone = re.sub(r'[^\d+\s()-]', '', phone)  # Keep only digits, +, spaces, parentheses, hyphens
        phone = re.sub(r'\s+', ' ', phone).strip()  # Normalize spaces
        
        return phone
    
    def is_postcode(self, line):
        """Check if a line contains a UK postcode and return it if found"""
        # More flexible UK postcode regex that handles common variations
        postcode_pattern = r'\b([A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2})\b'
        match = re.search(postcode_pattern, line.upper())
        if match:
            postcode = match.group(1)
            # Ensure proper spacing in postcode
            postcode = re.sub(r'([A-Z]\d+[A-Z]?)(\d[A-Z]{2})', r'\1 \2', postcode)
            return postcode
        return None
    
    def parse_jobs(self, text):
        """Parse BC04 jobs from text input"""
        self.jobs = []
        
        # Split by "Job Sheet" sections
        job_sections = re.split(r'Job Sheet\s*\n', text)
        
        # Remove empty sections
        job_sections = [section.strip() for section in job_sections if section.strip()]
        
        for section in job_sections:
            if section.strip():
                job = self.parse_single_job(section)
                if job and job.get('REG NUMBER'):  # Only add if we found a registration
                    self.jobs.append(job)
        
        return self.jobs
    
    def parse_single_job(self, job_text):
        job = {}
        # Initialize all fields
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

        # Extract job number and vehicle registration
        job_number_match = re.search(r'Job Number.*?(\d+/\d+)', job_text, re.DOTALL)
        if job_number_match:
            job['YOUR REF NO'] = job_number_match.group(1)
        # Robust registration extraction (UK reg: 2 letters, 2 digits, 3 letters)
        reg_match = re.search(r'([A-Z]{2}\d{2}[A-Z]{3})', job_text)
        if reg_match:
            job['REG NUMBER'] = reg_match.group(1)
        # VIN extraction (12+ digits, after reg)
        vin_match = re.search(rf'{job["REG NUMBER"]}\s+(\d{{9,}})', job_text) if job['REG NUMBER'] else None
        if vin_match:
            job['VIN'] = vin_match.group(1)
        # --- MAKE and MODEL are always blank for BC04 ---
        job['MAKE'] = ''
        job['MODEL'] = ''
        # Price extraction
        price_matches = re.findall(r'£?\s*(\d+\.\d{2})', job_text)
        if len(price_matches) >= 2:
            job['PRICE'] = price_matches[1]
        elif price_matches:
            job['PRICE'] = price_matches[0]

        # --- Improved Address Extraction for BC04 ---
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
            # Assign collection address (ADDR1-3, ADDR4=town, POSTCODE)
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
            # Assign delivery address (ADDR1-3, ADDR4=town, POSTCODE)
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
        # --- Phone Extraction and Dynamic Date Extraction ---
        # Find the line with two phone numbers before the date lines
        phone_line = ''
        found_dates = False
        for i, line in enumerate(lines):
            # Look for two phone numbers (8+ digits) separated by space or tab
            phones = re.findall(r'\d{8,}', line)
            if len(phones) >= 2:
                # Check if the next line is a date (dd/mm/yyyy)
                if i+1 < len(lines) and re.match(r'\d{2}/\d{2}/\d{4}', lines[i+1]):
                    job['COLLECTION PHONE'] = phones[0]
                    job['DELIVERY CONTACT PHONE'] = phones[1]
                    # Now extract the first two dates after this line
                    date_matches = []
                    for l in lines[i+1:i+5]:
                        date_matches += re.findall(r'\d{2}/\d{2}/\d{4}', l)
                        if len(date_matches) >= 2:
                            break
                    if len(date_matches) >= 2:
                        job['COLLECTION DATE'] = date_matches[0]
                        job['DELIVERY DATE'] = date_matches[1]
                    found_dates = True
                    break
        # If not found, fallback to default dates
        if not found_dates:
            job['COLLECTION DATE'] = self.collection_date
            job['DELIVERY DATE'] = self.delivery_date
        return job

# Always look for credentials.json in the same directory as this script
CREDENTIALS_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credentials.json')
TOKEN_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'token.json')

if __name__ == "__main__":
    try:
        # Import needed for Pillow if it exists (for better visuals)
        try:
            from PIL import ImageTk, Image
            has_pillow = True
        except ImportError:
            has_pillow = False
        
        # Check for the holidays package which might be missing
        try:
            import holidays
        except ImportError:
            # Simple fallback for holidays
            class DummyHolidays:
                def __init__(self, *args, **kwargs):
                    pass
                def __contains__(self, date):
                    return False
            # Monkey patch the holidays module
            import sys
            sys.modules['holidays'] = type('', (), {'UK': DummyHolidays})
            import holidays
        
        root = tk.Tk()
        app = VehicleTransportApp(root)
        root.mainloop()
    except Exception as e:
        # Create a simple error window if the main app fails to load
        import traceback
        error_msg = traceback.format_exc()
        
        error_root = tk.Tk()
        error_root.title("Error Starting Application")
        error_root.geometry("600x400")
        
        error_frame = tk.Frame(error_root, padx=20, pady=20)
        error_frame.pack(fill=tk.BOTH, expand=True)
        
        error_label = tk.Label(
            error_frame,
            text="Error Starting Application",
            font=('Segoe UI', 14, 'bold'),
            fg="#F44336" 
        )
        error_label.pack(pady=10)
        
        error_message = tk.Label(
            error_frame,
            text="The Vehicle Transport Parser couldn't start due to an error.\nPlease report this issue.",
            wraplength=550
        )
        error_message.pack(pady=10)
        
        error_details = tk.Text(
            error_frame,
            wrap=tk.WORD,
            height=15,
            width=70
        )
        error_details.insert(tk.END, f"Error details:\n\n{error_msg}")
        error_details.config(state=tk.DISABLED)
        error_details.pack(fill=tk.BOTH, expand=True, pady=10)
        
        def close_app():
            error_root.destroy()
        
        close_button = tk.Button(
            error_frame,
            text="Close",
            command=close_app,
            bg="#F44336", 
            fg="white",
            padx=20,
            pady=5
        )
        close_button.pack(pady=10)
        
        error_root.mainloop() 