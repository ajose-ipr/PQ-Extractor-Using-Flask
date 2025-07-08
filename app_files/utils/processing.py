import re
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import logging

# Set up logging
logger = logging.getLogger(__name__)

# Constants
EXPECTED_HARMONICS = set(range(2, 51))
FAIL_COLOR = 'color: darkred; font-weight: bold; background-color: #ffebeb'
HIGHLIGHT_COLOR = 'font-weight: bold; background-color: #fff3cd'

# Column definitions for all table types
VOLTAGE_COLUMNS = [
    "Harmonic", "Time Percent Limit[%]", "Reg Max[%]",
    "Measured_V1N", "Measured_V2N", "Measured_V3N", 
    "Result_V1N", "Result_V2N", "Result_V3N"
]
CURRENT_COLUMNS = [
    "Harmonic", "Time Percent Limit[%]", "Reg Max[%]",
    "Measured_I1", "Measured_I2", "Measured_I3", 
    "Result_I1", "Result_I2", "Result_I3"
]

# Updated section boundaries for all tables
SECTION_BOUNDARIES = {
    "HARMONIC VOLTAGE FULL TIME RANGE": [
        "SUMMARY", "TOTAL HARMONIC VOLTAGE FULL TIME RANGE", 
        "TOTAL HARMONIC DISTORTION FULL TIME RANGE", "HARMONIC CURRENT FULL TIME RANGE"
    ],
    "HARMONIC CURRENT FULL TIME RANGE": [
        "TOTAL HARMONIC DISTORTION DAILY", "TDD FULL TIME RANGE",
        "HARMONIC VOLTAGE DAILY", "TRANSIENT"
    ],
    "HARMONIC VOLTAGE DAILY": [
        "TOTAL HARMONIC DISTORTION FULL TIME RANGE", 
        "TOTAL HARMONIC VOLTAGE FULL TIME RANGE", "HARMONIC CURRENT DAILY", "TOTAL HARMONIC DISTORTION DAILY"
    ],
    "HARMONIC CURRENT DAILY": [
        "TDD FULL TIME RANGE", "TDD DAILY", "TRANSIENT", "FLICKER SEVERITY"
    ]
}

# All supported table names
SUPPORTED_TABLES = [
    "Harmonic Voltage Full Time Range",
    "Harmonic Current Full Time Range", 
    "Harmonic Voltage Daily",
    "Harmonic Current Daily"
]

# Enhanced regex patterns for text extraction
TEXT_EXTRACTION_PATTERNS = [
    # Standard pattern with Pass/Fail
    re.compile(
        r'(\d+)\s*,?\s*'  # Harmonic
        r'(\d+)\s*,?\s*'  # Time percent
        r'([\d.]+)\s*,?\s*'  # Reg max
        r'([\d.]+)\s*,?\s*'  # Measured 1
        r'([\d.]+)\s*,?\s*'  # Measured 2
        r'([\d.]+)\s*,?\s*'  # Measured 3
        r'(Pass|Fail)\s*\(([\d.%]+)\)\s*,?\s*'  # Result 1
        r'(Pass|Fail)\s*\(([\d.%]+)\)\s*,?\s*'  # Result 2
        r'(Pass|Fail)\s*\(([\d.%]+)\)',  # Result 3
        re.IGNORECASE
    ),
    # Pattern without explicit Pass/Fail words
    re.compile(
        r'(\d+)\s*,?\s*'
        r'(\d+)\s*,?\s*'
        r'([\d.]+)\s*,?\s*'
        r'([\d.]+)\s*,?\s*'
        r'([\d.]+)\s*,?\s*'
        r'([\d.]+)\s*,?\s*'
        r'\(([\d.%]+)\)\s*,?\s*'
        r'\(([\d.%]+)\)\s*,?\s*'
        r'\(([\d.%]+)\)',
        re.IGNORECASE
    ),
    # Pattern for multiline data
    re.compile(
        r'(\d+)\s*,?\s*(\d+)\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)',
        re.IGNORECASE | re.MULTILINE
    )
]

def safe_float_convert(val):
    try:
        return float(val)
    except Exception:
        return None

# --- THD SUMMARY TABLE EXTRACTION AND GENERATION ---

def extract_thd_daily_data_from_pdf(pdf_file):
    """Extract THD Daily data from PDF"""
    voltage_thd_daily = []
    current_tdd_daily = []
    try:
        with pdfplumber.open(pdf_file if isinstance(pdf_file, str) else pdf_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                # Voltage THD
                if "Total Harmonic Distortion Daily" in text and "3sec THD" in text:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            for row in table:
                                if row and len(row) >= 6:
                                    day = str(row[0]).strip() if row[0] else ""
                                    if re.match(r'\d{2}-\d{2}-\d{4}', day):
                                        voltage_thd_daily.append({
                                            "Day": day,
                                            "V1N": safe_float_convert(row[3]),
                                            "V2N": safe_float_convert(row[4]),
                                            "V3N": safe_float_convert(row[5])
                                        })
                # Current TDD
                if "TDD Daily" in text and "3sec TDD" in text:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            for row in table:
                                if row and len(row) >= 6:
                                    day = str(row[0]).strip() if row[0] else ""
                                    if re.match(r'\d{2}-\d{2}-\d{4}', day):
                                        current_tdd_daily.append({
                                            "Day": day,
                                            "I1": safe_float_convert(row[3]),
                                            "I2": safe_float_convert(row[4]),
                                            "I3": safe_float_convert(row[5])
                                        })
        return voltage_thd_daily, current_tdd_daily
    except Exception as e:
        logger.error(f"Error extracting THD data: {str(e)}")
        return [], []

def generate_thd_summary_tables_from_pdf(pdf_file):
    """Generate THD summary tables from PDF"""
    voltage_thd_daily, current_tdd_daily = extract_thd_daily_data_from_pdf(pdf_file)
    voltage_data = [
        {
            "Day": day_data["Day"],
            "Recommended limit (%)": 7.5,
            "R Phase (%)": day_data["V1N"],
            "Y Phase (%)": day_data["V2N"],
            "B Phase (%)": day_data["V3N"],
            "Remarks": "All values within limits" if all(
                v is not None and v <= 7.5 for v in [day_data["V1N"], day_data["V2N"], day_data["V3N"]]
            ) else "Some values exceed limits"
        }
        for day_data in voltage_thd_daily
    ]
    current_data = [
        {
            "Day": day_data["Day"],
            "Recommended limit (%)": 10.0,
            "R Phase (%)": day_data["I1"],
            "Y Phase (%)": day_data["I2"],
            "B Phase (%)": day_data["I3"],
            "Remarks": "All values within limits" if all(
                i is not None and i <= 10.0 for i in [day_data["I1"], day_data["I2"], day_data["I3"]]
            ) else "Some values exceed limits"
        }
        for day_data in current_tdd_daily
    ]
    return pd.DataFrame(voltage_data), pd.DataFrame(current_data)

# --- END THD SUMMARY TABLE EXTRACTION AND GENERATION ---

def extract_metadata(pdf_file, filename):
    """Extract metadata from PDF file"""
    name = filename if isinstance(filename, str) else filename.name
    component_info = re.findall(r"\((.*?)\)", str(name))
    component_text = component_info[0] if component_info else "Not found"

    report_info = {
        "start_time": "Not found", 
        "end_time": "Not found", 
        "gmt": "Not found", 
        "version": "Not found"
    }

    try:
        with pdfplumber.open(pdf_file if isinstance(pdf_file, str) else pdf_file) as pdf:
            text0 = pdf.pages[0].extract_text() or ""
        
        time_pattern = re.compile(
            r"Start time:\s*(\d{2}-\d{2}-\d{4}\s*\d{2}:\d{2}:\d{2}\s*[AP]M)\s*"
            r"End time:\s*(\d{2}-\d{2}-\d{4}\s*\d{2}:\d{2}:\d{2}\s*[AP]M)\s*"
            r"GMT:\s*([+-]\d{2}:\d{2})\s*"
            r"Report Version:\s*([\d.]+)"
        )
        
        match = time_pattern.search(text0)
        if match:
            report_info = {
                "start_time": match.group(1),
                "end_time": match.group(2),
                "gmt": match.group(3),
                "version": match.group(4)
            }
            
        combined_text = (str(name) + " " + text0).upper()
        block_match = re.search(r"\bBLOCK[-\s]*(\d{1,3})\b", combined_text)
        feeder_match = re.search(r"\b(FEEDER|BAY)[-\s]*(\d{1,3})\b", combined_text)
        company_match = re.search(r"\b(TATA|ADANI|NTPC|RELIANCE|POWERGRID|TORRENT)\b", combined_text)

        block = block_match.group(1) if block_match else "Not found"
        feeder = feeder_match.group(2) if feeder_match else "Not found"
        company = company_match.group(1) if company_match else "Not found"

        return component_text, block, feeder, company, report_info

    except Exception as e:
        logger.error(f"Error extracting metadata: {str(e)}")
        return "Not found", "Error", "Error", "Error", report_info

def extract_table_data_from_text(text, has_results=True):
    """Enhanced text extraction for all table types"""
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'(Pass|Fail)\s*\(\s*([\d.%]+)\s*\)', r'\1(\2)', text)
    
    data = []
    
    # Try each pattern
    for pattern in TEXT_EXTRACTION_PATTERNS:
        for match in pattern.finditer(text):
            try:
                harmonic = int(match.group(1))
                
                # FILTER 1: Skip harmonic 1 (fundamental frequency)
                if harmonic == 1:
                    continue
                
                # FILTER 2: Skip non-harmonic data (years, dates, etc.)
                # Valid harmonics are 2-50, anything outside this range is likely date/year data
                if harmonic < 2 or harmonic > 50:
                    continue
                
                # FILTER 3: Additional validation - if the "harmonic" looks like a year (>1000)
                if harmonic > 1000:
                    continue
                
                groups = match.groups()
                
                if len(groups) >= 12:  # Full pattern with Pass/Fail
                    row = [
                        harmonic, match.group(2), match.group(3), match.group(4),
                        match.group(5), match.group(6), f"{match.group(7)}({match.group(8)})",
                        f"{match.group(9)}({match.group(10)})", f"{match.group(11)}({match.group(12)})"
                    ]
                elif len(groups) >= 9:  # Pattern without Pass/Fail
                    row = [
                        harmonic, match.group(2), match.group(3), match.group(4),
                        match.group(5), match.group(6), f"Pass({match.group(7)})",
                        f"Pass({match.group(8)})", f"Pass({match.group(9)})"
                    ]
                elif len(groups) >= 6 and not has_results:  # Just measurements
                    row = [
                        harmonic, match.group(2), match.group(3), match.group(4),
                        match.group(5), match.group(6), "N/A", "N/A", "N/A"
                    ]
                else:
                    continue
                    
                data.append(row)
            except (ValueError, IndexError):
                continue
    
    return data

def _extract_structured_data(page_tables, tables, active_table):
    """Helper function to extract structured table data"""
    for table in page_tables:
        if len(table) > 1:
            for row in table:
                if row and str(row[0]).strip().isdigit():
                    try:
                        harmonic = int(row[0])
                        
                        # FILTER 1: Skip harmonic 1 (fundamental frequency)
                        if harmonic == 1:
                            continue
                            
                        # FILTER 2: Only accept valid harmonic range (2-50)
                        if harmonic < 2 or harmonic > 50:
                            continue
                            
                        # FILTER 3: Skip year-like numbers (>1000)
                        if harmonic > 1000:
                            continue
                        
                        if len(row) >= 9:
                            clean_row = [str(cell).strip() if cell is not None else "" for cell in row[:9]]
                            tables[active_table].append(clean_row)
                    except (ValueError, IndexError):
                        continue

def _extract_text_data(text, tables, active_table):
    """Helper function to extract text-based data"""
    text_data = extract_table_data_from_text(text)
    if text_data:
        existing_harmonics = {int(row[0]) for row in tables[active_table] if row and str(row[0]).isdigit()}
        for new_row in text_data:
            try:
                harmonic_value = int(new_row[0])
                
                # ADDITIONAL FILTER: Ensure we only add valid harmonics (2-50)
                if 2 <= harmonic_value <= 50 and harmonic_value not in existing_harmonics:
                    tables[active_table].append(new_row)
            except (ValueError, IndexError):
                continue

def _check_boundary_hit(upper_text, active_table):
    """Helper function to check if section boundary is hit"""
    active_upper = active_table.upper()
    boundaries = SECTION_BOUNDARIES.get(active_upper, [])
    
    # For Harmonic Current Daily, exclude "HARMONIC 5:" from boundaries
    if active_upper == "HARMONIC CURRENT DAILY":
        boundaries = [b for b in boundaries if "HARMONIC 5" not in b.upper()]
    
    return any(boundary in upper_text for boundary in boundaries)

def extract_tables_from_pdf(file):
    """Extract all harmonic tables from PDF starting from page 2"""
    tables = {table_name: [] for table_name in SUPPORTED_TABLES}
    
    try:
        with pdfplumber.open(file if isinstance(file, str) else file) as pdf:
            active_table = None
            
            # Skip first page as requested
            for page_num, page in enumerate(pdf.pages):
                if page_num == 0:  # Skip first page
                    continue
                    
                page_text = page.extract_text() or ""
                upper_text = page_text.upper()
                page_tables = page.extract_tables()
                
                # Check for table headers
                for table_name in tables:
                    table_name_upper = table_name.upper()
                    if table_name_upper in upper_text:
                        start_idx = upper_text.find(table_name_upper)
                        end_idx = len(page_text)
                        
                        # Find section boundaries
                        for boundary in SECTION_BOUNDARIES.get(table_name_upper, []):
                            boundary_idx = upper_text.find(boundary, start_idx + len(table_name))
                            if boundary_idx != -1:
                                end_idx = min(end_idx, boundary_idx)
                        
                        section_text = page_text[start_idx:end_idx]
                        active_table = table_name
                        
                        # Extract structured tables
                        _extract_structured_data(page_tables, tables, active_table)
                        
                        # Extract from text as fallback
                        _extract_text_data(section_text, tables, active_table)
                        continue
                
                # Continue extracting for active table
                if active_table and not _check_boundary_hit(upper_text, active_table):
                    _extract_structured_data(page_tables, tables, active_table)
                    _extract_text_data(page_text, tables, active_table)
                else:
                    # Only reset active_table if we actually hit a real boundary, not "HARMONIC 5:"
                    if active_table and _check_boundary_hit(upper_text, active_table):
                        # Special case: Don't stop for "HARMONIC 5:" when processing Harmonic Current Daily
                        if active_table == "Harmonic Current Daily" and "HARMONIC 5:" in upper_text:
                            # Continue processing this page for the current table
                            _extract_structured_data(page_tables, tables, active_table)
                            _extract_text_data(page_text, tables, active_table)
                        else:
                            active_table = None

    except Exception as e:
        logger.error(f"Error processing PDF: {str(e)}")
    
    return tables

def process_table_data(table_data, table_name=None):
    """Process and validate table data"""
    columns = CURRENT_COLUMNS if table_name and "Current" in table_name else VOLTAGE_COLUMNS
    
    if not table_data:
        return pd.DataFrame(columns=columns)

    try:
        df = pd.DataFrame(table_data, columns=columns)
        df['Harmonic'] = pd.to_numeric(df['Harmonic'], errors='coerce')
        df = df.dropna(subset=['Harmonic'])
        
        # CRITICAL FILTER: Remove fundamental frequency and invalid harmonics
        df = df[df['Harmonic'] != 1]
        
        # ADDITIONAL FILTER: Only keep valid harmonic range (2-50)
        # This removes any year data (2024, 2025, etc.) that might have been captured
        df = df[(df['Harmonic'] >= 2) & (df['Harmonic'] <= 50)]
        
        df = df.drop_duplicates(subset=['Harmonic', 'Time Percent Limit[%]'])
        
        found_harmonics = set(df['Harmonic'].astype(int))
        missing = sorted(EXPECTED_HARMONICS - found_harmonics)
        
        # Log missing harmonics for debugging
        if missing:
            logger.info(f"Missing harmonics in {table_name}: {missing[:10]}{'...' if len(missing) > 10 else ''}")
        
        numeric_cols = ["Harmonic", "Time Percent Limit[%]", "Reg Max[%]", columns[3], columns[4], columns[5]]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce")
        
        return df.dropna()
    
    except Exception as e:
        logger.error(f"Error processing table data: {str(e)}")
        return pd.DataFrame(columns=columns)

def split_table(df):
    """Split table by time limits and odd/even harmonics"""
    if df.empty:
        return {"95": (pd.DataFrame(), pd.DataFrame()), "99": (pd.DataFrame(), pd.DataFrame())}
    
    def split_odd_even(df_subset):
        if df_subset.empty:
            return pd.DataFrame(), pd.DataFrame()
        df_subset = df_subset.sort_values("Harmonic")
        odd = df_subset[df_subset["Harmonic"] % 2 == 1].reset_index(drop=True)
        even = df_subset[df_subset["Harmonic"] % 2 == 0].reset_index(drop=True)
        return odd, even
    
    df_95 = df[df["Time Percent Limit[%]"] == 95.0].copy()
    df_99 = df[df["Time Percent Limit[%]"] == 99.0].copy()
    
    return {"95": split_odd_even(df_95), "99": split_odd_even(df_99)}

def analyze_failures(df):
    """Identify and summarize all harmonic violations"""
    if df.empty:
        return pd.DataFrame()
    
    violations = []
    measured_cols = [col for col in df.columns if col.startswith('Measured_')]
    
    for _, row in df.iterrows():
        try:
            threshold = float(row['Reg Max[%]'])
            harmonic = int(row['Harmonic'])
            time_limit = row['Time Percent Limit[%]']
            
            for col in measured_cols:
                phase = col.split('_')[-1]  # Gets V1N/V2N/V3N or I1/I2/I3
                value = float(row[col])
                
                if value > threshold:
                    violations.append({
                        'Harmonic': harmonic,
                        'Phase': phase,
                        'Time Limit (%)': time_limit,
                        'Allowed (%)': threshold,
                        'Measured (%)': value,
                        'Exceedance (%)': round(value - threshold, 2)
                    })
        except (ValueError, KeyError):
            continue
    
    return pd.DataFrame(violations)

def highlight_fails_in_excel(df, ws, start_row=1):
    """Apply conditional formatting to highlight failed harmonics in Excel"""
    if df.empty:
        return
    
    fail_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    fail_font = Font(color='9C0006', bold=True)
    harmonic_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    reg_max_col = None
    measured_cols = []
    result_cols = []
    
    for idx, col in enumerate(df.columns, 1):
        if "Reg Max" in col:
            reg_max_col = idx
        elif col.startswith(('Measured_')):
            measured_cols.append(idx)
        elif col.startswith('Result_'):
            result_cols.append(idx)
    
    if not reg_max_col or not measured_cols:
        return
    
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
        if r_idx == start_row:
            continue
            
        try:
            threshold = float(ws.cell(row=r_idx, column=reg_max_col).value)
        except (ValueError, TypeError):
            threshold = 0.0
        
        for i, m_col in enumerate(measured_cols):
            try:
                cell = ws.cell(row=r_idx, column=m_col)
                value = float(cell.value) if cell.value else 0.0
                
                result_value = None
                if i < len(result_cols):
                    result_cell = ws.cell(row=r_idx, column=result_cols[i])
                    result_value = str(result_cell.value).lower() if result_cell.value else ""
                
                if value > threshold or (result_value and "fail" in result_value):
                    cell.fill = fail_fill
                    cell.font = fail_font
                    
                    harmonic_cell = ws.cell(row=r_idx, column=1)
                    harmonic_cell.fill = harmonic_fill
            except:
                continue

def parse_filename_for_sheet_name(filename):
    """Parse filename to extract day number and time period for concise sheet naming"""
    filename_upper = filename.upper()
    
    if "7" in filename_upper and "DAY" in filename_upper:
        return "7Days"
    
    day_pattern = re.search(r'DAY\s*(\d+)\s*(DAY|NIGHT)', filename_upper)
    if day_pattern:
        day_num = day_pattern.group(1)
        period = "D" if "DAY" in day_pattern.group(2) else "N"
        return f"{day_num}{period}"
    
    day_only_pattern = re.search(r'DAY\s*(\d+)', filename_upper)
    if day_only_pattern:
        return f"{day_only_pattern.group(1)}D"
    
    clean_name = re.sub(r'[^\w]', '', filename.replace('.pdf', ''))
    return clean_name[:4]

def get_table_abbreviation(table_name):
    """Convert full table names to abbreviations for concise sheet naming"""
    table_upper = table_name.upper()
    
    if "CURRENT" in table_upper:
        prefix = "I"
    elif "VOLTAGE" in table_upper:
        prefix = "V"
    else:
        prefix = "X"
    
    if "FULL TIME RANGE" in table_upper:
        suffix = "F"
    elif "DAILY" in table_upper:
        suffix = "D"
    else:
        suffix = "X"
    
    return f"{prefix}{suffix}"

def create_enhanced_excel_download(tables_data, thd_tdd_tables, filename):
    """Create Excel file with both individual harmonic tables and THD/TDD summary tables"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # Add THD/TDD Summary Tables first
        for table_name, df in thd_tdd_tables.items():
            if not df.empty:
                sheet_name = f"THD_TDD_{table_name[:20]}"[:31]  # Ensure valid sheet name
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Apply formatting
                workbook = writer.book
                worksheet = workbook[sheet_name]
                
                # Header formatting
                header_fill = PatternFill(start_color='E3F2FD', end_color='E3F2FD', fill_type='solid')
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = Font(bold=True)
                
                # Violation highlighting for remarks column
                for row in worksheet.iter_rows(min_row=2):
                    remarks_cell = row[-1]  # Last column should be remarks
                    if remarks_cell.value and "Exceeding" in str(remarks_cell.value):
                        for cell in row:
                            cell.fill = PatternFill(start_color='FFEBEE', end_color='FFEBEE', fill_type='solid')
                            cell.font = Font(color='D32F2F', bold=True)
        
        # Add individual harmonic tables
        for table_name, table_data in tables_data.items():
            if table_data:
                df = process_table_data(table_data, table_name)
                if not df.empty:
                    split_dfs = split_table(df)
                    
                    table_prefix = "I" if "Current" in table_name else "V"
                    table_suffix = "D" if "Daily" in table_name else "F"
                    
                    for limit in ["95", "99"]:
                        odd_df, even_df = split_dfs[limit]
                        
                        for df_data, suffix in [(odd_df, 'O'), (even_df, 'E')]:
                            if not df_data.empty:
                                sheet_name = f"H_{table_prefix}{table_suffix}_{limit}_{suffix}"[:31]
                                df_data.to_excel(writer, sheet_name=sheet_name, index=False)
                                
                                workbook = writer.book
                                worksheet = workbook[sheet_name]
                                highlight_fails_in_excel(df_data, worksheet, start_row=2)
    
    output.seek(0)
    return output.getvalue()

def create_bulk_excel_download(all_files_data):
    """Create Excel file with all PDFs using concise sheet naming and highlighting"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        sheet_file_map = {}
        
        for file_name, tables_data in all_files_data.items():
            file_prefix = parse_filename_for_sheet_name(file_name)
            
            for table_name, table_data in tables_data.items():
                if table_data:
                    df = process_table_data(table_data, table_name)
                    if not df.empty:
                        table_abbrev = get_table_abbreviation(table_name)
                        sheet_name = f"{file_prefix}_H_{table_abbrev}"
                        
                        if len(sheet_name) > 31:
                            sheet_name = sheet_name[:31]
                        
                        original_sheet_name = sheet_name
                        counter = 1
                        while sheet_name in sheet_file_map:
                            sheet_name = f"{original_sheet_name}_{counter}"
                            if len(sheet_name) > 31:
                                truncated = original_sheet_name[:31-len(f"_{counter}")]
                                sheet_name = f"{truncated}_{counter}"
                            counter += 1
                        
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheet_file_map[sheet_name] = file_name
                        
                        workbook = writer.book
                        worksheet = workbook[sheet_name]
                        highlight_fails_in_excel(df, worksheet, start_row=2)
    
    output.seek(0)
    wb = load_workbook(output)
    
    for sheet_name, file_name in sheet_file_map.items():
        ws = wb[sheet_name]
        ws.insert_rows(1)
        ws.cell(row=1, column=1, value=f"File: {file_name}")
    
    output2 = BytesIO()
    wb.save(output2)
    output2.seek(0)
    return output2.getvalue()

# Add these functions to your processing.py file

def extract_thd_tdd_summary_tables_from_pdf(pdf_file):
    """
    Extract all THD/TDD summary tables from PDF
    Returns: Dict with 4 table types
    """
    tables_data = {
        'voltage_thd_full_95': [],      # 10min THD 95th percentile
        'voltage_thd_daily_99': [],     # 3sec THD 99th percentile  
        'current_tdd_full_95': [],      # 10min TDD 95th percentile
        'current_tdd_full_99': [],      # 10min TDD 99th percentile
        'current_tdd_daily_99': []      # 3sec TDD 99th percentile
    }
    
    try:
        with pdfplumber.open(pdf_file if isinstance(pdf_file, str) else pdf_file) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text() or ""
                text_upper = text.upper()
                
                # Extract Voltage THD Full Time Range (95th percentile)
                if "TOTAL HARMONIC DISTORTION FULL TIME RANGE" in text_upper and "10MIN THD" in text_upper:
                    tables_data['voltage_thd_full_95'].extend(
                        _extract_thd_table_data(page, text, "voltage", "95")
                    )
                
                # Extract Voltage THD Daily (99th percentile)  
                if "TOTAL HARMONIC DISTORTION DAILY" in text_upper and "3SEC THD" in text_upper:
                    tables_data['voltage_thd_daily_99'].extend(
                        _extract_thd_table_data(page, text, "voltage", "99")
                    )
                
                # Extract Current TDD Full Time Range (both 95th and 99th)
                if "TDD FULL TIME RANGE" in text_upper and "10MIN TDD" in text_upper:
                    # Check for both percentiles in the table
                    if "95%" in text and "99%" in text:
                        tables_data['current_tdd_full_95'].extend(
                            _extract_thd_table_data(page, text, "current", "95")
                        )
                        tables_data['current_tdd_full_99'].extend(
                            _extract_thd_table_data(page, text, "current", "99")
                        )
                
                # Extract Current TDD Daily (99th percentile)
                if "TDD DAILY" in text_upper and "3SEC TDD" in text_upper:
                    tables_data['current_tdd_daily_99'].extend(
                        _extract_thd_table_data(page, text, "current", "99")
                    )
                    
        return tables_data
        
    except Exception as e:
        logger.error(f"Error extracting THD/TDD summary tables: {str(e)}")
        return tables_data

def _extract_thd_table_data(page, text, table_type, percentile):
    """
    Helper function to extract THD/TDD data from a page
    """
    extracted_data = []
    
    try:
        # Try structured table extraction first
        tables = page.extract_tables()
        for table in tables:
            if table and len(table) > 1:
                # Look for header row to identify the correct table
                header_row = None
                for i, row in enumerate(table):
                    if row and any(cell and ("day" in str(cell).lower() or "date" in str(cell).lower()) for cell in row):
                        header_row = i
                        break
                
                if header_row is not None:
                    # Process data rows
                    for row in table[header_row + 1:]:
                        if row and len(row) >= 6:
                            day_cell = str(row[0]).strip() if row[0] else ""
                            
                            # Check if this is a valid date row
                            if re.match(r'\d{1,2}[-/]\d{1,2}[-/]\d{4}', day_cell):
                                data_row = {
                                    "Day": day_cell,
                                    "Percentile": percentile,
                                    "Table_Type": table_type,
                                    "R_Phase": safe_float_convert(row[3]) if len(row) > 3 else None,
                                    "Y_Phase": safe_float_convert(row[4]) if len(row) > 4 else None,
                                    "B_Phase": safe_float_convert(row[5]) if len(row) > 5 else None,
                                    "Limit": 7.5 if table_type == "voltage" else 10.0
                                }
                                extracted_data.append(data_row)
        
        # Fallback: text pattern extraction if table extraction fails
        if not extracted_data:
            extracted_data = _extract_thd_from_text_patterns(text, table_type, percentile)
            
    except Exception as e:
        logger.error(f"Error in _extract_thd_table_data: {str(e)}")
    
    return extracted_data

def _extract_thd_from_text_patterns(text, table_type, percentile):
    """
    Extract THD/TDD data using regex patterns as fallback
    """
    extracted_data = []
    
    # Pattern to match date and values
    date_pattern = r'(\d{1,2}[-/]\d{1,2}[-/]\d{4})\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)\s*,?\s*([\d.]+)'
    
    matches = re.finditer(date_pattern, text)
    for match in matches:
        try:
            data_row = {
                "Day": match.group(1),
                "Percentile": percentile,
                "Table_Type": table_type,
                "R_Phase": safe_float_convert(match.group(3)),
                "Y_Phase": safe_float_convert(match.group(4)),
                "B_Phase": safe_float_convert(match.group(5)),
                "Limit": 7.5 if table_type == "voltage" else 10.0
            }
            extracted_data.append(data_row)
        except (IndexError, ValueError):
            continue
    
    return extracted_data

def process_thd_tdd_summary_data(tables_data):
    """
    Process extracted THD/TDD data into formatted DataFrames
    """
    processed_tables = {}
    
    for table_key, raw_data in tables_data.items():
        if raw_data:
            df_data = []
            for row in raw_data:
                processed_row = {
                    "Day": row["Day"],
                    "Recommended limit (%)": row["Limit"],
                    "R Phase (%)": row["R_Phase"],
                    "Y Phase (%)": row["Y_Phase"], 
                    "B Phase (%)": row["B_Phase"],
                    "Remarks": _generate_thd_remarks(row)
                }
                df_data.append(processed_row)
            
            processed_tables[table_key] = pd.DataFrame(df_data)
        else:
            processed_tables[table_key] = pd.DataFrame()
    
    return processed_tables

def _generate_thd_remarks(row):
    """
    Generate remarks for THD/TDD rows based on limit compliance
    """
    limit = row["Limit"]
    phases = [row["R_Phase"], row["Y_Phase"], row["B_Phase"]]
    
    violations = []
    for i, (phase_val, phase_name) in enumerate(zip(phases, ["R", "Y", "B"])):
        if phase_val is not None and phase_val > limit:
            violations.append(f"{phase_name}({phase_val:.2f}%)")
    
    if violations:
        return f"Exceeding limit: {', '.join(violations)}"
    else:
        return "All values within limits"



