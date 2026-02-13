import pandas as pd
import re
from datetime import datetime
from dateutil import parser
def get_style_numbers_from_plan(file_path):
    """
    Reads the plan, and returns a list of all style numbers in that plan. Also checks that the 'Style #' in the first column of each sheet
    matches the sheet name, and outputs an error message if not.

    Args:
        file_path: The full path to the Excel file.

    Returns:
        A list of sheet names that passed the validation.
    """
    
    # --- Step 1: Try to open the Excel file ---
    try:
        # The 'ExcelFile' function helps us see all the sheets in the file.
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
    except FileNotFoundError:
        # This error happens if the file path is wrong or the file doesn't exist.
        print(f"Error: The file '{file_path}' was not found.")
        return [] # Return an empty list because we can't continue.
    except Exception as e:
        # This handles any other problems with opening the file.
        print(f"An error occurred while trying to read the file: {e}")
        return []

    # This list will hold the names of the sheets that are correct.
    valid_sheets = []

    # --- Step 2: Loop through each sheet to check it ---
    for sheet_name in sheet_names:
        try:
            # Read just the first column ('A') of the current sheet.
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols="A")
            
            blank_cells = []
            # --- Step 3: Loop through each row in the first column ---
            # We use a simple loop to go through each value in the column.
            # The 'values' attribute gives us all items from that column.
            for index, cell_value in enumerate(df['Style #'].values):

                
                # The Excel row number is the current index + 2.
                # (+1 because lists start at 0, and +1 to skip the header row).
                excel_row_number = index + 2

                # Check if the cell is blank
                if pd.isna(cell_value):
                    blank_cells.append(excel_row_number)
                    continue # Skip to the next row
                
                # Compare the value in the cell to the name of the sheet.
                # We use str() to make sure we are comparing text.
                if str(cell_value) != sheet_name:
                    print(f"In sheet {sheet_name}, cell A{excel_row_number} has '{cell_value}', but it should be '{sheet_name}'. We are ignoring it for now, but please fix urgently.")

            # BLANK CELL HANDLING:
            # Remove trailing blank cells from the blank_cells list.
            # We only want to report blank cells that are in the middle of the data,
            # not the ones at the very end of the sheet.
            if blank_cells:
                # Find the last non-blank row number in the entire column.
                # We loop backwards through the dataframe values to find the last non-blank cell.
                last_non_blank_index = -1
                for index in range(len(df['Style #'].values) - 1, -1, -1):
                    if not pd.isna(df['Style #'].values[index]):
                        last_non_blank_index = index
                        break
                
                # Convert the last non-blank index to Excel row number.
                last_non_blank_row = last_non_blank_index + 2
                
                # Filter out blank cells that come after the last non-blank row.
                # We only keep blank cells that are before or equal to the last non-blank row.
                blank_cells = [row for row in blank_cells if row <= last_non_blank_row]
                
                # Now print the error message only for the remaining blank cells (middle blanks).
                if blank_cells:
                    cells_str = ", ".join(f"A{row}" for row in blank_cells)
                    print(f"For style {sheet_name}, the following cells are empty: {cells_str}. We are ignoring it for now, but please fix.")
            # END OF BLANK CELL HANDLING

            
            # If, after checking all rows, the sheet is still valid, we add it to our list.
            valid_sheets.append(sheet_name)

        except Exception as e:
            print(f"  -> Error: Could not process sheet '{sheet_name}'. Reason: {e}")

    return valid_sheets
def get_row_wise_data_from_plan(file_path, style_number):
    """
    Extracts all row data for a given style number from an Excel file.
    Returns a list of dictionaries with Style No, PO, Colour, Date, and planned quantities.
    
    Args:
        file_path: The full path to the Excel file.
        style_number: The style number (sheet name) to extract data from.
    
    Returns:
        A list of dictionaries, where each dictionary represents one row with keys:
        'Style No', 'PO', 'Colour', 'Date', 'Planned Cutting', 'Planned Sewing', 
        'Planned Washing', 'Planned Finishing', 'Planned Packing', 'Source Sheet', 'Source Row'
    """
    
    # --- Step 1: Try to open the Excel file and locate the sheet ---
    try:
        xls = pd.ExcelFile(file_path)
        
        # Check if the style number exists as a sheet in the file.
        if style_number not in xls.sheet_names:
            print(f"Error: Style number '{style_number}' not found in the file.")
            print(f"Available sheets: {xls.sheet_names}")
            return []
            
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return []
    except Exception as e:
        print(f"An error occurred while trying to read the file: {e}")
        return []
    
    # --- Step 2: Read the specific sheet with columns A through M ---
    try:
        # Read columns A through M (includes all the data we need).
        # Column positions: A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12
        df = pd.read_excel(xls, sheet_name=style_number, usecols="A:M")
        
        # Directly assign the quantity columns by position (no searching needed!)
        # E=4, G=6, I=8, K=10, M=12 (0-indexed)
        quantity_columns = {
            'Planned Cutting': df.columns[4],    # Column E
            'Planned Sewing': df.columns[6],     # Column G
            'Planned Washing': df.columns[8],    # Column I
            'Planned Finishing': df.columns[10], # Column K
            'Planned Packing': df.columns[12]    # Column M
        }
        
        # Check if the required columns exist for Style, PO, Colour, Date
        required_columns = ['Style #', 'PO#', 'Colour', 'Date']
        for col in required_columns:
            if col not in df.columns:
                print(f"Error: Required column '{col}' not found in sheet '{style_number}'.")
                return []
                
    except Exception as e:
        print(f"Error: Could not read sheet '{style_number}'. Reason: {e}")
        return []
    
    # --- Step 3: Find the last non-blank row to identify trailing blanks ---
    # We need to find the last row that has at least one non-blank value in the main columns.
    last_non_blank_index = -1
    for index in range(len(df) - 1, -1, -1):
        row = df.iloc[index]
        # Check if at least one of the four main columns has a value.
        if not (pd.isna(row['Style #']) and pd.isna(row['PO#']) and 
                pd.isna(row['Colour']) and pd.isna(row['Date'])):
            last_non_blank_index = index
            break
    
    # --- Step 4: Initialize tracking variables ---
    result_list = []  # This will hold all the row dictionaries.
    blank_cells = []  # This will track cells that are blank and need attention.
    skipped_rows = []  # This will track rows that were completely skipped due to errors.
    decimal_quantities = []  # This will track quantities that had to be rounded.
    
    # This will store successfully parsed dates with their row indices for year correction.
    parsed_dates_with_indices = []
    
    # --- Step 5: Loop through each row and extract data ---
    for index, row in df.iterrows():
        # The Excel row number (accounting for header row).
        excel_row_number = index + 2
        
        # Extract values from each column.
        style_value = row['Style #']
        po_value = row['PO#']
        colour_value = row['Colour']
        date_value = row['Date']
        
        # --- Step 5.1: Check if the entire row is blank ---
        # If all four main values are blank, we need to check if this is a trailing blank.
        if pd.isna(style_value) and pd.isna(po_value) and pd.isna(colour_value) and pd.isna(date_value):
            # Check if this row comes after the last non-blank row.
            # If yes, it's a trailing blank and we skip it without reporting.
            # If no, it's a blank row in the middle and we should report it.
            if index > last_non_blank_index:
                # This is a trailing blank row at the end. Skip without reporting.
                continue
            else:
                # This is a blank row in the middle of the data. Report it.
                blank_cells.append(f"Row {excel_row_number} (entire row is blank)")
                # We still skip processing this row, but we've reported it.
                continue
        
        # --- Step 5.2: Process Style # ---
        # Convert to string, handling blank values.
        if pd.isna(style_value):
            style_str = ""
            blank_cells.append(f"A{excel_row_number} (Style #)")
        else:
            style_str = str(style_value).strip()
        
        # --- Step 5.3: Process PO# ---
        # Convert to string, handling blank values.
        if pd.isna(po_value):
            po_str = ""
            blank_cells.append(f"B{excel_row_number} (PO#)")
        else:
            # PO numbers might be integers or strings, so we convert carefully.
            po_str = str(int(po_value)) if isinstance(po_value, float) else str(po_value).strip()
        
        # --- Step 5.4: Process Colour ---
        # Normalize the colour: lowercase, remove extra spaces.
        if pd.isna(colour_value):
            colour_normalized = ""
            blank_cells.append(f"C{excel_row_number} (Colour)")
        else:
            # Convert to string and normalize.
            colour_str = str(colour_value)
            # Remove leading and trailing spaces.
            colour_str = colour_str.strip()
            # Replace multiple spaces in the middle with a single space.
            colour_str = re.sub(r'\s+', ' ', colour_str)
            # Convert to lowercase for consistency.
            colour_normalized = colour_str.lower()
        
        # --- Step 5.5: Process Date ---
        # This is the most complex part due to various date formats.
        # Date is ESSENTIAL - if it cannot be parsed, we skip the entire row.
        date_parsed = None
        
        if pd.isna(date_value):
            # Date is blank - this is an error. Skip this row.
            blank_cells.append(f"D{excel_row_number} (Date)")
            print(f"ERROR: Date is blank in cell D{excel_row_number}. Skipping this entire row. Please fill in the date.")
            skipped_rows.append(excel_row_number)
            continue  # Skip to the next row.
        else:
            # Try to parse the date.
            try:
                # Check if it's already a datetime object (pandas sometimes does this).
                if isinstance(date_value, datetime):
                    date_parsed = date_value
                else:
                    # Use dateutil.parser which is very flexible with date formats.
                    # It can handle MM/DD/YYYY, DD/MM/YYYY, and many other formats.
                    date_parsed = parser.parse(str(date_value), dayfirst=False)
                
                # ========== YEAR ERROR DETECTION AND CORRECTION ==========
                # Check if the parsed year looks suspicious (e.g., year < 2000 or year > 2100).
                # This catches cases like "9/16/202" being parsed as year 0002 or 0202.
                if date_parsed.year < 2020 or date_parsed.year > 2050:
                    print(f"WARNING: Suspicious year detected in cell D{excel_row_number}. Original value: '{date_value}', Parsed as: {date_parsed.strftime('%d/%b/%Y')}.")
                    
                    # Try to correct the year by looking at adjacent rows.
                    # We'll check the row above and below to see if they have the same year.
                    year_above = None
                    year_below = None
                    
                    # Look at the row above (if it exists and has been processed).
                    if len(parsed_dates_with_indices) > 0:
                        # Get the most recent successfully parsed date.
                        year_above = parsed_dates_with_indices[-1]['date'].year
                    
                    # Look at the row below (if it exists).
                    # We need to peek ahead in the dataframe.
                    if index + 1 < len(df):
                        next_date_value = df.iloc[index + 1]['Date']
                        if not pd.isna(next_date_value):
                            try:
                                if isinstance(next_date_value, datetime):
                                    next_date_parsed = next_date_value
                                else:
                                    next_date_parsed = parser.parse(str(next_date_value), dayfirst=False)
                                # Only use this year if it looks reasonable.
                                if 2000 <= next_date_parsed.year <= 2100:
                                    year_below = next_date_parsed.year
                            except:
                                pass  # If we can't parse the next row, just skip it.
                    
                    # Now decide what to do based on the years we found.
                    if year_above is not None and year_below is not None and year_above == year_below:
                        # Both adjacent rows have the same year. Use that year.
                        corrected_date = date_parsed.replace(year=year_above)
                        print(f"   -> Auto-corrected to: {corrected_date.strftime('%d/%b/%Y')} (using year from adjacent rows). Please verify.")
                        date_parsed = corrected_date
                    elif year_above is not None and year_below is None:
                        # Only the row above is available. Use its year.
                        corrected_date = date_parsed.replace(year=year_above)
                        print(f"   -> Auto-corrected to: {corrected_date.strftime('%d/%b/%Y')} (using year from row above). Please verify.")
                        date_parsed = corrected_date
                    elif year_above is None and year_below is not None:
                        # Only the row below is available. Use its year.
                        corrected_date = date_parsed.replace(year=year_below)
                        print(f"   -> Auto-corrected to: {corrected_date.strftime('%d/%b/%Y')} (using year from row below). Please verify.")
                        date_parsed = corrected_date
                    else:
                        # We can't determine the correct year. Skip this row.
                        print(f"   -> ERROR: Cannot determine correct year. Skipping this row.")
                        skipped_rows.append(excel_row_number)
                        continue
                # ========== END OF YEAR ERROR DETECTION AND CORRECTION ==========
                
                # ========== CONSECUTIVE DATE YEAR CHANGE VALIDATION ==========
                # Check if the year changed compared to the previous date.
                if len(parsed_dates_with_indices) > 0:
                    previous_date = parsed_dates_with_indices[-1]['date']
                    previous_row_number = parsed_dates_with_indices[-1]['index'] + 2
                    
                    if date_parsed.year > previous_date.year:
                        # Year increased. Check if it's a legitimate New Year transition.
                        # If the new date is in early January (first 2 weeks), it's probably okay.
                        if date_parsed.month == 1 and date_parsed.day <= 14:
                            # This looks like a natural New Year transition. No warning needed.
                            pass
                        else:
                            # Year increased but not in early January. This is suspicious.
                            print(f"⚠️  CAUTION: For style {style_number}, year increased from {previous_date.year} (row {previous_row_number}) to {date_parsed.year} (row {excel_row_number}), but the date is not in early January. Please verify this is correct.")
                    
                    elif date_parsed.year < previous_date.year:
                        # Year went backwards. This is almost always an error.
                        print(f"⚠️  WARNING: For style {style_number}, year decreased from {previous_date.year} (row {previous_row_number}) to {date_parsed.year} (row {excel_row_number}). Dates should generally be in chronological order. Please verify this is correct.")
                # ========== END OF CONSECUTIVE DATE YEAR CHANGE VALIDATION ==========
                
            except Exception as e:
                # If we can't parse the date at all, we need to skip this row.
                # Date is essential, so we cannot proceed without it.
                print(f"ERROR: Unable to parse date in cell D{excel_row_number}. Value: '{date_value}'. Skipping this entire row. Please fix the date format.")
                skipped_rows.append(excel_row_number)
                continue  # Skip to the next row.
        
        # --- Step 5.6: Process Planned Quantities ---
        # Extract quantities directly from columns E, G, I, K, M (no searching!)
        # These should be integers. Blank cells are treated as 0.
        # Decimal values are rounded and reported.
        
        quantities = {}
        
        # Column letter mapping for error reporting
        column_letters = {
            'Planned Cutting': 'E',
            'Planned Sewing': 'G',
            'Planned Washing': 'I',
            'Planned Finishing': 'K',
            'Planned Packing': 'M'
        }
        
        for quantity_name, column_name in quantity_columns.items():
            quantity_value = row[column_name]
            
            # Check if the cell is blank.
            if pd.isna(quantity_value):
                # Blank cell - store as 0.
                quantities[quantity_name] = 0
            else:
                # Try to convert to a number.
                try:
                    # Convert to float first.
                    quantity_float = float(quantity_value)
                    
                    # Check if it's a whole number (like 100.0).
                    if quantity_float == int(quantity_float):
                        # It's a whole number. Store as int.
                        quantities[quantity_name] = int(quantity_float)
                    else:
                        # It has a decimal part (like 100.5). Round it and report.
                        quantity_rounded = round(quantity_float)
                        quantities[quantity_name] = quantity_rounded
                        
                        # Use the column letter for reporting.
                        col_letter = column_letters[quantity_name]
                        decimal_quantities.append(f"{col_letter}{excel_row_number} ({quantity_name}): {quantity_float} rounded to {quantity_rounded}")
                        
                except Exception as e:
                    # If we can't convert to a number, treat as 0 and report.
                    quantities[quantity_name] = 0
                    col_letter = column_letters[quantity_name]
                    print(f"WARNING: Could not parse quantity in cell {col_letter}{excel_row_number} ({quantity_name}). Value: '{quantity_value}'. Using 0.")
        
        # --- Step 5.7: Create the dictionary for this row ---
        # We only reach this point if the date was successfully parsed.
        row_dict = {
            'Style No': style_str if style_str else style_number,  # Use the sheet name if blank.
            'PO': po_str,
            'Colour': colour_normalized,
            'Date': date_parsed.strftime('%d/%b/%y'),  # Date is guaranteed to exist here.
            'Planned Cutting': quantities.get('Planned Cutting', 0),
            'Planned Sewing': quantities.get('Planned Sewing', 0),
            'Planned Washing': quantities.get('Planned Washing', 0),
            'Planned Finishing': quantities.get('Planned Finishing', 0),
            'Planned Packing': quantities.get('Planned Packing', 0),
            'Source Sheet': style_number,          # Metadata: which sheet this row came from
            'Source Row': excel_row_number         # Metadata: which row in the Excel file
        }
        
        result_list.append(row_dict)
        
        # Store this successfully parsed date for future year correction reference.
        parsed_dates_with_indices.append({'index': index, 'date': date_parsed})
    
    # --- Step 6: Print all warnings and errors ---
    if blank_cells:
        print(f"\n⚠️  WARNING: For style {style_number}, the following cells are blank and should be filled:")
        for cell in blank_cells:
            print(f"   - {cell}")
    
    if decimal_quantities:
        print(f"\n⚠️  WARNING: For style {style_number}, the following quantities had decimal values and were rounded:")
        for item in decimal_quantities:
            print(f"   - {item}")
    
    if skipped_rows:
        print(f"\n❌ ERROR: For style {style_number}, the following rows were skipped due to missing or unparseable dates:")
        for row_num in skipped_rows:
            print(f"   - Row {row_num}")
        print("Please fix the dates in these rows and run the function again.")
    
    return result_list

def get_row_wise_data_from_daily_prod(file_path):
    """
    Extracts all row data from a daily production Excel file.
    The file has multiple sheets, each named with a date.
    Returns a list of dictionaries with Style No, PO, Colour, Date, and actual quantities.
    
    Args:
        file_path: The full path to the daily production Excel file.
    
    Returns:
        A list of dictionaries, where each dictionary represents one row with keys:
        'Style No', 'PO', 'Colour', 'Date', 'Actual Cutting', 'Actual Sewing',
        'Actual Finishing', 'Actual Washing', 'Actual Packing', 'Source Sheet', 'Source Row'
    """
    
    # --- Step 1: Try to open the Excel file ---
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
    except FileNotFoundError:
        print(f"Error: The file '{file_path}' was not found.")
        return []
    except Exception as e:
        print(f"An error occurred while trying to read the file: {e}")
        return []
    
    # --- Step 2: Initialize tracking variables ---
    result_list = []  # This will hold all the row dictionaries from all sheets.
    blank_cells = []  # This will track cells that are blank and need attention.
    skipped_sheets = []  # This will track sheets that were skipped due to errors.
    quantity_warnings = []  # This will track quantities with issues (decimals, negatives, unparseable values, etc.).
    
    # --- Step 3: Loop through each sheet ---
    for sheet_name in sheet_names:        
        # --- Step 3.1: Parse the date from the sheet name ---
        # The sheet name is the date. We need to parse it.
        try:
            # Use dateutil.parser to handle various date formats.
            # This can handle formats like "24-SEP-25", "25-SEP-26", "DD/MM/YYYY", etc.
            sheet_date_parsed = parser.parse(sheet_name, dayfirst=True)
        except Exception as e:
            # If we cannot parse the sheet name as a date, skip this sheet.
            print(f"   ERROR: Cannot parse sheet name '{sheet_name}' as a date. Skipping this sheet.")
            print(f"   Please fix the sheet name to be a valid date format.")
            skipped_sheets.append(sheet_name)
            continue  # Skip to the next sheet.
        
        # --- Step 3.2: Read the sheet data ---
        try:
            # Read columns A (PO#), E (Style Number), F (Colour), and H, I, J, K, L (quantities).
            # Column H: Cutting Quantity
            # Column I: Sewing Quantity
            # Column J: Finishing Quantity
            # Column K: Washing Quantity
            # Column L: Packing Quantity
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols="A,E,F,G,H,I,J,K,L")  # ← CHANGED: Added G for Order Quantity
            
            # Check if we got the expected number of columns.
            if len(df.columns) < 9:  # ← CHANGED: Was 8, now 9 because we added Order Quantity
                print(f"   ERROR: Did not find the expected columns in sheet '{sheet_name}'. Skipping this sheet.")
                skipped_sheets.append(sheet_name)
                continue
            
            # Rename columns for consistency.
            # The order should be: A, E, F, G, H, I, J, K, L
            df.columns = ['PO#', 'Style Number', 'Colour', 'Order Quantity',  # ← CHANGED: Added 'Order Quantity'
                         'Cutting Quantity', 'Sewing Quantity', 
                         'Finishing Quantity', 'Washing Quantity', 'Packing Quantity']
            
        except Exception as e:
            print(f"   ERROR: Could not read sheet '{sheet_name}'. Reason: {e}")
            skipped_sheets.append(sheet_name)
            continue
        
        # --- Step 3.3: Find the last non-blank row to identify trailing blanks ---
        # We need to find the last row that has at least one non-blank value in the main columns.
        last_non_blank_index = -1
        for index in range(len(df) - 1, -1, -1):
            row = df.iloc[index]
            # Check if at least one of the three main columns has a value.
            if not (pd.isna(row['PO#']) and pd.isna(row['Style Number']) and pd.isna(row['Colour'])):
                last_non_blank_index = index
                break
        
        # --- Step 3.4: Loop through each row in the sheet ---
        for index, row in df.iterrows():
            # The Excel row number (accounting for header row).
            # Note: In the image, the data starts at row 2, so we add 2.
            excel_row_number = index + 2
            
            # Extract values from each column.
            po_value = row['PO#']
            style_value = row['Style Number']
            colour_value = row['Colour']
            
            # --- Step 3.4.1: Check if the entire row is blank ---
            # If all three main values are blank, check if this is a trailing blank.
            if pd.isna(po_value) and pd.isna(style_value) and pd.isna(colour_value):
                # Check if this row comes after the last non-blank row.
                if index > last_non_blank_index:
                    # This is a trailing blank row at the end. Skip without reporting.
                    continue
                else:
                    # This is a blank row in the middle of the data. Report it.
                    blank_cells.append(f"Sheet '{sheet_name}', Row {excel_row_number} (entire row is blank)")
                    continue
            
            # --- Step 3.4.2: Process PO# ---
            # Convert to string, handling blank values.
            if pd.isna(po_value):
                po_str = ""
                blank_cells.append(f"Sheet '{sheet_name}', Cell A{excel_row_number} (PO#)")
            else:
                # PO numbers might be integers or strings, so we convert carefully.
                po_str = str(int(po_value)) if isinstance(po_value, float) else str(po_value).strip()
            
            # --- Step 3.4.3: Process Style Number ---
            # The style number is in format "PID-9KLXL8" or "PID9KLXL8".
            # We need to extract only the part after "PID-" or "PID".
            if pd.isna(style_value):
                style_str = ""
                blank_cells.append(f"Sheet '{sheet_name}', Cell E{excel_row_number} (Style Number)")
            else:
                style_full = str(style_value).strip()
                
                # Check if it contains "PID-" (with dash).
                if "PID-" in style_full:
                    # Extract the part after "PID-".
                    style_str = style_full.split("PID-")[1].strip()
                # Check if it contains "PID" (without dash).
                elif "PID" in style_full:
                    # Extract the part after "PID".
                    style_str = style_full.split("PID")[1].strip()
                else:
                    # If neither "PID-" nor "PID" is found, use the full value.
                    # But print a warning.
                    print(f"   WARNING: Style number in sheet '{sheet_name}', row {excel_row_number} does not contain 'PID-' or 'PID'. Value: '{style_full}'. Using as-is.")
                    style_str = style_full
            
            # --- Step 3.4.4: Process Colour ---
            # Normalize the colour: lowercase, remove extra spaces.
            if pd.isna(colour_value):
                colour_normalized = ""
                blank_cells.append(f"Sheet '{sheet_name}', Cell F{excel_row_number} (Colour)")
            else:
                # Convert to string and normalize.
                colour_str = str(colour_value)
                # Remove leading and trailing spaces.
                colour_str = colour_str.strip()
                # Replace multiple spaces in the middle with a single space.
                colour_str = re.sub(r'\s+', ' ', colour_str)
                # Convert to lowercase for consistency.
                colour_normalized = colour_str.lower()
            
            # --- Step 3.4.5: Process Actual Quantities ---
            # Extract quantities from columns H, I, J, K, L.
            # These should be integers. Blank cells are treated as 0 WITHOUT warning.
            # Decimal values are rounded and reported.
            # Negative values are treated as 0 and reported.
            # Unparseable values are treated as 0 and reported.
            
            # Helper function to process a quantity value.
            def process_quantity(value, col_letter, quantity_name):
                # Check if the cell is blank.
                if pd.isna(value):
                    # Blank cell - store as 0 WITHOUT warning (as requested).
                    return 0
                else:
                    # Try to convert to a number.
                    try:
                        # Convert to float first.
                        quantity_float = float(value)
                        
                        # Check if the number is negative.
                        if quantity_float < 0:
                            # Negative numbers are not allowed. Store as 0 and report.
                            quantity_warnings.append(f"Sheet '{sheet_name}', Cell {col_letter}{excel_row_number} ({quantity_name}): '{quantity_float}' is negative, using 0")
                            return 0
                        
                        # Check if it's a whole number (like 100.0).
                        if quantity_float == int(quantity_float):
                            # It's a whole number. Store as int.
                            return int(quantity_float)
                        else:
                            # It has a decimal part (like 100.5). Round it and report.
                            quantity_rounded = round(quantity_float)
                            quantity_warnings.append(f"Sheet '{sheet_name}', Cell {col_letter}{excel_row_number} ({quantity_name}): '{quantity_float}' rounded to {quantity_rounded}")
                            return quantity_rounded
                            
                    except Exception as e:
                        # If we can't convert to a number, treat as 0 and report.
                        quantity_warnings.append(f"Sheet '{sheet_name}', Cell {col_letter}{excel_row_number} ({quantity_name}): '{value}' is not a valid number, using 0")
                        return 0
            
            # Process each quantity column.
            order_quantity = process_quantity(row['Order Quantity'], 'G', 'Order Quantity')  # ← ADDED
            actual_cutting = process_quantity(row['Cutting Quantity'], 'H', 'Actual Cutting')
            actual_sewing = process_quantity(row['Sewing Quantity'], 'I', 'Actual Sewing')
            actual_finishing = process_quantity(row['Finishing Quantity'], 'J', 'Actual Finishing')
            actual_washing = process_quantity(row['Washing Quantity'], 'K', 'Actual Washing')
            actual_packing = process_quantity(row['Packing Quantity'], 'L', 'Actual Packing')
            
            # --- Step 3.4.6: Create the dictionary for this row ---
            # The date comes from the sheet name, which we already parsed.
            row_dict = {
                'Style No': style_str,
                'PO': po_str,
                'Colour': colour_normalized,
                'Date': sheet_date_parsed.strftime('%d/%b/%y'),
                'Order Quantity': order_quantity,  # ← ADDED
                'Actual Cutting': actual_cutting,
                'Actual Sewing': actual_sewing,
                'Actual Finishing': actual_finishing,
                'Actual Washing': actual_washing,
                'Actual Packing': actual_packing,
                'Source Sheet': sheet_name,        # Metadata: which sheet this row came from
                'Source Row': excel_row_number     # Metadata: which row in the Excel file
            }
            
            result_list.append(row_dict)
    
    # --- Step 4: Print all warnings and errors ---    
    if blank_cells:
        print(f"\n⚠️  WARNING: The following cells are blank and should be filled:")
        for cell in blank_cells:
            print(f"   - {cell}")
    
    if quantity_warnings:
        print(f"\n⚠️  WARNING: The following quantity cells had issues:")
        for item in quantity_warnings:
            print(f"   - {item}")
    
    if skipped_sheets:
        print(f"\n❌ ERROR: The following sheets were skipped due to errors:")
        for sheet in skipped_sheets:
            print(f"   - Sheet '{sheet}'")
        print("Please fix the issues in these sheets and run the function again.")
    
    return result_list

def convert_cumulative_to_daywise_quantities_for_daily_prod(daily_prod_data):
    """
    Converts cumulative quantities in daily production data to day-wise quantities.
    
    The daily production file contains CUMULATIVE quantities (running totals).
    This function converts them to day-wise quantities by subtracting the previous
    occurrence of the same (Style, PO, Colour) combination.
    
    Logic:
    - For each row, find the most recent previous row with the same (Style, PO, Colour)
    - Subtract the previous quantities from the current quantities
    - If no previous row exists, the quantities remain as-is (first occurrence)
    
    Args:
        daily_prod_data: List of dictionaries from get_row_wise_data_from_daily_prod()
                        Each dict has keys: 'Style No', 'PO', 'Colour', 'Date',
                        'Actual Cutting', 'Actual Sewing', 'Actual Finishing',
                        'Actual Washing', 'Actual Packing', 'Source Sheet', 'Source Row'
    
    Returns:
        A list of dictionaries with the same structure, but with day-wise quantities
        instead of cumulative quantities.
    """
    
    from datetime import datetime
    
    # --- Step 1: Sort the data by date (chronologically) ---
    # We need to process rows in chronological order to properly calculate day-wise quantities.    
    # Create a list of tuples: (date_object, row)
    # This allows us to sort by actual date.
    dated_rows = []
    
    for row in daily_prod_data:
        date_str = row['Date']
        
        # Parse the date string to a datetime object.
        # The format is 'DD/Mon/YY' (e.g., '24/Sep/25')
        try:
            date_obj = datetime.strptime(date_str, '%d/%b/%y')
            dated_rows.append((date_obj, row))
        except Exception as e:
            print(f"   WARNING: Could not parse date '{date_str}' for row. Skipping this row.")
            print(f"   Row details: Style={row.get('Style No')}, PO={row.get('PO')}, Colour={row.get('Colour')}")
            continue
    
    # Sort by date (earliest first).
    dated_rows.sort(key=lambda x: x[0])
    
    # Removed: Processing count message
    
    # --- Step 2: Build a dictionary to track the most recent quantities for each combo ---
    # Structure: {(style, po, colour): {'date': date_obj, 'quantities': {...}}}
    # This will help us find the previous occurrence of each combo.
    
    previous_quantities_by_combo = {}
    
    # --- Step 3: Process each row and convert cumulative to day-wise ---
    result_list = []
    
    for date_obj, row in dated_rows:
        # Extract the key fields.
        style = row['Style No'].strip().upper()
        po = row['PO'].strip()
        colour = row['Colour'].strip().lower()
        date_str = row['Date']
        
        # Create a unique key for this combination.
        combo_key = (style, po, colour)
        
        # Get the current cumulative quantities.
        cumulative_cutting = row['Actual Cutting']
        cumulative_sewing = row['Actual Sewing']
        cumulative_finishing = row['Actual Finishing']
        cumulative_washing = row['Actual Washing']
        cumulative_packing = row['Actual Packing']
        
        # --- Step 3.1: Check if we have a previous occurrence of this combo ---
        if combo_key in previous_quantities_by_combo:
            # We have a previous occurrence. Subtract to get day-wise quantities.
            previous_data = previous_quantities_by_combo[combo_key]
            previous_date = previous_data['date']
            previous_quantities = previous_data['quantities']
            
            # Calculate day-wise quantities by subtracting previous cumulative from current cumulative.
            daywise_cutting = cumulative_cutting - previous_quantities['cutting']
            daywise_sewing = cumulative_sewing - previous_quantities['sewing']
            daywise_finishing = cumulative_finishing - previous_quantities['finishing']
            daywise_washing = cumulative_washing - previous_quantities['washing']
            daywise_packing = cumulative_packing - previous_quantities['packing']
            
            # Check if any quantity decreased (which would be unusual).
            if (daywise_cutting < 0 or daywise_sewing < 0 or daywise_finishing < 0 or 
                daywise_washing < 0 or daywise_packing < 0):
                # Print as single formatted warning message
                warning_msg = (
                    f"\n   ⚠️  WARNING: Negative day-wise quantity detected!\n"
                    f"Style: {style}, PO: {po}, Colour: {colour}\n"
                    f"Current date: {date_str}, Previous date: {previous_date.strftime('%d/%b/%y')}\n"
                    f"Using 0 for negative values."
                )
                print(f"\n{warning_msg}\n")

                # Set negative values to 0 (can't have negative production).
                daywise_cutting = max(0, daywise_cutting)
                daywise_sewing = max(0, daywise_sewing)
                daywise_finishing = max(0, daywise_finishing)
                daywise_washing = max(0, daywise_washing)
                daywise_packing = max(0, daywise_packing)
        else:
            # This is the first occurrence of this combo. Use cumulative as day-wise.
            daywise_cutting = cumulative_cutting
            daywise_sewing = cumulative_sewing
            daywise_finishing = cumulative_finishing
            daywise_washing = cumulative_washing
            daywise_packing = cumulative_packing
        
        # --- Step 3.2: Create the updated row with day-wise quantities ---
        updated_row = {
            'Style No': row['Style No'],  # Keep original formatting
            'PO': row['PO'],
            'Colour': row['Colour'],
            'Date': row['Date'],
            'Actual Cutting': daywise_cutting,
            'Actual Sewing': daywise_sewing,
            'Actual Finishing': daywise_finishing,
            'Actual Washing': daywise_washing,
            'Actual Packing': daywise_packing,
            'Source Sheet': row.get('Source Sheet', ''),
            'Source Row': row.get('Source Row', '')
        }
        
        result_list.append(updated_row)
        
        # --- Step 3.3: Update the tracking dictionary with current cumulative quantities ---
        # This becomes the "previous" for the next occurrence of this combo.
        previous_quantities_by_combo[combo_key] = {
            'date': date_obj,
            'quantities': {
                'cutting': cumulative_cutting,
                'sewing': cumulative_sewing,
                'finishing': cumulative_finishing,
                'washing': cumulative_washing,
                'packing': cumulative_packing
            }
        }
    return result_list

def match_plan_with_actual(plan_data, daily_prod_data, style_number):
    """
    Matches plan data with actual production data for a given style.
    
    For each unique (Style, PO, Colour) combination, this function:
    - Combines all dates from both plan and actual production
    - Creates a row for each date with both planned and actual quantities
    - Handles cases where dates exist in one dataset but not the other
    - Validates that actual production only happens for planned style/PO/colour combos
    
    Args:
        plan_data: List of dictionaries from get_row_wise_data_from_plan()
                   Each dict has keys: 'Style No', 'PO', 'Colour', 'Date',
                   'Planned Cutting', 'Planned Sewing', 'Planned Washing',
                   'Planned Finishing', 'Planned Packing', 'Source Sheet', 'Source Row'
        
        daily_prod_data: List of dictionaries from get_row_wise_data_from_daily_prod()
                        Each dict has keys: 'Style No', 'PO', 'Colour', 'Date',
                        'Actual Cutting', 'Actual Sewing', 'Actual Finishing',
                        'Actual Washing', 'Actual Packing', 'Source Sheet', 'Source Row'
        
        style_number: The style number to process (string)
    
    Returns:
        A list of dictionaries, where each dictionary represents one matched row with keys:
        'Style No', 'PO', 'Colour', 'Date',
        'Planned Cutting', 'Planned Sewing', 'Planned Washing', 'Planned Finishing', 'Planned Packing',
        'Actual Cutting', 'Actual Sewing', 'Actual Finishing', 'Actual Washing', 'Actual Packing'
    """
    
    # --- Step 1: Filter data for the specific style ---
    # We only want to process data for the given style number.
    
    # Filter plan data for this style.
    plan_for_style = []
    for row in plan_data:
        # Check if this row belongs to the current style.
        # The style might be stored with or without leading/trailing spaces.
        if row['Style No'].strip().upper() == style_number.strip().upper():
            plan_for_style.append(row)
    
    # Filter daily production data for this style.
    actual_for_style = []
    for row in daily_prod_data:
        # Check if this row belongs to the current style.
        if row['Style No'].strip().upper() == style_number.strip().upper():
            actual_for_style.append(row)
    
    
    # --- Step 2: Build dictionaries organized by (PO, Colour) combination ---
    # We'll create a structure like:
    # {
    #   ('4201959', 'black'): {
    #       'plan_dates': {'15/Sep/25': {...row data...}, '16/Sep/25': {...row data...}},
    #       'actual_dates': {'15/Sep/25': {...row data...}, '16/Sep/25': {...row data...}}
    #   },
    #   ...
    # }
    
    plan_by_combo = {}  # Dictionary to organize plan data
    actual_by_combo = {}  # Dictionary to organize actual data
    
    # --- Step 2.1: Organize plan data ---
    # Also build sets of all POs and colours that exist in the plan for this style.
    plan_pos = set()  # All PO numbers in the plan
    plan_colours_by_po = {}  # Dictionary: PO -> set of colours for that PO
    
    for row in plan_for_style:
        # Create a unique key for this PO/Colour combination.
        # We use lowercase and stripped values to ensure matching works correctly.
        po = row['PO'].strip()
        colour = row['Colour'].strip().lower()
        combo_key = (po, colour)
        
        # Track this PO and colour.
        plan_pos.add(po)
        if po not in plan_colours_by_po:
            plan_colours_by_po[po] = set()
        plan_colours_by_po[po].add(colour)
        
        # If this is the first time we see this combo, initialize the dictionary.
        if combo_key not in plan_by_combo:
            plan_by_combo[combo_key] = {}
        
        # Store this row indexed by its date.
        date = row['Date']
        plan_by_combo[combo_key][date] = row
    
    # --- Step 2.2: Organize actual data ---
    for row in actual_for_style:
        # Create a unique key for this PO/Colour combination.
        po = row['PO'].strip()
        colour = row['Colour'].strip().lower()
        combo_key = (po, colour)
        
        # If this is the first time we see this combo, initialize the dictionary.
        if combo_key not in actual_by_combo:
            actual_by_combo[combo_key] = {}
        
        # Store this row indexed by its date.
        date = row['Date']
        actual_by_combo[combo_key][date] = row
    
    # --- Step 3: Validate actual production with detailed error messages ---
    # Check each actual production row and determine exactly what's wrong if it's not in the plan.
    
    unplanned_rows = []  # List to track all unplanned production with details
    
    for combo_key in actual_by_combo.keys():
        po, colour = combo_key
        
        # Check if this combo exists in the plan.
        if combo_key not in plan_by_combo:
            # This combo is NOT in the plan. Determine what specifically is wrong.
            
            # Get all rows for this combo from actual production (to report source info).
            for date, actual_row in actual_by_combo[combo_key].items():
                source_sheet = actual_row.get('Source Sheet', 'Unknown')
                source_row = actual_row.get('Source Row', 'Unknown')
                
                # Determine the specific issue.
                if po not in plan_pos:
                    # The PO itself doesn't exist in the plan for this style.
                    error_detail = f"PO '{po}' does not exist in the plan for style {style_number}"
                elif po in plan_colours_by_po and colour not in plan_colours_by_po[po]:
                    # The PO exists, but this colour doesn't exist for this PO.
                    available_colours = ', '.join(sorted(plan_colours_by_po[po]))
                    error_detail = f"Colour '{colour}' does not exist for PO '{po}' in the plan (available colours: {available_colours})"
                else:
                    # This shouldn't happen, but just in case.
                    error_detail = f"PO '{po}' with colour '{colour}' does not exist in the plan"
                
                # Record this unplanned row.
                unplanned_rows.append({
                    'po': po,
                    'colour': colour,
                    'date': date,
                    'source_sheet': source_sheet,
                    'source_row': source_row,
                    'error_detail': error_detail
                })
    
    # Print detailed warnings for unplanned production.
    if unplanned_rows:
        print(f"\n⚠️  WARNING: Found actual production for style {style_number} that is NOT in the plan:")
        print("   These rows will be IGNORED as we should only produce planned items.\n")
        
        for item in unplanned_rows:
            print(f"   ❌ Sheet '{item['source_sheet']}', Row {item['source_row']}:")
            print(f"      {item['error_detail']}")
            print(f"      (Date: {item['date']}, PO: {item['po']}, Colour: {item['colour']})\n")
    
    # --- Step 4: Build the matched rows ---
    # For each PO/Colour combo in the PLAN, create rows for all dates.
    matched_rows = []
    
    # Get all PO/Colour combinations from the plan (these are the only ones we care about).
    for combo_key in sorted(plan_by_combo.keys()):
        po, colour = combo_key
        
        # Get all dates from plan for this combo.
        plan_dates = set(plan_by_combo[combo_key].keys())
        
        # Get all dates from actual for this combo (if it exists).
        actual_dates = set()
        if combo_key in actual_by_combo:
            actual_dates = set(actual_by_combo[combo_key].keys())
        
        # Combine all dates (union of plan and actual).
        all_dates = plan_dates.union(actual_dates)
        
        # Sort the dates chronologically.
        # The dates are in format 'DD/Mon/YY' (e.g., '15/Sep/25').
        # We need to convert them to datetime objects for sorting.
        from datetime import datetime
        
        # Convert date strings to datetime objects for sorting.
        date_objects = []
        for date_str in all_dates:
            try:
                date_obj = datetime.strptime(date_str, '%d/%b/%y')
                date_objects.append((date_str, date_obj))
            except:
                # If we can't parse the date, just skip it.
                print(f"   WARNING: Could not parse date '{date_str}' for sorting. Skipping.")
                continue
        
        # Sort by the datetime object.
        date_objects.sort(key=lambda x: x[1])
        sorted_dates = [date_str for date_str, date_obj in date_objects]
        
        # --- Step 4.1: Create a row for each date ---
        for date in sorted_dates:
            # Initialize the matched row with basic information.
            matched_row = {
                'Style No': style_number,
                'PO': po,
                'Colour': colour,
                'Date': date
            }
            
            # --- Step 4.1.1: Get planned quantities for this date ---
            if date in plan_by_combo[combo_key]:
                # This date exists in the plan. Get the planned quantities.
                plan_row = plan_by_combo[combo_key][date]
                matched_row['Planned Cutting'] = plan_row['Planned Cutting']
                matched_row['Planned Sewing'] = plan_row['Planned Sewing']
                matched_row['Planned Washing'] = plan_row['Planned Washing']
                matched_row['Planned Finishing'] = plan_row['Planned Finishing']
                matched_row['Planned Packing'] = plan_row['Planned Packing']
            else:
                # This date does NOT exist in the plan. Set planned quantities to 0.
                # This happens when actual production extends beyond the plan dates.
                matched_row['Planned Cutting'] = 0
                matched_row['Planned Sewing'] = 0
                matched_row['Planned Washing'] = 0
                matched_row['Planned Finishing'] = 0
                matched_row['Planned Packing'] = 0
            
            # --- Step 4.1.2: Get actual quantities for this date ---
            if combo_key in actual_by_combo and date in actual_by_combo[combo_key]:
                # This date exists in actual production. Get the actual quantities.
                actual_row = actual_by_combo[combo_key][date]
                matched_row['Actual Cutting'] = actual_row['Actual Cutting']
                matched_row['Actual Sewing'] = actual_row['Actual Sewing']
                matched_row['Actual Finishing'] = actual_row['Actual Finishing']
                matched_row['Actual Washing'] = actual_row['Actual Washing']
                matched_row['Actual Packing'] = actual_row['Actual Packing']
            else:
                # This date does NOT exist in actual production. Set actual quantities to 0.
                # This happens when the plan includes dates that haven't been produced yet.
                matched_row['Actual Cutting'] = 0
                matched_row['Actual Sewing'] = 0
                matched_row['Actual Finishing'] = 0
                matched_row['Actual Washing'] = 0
                matched_row['Actual Packing'] = 0
            
            # Add this matched row to our result list.
            matched_rows.append(matched_row)

    return matched_rows

def delete_empty_rows(matched_data):
    """
    Removes rows from matched data where ALL planned and actual quantities are zero.
    
    A row is kept if ANY of the following quantities is non-zero:
    - Planned Cutting, Planned Sewing, Planned Washing, Planned Finishing, Planned Packing
    - Actual Cutting, Actual Sewing, Actual Finishing, Actual Washing, Actual Packing
    
    A row is deleted ONLY if ALL 10 quantities are zero.
    
    Args:
        matched_data: List of dictionaries from match_plan_with_actual()
                     Each dict has keys: 'Style No', 'PO', 'Colour', 'Date',
                     'Planned Cutting', 'Planned Sewing', 'Planned Washing', 'Planned Finishing', 'Planned Packing',
                     'Actual Cutting', 'Actual Sewing', 'Actual Finishing', 'Actual Washing', 'Actual Packing'
    
    Returns:
        A new list of dictionaries with the same structure, but with empty rows removed.
    """
    
    # Create a new list to hold the filtered rows.
    filtered_rows = []
    
    # Track how many rows we deleted for reporting.
    deleted_count = 0
    
    # --- Step 1: Loop through each row and check if it's empty ---
    for row in matched_data:
        
        # --- Step 1.1: Extract all planned quantities ---
        planned_cutting = row.get('Planned Cutting', 0)
        planned_sewing = row.get('Planned Sewing', 0)
        planned_washing = row.get('Planned Washing', 0)
        planned_finishing = row.get('Planned Finishing', 0)
        planned_packing = row.get('Planned Packing', 0)
        
        # --- Step 1.2: Extract all actual quantities ---
        actual_cutting = row.get('Actual Cutting', 0)
        actual_sewing = row.get('Actual Sewing', 0)
        actual_finishing = row.get('Actual Finishing', 0)
        actual_washing = row.get('Actual Washing', 0)
        actual_packing = row.get('Actual Packing', 0)
        
        # --- Step 1.3: Check if ALL quantities are zero ---
        # We check if the sum of all quantities is zero.
        # This is simpler than checking each one individually.
        total_quantity = (
            planned_cutting + planned_sewing + planned_washing + planned_finishing + planned_packing +
            actual_cutting + actual_sewing + actual_finishing + actual_washing + actual_packing
        )
        
        # --- Step 1.4: Decide whether to keep or delete this row ---
        if total_quantity == 0:
            # All quantities are zero. Delete this row (don't add it to filtered_rows).
            deleted_count += 1
        else:
            # At least one quantity is non-zero. Keep this row.
            filtered_rows.append(row)
    
    return filtered_rows

def add_cumulative_columns_to_matched_dict(matched_data):
    """
    Adds cumulative columns and day-wise difference columns to matched data.
    
    For each row, this function adds:
    - Day (day of the week)
    - Day Actual - Day Planned (for each process: Cutting, Sewing, Washing, Finishing, Packing)
    - Cumulative Planned (for each process)
    - Cumulative Actual (for each process)
    - Cumulative Actual - Cumulative Planned (for each process)
    
    IMPORTANT: Cumulative quantities are calculated at the STYLE/PO/COLOUR level.
    Each unique combination of Style No, PO, and Colour has its own cumulative tracking.
    
    Args:
        matched_data: List of dictionaries from delete_empty_rows()
                     Each dict has keys: 'Style No', 'PO', 'Colour', 'Date',
                     'Planned Cutting', 'Planned Sewing', 'Planned Washing', 'Planned Finishing', 'Planned Packing',
                     'Actual Cutting', 'Actual Sewing', 'Actual Finishing', 'Actual Washing', 'Actual Packing'
    
    Returns:
        A new list of dictionaries with all the original columns plus the new cumulative columns.
        The rows are sorted by Style No, PO, Colour, and then by date.
    """
    
    # If the input is empty, return empty list.
    if not matched_data:
        return []
    
    # --- Step 1: Parse dates and group rows by Style/PO/Colour ---
    # Create a dictionary: {(Style No, PO, Colour): [(row, date_obj), ...]}
    
    grouped_rows = {}
    
    for row in matched_data:
        date_str = row['Date']
        
        try:
            # Parse the date string (format: DD/Mon/YY, e.g., '15/Sep/25').
            date_obj = datetime.strptime(date_str, '%d/%b/%y')
            
            # Create a key for this group: (Style No, PO, Colour)
            group_key = (row.get('Style No', ''), row.get('PO', ''), row.get('Colour', ''))
            
            # Add this row to the appropriate group.
            if group_key not in grouped_rows:
                grouped_rows[group_key] = []
            
            grouped_rows[group_key].append((row, date_obj))
            
        except Exception as e:
            # If we can't parse the date, print a warning and skip this row.
            print(f"WARNING: Could not parse date '{date_str}'. Skipping this row.")
            continue
    
    # --- Step 2: Process each group separately ---
    result_rows = []
    
    for group_key, rows_with_dates in grouped_rows.items():
        
        # Sort this group's rows by date (chronologically).
        rows_with_dates.sort(key=lambda x: x[1])
        
        # --- Step 2.1: Initialize cumulative tracking variables for this group ---
        cumulative_planned_cutting = 0
        cumulative_planned_sewing = 0
        cumulative_planned_washing = 0
        cumulative_planned_finishing = 0
        cumulative_planned_packing = 0
        
        cumulative_actual_cutting = 0
        cumulative_actual_sewing = 0
        cumulative_actual_washing = 0
        cumulative_actual_finishing = 0
        cumulative_actual_packing = 0
        
        # Track the current date we're processing.
        current_date = None
        
        # Track the cumulative values for the current date.
        # All rows with the same date will get these values.
        current_date_cumulative_planned_cutting = 0
        current_date_cumulative_planned_sewing = 0
        current_date_cumulative_planned_washing = 0
        current_date_cumulative_planned_finishing = 0
        current_date_cumulative_planned_packing = 0
        
        current_date_cumulative_actual_cutting = 0
        current_date_cumulative_actual_sewing = 0
        current_date_cumulative_actual_washing = 0
        current_date_cumulative_actual_finishing = 0
        current_date_cumulative_actual_packing = 0
        
        # Track the total quantities for the current date.
        current_date_total_planned_cutting = 0
        current_date_total_planned_sewing = 0
        current_date_total_planned_washing = 0
        current_date_total_planned_finishing = 0
        current_date_total_planned_packing = 0
        
        current_date_total_actual_cutting = 0
        current_date_total_actual_sewing = 0
        current_date_total_actual_washing = 0
        current_date_total_actual_finishing = 0
        current_date_total_actual_packing = 0
        
        # --- Step 2.2: Process each row in this group ---
        for row, date_obj in rows_with_dates:
            
            # --- Step 2.2.1: Check if we've moved to a new date ---
            if current_date is None or date_obj != current_date:
                # We've moved to a new date.
                
                # If this is not the first date, update the cumulative values.
                if current_date is not None:
                    # Add the totals from the previous date to the cumulative values.
                    cumulative_planned_cutting += current_date_total_planned_cutting
                    cumulative_planned_sewing += current_date_total_planned_sewing
                    cumulative_planned_washing += current_date_total_planned_washing
                    cumulative_planned_finishing += current_date_total_planned_finishing
                    cumulative_planned_packing += current_date_total_planned_packing
                    
                    cumulative_actual_cutting += current_date_total_actual_cutting
                    cumulative_actual_sewing += current_date_total_actual_sewing
                    cumulative_actual_washing += current_date_total_actual_washing
                    cumulative_actual_finishing += current_date_total_actual_finishing
                    cumulative_actual_packing += current_date_total_actual_packing
                
                # Reset the current date totals.
                current_date_total_planned_cutting = 0
                current_date_total_planned_sewing = 0
                current_date_total_planned_washing = 0
                current_date_total_planned_finishing = 0
                current_date_total_planned_packing = 0
                
                current_date_total_actual_cutting = 0
                current_date_total_actual_sewing = 0
                current_date_total_actual_washing = 0
                current_date_total_actual_finishing = 0
                current_date_total_actual_packing = 0
                
                # Update the current date.
                current_date = date_obj
                
                # First pass: Calculate the total quantities for this date.
                # We need to look ahead and sum all rows with the same date.
                for future_row, future_date_obj in rows_with_dates:
                    if future_date_obj == current_date:
                        current_date_total_planned_cutting += future_row.get('Planned Cutting', 0)
                        current_date_total_planned_sewing += future_row.get('Planned Sewing', 0)
                        current_date_total_planned_washing += future_row.get('Planned Washing', 0)
                        current_date_total_planned_finishing += future_row.get('Planned Finishing', 0)
                        current_date_total_planned_packing += future_row.get('Planned Packing', 0)
                        
                        current_date_total_actual_cutting += future_row.get('Actual Cutting', 0)
                        current_date_total_actual_sewing += future_row.get('Actual Sewing', 0)
                        current_date_total_actual_washing += future_row.get('Actual Washing', 0)
                        current_date_total_actual_finishing += future_row.get('Actual Finishing', 0)
                        current_date_total_actual_packing += future_row.get('Actual Packing', 0)
                
                # Calculate the cumulative values for this date.
                # This includes the previous cumulative + the total for this date.
                current_date_cumulative_planned_cutting = cumulative_planned_cutting + current_date_total_planned_cutting
                current_date_cumulative_planned_sewing = cumulative_planned_sewing + current_date_total_planned_sewing
                current_date_cumulative_planned_washing = cumulative_planned_washing + current_date_total_planned_washing
                current_date_cumulative_planned_finishing = cumulative_planned_finishing + current_date_total_planned_finishing
                current_date_cumulative_planned_packing = cumulative_planned_packing + current_date_total_planned_packing
                
                current_date_cumulative_actual_cutting = cumulative_actual_cutting + current_date_total_actual_cutting
                current_date_cumulative_actual_sewing = cumulative_actual_sewing + current_date_total_actual_sewing
                current_date_cumulative_actual_washing = cumulative_actual_washing + current_date_total_actual_washing
                current_date_cumulative_actual_finishing = cumulative_actual_finishing + current_date_total_actual_finishing
                current_date_cumulative_actual_packing = cumulative_actual_packing + current_date_total_actual_packing
            
            # --- Step 2.2.2: Extract the planned and actual quantities for this row ---
            planned_cutting = row.get('Planned Cutting', 0)
            planned_sewing = row.get('Planned Sewing', 0)
            planned_washing = row.get('Planned Washing', 0)
            planned_finishing = row.get('Planned Finishing', 0)
            planned_packing = row.get('Planned Packing', 0)
            
            actual_cutting = row.get('Actual Cutting', 0)
            actual_sewing = row.get('Actual Sewing', 0)
            actual_washing = row.get('Actual Washing', 0)
            actual_finishing = row.get('Actual Finishing', 0)
            actual_packing = row.get('Actual Packing', 0)
            
            # --- Step 2.2.3: Calculate the day of the week ---
            day_of_week = date_obj.strftime('%A')  # e.g., 'Monday', 'Tuesday', etc.
            
            # --- Step 2.2.4: Calculate day-wise differences (Actual - Planned) ---
            day_diff_cutting = actual_cutting - planned_cutting
            day_diff_sewing = actual_sewing - planned_sewing
            day_diff_washing = actual_washing - planned_washing
            day_diff_finishing = actual_finishing - planned_finishing
            day_diff_packing = actual_packing - planned_packing
            
            # --- Step 2.2.5: Calculate cumulative differences (Cumulative Actual - Cumulative Planned) ---
            cumulative_diff_cutting = current_date_cumulative_actual_cutting - current_date_cumulative_planned_cutting
            cumulative_diff_sewing = current_date_cumulative_actual_sewing - current_date_cumulative_planned_sewing
            cumulative_diff_washing = current_date_cumulative_actual_washing - current_date_cumulative_planned_washing
            cumulative_diff_finishing = current_date_cumulative_actual_finishing - current_date_cumulative_planned_finishing
            cumulative_diff_packing = current_date_cumulative_actual_packing - current_date_cumulative_planned_packing
            
            # --- Step 2.2.6: Create the new row with all columns ---
            new_row = {
                'Style No': row.get('Style No', ''),
                'PO': row.get('PO', ''),
                'Colour': row.get('Colour', ''),
                'Date': row.get('Date', ''),
                'Day': day_of_week,
                
                # Cutting
                'Planned Cutting': planned_cutting,
                'Actual Cutting': actual_cutting,
                'Day Actual - Day Planned Cutting': day_diff_cutting,
                'Cumulative Planned Cutting': current_date_cumulative_planned_cutting,
                'Cumulative Actual Cutting': current_date_cumulative_actual_cutting,
                'Cumulative Actual - Cumulative Planned Cutting': cumulative_diff_cutting,
                
                # Sewing
                'Planned Sewing': planned_sewing,
                'Actual Sewing': actual_sewing,
                'Day Actual - Day Planned Sewing': day_diff_sewing,
                'Cumulative Planned Sewing': current_date_cumulative_planned_sewing,
                'Cumulative Actual Sewing': current_date_cumulative_actual_sewing,
                'Cumulative Actual - Cumulative Planned Sewing': cumulative_diff_sewing,
                
                # Washing
                'Planned Washing': planned_washing,
                'Actual Washing': actual_washing,
                'Day Actual - Day Planned Washing': day_diff_washing,
                'Cumulative Planned Washing': current_date_cumulative_planned_washing,
                'Cumulative Actual Washing': current_date_cumulative_actual_washing,
                'Cumulative Actual - Cumulative Planned Washing': cumulative_diff_washing,
                
                # Finishing
                'Planned Finishing': planned_finishing,
                'Actual Finishing': actual_finishing,
                'Day Actual - Day Planned Finishing': day_diff_finishing,
                'Cumulative Planned Finishing': current_date_cumulative_planned_finishing,
                'Cumulative Actual Finishing': current_date_cumulative_actual_finishing,
                'Cumulative Actual - Cumulative Planned Finishing': cumulative_diff_finishing,
                
                # Packing
                'Planned Packing': planned_packing,
                'Actual Packing': actual_packing,
                'Day Actual - Day Planned Packing': day_diff_packing,
                'Cumulative Planned Packing': current_date_cumulative_planned_packing,
                'Cumulative Actual Packing': current_date_cumulative_actual_packing,
                'Cumulative Actual - Cumulative Planned Packing': cumulative_diff_packing
            }
            
            result_rows.append(new_row)
    
    # --- Step 3: Sort the final result by Style No, PO, Colour, and Date ---
    # This ensures the output is organized and easy to read.
    result_rows.sort(key=lambda x: (
        x['Style No'],
        x['PO'],
        x['Colour'],
        datetime.strptime(x['Date'], '%d/%b/%y')
    ))
    
    return result_rows


from datetime import datetime
import pandas as pd

def write_production_report_to_excel(matched_data_by_style, output_file_path):
    """
    Writes the matched production data to an Excel file with one sheet per style.
    
    Each sheet contains all 35 columns including:
    - Style No, PO, Colour, Date, Day
    - Planned quantities, Actual quantities, Day differences, Cumulative quantities, Cumulative differences
    - For each process: Cutting, Sewing, Washing, Finishing, Packing
    
    Rows are sorted by date (chronologically), so all rows with the same date appear together.
    
    Args:
        matched_data_by_style: Dictionary where keys are style numbers (strings)
                              and values are lists of row dictionaries.
                              Example: {
                                  '9KLXL8': [{row1}, {row2}, ...],
                                  'YPYWM3': [{row1}, {row2}, ...]
                              }
        
        output_file_path: Full path where the Excel file should be saved.
                         Example: '/path/to/production_report.xlsx'
    
    Returns:
        None. The Excel file is created at the specified path.
    """
    
    # --- Step 1: Create an Excel writer object ---
    # This allows us to write multiple sheets to the same Excel file.
    try:
        writer = pd.ExcelWriter(output_file_path, engine='openpyxl')
    except Exception as e:
        print(f"   ❌ ERROR: Could not create Excel writer. Error: {e}")
        return
    
    # --- Step 2: Process each style and create a sheet ---
    total_styles = len(matched_data_by_style)
    
    for style_number, rows in matched_data_by_style.items():
        
        # --- Step 2.1: Check if we have data ---
        if len(rows) == 0:
            print(f"      ⚠️  WARNING: No data for style {style_number}. Skipping this sheet.")
            continue
        
        # --- Step 2.2: Sort rows by date ---
        # We need to sort the rows chronologically before writing to Excel.
        # This ensures all rows with the same date appear together.
        
        rows_with_dates = []
        
        for row in rows:
            date_str = row.get('Date', '')
            
            try:
                # Parse the date string (format: DD/Mon/YY, e.g., '15/Sep/25').
                date_obj = datetime.strptime(date_str, '%d/%b/%y')
                rows_with_dates.append((row, date_obj))
            except Exception as e:
                # If we can't parse the date, put this row at the end.
                print(f"      WARNING: Could not parse date '{date_str}' for style {style_number}. Row will be placed at the end.")
                rows_with_dates.append((row, datetime.max))
        
        # Sort by date (chronologically).
        rows_with_dates.sort(key=lambda x: x[1])
        
        # Extract just the rows (without the date objects).
        sorted_rows = [row for row, date_obj in rows_with_dates]
        
        # --- Step 2.3: Convert the list of dictionaries to a DataFrame ---
        # pandas makes it easy to write DataFrames to Excel.
        df = pd.DataFrame(sorted_rows)
        
        # --- Step 2.4: Reorder columns to match the desired layout ---
        # We want the columns in the exact order specified by the user.
        column_order = [
            'Style No',
            'PO',
            'Colour',
            'Order Quantity',  # ← ADDED
            'Date',
            'Day',
            
            # Cutting
            'Planned Cutting',
            'Actual Cutting',
            'Day Actual - Day Planned Cutting',
            'Cumulative Planned Cutting',
            'Cumulative Actual Cutting',
            'Cumulative Actual - Cumulative Planned Cutting',
            
            # Sewing
            'Planned Sewing',
            'Actual Sewing',
            'Day Actual - Day Planned Sewing',
            'Cumulative Planned Sewing',
            'Cumulative Actual Sewing',
            'Cumulative Actual - Cumulative Planned Sewing',
            
            # Washing
            'Planned Washing',
            'Actual Washing',
            'Day Actual - Day Planned Washing',
            'Cumulative Planned Washing',
            'Cumulative Actual Washing',
            'Cumulative Actual - Cumulative Planned Washing',
            
            # Finishing
            'Planned Finishing',
            'Actual Finishing',
            'Day Actual - Day Planned Finishing',
            'Cumulative Planned Finishing',
            'Cumulative Actual Finishing',
            'Cumulative Actual - Cumulative Planned Finishing',
            
            # Packing
            'Planned Packing',
            'Actual Packing',
            'Day Actual - Day Planned Packing',
            'Cumulative Planned Packing',
            'Cumulative Actual Packing',
            'Cumulative Actual - Cumulative Planned Packing'
        ]
        
        # Check if all expected columns exist.
        missing_columns = [col for col in column_order if col not in df.columns]
        if missing_columns:
            print(f"      ⚠️  WARNING: Missing columns for style {style_number}: {missing_columns}")
            print(f"      Available columns: {list(df.columns)}")
            print(f"      Using available columns only.")
            # Use only the columns that exist.
            column_order = [col for col in column_order if col in df.columns]
        
        # Reorder the DataFrame columns.
        df = df[column_order]
        
        # --- Step 2.5: Write the DataFrame to a sheet ---
        # The sheet name is the style number.
        # Excel sheet names have a 31-character limit, so we'll truncate if needed.
        sheet_name = str(style_number)[:31]
        
        try:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        except Exception as e:
            print(f"      ❌ ERROR: Could not write sheet '{sheet_name}'. Error: {e}")
            continue
    
    # --- Step 3: Save and close the Excel file ---
    try:
        writer.close()
        print(f"\n✅ Excel file created successfully!")
    except Exception as e:
        print(f"\n❌ ERROR: Could not save Excel file. Error: {e}")
        return

def do_everything(plan_file_path, daily_prod_file_path, output_file_path):
    """
    Master function that orchestrates the entire production report generation workflow.
    
    This function:
    1. Extracts style numbers from the plan file
    2. Extracts row-wise data from the plan for each style
    3. Extracts row-wise data from the daily production file
    4. Converts cumulative quantities to day-wise quantities
    5. Matches plan data with actual production data for each style
    6. Writes the final production report to an Excel file
    
    Args:
        plan_file_path: Full path to the plan Excel file.
        daily_prod_file_path: Full path to the daily production Excel file.
        output_file_path: Full path where the output Excel file should be saved.
    
    Returns:
        None. The production report Excel file is created at the specified output path.
    """
    
    # Step 1: Get list of style numbers from the plan.
    style_numbers = get_style_numbers_from_plan(plan_file_path)
    
    if not style_numbers:
        print("ERROR: No valid style numbers found in the plan file.")
        return
    
    # Step 2: Extract row-wise data from daily production file.
    daily_prod_data = get_row_wise_data_from_daily_prod(daily_prod_file_path)
    
    if not daily_prod_data:
        print("ERROR: No data extracted from daily production file.")
        return
    
    # ========== ADDED BLOCK START: Build Order Quantity lookup from raw DPR data ==========
    # We build this lookup BEFORE converting to day-wise, because Order Quantity is a fixed
    # value per (Style, PO, Colour) combination and doesn't change across dates.
    order_quantity_lookup = {}
    for row in daily_prod_data:
        style_upper = row['Style No'].strip().upper()
        po = row['PO'].strip()
        colour_lower = row['Colour'].strip().lower()
        combo_key = (style_upper, po, colour_lower)
        # Only store the first occurrence (Order Quantity is the same for all dates)
        if combo_key not in order_quantity_lookup:
            order_quantity_lookup[combo_key] = row.get('Order Quantity', 0)
    # ========== ADDED BLOCK END ==========

    # Step 3: Convert cumulative quantities to day-wise quantities.
    daily_prod_daywise = convert_cumulative_to_daywise_quantities_for_daily_prod(daily_prod_data)
    
    if not daily_prod_daywise:
        print("ERROR: Conversion to day-wise quantities failed.")
        return
    
    # Step 4: Process each style - extract plan data and match with actual.
    matched_data_by_style = {}
    
    for style_number in style_numbers:
        # Extract plan data for this style.
        plan_data_for_style = get_row_wise_data_from_plan(plan_file_path, style_number)
        
        if not plan_data_for_style:
            continue
        
        # Match plan with actual production.
        matched_rows = match_plan_with_actual(plan_data_for_style, daily_prod_daywise, style_number)
        
        if not matched_rows:
            continue

        matched_rows = delete_empty_rows(matched_rows)

        if not matched_rows:
            continue

        matched_rows = add_cumulative_columns_to_matched_dict(matched_rows)

        if not matched_rows:
            continue
        
        # ========== ADDED BLOCK START: Stamp Order Quantity onto each matched row ==========
        for row in matched_rows:
            style_upper = row['Style No'].strip().upper()
            po = row['PO'].strip()
            colour_lower = row['Colour'].strip().lower()
            combo_key = (style_upper, po, colour_lower)
            row['Order Quantity'] = order_quantity_lookup.get(combo_key, 0)
        # ========== ADDED BLOCK END ==========

        # Store the matched data.
        matched_data_by_style[style_number] = matched_rows
    
    # Check if we have any data to write.
    if not matched_data_by_style:
        print("ERROR: No matched data for any style.")
        return
    
    # Step 5: Write the production report to Excel.
    write_production_report_to_excel(matched_data_by_style, output_file_path)
    
    # Done!
    print(f"\n✅ Generated Successfully: {output_file_path}")

do_everything("new_plan.xlsx", "daily_prod_report_2.xlsx", "collated_production.xlsx")
