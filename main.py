import google.generativeai as genai
import pandas as pd
import json # To parse JSON response
import io
import re
import os
from mistralai import Mistral
import base64
import markdown



api_key = #enter mistral ocr key here
client = Mistral(api_key=api_key)

ocr_response = client.ocr.process(
    model="mistral-ocr-latest",
    document={
        "type": "document_url",
        "document_url": "https://www.tatasteel.com/media/21242/business-responsibility-and-sustainability-report.pdf"
    },
    include_image_base64=True
)
for page in ocr_response.pages:
    formatted_text = markdown.markdown(page.markdown)
    print(formatted_text)
# --- Configuration ---
API_KEY = # PASTE YOUR API KEY HERE
HTML_FILE_PATH = r"C:\Unified coding\Projects\Hackechino\document.md"
OUTPUT_EXCEL_PATH = r"C:\Unified coding\Projects\Hackechino\output_multisheet_json.xlsx"
# Using Gemini 1.5 Flash now as requested (was 1.0 previously)
LLM_MODEL = 'gemini-2.5-flash' # Use 'gemini-1.5-flash' or other suitable model
# --- End Configuration ---

# Configure the generative AI client
try:
    genai.configure(api_key=API_KEY)
    # Ensure you are using the desired model version
    model = genai.GenerativeModel(LLM_MODEL)
    print(f"Using LLM Model: {LLM_MODEL}")
except Exception as e:
    print(f"Error configuring GenAI: {e}")
    exit()

# Read the input markdown/HTML file
try:
    with open(HTML_FILE_PATH, 'r', encoding='utf-8') as file:
        html_content = file.read()
except FileNotFoundError:
    print(f"Error: Input file not found at {HTML_FILE_PATH}")
    exit()
except Exception as e:
    print(f"Error reading file {HTML_FILE_PATH}: {e}")
    exit()

# --- Modified Prompt for JSON Output ---
prompt = f"""Analyze the following HTML content and extract the data for the specified tables.

Output the result ONLY as a valid JSON list. Each item in the list should be a JSON object representing one table.
Each table object must have two keys:
1.  "table_name": A string containing the descriptive name of the table (e.g., "Section: Employees M F", "Q.19", "Principle 6: Total Electricity Consumption"). Use the names provided in the requirements.
2.  "data": A list of lists, where each inner list represents a single row of the table's data.
    *   The FIRST inner list MUST be the header row.
    *   Subsequent inner lists are the data rows.
    *   Ensure all inner lists (rows) within a single table have the same number of elements (columns).
    *   Represent empty cells as empty strings ("") or null.

Example of a table object in the JSON list:
{{
  "table_name": "Section: Employees M F",
  "data": [
    ["S. No.", "Particulars", "Total (A)", "Male No. (B)", "Male % (B/A)", "Female No. (C)", "Female % (C/A)", "Others No. (D)", "Others % (D/A)"],
    ["1", "Permanent (E)", "74,705", "68,252", "91.4", "6,366", "8.5", "87", "0.1"],
    ["2", "Other than Permanent (F)", "3,347", "2,295", "68.6", "1,052", "31.4", "-", "-"],
    ["3", "Total Employees (E+ F)", "78,052", "70,547", "90.4", "7,418", "9.5", "87", "0.1"]
  ]
}}

If a required table cannot be extracted or data is missing (e.g., specific years not found), the "data" list for that table object should contain only the header row followed by a single row explaining the issue, like:
{{
  "table_name": "Principle 6: Total Fuel Consumption for 21-22",
  "data": [
      ["Status"],
      ["Data for FY 21-22 not found in source"]
  ]
}}
Or, if appropriate, return an empty list for "data": [].

Do NOT include any text before the opening bracket `[` of the JSON list or after the closing bracket `]`. Ensure the entire output is valid JSON.

Required tables:
- Section: Employees M F
- Workers M F
- Differently Abled M F (Both Employees & Workers)
- Q.19
- Q.20: Turnover Rate (for 21-22, 22-23, 23-24) Perm. Employees / Perm. Workers
- Principle 3 Q.5: Return to Work Rate (Employees, Workers) and Retention Rate (Employees, Workers)
- Principle 5 Q.43: Median Wage (Male / Female, KMP, Employees, Workers)
- Principle 6: Total Electricity Consumption (GJ) for 21-22, 22-23, 23-24
- Principle 6: Total Fuel Consumption for 21-22
- Principle 6: Total Volume of Water Consumed for 21-22, 22-23, all years
- Principle 6: Total Emissions (million ton CO2e) Scope-1 and Scope-2 for 21-22 and 22-23

md Content:
{html_content}
"""

# Generate content using the LLM
print("Generating content from LLM...")
try:
    response = model.generate_content(prompt)
    llm_output_text = response.text

    # --- Clean the LLM Output ---
    # Remove potential markdown fences (```json ... ```)
    match = re.search(r'```json\s*([\s\S]*?)\s*```', llm_output_text, re.DOTALL)
    if match:
        print("Found JSON within markdown fences, extracting...")
        cleaned_output_text = match.group(1).strip()
    else:
        # Assume the whole text is JSON, strip whitespace
        cleaned_output_text = llm_output_text.strip()

    # Ensure it starts with [ and ends with ] (basic validation)
    if not (cleaned_output_text.startswith('[') and cleaned_output_text.endswith(']')):
         print("Warning: LLM output doesn't seem to be a valid JSON list structure.")
         # Attempt to find the start/end if wrapped in other text (more aggressive)
         start_index = cleaned_output_text.find('[')
         end_index = cleaned_output_text.rfind(']')
         if start_index != -1 and end_index != -1 and start_index < end_index:
             print("Attempting to extract JSON list from detected start/end brackets.")
             cleaned_output_text = cleaned_output_text[start_index : end_index + 1]
         else:
             print("Could not reliably find JSON list brackets. Parsing might fail.")


    print("Cleaned LLM Output (first 500 chars):\n", cleaned_output_text[:500])
    print("-" * 20)

    # --- Parse the JSON Output ---
    try:
        tables_data = json.loads(cleaned_output_text)
        if not isinstance(tables_data, list):
            print("Error: Parsed JSON is not a list as expected.")
            exit()
    except json.JSONDecodeError as json_err:
        print(f"Fatal Error: Failed to decode JSON from LLM response: {json_err}")
        print("--- Problematic Text ---")
        print(cleaned_output_text)
        print("--- End Problematic Text ---")
        exit()

except Exception as e:
    print(f"Error generating content or initial cleaning from LLM: {e}")
    # print("Full response object:", response) # Uncomment for debugging if needed
    exit()


# --- Process the Parsed JSON and Write to Excel ---
print(f"Processing parsed JSON and writing to {OUTPUT_EXCEL_PATH}...")

try:
    with pd.ExcelWriter(OUTPUT_EXCEL_PATH, engine='openpyxl') as writer:
        sheet_counter = 1
        processed_sheet_names = set() # Keep track of names used

        for i, table_info in enumerate(tables_data):
            if not isinstance(table_info, dict):
                print(f"Skipping item {i+1}: Expected a dictionary (object), but got {type(table_info)}")
                continue

            table_name = table_info.get('table_name', f'Unnamed_Table_{i+1}')
            table_data = table_info.get('data', []) # Default to empty list if 'data' key is missing

            print(f"\nProcessing Table: '{table_name}'")

            if not isinstance(table_data, list):
                 print(f"  Skipping table '{table_name}': 'data' field is not a list.")
                 continue

            # Clean the table name for use as an Excel sheet name
            clean_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', table_name) # Replace invalid chars
            clean_sheet_name = clean_sheet_name[:31] # Limit length
            if not clean_sheet_name: # Handle empty name after cleaning
                clean_sheet_name = f'Sheet_{sheet_counter}'

            # Ensure sheet name uniqueness
            original_clean_name = clean_sheet_name
            suffix = 1
            while clean_sheet_name in processed_sheet_names:
                suffix_str = f"_{suffix}"
                trunc_len = 31 - len(suffix_str)
                clean_sheet_name = f"{original_clean_name[:trunc_len]}{suffix_str}"
                suffix += 1
            processed_sheet_names.add(clean_sheet_name)


            try:
                if not table_data: # Handle explicitly empty data list
                    print(f"  Table '{table_name}' has empty 'data' list. Creating empty sheet.")
                    df = pd.DataFrame() # Create an empty DataFrame
                elif len(table_data) == 1: # Only a header row or maybe a status message
                     print(f"  Table '{table_name}' has only one row. Writing as is.")
                     header = table_data[0] if isinstance(table_data[0], list) else [str(table_data[0])] # Handle non-list single item
                     df = pd.DataFrame([], columns=header) # DataFrame with only header
                     # If you want the single row as data instead:
                     # df = pd.DataFrame([table_data[0]]) # Treat the single row as data without header
                else:
                    # Assume first row is header, rest is data
                    header = table_data[0]
                    data_rows = table_data[1:]

                    # Basic validation: check if header is a list
                    if not isinstance(header, list):
                         print(f"  Skipping table '{table_name}': Header (first item in 'data') is not a list.")
                         continue

                    # Basic validation: Check if all data rows are lists and have same length as header
                    num_cols = len(header)
                    valid_rows = []
                    for r_idx, row in enumerate(data_rows):
                        if isinstance(row, list) and len(row) == num_cols:
                            valid_rows.append(row)
                        else:
                            print(f"  Warning for table '{table_name}': Row {r_idx+1} is invalid (not a list or incorrect column count: expected {num_cols}, got {len(row) if isinstance(row,list) else 'Not a list'}). Skipping row.")
                            # Optionally append placeholders or handle differently
                            # valid_rows.append(['ERROR'] * num_cols)

                    if not valid_rows and data_rows: # If all data rows were invalid
                         print(f"  No valid data rows found for table '{table_name}' after header. Creating sheet with only header.")
                         df = pd.DataFrame([], columns=header)
                    elif not valid_rows and not data_rows: # No data rows to begin with
                         print(f"  No data rows found for table '{table_name}'. Creating sheet with only header.")
                         df = pd.DataFrame([], columns=header)
                    else: # We have valid rows
                        df = pd.DataFrame(valid_rows, columns=header)


                # Write DataFrame to the Excel sheet
                df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
                print(f"  Successfully wrote data to sheet: '{clean_sheet_name}'")
                sheet_counter += 1

            except Exception as df_err:
                print(f"  Error creating DataFrame or writing sheet for '{table_name}': {df_err}")
                print(f"  Problematic Data Structure:\n{table_data}\n---")
                # Optionally write raw data to a sheet for inspection
                try:
                     error_sheet_name = f"ERROR_{sheet_counter}"
                     # Ensure error sheet name is unique too
                     original_error_name = error_sheet_name
                     suffix = 1
                     while error_sheet_name in processed_sheet_names:
                          suffix_str = f"_{suffix}"
                          trunc_len = 31 - len(suffix_str)
                          error_sheet_name = f"{original_error_name[:trunc_len]}{suffix_str}"
                          suffix += 1
                     processed_sheet_names.add(error_sheet_name)

                     pd.DataFrame(table_data).to_excel(writer, sheet_name=error_sheet_name, index=False, header=False)
                     print(f"  Wrote raw problematic data to sheet: '{error_sheet_name}'")
                     sheet_counter += 1
                except Exception as raw_write_err:
                     print(f"  Could not write raw data for table '{table_name}' to error sheet: {raw_write_err}")


    print(f"\nSuccessfully finished processing. Excel file saved to: {OUTPUT_EXCEL_PATH}")

except Exception as e:
    print(f"An unexpected error occurred during Excel file writing or processing: {e}")
