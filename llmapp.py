from langchain_community.llms import Ollama
import streamlit as st
import os
from dotenv import load_dotenv
import json
import re
import tempfile
from openpyxl import load_workbook
from utils import  converter
# Load environment variables
load_dotenv()
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_API_KEY"] = os.getenv("LANGCHAIN_API_KEY")

# Initialize Ollama model (Mistral)
llm = Ollama(model="mistral", base_url="http://127.0.0.1:11434")
def clean_and_parse_json(response):
    """Cleans and parses JSON safely, handling errors."""
    try:
        if not response or not isinstance(response, str):
            raise ValueError("Response is empty or not a string.")

        # Remove JavaScript-style comments (// ...)
        response = re.sub(r"//.*", "", response)

        # Ensure all keys and values use double quotes
        response = response.replace("'", '"')  # Convert single to double quotes

        # Extract only valid JSON (between curly braces)
        match = re.search(r"\{.*\}", response, re.DOTALL)
        if match:
            return json.loads(match.group(0))  # Convert to Python dict

    except json.JSONDecodeError as e:
        print(f"JSON parsing error: {e}")
    except ValueError as e:
        print(f"Value error: {e}")

    return None 
# Function to interpret user command using LangChain's Ollama wrapper
def update_excel_mapping(data_dict, user_query):
    prompt = f"""
    You are an AI assistant that updates an Excel mapping dictionary based on user instructions.
    The dictionary maps Excel cell names to their positions in an Excel sheet.

    Given the following dictionary:

    {json.dumps(data_dict, indent=4)}

    User instruction: "{user_query}"

    Modify the dictionary according to the user instruction while strictly preserving the format:
    - If a value is enclosed in single quotes (' '), it must remain in single quotes.
    - If a value is not enclosed in quotes, it must remain without quotes.
    - Do not change the structure, spacing, or ordering of the dictionary.

    Output the **complete updated dictionary** in valid JSON format.
    Include both modified and unchanged key-value pairs.
    Do not add explanations, comments, or formatting changes—just return valid JSON.
    """
    response = llm.invoke(prompt) 
    print(response)
    st.write("Raw Response from Ollama:", response)
    response = clean_and_parse_json(response)
    # response = json.loads(response)
    print(response)
    # # Extract the text response
    # response_text = response["message"]["content"].strip()

    # Extract only the JSON part
    # match = re.search(r"\{.*?\}", response, re.DOTALL)
    # if match:
    #     return json.loads(match.group(0))  # Return parsed JSON
    # return {} 
    return response

# Function to update Excel file
def update_excel(file_path, parsed_data):
    try:
        # Ensure extracted data is valid JSON
        data = json.loads(parsed_data)

        sheet_name = "Model Inputs"  # Fixed sheet name
        cell_name, new_value = data.get("cell_name"), data.get("new_value")

        if not cell_name or not new_value:
            return "❌ Error: Invalid JSON format from AI."

        wb = load_workbook(file_path, keep_vba=True)
        if sheet_name not in wb.sheetnames:
            return f"❌ Sheet '{sheet_name}' not found."

        sheet = wb[sheet_name]
        found = False

        # Normalize cell name for case-insensitive matching
        cell_name_normalized = str(cell_name).strip().casefold()

        # Find the cell with the name (including merged cells)
        for row in sheet.iter_rows(values_only=False):
            for cell in row:
                if cell.value and str(cell.value).strip().casefold() == cell_name_normalized:
                    # If the cell is in a merged range, find the top-left cell
                    for merged_range in sheet.merged_cells.ranges:
                        if cell.coordinate in merged_range:
                            cell = sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
                            break
                    
                    # Get the first non-merged cell in the same row
                    col = cell.column + 1
                    while sheet.cell(row=cell.row, column=col).coordinate in sheet.merged_cells:
                        col += 1
                    
                    # Assign new value
                    sheet.cell(row=cell.row, column=col).value = new_value
                    found = True
                    break

            if found:
                break

        if not found:
            return f"❌ Cell name '{cell_name}' not found in the sheet."

        wb.save(file_path)
        return f"✅ Updated value next to '{cell_name}' to '{new_value}'"

    except json.JSONDecodeError as e:
        return f"❌ JSON Error: {str(e)}"

    except Exception as e:
        return f"❌ Error: {str(e)}"
def update_excel_from_json(json_data, excel_file, output_file, sheet_name):
    # Load the Excel file and keep the VBA macros if any
    wb = load_workbook(excel_file, keep_vba=True)

    # Access the specific sheet you want to modify
    if sheet_name not in wb.sheetnames:
        print(f"Sheet '{sheet_name}' does not exist in the Excel file.")
        return

    sheet = wb[sheet_name]

    # Step through each entry in the JSON data and update corresponding cells
    for key, cell_refs in json_data.items():
        # If the value is a string (cell reference like 'D12' or 'G12, G19, G25')
        if isinstance(cell_refs, str):
            # Split the cell references in case multiple references are provided
            cell_list = [cell.strip() for cell in cell_refs.split(',')]

            for cell_ref in cell_list:
                try:
                    sheet[cell_ref] = key  # Write the key's value to the cell
                except ValueError:
                    print(f"Invalid cell reference: {cell_ref}")

        else:
            # If the value is a number, retrieve corresponding cell references
            if isinstance(cell_refs, str):
                cell_list = [cell.strip() for cell in cell_refs.split(',')]
                for cell_ref in cell_list:
                    try:
                        sheet[cell_ref] = key  # Write the key's value to the cell
                    except ValueError:
                        print(f"Invalid cell reference: {cell_ref}")


    wb.save(output_file)
# Streamlit UI
st.title("Excel Modifier App (LangChain + Mistral) 📝")
st.write("Upload a `.xlsm` file and modify it using natural language.")

uploaded_file = st.file_uploader("Upload your .xlsm file", type=["xlsm"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as temp_file:
        temp_file.write(uploaded_file.read())
        file_path = temp_file.name

    user_command = st.text_input("Enter your modification command (e.g., 'Update Male Population to 5000')")

    if st.button("Modify Excel"):
        if user_command:
            wb = load_workbook(file_path, keep_vba=True) 
            sheet = wb['Model Inputs']
            json_file = converter(sheet)
            sheet_name = "Model Inputs"
            print(json_file)
            data_dict_fixed = {str(k): v for k, v in json_file.items()}
            json_updated = update_excel_mapping(data_dict_fixed, user_command)
            output_file_path = file_path.replace(".xlsm", "_modified.xlsm") 
            result = update_excel_from_json(json_file, file_path,output_file_path , sheet_name)
            st.success(result)

            # Provide a download link
            with open(output_file_path, "rb") as f:
                st.download_button("Download Modified Excel", f, file_name="modified.xlsm", mime="application/vnd.ms-excel")
        else:
            st.warning("Please enter a command to modify the file.")
