from langchain_community.llms import Ollama
import streamlit as st
import os
from dotenv import load_dotenv
import json
import re
import tempfile
from openpyxl import load_workbook
from utils import  converter

load_dotenv()
os.environ["LANGCHAIN_TRACING_V2"] = "true"
os.environ["LANGCHAIN_API_KEY"] = os.getenv("LANGCHAIN_API_KEY")


llm = Ollama(model="mistral", base_url="http://127.0.0.1:11434")
def clean_and_parse_json(response):
    """Cleans and parses JSON safely, handling errors."""
    try:
        if not response or not isinstance(response, str):
            raise ValueError("Response is empty or not a string.")

 
        response = re.sub(r"//.*", "", response)


        response = response.replace("'", '"')  


        match = re.search(r"\{.*\}", response, re.DOTALL)
        if match:
            return json.loads(match.group(0))  

    except json.JSONDecodeError as e:
        print(f"JSON parsing error: {e}")
    except ValueError as e:
        print(f"Value error: {e}")

    return None 

def update_excel_mapping(data_dict, user_query):
    prompt = f"""
    You are an AI assistant that updates an Excel mapping dictionary based on user instructions.
    The dictionary maps Excel cell names to their positions in an Excel sheet.

    Given the following dictionary:

    {json.dumps(data_dict, indent=4)}

    User instruction: {user_query}

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

    print(response)

    return response


def update_excel_from_json(json_data, excel_file, output_file, sheet_name):

    wb = load_workbook(excel_file, keep_vba=True)


    if sheet_name not in wb.sheetnames:
        print(f"Sheet '{sheet_name}' does not exist in the Excel file.")
        return

    sheet = wb[sheet_name]

 
    for key, cell_refs in json_data.items():
 
        if isinstance(cell_refs, str):
   
            cell_list = [cell.strip() for cell in cell_refs.split(',')]

            for cell_ref in cell_list:
                try:
                    sheet[cell_ref] = key 
                except ValueError:
                    print(f"Invalid cell reference: {cell_ref}")

        else:
        
            if isinstance(cell_refs, str):
                cell_list = [cell.strip() for cell in cell_refs.split(',')]
                for cell_ref in cell_list:
                    try:
                        sheet[cell_ref] = key   
                    except ValueError:
                        print(f"Invalid cell reference: {cell_ref}")


    wb.save(output_file)

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

           
            with open(output_file_path, "rb") as f:
                st.download_button("Download Modified Excel", f, file_name="modified.xlsm", mime="application/vnd.ms-excel")
        else:
            st.warning("Please enter a command to modify the file.")
