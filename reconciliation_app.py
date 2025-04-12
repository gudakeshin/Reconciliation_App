import os
import pandas as pd
import json
import io
import datetime
import base64
import requests # Still needed for potential other HTTP requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS # Import CORS
from openai import OpenAI # Import OpenAI library
from dotenv import load_dotenv # Import dotenv to load environment variables

# --- Configuration ---
load_dotenv() # Load environment variables from .env file
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# --- Flask App Initialization ---
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024 # Limit uploads to 30 MB total
CORS(app) # Enable CORS for all routes

# --- OpenAI Client Initialization ---
# IMPORTANT: Set your OpenAI API key as an environment variable named OPENAI_API_KEY
# Or, replace the placeholder below, but environment variables are recommended for security.
try:
    client = OpenAI(
        api_key=os.environ.get("OPENAI_API_KEY", "YOUR_OPENAI_API_KEY") # Use environment variable or placeholder
    )
    # Add a check if the placeholder is still being used
    if client.api_key == "YOUR_OPENAI_API_KEY":
        print("\n*********************************************************************")
        print("WARNING: Using placeholder OpenAI API key.")
        print("Please set the OPENAI_API_KEY environment variable or replace")
        print("'YOUR_OPENAI_API_KEY' in the code with your actual OpenAI API key.")
        print("*********************************************************************\n")
except Exception as e:
    print(f"Error initializing OpenAI client: {e}")
    client = None # Set client to None if initialization fails


# --- Helper Functions (normalize_value, create_composite_key, highlight_discrepancies) ---
# These functions remain the same as in the original code.

def normalize_value(value):
    """Converts value to string, lowercases, and trims whitespace."""
    if pd.isna(value):
        return "" # Represent NaN/None as empty string for consistent comparison
    return str(value).strip().lower()

def create_composite_key(row, key_columns):
    """Creates a composite key from specified columns for a row."""
    # Use normalized values for key creation
    key_parts = [normalize_value(row.get(col, '')) for col in key_columns]
    return "_||_".join(key_parts) # Use a unique separator

def highlight_discrepancies(row, df1_row, df2_row, all_headers1, all_headers2, key_headers1):
    """Identifies differing values between two matched rows, excluding key columns."""
    discrepancies = {}
    non_key_headers1 = [h for h in all_headers1 if h not in key_headers1]
    non_key_headers2 = [h for h in all_headers2 if h not in key_headers1] # Compare based on File 1 non-key headers

    for col1 in non_key_headers1:
        col2 = col1 # Assumes same name, needs mapping for different names
        val1 = normalize_value(df1_row.get(col1, ''))
        val2 = normalize_value(df2_row.get(col2, ''))

        if val1 != val2:
            discrepancies[col1] = {'file1': str(df1_row.get(col1, '')), 'file2': str(df2_row.get(col2, ''))}

    return discrepancies if discrepancies else None

# --- Flask Routes ---

@app.route('/')
def index():
    # Index route remains the same
    return jsonify({
        "message": "Excel Reconciliation API with OpenAI Integration",
        "endpoints": {
            "/reconcile": { "method": "POST", "description": "Reconcile two Excel files" },
            "/analyze": { "method": "POST", "description": "Analyze reconciliation summary using OpenAI" },
            "/recommend-columns": { "method": "POST", "description": "Recommend reconciliation columns using OpenAI" }
        }
    })

@app.route('/reconcile', methods=['POST'])
def reconcile_files():
    """
    API endpoint to receive two Excel files, key headers, and mapping,
    perform reconciliation, and return the report.
    (This core reconciliation logic remains unchanged from the original)
    """
    start_time = datetime.datetime.now()
    print(f"[{start_time.strftime('%Y-%m-%d %H:%M:%S')}] Received reconciliation request.")

    # --- 1. Input Validation (Same as before) ---
    if 'file1' not in request.files or 'file2' not in request.files:
        return jsonify({"error": "Both file1 and file2 are required."}), 400
    # ... (rest of input validation is identical to the original code) ...
    file1 = request.files['file1']
    file2 = request.files['file2']
    file1_name = file1.filename
    file2_name = file2.filename

    if not (file1_name and (file1_name.endswith('.xlsx') or file1_name.endswith('.xls'))):
        return jsonify({"error": "File 1 must be a .xlsx or .xls file."}), 400
    if not (file2_name and (file2_name.endswith('.xlsx') or file2_name.endswith('.xls'))):
        return jsonify({"error": "File 2 must be a .xlsx or .xls file."}), 400

    try:
        file1_key_headers = json.loads(request.form.get('file1_key_headers', '[]'))
        file2_key_headers = json.loads(request.form.get('file2_key_headers', '[]'))
        column_mapping = json.loads(request.form.get('column_mapping', '{}'))

        if not (file1_key_headers and len(file1_key_headers) == 3):
             return jsonify({"error": "Exactly 3 key headers must be provided for File 1."}), 400
        if not (file2_key_headers and len(file2_key_headers) == 3):
             return jsonify({"error": "Exactly 3 key headers must be provided for File 2."}), 400
        if not (column_mapping and len(column_mapping) == 3):
            return jsonify({"error": "Exactly 3 column mappings are required."}), 400
        if set(file1_key_headers) != set(column_mapping.keys()):
             return jsonify({"error": "Mapping keys must exactly match File 1 selected key headers."}), 400
        if set(file2_key_headers) != set(column_mapping.values()):
             return jsonify({"error": "Mapping values must exactly match File 2 selected key headers."}), 400

        print(f"File 1: {file1_name}, Keys: {file1_key_headers}")
        print(f"File 2: {file2_name}, Keys: {file2_key_headers}")
        print(f"Mapping: {column_mapping}")

    except json.JSONDecodeError:
        return jsonify({"error": "Invalid JSON format for headers or mapping data."}), 400
    except Exception as e:
        return jsonify({"error": f"Error processing input parameters: {e}"}), 400


    # --- 2. File Reading and Preparation (Same as before) ---
    try:
        print("Reading Excel files into pandas DataFrames...")
        df1 = pd.read_excel(file1, engine='openpyxl' if file1_name.endswith('.xlsx') else None)
        df2 = pd.read_excel(file2, engine='openpyxl' if file2_name.endswith('.xlsx') else None)
        all_headers1 = df1.columns.tolist()
        all_headers2 = df2.columns.tolist()
        mapped_key_headers2 = [column_mapping[h1] for h1 in file1_key_headers]

        print(f"Creating composite keys for File 1 using: {file1_key_headers}")
        df1['_composite_key'] = df1.apply(lambda row: create_composite_key(row, file1_key_headers), axis=1)
        print(f"Creating composite keys for File 2 using mapped headers: {mapped_key_headers2}")
        df2['_composite_key'] = df2.apply(lambda row: create_composite_key(row, mapped_key_headers2), axis=1)

        # Optional: Check for duplicate keys (same as before)
        if df1['_composite_key'].duplicated().any(): print("Warning: Duplicate keys found in File 1.")
        if df2['_composite_key'].duplicated().any(): print("Warning: Duplicate keys found in File 2.")

    except Exception as e:
        print(f"Error reading or processing Excel files: {e}")
        return jsonify({"error": f"Failed to read or process Excel files: {e}"}), 500

    # --- 3. Reconciliation using Merge (Same as before) ---
    try:
        print("Performing merge operation...")
        merged_df = pd.merge(
            df1.add_prefix('f1_'),
            df2.add_prefix('f2_'),
            left_on='f1__composite_key',
            right_on='f2__composite_key',
            how='outer',
            indicator=True
        )
        print(f"Merge complete. Merged shape: {merged_df.shape}")

        # --- 4. Categorize Records (Same as before) ---
        print("Categorizing records...")
        unique_file1_df = merged_df[merged_df['_merge'] == 'left_only'].copy()
        unique_file2_df = merged_df[merged_df['_merge'] == 'right_only'].copy()
        matched_df = merged_df[merged_df['_merge'] == 'both'].copy()

        # Clean up unique DFs (same as before)
        unique_file1_df = unique_file1_df[[f'f1_{col}' for col in all_headers1]].rename(columns=lambda x: x[3:])
        unique_file2_df = unique_file2_df[[f'f2_{col}' for col in all_headers2]].rename(columns=lambda x: x[3:])

        # --- 5. Identify Discrepancies in Matched Records (Same as before) ---
        print("Identifying discrepancies in matched records...")
        matched_records_identical = []
        matched_records_discrepancies = []
        discrepancy_details = []

        for index, row in matched_df.iterrows():
            original_row1 = {col: row[f'f1_{col}'] for col in all_headers1}
            original_row2 = {col: row[f'f2_{col}'] for col in all_headers2}
            is_identical = True
            discrepancy_info = {}

            for col1 in all_headers1:
                col2 = column_mapping.get(col1, col1) # Use mapping for keys, assume same name otherwise
                val1_norm = normalize_value(original_row1.get(col1))
                val2_norm = normalize_value(original_row2.get(col2))

                if val1_norm != val2_norm:
                    is_identical = False
                    discrepancy_info[col1] = {
                        'file1_value': str(original_row1.get(col1, '')),
                        'file2_value': str(original_row2.get(col2, ''))
                    }

            # Combine data (same as before)
            combined_row_data = {}
            for col in all_headers1: combined_row_data[f'{col} (File 1)'] = original_row1.get(col, '')
            for col in all_headers2:
                 mapped_col1 = next((k for k, v in column_mapping.items() if v == col), None)
                 if mapped_col1 is None or mapped_col1 not in all_headers1:
                    combined_row_data[f'{col} (File 2)'] = original_row2.get(col, '')

            if is_identical:
                matched_records_identical.append(combined_row_data)
            else:
                combined_row_data['_DISCREPANCIES'] = json.dumps(discrepancy_info)
                matched_records_discrepancies.append(combined_row_data)
                # Prep details for report (same as before)
                detail_row = {}
                for col1 in all_headers1:
                    col2 = column_mapping.get(col1, col1)
                    detail_row[f'{col1} (File 1)'] = original_row1.get(col1, '')
                    if col1 in discrepancy_info:
                         detail_row[f'{col2} (File 2)'] = f"DIFFERS: {discrepancy_info[col1]['file2_value']}"
                    else:
                         detail_row[f'{col2} (File 2)'] = original_row2.get(col2, '')
                discrepancy_details.append(detail_row)

        matched_identical_df = pd.DataFrame(matched_records_identical)
        matched_discrepancies_detail_df = pd.DataFrame(discrepancy_details)

    except Exception as e:
        print(f"Error during reconciliation or categorization: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Error during reconciliation: {e}"}), 500

    # --- 6. Generate Summary (Same as before) ---
    summary = {
        "file1_name": file1_name,
        "file2_name": file2_name,
        "total_file1": df1.shape[0],
        "total_file2": df2.shape[0],
        "matched_identical": len(matched_records_identical),
        "matched_discrepancies": len(matched_records_discrepancies),
        "unique_file1": unique_file1_df.shape[0],
        "unique_file2": unique_file2_df.shape[0],
        "timestamp": start_time.isoformat(),
        "file1_key_headers": file1_key_headers,
        "file2_key_headers": mapped_key_headers2,
        "mapping": column_mapping,
    }
    print("Reconciliation Summary:", json.dumps(summary, indent=2))

    # --- 7. Generate Excel Report (Same as before) ---
    try:
        print("Generating Excel report...")
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
            # Summary Sheet (same as before)
            summary_df = pd.DataFrame([
                {"Metric": "File 1 Name", "Value": summary["file1_name"]},
                # ... (rest of summary metrics) ...
                 {"Metric": "File 2 Key Columns Used (Mapped)", "Value": ", ".join(summary["file2_key_headers"])},
            ])
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

            # Discrepancies Sheet (same as before, including highlighting)
            if not matched_discrepancies_detail_df.empty:
                 matched_discrepancies_detail_df.to_excel(writer, sheet_name='Matched (Discrepancies)', index=False)
                 workbook = writer.book
                 worksheet = writer.sheets['Matched (Discrepancies)']
                 highlight_format = workbook.add_format({'bg_color': '#FFDDC1', 'bold': True})
                 for row_num, row_data in enumerate(matched_discrepancies_detail_df.values):
                     for col_num, cell_value in enumerate(row_data):
                         if isinstance(cell_value, str) and cell_value.startswith("DIFFERS:"):
                             worksheet.write(row_num + 1, col_num, cell_value, highlight_format)
            else:
                 pd.DataFrame([{"Status": "No matched records with discrepancies found."}]).to_excel(writer, sheet_name='Matched (Discrepancies)', index=False)

            # Other Sheets (Identical, Unique File 1, Unique File 2 - same as before)
            if not matched_identical_df.empty:
                matched_identical_df.to_excel(writer, sheet_name='Matched (Identical)', index=False)
            else:
                pd.DataFrame([{"Status": "No identical matched records found."}]).to_excel(writer, sheet_name='Matched (Identical)', index=False)
            if not unique_file1_df.empty:
                unique_file1_df.to_excel(writer, sheet_name='File 1 Unique Records', index=False)
            else:
                 pd.DataFrame([{"Status": f"No records found only in {file1_name}."}]).to_excel(writer, sheet_name='File 1 Unique Records', index=False)
            if not unique_file2_df.empty:
                unique_file2_df.to_excel(writer, sheet_name='File 2 Unique Records', index=False)
            else:
                 pd.DataFrame([{"Status": f"No records found only in {file2_name}."}]).to_excel(writer, sheet_name='File 2 Unique Records', index=False)

        output_buffer.seek(0)
        print("Excel report generated successfully.")

        # --- 8. Return File Response (Same as before) ---
        summary_json = json.dumps(summary)
        summary_base64 = base64.b64encode(summary_json.encode('utf-8')).decode('utf-8')
        response = send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'Reconciliation_{os.path.splitext(file1_name)[0]}_vs_{os.path.splitext(file2_name)[0]}.xlsx'
        )
        response.headers['X-Reconciliation-Summary'] = summary_base64
        response.headers['Access-Control-Expose-Headers'] = 'X-Reconciliation-Summary'
        end_time = datetime.datetime.now()
        print(f"[{end_time.strftime('%Y-%m-%d %H:%M:%S')}] Reconciliation complete. Duration: {end_time - start_time}")
        return response

    except Exception as e:
        print(f"Error generating Excel report: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({"error": f"Failed to generate Excel report: {e}"}), 500


@app.route('/analyze', methods=['POST'])
def analyze_reconciliation():
    """
    Analyzes the reconciliation results using OpenAI API
    and provides actionable insights.
    """
    # Check if OpenAI client is initialized
    if client is None or client.api_key == "YOUR_OPENAI_API_KEY":
        return jsonify({
            "error": "OpenAI API key not configured.",
            "details": "Please set the OPENAI_API_KEY environment variable or update the placeholder in the code."
        }), 503 # Service Unavailable

    try:
        # Get the reconciliation summary from the request
        summary = request.json.get('summary')
        if not summary:
            return jsonify({"error": "No summary data provided"}), 400

        # --- Prepare the prompt for OpenAI ---
        # Use a system message to set the context for the AI
        system_message = "You are an expert financial analyst providing actionable insights based on reconciliation data."
        # User message containing the specific data and request
        user_prompt = f"""
        Analyze this reconciliation summary and provide actionable insights:

        File 1: {summary['file1_name']} ({summary['total_file1']} records)
        File 2: {summary['file2_name']} ({summary['total_file2']} records)

        Matched Records (Identical): {summary['matched_identical']}
        Matched Records (with Discrepancies): {summary['matched_discrepancies']}
        Records Unique to File 1: {summary['unique_file1']}
        Records Unique to File 2: {summary['unique_file2']}

        Please provide:
        1. A brief analysis of the reconciliation results (potential issues, significance of numbers).
        2. Specific, actionable steps to investigate and resolve:
           - Records with discrepancies
           - Records unique to File 1
           - Records unique to File 2
        3. Recommendations for process improvements to prevent similar reconciliation issues in the future.

        Format your response clearly, using bullet points or numbered lists for actions and recommendations.
        """

        print("Sending analysis request to OpenAI...")
        # --- Call OpenAI API ---
        try:
            completion = client.chat.completions.create(
                model="gpt-3.5-turbo", # Or use "gpt-4" or other suitable models
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.5, # Adjust temperature for creativity vs consistency
                max_tokens=500 # Set a reasonable limit for the response length
            )

            # Extract the response content
            analysis = completion.choices[0].message.content.strip()
            print("Received analysis from OpenAI.")

            # Return the analysis from OpenAI
            return jsonify({
                "analysis": analysis,
                "timestamp": datetime.datetime.now().isoformat()
            })

        # --- Handle potential OpenAI API errors ---
        except Exception as e: # Catch OpenAI specific errors if needed, e.g., openai.APIError
            error_message = f"Error communicating with OpenAI API: {e}"
            print(error_message)
            # Check for specific common errors like authentication
            if "AuthenticationError" in str(e) or "Incorrect API key" in str(e):
                 details = "Please check if your OpenAI API key is correct and has sufficient credits."
            else:
                 details = str(e)
            return jsonify({
                "error": "Failed to get analysis from OpenAI.",
                "details": details
            }), 500

    except Exception as e:
        # Handle other potential errors in the endpoint logic
        print(f"Error in analyze_reconciliation endpoint: {str(e)}")
        return jsonify({
            "error": "Failed to process analysis request",
            "details": str(e)
        }), 500


@app.route('/recommend-columns', methods=['POST'])
def recommend_columns():
    """
    Analyzes the Excel files using OpenAI API and recommends columns for reconciliation.
    """
    # Check if OpenAI client is initialized
    if client is None or client.api_key == "YOUR_OPENAI_API_KEY":
        return jsonify({
            "error": "OpenAI API key not configured.",
            "details": "Please set the OPENAI_API_KEY environment variable or update the placeholder in the code."
        }), 503 # Service Unavailable

    try:
        # Get the files from the request
        if 'file1' not in request.files or 'file2' not in request.files:
            return jsonify({"error": "Both files are required"}), 400

        file1 = request.files['file1']
        file2 = request.files['file2']

        # --- Read only headers to avoid loading large files ---
        try:
             df1 = pd.read_excel(file1, nrows=1) # Read only the header row
             df2 = pd.read_excel(file2, nrows=1)
             file1_columns = df1.columns.tolist()
             file2_columns = df2.columns.tolist()
        except Exception as e:
            print(f"Error reading file headers: {e}")
            return jsonify({"error": f"Failed to read headers from files: {e}"}), 400

        # --- Prepare the prompt for OpenAI ---
        system_message = """
        You are an expert assistant specializing in data reconciliation between Excel files.
        Your task is to recommend the **best 3 columns** from each file to use as keys for matching records accurately.
        Focus on columns that are likely to uniquely identify a transaction or entity across both files.
        Consider common identifiers like Transaction IDs, Order Numbers, Customer IDs, Product Codes, Dates, etc.
        Avoid columns with highly variable data (like descriptions, amounts, statuses) unless they are part of a composite key.
        The output MUST be a valid JSON object.
        """

        user_prompt = f"""
        Analyze the column headers from two Excel files and recommend the best 3 columns from each to use as reconciliation keys.

        File 1 Columns: {json.dumps(file1_columns)}
        File 2 Columns: {json.dumps(file2_columns)}

        Based on these column names, provide your recommendations.

        Return your response ONLY as a JSON object with the following structure:
        {{
          "file1_recommendations": ["ColumnA", "ColumnB", "ColumnC"], // List of exactly 3 strings (column names from File 1)
          "file2_recommendations": ["ColumnX", "ColumnY", "ColumnZ"], // List of exactly 3 strings (column names from File 2)
          "explanation": "Brief reasoning for choosing these columns (e.g., combination likely unique).",
          "considerations": "Potential issues (e.g., date format differences, case sensitivity, leading/trailing spaces)."
        }}

        Ensure the recommended column names exactly match the names provided in the input lists.
        Do not include any text before or after the JSON object.
        """

        print("Sending column recommendation request to OpenAI...")
        # --- Call OpenAI API ---
        try:
            completion = client.chat.completions.create(
                model="gpt-3.5-turbo", # Or "gpt-4"
                messages=[
                    {"role": "system", "content": system_message},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.2, # Lower temperature for more deterministic JSON output
                max_tokens=300,
                response_format={ "type": "json_object" } # Request JSON output if using compatible models
            )

            # Extract the response content (which should be JSON)
            recommendation_json_str = completion.choices[0].message.content.strip()
            print("Received recommendations from OpenAI.")

            # --- Parse the JSON response ---
            try:
                recommendations = json.loads(recommendation_json_str)
                # Basic validation of the returned structure
                if not all(k in recommendations for k in ["file1_recommendations", "file2_recommendations", "explanation", "considerations"]) or \
                   len(recommendations["file1_recommendations"]) != 3 or \
                   len(recommendations["file2_recommendations"]) != 3:
                    raise ValueError("OpenAI response did not match the required JSON structure or list lengths.")

                # Return the parsed recommendations along with all columns
                return jsonify({
                    "recommendations": recommendations,
                     "all_columns": { # Also return all columns for the frontend dropdowns
                        "file1": file1_columns,
                        "file2": file2_columns
                     },
                    "timestamp": datetime.datetime.now().isoformat()
                })

            except (json.JSONDecodeError, ValueError) as e:
                print(f"Error parsing OpenAI JSON response: {e}")
                print(f"Received string: {recommendation_json_str}")
                return jsonify({
                    "error": "Invalid response format received from OpenAI.",
                    "details": f"Could not parse recommendations. Raw response: {recommendation_json_str}",
                    "all_columns": { "file1": file1_columns, "file2": file2_columns } # Still return all columns
                }), 500

        # --- Handle potential OpenAI API errors ---
        except Exception as e:
            error_message = f"Error communicating with OpenAI API: {e}"
            print(error_message)
            if "AuthenticationError" in str(e) or "Incorrect API key" in str(e):
                 details = "Please check if your OpenAI API key is correct and has sufficient credits."
            else:
                 details = str(e)
            return jsonify({
                "error": "Failed to get recommendations from OpenAI.",
                "details": details,
                 "all_columns": { "file1": file1_columns, "file2": file2_columns } # Still return all columns
            }), 500

    except Exception as e:
        # Handle other potential errors in the endpoint logic
        print(f"Error in recommend_columns endpoint: {str(e)}")
        return jsonify({
            "error": "Failed to process recommendation request",
            "details": str(e)
        }), 500


# --- Main Execution ---
if __name__ == '__main__':
    print("Starting Flask server for Excel Reconciliation App (with OpenAI)...")
    print("Backend accessible at http://127.0.0.1:5000")
    # Use debug=False for production environments
    # Consider using a proper WSGI server like Gunicorn or Waitress for production
    app.run(debug=True) # debug=True enables auto-reloading and detailed errors
