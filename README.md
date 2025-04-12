# Excel Reconciliation App with OpenAI Integration

This application provides a web interface to reconcile two Excel files based on selected key columns and utilizes OpenAI for analyzing results and recommending reconciliation columns.

## Features

* Upload two Excel files (.xlsx or .xls).
* Manually select 3 key columns from each file for reconciliation.
* Map the selected key columns between the two files.
* (Optional) Get AI-powered recommendations for key columns using OpenAI.
* (Optional) Get AI-powered analysis of the reconciliation results using OpenAI.
* Perform reconciliation based on the composite keys derived from selected columns.
* Categorizes records into:
    * Matched (Identical)
    * Matched (with Discrepancies)
    * Unique to File 1
    * Unique to File 2
* Download a detailed reconciliation report in Excel format, highlighting discrepancies.

## Setup

1.  **Clone or Download:** Get the application files (`reconciliation_app.py`, `reconciler.html`, etc.).
2.  **Create Environment (Recommended):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```
3.  **Install Dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **Configure OpenAI API Key:**
    * Create a file named `.env` in the `Reco_App` directory.
    * Add your OpenAI API key to the `.env` file like this:
        ```
        OPENAI_API_KEY='your_actual_openai_api_key_here'
        ```
    * **Alternatively (Not Recommended for Security):** You can replace the placeholder `"YOUR_OPENAI_API_KEY"` directly in the `reconciliation_app.py` file, but using environment variables is safer.
    * Ensure your OpenAI account has sufficient credits/quota.

## Running the Application

1.  **Start the Backend Server:**
    ```bash
    python reconciliation_app.py
    ```
    The server will typically start on `http://127.0.0.1:5000`. Check the terminal output for the exact address.

2.  **Open the Frontend:**
    * Open the `reconciler.html` file directly in your web browser (e.g., by double-clicking it or using `File -> Open`).

## How to Use

1.  Open `reconciler.html` in your browser.
2.  Use the "Browse File" buttons or drag-and-drop areas to upload your two Excel files.
3.  Wait for the headers to load.
4.  **Column Selection:**
    * **AI Recommendation (Optional):** If both files are uploaded, an AI recommendation section may appear. Review the recommendations and click "Accept Recommendations" to use them or "Choose My Own" to select manually.
    * **Manual Selection:** Click on exactly 3 column headers from each file's list. These columns will be used to uniquely identify matching rows.
5.  **Column Mapping:** If 3 columns are selected for both files, a mapping section will appear. For each File 1 column selected, choose the corresponding column from File 2 using the dropdown menus. Ensure each File 2 column is mapped only once.
6.  **Reconcile:** Once files are uploaded, columns are selected, and mapping is complete, the "Reconcile Files" button will be enabled. Click it.
7.  **Results:** Wait for the processing to complete. A summary will be displayed.
8.  **Download Report:** Click the "Download Reconciliation Report (.xlsx)" button to get the detailed Excel report.
9.  **AI Analysis (Optional):** Click the "Get AI Analysis" button to receive insights from OpenAI based on the reconciliation summary.

## Notes

* The reconciliation logic performs case-insensitive and whitespace-trimmed comparisons for the key columns to improve matching accuracy.
* Discrepancies are identified by comparing all *other* columns between matched rows.
* Ensure the OpenAI API key is kept secure and not committed to public repositories.
* For production deployment, use a proper WSGI server (like Gunicorn or Waitress) instead of the Flask development server (`app.run(debug=True)`). Set `debug=False` in `app.run()` for production.
