from flask import Flask, request, render_template, redirect, url_for, jsonify
import openpyxl
import pandas as pd
import os
import time

app = Flask(__name__)


# Define the root path for the Excel files
station1_excel_path = os.path.join(os.getcwd(), 'templates', 'Excel', 'ST1')



excel_files = {
    "ST1_Maintainance": "templates/Excel/ST1/ST1_Maintainance.xlsx",
    "ST1_Line_Rejection": "templates/Excel/ST1/ST1_Line_Rejection.xlsx",
    "ST1_Linesetup": "templates/Excel/ST1/ST1_Line_setup.xlsx",
    "ST1_Poka_yoke": "templates/Excel/ST1/ST1_POKA-YOKE.xlsx",    
    "ST1_Tool_Monitoring": "templates/ST1/ST1_Excel/Tool_Monitoring.xlsx",
    
}

def save_to_excel(app_name, data):
    file_name = excel_files[app_name]
    df = pd.DataFrame([data])
    if os.path.exists(file_name):
        existing_df = pd.read_excel(file_name)
        df = pd.concat([existing_df, df], ignore_index=True)
    df.to_excel(file_name, index=False)
    
@app.route("/")
def index():
    return render_template("index.html")
    
@app.route("/AhpAllStation")
def AhpAllStation():
    return render_template("AhpAllStation.html")      

@app.route("/Station1")
def Station1():
    return render_template("Station1.html")       
    
@app.route("/ST1_Line_Rejection")
def ST1_Line_Rejection():
    return render_template("ST1_Line_Rejection.html")

@app.route("/ST1_Linesetup")
def ST1_Linesetup():
    return render_template("ST1_Linesetup.html")
    
@app.route("/ST1_Maintainance")
def ST1_Maintainance():
    return render_template("ST1_Maintainance.html")

@app.route("/ST1_Poka_yoke")
def ST1_Poka_yoke():
    return render_template("ST1_Poka_yoke.html")

@app.route("/ST1_Tool_Monitoring")
def ST1_Tool_Monitoring():
    return render_template("ST1_Tool_Monitoring.html")
    
@app.route("/ST1_Report")
def Report():
    return render_template("ST1_Report.html")    

@app.route('/get_excel_data/<filename>', methods=['GET'])
def get_excel_data(filename):
    try:
        file_path = os.path.join(station1_excel_path, f'{filename}.xlsx')
        
        if not os.path.exists(file_path):
            return jsonify({'error': f'File {filename}.xlsx not found.'}), 404
        
        df = pd.read_excel(file_path)
        
        # Get filtering parameters
        start_date = request.args.get('start_date')  # Format: YYYY-MM-DD
        end_date = request.args.get('end_date')      # Format: YYYY-MM-DD
        query = request.args.get('query')           # Text to search
        
        # Convert 'Date' column to datetime if it exists
        if 'Date' in df.columns:
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

            # Filter by date range
            if start_date and end_date:
                start_date = pd.to_datetime(start_date)
                end_date = pd.to_datetime(end_date)
                df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]

            # Sort by Date (ascending order)
            df = df.sort_values(by='Date', ascending=True)
        
        # Filter by text query
        if query:
            df = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        
        # Convert DataFrame to HTML
        table_html = df.to_html(classes='table table-bordered', index=False)
        table_html = table_html.replace('<table', '<table style="text-align: center;"')
        
        return jsonify({'table_html': table_html})  
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route("/<path:app_name>", methods=["POST"])  # Allows slashes in app_name
def submit(app_name):
    actual_app_name = os.path.basename(app_name)  # Extracts last part after the final slash
    
    print(f"\n---- Processing request for: {actual_app_name} ----")  

    if actual_app_name not in excel_files:
        return "Invalid application", 400
    
    data = request.form.to_dict()
    save_to_excel(actual_app_name, data)  # Pass only the last part
    
    print(f"Received data for {actual_app_name}: {data}")  
    
    time.sleep(2)    
    return render_template(f"{app_name}.html")


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
