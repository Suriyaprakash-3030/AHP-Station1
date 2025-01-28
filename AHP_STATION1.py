from flask import Flask, request, render_template, redirect, url_for, jsonify
import openpyxl
import pandas as pd
import os
import time

app = Flask(__name__)


# Define the root path for the Excel files
EXCEL_FOLDER = os.path.join(os.getcwd(), 'templates', 'Excel')


excel_files = {
    "Maintainance": "templates/Excel/Maintainance.xlsx",
    "Line_Rejection": "templates/Excel/Line_Rejection.xlsx",
    "Linesetup": "templates/Excel/Line_setup.xlsx",
    "Poka_yoke": "templates/Excel/POKA-YOKE.xlsx",    
    "Tool_Monitoring": "templates/Excel/Tool_Monitoring.xlsx",
    
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
    return render_template("Front_page.html")
    
@app.route("/Line_Rejection")
def Line_Rejection():
    return render_template("Line_Rejection.html")

@app.route("/Linesetup")
def Linesetup():
    return render_template("Linesetup.html")
    
@app.route("/Maintainance")
def Maintainance():
    return render_template("Maintainance.html")

@app.route("/Poka_yoke")
def Poka_yoke():
    return render_template("Poka_yoke.html")

@app.route("/Tool_Monitoring")
def Tool_Monitoring():
    return render_template("Tool_Monitoring.html")
    
@app.route("/Report")
def Report():
    return render_template("Report.html")    

@app.route('/get_excel_data/<filename>')
def get_excel_data(filename):
    try:
        # Build the file path for the Excel file
        file_path = os.path.join(EXCEL_FOLDER, f'{filename}.xlsx')
        
        # Check if the file exists
        if not os.path.exists(file_path):
            return jsonify({'error': f'File {filename}.xlsx not found.'}), 404
        
        # Read the Excel file using pandas
        df = pd.read_excel(file_path)
        
        # Convert the dataframe to HTML with center-aligned cells
        table_html = df.to_html(classes='table table-bordered', index=False, 
                                justify='center')  # Center align all columns
        
        # You can also manually center-align if required
        table_html = table_html.replace('<table', '<table style="text-align: center;"')


        return jsonify({'table_html': table_html})  # Return the HTML table data as JSON
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route("/<app_name>", methods=["POST"])
def submit(app_name):
    if app_name not in excel_files:
        return "Invalid application", 400
    data = request.form.to_dict()
    save_to_excel(app_name, data)
    print(f"Received data for {app_name}: {data}")   
    time.sleep(2)    
    return render_template(f"{app_name}.html")

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
