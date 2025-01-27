from flask import Flask, request, render_template, redirect, url_for
import openpyxl
import pandas as pd
import os

app = Flask(__name__)

excel_files = {
    "Maintainance": "templates/Excel/Maintainance.xlsx",   
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

@app.route("/submit/<app_name>", methods=["POST"])
def submit(app_name):
    if app_name not in excel_files:
        return "Invalid application", 400
    data = request.form.to_dict()
    save_to_excel(app_name, data)
    print(f"Received data for {app_name}: {data}")    
    return redirect('Maintainance')

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
