from flask import Flask, request, render_template, redirect, url_for, jsonify, send_from_directory
import openpyxl
import schedule
import pandas as pd
import os
import time
import threading
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta

app = Flask(__name__)


# Define the root path for the Excel files
station1_excel_path = os.path.join(os.getcwd(), 'templates', 'Excel')



excel_files = {
 
    "ST1_Maintainance": "templates/Excel/ST1_Maintainance.xlsx",
    "ST2_Maintainance": "templates/Excel/ST2_Maintainance.xlsx",
    "ST3_Maintainance": "templates/Excel/ST3_Maintainance.xlsx",
    "ST4_Maintainance": "templates/Excel/ST4_Maintainance.xlsx",
    "ST1_Line_Rejection": "templates/Excel/ST1_Line_Rejection.xlsx",
    "ST2_Line_Rejection": "templates/Excel/ST2_Line_Rejection.xlsx",
    "ST3_Line_Rejection": "templates/Excel/ST3_Line_Rejection.xlsx",
    "ST4_Line_Rejection": "templates/Excel/ST4_Line_Rejection.xlsx",
    "ST1_Linesetup": "templates/Excel/ST1_Line_setup.xlsx",
    "ST2_Linesetup": "templates/Excel/ST2_Line_setup.xlsx",
    "ST3_Linesetup": "templates/Excel/ST3_Line_setup.xlsx",
    "ST4_Linesetup": "templates/Excel/ST4_Line_setup.xlsx",
    "ST1_Poka_yoke": "templates/Excel/ST1_POKA-YOKE.xlsx",    
    "ST2_Poka_yoke": "templates/Excel/ST2_POKA-YOKE.xlsx",    
    "ST3_Poka_yoke": "templates/Excel/ST3_POKA-YOKE.xlsx",    
    "ST4_Poka_yoke": "templates/Excel/ST4_POKA-YOKE.xlsx",    
    "ST1_Tool_Monitoring": "templates/Excel/ST1_Tool_Monitoring.xlsx",
    "ST2_Tool_Monitoring": "templates/Excel/ST2_Tool_Monitoring.xlsx",
    "ST3_Tool_Monitoring": "templates/Excel/ST3_Tool_Monitoring.xlsx",
    "ST4_Tool_Monitoring": "templates/Excel/ST4_Tool_Monitoring.xlsx",
    "AHP_FI": "templates/Excel/AHP_FI.xlsx",
  
}

shift_timings = {
    "SHIFT_1": (6, 30, 14, 30),
    "SHIFT_2": (14, 30, 22, 30),   
    "SHIFT_3": (22, 30, 6, 30),
}

# Set shift end times (Modify these for testing)
shift_end_times = {
    "SHIFT_1": "14:30",  # Change for testing
    "SHIFT_2": "22:30",  # Change for testing
    "SHIFT_3": "06:30",  # Change for testing
}


POKA_YOKE_APPS = {"ST1_Poka_yoke", "ST2_Poka_yoke", "ST3_Poka_yoke", "ST4_Poka_yoke"}
LINE_REJECTION_APPS = {"ST1_Line_Rejection", "ST2_Line_Rejection", "ST3_Line_Rejection", "ST4_Line_Rejection", "AHP_FI"}

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

@app.route("/Station2")
def Station2():
    return render_template("Station2.html") 

@app.route("/Station3")
def Station3():
    return render_template("Station3.html") 


@app.route("/Station4")
def Station4():
    return render_template("Station4.html") 

@app.route("/AHP_FI")
def AHP_FI():
    return render_template("AHP_FI.html")
    
@app.route("/ST1_Line_Rejection")
def ST1_Line_Rejection():
    return render_template("ST1_Line_Rejection.html")
    
@app.route("/ST2_Line_Rejection")
def ST2_Line_Rejection():
    return render_template("ST2_Line_Rejection.html")

@app.route("/ST3_Line_Rejection")
def ST3_Line_Rejection():
    return render_template("ST3_Line_Rejection.html")


@app.route("/ST4_Line_Rejection")
def ST4_Line_Rejection():
    return render_template("ST4_Line_Rejection.html")
    

@app.route("/ST1_Linesetup")
def ST1_Linesetup():
    return render_template("ST1_Linesetup.html")
    
@app.route("/ST2_Linesetup")
def ST2_Linesetup():
    return render_template("ST2_Linesetup.html")

@app.route("/ST3_Linesetup")
def ST3_Linesetup():
    return render_template("ST3_Linesetup.html")

@app.route("/ST4_Linesetup")
def ST4_Linesetup():
    return render_template("ST4_Linesetup.html")    
    
@app.route("/ST1_Maintainance")
def ST1_Maintainance():
    return render_template("ST1_Maintainance.html")
    
@app.route("/ST2_Maintainance")
def ST2_Maintainance():
    return render_template("ST2_Maintainance.html")

@app.route("/ST3_Maintainance")
def ST3_Maintainance():
    return render_template("ST3_Maintainance.html")

@app.route("/ST4_Maintainance")
def ST4_Maintainance():
    return render_template("ST4_Maintainance.html")    

@app.route("/ST1_Poka_yoke")
def ST1_Poka_yoke():
    return render_template("ST1_Poka_yoke.html")

@app.route("/ST2_Poka_yoke")
def ST2_Poka_yoke():
    return render_template("ST2_Poka_yoke.html")

@app.route("/ST3_Poka_yoke")
def ST3_Poka_yoke():
    return render_template("ST3_Poka_yoke.html")


@app.route("/ST4_Poka_yoke")
def ST4_Poka_yoke():
    return render_template("ST4_Poka_yoke.html")    

@app.route("/ST1_Tool_Monitoring")
def ST1_Tool_Monitoring():
    return render_template("ST1_Tool_Monitoring.html")
    
@app.route("/ST2_Tool_Monitoring")
def ST2_Tool_Monitoring():
    return render_template("ST2_Tool_Monitoring.html")

@app.route("/ST3_Tool_Monitoring")
def ST3_Tool_Monitoring():
    return render_template("ST3_Tool_Monitoring.html")


@app.route("/ST4_Tool_Monitoring")
def ST4_Tool_Monitoring():
    return render_template("ST4_Tool_Monitoring.html")    
    
@app.route("/ST1_Report")
def ST1_Report():
    return render_template("ST1_Report.html")

@app.route("/ST2_Report")
def ST2_Report():
    return render_template("ST2_Report.html")

@app.route("/ST3_Report")
def ST3_Report():
    return render_template("ST3_Report.html")

@app.route("/ST4_Report")
def ST4_Report():
    return render_template("ST4_Report.html")    
    
@app.route("/AHP_REVISION")
def AHP_REVISION():
    return render_template("AHP_REVISION.html")      

@app.route("/FI_REPORT")
def FI_REPORT():
    return render_template("FI_REPORT.html")      

@app.route("/AHP_ESCALATION")
def AHP_ESCALATION():
    return render_template("AHP_ESCALATION.html")


@app.route('/Excel/<filename>')
def serve_excel(filename):
    return send_from_directory('templates/Excel', filename)


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


def PY_send_email(station_name, failed_params):
    sender_email = "dmsprebo@gmail.com"  # Replace with your email
    recipient_emails = ["shopfloor.prebo01@prettl.com", "basavaraj.hiremath@prettl.com" , "shishir.cn@prettl.com" , "nagaraj.cm@prettl.com"]  # Replace with the recipient's email
    password = "smjqsoibldztfggk"  # Replace with your email password

    subject = f"🚨Alert: {station_name} - Failed Parameters🚨"
    body = f"Station {station_name} has the following failed parameters:\n\n"
    body += "\n".join(f"- {param}" for param in failed_params)
    body += "\n\nPlease check immediately."

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(recipient_emails)  # Join the recipients with commas
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(message)
        print(f"Email sent to {', '.join(recipient_emails)} for station {station_name}.")
    except Exception as e:
        print(f"Failed to send email: {e}")

def LR_send_email(station_name, failed_params, failure_percentage, recipient_emails):
    sender_email = "dmsprebo@gmail.com"
    password = "smjqsoibldztfggk"

    subject = f"🚨Alert: {station_name} - {failure_percentage:.2f}% Failure Rate🚨"
    body = f"Station {station_name} has the following failed parameters ({failure_percentage:.2f}%):\n\n"
    body += "\n".join(f"- {param}" for param in failed_params)
    body += "\n\nPlease check immediately."

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(recipient_emails)
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(message)
        print(f"Email sent to {', '.join(recipient_emails)} for station {station_name}.")
    except Exception as e:
        print(f"Failed to send email: {e}")




@app.route("/<path:app_name>", methods=["POST"])
def submit(app_name):
    actual_app_name = os.path.basename(app_name)
    print(f"\n---- Processing request for: {actual_app_name} ----")

    if actual_app_name not in excel_files:
        return "Invalid application", 400

    data = request.form.to_dict()
    save_to_excel(actual_app_name, data)  # Save data to Excel
    print(f"Received data for {actual_app_name}: {data}")

    # Call poka-yoke email check only if the application is in POKA_YOKE_APPS
    if actual_app_name in POKA_YOKE_APPS:
        check_and_send_poka_yoke_email(actual_app_name, data)
    elif actual_app_name in LINE_REJECTION_APPS:
        check_and_send_line_rejection_email(actual_app_name, data)
        

    time.sleep(2)
    return render_template(f"{app_name}.html")

def check_and_send_poka_yoke_email(actual_app_name, data):
    failed_params = [key for key, value in data.items() if value.lower() == "nok"]

    if failed_params:
        print(f"Alert: {actual_app_name} has the following failed parameters: {failed_params}")
        PY_send_email(actual_app_name, failed_params)  # Send email with failed parameters

def check_and_send_line_rejection_email(actual_app_name, data):
    total_count = int(data.get('totalCount', 0))
    failed_params = [key for key, value in list(data.items())[3:-3] if float(value) > 0]
    failure_count = len(failed_params)

    if total_count > 0:
        failure_percentage = (failure_count / total_count) * 100
    else:
        failure_percentage = 0

    if failure_percentage >= 3:  # Example threshold
        recipient_emails = ["anurag.khurana@prettl.com", "mahesh.bv@prettl.com", "nagaraj.cm@prettl.com", "shishir.cn@prettl.com", "basavaraj.hiremath@prettl.com", "shopfloor.prebo01@prettl.com"]
    elif failure_percentage >= 2:
        recipient_emails = ["mahesh.bv@prettl.com", "nagaraj.cm@prettl.com", "shishir.cn@prettl.com", "basavaraj.hiremath@prettl.com", "shopfloor.prebo01@prettl.com"]
    elif failure_percentage >= 1.5:
        recipient_emails = ["nagaraj.cm@prettl.com", "shishir.cn@prettl.com", "basavaraj.hiremath@prettl.com", "shopfloor.prebo01@prettl.com"]    
    elif failure_percentage >= 1:
        recipient_emails = ["basavaraj.hiremath@prettl.com", "shopfloor.prebo01@prettl.com"]
    elif failure_percentage >= 0.75:
        recipient_emails = ["shopfloor.prebo01@prettl.com"]
    else:
        recipient_emails = ["shopfloor.prebo01@prettl.com"]

    if failed_params:
        print(f"Alert: {actual_app_name} has the following failed parameters ({failure_percentage:.2f}%): {failed_params}")
        LR_send_email(actual_app_name, failed_params, failure_percentage, recipient_emails)
        

# ✅ Function to get current shift end time
def get_shift_end_time():
    now = datetime.now()
    for shift, time_str in shift_end_times.items():
        end_time = datetime.strptime(time_str, "%H:%M").replace(
            year=now.year, month=now.month, day=now.day
        )
        if now < end_time:
            return end_time
    return now + timedelta(days=1)  # If no shift matches, move to next day

# ✅ Function to check missing entries
def check_entries(shift_start, shift_end):
    missing_entries = []

    for station, file_path in excel_files.items():
        if not os.path.exists(file_path):
            print(f"❌ File {file_path} does not exist")
            missing_entries.append(station)
            continue

        try:
            df = pd.read_excel(file_path)

            # Check if "Date" column exists
            if df.empty or "Date" not in df.columns or df["Date"].dropna().empty:
                print(f"❌ No valid 'Date' column in {station}")
                missing_entries.append(station)
                continue

            # Convert "Date" column to datetime
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date"])  # Drop invalid dates

            if df.empty:
                print(f"❌ All dates invalid in {station}")
                missing_entries.append(station)
                continue

            # 🔹 Filter only entries within shift time
            shift_entries = df[(df["Date"] >= shift_start) & (df["Date"] <= shift_end)]

            if shift_entries.empty:
                print(f"🚨 No entries found for {station} during shift {shift_start} to {shift_end}")
                missing_entries.append(station)
            else:
                print(f"✅ Data found for {station}, skipping.")

        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            missing_entries.append(station)

    return missing_entries

# ✅ Function to monitor shifts
def monitor_shifts():
    while True:
        shift_end = get_shift_end_time()
        shift_start = shift_end - timedelta(hours=8)  # Assume 8-hour shift

        sleep_time = (shift_end - datetime.now()).total_seconds()
        print(f"🕒 Waiting until {shift_end} to check shift entries...")
        time.sleep(sleep_time)

        missing_entries = check_entries(shift_start, shift_end)

        if missing_entries:
            send_email(missing_entries, shift_start, shift_end)

# ✅ Start monitoring in a background thread
def start_monitoring():
    monitor_thread = threading.Thread(target=monitor_shifts, daemon=True)
    monitor_thread.start()


def send_email(missing_entries, shift_start, shift_end):
    sender_email = "dmsprebo@gmail.com"
    recipient_emails = ["basavaraj.hiremath@prettl.com"]
    password = "smjqsoibldztfggk"  # Replace with your email password

    subject = f"🚨 No Entries from {shift_start.strftime('%H:%M')} to {shift_end.strftime('%H:%M')} 🚨"
    body = f"The following stations have no entries between {shift_start.strftime('%H:%M')} and {shift_end.strftime('%H:%M')}:\n\n" + "\n".join(missing_entries)

    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(recipient_emails)  # Join the recipients with commas
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:  # Replace with your SMTP server
            server.starttls()
            server.login(sender_email, password)
            server.send_message(message)
            print(f"✅ Email sent successfully to {recipient_emails}")
    except Exception as e:
        print(f"❌ Failed to send email: {e}")



if __name__ == "__main__":
    start_monitoring()   
    app.run(debug=True)
