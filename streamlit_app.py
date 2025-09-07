import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from io import BytesIO
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# import smtplib, ssl
# from email.message import EmailMessage

# Google Sheets setup (email related removed)
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(
    'attendance-compliance-shashwat-97b0b49a6cdb.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Database").sheet1

# YOUR_EMAIL = "scripster1218@gmail.com"
# SENDGRID_API_KEY = "YOUR_SENDGRID_API_KEY"
# SMTP_SERVER = "smtp.sendgrid.net"
# SMTP_PORT = 465

# def send_email(to, subject, body):
#     msg = EmailMessage()
#     msg.set_content(body)
#     msg['Subject'] = subject
#     msg['From'] = YOUR_EMAIL
#     msg['To'] = to
#     context = ssl.create_default_context()
#     try:
#         with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
#             server.login("apikey", SENDGRID_API_KEY)
#             server.send_message(msg)
#         st.write(f"Sent email to {to} with subject '{subject}'")
#     except Exception as e:
#         st.error(f"Failed to send email to {to}: {e}")

def emp_id_clean(id_val):
    try:
        return int(float(str(id_val).strip()))
    except Exception:
        return str(id_val).strip()

def custom_round(value):
    integer_part = int(np.floor(value))
    decimal_part = value - integer_part
    if decimal_part <= 0.4:
        return integer_part
    else:
        return integer_part + 1

def process_data(roster_file, attendance_file):
    roster_df = pd.read_excel(roster_file, sheet_name='Sheet1', header=None)
    attendance_df = pd.read_excel(attendance_file, sheet_name='Sheet1', header=None)
    cols_roster = ['Dept', 'Employee ID', 'First Name', 'Last Name'] + list(range(4, 32))
    roster_df.columns = cols_roster
    roster_df = roster_df.iloc[3:].reset_index(drop=True)
    cols_attendance = ['Dept', 'Employee ID', 'First Name', 'Last Name', 'Attendance']
    attendance_df.columns = cols_attendance
    attendance_df = attendance_df.iloc[3:].reset_index(drop=True)
    shifts = roster_df.loc[:, list(range(4, 32))]
    day_labels = pd.read_excel(roster_file, sheet_name='Sheet1', header=None).iloc[2, 4:32].values
    weekday_mask = np.isin(day_labels, ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'])
    working_days, office_days, sl_count, al_count, hsl_count = [], [], [], [], []
    for idx, row in shifts.iterrows():
        arr = row.values
        # Count M, D, N all days (Mon-Sun)
        wd_base = np.sum(np.isin(arr, ['M', 'D', 'N']))
        # Add half-day for each HSL (half-day sick leave)
        hsl_total = np.sum(arr == 'HSL')
        wd_final = wd_base + 0.5 * hsl_total  # Do NOT subtract SL or A/L
        working_days.append(round(wd_final, 2))
        sl_count.append(np.sum(arr == 'SL'))        # For info only
        al_count.append(np.sum(arr == 'A/L'))       # For info only
        hsl_count.append(hsl_total)
        # Working days from office = 60% of (M+D + 0.5*HSL on Mon-Fri)
        arr_weekday = arr[weekday_mask]
        wd_compliance_base = np.sum((arr_weekday == 'M') | (arr_weekday == 'D')) + np.sum(arr_weekday == 'HSL') * 0.5
        rounded_val = wd_compliance_base * 0.6
        office_days.append(custom_round(rounded_val))
    actual_present = attendance_df['Attendance'].astype(float).round(2)
    office_days_np = np.array(office_days)
    with np.errstate(divide='ignore', invalid='ignore'):
        adjusted_attendance = np.where(office_days_np == 0, 0, actual_present / office_days_np * 5)
    adjusted_attendance = np.round(adjusted_attendance, 2)
    compliant = adjusted_attendance >= 3
    days_needed = np.ceil(3 * office_days_np / 5)
    days_missed = (days_needed - actual_present).clip(lower=0).astype(int)
    results = pd.DataFrame({
        'Employee ID': roster_df['Employee ID'].apply(lambda x: emp_id_clean(x)),
        'First Name': roster_df['First Name'],
        'Last Name': roster_df['Last Name'],
        'Working Days': working_days,
        'Working Days (from Office)': office_days,
        'Actual Present': actual_present,
        'Adjusted Attendance': adjusted_attendance,
        'Compliant': np.where(compliant, 'Yes', 'No'),
        'Days Missed for Compliance': days_missed,
        'Sick Leave Days': sl_count,
        'Annual Leave Days': al_count,
        'Half-Day Sick Leave Days': hsl_count
    })
    return results

def styled_dataframe(df):
    df = df.copy()
    df.index = np.arange(1, len(df) + 1)
    float_cols = ['Working Days', 'Working Days (from Office)', 'Actual Present', 'Adjusted Attendance']
    for col in float_cols:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{float(x):.2f}")
    styler = df.style.set_properties(**{'text-align': 'center'})
    styler = styler.set_table_styles([
        {'selector': 'th', 'props': [('text-align', 'center'), ('vertical-align', 'middle')]},
        {'selector': 'td', 'props': [('vertical-align', 'middle')]}
    ])
    return styler

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ComplianceSummary')
    return output.getvalue()

def plotly_fig_to_png(fig):
    import plotly.io as pio
    return pio.to_image(fig, format="png")

def add_table_to_pdf(pdf, df, title):
    fig, ax = plt.subplots(figsize=(12, 0.7 + len(df) * 0.4))
    ax.axis('off')
    table = ax.table(
        cellText=df.values,
        colLabels=df.columns,
        loc='center',
        cellLoc='center',
        colColours=["#c7e9fc"] * df.shape[1]
    )
    plt.title(title, fontsize=14)
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.auto_set_column_width(col=list(range(len(df.columns))))
    for key, cell in table.get_celld().items():
        cell.set_text_props(ha='center', va='center')
    pdf.savefig(fig, bbox_inches='tight')
    plt.close(fig)

def generate_pdf_report(full_table, top_performers, non_compliant):
    pdf_buffer = BytesIO()
    with PdfPages(pdf_buffer) as pdf:
        plt.figure(figsize=(10, 4))
        plt.axis('off')
        plt.title("BT Group Attendance Compliance Report", fontsize=20)
        stats = [
            f"Total Employees: {len(full_table)}",
            f"Compliant Employees: {full_table['Compliant'].value_counts().get('Yes', 0)}",
            f"Non-Compliant Employees: {full_table['Compliant'].value_counts().get('No', 0)}",
            f"Average Adjusted Attendance: {full_table['Adjusted Attendance'].mean():.2f}",
        ]
        for i, stat in enumerate(stats):
            plt.text(0, 0.7 - i * 0.15, stat, fontsize=14)
        pdf.savefig()
        plt.close()
        for fig in [
            px.pie(full_table, names='Compliant', title='Compliance Distribution',
                   color='Compliant', color_discrete_map={'Yes': 'green', 'No': 'red'}),
            px.bar(full_table.sort_values('Working Days', ascending=False), x='First Name', y='Working Days',
                   title="Total Working Days per Employee", color='Compliant',
                   color_discrete_map={'Yes': 'green', 'No': 'red'}),
            px.bar(full_table.sort_values('Days Missed for Compliance', ascending=False), x='First Name',
                   y='Days Missed for Compliance', title="Days Missed for Compliance", color='Compliant',
                   color_discrete_map={'Yes': 'green', 'No': 'red'})
        ]:
            try:
                img_bytes = plotly_fig_to_png(fig)
                fig_plt = plt.figure(figsize=(10, 6))
                plt.imshow(plt.imread(BytesIO(img_bytes)))
                plt.axis('off')
                pdf.savefig(fig_plt, bbox_inches='tight')
                plt.close(fig_plt)
            except Exception:
                pass
        add_table_to_pdf(pdf, full_table, "Full Compliance Table")
        if not top_performers.empty:
            add_table_to_pdf(pdf, top_performers, "Top Performers")
        if not non_compliant.empty:
            add_table_to_pdf(pdf, non_compliant, "Non-Compliant Employees")
    pdf_buffer.seek(0)
    return pdf_buffer.read()

def notify_strikes(results, compliance_month):
    all_records = sheet.get_all_records()
    strikes_updated = 0
    for _, row in results.iterrows():
        emp_id = emp_id_clean(row['Employee ID'])
        non_compliant = str(row['Compliant']).strip().lower() == 'no'
        db_row = next((r for r in all_records if emp_id_clean(r['Employee ID']) == emp_id), None)
        if not db_row:
            st.warning(f"Employee ID {emp_id} not found in database sheet. Skipping.")
            continue
        history = db_row.get('Compliance History', "")
        history_list = history.split(',') if history else []
        strike_count = int(db_row.get('Strike Count', '0'))
        if non_compliant and compliance_month not in history_list:
            strike_count += 1
            history_list.append(compliance_month)
            row_index = all_records.index(db_row) + 2
            sheet.update_cell(row_index, sheet.find('Strike Count').col, str(strike_count))
            sheet.update_cell(row_index, sheet.find('Compliance History').col, ','.join(history_list))
            # recipients = [db_row['Email']]
            # if strike_count >= 2:
            #     recipients.append(db_row['Manager Email'])
            # if strike_count >= 3:
            #     recipients.append(db_row['HR Email'])
            # subject = f"[Attendance Compliance] Strike {strike_count} for {db_row['First Name']} {db_row['Last Name']}"
            # body = (
            #     f"Dear {db_row['First Name']},\n\n"
            #     f"You have received strike {strike_count} for attendance non-compliance for {compliance_month}.\n"
            #     f"Strike History Dates: {', '.join(history_list)}\n"
            # )
            # if strike_count == 1:
            #     body += "\nYou have 30 days to fix your attendance compliance."
            # elif strike_count == 2:
            #     body += "\nYour manager has been notified. Please address this immediately."
            # elif strike_count >= 3:
            #     body += "\nHR has been notified. Immediate action required to avoid escalation."
            # for recipient in recipients:
            #     if recipient:
            #         send_email(recipient, subject, body)
            strikes_updated += 1
    if strikes_updated == 0:
        st.info("No new strikes to notify.")
    else:
        st.success(f"Strike records updated for {strikes_updated} employees (email sending skipped).")

def main():
    st.set_page_config(page_title="BT Attendance Compliance Dashboard", layout="wide", page_icon="BT.png")
    st.title("BT Attendance Compliance Dashboard")
    roster_file = st.file_uploader("Upload Roster Excel file", type=["xls", "xlsx"])
    attendance_file = st.file_uploader("Upload Attendance Excel file", type=["xls", "xlsx"])
    with st.form("calc_form"):
        submit = st.form_submit_button("Calculate Compliance")
    results = None
    if submit and roster_file and attendance_file:
        with st.spinner("Processing data, please wait..."):
            results = process_data(roster_file, attendance_file)
            st.session_state['results'] = results
            num_top = min(5, len(results))
            top_performers = results.sort_values(by='Adjusted Attendance', ascending=False).head(num_top)
            non_compliant = results[results['Compliant'] == 'No']
            pdf_bytes = generate_pdf_report(results, top_performers, non_compliant)
            st.session_state['pdf_bytes'] = pdf_bytes
    if 'results' in st.session_state:
        results = st.session_state['results']
        st.subheader(f"Compliance Data for {len(results)} Employees")
        st.dataframe(styled_dataframe(results), width=1500)
        tabs = st.tabs(["Summary", "Visualizations", "Top Performers",
                        "Non-Compliant Employees", "Export", "Notify"])
        with tabs[0]:
            st.markdown("### 📊 Summary Statistics")
            st.metric("Total Employees", len(results))
            st.metric("Compliant Employees", results['Compliant'].value_counts().get('Yes', 0))
            st.metric("Non-Compliant Employees", results['Compliant'].value_counts().get('No', 0))
            st.metric("Average Adjusted Attendance", results['Adjusted Attendance'].astype(float).mean().round(2))
        with tabs[1]:
            st.markdown("### 📈 Visualizations")
            col1, col2 = st.columns(2)
            with col1:
                fig_pie = px.pie(results, names='Compliant', title='Compliance Distribution',
                                 color='Compliant', color_discrete_map={'Yes': 'green', 'No': 'red'})
                st.plotly_chart(fig_pie, use_container_width=True)
                fig_bar_days = px.bar(results.sort_values('Working Days', ascending=False),
                                      x='First Name', y='Working Days', title="Total Working Days per Employee",
                                      color='Compliant')
                st.plotly_chart(fig_bar_days, use_container_width=True)
            with col2:
                fig_bar_missed = px.bar(results.sort_values('Days Missed for Compliance', ascending=False),
                                        x='First Name', y='Days Missed for Compliance',
                                        title="Days Missed for Compliance", color='Compliant')
                st.plotly_chart(fig_bar_missed, use_container_width=True)
        with tabs[2]:
            st.markdown("### 🏆 Top Performers")
            max_top = min(20, len(results))
            num_top = st.number_input("Enter number of top performers:", 1, max_top, 5, key="top_performer")
            top_performers = results.sort_values(by='Adjusted Attendance', ascending=False).head(num_top)
            st.dataframe(styled_dataframe(top_performers[['Employee ID', 'First Name', 'Last Name',
                                                          'Adjusted Attendance', 'Compliant']]), width=1000)
        with tabs[3]:
            st.markdown("### ❌ Non-Compliant Employees")
            non_compliant = results[results['Compliant'] == 'No']
            st.dataframe(styled_dataframe(non_compliant[ 
                ['Employee ID', 'First Name', 'Last Name', 'Days Missed for Compliance', 'Sick Leave Days',
                 'Annual Leave Days', 'Half-Day Sick Leave Days']]), width=1000)
        with tabs[4]:
            st.markdown("### 📥 Export Data and Visual Report")
            excel_bytes = to_excel(results)
            st.download_button("Download Data as Excel", excel_bytes,
                               file_name="compliance_data.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            if 'pdf_bytes' in st.session_state:
                st.download_button(label="Download PDF Report",
                                   data=st.session_state['pdf_bytes'],
                                   file_name="compliance_visual_report.pdf",
                                   mime="application/pdf")
        with tabs[5]:
            st.markdown("### 📧 Notify Non-Compliant Employees")
            compliance_month = st.text_input("Compliance Month (YYYY-MM) - for strike tracking",
                                            value=pd.Timestamp.today().strftime("%Y-%m"))
            st.write("Click the button below to notify all newly non-compliant employees and update strike records.")
            if st.button("Send Notifications"):
                notify_strikes(results, compliance_month)
            if st.button("Reset Strike Records (DANGER)"):
                st.warning("You are resetting all strike counts and compliance history!")
                all_records = sheet.get_all_records()
                for i in range(len(all_records)):
                    row_index = i + 2
                    sheet.update_cell(row_index, sheet.find('Strike Count').col, '0')
                    sheet.update_cell(row_index, sheet.find('Compliance History').col, '')
                st.success("All strike records reset successfully.")
    else:
        st.info("Please upload both Roster and Attendance files and click 'Calculate Compliance'.")

if __name__ == "__main__":
    main()
