from flask import Flask, render_template, send_from_directory, send_file, request
from createDoc import create_month_doc, create_month_html
import calendar
from datetime import date

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/monthy-doc')
def monthDoc():
    months = [(i, calendar.month_name[i]) for i in range(1, 13)]
    current_month = date.today().month

    return render_template('monthoverview.html', months=months, current_month=current_month)

@app.route('/create-monthly-doc', methods=["POST"])
def createMonthDoc():
    selected_months = request.form.getlist("months")
    showHolidays = 'showHolidays' in request.form
    holiday_url = request.form.get("holidayURL") or "https://calendar.google.com/calendar/ical/de.german%23holiday@group.v.calendar.google.com/public/basic.ics"

    if selected_months:
        months = [int(m) for m in selected_months]
    else:
        months = None # --> current month

    file_stream, filename = create_month_doc(months, showHolidays, url=holiday_url)

    return send_file(file_stream, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.route('/preview', methods=["POST"])
def preview():
    selected_months = request.form.getlist("months")
    showHolidays = 'showHolidays' in request.form
    months = [int(m) for m in selected_months] if selected_months else None

    return create_month_html(months, showHolidays)

if __name__ == "__main__":
    app.run(debug=True, port=5002, host="0.0.0.0")