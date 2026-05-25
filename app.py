from flask import Flask, flash, redirect, render_template, request, url_for  # type: ignore

from business import (
    current_week_range,
    delete_wfh_excel_file,
    fill_empty_weekdays_with_wfo,
    get_wfh_plan_emails,
    get_wfh_plan_table,
    next_week_range,
    submit_wfh_request,
)
from send_mail import send_wfh_summary_mail

app = Flask(__name__)
app.secret_key = "change-this-secret-key"


@app.route("/", methods=["GET", "POST"])
def index():
    next_week_start, next_week_end = next_week_range()

    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip()
        emp_id = request.form.get("emp_id", "").strip()
        wfh = request.form.get("wfh")
        result = submit_wfh_request(
            name,
            email,
            emp_id,
            wfh,
            request.form.get("wfh_date", ""),
        )
        flash(result.message, result.category)
        return redirect(url_for("index"))

    return render_template(
        "index.html",
        week_start=next_week_start.isoformat(),
        week_end=next_week_end.isoformat(),
    )


@app.route("/wfh-plan")
def wfh_plan():
    table_data = get_wfh_plan_table()
    return render_template("wfh_plan.html", table_data=table_data)


@app.route("/send-summary-mail", methods=["POST"])
def send_summary_mail():
    fill_empty_weekdays_with_wfo()
    table_data = get_wfh_plan_table()
    recipients = get_wfh_plan_emails()

    if not recipients:
        flash("No email IDs found in the Excel sheet.", "error")
        return redirect(url_for("wfh_plan"))

    try:
        send_wfh_summary_mail(recipients, table_data)
    except Exception as error:
        flash(f"Mail sending failed: {error}", "error")
        return redirect(url_for("wfh_plan"))

    delete_wfh_excel_file()
    flash("Form submitted successfully.", "success")
    flash(f"Summary mail sent to {len(recipients)} members.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    app.run(
        debug=True,
        host="localhost",
        port=5000
    )
