from flask import Flask, flash, redirect, render_template, request, url_for  # type: ignore

from business import (
    delete_wfh_excel_file,
    fill_empty_weekdays_with_wfo,
    get_wfh_selection_table,
    get_wfh_plan_emails,
    get_wfh_plan_table,
    submit_weekly_wfh_plan,
)
from send_mail import send_wfh_summary_mail

app = Flask(__name__)
app.secret_key = "change-this-secret-key"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        result = submit_weekly_wfh_plan(
            request.form.getlist("wfh_selection"),
            request.form.get("week_start", ""),
            request.form.get("week_end", ""),
        )
        flash(result.message, result.category)
        return redirect(url_for("index"))

    return render_template(
        "index.html",
        table_data=get_wfh_selection_table(
            request.args.get("week_start"),
            request.args.get("week_end"),
        ),
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
