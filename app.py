from flask import Flask, flash, jsonify, redirect, render_template, request, url_for  # type: ignore

from business import (
    add_user,
    delete_user,
    delete_wfh_excel_file,
    fill_empty_weekdays_with_wfo,
    get_user_records,
    get_wfh_selection_table,
    get_wfh_plan_emails,
    get_wfh_plan_table,
    submit_weekly_wfh_plan,
    update_user,
)
from send_mail import send_wfh_summary_mail

app = Flask(__name__)
app.secret_key = "change-this-secret-key"


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        week_start = request.form.get("week_start", "")
        week_end = request.form.get("week_end", "")
        result = submit_weekly_wfh_plan(
            request.form.getlist("wfh_selection"),
            week_start,
            week_end,
        )
        flash(result.message, result.category)
        return redirect(url_for("index", week_start=week_start, week_end=week_end))

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


@app.route("/users")
def manage_users():
    return render_template("manage_users.html", users=get_user_records())


@app.route("/add-user")
def add_user_page():
    return redirect(url_for("manage_users"))


@app.route("/api/users", methods=["POST"])
def create_user():
    payload = request.get_json(silent=True) or request.form
    result, employee = add_user(
        payload.get("name", ""),
        payload.get("tid", ""),
        payload.get("email", ""),
    )

    if request.is_json:
        status_code = 201 if result.category == "success" else 400
        return jsonify({
            "message": result.message,
            "category": result.category,
            "user": employee,
        }), status_code

    flash(result.message, result.category)
    return redirect(url_for("manage_users"))


@app.route("/users/<int:user_index>/update", methods=["POST"])
def edit_user(user_index):
    result, _ = update_user(
        user_index,
        request.form.get("name", ""),
        request.form.get("tid", ""),
        request.form.get("email", ""),
    )
    flash(result.message, result.category)
    return redirect(url_for("manage_users"))


@app.route("/users/<int:user_index>/delete", methods=["POST"])
def remove_user(user_index):
    result = delete_user(user_index)
    flash(result.message, result.category)
    return redirect(url_for("manage_users"))


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
