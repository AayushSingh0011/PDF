from flask import Flask, render_template, request, redirect, send_file, flash
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = "supersecretkey"

EXCEL_FILE = "students_data.xlsx"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        data = {
            "Name": request.form.get("name")
        }

        df_new = pd.DataFrame([data])

        if os.path.exists(EXCEL_FILE):
            df_old = pd.read_excel(EXCEL_FILE)
            df = pd.concat([df_old, df_new], ignore_index=True)
        else:
            df = df_new

        df.to_excel(EXCEL_FILE, index=False)
        flash("Submitted successfully")
        return redirect("/")

    return render_template("index.html")


@app.route("/download")
def download_excel():
    return send_file(EXCEL_FILE, as_attachment=True)


if __name__ == "__main__":
    app.run()
