from flask import Flask, render_template, request, redirect, url_for, session, send_file
import sqlite3
from openpyxl import Workbook
from datetime import datetime

app = Flask(__name__)

DATABASE = "database.db"

ADMIN_PASSWORD = "NutriCheck@2026"

app.secret_key = "NutriCheck_Project_Secret_2026"


# =========================================================
# DATABASE
# =========================================================

def init_db():

    conn = sqlite3.connect(DATABASE)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS assessments (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            name TEXT,

            submitted_at TEXT,

            age_group TEXT,
            gender TEXT,
            occupation TEXT,
            diet TEXT,
            meals TEXT,
            fruitveg TEXT,
            dairy TEXT,
            protein TEXT,
            breakfast TEXT,
            processed_food TEXT,
            sunlight TEXT,
            diet_rating TEXT,
            water TEXT,
            nutrition_knowledge TEXT,
            diet_barrier TEXT,
            diagnosed TEXT,
            diagnosed_type TEXT,
            nutrition_source TEXT,
            supplements TEXT,
            awareness TEXT
        )
    """)

    # -----------------------------------------------------
    # Migration for your existing database
    # -----------------------------------------------------

    columns = conn.execute(
        "PRAGMA table_info(assessments)"
    ).fetchall()

    column_names = [column[1] for column in columns]

    # Add name if old database does not have it
    if "name" not in column_names:

        conn.execute(
            "ALTER TABLE assessments ADD COLUMN name TEXT"
        )

    # Add submitted_at if old database does not have it
    if "submitted_at" not in column_names:

        conn.execute(
            "ALTER TABLE assessments ADD COLUMN submitted_at TEXT"
        )

    conn.commit()

    conn.close()


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    return render_template("home.html")


# =========================================================
# ASSESSMENT
# =========================================================

@app.route("/assessment")
def assessment():

    return render_template("assessment.html")


# =========================================================
# RISK CALCULATION
# =========================================================

def calculate_risk(data):
    iron = 0
    b12 = 0
    vitamin_d = 0
    calcium = 0

    # Diet
    if data["diet"] == "Vegan":
        b12 += 2
        calcium += 1

    # Fruits and vegetables
    if data["fruitveg"] == "Rarely":
        iron += 2
    elif data["fruitveg"] == "1–3 days a week":
        iron += 1

    # Protein-rich foods
    if data["protein"] == "Never":
        iron += 2
        b12 += 2
    elif data["protein"] == "Rarely":
        iron += 1
        b12 += 1
    elif data["protein"] == "1–3 days a week":
        iron += 1

    # Milk / dairy / fortified alternatives
    if data["dairy"] == "Never":
        calcium += 3
        b12 += 1
    elif data["dairy"] == "Rarely":
        calcium += 2
        b12 += 1
    elif data["dairy"] == "1–3 days a week":
        calcium += 1

    # Sunlight / outdoor daylight
    if data["sunlight"] == "Rarely":
        vitamin_d += 3
    elif data["sunlight"] == "1–3 days a week":
        vitamin_d += 2
    elif data["sunlight"] == "4–6 days a week":
        vitamin_d += 1

    # Overall diet
    if data["diet_rating"] == "Poor":
        iron += 1
        b12 += 1
        vitamin_d += 1
        calcium += 1
    elif data["diet_rating"] == "Average":
        iron += 1
        calcium += 1

    # Meals
    if data["meals"] == "1":
        iron += 1
        b12 += 1
    elif data["meals"] == "2":
        iron += 1

    # Breakfast
    if data["breakfast"] == "Never":
        iron += 1
    elif data["breakfast"] == "Rarely":
        iron += 1

    # Processed food
    if data["processed_food"] == "Daily":
        iron += 1
        calcium += 1
    elif data["processed_food"] == "Often":
        calcium += 1

    # Nutrition knowledge
    if data["nutrition_knowledge"] in ["Poor", "Very poor"]:
        iron += 1
        b12 += 1
        vitamin_d += 1
        calcium += 1

    # Previously diagnosed deficiency
    diagnosed_type = data.get("diagnosed_type", "")

    if diagnosed_type == "Iron":
        iron += 3
    elif diagnosed_type == "B12":
        b12 += 3
    elif diagnosed_type == "D":
        vitamin_d += 3
    elif diagnosed_type == "Calcium":
        calcium += 3

    def level(score):
        if score >= 5:
            return "Higher possible risk"
        elif score >= 2:
            return "Moderate possible risk"
        else:
            return "Lower possible risk"

    return {
        "Iron": level(iron),
        "Vitamin B12": level(b12),
        "Vitamin D": level(vitamin_d),
        "Calcium": level(calcium)
    }


# =========================================================
# SUBMIT ASSESSMENT
# =========================================================

@app.route("/submit", methods=["POST"])
def submit():

    name = request.form.get("name", "").strip()

    if not name:
        return redirect(url_for("assessment"))

    data = {
        "age_group": request.form.get("age_group"),
        "gender": request.form.get("gender"),
        "occupation": request.form.get("occupation"),
        "diet": request.form.get("diet"),
        "meals": request.form.get("meals"),
        "fruitveg": request.form.get("fruitveg"),
        "dairy": request.form.get("dairy"),
        "protein": request.form.get("protein"),
        "breakfast": request.form.get("breakfast"),
        "processed_food": request.form.get("processed_food"),
        "sunlight": request.form.get("sunlight"),
        "diet_rating": request.form.get("diet_rating"),
        "water": request.form.get("water"),
        "nutrition_knowledge": request.form.get("nutrition_knowledge"),
        "diet_barrier": request.form.get("diet_barrier"),
        "diagnosed": request.form.get("diagnosed"),
        "diagnosed_type": request.form.get("diagnosed_type"),
        "nutrition_source": request.form.get("nutrition_source"),
        "supplements": request.form.get("supplements"),
        "awareness": request.form.get("awareness")
    }

    results = calculate_risk(data)

    # Save response to database
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    submitted_at = datetime.now().strftime("%d-%m-%Y %I:%M %p")

    cursor.execute("""
        INSERT INTO assessments (
            name, submitted_at, age_group, gender, occupation, diet,
            meals, fruitveg, dairy, protein, breakfast, processed_food,
            sunlight, diet_rating, water, nutrition_knowledge,
            diet_barrier, diagnosed, diagnosed_type, nutrition_source,
            supplements, awareness
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        name,
        submitted_at,
        data["age_group"],
        data["gender"],
        data["occupation"],
        data["diet"],
        data["meals"],
        data["fruitveg"],
        data["dairy"],
        data["protein"],
        data["breakfast"],
        data["processed_food"],
        data["sunlight"],
        data["diet_rating"],
        data["water"],
        data["nutrition_knowledge"],
        data["diet_barrier"],
        data["diagnosed"],
        data["diagnosed_type"],
        data["nutrition_source"],
        data["supplements"],
        data["awareness"]
    ))

    conn.commit()
    conn.close()

    return render_template(
        "result.html",
        results=results,
        name=name
    )


# =========================================================
# ADMIN LOGIN
# =========================================================

@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():

    if session.get("admin_logged_in"):

        return redirect(url_for("dashboard"))

    error = None

    if request.method == "POST":

        password = request.form.get("password")

        if password == ADMIN_PASSWORD:

            session["admin_logged_in"] = True

            return redirect(url_for("dashboard"))

        else:

            error = "Incorrect password."

    return render_template(
        "admin_login.html",
        error=error
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
def dashboard():

    if not session.get("admin_logged_in"):

        return redirect(url_for("admin_login"))

    conn = sqlite3.connect(DATABASE)

    conn.row_factory = sqlite3.Row

    responses = conn.execute("""
        SELECT *
        FROM assessments
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    total = len(responses)

    age_groups = {}
    diets = {}

    iron_risk = 0
    b12_risk = 0
    vitamin_d_risk = 0
    calcium_risk = 0

    # -----------------------------------------------------
    # Analyze every response
    # -----------------------------------------------------

    for row in responses:

        # Age
        age = row["age_group"]

        if age:

            age_groups[age] = age_groups.get(age, 0) + 1

        # Diet
        diet = row["diet"]

        if diet:

            diets[diet] = diets.get(diet, 0) + 1

        # Risk calculation

        iron = 0
        b12 = 0
        vitamin_d = 0
        calcium = 0

        if row["fruitveg"] == "rarely":

            iron += 2

        if row["protein"] == "rarely":

            iron += 2
            b12 += 1

        if row["protein"] == "never":

            iron += 2
            b12 += 2

        if row["diet"] == "Vegan":

            b12 += 2

        if row["dairy"] == "rarely":

            calcium += 2
            b12 += 1

        if row["dairy"] == "never":

            calcium += 3
            b12 += 1

        if row["sunlight"] == "rarely":

            vitamin_d += 3

        if row["sunlight"] == "1-3":

            vitamin_d += 1

        if iron >= 2:
            iron_risk += 1

        if b12 >= 2:
            b12_risk += 1

        if vitamin_d >= 2:
            vitamin_d_risk += 1

        if calcium >= 2:
            calcium_risk += 1

    return render_template(

        "dashboard.html",

        total=total,

        age_groups=age_groups,

        diets=diets,

        iron_risk=iron_risk,

        b12_risk=b12_risk,

        vitamin_d_risk=vitamin_d_risk,

        calcium_risk=calcium_risk,

        responses=responses

    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route("/export-excel")
def export_excel():

    if not session.get("admin_logged_in"):

        return redirect(url_for("admin_login"))

    conn = sqlite3.connect(DATABASE)

    cursor = conn.cursor()

    cursor.execute("""
        SELECT

            id,
            name,
            submitted_at,

            age_group,
            gender,
            occupation,
            diet,
            meals,
            fruitveg,
            dairy,
            protein,
            breakfast,
            processed_food,
            sunlight,
            diet_rating,
            water,
            nutrition_knowledge,
            diet_barrier,
            diagnosed,
            diagnosed_type,
            nutrition_source,
            supplements,
            awareness

        FROM assessments

        ORDER BY id
    """)

    rows = cursor.fetchall()

    conn.close()

    # -----------------------------------------------------
    # Create Excel workbook
    # -----------------------------------------------------

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "NutriCheck Responses"

    headers = [

        "ID",
        "Name",
        "Submission Date & Time",

        "Age Group",
        "Gender",
        "Occupation",
        "Diet",
        "Meals Per Day",
        "Fruits & Vegetables",
        "Dairy",
        "Protein",
        "Breakfast",
        "Processed Food",
        "Daylight Exposure",
        "Diet Rating",
        "Water",
        "Nutrition Knowledge",
        "Diet Barrier",
        "Previously Diagnosed",
        "Diagnosed Type",
        "Nutrition Information Source",
        "Supplements",
        "Nutrition Awareness"
    ]

    worksheet.append(headers)

    for row in rows:

        worksheet.append(list(row))

    # -----------------------------------------------------
    # Style Excel
    # -----------------------------------------------------

    for cell in worksheet[1]:

        cell.font = cell.font.copy(
            bold=True
        )

    # Column widths

    for column in worksheet.columns:

        max_length = 0

        column_letter = column[0].column_letter

        for cell in column:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max_length + 3,
            35
        )

    excel_file = "NutriCheck_Responses.xlsx"

    workbook.save(excel_file)

    return send_file(

        excel_file,

        as_attachment=True,

        download_name="NutriCheck_Responses.xlsx",

        mimetype=
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.pop(
        "admin_logged_in",
        None
    )

    return redirect(
        url_for("home")
    )


# =========================================================
# NUTRITION GUIDE
# =========================================================

@app.route("/nutrition-guide")
def nutrition_guide():

    return render_template(
        "nutrition_guide.html"
    )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    init_db()

    app.run(debug=True)