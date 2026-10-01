from flask import Flask, render_template, request, redirect, url_for, session, send_file
import psycopg
from psycopg.rows import dict_row
from openpyxl import Workbook
from datetime import datetime
import os

app = Flask(__name__)

app.secret_key = "NutriCheck_Project_Secret_2026"

ADMIN_PASSWORD = "NutriCheck@2026"

DATABASE_URL = os.environ.get("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError("DATABASE_URL environment variable is not set")


def get_db():
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)


# --------------------------------------------------
# DATABASE INITIALIZATION
# --------------------------------------------------

def init_db():
    conn = get_db()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS assessments (
            id SERIAL PRIMARY KEY,
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

    conn.commit()
    conn.close()


# --------------------------------------------------
# RISK CALCULATION
# --------------------------------------------------

def calculate_risk(data):

    iron = 0
    b12 = 0
    vitamin_d = 0
    calcium = 0

    diet = data.get("diet", "")
    fruitveg = data.get("fruitveg", "")
    dairy = data.get("dairy", "")
    protein = data.get("protein", "")
    sunlight = data.get("sunlight", "")
    diet_rating = data.get("diet_rating", "")
    meals = data.get("meals", "")
    breakfast = data.get("breakfast", "")
    processed_food = data.get("processed_food", "")
    nutrition_knowledge = data.get("nutrition_knowledge", "")
    diagnosed_type = data.get("diagnosed_type", "")

    # Iron
    if fruitveg in ["Rarely"]:
        iron += 2

    if protein == "Rarely":
        iron += 2
        b12 += 1

    if protein == "Never":
        iron += 2
        b12 += 2

    # Vitamin B12
    if diet in ["Vegan"]:
        b12 += 2

    if dairy == "Rarely":
        b12 += 1

    if dairy == "Never":
        b12 += 1

    # Vitamin D
    if sunlight == "Rarely":
        vitamin_d += 3

    elif sunlight == "1–3 days a week":
        vitamin_d += 1

    # Calcium
    if dairy == "Rarely":
        calcium += 2

    elif dairy == "Never":
        calcium += 3

    # Overall diet
    if diet_rating == "Poor":
        iron += 1
        b12 += 1
        vitamin_d += 1
        calcium += 1

    # Meals
    if meals == "1":
        iron += 1
        b12 += 1
        calcium += 1

    # Breakfast
    if breakfast == "Rarely":
        iron += 1
        b12 += 1

    elif breakfast == "Never":
        iron += 1
        b12 += 1

    # Processed food
    if processed_food == "Daily":
        iron += 1
        calcium += 1

    elif processed_food == "Often":
        iron += 1

    # Nutrition knowledge
    if nutrition_knowledge in ["Poor", "Very poor"]:
        iron += 1
        b12 += 1
        vitamin_d += 1
        calcium += 1

    # Previously diagnosed category
    if "Iron" in diagnosed_type:
        iron += 2

    if "B12" in diagnosed_type:
        b12 += 2

    if "D" in diagnosed_type:
        vitamin_d += 2

    if "Calcium" in diagnosed_type:
        calcium += 2

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


# --------------------------------------------------
# HOME
# --------------------------------------------------

@app.route("/")
def home():
    return render_template("home.html")


# --------------------------------------------------
# ASSESSMENT
# --------------------------------------------------

@app.route("/assessment")
def assessment():
    return render_template("assessment.html")


# --------------------------------------------------
# SUBMIT ASSESSMENT
# --------------------------------------------------

@app.route("/submit", methods=["POST"])
def submit():

    name = request.form.get("name", "").strip()

    if not name:
        return redirect(url_for("assessment"))

    data = {
        "age_group": request.form.get("age_group", ""),
        "gender": request.form.get("gender", ""),
        "occupation": request.form.get("occupation", ""),
        "diet": request.form.get("diet", ""),
        "meals": request.form.get("meals", ""),
        "fruitveg": request.form.get("fruitveg", ""),
        "dairy": request.form.get("dairy", ""),
        "protein": request.form.get("protein", ""),
        "breakfast": request.form.get("breakfast", ""),
        "processed_food": request.form.get("processed_food", ""),
        "sunlight": request.form.get("sunlight", ""),
        "diet_rating": request.form.get("diet_rating", ""),
        "water": request.form.get("water", ""),
        "nutrition_knowledge": request.form.get("nutrition_knowledge", ""),
        "diet_barrier": request.form.get("diet_barrier", ""),
        "diagnosed": request.form.get("diagnosed", ""),
        "diagnosed_type": request.form.get("diagnosed_type", ""),
        "nutrition_source": request.form.get("nutrition_source", ""),
        "supplements": request.form.get("supplements", ""),
        "awareness": request.form.get("awareness", "")
    }

    submitted_at = datetime.now().strftime("%d-%m-%Y %I:%M %p")

    conn = get_db()

    conn.execute("""
        INSERT INTO assessments (
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
        )
        VALUES (
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
            %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
            %s, %s
        )
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

    results = calculate_risk(data)

    return render_template(
        "result.html",
        results=results,
        name=name
    )


# --------------------------------------------------
# ADMIN LOGIN
# --------------------------------------------------

@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():

    if request.method == "POST":

        password = request.form.get("password", "")

        if password == ADMIN_PASSWORD:
            session["admin_logged_in"] = True
            return redirect(url_for("dashboard"))

        return render_template(
            "admin_login.html",
            error="Incorrect password"
        )

    return render_template("admin_login.html")


# --------------------------------------------------
# DASHBOARD
# --------------------------------------------------

@app.route("/dashboard")
def dashboard():

    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    conn = get_db()

    responses = conn.execute("""
        SELECT *
        FROM assessments
        ORDER BY id DESC
    """).fetchall()

    age_rows = conn.execute("""
        SELECT age_group, COUNT(*) AS count
        FROM assessments
        GROUP BY age_group
        ORDER BY age_group
    """).fetchall()

    diet_rows = conn.execute("""
        SELECT diet, COUNT(*) AS count
        FROM assessments
        GROUP BY diet
        ORDER BY diet
    """).fetchall()

    conn.close()

    age_groups = {
        row["age_group"] or "Not specified": row["count"]
        for row in age_rows
    }

    diets = {
        row["diet"] or "Not specified": row["count"]
        for row in diet_rows
    }

    iron_risk = 0
    b12_risk = 0
    vitamin_d_risk = 0
    calcium_risk = 0

    for response in responses:

        result = calculate_risk(response)

        if result["Iron"] != "Lower possible risk":
            iron_risk += 1

        if result["Vitamin B12"] != "Lower possible risk":
            b12_risk += 1

        if result["Vitamin D"] != "Lower possible risk":
            vitamin_d_risk += 1

        if result["Calcium"] != "Lower possible risk":
            calcium_risk += 1

    return render_template(
        "dashboard.html",
        responses=responses,
        total=len(responses),
        age_groups=age_groups,
        diets=diets,
        iron_risk=iron_risk,
        b12_risk=b12_risk,
        vitamin_d_risk=vitamin_d_risk,
        calcium_risk=calcium_risk
    )


# --------------------------------------------------
# EXPORT EXCEL
# --------------------------------------------------

@app.route("/export-excel")
def export_excel():

    if not session.get("admin_logged_in"):
        return redirect(url_for("admin_login"))

    conn = get_db()

    responses = conn.execute("""
        SELECT *
        FROM assessments
        ORDER BY id
    """).fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "NutriCheck Responses"

    headers = [
        "ID",
        "Name",
        "Submission Date & Time",
        "Age Group",
        "Gender",
        "Occupation",
        "Diet",
        "Meals",
        "Fruits & Vegetables",
        "Dairy",
        "Protein",
        "Breakfast",
        "Processed Food",
        "Sunlight",
        "Diet Rating",
        "Water",
        "Nutrition Knowledge",
        "Diet Barrier",
        "Diagnosed",
        "Diagnosed Type",
        "Nutrition Source",
        "Supplements",
        "Awareness"
    ]

    ws.append(headers)

    for row in responses:
        ws.append([
            row["id"],
            row["name"],
            row["submitted_at"],
            row["age_group"],
            row["gender"],
            row["occupation"],
            row["diet"],
            row["meals"],
            row["fruitveg"],
            row["dairy"],
            row["protein"],
            row["breakfast"],
            row["processed_food"],
            row["sunlight"],
            row["diet_rating"],
            row["water"],
            row["nutrition_knowledge"],
            row["diet_barrier"],
            row["diagnosed"],
            row["diagnosed_type"],
            row["nutrition_source"],
            row["supplements"],
            row["awareness"]
        ])

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(
            max_length + 2,
            40
        )

    filename = "NutriCheck_Responses.xlsx"

    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True,
        download_name=filename
    )


# --------------------------------------------------
# LOGOUT
# --------------------------------------------------

@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("home"))


# --------------------------------------------------
# NUTRITION GUIDE
# --------------------------------------------------

@app.route("/nutrition-guide")
def nutrition_guide():
    return render_template("nutrition_guide.html")


# --------------------------------------------------
# START APP
# --------------------------------------------------

init_db()

if __name__ == "__main__":
    app.run(debug=True)