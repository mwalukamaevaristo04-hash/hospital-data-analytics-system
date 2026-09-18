from io import BytesIO, StringIO
import os

from flask import send_file

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak,
    Table,
    TableStyle
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

import csv

from flask import (
    Flask,
    render_template,
    request,
    redirect,
    session,
    url_for,
    jsonify
)

from werkzeug.security import (
    generate_password_hash,
    check_password_hash
)

import sqlite3

from functools import wraps

from datetime import (
    date,
    datetime,
    timedelta
)


app = Flask(__name__)

app.secret_key = "hospital_secret_key_2026"

# Basic secure session-cookie settings.
# HTTPS should be enabled before setting SESSION_COOKIE_SECURE=True.

app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

# =========================================================
# DATABASE CONNECTION
# =========================================================

def get_db_connection():

    database_path = os.path.join(
        app.root_path,
        "hospital.db"
    )

    conn = sqlite3.connect(
        database_path
    )

    conn.row_factory = sqlite3.Row

    return conn
# ==================================================
# INITIALIZE DATABASE
# =================================================

def init_db():

    conn = get_db_connection()

    # =====================================================
    # USERS TABLE
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS users (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            username TEXT NOT NULL UNIQUE,

            password TEXT NOT NULL,

            role TEXT NOT NULL,

            doctor_id INTEGER,

            patient_id INTEGER

        )
    """)

    # =====================================================
    # USERS DATABASE MIGRATION
    # =====================================================

    columns = conn.execute(
        "PRAGMA table_info(users)"
    ).fetchall()

    column_names = [
        column["name"]
        for column in columns
    ]

    if "doctor_id" not in column_names:

        conn.execute("""
            ALTER TABLE users
            ADD COLUMN doctor_id INTEGER
        """)

    if "patient_id" not in column_names:

        conn.execute("""
            ALTER TABLE users
            ADD COLUMN patient_id INTEGER
        """)

    # =====================================================
    # PATIENTS TABLE
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS patients (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            full_name TEXT NOT NULL,

            age INTEGER NOT NULL,

            gender TEXT NOT NULL,

            phone TEXT,

            address TEXT

        )
    """)

    # =====================================================
    # DOCTORS TABLE
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS doctors (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            full_name TEXT NOT NULL,

            gender TEXT NOT NULL,

            specialization TEXT NOT NULL,

            phone TEXT,

            email TEXT,

            department TEXT

        )
    """)

    # =====================================================
    # APPOINTMENTS TABLE
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS appointments (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            patient_id INTEGER NOT NULL,

            doctor_id INTEGER NOT NULL,

            appointment_date TEXT NOT NULL,

            appointment_time TEXT NOT NULL,

            reason TEXT,

            status TEXT NOT NULL,

            FOREIGN KEY (patient_id)
                REFERENCES patients(id),

            FOREIGN KEY (doctor_id)
                REFERENCES doctors(id)

        )
    """)

    # =====================================================
    # MEDICAL RECORDS TABLE
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS medical_records (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            patient_id INTEGER NOT NULL,

            doctor_id INTEGER NOT NULL,

            appointment_id INTEGER,

            visit_date TEXT NOT NULL,

            chief_complaint TEXT,

            diagnosis TEXT,

            symptoms TEXT,

            treatment TEXT,

            prescription TEXT,

            notes TEXT,

            FOREIGN KEY (patient_id)
                REFERENCES patients(id),

            FOREIGN KEY (doctor_id)
                REFERENCES doctors(id),

            FOREIGN KEY (appointment_id)
                REFERENCES appointments(id)

        )
    """)

    # =====================================================
    # PHARMACY - MEDICINES
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS medicines (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            medicine_name TEXT NOT NULL,

            category TEXT,

            unit TEXT,

            quantity INTEGER NOT NULL DEFAULT 0,

            reorder_level INTEGER NOT NULL DEFAULT 10,

            unit_price REAL NOT NULL DEFAULT 0,

            expiry_date TEXT,

            supplier TEXT,

            batch_number TEXT,

            description TEXT

        )
    """)

    # =====================================================
    # PHARMACY - PRESCRIPTIONS
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS prescriptions (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            patient_id INTEGER NOT NULL,

            doctor_id INTEGER NOT NULL,

            medical_record_id INTEGER,

            medicine_id INTEGER NOT NULL,

            dosage TEXT,

            frequency TEXT,

            duration TEXT,

            quantity INTEGER NOT NULL DEFAULT 1,

            instructions TEXT,

            status TEXT NOT NULL DEFAULT 'Pending',

            prescribed_date TEXT NOT NULL,

            dispensed_date TEXT,

            dispensed_by TEXT,

            FOREIGN KEY (patient_id)
                REFERENCES patients(id),

            FOREIGN KEY (doctor_id)
                REFERENCES doctors(id),

            FOREIGN KEY (medical_record_id)
                REFERENCES medical_records(id),

            FOREIGN KEY (medicine_id)
                REFERENCES medicines(id)

        )
    """)

    # =====================================================
    # LABORATORY - TEST TYPES
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS laboratory_tests (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            test_name TEXT NOT NULL,

            category TEXT,

            description TEXT,

            normal_range TEXT,

            unit TEXT,

            price REAL NOT NULL DEFAULT 0,

            recommended_specialty TEXT

        )
    """)

    # =====================================================
    # LABORATORY - TEST REQUESTS
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS laboratory_requests (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            patient_id INTEGER NOT NULL,

            test_id INTEGER NOT NULL,

            request_date TEXT NOT NULL,

            status TEXT NOT NULL DEFAULT 'Pending',

            notes TEXT,

            created_by TEXT,

            FOREIGN KEY (patient_id)
                REFERENCES patients(id),

            FOREIGN KEY (test_id)
                REFERENCES laboratory_tests(id)

        )
    """)

    # =====================================================
    # LABORATORY - RESULTS
    # =====================================================

    conn.execute("""
        CREATE TABLE IF NOT EXISTS laboratory_results (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            request_id INTEGER NOT NULL,

            result_value TEXT,

            finding TEXT,

            interpretation TEXT,

            result_date TEXT NOT NULL,

            performed_by TEXT,

            recommended_specialty TEXT,

            doctor_id INTEGER,

            FOREIGN KEY (request_id)
                REFERENCES laboratory_requests(id),

            FOREIGN KEY (doctor_id)
                REFERENCES doctors(id)

        )
    """)

    # =====================================================
    # LABORATORY RESULT MIGRATION
    # =====================================================

    result_columns = conn.execute(
        "PRAGMA table_info(laboratory_results)"
    ).fetchall()

    result_column_names = [
        column["name"]
        for column in result_columns
    ]

    if "doctor_id" not in result_column_names:

        conn.execute("""
            ALTER TABLE laboratory_results
            ADD COLUMN doctor_id INTEGER
        """)

    # =====================================================
    # DEFAULT LABORATORY TESTS
    # =====================================================

    laboratory_test_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_tests
    """).fetchone()["count"]

    if laboratory_test_count == 0:

        default_tests = [

            (
                "Complete Blood Count (CBC)",
                "Hematology",
                "Blood test used to evaluate blood cells.",
                "Varies by parameter",
                "",
                10000,
                "Hematology"
            ),

            (
                "Malaria Test",
                "Parasitology",
                "Test for malaria parasites.",
                "Negative",
                "",
                5000,
                "Internal Medicine"
            ),

            (
                "Blood Glucose",
                "Biochemistry",
                "Measurement of blood glucose level.",
                "70-99 mg/dL fasting",
                "mg/dL",
                5000,
                "Internal Medicine"
            ),

            (
                "Urinalysis",
                "Clinical Laboratory",
                "Examination of urine for abnormalities.",
                "Normal",
                "",
                5000,
                "Internal Medicine"
            ),

            (
                "X-Ray - Bone",
                "Diagnostic Imaging",
                "X-Ray examination of bones.",
                "Report based",
                "",
                15000,
                "Orthopedic"
            ),

            (
                "Chest X-Ray",
                "Diagnostic Imaging",
                "Imaging examination of the chest.",
                "Report based",
                "",
                20000,
                "Pulmonology"
            ),

            (
                "Eye Imaging",
                "Diagnostic Imaging",
                "Diagnostic imaging related to the eye.",
                "Report based",
                "",
                20000,
                "Ophthalmology"
            )

        ]

        conn.executemany("""
            INSERT INTO laboratory_tests
            (
                test_name,
                category,
                description,
                normal_range,
                unit,
                price,
                recommended_specialty
            )

            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, default_tests)

    # =====================================================
    # DEFAULT ADMIN
    # =====================================================

    admin = conn.execute(
        "SELECT * FROM users WHERE username = ?",
        ("admin",)
    ).fetchone()

    if admin is None:

        conn.execute("""
            INSERT INTO users
            (
                username,
                password,
                role,
                doctor_id,
                patient_id
            )

            VALUES (?, ?, ?, ?, ?)
        """, (
            "admin",
            generate_password_hash("admin123"),
            "Admin",
            None,
            None
        ))

    # =====================================================
    # PASSWORD HASH MIGRATION
    # =====================================================

    all_users = conn.execute("""
        SELECT id, password
        FROM users
    """).fetchall()

    for existing_user in all_users:

        stored_password = existing_user["password"]

        if (
            stored_password
            and not stored_password.startswith(
                ("scrypt:", "pbkdf2:")
            )
        ):

            conn.execute("""
                UPDATE users

                SET password = ?

                WHERE id = ?
            """, (
                generate_password_hash(
                    stored_password
                ),
                existing_user["id"]
            ))

    # =====================================================
    # OLD APPOINTMENT STATUS MIGRATION
    # =====================================================

    conn.execute("""
        UPDATE appointments

        SET status = 'Accepted'

        WHERE status = 'Confirmed'
    """)

    # =====================================================
    # SAVE DATABASE
    # =====================================================

    conn.commit()

    conn.close()


# =========================================================
# LOGIN REQUIRED
# =========================================================

def login_required(function):

    @wraps(function)
    def decorated_function(*args, **kwargs):

        if "user_id" not in session:

            return redirect(url_for("login"))

        return function(*args, **kwargs)

    return decorated_function


# =========================================================
# ROLE REQUIRED
# =========================================================

def role_required(*allowed_roles):

    def decorator(function):

        @wraps(function)
        def decorated_function(*args, **kwargs):

            if "user_id" not in session:

                return redirect(url_for("login"))

            current_role = session.get("role")

            if current_role not in allowed_roles:

                return """
                    <h1>Access Denied</h1>

                    <p>
                        You do not have permission
                        to access this page.
                    </p>

                    <p>
                        Your current role:
                        <strong>{}</strong>
                    </p>

                    <a href="/">
                        Back to Dashboard
                    </a>
                """.format(current_role)

            return function(*args, **kwargs)

        return decorated_function

    return decorator


# =========================================================
# ADMIN ONLY
# =========================================================

def admin_required(function):

    @wraps(function)
    def decorated_function(*args, **kwargs):

        if "user_id" not in session:

            return redirect(url_for("login"))

        if session.get("role") != "Admin":

            return """
                <h1>Access Denied</h1>

                <p>
                    Only administrators can access
                    User Management.
                </p>

                <a href="/">
                    Back to Dashboard
                </a>
            """

        return function(*args, **kwargs)

    return decorated_function


# =========================================================
# LOGIN
# =========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"].strip()

        password = request.form["password"]

        conn = get_db_connection()

        user = conn.execute("""
            SELECT *
            FROM users
            WHERE username = ?
        """, (
            username,
        )).fetchone()

        if user and check_password_hash(user["password"], password):

            conn.close()

            session.clear()

            session["user_id"] = user["id"]

            session["username"] = user["username"]

            session["role"] = user["role"]

            session["doctor_id"] = user["doctor_id"]

            session["patient_id"] = user["patient_id"]

            if user["role"] == "Patient":
                return redirect(url_for("patient_dashboard"))

            return redirect(url_for("home"))

        conn.close()

        return render_template(
            "login.html",
            error="Invalid username or password."
        )

    return render_template("login.html")



# =========================================================
# PATIENT SELF REGISTRATION
# =========================================================

@app.route("/patient_register", methods=["GET", "POST"])
def patient_register():

    if request.method == "POST":

        full_name = request.form["full_name"].strip()
        age = request.form["age"]
        gender = request.form["gender"]
        phone = request.form["phone"].strip()
        address = request.form.get("address", "").strip()
        username = request.form["username"].strip()
        password = request.form["password"]
        confirm_password = request.form["confirm_password"]

        if password != confirm_password:
            return render_template(
                "patient_register.html",
                error="Passwords do not match."
            )

        if len(password) < 6:
            return render_template(
                "patient_register.html",
                error="Password must contain at least 6 characters."
            )

        conn = get_db_connection()

        existing_user = conn.execute("""
            SELECT id
            FROM users
            WHERE username = ?
        """, (username,)).fetchone()

        if existing_user:
            conn.close()
            return render_template(
                "patient_register.html",
                error="Username already exists. Please choose another username."
            )

        cursor = conn.execute("""
            INSERT INTO patients
            (
                full_name,
                age,
                gender,
                phone,
                address
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            full_name,
            age,
            gender,
            phone,
            address
        ))

        patient_id = cursor.lastrowid

        conn.execute("""
            INSERT INTO users
            (
                username,
                password,
                role,
                doctor_id,
                patient_id
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            username,
            generate_password_hash(password),
            "Patient",
            None,
            patient_id
        ))

        conn.commit()
        conn.close()

        return redirect(url_for("login", registered="1"))

    return render_template("patient_register.html")


# =========================================================
# PATIENT DASHBOARD
# =========================================================

@app.route("/patient_dashboard")
@role_required("Patient")
def patient_dashboard():

    patient_id = session.get("patient_id")

    if not patient_id:
        return """
            <h1>Patient Profile Not Connected</h1>
            <p>Your account is not connected to a Patient Profile.</p>
            <a href="/logout">Logout</a>
        """

    conn = get_db_connection()

    patient = conn.execute("""
        SELECT *
        FROM patients
        WHERE id = ?
    """, (patient_id,)).fetchone()

    appointment_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE patient_id = ?
    """, (patient_id,)).fetchone()["count"]

    pending_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE patient_id = ?
        AND status = 'Pending'
    """, (patient_id,)).fetchone()["count"]

    lab_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE patient_id = ?
    """, (patient_id,)).fetchone()["count"]

    completed_labs = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE patient_id = ?
        AND status = 'Completed'
    """, (patient_id,)).fetchone()["count"]

    conn.close()

    return render_template(
        "patient_dashboard.html",
        patient=patient,
        appointment_count=appointment_count,
        pending_appointments=pending_appointments,
        lab_requests=lab_requests,
        completed_labs=completed_labs,
        username=session.get("username"),
        role=session.get("role")
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    session.clear()

    return redirect(url_for("login"))


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/")
@login_required
def home():

    conn = get_db_connection()

    patient_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM patients
    """).fetchone()["count"]

    doctor_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM doctors
    """).fetchone()["count"]

    appointment_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
    """).fetchone()["count"]

    medicine_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
    """).fetchone()["count"]

    low_stock_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
        WHERE quantity <= reorder_level
    """).fetchone()["count"]

    laboratory_test_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_tests
    """).fetchone()["count"]

    pending_lab_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Pending'
    """).fetchone()["count"]

    processing_lab_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Processing'
    """).fetchone()["count"]

    completed_lab_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Completed'
    """).fetchone()["count"]

    conn.close()

    return render_template(
        "index.html",

        patient_count=patient_count,

        doctor_count=doctor_count,

        appointment_count=appointment_count,

        medicine_count=medicine_count,

        low_stock_count=low_stock_count,

        laboratory_test_count=laboratory_test_count,

        pending_lab_count=pending_lab_count,

        processing_lab_count=processing_lab_count,

        completed_lab_count=completed_lab_count,

        username=session["username"],

        role=session["role"]
    )


# =========================================================
# USER MANAGEMENT
# ADMIN ONLY
# =========================================================

@app.route("/users", methods=["GET", "POST"])
@admin_required
def users():

    conn = get_db_connection()

    if request.method == "POST":

        username = request.form["username"]

        password = request.form.get("password", "").strip()

        if not password:

            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()

            patients = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            users_list = conn.execute("""
                SELECT *
                FROM users
                ORDER BY id
            """).fetchall()

            conn.close()

            return render_template(
                "users.html",
                users=users_list,
                doctors=doctors,
                patients=patients,
                error="Password is required for a new user."
            )

        hashed_password = generate_password_hash(password)

        role = request.form["role"]

        allowed_user_roles = [
            "Admin",
            "Doctor",
            "Receptionist",
            "Data Analyst",
            "Patient",
            "Laboratorian",
            "Pharmacist"
        ]

        if role not in allowed_user_roles:

            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()

            patients = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            users_list = conn.execute("""
                SELECT *
                FROM users
                ORDER BY id
            """).fetchall()

            conn.close()

            return render_template(
                "users.html",
                users=users_list,
                doctors=doctors,
                patients=patients,
                error="Invalid user role selected."
            )

        doctor_id = None

        patient_id = None

        if role == "Doctor":

            doctor_id = request.form.get("doctor_id")

            if not doctor_id:

                doctors = conn.execute("""
                    SELECT *
                    FROM doctors
                    ORDER BY full_name
                """).fetchall()

                patients = conn.execute("""
                    SELECT *
                    FROM patients
                    ORDER BY full_name
                """).fetchall()

                users_list = conn.execute("""
                    SELECT *
                    FROM users
                    ORDER BY id
                """).fetchall()

                conn.close()

                return render_template(
                    "users.html",
                    users=users_list,
                    doctors=doctors,
                    patients=patients,
                    error="Please select a Doctor Profile."
                )

        elif role == "Patient":

            patient_id = request.form.get("patient_id")

            if not patient_id:

                doctors = conn.execute("""
                    SELECT *
                    FROM doctors
                    ORDER BY full_name
                """).fetchall()

                patients = conn.execute("""
                    SELECT *
                    FROM patients
                    ORDER BY full_name
                """).fetchall()

                users_list = conn.execute("""
                    SELECT *
                    FROM users
                    ORDER BY id
                """).fetchall()

                conn.close()

                return render_template(
                    "users.html",
                    users=users_list,
                    doctors=doctors,
                    patients=patients,
                    error="Please select a Patient Profile."
                )

        try:

            conn.execute("""
                INSERT INTO users
                (
                    username,
                    password,
                    role,
                    doctor_id,
                    patient_id
                )

                VALUES (?, ?, ?, ?, ?)
            """, (
                username,
                hashed_password,
                role,
                doctor_id,
                patient_id
            ))

            conn.commit()

        except sqlite3.IntegrityError:

            users_list = conn.execute("""
                SELECT *
                FROM users
                ORDER BY id
            """).fetchall()

            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()

            patients = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            conn.close()

            return render_template(
                "users.html",
                users=users_list,
                doctors=doctors,
                patients=patients,
                error="Username already exists."
            )

    users_list = conn.execute("""
        SELECT

            users.*,

            doctors.full_name AS doctor_name,

            patients.full_name AS patient_name

        FROM users

        LEFT JOIN doctors
        ON users.doctor_id = doctors.id

        LEFT JOIN patients
        ON users.patient_id = patients.id

        ORDER BY users.id
    """).fetchall()

    doctors = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY full_name
    """).fetchall()

    patients = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "users.html",
        users=users_list,
        doctors=doctors,
        patients=patients
    )


# =========================================================
# EDIT USER
# ADMIN ONLY
# =========================================================

@app.route(
    "/edit_user/<int:user_id>",
    methods=["GET", "POST"]
)
@admin_required
def edit_user(user_id):

    conn = get_db_connection()

    user = conn.execute("""
        SELECT *
        FROM users
        WHERE id = ?
    """, (user_id,)).fetchone()

    if user is None:

        conn.close()

        return "User not found"

    if request.method == "POST":

        username = request.form["username"]

        password = request.form.get("password", "").strip()

        if not password:

            hashed_password = user["password"]

        else:

            hashed_password = generate_password_hash(password)

        role = request.form["role"]

        allowed_user_roles = [
            "Admin",
            "Doctor",
            "Receptionist",
            "Data Analyst",
            "Patient",
            "Laboratorian",
            "Pharmacist"
        ]

        if role not in allowed_user_roles:

            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()

            patients = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            conn.close()

            return render_template(
                "edit_user.html",
                user=user,
                doctors=doctors,
                patients=patients,
                error="Invalid user role selected."
            )

        doctor_id = None

        patient_id = None

        if role == "Doctor":

            doctor_id = request.form.get("doctor_id")

            if not doctor_id:

                doctors = conn.execute("""
                    SELECT *
                    FROM doctors
                    ORDER BY full_name
                """).fetchall()

                patients = conn.execute("""
                    SELECT *
                    FROM patients
                    ORDER BY full_name
                """).fetchall()

                conn.close()

                return render_template(
                    "edit_user.html",
                    user=user,
                    doctors=doctors,
                    patients=patients,
                    error="Please select a Doctor Profile."
                )

        elif role == "Patient":

            patient_id = request.form.get("patient_id")

            if not patient_id:

                doctors = conn.execute("""
                    SELECT *
                    FROM doctors
                    ORDER BY full_name
                """).fetchall()

                patients = conn.execute("""
                    SELECT *
                    FROM patients
                    ORDER BY full_name
                """).fetchall()

                conn.close()

                return render_template(
                    "edit_user.html",
                    user=user,
                    doctors=doctors,
                    patients=patients,
                    error="Please select a Patient Profile."
                )

        try:

            conn.execute("""
                UPDATE users

                SET

                    username = ?,

                    password = ?,

                    role = ?,

                    doctor_id = ?,

                    patient_id = ?

                WHERE id = ?
            """, (
                username,
                hashed_password,
                role,
                doctor_id,
                patient_id,
                user_id
            ))

            conn.commit()

        except sqlite3.IntegrityError:

            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()

            patients = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            conn.close()

            return render_template(
                "edit_user.html",
                user=user,
                doctors=doctors,
                patients=patients,
                error="Username already exists."
            )

        conn.close()

        return redirect(url_for("users"))

    doctors = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY full_name
    """).fetchall()

    patients = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "edit_user.html",
        user=user,
        doctors=doctors,
        patients=patients
    )


# =========================================================
# DELETE USER
# ADMIN ONLY
# =========================================================

@app.route(
    "/delete_user/<int:user_id>",
    methods=["POST"]
)
@admin_required
def delete_user(user_id):

    conn = get_db_connection()

    user = conn.execute("""
        SELECT *
        FROM users
        WHERE id = ?
    """, (user_id,)).fetchone()

    if user:

        if user["id"] != session["user_id"]:

            conn.execute("""
                DELETE FROM users
                WHERE id = ?
            """, (user_id,))

            conn.commit()

    conn.close()

    return redirect(url_for("users"))


# =========================================================
# PATIENTS - VIEW
# =========================================================
@app.route("/patients")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def patients():

    conn = get_db_connection()

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "patients.html",
        patients=patients_list,
        username=session["username"],
        role=session["role"]
    )

# =========================================================
# ADD PATIENT
# ADMIN + RECEPTIONIST
# =========================================================

@app.route(
    "/add_patient",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Receptionist"
)
def add_patient():

    full_name = request.form["full_name"].strip()

    age = request.form["age"]

    gender = request.form["gender"]

    phone = request.form["phone"].strip()

    address = request.form["address"].strip()

    username = request.form.get("username", "").strip()

    password = request.form.get("password", "").strip()

    conn = get_db_connection()

    if not username or not password:

        patients_list = conn.execute("""
            SELECT *
            FROM patients
            ORDER BY id DESC
        """).fetchall()

        conn.close()

        return render_template(
            "patients.html",
            patients=patients_list,
            username=session.get("username"),
            role=session.get("role"),
            error="Patient username and password are required."
        )

    existing_user = conn.execute("""
        SELECT id
        FROM users
        WHERE username = ?
    """, (username,)).fetchone()

    if existing_user:

        patients_list = conn.execute("""
            SELECT *
            FROM patients
            ORDER BY id DESC
        """).fetchall()

        conn.close()

        return render_template(
            "patients.html",
            patients=patients_list,
            username=session.get("username"),
            role=session.get("role"),
            error="Username already exists. Please choose another username."
        )

    cursor = conn.execute("""
        INSERT INTO patients
        (
            full_name,
            age,
            gender,
            phone,
            address
        )

        VALUES (?, ?, ?, ?, ?)
    """, (
        full_name,
        age,
        gender,
        phone,
        address
    ))

    patient_id = cursor.lastrowid

    conn.execute("""
        INSERT INTO users
        (
            username,
            password,
            role,
            doctor_id,
            patient_id
        )

        VALUES (?, ?, ?, ?, ?)
    """, (
        username,
        generate_password_hash(password),
        "Patient",
        None,
        patient_id
    ))

    conn.commit()

    conn.close()

    return redirect(url_for("patients"))


# =========================================================
# EDIT PATIENT
# =========================================================

@app.route(
    "/edit_patient/<int:patient_id>",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Receptionist"
)
def edit_patient(patient_id):

    conn = get_db_connection()

    patient = conn.execute("""
        SELECT *
        FROM patients
        WHERE id = ?
    """, (patient_id,)).fetchone()

    if patient is None:

        conn.close()

        return "Patient not found"

    if request.method == "POST":

        full_name = request.form["full_name"]

        age = request.form["age"]

        gender = request.form["gender"]

        phone = request.form["phone"]

        address = request.form["address"]

        conn.execute("""
            UPDATE patients

            SET

                full_name = ?,

                age = ?,

                gender = ?,

                phone = ?,

                address = ?

            WHERE id = ?
        """, (
            full_name,
            age,
            gender,
            phone,
            address,
            patient_id
        ))

        conn.commit()

        conn.close()

        return redirect(url_for("patients"))

    conn.close()

    return render_template(
        "edit_patient.html",
        patient=patient
    )


# =========================================================
# DELETE PATIENT
# =========================================================

@app.route(
    "/delete_patient/<int:patient_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Receptionist"
)
def delete_patient(patient_id):

    conn = get_db_connection()

    conn.execute("""
        DELETE FROM patients
        WHERE id = ?
    """, (patient_id,))

    conn.commit()

    conn.close()

    return redirect(url_for("patients"))


# =========================================================
# DOCTORS - VIEW
# =========================================================
@app.route("/doctors")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def doctors():

    conn = get_db_connection()

    doctors_list = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY id DESC
    """).fetchall()

    conn.close()

    return render_template(
        "doctors.html",
        doctors=doctors_list,
        username=session["username"],
        role=session["role"]
    )
# =========================================================
# ADD DOCTOR
# ADMIN ONLY
# =========================================================

@app.route(
    "/add_doctor",
    methods=["POST"]
)
@admin_required
def add_doctor():

    full_name = request.form["full_name"]

    gender = request.form["gender"]

    specialization = request.form["specialization"]

    phone = request.form["phone"]

    email = request.form["email"]

    department = request.form["department"]

    conn = get_db_connection()

    conn.execute("""
        INSERT INTO doctors
        (
            full_name,
            gender,
            specialization,
            phone,
            email,
            department
        )

        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        full_name,
        gender,
        specialization,
        phone,
        email,
        department
    ))

    conn.commit()

    conn.close()

    return redirect(url_for("doctors"))


# =========================================================
# EDIT DOCTOR
# =========================================================

@app.route(
    "/edit_doctor/<int:doctor_id>",
    methods=["GET", "POST"]
)
@admin_required
def edit_doctor(doctor_id):

    conn = get_db_connection()

    doctor = conn.execute("""
        SELECT *
        FROM doctors
        WHERE id = ?
    """, (doctor_id,)).fetchone()

    if doctor is None:

        conn.close()

        return "Doctor not found"

    if request.method == "POST":

        full_name = request.form["full_name"]

        gender = request.form["gender"]

        specialization = request.form["specialization"]

        phone = request.form["phone"]

        email = request.form["email"]

        department = request.form["department"]

        conn.execute("""
            UPDATE doctors

            SET

                full_name = ?,

                gender = ?,

                specialization = ?,

                phone = ?,

                email = ?,

                department = ?

            WHERE id = ?
        """, (
            full_name,
            gender,
            specialization,
            phone,
            email,
            department,
            doctor_id
        ))

        conn.commit()

        conn.close()

        return redirect(url_for("doctors"))

    conn.close()

    return render_template(
        "edit_doctor.html",
        doctor=doctor
    )


# =========================================================
# DELETE DOCTOR
# =========================================================

@app.route(
    "/delete_doctor/<int:doctor_id>",
    methods=["POST"]
)
@admin_required
def delete_doctor(doctor_id):

    conn = get_db_connection()

    conn.execute("""
        DELETE FROM doctors
        WHERE id = ?
    """, (doctor_id,))

    conn.commit()

    conn.close()

    return redirect(url_for("doctors"))


# =========================================================
# APPOINTMENTS
# =========================================================

@app.route("/appointments")
@login_required
def appointments():

    conn = get_db_connection()

    current_role = session.get("role")

    # =====================================================
    # APPOINTMENT LIST
    # =====================================================

    appointments_list = conn.execute("""
        SELECT

            appointments.id,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name,

            appointments.appointment_date,

            appointments.appointment_time,

            appointments.reason,

            appointments.status

        FROM appointments

        JOIN patients
        ON appointments.patient_id = patients.id

        JOIN doctors
        ON appointments.doctor_id = doctors.id

        ORDER BY
            appointments.appointment_date DESC,
            appointments.appointment_time DESC
    """).fetchall()

    # =====================================================
    # APPOINTMENT SUMMARY
    # =====================================================

    total_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
    """).fetchone()["count"]

    pending_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE status = 'Pending'
    """).fetchone()["count"]

    accepted_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE status = 'Accepted'
    """).fetchone()["count"]

    rejected_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE status = 'Rejected'
    """).fetchone()["count"]

    completed_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE status = 'Completed'
    """).fetchone()["count"]

    today_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE appointment_date = date('now')
    """).fetchone()["count"]

    # =====================================================
    # PATIENTS
    # =====================================================

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    # =====================================================
    # DOCTORS
    # =====================================================

    doctors_list = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "appointments.html",

        appointments=appointments_list,

        patients=patients_list,

        doctors=doctors_list,

        total_appointments=total_appointments,

        pending_appointments=pending_appointments,

        accepted_appointments=accepted_appointments,

        rejected_appointments=rejected_appointments,

        completed_appointments=completed_appointments,

        today_appointments=today_appointments,

        username=session.get("username"),

        role=current_role
    )


# =========================================================
# PATIENT REQUEST APPOINTMENT
# =========================================================

@app.route("/patient_request_appointment", methods=["GET", "POST"])
@role_required("Patient")
def patient_request_appointment():

    patient_id = session.get("patient_id")

    if not patient_id:
        return """
            <h1>Patient Profile Not Connected</h1>
            <p>Your account is not connected to a Patient Profile.</p>
            <a href="/patient_dashboard">Back to Patient Dashboard</a>
        """

    conn = get_db_connection()

    if request.method == "POST":

        doctor_id = request.form["doctor_id"]
        appointment_date = request.form["appointment_date"]
        appointment_time = request.form["appointment_time"]
        reason = request.form.get("reason", "").strip()

        doctor = conn.execute("""
            SELECT id
            FROM doctors
            WHERE id = ?
        """, (doctor_id,)).fetchone()

        if doctor is None:
            doctors = conn.execute("""
                SELECT *
                FROM doctors
                ORDER BY full_name
            """).fetchall()
            conn.close()
            return render_template(
                "patient_request_appointment.html",
                doctors=doctors,
                error="Selected doctor was not found."
            )

        conn.execute("""
            INSERT INTO appointments
            (
                patient_id,
                doctor_id,
                appointment_date,
                appointment_time,
                reason,
                status
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            patient_id,
            doctor_id,
            appointment_date,
            appointment_time,
            reason,
            "Pending"
        ))

        conn.commit()
        conn.close()

        return redirect(url_for("patient_dashboard"))

    doctors = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "patient_request_appointment.html",
        doctors=doctors
    )


# =========================================================
# ADD APPOINTMENT
# =========================================================

@app.route(
    "/add_appointment",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Receptionist"
)
def add_appointment():

    patient_id = request.form["patient_id"]

    doctor_id = request.form["doctor_id"]

    appointment_date = request.form["appointment_date"]

    appointment_time = request.form["appointment_time"]

    reason = request.form["reason"]

    conn = get_db_connection()

    conn.execute("""
        INSERT INTO appointments
        (
            patient_id,
            doctor_id,
            appointment_date,
            appointment_time,
            reason,
            status
        )

        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        patient_id,
        doctor_id,
        appointment_date,
        appointment_time,
        reason,
        "Pending"
    ))

    conn.commit()

    conn.close()

    return redirect(url_for("appointments"))
# =========================================================
# CANCEL APPOINTMENT
# =========================================================

@app.route(
    "/cancel_appointment/<int:appointment_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Receptionist"
)
def cancel_appointment(appointment_id):

    conn = get_db_connection()

    appointment = conn.execute("""
        SELECT *
        FROM appointments
        WHERE id = ?
    """, (
        appointment_id,
    )).fetchone()

    if appointment is None:

        conn.close()

        return """
            <h1>Appointment Not Found</h1>

            <p>
                The selected appointment
                does not exist.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if appointment["status"] in [
        "Completed",
        "Cancelled",
        "Rejected"
    ]:

        conn.close()

        return redirect(
            url_for("appointments")
        )

    conn.execute("""
        UPDATE appointments
        SET status = 'Cancelled'
        WHERE id = ?
    """, (
        appointment_id,
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("appointments")
    )



# =========================================================
# ACCEPT APPOINTMENT
# DOCTOR ONLY
# =========================================================

@app.route(
    "/accept_appointment/<int:appointment_id>",
    methods=["POST"]
)
@role_required(
    "Doctor"
)
def accept_appointment(appointment_id):

    conn = get_db_connection()

    appointment = conn.execute("""
        SELECT *
        FROM appointments
        WHERE id = ?
    """, (
        appointment_id,
    )).fetchone()

    if appointment is None:

        conn.close()

        return """
            <h1>Appointment Not Found</h1>

            <p>
                The selected appointment
                does not exist.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if str(appointment["doctor_id"]) != str(
        session.get("doctor_id")
    ):

        conn.close()

        return """
            <h1>Access Denied</h1>

            <p>
                You can only accept appointments
                assigned to your Doctor Profile.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if appointment["status"] != "Pending":

        conn.close()

        return redirect(
            url_for("appointments")
        )

    conn.execute("""
        UPDATE appointments
        SET status = 'Accepted'
        WHERE id = ?
    """, (
        appointment_id,
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("appointments")
    )


# =========================================================
# REJECT APPOINTMENT
# DOCTOR ONLY
# =========================================================

@app.route(
    "/reject_appointment/<int:appointment_id>",
    methods=["POST"]
)
@role_required(
    "Doctor"
)
def reject_appointment(appointment_id):

    conn = get_db_connection()

    appointment = conn.execute("""
        SELECT *
        FROM appointments
        WHERE id = ?
    """, (
        appointment_id,
    )).fetchone()

    if appointment is None:

        conn.close()

        return """
            <h1>Appointment Not Found</h1>

            <p>
                The selected appointment
                does not exist.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if str(appointment["doctor_id"]) != str(
        session.get("doctor_id")
    ):

        conn.close()

        return """
            <h1>Access Denied</h1>

            <p>
                You can only reject appointments
                assigned to your Doctor Profile.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if appointment["status"] != "Pending":

        conn.close()

        return redirect(
            url_for("appointments")
        )

    conn.execute("""
        UPDATE appointments
        SET status = 'Rejected'
        WHERE id = ?
    """, (
        appointment_id,
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("appointments")
    )


# =========================================================
# COMPLETE APPOINTMENT
# DOCTOR ONLY
# =========================================================

@app.route(
    "/complete_appointment/<int:appointment_id>",
    methods=["POST"]
)
@role_required(
    "Doctor"
)
def complete_appointment(appointment_id):

    conn = get_db_connection()

    appointment = conn.execute("""
        SELECT *
        FROM appointments
        WHERE id = ?
    """, (
        appointment_id,
    )).fetchone()

    if appointment is None:

        conn.close()

        return """
            <h1>Appointment Not Found</h1>

            <p>
                The selected appointment
                does not exist.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if str(appointment["doctor_id"]) != str(
        session.get("doctor_id")
    ):

        conn.close()

        return """
            <h1>Access Denied</h1>

            <p>
                You can only complete appointments
                assigned to your Doctor Profile.
            </p>

            <a href="/appointments">
                Back to Appointments
            </a>
        """

    if appointment["status"] != "Accepted":

        conn.close()

        return redirect(
            url_for("appointments")
        )

    conn.execute("""
        UPDATE appointments
        SET status = 'Completed'
        WHERE id = ?
    """, (
        appointment_id,
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("appointments")
    )

# =========================================================
# MEDICAL RECORDS HELPER
# =========================================================

def render_medical_records_page(filter_type="all"):

    conn = get_db_connection()

    current_role = session.get("role")

    conditions = []

    parameters = []

    if current_role == "Doctor":

        conditions.append(
            "medical_records.doctor_id = ?"
        )

        parameters.append(
            session.get("doctor_id")
        )

    if filter_type == "today":

        conditions.append(
            "medical_records.visit_date = ?"
        )

        parameters.append(
            date.today().isoformat()
        )

    elif filter_type == "diagnosed":

        conditions.append("""
            medical_records.diagnosis IS NOT NULL
            AND TRIM(medical_records.diagnosis) != ''
        """)

    elif filter_type == "linked":

        conditions.append(
            "medical_records.appointment_id IS NOT NULL"
        )

    where_clause = ""

    if conditions:

        where_clause = (
            "WHERE " +
            " AND ".join(conditions)
        )

    query = f"""
        SELECT

            medical_records.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name,

            appointments.appointment_date,

            appointments.appointment_time

        FROM medical_records

        JOIN patients
        ON medical_records.patient_id =
           patients.id

        JOIN doctors
        ON medical_records.doctor_id =
           doctors.id

        LEFT JOIN appointments
        ON medical_records.appointment_id =
           appointments.id

        {where_clause}

        ORDER BY
            medical_records.visit_date DESC,
            medical_records.id DESC
    """

    records = conn.execute(
        query,
        parameters
    ).fetchall()

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    if current_role == "Doctor":

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            WHERE id = ?
        """, (
            session.get("doctor_id"),
        )).fetchall()

        appointments_list = conn.execute("""
            SELECT

                appointments.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name

            FROM appointments

            JOIN patients
            ON appointments.patient_id =
               patients.id

            JOIN doctors
            ON appointments.doctor_id =
               doctors.id

            WHERE appointments.doctor_id = ?

            ORDER BY
                appointments.appointment_date DESC,
                appointments.appointment_time DESC
        """, (
            session.get("doctor_id"),
        )).fetchall()

    else:

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            ORDER BY full_name
        """).fetchall()

        appointments_list = conn.execute("""
            SELECT

                appointments.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name

            FROM appointments

            JOIN patients
            ON appointments.patient_id =
               patients.id

            JOIN doctors
            ON appointments.doctor_id =
               doctors.id

            ORDER BY
                appointments.appointment_date DESC,
                appointments.appointment_time DESC
        """).fetchall()

    current_date = date.today().isoformat()

    page_titles = {
        "all": "Medical Records",
        "today": "Today's Medical Records",
        "diagnosed": "Diagnosed Medical Records",
        "linked": "Medical Records with Linked Appointments"
    }

    active_filter = filter_type

    page_title = page_titles.get(
        filter_type,
        "Medical Records"
    )

    conn.close()

    return render_template(
        "medical_records.html",

        medical_records=records,

        patients=patients_list,

        doctors=doctors_list,

        appointments=appointments_list,

        username=session.get("username"),

        role=current_role,

        current_date=current_date,

        page_title=page_title,

        active_filter=active_filter
    )

# =========================================================
# MEDICAL RECORDS - ALL
# =========================================================

@app.route("/medical_records")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medical_records():

    return render_medical_records_page("all")


# =========================================================
# MEDICAL RECORDS - TODAY
# =========================================================

@app.route("/medical_records/today")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medical_records_today():

    return render_medical_records_page("today")


# =========================================================
# MEDICAL RECORDS - DIAGNOSED
# =========================================================

@app.route("/medical_records/diagnosed")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medical_records_diagnosed():

    return render_medical_records_page("diagnosed")


# =========================================================
# MEDICAL RECORDS - LINKED APPOINTMENTS
# =========================================================

@app.route("/medical_records/linked")
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medical_records_linked():

    return render_medical_records_page("linked")


# =========================================================
# ADD MEDICAL RECORD
# =========================================================

@app.route(
    "/add_medical_record",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Doctor"
)
def add_medical_record():

    patient_id = request.form.get("patient_id")

    doctor_id = request.form.get("doctor_id")

    appointment_id = request.form.get("appointment_id")

    visit_date = request.form.get("visit_date")

    chief_complaint = request.form.get("chief_complaint")

    diagnosis = request.form.get("diagnosis")

    symptoms = request.form.get("symptoms")

    treatment = request.form.get("treatment")

    prescription = request.form.get("prescription")

    notes = request.form.get("notes")

    if session.get("role") == "Doctor":

        doctor_id = session.get("doctor_id")

    if not patient_id or not doctor_id or not visit_date:

        return redirect(
            url_for("medical_records")
        )

    conn = get_db_connection()

    patient = conn.execute("""
        SELECT id
        FROM patients
        WHERE id = ?
    """, (patient_id,)).fetchone()

    if patient is None:

        conn.close()

        return "Patient not found"

    doctor = conn.execute("""
        SELECT id
        FROM doctors
        WHERE id = ?
    """, (doctor_id,)).fetchone()

    if doctor is None:

        conn.close()

        return "Doctor not found"

    if appointment_id:

        appointment = conn.execute("""
            SELECT *
            FROM appointments
            WHERE id = ?
        """, (appointment_id,)).fetchone()

        if appointment is None:

            conn.close()

            return "Appointment not found"

        if str(appointment["patient_id"]) != str(patient_id):

            conn.close()

            return "Invalid appointment for selected patient."

        if str(appointment["doctor_id"]) != str(doctor_id):

            conn.close()

            return "Invalid appointment for selected doctor."

    conn.execute("""
        INSERT INTO medical_records
        (
            patient_id,
            doctor_id,
            appointment_id,
            visit_date,
            chief_complaint,
            diagnosis,
            symptoms,
            treatment,
            prescription,
            notes
        )

        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        patient_id,
        doctor_id,
        appointment_id if appointment_id else None,
        visit_date,
        chief_complaint,
        diagnosis,
        symptoms,
        treatment,
        prescription,
        notes
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("medical_records")
    )

# =========================================================
# VIEW MEDICAL RECORD
# =========================================================

@app.route(
    "/medical_record/<int:record_id>"
)
@role_required(
    "Admin",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medical_record_details(record_id):

    conn = get_db_connection()

    record = conn.execute("""
        SELECT

            medical_records.*,

            patients.full_name AS patient_name,

            patients.age AS patient_age,

            patients.gender AS patient_gender,

            patients.phone AS patient_phone,

            doctors.full_name AS doctor_name,

            doctors.specialization AS doctor_specialization,

            doctors.department AS doctor_department,

            appointments.appointment_date,

            appointments.appointment_time,

            appointments.reason AS appointment_reason,

            appointments.status AS appointment_status

        FROM medical_records

        JOIN patients
        ON medical_records.patient_id = patients.id

        JOIN doctors
        ON medical_records.doctor_id = doctors.id

        LEFT JOIN appointments
        ON medical_records.appointment_id = appointments.id

        WHERE medical_records.id = ?
    """, (record_id,)).fetchone()

    if record is None:

        conn.close()

        return "Medical Record not found"

    if session.get("role") == "Doctor":

        if str(record["doctor_id"]) != str(
            session.get("doctor_id")
        ):

            conn.close()

            return """
                <h1>Access Denied</h1>

                <p>
                    You can only view medical records
                    belonging to your Doctor Profile.
                </p>

                <a href="/medical_records">
                    Back to Medical Records
                </a>
            """

    conn.close()

    return render_template(
        "medical_record_details.html",

        record=record,

        username=session.get("username"),

        role=session.get("role")
    )

# =========================================================
# EDIT MEDICAL RECORD
# =========================================================

@app.route(
    "/edit_medical_record/<int:record_id>",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Doctor"
)
def edit_medical_record(record_id):

    conn = get_db_connection()

    record = conn.execute("""
        SELECT *
        FROM medical_records
        WHERE id = ?
    """, (record_id,)).fetchone()

    if record is None:

        conn.close()

        return "Medical Record not found"

    if session.get("role") == "Doctor":

        if str(record["doctor_id"]) != str(
            session.get("doctor_id")
        ):

            conn.close()

            return """
                <h1>Access Denied</h1>

                <p>
                    You can only edit medical records
                    belonging to your Doctor Profile.
                </p>

                <a href="/medical_records">
                    Back to Medical Records
                </a>
            """

    if request.method == "POST":

        patient_id = request.form.get("patient_id")

        doctor_id = request.form.get("doctor_id")

        appointment_id = request.form.get("appointment_id")

        visit_date = request.form.get("visit_date")

        chief_complaint = request.form.get("chief_complaint")

        diagnosis = request.form.get("diagnosis")

        symptoms = request.form.get("symptoms")

        treatment = request.form.get("treatment")

        prescription = request.form.get("prescription")

        notes = request.form.get("notes")

        if session.get("role") == "Doctor":

            doctor_id = session.get("doctor_id")

        if not patient_id or not doctor_id or not visit_date:

            patients_list = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            if session.get("role") == "Doctor":

                doctors_list = conn.execute("""
                    SELECT *
                    FROM doctors
                    WHERE id = ?
                """, (
                    session.get("doctor_id"),
                )).fetchall()

                appointments_list = conn.execute("""
                    SELECT

                        appointments.*,

                        patients.full_name AS patient_name,

                        doctors.full_name AS doctor_name

                    FROM appointments

                    JOIN patients
                    ON appointments.patient_id =
                       patients.id

                    JOIN doctors
                    ON appointments.doctor_id =
                       doctors.id

                    WHERE appointments.doctor_id = ?

                    ORDER BY
                        appointments.appointment_date DESC,
                        appointments.appointment_time DESC
                """, (
                    session.get("doctor_id"),
                )).fetchall()

            else:

                doctors_list = conn.execute("""
                    SELECT *
                    FROM doctors
                    ORDER BY full_name
                """).fetchall()

                appointments_list = conn.execute("""
                    SELECT

                        appointments.*,

                        patients.full_name AS patient_name,

                        doctors.full_name AS doctor_name

                    FROM appointments

                    JOIN patients
                    ON appointments.patient_id =
                       patients.id

                    JOIN doctors
                    ON appointments.doctor_id =
                       doctors.id

                    ORDER BY
                        appointments.appointment_date DESC,
                        appointments.appointment_time DESC
                """).fetchall()

            conn.close()

            return render_template(
                "edit_medical_record.html",

                record=record,

                patients=patients_list,

                doctors=doctors_list,

                appointments=appointments_list,

                username=session.get("username"),

                role=session.get("role"),

                error="Patient, Doctor and Visit Date are required."
            )

        if appointment_id:

            appointment = conn.execute("""
                SELECT *
                FROM appointments
                WHERE id = ?
            """, (appointment_id,)).fetchone()

            if appointment is None:

                conn.close()

                return "Appointment not found"

            if str(appointment["patient_id"]) != str(patient_id):

                conn.close()

                return "Invalid appointment for selected patient."

            if str(appointment["doctor_id"]) != str(doctor_id):

                conn.close()

                return "Invalid appointment for selected doctor."

        conn.execute("""
            UPDATE medical_records

            SET

                patient_id = ?,

                doctor_id = ?,

                appointment_id = ?,

                visit_date = ?,

                chief_complaint = ?,

                diagnosis = ?,

                symptoms = ?,

                treatment = ?,

                prescription = ?,

                notes = ?

            WHERE id = ?
        """, (
            patient_id,
            doctor_id,
            appointment_id if appointment_id else None,
            visit_date,
            chief_complaint,
            diagnosis,
            symptoms,
            treatment,
            prescription,
            notes,
            record_id
        ))

        conn.commit()

        conn.close()

        return redirect(
           url_for("medical_record_details", record_id=record_id)
        )

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    if session.get("role") == "Doctor":

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            WHERE id = ?
        """, (
            session.get("doctor_id"),
        )).fetchall()

        appointments_list = conn.execute("""
            SELECT

                appointments.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name

            FROM appointments

            JOIN patients
            ON appointments.patient_id =
               patients.id

            JOIN doctors
            ON appointments.doctor_id =
               doctors.id

            WHERE appointments.doctor_id = ?

            ORDER BY
                appointments.appointment_date DESC,
                appointments.appointment_time DESC
        """, (
            session.get("doctor_id"),
        )).fetchall()

    else:

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            ORDER BY full_name
        """).fetchall()

        appointments_list = conn.execute("""
            SELECT

                appointments.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name

            FROM appointments

            JOIN patients
            ON appointments.patient_id =
               patients.id

            JOIN doctors
            ON appointments.doctor_id =
               doctors.id

            ORDER BY
                appointments.appointment_date DESC,
                appointments.appointment_time DESC
        """).fetchall()

    conn.close()

    return render_template(
        "edit_medical_record.html",

        record=record,

        patients=patients_list,

        doctors=doctors_list,

        appointments=appointments_list,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# DELETE MEDICAL RECORD
# ADMIN ONLY
# =========================================================

@app.route(
    "/delete_medical_record/<int:record_id>",
    methods=["POST"]
)
@admin_required
def delete_medical_record(record_id):

    conn = get_db_connection()

    conn.execute("""
        DELETE FROM medical_records
        WHERE id = ?
    """, (record_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("medical_records")
    )


# =========================================================
# PHARMACY DASHBOARD
# =========================================================

@app.route("/pharmacy")
@role_required(
    "Admin",
    "Pharmacist",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def pharmacy():

    conn = get_db_connection()

    total_medicines = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
    """).fetchone()["count"]

    total_stock = conn.execute("""
        SELECT COALESCE(SUM(quantity), 0) AS total
        FROM medicines
    """).fetchone()["total"]

    low_stock = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
        WHERE quantity <= reorder_level
    """).fetchone()["count"]

    expired_medicines = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
        WHERE expiry_date IS NOT NULL
        AND expiry_date != ''
        AND date(expiry_date) < date('now')
    """).fetchone()["count"]

    pending_prescriptions = conn.execute("""
        SELECT COUNT(*) AS count
        FROM prescriptions
        WHERE status = 'Pending'
    """).fetchone()["count"]

    medicines_list = conn.execute("""
        SELECT *
        FROM medicines
        ORDER BY medicine_name
    """).fetchall()

    low_stock_medicines = conn.execute("""
        SELECT *
        FROM medicines
        WHERE quantity <= reorder_level
        ORDER BY quantity ASC
    """).fetchall()

    expired_medicine_list = conn.execute("""
        SELECT *
        FROM medicines
        WHERE expiry_date IS NOT NULL
        AND expiry_date != ''
        AND date(expiry_date) < date('now')
        ORDER BY expiry_date ASC
    """).fetchall()

    current_date = date.today().isoformat()

    conn.close()

    return render_template(
        "pharmacy.html",

        total_medicines=total_medicines,

        total_stock=total_stock,

        low_stock=low_stock,

        expired_medicines=expired_medicines,

        pending_prescriptions=pending_prescriptions,

        medicines=medicines_list,

        low_stock_medicines=low_stock_medicines,

        expired_medicine_list=expired_medicine_list,

        current_date=current_date,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# MEDICINES
# =========================================================

@app.route("/medicines")
@role_required(
    "Admin",
    "Pharmacist",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def medicines():

    conn = get_db_connection()

    medicines_list = conn.execute("""
        SELECT *
        FROM medicines
        ORDER BY medicine_name
    """).fetchall()

    conn.close()

    return render_template(
        "medicines.html",

        medicines=medicines_list,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# ADD MEDICINE
# =========================================================

@app.route(
    "/add_medicine",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Pharmacist"
)
def add_medicine():

    if request.method == "GET":

        return render_template(
            "add_medicine.html",

            username=session.get("username"),

            role=session.get("role")
        )

    medicine_name = request.form.get("medicine_name")

    category = request.form.get("category")

    unit = request.form.get("unit")

    quantity = request.form.get("quantity")

    reorder_level = request.form.get("reorder_level")

    unit_price = request.form.get("unit_price")

    expiry_date = request.form.get("expiry_date")

    supplier = request.form.get("supplier")

    batch_number = request.form.get("batch_number")

    description = request.form.get("description")

    if not medicine_name:

        return redirect(
            url_for("medicines")
        )

    try:

        quantity = int(quantity or 0)

        reorder_level = int(reorder_level or 10)

        unit_price = float(unit_price or 0)

    except ValueError:

        return "Invalid quantity, reorder level or price."

    conn = get_db_connection()

    conn.execute("""
        INSERT INTO medicines
        (
            medicine_name,
            category,
            unit,
            quantity,
            reorder_level,
            unit_price,
            expiry_date,
            supplier,
            batch_number,
            description
        )

        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        medicine_name,
        category,
        unit,
        quantity,
        reorder_level,
        unit_price,
        expiry_date,
        supplier,
        batch_number,
        description
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("medicines")
    )


# =========================================================
# EDIT MEDICINE
# =========================================================

@app.route(
    "/edit_medicine/<int:medicine_id>",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Pharmacist"
)
def edit_medicine(medicine_id):

    conn = get_db_connection()

    medicine = conn.execute("""
        SELECT *
        FROM medicines
        WHERE id = ?
    """, (medicine_id,)).fetchone()

    if medicine is None:

        conn.close()

        return "Medicine not found"

    if request.method == "POST":

        medicine_name = request.form.get("medicine_name")

        category = request.form.get("category")

        unit = request.form.get("unit")

        quantity = request.form.get("quantity")

        reorder_level = request.form.get("reorder_level")

        unit_price = request.form.get("unit_price")

        expiry_date = request.form.get("expiry_date")

        supplier = request.form.get("supplier")

        batch_number = request.form.get("batch_number")

        description = request.form.get("description")

        try:

            quantity = int(quantity or 0)

            reorder_level = int(reorder_level or 10)

            unit_price = float(unit_price or 0)

        except ValueError:

            conn.close()

            return "Invalid quantity, reorder level or price."

        conn.execute("""
            UPDATE medicines

            SET

                medicine_name = ?,

                category = ?,

                unit = ?,

                quantity = ?,

                reorder_level = ?,

                unit_price = ?,

                expiry_date = ?,

                supplier = ?,

                batch_number = ?,

                description = ?

            WHERE id = ?
        """, (
            medicine_name,
            category,
            unit,
            quantity,
            reorder_level,
            unit_price,
            expiry_date,
            supplier,
            batch_number,
            description,
            medicine_id
        ))

        conn.commit()

        conn.close()

        return redirect(
            url_for("medicines")
        )

    conn.close()

    return render_template(
        "edit_medicine.html",

        medicine=medicine,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# DELETE MEDICINE
# =========================================================

@app.route(
    "/delete_medicine/<int:medicine_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Pharmacist"
)
def delete_medicine(medicine_id):

    conn = get_db_connection()

    prescription_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM prescriptions
        WHERE medicine_id = ?
    """, (medicine_id,)).fetchone()["count"]

    if prescription_count > 0:

        conn.close()

        return """
            <h1>Cannot Delete Medicine</h1>

            <p>
                This medicine is already linked to
                pharmacy prescription records.
            </p>

            <a href="/medicines">
                Back to Medicines
            </a>
        """

    conn.execute("""
        DELETE FROM medicines
        WHERE id = ?
    """, (medicine_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("medicines")
    )


# =========================================================
# PHARMACY PRESCRIPTIONS
# =========================================================

@app.route("/pharmacy_prescriptions")
@role_required(
    "Admin",
    "Pharmacist",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def pharmacy_prescriptions():

    conn = get_db_connection()

    current_role = session.get("role")

    if current_role == "Doctor":

        prescriptions = conn.execute("""
            SELECT

                prescriptions.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name,

                medicines.medicine_name,

                medicines.unit,

                medicines.quantity AS stock_quantity

            FROM prescriptions

            JOIN patients
            ON prescriptions.patient_id =
               patients.id

            JOIN doctors
            ON prescriptions.doctor_id =
               doctors.id

            JOIN medicines
            ON prescriptions.medicine_id =
               medicines.id

            WHERE prescriptions.doctor_id = ?

            ORDER BY prescriptions.id DESC
        """, (
            session.get("doctor_id"),
        )).fetchall()

    else:

        prescriptions = conn.execute("""
            SELECT

                prescriptions.*,

                patients.full_name AS patient_name,

                doctors.full_name AS doctor_name,

                medicines.medicine_name,

                medicines.unit,

                medicines.quantity AS stock_quantity

            FROM prescriptions

            JOIN patients
            ON prescriptions.patient_id =
               patients.id

            JOIN doctors
            ON prescriptions.doctor_id =
               doctors.id

            JOIN medicines
            ON prescriptions.medicine_id =
               medicines.id

            ORDER BY prescriptions.id DESC
        """).fetchall()

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    medicines_list = conn.execute("""
        SELECT *
        FROM medicines
        WHERE quantity > 0
        ORDER BY medicine_name
    """).fetchall()

    if current_role == "Doctor":

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            WHERE id = ?
        """, (
            session.get("doctor_id"),
        )).fetchall()

    else:

        doctors_list = conn.execute("""
            SELECT *
            FROM doctors
            ORDER BY full_name
        """).fetchall()

    conn.close()

    return render_template(
        "pharmacy_prescriptions.html",

        prescriptions=prescriptions,

        patients=patients_list,

        medicines=medicines_list,

        doctors=doctors_list,

        username=session.get("username"),

        role=current_role
    )


# =========================================================
# ADD PHARMACY PRESCRIPTION
# =========================================================

@app.route(
    "/add_pharmacy_prescription",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Doctor"
)
def add_pharmacy_prescription():

    patient_id = request.form.get("patient_id")

    doctor_id = request.form.get("doctor_id")

    medical_record_id = request.form.get(
        "medical_record_id"
    )

    medicine_id = request.form.get("medicine_id")

    dosage = request.form.get("dosage")

    frequency = request.form.get("frequency")

    duration = request.form.get("duration")

    quantity = request.form.get("quantity")

    instructions = request.form.get("instructions")

    prescribed_date = request.form.get(
        "prescribed_date"
    )

    if session.get("role") == "Doctor":

        doctor_id = session.get("doctor_id")

    if not patient_id or not doctor_id or not medicine_id:

        return redirect(
            url_for("pharmacy_prescriptions")
        )

    try:

        quantity = int(quantity or 1)

    except ValueError:

        return "Invalid quantity."

    if quantity <= 0:

        return "Quantity must be greater than zero."

    conn = get_db_connection()

    patient = conn.execute("""
        SELECT id
        FROM patients
        WHERE id = ?
    """, (patient_id,)).fetchone()

    if patient is None:

        conn.close()

        return "Patient not found."

    doctor = conn.execute("""
        SELECT id
        FROM doctors
        WHERE id = ?
    """, (doctor_id,)).fetchone()

    if doctor is None:

        conn.close()

        return "Doctor not found."

    medicine = conn.execute("""
        SELECT *
        FROM medicines
        WHERE id = ?
    """, (medicine_id,)).fetchone()

    if medicine is None:

        conn.close()

        return "Medicine not found."

    if not prescribed_date:

        prescribed_date = date.today().isoformat()

    conn.execute("""
        INSERT INTO prescriptions
        (
            patient_id,
            doctor_id,
            medical_record_id,
            medicine_id,
            dosage,
            frequency,
            duration,
            quantity,
            instructions,
            status,
            prescribed_date
        )

        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        patient_id,
        doctor_id,
        medical_record_id
        if medical_record_id else None,
        medicine_id,
        dosage,
        frequency,
        duration,
        quantity,
        instructions,
        "Pending",
        prescribed_date
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("pharmacy_prescriptions")
    )


# =========================================================
# DISPENSE PRESCRIPTION
# =========================================================

@app.route(
    "/dispense_prescription/<int:prescription_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Pharmacist"
)
def dispense_prescription(prescription_id):

    conn = get_db_connection()

    prescription = conn.execute("""
        SELECT *
        FROM prescriptions
        WHERE id = ?
    """, (prescription_id,)).fetchone()

    if prescription is None:

        conn.close()

        return "Prescription not found."

    if prescription["status"] == "Dispensed":

        conn.close()

        return "This prescription has already been dispensed."

    if prescription["status"] == "Cancelled":

        conn.close()

        return "A cancelled prescription cannot be dispensed."

    medicine = conn.execute("""
        SELECT *
        FROM medicines
        WHERE id = ?
    """, (prescription["medicine_id"],)).fetchone()

    if medicine is None:

        conn.close()

        return "Medicine not found."

    required_quantity = prescription["quantity"]

    current_stock = medicine["quantity"]

    if current_stock < required_quantity:

        conn.close()

        return """
            <h1>Insufficient Stock</h1>

            <p>
                The available stock is not enough
                to dispense this prescription.
            </p>

            <a href="/pharmacy_prescriptions">
                Back to Pharmacy Prescriptions
            </a>
        """

    new_quantity = current_stock - required_quantity

    dispensed_date = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn.execute("""
        UPDATE medicines

        SET quantity = ?

        WHERE id = ?
    """, (
        new_quantity,
        medicine["id"]
    ))

    conn.execute("""
        UPDATE prescriptions

        SET

            status = 'Dispensed',

            dispensed_date = ?,

            dispensed_by = ?

        WHERE id = ?
    """, (
        dispensed_date,
        session.get("username"),
        prescription_id
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("pharmacy_prescriptions")
    )


# =========================================================
# CANCEL PRESCRIPTION
# =========================================================

@app.route(
    "/cancel_prescription/<int:prescription_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Doctor"
)
def cancel_prescription(prescription_id):

    conn = get_db_connection()

    prescription = conn.execute("""
        SELECT *
        FROM prescriptions
        WHERE id = ?
    """, (prescription_id,)).fetchone()

    if prescription is None:

        conn.close()

        return "Prescription not found."

    if prescription["status"] == "Dispensed":

        conn.close()

        return "A dispensed prescription cannot be cancelled."

    if session.get("role") == "Doctor":

        if str(prescription["doctor_id"]) != str(
            session.get("doctor_id")
        ):

            conn.close()

            return """
                <h1>Access Denied</h1>

                <p>
                    You can only cancel prescriptions
                    created under your Doctor Profile.
                </p>

                <a href="/pharmacy_prescriptions">
                    Back to Pharmacy
                </a>
            """

    conn.execute("""
        UPDATE prescriptions

        SET status = 'Cancelled'

        WHERE id = ?
    """, (prescription_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("pharmacy_prescriptions")
    )


# =========================================================
# PHARMACY REPORT
# =========================================================

@app.route("/pharmacy_report")
@role_required(
    "Admin",
    "Pharmacist",
    "Doctor",
    "Data Analyst"
)
def pharmacy_report():

    conn = get_db_connection()

    total_medicines = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medicines
    """).fetchone()["count"]

    total_stock_value = conn.execute("""
        SELECT
            COALESCE(
                SUM(quantity * unit_price),
                0
            ) AS total
        FROM medicines
    """).fetchone()["total"]

    dispensed_prescriptions = conn.execute("""
        SELECT COUNT(*) AS count
        FROM prescriptions
        WHERE status = 'Dispensed'
    """).fetchone()["count"]

    pending_prescriptions = conn.execute("""
        SELECT COUNT(*) AS count
        FROM prescriptions
        WHERE status = 'Pending'
    """).fetchone()["count"]

    cancelled_prescriptions = conn.execute("""
        SELECT COUNT(*) AS count
        FROM prescriptions
        WHERE status = 'Cancelled'
    """).fetchone()["count"]

    category_data = conn.execute("""
        SELECT

            category,

            COUNT(*) AS total_medicines,

            COALESCE(
                SUM(quantity),
                0
            ) AS total_stock

        FROM medicines

        GROUP BY category

        ORDER BY total_stock DESC
    """).fetchall()

    most_dispensed = conn.execute("""
        SELECT

            medicines.medicine_name,

            SUM(
                prescriptions.quantity
            ) AS total_quantity

        FROM prescriptions

        JOIN medicines
        ON prescriptions.medicine_id =
           medicines.id

        WHERE prescriptions.status = 'Dispensed'

        GROUP BY medicines.id

        ORDER BY total_quantity DESC
    """).fetchall()

    conn.close()

    return render_template(
        "pharmacy_report.html",

        total_medicines=total_medicines,

        total_stock_value=total_stock_value,

        dispensed_prescriptions=dispensed_prescriptions,

        pending_prescriptions=pending_prescriptions,

        cancelled_prescriptions=cancelled_prescriptions,

        category_data=category_data,

        most_dispensed=most_dispensed,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# LABORATORY DASHBOARD
# =========================================================

@app.route("/laboratory")
@role_required(
    "Admin",
    "Laboratorian",
    "Patient",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def laboratory():

    conn = get_db_connection()

    current_role = session.get("role")

    if current_role == "Patient":

        patient_id = session.get("patient_id")

        if not patient_id:

            conn.close()

            return """
                <h1>Patient Profile Not Connected</h1>

                <p>
                    Your Patient account is not linked
                    to a Patient Profile.
                </p>

                <a href="/">
                    Back to Dashboard
                </a>
            """

        total_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE patient_id = ?
        """, (patient_id,)).fetchone()["count"]

        pending_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE patient_id = ?
            AND status = 'Pending'
        """, (patient_id,)).fetchone()["count"]

        processing_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE patient_id = ?
            AND status = 'Processing'
        """, (patient_id,)).fetchone()["count"]

        completed_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE patient_id = ?
            AND status = 'Completed'
        """, (patient_id,)).fetchone()["count"]

    else:

        total_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
        """).fetchone()["count"]

        pending_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE status = 'Pending'
        """).fetchone()["count"]

        processing_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE status = 'Processing'
        """).fetchone()["count"]

        completed_requests = conn.execute("""
            SELECT COUNT(*) AS count
            FROM laboratory_requests
            WHERE status = 'Completed'
        """).fetchone()["count"]

    total_tests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_tests
    """).fetchone()["count"]

    conn.close()

    return render_template(
        "laboratory.html",

        total_tests=total_tests,

        total_requests=total_requests,

        pending_requests=pending_requests,

        processing_requests=processing_requests,

        completed_requests=completed_requests,

        username=session.get("username"),

        role=current_role
    )


# =========================================================
# LABORATORY TEST TYPES
# =========================================================

@app.route("/laboratory_tests")
@role_required(
    "Admin",
    "Laboratorian",
    "Data Analyst",
    "Patient",
    "Doctor",
    "Receptionist"
)
def laboratory_tests():

    conn = get_db_connection()

    tests = conn.execute("""
        SELECT *
        FROM laboratory_tests
        ORDER BY test_name
    """).fetchall()

    conn.close()

    return render_template(
        "laboratory_tests.html",

        tests=tests,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# ADD LABORATORY TEST
# =========================================================

@app.route(
    "/add_laboratory_test",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def add_laboratory_test():

    if request.method == "GET":

        return render_template(
            "add_laboratory_test.html",

            username=session.get("username"),

            role=session.get("role")
        )

    test_name = request.form.get("test_name")

    category = request.form.get("category")

    description = request.form.get("description")

    normal_range = request.form.get("normal_range")

    unit = request.form.get("unit")

    price = request.form.get("price")

    recommended_specialty = request.form.get(
        "recommended_specialty"
    )

    if not test_name:

        return "Test name is required."

    try:

        price = float(price or 0)

    except ValueError:

        return "Invalid test price."

    conn = get_db_connection()

    conn.execute("""
        INSERT INTO laboratory_tests
        (
            test_name,
            category,
            description,
            normal_range,
            unit,
            price,
            recommended_specialty
        )

        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (
        test_name,
        category,
        description,
        normal_range,
        unit,
        price,
        recommended_specialty
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for("laboratory_tests")
    )


# =========================================================
# EDIT LABORATORY TEST
# =========================================================

@app.route(
    "/edit_laboratory_test/<int:test_id>",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def edit_laboratory_test(test_id):

    conn = get_db_connection()

    test = conn.execute("""
        SELECT *
        FROM laboratory_tests
        WHERE id = ?
    """, (test_id,)).fetchone()

    if test is None:

        conn.close()

        return "Laboratory test not found."

    if request.method == "POST":

        test_name = request.form.get("test_name")

        category = request.form.get("category")

        description = request.form.get("description")

        normal_range = request.form.get("normal_range")

        unit = request.form.get("unit")

        price = request.form.get("price")

        recommended_specialty = request.form.get(
            "recommended_specialty"
        )

        try:

            price = float(price or 0)

        except ValueError:

            conn.close()

            return "Invalid test price."

        conn.execute("""
            UPDATE laboratory_tests

            SET

                test_name = ?,

                category = ?,

                description = ?,

                normal_range = ?,

                unit = ?,

                price = ?,

                recommended_specialty = ?

            WHERE id = ?
        """, (
            test_name,
            category,
            description,
            normal_range,
            unit,
            price,
            recommended_specialty,
            test_id
        ))

        conn.commit()

        conn.close()

        return redirect(
            url_for("laboratory_tests")
        )

    conn.close()

    return render_template(
        "edit_laboratory_test.html",

        test=test,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# DELETE LABORATORY TEST
# =========================================================

@app.route(
    "/delete_laboratory_test/<int:test_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def delete_laboratory_test(test_id):

    conn = get_db_connection()

    request_count = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE test_id = ?
    """, (test_id,)).fetchone()["count"]

    if request_count > 0:

        conn.close()

        return """
            <h1>Cannot Delete Laboratory Test</h1>

            <p>
                This laboratory test is already linked
                to laboratory requests.
            </p>

            <a href="/laboratory_tests">
                Back to Laboratory Tests
            </a>
        """

    conn.execute("""
        DELETE FROM laboratory_tests
        WHERE id = ?
    """, (test_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("laboratory_tests")
    )


# =========================================================
# REQUEST LABORATORY TEST
# PATIENT + ADMIN
# =========================================================

@app.route(
    "/request_laboratory_test",
    methods=["GET", "POST"]
)
@role_required(
    "Patient",
    "Admin"
)
def request_laboratory_test():

    conn = get_db_connection()

    if request.method == "POST":

        if session.get("role") == "Patient":

            patient_id = session.get("patient_id")

        else:

            patient_id = request.form.get("patient_id")

        test_id = request.form.get("test_id")

        notes = request.form.get("notes")

        if not patient_id or not test_id:

            patients_list = conn.execute("""
                SELECT *
                FROM patients
                ORDER BY full_name
            """).fetchall()

            tests = conn.execute("""
                SELECT *
                FROM laboratory_tests
                ORDER BY test_name
            """).fetchall()

            conn.close()

            return render_template(
                "request_laboratory_test.html",

                patients=patients_list,

                tests=tests,

                error="Patient and Test are required.",

                username=session.get("username"),

                role=session.get("role")
            )

        patient = conn.execute("""
            SELECT id
            FROM patients
            WHERE id = ?
        """, (patient_id,)).fetchone()

        if patient is None:

            conn.close()

            return "Patient not found."

        test = conn.execute("""
            SELECT id
            FROM laboratory_tests
            WHERE id = ?
        """, (test_id,)).fetchone()

        if test is None:

            conn.close()

            return "Laboratory test not found."

        request_date = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        conn.execute("""
            INSERT INTO laboratory_requests
            (
                patient_id,
                test_id,
                request_date,
                status,
                notes,
                created_by
            )

            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            patient_id,
            test_id,
            request_date,
            "Pending",
            notes,
            session.get("username")
        ))

        conn.commit()

        conn.close()

        return redirect(
            url_for("laboratory_requests")
        )

    patients_list = conn.execute("""
        SELECT *
        FROM patients
        ORDER BY full_name
    """).fetchall()

    tests = conn.execute("""
        SELECT *
        FROM laboratory_tests
        ORDER BY test_name
    """).fetchall()

    conn.close()

    return render_template(
        "request_laboratory_test.html",

        patients=patients_list,

        tests=tests,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# LABORATORY REQUESTS
# =========================================================

@app.route("/laboratory_requests")
@role_required(
    "Admin",
    "Laboratorian",
    "Patient",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def laboratory_requests():

    conn = get_db_connection()

    current_role = session.get("role")

    if current_role == "Patient":

        patient_id = session.get("patient_id")

        requests_list = conn.execute("""
            SELECT

                laboratory_requests.*,

                patients.full_name AS patient_name,

                laboratory_tests.test_name,

                laboratory_tests.category,

                laboratory_tests.normal_range,

                laboratory_tests.unit,

                laboratory_tests.recommended_specialty,

                laboratory_results.id AS result_id,

                laboratory_results.result_value,

                laboratory_results.finding,

                laboratory_results.interpretation,

                laboratory_results.result_date,

                laboratory_results.performed_by,

                laboratory_results.recommended_specialty
                    AS result_specialty

            FROM laboratory_requests

            JOIN patients
            ON laboratory_requests.patient_id =
               patients.id

            JOIN laboratory_tests
            ON laboratory_requests.test_id =
               laboratory_tests.id

            LEFT JOIN laboratory_results
            ON laboratory_requests.id =
               laboratory_results.request_id

            WHERE laboratory_requests.patient_id = ?

            ORDER BY laboratory_requests.id DESC
        """, (patient_id,)).fetchall()

    elif current_role == "Doctor":

        doctor = conn.execute("""
            SELECT *
            FROM doctors
            WHERE id = ?
        """, (
            session.get("doctor_id"),
        )).fetchone()

        if doctor is None:

            conn.close()

            return """
                <h1>Doctor Profile Not Found</h1>

                <p>
                    Your Doctor account is not connected
                    to a Doctor Profile.
                </p>

                <a href="/">
                    Back to Dashboard
                </a>
            """

        specialization = doctor["specialization"]

        requests_list = conn.execute("""
            SELECT

                laboratory_requests.*,

                patients.full_name AS patient_name,

                laboratory_tests.test_name,

                laboratory_tests.category,

                laboratory_tests.normal_range,

                laboratory_tests.unit,

                laboratory_tests.recommended_specialty,

                laboratory_results.id AS result_id,

                laboratory_results.result_value,

                laboratory_results.finding,

                laboratory_results.interpretation,

                laboratory_results.result_date,

                laboratory_results.performed_by,

                laboratory_results.recommended_specialty
                    AS result_specialty

            FROM laboratory_requests

            JOIN patients
            ON laboratory_requests.patient_id =
               patients.id

            JOIN laboratory_tests
            ON laboratory_requests.test_id =
               laboratory_tests.id

            LEFT JOIN laboratory_results
            ON laboratory_requests.id =
               laboratory_results.request_id

            WHERE laboratory_requests.status = 'Completed'

            AND (

                LOWER(
                    COALESCE(
                        laboratory_results.recommended_specialty,
                        laboratory_tests.recommended_specialty,
                        ''
                    )
                ) LIKE '%' || LOWER(?) || '%'

                OR

                LOWER(
                    COALESCE(
                        laboratory_tests.recommended_specialty,
                        ''
                    )
                ) LIKE '%' || LOWER(?) || '%'

            )

            ORDER BY laboratory_requests.id DESC
        """, (
            specialization,
            specialization
        )).fetchall()

    else:

        requests_list = conn.execute("""
            SELECT

                laboratory_requests.*,

                patients.full_name AS patient_name,

                laboratory_tests.test_name,

                laboratory_tests.category,

                laboratory_tests.normal_range,

                laboratory_tests.unit,

                laboratory_tests.recommended_specialty,

                laboratory_results.id AS result_id,

                laboratory_results.result_value,

                laboratory_results.finding,

                laboratory_results.interpretation,

                laboratory_results.result_date,

                laboratory_results.performed_by,

                laboratory_results.recommended_specialty
                    AS result_specialty

            FROM laboratory_requests

            JOIN patients
            ON laboratory_requests.patient_id =
               patients.id

            JOIN laboratory_tests
            ON laboratory_requests.test_id =
               laboratory_tests.id

            LEFT JOIN laboratory_results
            ON laboratory_requests.id =
               laboratory_results.request_id

            ORDER BY laboratory_requests.id DESC
        """).fetchall()

    tests = conn.execute("""
        SELECT *
        FROM laboratory_tests
        ORDER BY test_name
    """).fetchall()

    conn.close()

    return render_template(
        "laboratory_requests.html",

        requests_list=requests_list,

        tests=tests,

        username=session.get("username"),

        role=current_role
    )


# =========================================================
# PROCESS LABORATORY REQUEST
# =========================================================

@app.route(
    "/process_laboratory_request/<int:request_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def process_laboratory_request(request_id):

    conn = get_db_connection()

    laboratory_request = conn.execute("""
        SELECT *
        FROM laboratory_requests
        WHERE id = ?
    """, (request_id,)).fetchone()

    if laboratory_request is None:

        conn.close()

        return "Laboratory request not found."

    if laboratory_request["status"] == "Completed":

        conn.close()

        return "This laboratory request is already completed."

    if laboratory_request["status"] == "Cancelled":

        conn.close()

        return "A cancelled laboratory request cannot be processed."

    conn.execute("""
        UPDATE laboratory_requests

        SET status = 'Processing'

        WHERE id = ?
    """, (request_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("laboratory_requests")
    )


# =========================================================
# CANCEL LABORATORY REQUEST
# =========================================================

@app.route(
    "/cancel_laboratory_request/<int:request_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def cancel_laboratory_request(request_id):

    conn = get_db_connection()

    laboratory_request = conn.execute("""
        SELECT *
        FROM laboratory_requests
        WHERE id = ?
    """, (request_id,)).fetchone()

    if laboratory_request is None:

        conn.close()

        return "Laboratory request not found."

    if laboratory_request["status"] == "Completed":

        conn.close()

        return "A completed laboratory request cannot be cancelled."

    conn.execute("""
        UPDATE laboratory_requests

        SET status = 'Cancelled'

        WHERE id = ?
    """, (request_id,))

    conn.commit()

    conn.close()

    return redirect(
        url_for("laboratory_requests")
    )


# =========================================================
# ENTER LABORATORY RESULT
# =========================================================

@app.route(
    "/enter_laboratory_result/<int:request_id>",
    methods=["GET", "POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def enter_laboratory_result(request_id):

    conn = get_db_connection()

    laboratory_request = conn.execute("""
        SELECT

            laboratory_requests.*,

            patients.full_name AS patient_name,

            laboratory_tests.test_name,

            laboratory_tests.category,

            laboratory_tests.normal_range,

            laboratory_tests.unit,

            laboratory_tests.recommended_specialty

        FROM laboratory_requests

        JOIN patients
        ON laboratory_requests.patient_id =
           patients.id

        JOIN laboratory_tests
        ON laboratory_requests.test_id =
           laboratory_tests.id

        WHERE laboratory_requests.id = ?
    """, (request_id,)).fetchone()

    if laboratory_request is None:

        conn.close()

        return "Laboratory request not found."

    existing_result = conn.execute("""
        SELECT *
        FROM laboratory_results
        WHERE request_id = ?
    """, (request_id,)).fetchone()

    if laboratory_request["status"] == "Cancelled":

        conn.close()

        return "A cancelled laboratory request cannot receive a result."

    if request.method == "POST":

        result_value = request.form.get("result_value")

        finding = request.form.get("finding")

        interpretation = request.form.get(
            "interpretation"
        )

        recommended_specialty = request.form.get(
            "recommended_specialty"
        )

        doctor_id = request.form.get("doctor_id")

        if not recommended_specialty:

            recommended_specialty = (
                laboratory_request[
                    "recommended_specialty"
                ]
            )

        if doctor_id == "":

            doctor_id = None

        if doctor_id:

            doctor = conn.execute("""
                SELECT id
                FROM doctors
                WHERE id = ?
            """, (doctor_id,)).fetchone()

            if doctor is None:

                conn.close()

                return "Selected doctor not found."

        result_date = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        if existing_result:

            conn.execute("""
                UPDATE laboratory_results

                SET

                    result_value = ?,

                    finding = ?,

                    interpretation = ?,

                    result_date = ?,

                    performed_by = ?,

                    recommended_specialty = ?,

                    doctor_id = ?

                WHERE id = ?
            """, (
                result_value,
                finding,
                interpretation,
                result_date,
                session.get("username"),
                recommended_specialty,
                doctor_id,
                existing_result["id"]
            ))

        else:

            conn.execute("""
                INSERT INTO laboratory_results
                (
                    request_id,
                    result_value,
                    finding,
                    interpretation,
                    result_date,
                    performed_by,
                    recommended_specialty,
                    doctor_id
                )

                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                request_id,
                result_value,
                finding,
                interpretation,
                result_date,
                session.get("username"),
                recommended_specialty,
                doctor_id
            ))

        conn.execute("""
            UPDATE laboratory_requests

            SET status = 'Completed'

            WHERE id = ?
        """, (request_id,))

        conn.commit()

        conn.close()

        return redirect(
            url_for("laboratory_requests")
        )

    doctors = conn.execute("""
        SELECT *
        FROM doctors
        ORDER BY full_name
    """).fetchall()

    conn.close()

    return render_template(
        "enter_laboratory_result.html",

        laboratory_request=laboratory_request,

        existing_result=existing_result,

        doctors=doctors,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# VIEW LABORATORY RESULT
# =========================================================

@app.route(
    "/laboratory_result/<int:result_id>"
)
@role_required(
    "Admin",
    "Laboratorian",
    "Patient",
    "Doctor",
    "Receptionist",
    "Data Analyst"
)
def view_laboratory_result(result_id):

    conn = get_db_connection()

    result = conn.execute("""
        SELECT

            laboratory_results.*,

            laboratory_requests.patient_id,

            laboratory_requests.request_date,

            laboratory_requests.status,

            laboratory_requests.notes,

            patients.full_name AS patient_name,

            patients.age AS patient_age,

            patients.gender AS patient_gender,

            patients.phone AS patient_phone,

            laboratory_tests.test_name,

            laboratory_tests.category,

            laboratory_tests.normal_range,

            laboratory_tests.unit,

            laboratory_tests.price,

            laboratory_tests.recommended_specialty
                AS test_specialty,

            doctors.full_name AS doctor_name,

            doctors.specialization AS doctor_specialization

        FROM laboratory_results

        JOIN laboratory_requests
        ON laboratory_results.request_id =
           laboratory_requests.id

        JOIN patients
        ON laboratory_requests.patient_id =
           patients.id

        JOIN laboratory_tests
        ON laboratory_requests.test_id =
           laboratory_tests.id

        LEFT JOIN doctors
        ON laboratory_results.doctor_id =
           doctors.id

        WHERE laboratory_results.id = ?
    """, (result_id,)).fetchone()

    if result is None:

        conn.close()

        return "Laboratory result not found."

    if session.get("role") == "Patient":

        if str(result["patient_id"]) != str(
            session.get("patient_id")
        ):

            conn.close()

            return """
                <h1>Access Denied</h1>

                <p>
                    You can only view your own
                    laboratory results.
                </p>

                <a href="/laboratory_requests">
                    Back to Laboratory
                </a>
            """

    if session.get("role") == "Doctor":

        doctor = conn.execute("""
            SELECT *
            FROM doctors
            WHERE id = ?
        """, (
            session.get("doctor_id"),
        )).fetchone()

        if doctor is None:

            conn.close()

            return "Doctor profile not found."

        specialization = doctor["specialization"]

        result_specialty = (
            result["recommended_specialty"]
            or result["test_specialty"]
            or ""
        )

        if (
            result_specialty
            and
            specialization.lower()
            not in result_specialty.lower()
            and
            result_specialty.lower()
            not in specialization.lower()
        ):

            conn.close()

            return """
                <h1>Access Denied</h1>

                <p>
                    This laboratory result is not
                    assigned to your specialty.
                </p>

                <a href="/laboratory_requests">
                    Back to Laboratory
                </a>
            """

    recommended_specialty = (
        result["recommended_specialty"]
        or result["test_specialty"]
        or ""
    )

    if recommended_specialty:

        doctors = conn.execute("""
            SELECT *
            FROM doctors
            WHERE LOWER(specialization)
            LIKE '%' || LOWER(?) || '%'
            ORDER BY full_name
        """, (
            recommended_specialty,
        )).fetchall()

    else:

        doctors = conn.execute("""
            SELECT *
            FROM doctors
            ORDER BY full_name
        """).fetchall()

    conn.close()

    return render_template(
        "laboratory_result_details.html",

        result=result,

        doctors=doctors,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# ASSIGN LABORATORY RESULT TO DOCTOR
# =========================================================

@app.route(
    "/assign_laboratory_result/<int:result_id>",
    methods=["POST"]
)
@role_required(
    "Admin",
    "Laboratorian"
)
def assign_laboratory_result(result_id):

    doctor_id = request.form.get("doctor_id")

    conn = get_db_connection()

    result = conn.execute("""
        SELECT *
        FROM laboratory_results
        WHERE id = ?
    """, (result_id,)).fetchone()

    if result is None:

        conn.close()

        return "Laboratory result not found."

    if not doctor_id:

        conn.close()

        return "Please select a doctor."

    doctor = conn.execute("""
        SELECT *
        FROM doctors
        WHERE id = ?
    """, (doctor_id,)).fetchone()

    if doctor is None:

        conn.close()

        return "Doctor not found."

    conn.execute("""
        UPDATE laboratory_results

        SET doctor_id = ?

        WHERE id = ?
    """, (
        doctor_id,
        result_id
    ))

    conn.commit()

    conn.close()

    return redirect(
        url_for(
            "view_laboratory_result",
            result_id=result_id
        )
    )


# =========================================================
# LABORATORY REPORT
# =========================================================

@app.route("/laboratory_report")
@role_required(
    "Admin",
    "Laboratorian",
    "Data Analyst"
)
def laboratory_report():

    conn = get_db_connection()

    total_tests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_tests
    """).fetchone()["count"]

    total_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
    """).fetchone()["count"]

    pending_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Pending'
    """).fetchone()["count"]

    processing_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Processing'
    """).fetchone()["count"]

    completed_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Completed'
    """).fetchone()["count"]

    cancelled_requests = conn.execute("""
        SELECT COUNT(*) AS count
        FROM laboratory_requests
        WHERE status = 'Cancelled'
    """).fetchone()["count"]

    test_usage = conn.execute("""
        SELECT

            laboratory_tests.test_name,

            laboratory_tests.category,

            COUNT(
                laboratory_requests.id
            ) AS total_requests,

            SUM(
                CASE
                    WHEN laboratory_requests.status =
                        'Completed'
                    THEN 1
                    ELSE 0
                END
            ) AS completed

        FROM laboratory_tests

        LEFT JOIN laboratory_requests
        ON laboratory_tests.id =
           laboratory_requests.test_id

        GROUP BY laboratory_tests.id

        ORDER BY total_requests DESC
    """).fetchall()

    specialty_data = conn.execute("""
        SELECT

            COALESCE(
                laboratory_results.recommended_specialty,
                laboratory_tests.recommended_specialty,
                'Not Assigned'
            ) AS specialty,

            COUNT(laboratory_requests.id) AS total

        FROM laboratory_requests

        JOIN laboratory_tests
        ON laboratory_requests.test_id =
           laboratory_tests.id

        LEFT JOIN laboratory_results
        ON laboratory_requests.id =
           laboratory_results.request_id

        GROUP BY specialty

        ORDER BY total DESC
    """).fetchall()

    conn.close()

    return render_template(
        "laboratory_report.html",

        total_tests=total_tests,

        total_requests=total_requests,

        pending_requests=pending_requests,

        processing_requests=processing_requests,

        completed_requests=completed_requests,

        cancelled_requests=cancelled_requests,

        test_usage=test_usage,

        specialty_data=specialty_data,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# REPORTS HELPER
# =========================================================

def get_report_scope():

    current_role = session.get("role")

    if current_role == "Doctor":

        return (
            " AND appointments.doctor_id = ? ",
            [session.get("doctor_id")]
        )

    return (
        "",
        []
    )


def get_medical_record_report_scope():

    current_role = session.get("role")

    if current_role == "Doctor":

        return (
            " AND medical_records.doctor_id = ? ",
            [session.get("doctor_id")]
        )

    return (
        "",
        []
    )


# =========================================================
# REPORTS DASHBOARD
# =========================================================

@app.route("/reports")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def reports():

    conn = get_db_connection()

    total_patients = conn.execute("""
        SELECT COUNT(*) AS count
        FROM patients
    """).fetchone()["count"]

    total_doctors = conn.execute("""
        SELECT COUNT(*) AS count
        FROM doctors
    """).fetchone()["count"]

    total_appointments = conn.execute("""
        SELECT COUNT(*) AS count
        FROM appointments
    """).fetchone()["count"]

    total_medical_records = conn.execute("""
        SELECT COUNT(*) AS count
        FROM medical_records
    """).fetchone()["count"]

    conn.close()

    return render_template(
        "reports.html",

        total_patients=total_patients,

        total_doctors=total_doctors,

        total_appointments=total_appointments,

        total_medical_records=total_medical_records,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# DAILY REPORT
# =========================================================

@app.route("/reports/daily")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def daily_report():

    report_date = request.args.get(
        "report_date",
        date.today().isoformat()
    )

    try:

        datetime.strptime(
            report_date,
            "%Y-%m-%d"
        )

    except ValueError:

        report_date = date.today().isoformat()

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = get_medical_record_report_scope()

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count
        FROM appointments
        WHERE appointment_date = ?
        """ +
        appointment_scope,
        [report_date] + appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE appointment_date = ?
        """ +
        appointment_scope,
        [report_date] + appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date = ?
        """ +
        record_scope,
        [report_date] + record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date = ?
        AND diagnosis IS NOT NULL
        AND TRIM(diagnosis) != ''
        """ +
        record_scope,
        [report_date] + record_params
    ).fetchone()["count"]

    appointments_list = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date = ?
        """ +
        appointment_scope +
        """

        ORDER BY
            appointments.appointment_time ASC,
            appointments.id ASC
        """,
        [report_date] + appointment_params
    ).fetchall()

    medical_records_list = conn.execute(
        """
        SELECT

            medical_records.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM medical_records

        JOIN patients
        ON medical_records.patient_id =
           patients.id

        JOIN doctors
        ON medical_records.doctor_id =
           doctors.id

        WHERE medical_records.visit_date = ?
        """ +
        record_scope +
        """

        ORDER BY medical_records.id DESC
        """,
        [report_date] + record_params
    ).fetchall()

    conn.close()

    return render_template(
        "daily_report.html",

        report_date=report_date,

        current_date=date.today().isoformat(),

        total_patients=total_patients,

        total_appointments=total_appointments,

        total_medical_records=total_medical_records,

        total_diagnoses=total_diagnoses,

        appointments=appointments_list,

        medical_records=medical_records_list,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# WEEKLY REPORT
# =========================================================

@app.route("/reports/weekly")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def weekly_report():

    today = date.today()

    default_start = today - timedelta(
        days=today.weekday()
    )

    default_end = default_start + timedelta(
        days=6
    )

    start_date = request.args.get(
        "start_date",
        default_start.isoformat()
    )

    end_date = request.args.get(
        "end_date",
        default_end.isoformat()
    )

    try:

        start_object = datetime.strptime(
            start_date,
            "%Y-%m-%d"
        ).date()

        end_object = datetime.strptime(
            end_date,
            "%Y-%m-%d"
        ).date()

        if start_object > end_object:

            start_object, end_object = (
                end_object,
                start_object
            )

            start_date = start_object.isoformat()

            end_date = end_object.isoformat()

    except ValueError:

        start_object = default_start

        end_object = default_end

        start_date = start_object.isoformat()

        end_date = end_object.isoformat()

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = get_medical_record_report_scope()

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count
        FROM appointments
        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,
        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,
        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date BETWEEN ? AND ?
        """ +
        record_scope,
        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date BETWEEN ? AND ?
        AND diagnosis IS NOT NULL
        AND TRIM(diagnosis) != ''
        """ +
        record_scope,
        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    appointments_list = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date
        BETWEEN ? AND ?
        """ +
        appointment_scope +
        """

        ORDER BY
            appointments.appointment_date ASC,
            appointments.appointment_time ASC
        """,
        [start_date, end_date] +
        appointment_params
    ).fetchall()

    daily_data = []

    current_day = start_object

    while current_day <= end_object:

        current_date = current_day.isoformat()

        day_patients = conn.execute(
            """
            SELECT COUNT(DISTINCT patient_id) AS count
            FROM appointments
            WHERE appointment_date = ?
            """ +
            appointment_scope,
            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_appointments = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM appointments
            WHERE appointment_date = ?
            """ +
            appointment_scope,
            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_medical_records = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM medical_records
            WHERE visit_date = ?
            """ +
            record_scope,
            [current_date] +
            record_params
        ).fetchone()["count"]

        day_diagnoses = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM medical_records
            WHERE visit_date = ?
            AND diagnosis IS NOT NULL
            AND TRIM(diagnosis) != ''
            """ +
            record_scope,
            [current_date] +
            record_params
        ).fetchone()["count"]

        daily_data.append({
            "date": current_date,
            "patients": day_patients,
            "appointments": day_appointments,
            "medical_records": day_medical_records,
            "diagnoses": day_diagnoses
        })

        current_day += timedelta(days=1)

    conn.close()

    return render_template(
        "weekly_report.html",

        start_date=start_date,

        end_date=end_date,

        total_patients=total_patients,

        total_appointments=total_appointments,

        total_medical_records=total_medical_records,

        total_diagnoses=total_diagnoses,

        daily_data=daily_data,

        appointments=appointments_list,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# MONTHLY REPORT
# =========================================================

@app.route("/reports/monthly")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def monthly_report():

    today = date.today()

    try:

        selected_month = int(
            request.args.get(
                "month",
                today.month
            )
        )

        selected_year = int(
            request.args.get(
                "year",
                today.year
            )
        )

        if selected_month < 1 or selected_month > 12:

            selected_month = today.month

        if selected_year < 2000 or selected_year > 2100:

            selected_year = today.year

    except ValueError:

        selected_month = today.month

        selected_year = today.year

    start_object = date(
        selected_year,
        selected_month,
        1
    )

    if selected_month == 12:

        next_month = date(
            selected_year + 1,
            1,
            1
        )

    else:

        next_month = date(
            selected_year,
            selected_month + 1,
            1
        )

    end_object = next_month - timedelta(
        days=1
    )

    start_date = start_object.isoformat()

    end_date = end_object.isoformat()

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = get_medical_record_report_scope()

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count
        FROM appointments
        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,
        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM appointments
        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,
        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date BETWEEN ? AND ?
        """ +
        record_scope,
        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM medical_records
        WHERE visit_date BETWEEN ? AND ?
        AND diagnosis IS NOT NULL
        AND TRIM(diagnosis) != ''
        """ +
        record_scope,
        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    appointments_list = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date
        BETWEEN ? AND ?
        """ +
        appointment_scope +
        """

        ORDER BY
            appointments.appointment_date ASC,
            appointments.appointment_time ASC
        """,
        [start_date, end_date] +
        appointment_params
    ).fetchall()

    diagnoses = conn.execute(
        """
        SELECT

            diagnosis,

            COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?

        AND diagnosis IS NOT NULL

        AND TRIM(diagnosis) != ''

        """ +
        record_scope +
        """

        GROUP BY diagnosis

        ORDER BY count DESC
        """,
        [start_date, end_date] +
        record_params
    ).fetchall()

    daily_data = []

    current_day = start_object

    while current_day <= end_object:

        current_date = current_day.isoformat()

        day_patients = conn.execute(
            """
            SELECT COUNT(DISTINCT patient_id) AS count
            FROM appointments
            WHERE appointment_date = ?
            """ +
            appointment_scope,
            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_appointments = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM appointments
            WHERE appointment_date = ?
            """ +
            appointment_scope,
            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_medical_records = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM medical_records
            WHERE visit_date = ?
            """ +
            record_scope,
            [current_date] +
            record_params
        ).fetchone()["count"]

        day_diagnoses = conn.execute(
            """
            SELECT COUNT(*) AS count
            FROM medical_records
            WHERE visit_date = ?
            AND diagnosis IS NOT NULL
            AND TRIM(diagnosis) != ''
            """ +
            record_scope,
            [current_date] +
            record_params
        ).fetchone()["count"]

        daily_data.append({
            "date": current_date,
            "patients": day_patients,
            "appointments": day_appointments,
            "medical_records": day_medical_records,
            "diagnoses": day_diagnoses
        })

        current_day += timedelta(days=1)

    conn.close()

    return render_template(
        "monthly_report.html",

        selected_month=selected_month,

        selected_year=selected_year,

        total_patients=total_patients,

        total_appointments=total_appointments,

        total_medical_records=total_medical_records,

        total_diagnoses=total_diagnoses,

        daily_data=daily_data,

        appointments=appointments_list,

        diagnoses=diagnoses,

        username=session.get("username"),

        role=session.get("role")
    )


# =========================================================
# REPORT EXPORT HELPERS
# PDF + EXCEL + CSV
# =========================================================

def parse_report_date(value, default_value):

    if not value:
        return default_value

    try:

        datetime.strptime(
            value,
            "%Y-%m-%d"
        )

        return value

    except ValueError:

        return default_value


# =========================================================
# DAILY REPORT DATA FOR EXPORT
# =========================================================

def get_daily_report_data(report_date):

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = (
        get_medical_record_report_scope()
    )

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count

        FROM appointments

        WHERE appointment_date = ?
        """ +
        appointment_scope,

        [report_date] +
        appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM appointments

        WHERE appointment_date = ?
        """ +
        appointment_scope,

        [report_date] +
        appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date = ?
        """ +
        record_scope,

        [report_date] +
        record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date = ?

        AND diagnosis IS NOT NULL

        AND TRIM(diagnosis) != ''
        """ +
        record_scope,

        [report_date] +
        record_params
    ).fetchone()["count"]

    appointments = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date = ?
        """ +
        appointment_scope +
        """

        ORDER BY

            appointments.appointment_time ASC,

            appointments.id ASC
        """,

        [report_date] +
        appointment_params
    ).fetchall()

    medical_records = conn.execute(
        """
        SELECT

            medical_records.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM medical_records

        JOIN patients
        ON medical_records.patient_id =
           patients.id

        JOIN doctors
        ON medical_records.doctor_id =
           doctors.id

        WHERE medical_records.visit_date = ?
        """ +
        record_scope +
        """

        ORDER BY medical_records.id DESC
        """,

        [report_date] +
        record_params
    ).fetchall()

    conn.close()

    return {
        "report_date": report_date,

        "total_patients": total_patients,

        "total_appointments": total_appointments,

        "total_medical_records": total_medical_records,

        "total_diagnoses": total_diagnoses,

        "appointments": appointments,

        "medical_records": medical_records
    }


# =========================================================
# WEEKLY REPORT DATA FOR EXPORT
# =========================================================

def get_weekly_report_data(start_date, end_date):

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = (
        get_medical_record_report_scope()
    )

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count

        FROM appointments

        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,

        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM appointments

        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,

        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?
        """ +
        record_scope,

        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?

        AND diagnosis IS NOT NULL

        AND TRIM(diagnosis) != ''
        """ +
        record_scope,

        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    appointments = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date
        BETWEEN ? AND ?
        """ +
        appointment_scope +
        """

        ORDER BY

            appointments.appointment_date ASC,

            appointments.appointment_time ASC
        """,

        [start_date, end_date] +
        appointment_params
    ).fetchall()

    daily_data = []

    start_object = datetime.strptime(
        start_date,
        "%Y-%m-%d"
    ).date()

    end_object = datetime.strptime(
        end_date,
        "%Y-%m-%d"
    ).date()

    current_day = start_object

    while current_day <= end_object:

        current_date = current_day.isoformat()

        day_patients = conn.execute(
            """
            SELECT COUNT(DISTINCT patient_id) AS count

            FROM appointments

            WHERE appointment_date = ?
            """ +
            appointment_scope,

            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_appointments = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM appointments

            WHERE appointment_date = ?
            """ +
            appointment_scope,

            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_medical_records = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM medical_records

            WHERE visit_date = ?
            """ +
            record_scope,

            [current_date] +
            record_params
        ).fetchone()["count"]

        day_diagnoses = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM medical_records

            WHERE visit_date = ?

            AND diagnosis IS NOT NULL

            AND TRIM(diagnosis) != ''
            """ +
            record_scope,

            [current_date] +
            record_params
        ).fetchone()["count"]

        daily_data.append({
            "date": current_date,

            "patients": day_patients,

            "appointments": day_appointments,

            "medical_records": day_medical_records,

            "diagnoses": day_diagnoses
        })

        current_day += timedelta(days=1)

    conn.close()

    return {
        "start_date": start_date,

        "end_date": end_date,

        "total_patients": total_patients,

        "total_appointments": total_appointments,

        "total_medical_records": total_medical_records,

        "total_diagnoses": total_diagnoses,

        "daily_data": daily_data,

        "appointments": appointments
    }


# =========================================================
# MONTHLY REPORT DATA FOR EXPORT
# =========================================================

def get_monthly_report_data(
    selected_month,
    selected_year
):

    start_object = date(
        selected_year,
        selected_month,
        1
    )

    if selected_month == 12:

        next_month = date(
            selected_year + 1,
            1,
            1
        )

    else:

        next_month = date(
            selected_year,
            selected_month + 1,
            1
        )

    end_object = next_month - timedelta(
        days=1
    )

    start_date = start_object.isoformat()

    end_date = end_object.isoformat()

    conn = get_db_connection()

    appointment_scope, appointment_params = get_report_scope()

    record_scope, record_params = (
        get_medical_record_report_scope()
    )

    total_patients = conn.execute(
        """
        SELECT COUNT(DISTINCT patient_id) AS count

        FROM appointments

        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,

        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_appointments = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM appointments

        WHERE appointment_date BETWEEN ? AND ?
        """ +
        appointment_scope,

        [start_date, end_date] +
        appointment_params
    ).fetchone()["count"]

    total_medical_records = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?
        """ +
        record_scope,

        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    total_diagnoses = conn.execute(
        """
        SELECT COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?

        AND diagnosis IS NOT NULL

        AND TRIM(diagnosis) != ''
        """ +
        record_scope,

        [start_date, end_date] +
        record_params
    ).fetchone()["count"]

    appointments = conn.execute(
        """
        SELECT

            appointments.*,

            patients.full_name AS patient_name,

            doctors.full_name AS doctor_name

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        WHERE appointments.appointment_date
        BETWEEN ? AND ?
        """ +
        appointment_scope +
        """

        ORDER BY

            appointments.appointment_date ASC,

            appointments.appointment_time ASC
        """,

        [start_date, end_date] +
        appointment_params
    ).fetchall()

    diagnoses = conn.execute(
        """
        SELECT

            diagnosis,

            COUNT(*) AS count

        FROM medical_records

        WHERE visit_date BETWEEN ? AND ?

        AND diagnosis IS NOT NULL

        AND TRIM(diagnosis) != ''

        """ +
        record_scope +
        """

        GROUP BY diagnosis

        ORDER BY count DESC
        """,

        [start_date, end_date] +
        record_params
    ).fetchall()

    daily_data = []

    current_day = start_object

    while current_day <= end_object:

        current_date = current_day.isoformat()

        day_patients = conn.execute(
            """
            SELECT COUNT(DISTINCT patient_id) AS count

            FROM appointments

            WHERE appointment_date = ?
            """ +
            appointment_scope,

            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_appointments = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM appointments

            WHERE appointment_date = ?
            """ +
            appointment_scope,

            [current_date] +
            appointment_params
        ).fetchone()["count"]

        day_medical_records = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM medical_records

            WHERE visit_date = ?
            """ +
            record_scope,

            [current_date] +
            record_params
        ).fetchone()["count"]

        day_diagnoses = conn.execute(
            """
            SELECT COUNT(*) AS count

            FROM medical_records

            WHERE visit_date = ?

            AND diagnosis IS NOT NULL

            AND TRIM(diagnosis) != ''
            """ +
            record_scope,

            [current_date] +
            record_params
        ).fetchone()["count"]

        daily_data.append({
            "date": current_date,

            "patients": day_patients,

            "appointments": day_appointments,

            "medical_records": day_medical_records,

            "diagnoses": day_diagnoses
        })

        current_day += timedelta(days=1)

    conn.close()

    return {
        "month": selected_month,

        "year": selected_year,

        "start_date": start_date,

        "end_date": end_date,

        "total_patients": total_patients,

        "total_appointments": total_appointments,

        "total_medical_records": total_medical_records,

        "total_diagnoses": total_diagnoses,

        "daily_data": daily_data,

        "appointments": appointments,

        "diagnoses": diagnoses
    }


# =========================================================
# PDF HELPERS
# =========================================================

def pdf_title(title):

    styles = getSampleStyleSheet()

    return Paragraph(
        title,
        ParagraphStyle(
            "ReportTitle",
            parent=styles["Title"],
            alignment=TA_CENTER,
            fontSize=18,
            spaceAfter=12
        )
    )


def pdf_table(
    data,
    column_widths=None
):

    table = Table(
        data,
        colWidths=column_widths,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.grey
            ),

            (
                "TEXTCOLOR",
                (0, 0),
                (-1, 0),
                colors.white
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            ),

            (
                "GRID",
                (0, 0),
                (-1, -1),
                0.5,
                colors.black
            ),

            (
                "ALIGN",
                (0, 0),
                (-1, -1),
                "LEFT"
            ),

            (
                "VALIGN",
                (0, 0),
                (-1, -1),
                "TOP"
            ),

            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                8
            ),

            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, 0),
                7
            ),

            (
                "TOPPADDING",
                (0, 0),
                (-1, 0),
                7
            )
        ])
    )

    return table


def build_daily_pdf(data):

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    story = []

    story.append(
        pdf_title(
            "HOSPITAL DATA ANALYTICS SYSTEM"
        )
    )

    story.append(
        Paragraph(
            f"DAILY REPORT - {data['report_date']}",
            getSampleStyleSheet()["Heading2"]
        )
    )

    story.append(Spacer(1, 12))

    summary = [
        ["Metric", "Value"],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ]
    ]

    story.append(
        pdf_table(
            summary,
            [250, 100]
        )
    )

    story.append(Spacer(1, 18))

    story.append(
        Paragraph(
            "APPOINTMENTS",
            getSampleStyleSheet()["Heading2"]
        )
    )

    appointment_rows = [
        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    for row in data["appointments"]:

        appointment_rows.append([
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ])

    if len(appointment_rows) == 1:

        appointment_rows.append([
            "-",
            "-",
            "No appointments",
            "-",
            "-",
            "-"
        ])

    story.append(
        pdf_table(
            appointment_rows
        )
    )

    story.append(PageBreak())

    story.append(
        Paragraph(
            "MEDICAL RECORDS",
            getSampleStyleSheet()["Heading2"]
        )
    )

    record_rows = [
        [
            "Date",
            "Patient",
            "Doctor",
            "Chief Complaint",
            "Diagnosis",
            "Treatment"
        ]
    ]

    for row in data["medical_records"]:

        record_rows.append([
            row["visit_date"],
            row["patient_name"],
            row["doctor_name"],
            row["chief_complaint"],
            row["diagnosis"],
            row["treatment"]
        ])

    if len(record_rows) == 1:

        record_rows.append([
            "-",
            "No medical records",
            "-",
            "-",
            "-",
            "-"
        ])

    story.append(
        pdf_table(
            record_rows
        )
    )

    document.build(story)

    buffer.seek(0)

    return buffer


def build_weekly_pdf(data):

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    story = []

    story.append(
        pdf_title(
            "HOSPITAL DATA ANALYTICS SYSTEM"
        )
    )

    story.append(
        Paragraph(
            f"WEEKLY REPORT - "
            f"{data['start_date']} TO "
            f"{data['end_date']}",
            getSampleStyleSheet()["Heading2"]
        )
    )

    story.append(Spacer(1, 12))

    summary = [
        ["Metric", "Value"],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ]
    ]

    story.append(
        pdf_table(
            summary,
            [250, 100]
        )
    )

    story.append(Spacer(1, 18))

    story.append(
        Paragraph(
            "DAILY SUMMARY",
            getSampleStyleSheet()["Heading2"]
        )
    )

    daily_rows = [
        [
            "Date",
            "Patients",
            "Appointments",
            "Medical Records",
            "Diagnoses"
        ]
    ]

    for row in data["daily_data"]:

        daily_rows.append([
            row["date"],
            row["patients"],
            row["appointments"],
            row["medical_records"],
            row["diagnoses"]
        ])

    story.append(
        pdf_table(
            daily_rows
        )
    )

    story.append(Spacer(1, 18))

    story.append(
        Paragraph(
            "APPOINTMENTS",
            getSampleStyleSheet()["Heading2"]
        )
    )

    appointment_rows = [
        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    for row in data["appointments"]:

        appointment_rows.append([
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ])

    if len(appointment_rows) == 1:

        appointment_rows.append([
            "-",
            "-",
            "No appointments",
            "-",
            "-",
            "-"
        ])

    story.append(
        pdf_table(
            appointment_rows
        )
    )

    document.build(story)

    buffer.seek(0)

    return buffer


def build_monthly_pdf(data):

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    story = []

    story.append(
        pdf_title(
            "HOSPITAL DATA ANALYTICS SYSTEM"
        )
    )

    story.append(
        Paragraph(
            f"MONTHLY REPORT - "
            f"{data['year']}-{data['month']:02d}",
            getSampleStyleSheet()["Heading2"]
        )
    )

    story.append(Spacer(1, 12))

    summary = [
        ["Metric", "Value"],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ]
    ]

    story.append(
        pdf_table(
            summary,
            [250, 100]
        )
    )

    story.append(Spacer(1, 18))

    story.append(
        Paragraph(
            "DAILY SUMMARY",
            getSampleStyleSheet()["Heading2"]
        )
    )

    daily_rows = [
        [
            "Date",
            "Patients",
            "Appointments",
            "Medical Records",
            "Diagnoses"
        ]
    ]

    for row in data["daily_data"]:

        daily_rows.append([
            row["date"],
            row["patients"],
            row["appointments"],
            row["medical_records"],
            row["diagnoses"]
        ])

    story.append(
        pdf_table(
            daily_rows
        )
    )

    story.append(PageBreak())

    story.append(
        Paragraph(
            "DIAGNOSES",
            getSampleStyleSheet()["Heading2"]
        )
    )

    diagnosis_rows = [
        [
            "Diagnosis",
            "Count"
        ]
    ]

    for row in data["diagnoses"]:

        diagnosis_rows.append([
            row["diagnosis"],
            row["count"]
        ])

    if len(diagnosis_rows) == 1:

        diagnosis_rows.append([
            "No diagnoses",
            0
        ])

    story.append(
        pdf_table(
            diagnosis_rows
        )
    )

    story.append(Spacer(1, 18))

    story.append(
        Paragraph(
            "APPOINTMENTS",
            getSampleStyleSheet()["Heading2"]
        )
    )

    appointment_rows = [
        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    for row in data["appointments"]:

        appointment_rows.append([
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ])

    if len(appointment_rows) == 1:

        appointment_rows.append([
            "-",
            "-",
            "No appointments",
            "-",
            "-",
            "-"
        ])

    story.append(
        pdf_table(
            appointment_rows
        )
    )

    document.build(story)

    buffer.seek(0)

    return buffer


# =========================================================
# EXCEL HELPER
# =========================================================

def excel_file(sheets, filename):

    workbook = Workbook()

    first_sheet = True

    for sheet_name, rows in sheets:

        if first_sheet:

            worksheet = workbook.active

            worksheet.title = sheet_name

            first_sheet = False

        else:

            worksheet = workbook.create_sheet(
                title=sheet_name
            )

        for row in rows:

            worksheet.append(row)

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

        for column in worksheet.columns:

            max_length = 0

            column_letter = column[0].column_letter

            for cell in column:

                try:

                    cell_length = len(
                        str(cell.value)
                    )

                    if cell_length > max_length:

                        max_length = cell_length

                except Exception:

                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 3,
                40
            )

    output = BytesIO()

    workbook.save(output)

    output.seek(0)

    return output


# =========================================================
# CSV HELPER
# =========================================================

def csv_file(rows):

    output = StringIO()

    writer = csv.writer(
        output
    )

    for row in rows:

        writer.writerow(row)

    memory_file = BytesIO()

    memory_file.write(
        output.getvalue().encode(
            "utf-8-sig"
        )
    )

    memory_file.seek(0)

    return memory_file


# =========================================================
# ANALYTICS EXPORT HELPERS
# =========================================================

def get_analytics_export_data():

    dataset = request.args.get(
        "dataset",
        "patients"
    )

    analysis = request.args.get(
        "analysis",
        "bar"
    )

    x_variable = request.args.get(
        "x_variable",
        "gender"
    )

    y_variable = request.args.get(
        "y_variable",
        "count"
    )

    date_from = request.args.get(
        "date_from"
    )

    date_to = request.args.get(
        "date_to"
    )

    conn = get_db_connection()

    labels = []
    values = []
    title = ""

    try:

        # =====================================================
        # PATIENTS
        # =====================================================

        if dataset == "patients":

            title = "Patients Analytics"

            if x_variable == "age":

                rows = conn.execute("""
                    SELECT
                        age,
                        COUNT(*) AS count
                    FROM patients
                    GROUP BY age
                    ORDER BY age
                """).fetchall()

                labels = [
                    row["age"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute("""
                    SELECT
                        gender,
                        COUNT(*) AS count
                    FROM patients
                    GROUP BY gender
                    ORDER BY gender
                """).fetchall()

                labels = [
                    row["gender"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # DOCTORS
        # =====================================================

        elif dataset == "doctors":

            title = "Doctors Analytics"

            if x_variable == "gender":

                rows = conn.execute("""
                    SELECT
                        gender,
                        COUNT(*) AS count
                    FROM doctors
                    GROUP BY gender
                    ORDER BY gender
                """).fetchall()

                labels = [
                    row["gender"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute("""
                    SELECT
                        specialization,
                        COUNT(*) AS count
                    FROM doctors
                    GROUP BY specialization
                    ORDER BY specialization
                """).fetchall()

                labels = [
                    row["specialization"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # APPOINTMENTS
        # =====================================================

        elif dataset == "appointments":

            title = "Appointments Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "appointment_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "appointment_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "status":

                rows = conn.execute(
                    f"""
                    SELECT
                        status,
                        COUNT(*) AS count
                    FROM appointments
                    {where_sql}
                    GROUP BY status
                    ORDER BY status
                    """,
                    params
                ).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "doctor":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            doctors.full_name,
                            'Unknown Doctor'
                        ) AS doctor,
                        COUNT(*) AS count
                    FROM appointments
                    LEFT JOIN doctors
                        ON doctors.id =
                           appointments.doctor_id
                    {where_sql}
                    GROUP BY appointments.doctor_id
                    ORDER BY doctor
                    """,
                    params
                ).fetchall()

                labels = [
                    row["doctor"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute(
                    f"""
                    SELECT
                        appointment_date,
                        COUNT(*) AS count
                    FROM appointments
                    {where_sql}
                    GROUP BY appointment_date
                    ORDER BY appointment_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["appointment_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # MEDICAL RECORDS
        # =====================================================

        elif dataset == "medical_records":

            title = "Medical Records Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "visit_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "visit_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "diagnosis":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            NULLIF(diagnosis, ''),
                            'Not Specified'
                        ) AS diagnosis,
                        COUNT(*) AS count
                    FROM medical_records
                    {where_sql}
                    GROUP BY diagnosis
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["diagnosis"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "doctor":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            doctors.full_name,
                            'Unknown Doctor'
                        ) AS doctor,
                        COUNT(*) AS count
                    FROM medical_records
                    LEFT JOIN doctors
                        ON doctors.id =
                           medical_records.doctor_id
                    {where_sql}
                    GROUP BY medical_records.doctor_id
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["doctor"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute(
                    f"""
                    SELECT
                        visit_date,
                        COUNT(*) AS count
                    FROM medical_records
                    {where_sql}
                    GROUP BY visit_date
                    ORDER BY visit_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["visit_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # LABORATORY
        # =====================================================

        elif dataset == "laboratory":

            title = "Laboratory Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "laboratory_requests.request_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "laboratory_requests.request_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "status":

                rows = conn.execute(
                    f"""
                    SELECT
                        laboratory_requests.status,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    {where_sql}
                    GROUP BY laboratory_requests.status
                    ORDER BY laboratory_requests.status
                    """,
                    params
                ).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "date":

                rows = conn.execute(
                    f"""
                    SELECT
                        laboratory_requests.request_date,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    {where_sql}
                    GROUP BY laboratory_requests.request_date
                    ORDER BY laboratory_requests.request_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["request_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            laboratory_tests.test_name,
                            'Unknown Test'
                        ) AS test_name,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    LEFT JOIN laboratory_tests
                        ON laboratory_tests.id =
                           laboratory_requests.test_id
                    {where_sql}
                    GROUP BY laboratory_requests.test_id
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["test_name"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # PHARMACY
        # =====================================================

        elif dataset == "pharmacy":

            title = "Pharmacy Analytics"

            if x_variable == "status":

                rows = conn.execute("""
                    SELECT
                        status,
                        COUNT(*) AS count
                    FROM prescriptions
                    GROUP BY status
                    ORDER BY status
                """).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "category":

                rows = conn.execute("""
                    SELECT
                        COALESCE(
                            category,
                            'Uncategorized'
                        ) AS category,
                        SUM(quantity) AS quantity
                    FROM medicines
                    GROUP BY category
                    ORDER BY category
                """).fetchall()

                labels = [
                    row["category"]
                    for row in rows
                ]

                values = [
                    row["quantity"] or 0
                    for row in rows
                ]

            else:

                rows = conn.execute("""
                    SELECT
                        medicine_name,
                        quantity
                    FROM medicines
                    ORDER BY medicine_name
                """).fetchall()

                labels = [
                    row["medicine_name"]
                    for row in rows
                ]

                values = [
                    row["quantity"] or 0
                    for row in rows
                ]

        else:

            conn.close()

            return None

        conn.close()

        return {

            "dataset": dataset,

            "analysis": analysis,

            "x_variable": x_variable,

            "y_variable": y_variable,

            "title": title,

            "labels": labels,

            "values": values,

            "date_from": date_from,

            "date_to": date_to

        }

    except Exception:

        conn.close()

        return None


# =========================================================
# ANALYTICS PDF EXPORT
# =========================================================

@app.route("/analytics/export/pdf")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def analytics_export_pdf():

    data = get_analytics_export_data()

    if data is None:

        return "Invalid analytics dataset.", 400

    buffer = BytesIO()

    document = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )

    story = []

    story.append(
        pdf_title(
            "HOSPITAL DATA ANALYTICS SYSTEM"
        )
    )

    story.append(
        Paragraph(
            data["title"],
            getSampleStyleSheet()["Heading2"]
        )
    )

    if data["date_from"] or data["date_to"]:

        date_text = (
            f"Date range: "
            f"{data['date_from'] or 'Beginning'} "
            f"to "
            f"{data['date_to'] or 'Latest'}"
        )

        story.append(
            Paragraph(
                date_text,
                getSampleStyleSheet()["Normal"]
            )
        )

    story.append(
        Spacer(1, 18)
    )

    rows = [
        [
            "Category",
            "Value"
        ]
    ]

    for label, value in zip(
        data["labels"],
        data["values"]
    ):

        rows.append([
            label,
            value
        ])

    if len(rows) == 1:

        rows.append([
            "No data",
            0
        ])

    story.append(
        pdf_table(
            rows,
            [350, 120]
        )
    )

    story.append(
        Spacer(1, 18)
    )

    story.append(
        Paragraph(
            f"Total categories: {len(data['labels'])}",
            getSampleStyleSheet()["Normal"]
        )
    )

    document.build(story)

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="analytics_report.pdf",
        mimetype="application/pdf"
    )


# =========================================================
# ANALYTICS EXCEL EXPORT
# =========================================================

@app.route("/analytics/export/excel")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def analytics_export_excel():

    data = get_analytics_export_data()

    if data is None:

        return "Invalid analytics dataset.", 400

    rows = [
        [
            "Category",
            "Value"
        ]
    ]

    for label, value in zip(
        data["labels"],
        data["values"]
    ):

        rows.append([
            label,
            value
        ])

    if len(rows) == 1:

        rows.append([
            "No data",
            0
        ])

    sheets = [
        (
            "Analytics",
            rows
        )
    ]

    output = excel_file(
        sheets,
        "analytics_report.xlsx"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name="analytics_report.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        )
    )


# =========================================================
# ANALYTICS CSV EXPORT
# =========================================================

@app.route("/analytics/export/csv")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def analytics_export_csv():

    data = get_analytics_export_data()

    if data is None:

        return "Invalid analytics dataset.", 400

    rows = [
        [
            "Category",
            "Value"
        ]
    ]

    for label, value in zip(
        data["labels"],
        data["values"]
    ):

        rows.append([
            label,
            value
        ])

    if len(rows) == 1:

        rows.append([
            "No data",
            0
        ])

    output = csv_file(rows)

    return send_file(
        output,
        as_attachment=True,
        download_name="analytics_report.csv",
        mimetype="text/csv"
    )


# =========================================================
# DAILY PDF
# =========================================================

@app.route("/reports/daily/pdf")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def daily_report_pdf():

    report_date = parse_report_date(
        request.args.get("report_date"),
        date.today().isoformat()
    )

    data = get_daily_report_data(
        report_date
    )

    return send_file(
        build_daily_pdf(data),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=(
            f"daily_report_{report_date}.pdf"
        )
    )


# =========================================================
# DAILY EXCEL
# =========================================================

@app.route("/reports/daily/excel")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def daily_report_excel():

    report_date = parse_report_date(
        request.args.get("report_date"),
        date.today().isoformat()
    )

    data = get_daily_report_data(
        report_date
    )

    sheets = [
        (
            "Summary",
            [
                [
                    "Daily Report",
                    report_date
                ],

                [],

                [
                    "Metric",
                    "Value"
                ],

                [
                    "Total Patients",
                    data["total_patients"]
                ],

                [
                    "Total Appointments",
                    data["total_appointments"]
                ],

                [
                    "Medical Records",
                    data["total_medical_records"]
                ],

                [
                    "Diagnoses",
                    data["total_diagnoses"]
                ]
            ]
        ),

        (
            "Appointments",
            [
                [
                    "Date",
                    "Time",
                    "Patient",
                    "Doctor",
                    "Reason",
                    "Status"
                ]
            ]
            +
            [
                [
                    row["appointment_date"],
                    row["appointment_time"],
                    row["patient_name"],
                    row["doctor_name"],
                    row["reason"],
                    row["status"]
                ]

                for row in data["appointments"]
            ]
        ),

        (
            "Medical Records",
            [
                [
                    "Date",
                    "Patient",
                    "Doctor",
                    "Chief Complaint",
                    "Diagnosis",
                    "Symptoms",
                    "Treatment",
                    "Prescription",
                    "Notes"
                ]
            ]
            +
            [
                [
                    row["visit_date"],
                    row["patient_name"],
                    row["doctor_name"],
                    row["chief_complaint"],
                    row["diagnosis"],
                    row["symptoms"],
                    row["treatment"],
                    row["prescription"],
                    row["notes"]
                ]

                for row in data["medical_records"]
            ]
        )
    ]

    return send_file(
        excel_file(
            sheets,
            f"daily_report_{report_date}.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=(
            f"daily_report_{report_date}.xlsx"
        )
    )


# =========================================================
# DAILY CSV
# =========================================================

@app.route("/reports/daily/csv")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def daily_report_csv():

    report_date = parse_report_date(
        request.args.get("report_date"),
        date.today().isoformat()
    )

    data = get_daily_report_data(
        report_date
    )

    rows = [
        [
            "HOSPITAL DATA ANALYTICS SYSTEM"
        ],

        [
            "DAILY REPORT",
            report_date
        ],

        [],

        [
            "SUMMARY"
        ],

        [
            "Metric",
            "Value"
        ],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ],

        [],

        [
            "APPOINTMENTS"
        ],

        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    rows += [
        [
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ]

        for row in data["appointments"]
    ]

    rows += [
        [],

        [
            "MEDICAL RECORDS"
        ],

        [
            "Date",
            "Patient",
            "Doctor",
            "Chief Complaint",
            "Diagnosis",
            "Symptoms",
            "Treatment",
            "Prescription",
            "Notes"
        ]
    ]

    rows += [
        [
            row["visit_date"],
            row["patient_name"],
            row["doctor_name"],
            row["chief_complaint"],
            row["diagnosis"],
            row["symptoms"],
            row["treatment"],
            row["prescription"],
            row["notes"]
        ]

        for row in data["medical_records"]
    ]

    return send_file(
        csv_file(rows),
        mimetype="text/csv",
        as_attachment=True,
        download_name=(
            f"daily_report_{report_date}.csv"
        )
    )


# =========================================================
# WEEKLY PDF
# =========================================================

@app.route("/reports/weekly/pdf")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def weekly_report_pdf():

    today = date.today()

    default_start = (
        today -
        timedelta(
            days=today.weekday()
        )
    )

    default_end = (
        default_start +
        timedelta(days=6)
    )

    start_date = parse_report_date(
        request.args.get("start_date"),
        default_start.isoformat()
    )

    end_date = parse_report_date(
        request.args.get("end_date"),
        default_end.isoformat()
    )

    if start_date > end_date:

        start_date, end_date = (
            end_date,
            start_date
        )

    data = get_weekly_report_data(
        start_date,
        end_date
    )

    return send_file(
        build_weekly_pdf(data),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=(
            f"weekly_report_"
            f"{start_date}_to_{end_date}.pdf"
        )
    )


# =========================================================
# WEEKLY EXCEL
# =========================================================

@app.route("/reports/weekly/excel")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def weekly_report_excel():

    today = date.today()

    default_start = (
        today -
        timedelta(
            days=today.weekday()
        )
    )

    default_end = (
        default_start +
        timedelta(days=6)
    )

    start_date = parse_report_date(
        request.args.get("start_date"),
        default_start.isoformat()
    )

    end_date = parse_report_date(
        request.args.get("end_date"),
        default_end.isoformat()
    )

    if start_date > end_date:

        start_date, end_date = (
            end_date,
            start_date
        )

    data = get_weekly_report_data(
        start_date,
        end_date
    )

    sheets = [
        (
            "Summary",
            [
                [
                    "Weekly Report",
                    f"{start_date} to {end_date}"
                ],

                [],

                [
                    "Metric",
                    "Value"
                ],

                [
                    "Total Patients",
                    data["total_patients"]
                ],

                [
                    "Total Appointments",
                    data["total_appointments"]
                ],

                [
                    "Medical Records",
                    data["total_medical_records"]
                ],

                [
                    "Diagnoses",
                    data["total_diagnoses"]
                ]
            ]
        ),

        (
            "Daily Summary",
            [
                [
                    "Date",
                    "Patients",
                    "Appointments",
                    "Medical Records",
                    "Diagnoses"
                ]
            ]
            +
            [
                [
                    row["date"],
                    row["patients"],
                    row["appointments"],
                    row["medical_records"],
                    row["diagnoses"]
                ]

                for row in data["daily_data"]
            ]
        ),

        (
            "Appointments",
            [
                [
                    "Date",
                    "Time",
                    "Patient",
                    "Doctor",
                    "Reason",
                    "Status"
                ]
            ]
            +
            [
                [
                    row["appointment_date"],
                    row["appointment_time"],
                    row["patient_name"],
                    row["doctor_name"],
                    row["reason"],
                    row["status"]
                ]

                for row in data["appointments"]
            ]
        )
    ]

    return send_file(
        excel_file(
            sheets,
            f"weekly_report_"
            f"{start_date}_to_{end_date}.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=(
            f"weekly_report_"
            f"{start_date}_to_{end_date}.xlsx"
        )
    )


# =========================================================
# WEEKLY CSV
# =========================================================

@app.route("/reports/weekly/csv")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def weekly_report_csv():

    today = date.today()

    default_start = (
        today -
        timedelta(
            days=today.weekday()
        )
    )

    default_end = (
        default_start +
        timedelta(days=6)
    )

    start_date = parse_report_date(
        request.args.get("start_date"),
        default_start.isoformat()
    )

    end_date = parse_report_date(
        request.args.get("end_date"),
        default_end.isoformat()
    )

    if start_date > end_date:

        start_date, end_date = (
            end_date,
            start_date
        )

    data = get_weekly_report_data(
        start_date,
        end_date
    )

    rows = [
        [
            "HOSPITAL DATA ANALYTICS SYSTEM"
        ],

        [
            "WEEKLY REPORT",
            start_date,
            end_date
        ],

        [],

        [
            "SUMMARY"
        ],

        [
            "Metric",
            "Value"
        ],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ],

        [],

        [
            "DAILY SUMMARY"
        ],

        [
            "Date",
            "Patients",
            "Appointments",
            "Medical Records",
            "Diagnoses"
        ]
    ]

    rows += [
        [
            row["date"],
            row["patients"],
            row["appointments"],
            row["medical_records"],
            row["diagnoses"]
        ]

        for row in data["daily_data"]
    ]

    rows += [
        [],

        [
            "APPOINTMENTS"
        ],

        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    rows += [
        [
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ]

        for row in data["appointments"]
    ]

    return send_file(
        csv_file(rows),
        mimetype="text/csv",
        as_attachment=True,
        download_name=(
            f"weekly_report_"
            f"{start_date}_to_{end_date}.csv"
        )
    )


# =========================================================
# MONTHLY PDF
# =========================================================

@app.route("/reports/monthly/pdf")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def monthly_report_pdf():

    today = date.today()

    try:

        month = int(
            request.args.get(
                "month",
                today.month
            )
        )

    except ValueError:

        month = today.month

    try:

        year = int(
            request.args.get(
                "year",
                today.year
            )
        )

    except ValueError:

        year = today.year

    if month < 1 or month > 12:

        month = today.month

    if year < 2000 or year > 2100:

        year = today.year

    data = get_monthly_report_data(
        month,
        year
    )

    return send_file(
        build_monthly_pdf(data),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=(
            f"monthly_report_"
            f"{year}_{month:02d}.pdf"
        )
    )


# =========================================================
# MONTHLY EXCEL
# =========================================================

@app.route("/reports/monthly/excel")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def monthly_report_excel():

    today = date.today()

    try:

        month = int(
            request.args.get(
                "month",
                today.month
            )
        )

    except ValueError:

        month = today.month

    try:

        year = int(
            request.args.get(
                "year",
                today.year
            )
        )

    except ValueError:

        year = today.year

    if month < 1 or month > 12:

        month = today.month

    if year < 2000 or year > 2100:

        year = today.year

    data = get_monthly_report_data(
        month,
        year
    )

    sheets = [
        (
            "Summary",
            [
                [
                    "Monthly Report",
                    f"{year}-{month:02d}"
                ],

                [],

                [
                    "Metric",
                    "Value"
                ],

                [
                    "Total Patients",
                    data["total_patients"]
                ],

                [
                    "Total Appointments",
                    data["total_appointments"]
                ],

                [
                    "Medical Records",
                    data["total_medical_records"]
                ],

                [
                    "Diagnoses",
                    data["total_diagnoses"]
                ]
            ]
        ),

        (
            "Daily Summary",
            [
                [
                    "Date",
                    "Patients",
                    "Appointments",
                    "Medical Records",
                    "Diagnoses"
                ]
            ]
            +
            [
                [
                    row["date"],
                    row["patients"],
                    row["appointments"],
                    row["medical_records"],
                    row["diagnoses"]
                ]

                for row in data["daily_data"]
            ]
        ),

        (
            "Diagnoses",
            [
                [
                    "Diagnosis",
                    "Count"
                ]
            ]
            +
            [
                [
                    row["diagnosis"],
                    row["count"]
                ]

                for row in data["diagnoses"]
            ]
        ),

        (
            "Appointments",
            [
                [
                    "Date",
                    "Time",
                    "Patient",
                    "Doctor",
                    "Reason",
                    "Status"
                ]
            ]
            +
            [
                [
                    row["appointment_date"],
                    row["appointment_time"],
                    row["patient_name"],
                    row["doctor_name"],
                    row["reason"],
                    row["status"]
                ]

                for row in data["appointments"]
            ]
        )
    ]

    return send_file(
        excel_file(
            sheets,
            f"monthly_report_"
            f"{year}_{month:02d}.xlsx"
        ),
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=(
            f"monthly_report_"
            f"{year}_{month:02d}.xlsx"
        )
    )


# =========================================================
# MONTHLY CSV
# =========================================================

@app.route("/reports/monthly/csv")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def monthly_report_csv():

    today = date.today()

    try:

        month = int(
            request.args.get(
                "month",
                today.month
            )
        )

    except ValueError:

        month = today.month

    try:

        year = int(
            request.args.get(
                "year",
                today.year
            )
        )

    except ValueError:

        year = today.year

    if month < 1 or month > 12:

        month = today.month

    if year < 2000 or year > 2100:

        year = today.year

    data = get_monthly_report_data(
        month,
        year
    )

    rows = [
        [
            "HOSPITAL DATA ANALYTICS SYSTEM"
        ],

        [
            "MONTHLY REPORT",
            f"{year}-{month:02d}"
        ],

        [],

        [
            "SUMMARY"
        ],

        [
            "Metric",
            "Value"
        ],

        [
            "Total Patients",
            data["total_patients"]
        ],

        [
            "Total Appointments",
            data["total_appointments"]
        ],

        [
            "Medical Records",
            data["total_medical_records"]
        ],

        [
            "Diagnoses",
            data["total_diagnoses"]
        ],

        [],

        [
            "DAILY SUMMARY"
        ],

        [
            "Date",
            "Patients",
            "Appointments",
            "Medical Records",
            "Diagnoses"
        ]
    ]

    rows += [
        [
            row["date"],
            row["patients"],
            row["appointments"],
            row["medical_records"],
            row["diagnoses"]
        ]

        for row in data["daily_data"]
    ]

    rows += [
        [],

        [
            "DIAGNOSES"
        ],

        [
            "Diagnosis",
            "Count"
        ]
    ]

    rows += [
        [
            row["diagnosis"],
            row["count"]
        ]

        for row in data["diagnoses"]
    ]

    rows += [
        [],

        [
            "APPOINTMENTS"
        ],

        [
            "Date",
            "Time",
            "Patient",
            "Doctor",
            "Reason",
            "Status"
        ]
    ]

    rows += [
        [
            row["appointment_date"],
            row["appointment_time"],
            row["patient_name"],
            row["doctor_name"],
            row["reason"],
            row["status"]
        ]

        for row in data["appointments"]
    ]

    return send_file(
        csv_file(rows),
        mimetype="text/csv",
        as_attachment=True,
        download_name=(
            f"monthly_report_"
            f"{year}_{month:02d}.csv"
        )
    )


# =========================================================
# PATIENT REPORT
# TEMPORARY REPORT PAGE
# =========================================================

@app.route("/reports/patient")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def patient_report():

    return """
        <h1>Patient Report</h1>

        <p>
            Patient Report module will be connected
            in the next Reports stage.
        </p>

        <a href="/reports">
            Back to Reports
        </a>
    """


# =========================================================
# APPOINTMENT REPORT
# TEMPORARY REPORT PAGE
# =========================================================

@app.route("/reports/appointment")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def appointment_report():

    return """
        <h1>Appointment Report</h1>

        <p>
            Appointment Report module will be connected
            in the next Reports stage.
        </p>

        <a href="/reports">
            Back to Reports
        </a>
    """


# =========================================================
# DOCTOR REPORT
# TEMPORARY REPORT PAGE
# =========================================================

@app.route("/reports/doctor")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def doctor_report():

    return """
        <h1>Doctor Report</h1>

        <p>
            Doctor Report module will be connected
            in the next Reports stage.
        </p>

        <a href="/reports">
            Back to Reports
        </a>
    """


# =========================================================
# DIAGNOSIS REPORT
# TEMPORARY REPORT PAGE
# =========================================================

@app.route("/reports/diagnosis")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def diagnosis_report():

    return """
        <h1>Disease / Diagnosis Report</h1>

        <p>
            Disease and Diagnosis Report module
            will be connected in the next Reports stage.
        </p>

        <a href="/reports">
            Back to Reports
        </a>
    """


# =========================================================
# HOSPITAL STATISTICS
# TEMPORARY REPORT PAGE
# =========================================================

@app.route("/reports/statistics")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def hospital_statistics():

    return """
        <h1>Hospital Statistics</h1>

        <p>
            Hospital Statistics module will be connected
            in the next Reports stage.
        </p>

        <a href="/reports">
            Back to Reports
        </a>
    """


# =========================================================
# ANALYTICS
# =========================================================

@app.route("/analytics")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def analytics():

    conn = get_db_connection()

    gender_data = conn.execute("""
        SELECT
            gender,
            COUNT(*) AS total
        FROM patients
        GROUP BY gender
    """).fetchall()

    age_group_data = conn.execute("""
        SELECT

            CASE

                WHEN age < 18
                    THEN 'Below 18'

                WHEN age BETWEEN 18 AND 30
                    THEN '18-30'

                WHEN age BETWEEN 31 AND 45
                    THEN '31-45'

                WHEN age BETWEEN 46 AND 60
                    THEN '46-60'

                ELSE 'Above 60'

            END AS age_group,

            COUNT(*) AS total

        FROM patients

        GROUP BY age_group

        ORDER BY MIN(age)
    """).fetchall()

    status_data = conn.execute("""
        SELECT
            status,
            COUNT(*) AS total
        FROM appointments
        GROUP BY status
    """).fetchall()

    specialization_data = conn.execute("""
        SELECT
            specialization,
            COUNT(*) AS total
        FROM doctors
        GROUP BY specialization
    """).fetchall()

    department_data = conn.execute("""
        SELECT
            department,
            COUNT(*) AS total
        FROM doctors
        GROUP BY department
    """).fetchall()

    doctor_workload = conn.execute("""
        SELECT

            doctors.full_name AS doctor_name,

            COUNT(appointments.id) AS total

        FROM doctors

        LEFT JOIN appointments
        ON doctors.id = appointments.doctor_id

        GROUP BY doctors.id

        ORDER BY total DESC
    """).fetchall()

    patient_frequency = conn.execute("""
        SELECT

            patients.full_name AS patient_name,

            COUNT(appointments.id) AS total

        FROM patients

        LEFT JOIN appointments
        ON patients.id = appointments.patient_id

        GROUP BY patients.id

        ORDER BY total DESC
    """).fetchall()

    appointments_by_date = conn.execute("""
        SELECT

            appointment_date,

            COUNT(*) AS total

        FROM appointments

        GROUP BY appointment_date

        ORDER BY appointment_date
    """).fetchall()

    appointments_by_month = conn.execute("""
        SELECT

            substr(
                appointment_date,
                1,
                7
            ) AS month,

            COUNT(*) AS total

        FROM appointments

        GROUP BY month

        ORDER BY month
    """).fetchall()

    appointments_by_day = conn.execute("""
        SELECT

            CASE strftime(
                '%w',
                appointment_date
            )

                WHEN '0' THEN 'Sunday'

                WHEN '1' THEN 'Monday'

                WHEN '2' THEN 'Tuesday'

                WHEN '3' THEN 'Wednesday'

                WHEN '4' THEN 'Thursday'

                WHEN '5' THEN 'Friday'

                WHEN '6' THEN 'Saturday'

            END AS day_name,

            COUNT(*) AS total

        FROM appointments

        GROUP BY day_name

        ORDER BY
            strftime(
                '%w',
                appointment_date
            )
    """).fetchall()

    doctor_workload_by_date = conn.execute("""
        SELECT

            appointment_date,

            doctors.full_name AS doctor_name,

            COUNT(appointments.id) AS total

        FROM appointments

        JOIN doctors
        ON appointments.doctor_id =
           doctors.id

        GROUP BY
            appointment_date,
            doctors.id

        ORDER BY appointment_date
    """).fetchall()

    patient_visits_by_date = conn.execute("""
        SELECT

            appointment_date,

            patients.full_name AS patient_name,

            COUNT(appointments.id) AS total

        FROM appointments

        JOIN patients
        ON appointments.patient_id =
           patients.id

        GROUP BY
            appointment_date,
            patients.id

        ORDER BY appointment_date
    """).fetchall()

    appointment_status_by_date = conn.execute("""
        SELECT

            appointment_date,

            status,

            COUNT(*) AS total

        FROM appointments

        GROUP BY
            appointment_date,
            status

        ORDER BY appointment_date
    """).fetchall()

    laboratory_status_data = conn.execute("""
        SELECT

            status,

            COUNT(*) AS total

        FROM laboratory_requests

        GROUP BY status
    """).fetchall()

    laboratory_test_data = conn.execute("""
        SELECT

            laboratory_tests.test_name,

            COUNT(laboratory_requests.id) AS total

        FROM laboratory_tests

        LEFT JOIN laboratory_requests
        ON laboratory_tests.id =
           laboratory_requests.test_id

        GROUP BY laboratory_tests.id

        ORDER BY total DESC
    """).fetchall()

    laboratory_specialty_data = conn.execute("""
        SELECT

            COALESCE(
                laboratory_results.recommended_specialty,
                laboratory_tests.recommended_specialty,
                'Not Assigned'
            ) AS specialty,

            COUNT(laboratory_requests.id) AS total

        FROM laboratory_requests

        JOIN laboratory_tests
        ON laboratory_requests.test_id =
           laboratory_tests.id

        LEFT JOIN laboratory_results
        ON laboratory_requests.id =
           laboratory_results.request_id

        GROUP BY specialty

        ORDER BY total DESC
    """).fetchall()

    conn.close()

    return render_template(

        "analytics.html",

        gender_data=gender_data,

        age_group_data=age_group_data,

        status_data=status_data,

        specialization_data=specialization_data,

        department_data=department_data,

        doctor_workload=doctor_workload,

        patient_frequency=patient_frequency,

        appointments_by_date=appointments_by_date,

        appointments_by_month=appointments_by_month,

        appointments_by_day=appointments_by_day,

        doctor_workload_by_date=doctor_workload_by_date,

        patient_visits_by_date=patient_visits_by_date,

        appointment_status_by_date=appointment_status_by_date,

        laboratory_status_data=laboratory_status_data,

        laboratory_test_data=laboratory_test_data,

        laboratory_specialty_data=laboratory_specialty_data

    )


# =========================================================
# ANALYTICS DATA API
# =========================================================

@app.route("/analytics/data")
@role_required(
    "Admin",
    "Doctor",
    "Data Analyst"
)
def analytics_data():

    dataset = request.args.get(
        "dataset",
        "patients"
    )

    analysis = request.args.get(
        "analysis",
        "bar"
    )

    x_variable = request.args.get(
        "x_variable",
        "gender"
    )

    y_variable = request.args.get(
        "y_variable",
        "count"
    )

    date_from = request.args.get(
        "date_from"
    )

    date_to = request.args.get(
        "date_to"
    )

    conn = get_db_connection()

    labels = []
    values = []
    title = ""

    try:

        # =====================================================
        # PATIENTS
        # =====================================================

        if dataset == "patients":

            title = "Patients Analytics"

            if x_variable == "age":

                rows = conn.execute("""
                    SELECT age, COUNT(*) AS count
                    FROM patients
                    GROUP BY age
                    ORDER BY age
                """).fetchall()

                labels = [
                    row["age"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "gender":

                rows = conn.execute("""
                    SELECT
                        gender,
                        COUNT(*) AS count
                    FROM patients
                    GROUP BY gender
                    ORDER BY gender
                """).fetchall()

                labels = [
                    row["gender"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute("""
                    SELECT
                        COUNT(*) AS count
                    FROM patients
                """).fetchone()

                labels = ["Patients"]
                values = [rows["count"]]

        # =====================================================
        # DOCTORS
        # =====================================================

        elif dataset == "doctors":

            title = "Doctors Analytics"

            if x_variable == "specialization":

                rows = conn.execute("""
                    SELECT
                        specialization,
                        COUNT(*) AS count
                    FROM doctors
                    GROUP BY specialization
                    ORDER BY specialization
                """).fetchall()

                labels = [
                    row["specialization"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "gender":

                rows = conn.execute("""
                    SELECT
                        gender,
                        COUNT(*) AS count
                    FROM doctors
                    GROUP BY gender
                    ORDER BY gender
                """).fetchall()

                labels = [
                    row["gender"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                row = conn.execute("""
                    SELECT COUNT(*) AS count
                    FROM doctors
                """).fetchone()

                labels = ["Doctors"]
                values = [row["count"]]

        # =====================================================
        # APPOINTMENTS
        # =====================================================

        elif dataset == "appointments":

            title = "Appointments Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "appointment_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "appointment_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "status":

                rows = conn.execute(
                    f"""
                    SELECT
                        status,
                        COUNT(*) AS count
                    FROM appointments
                    {where_sql}
                    GROUP BY status
                    ORDER BY status
                    """,
                    params
                ).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "doctor":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            doctors.full_name,
                            'Unknown Doctor'
                        ) AS doctor,
                        COUNT(*) AS count
                    FROM appointments
                    LEFT JOIN doctors
                        ON doctors.id =
                           appointments.doctor_id
                    {where_sql}
                    GROUP BY appointments.doctor_id
                    ORDER BY doctor
                    """,
                    params
                ).fetchall()

                labels = [
                    row["doctor"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "date":

                rows = conn.execute(
                    f"""
                    SELECT
                        appointment_date,
                        COUNT(*) AS count
                    FROM appointments
                    {where_sql}
                    GROUP BY appointment_date
                    ORDER BY appointment_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["appointment_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                row = conn.execute(
                    f"""
                    SELECT COUNT(*) AS count
                    FROM appointments
                    {where_sql}
                    """,
                    params
                ).fetchone()

                labels = ["Appointments"]
                values = [row["count"]]

        # =====================================================
        # MEDICAL RECORDS
        # =====================================================

        elif dataset == "medical_records":

            title = "Medical Records Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "visit_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "visit_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "diagnosis":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            NULLIF(diagnosis, ''),
                            'Not Specified'
                        ) AS diagnosis,
                        COUNT(*) AS count
                    FROM medical_records
                    {where_sql}
                    GROUP BY diagnosis
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["diagnosis"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "doctor":

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            doctors.full_name,
                            'Unknown Doctor'
                        ) AS doctor,
                        COUNT(*) AS count
                    FROM medical_records
                    LEFT JOIN doctors
                        ON doctors.id =
                           medical_records.doctor_id
                    {where_sql}
                    GROUP BY medical_records.doctor_id
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["doctor"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "date":

                rows = conn.execute(
                    f"""
                    SELECT
                        visit_date,
                        COUNT(*) AS count
                    FROM medical_records
                    {where_sql}
                    GROUP BY visit_date
                    ORDER BY visit_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["visit_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                row = conn.execute(
                    f"""
                    SELECT COUNT(*) AS count
                    FROM medical_records
                    {where_sql}
                    """,
                    params
                ).fetchone()

                labels = ["Medical Records"]
                values = [row["count"]]

        # =====================================================
        # LABORATORY
        # =====================================================

        elif dataset == "laboratory":

            title = "Laboratory Analytics"

            where = []
            params = []

            if date_from:

                where.append(
                    "laboratory_requests.request_date >= ?"
                )

                params.append(date_from)

            if date_to:

                where.append(
                    "laboratory_requests.request_date <= ?"
                )

                params.append(date_to)

            where_sql = ""

            if where:

                where_sql = (
                    "WHERE " +
                    " AND ".join(where)
                )

            if x_variable == "status":

                rows = conn.execute(
                    f"""
                    SELECT
                        laboratory_requests.status,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    {where_sql}
                    GROUP BY laboratory_requests.status
                    ORDER BY laboratory_requests.status
                    """,
                    params
                ).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            elif x_variable == "date":

                rows = conn.execute(
                    f"""
                    SELECT
                        laboratory_requests.request_date,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    {where_sql}
                    GROUP BY laboratory_requests.request_date
                    ORDER BY laboratory_requests.request_date
                    """,
                    params
                ).fetchall()

                labels = [
                    row["request_date"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute(
                    f"""
                    SELECT
                        COALESCE(
                            laboratory_tests.test_name,
                            'Unknown Test'
                        ) AS test_name,
                        COUNT(*) AS count
                    FROM laboratory_requests
                    LEFT JOIN laboratory_tests
                        ON laboratory_tests.id =
                           laboratory_requests.test_id
                    {where_sql}
                    GROUP BY laboratory_requests.test_id
                    ORDER BY count DESC
                    """,
                    params
                ).fetchall()

                labels = [
                    row["test_name"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

        # =====================================================
        # PHARMACY
        # =====================================================

        elif dataset == "pharmacy":

            title = "Pharmacy Analytics"

            if x_variable == "category":

                rows = conn.execute("""
                    SELECT
                        COALESCE(
                            category,
                            'Uncategorized'
                        ) AS category,
                        SUM(quantity) AS quantity
                    FROM medicines
                    GROUP BY category
                    ORDER BY category
                """).fetchall()

                labels = [
                    row["category"]
                    for row in rows
                ]

                values = [
                    row["quantity"] or 0
                    for row in rows
                ]

            elif x_variable == "medicine":

                rows = conn.execute("""
                    SELECT
                        medicine_name,
                        quantity
                    FROM medicines
                    ORDER BY medicine_name
                """).fetchall()

                labels = [
                    row["medicine_name"]
                    for row in rows
                ]

                values = [
                    row["quantity"] or 0
                    for row in rows
                ]

            elif x_variable == "status":

                rows = conn.execute("""
                    SELECT
                        status,
                        COUNT(*) AS count
                    FROM prescriptions
                    GROUP BY status
                    ORDER BY status
                """).fetchall()

                labels = [
                    row["status"]
                    for row in rows
                ]

                values = [
                    row["count"]
                    for row in rows
                ]

            else:

                rows = conn.execute("""
                    SELECT
                        medicine_name,
                        quantity
                    FROM medicines
                    ORDER BY medicine_name
                """).fetchall()

                labels = [
                    row["medicine_name"]
                    for row in rows
                ]

                values = [
                    row["quantity"] or 0
                    for row in rows
                ]

        # =====================================================
        # INVALID DATASET
        # =====================================================

        else:

            conn.close()

            return jsonify({
                "success": False,
                "message": "Invalid analytics dataset."
            }), 400

        conn.close()

        # =====================================================
        # EMPTY RESULT
        # =====================================================

        if not labels and not values:

            return jsonify({
                "success": False,
                "message":
                    "No data found for the selected dataset and filters."
            })

        return jsonify({

            "success": True,

            "dataset": dataset,

            "analysis": analysis,

            "x_variable": x_variable,

            "y_variable": y_variable,

            "title": title,

            "labels": labels,

            "values": values,

            "date_from": date_from,

            "date_to": date_to

        })

    except Exception as e:

        conn.close()

        return jsonify({
            "success": False,
            "message": str(e)
        }), 500


# =========================================================
# START APPLICATION
# =========================================================

if __name__ == "__main__":

    init_db()

    app.run(debug=True)