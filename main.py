import os
import csv
import io
import base64
from datetime import datetime
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from flask import Flask, render_template, request, send_file
from fpdf import FPDF
from docx import Document
from docx.shared import Inches
import sqlite3
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from flask import redirect, url_for

app = Flask(__name__)

# Define the database file name
DB_FILE = 'obsidian.db'

# Flask secret key — set OBSIDIAN_SECRET_KEY env var in production!
_secret = os.environ.get('OBSIDIAN_SECRET_KEY')
if not _secret:
    import warnings
    warnings.warn(
        "OBSIDIAN_SECRET_KEY env var not set. Using insecure dev key. "
        "Set this variable before deploying!", stacklevel=1
    )
    _secret = 'obsidian-dev-key-change-in-production-a8f3b2c1'
app.config['SECRET_KEY'] = _secret

login_manager = LoginManager()
login_manager.init_app(app)
# Tell the app where to send users if they try to access a locked page
login_manager.login_view = 'login'

# Create a temporary User object that Flask-Login understands
class User(UserMixin):
    def __init__(self, id, username):
        self.id = id
        self.username = username

# This function tells Flask how to find a user in your SQLite database using their session ID
@login_manager.user_loader
def load_user(user_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT id, username FROM users WHERE id = ?", (user_id,))
    user_data = cursor.fetchone()
    conn.close()
    
    if user_data:
        return User(id=user_data[0], username=user_data[1])
    return None

def get_audio_tracks():
    audio_dir = os.path.join(app.root_path, "static", "audio")
    if os.path.exists(audio_dir):
        files = [f for f in sorted(os.listdir(audio_dir)) if f.lower().endswith(('.mp3', '.wav', '.ogg', '.m4a'))]
        if files:
            return [{"name": f, "url": f"/static/audio/{f}"} for f in files]
    return [{"name": "song.mp3", "url": "/static/audio/song.mp3"}]

@app.context_processor
def inject_global_data():
    return {
        'audio_tracks': get_audio_tracks()
    }

# ==========================================
# DATA MANAGEMENT & PERSISTENCE
# ==========================================
CSV_FILE = "Athlete_Data.csv"
COLUMNS = [
    "Timestamp", "Athlete's Name", "Activity Type", "Duration (min)", "RPE", 
    "Average Heart Rate", "Top Speed (km/h)", "Distance (km)", 
    "Exercise Name", "Weight (kg)", "Sets", "Reps"
]

def initialize_csv():
    if not os.path.isfile(CSV_FILE):
        df = pd.DataFrame(columns=COLUMNS)
        df.to_csv(CSV_FILE, index=False)

def add_entry(data_dict):
    initialize_csv()
    with open(CSV_FILE, mode='a', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writerow(data_dict)

initialize_csv()

# --- NEW DATABASE INITIALIZATION ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS workouts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            athlete_name TEXT,
            activity_type TEXT,
            duration REAL,
            rpe REAL,
            avg_hr REAL,
            top_speed REAL,
            distance REAL,
            exercise_name TEXT,
            weight REAL,
            sets INTEGER,
            reps INTEGER,
            user_id INTEGER
        )
    ''')
    
    conn.commit()
    conn.close()

init_db()
# -----------------------------------

# --- HELPER TO READ DATABASE INTO PANDAS ---
def get_dataframe():
    try:
        conn = sqlite3.connect(DB_FILE)
        query = '''
            SELECT 
                timestamp as "Timestamp",
                athlete_name as "Athlete's Name",
                activity_type as "Activity Type",
                duration as "Duration (min)",
                rpe as "RPE",
                avg_hr as "Average Heart Rate",
                top_speed as "Top Speed (km/h)",
                distance as "Distance (km)",
                exercise_name as "Exercise Name",
                weight as "Weight (kg)",
                sets as "Sets",
                reps as "Reps"
            FROM workouts
        '''
        df = pd.read_sql_query(query, conn)
        conn.close()
        return df
    except Exception as e:
        app.logger.error(f"get_dataframe error: {e}")
        return pd.DataFrame()
# -------------------------------------------



def get_global_telemetry():
    empty_telemetry = {
        "total_sessions": 0,
        "athletes": [],
        "athletes_count": 0,
        "total_distance": "0.0",
        "total_tonnage": "0",
        "records": {},
        "recent_logs": []
    }
    
    # Read from the database instead of the CSV
    df = get_dataframe()

    if df.empty or "Athlete's Name" not in df.columns:
        return empty_telemetry

    df_clean = df.copy()
    df_clean["Clean_Name"] = df_clean["Athlete's Name"].astype(str).str.strip().str.title()
    df_clean = df_clean[df_clean["Clean_Name"] != ""]
    df_clean = df_clean[df_clean["Clean_Name"].str.lower() != "nan"]

    if df_clean.empty:
        return empty_telemetry

    total_sessions = len(df_clean)
    athletes = sorted([a for a in df_clean["Clean_Name"].unique() if a])
    athletes_count = len(athletes)

    dist_series = pd.to_numeric(df_clean["Distance (km)"], errors='coerce').fillna(0)
    total_distance = dist_series.sum()

    w = pd.to_numeric(df_clean["Weight (kg)"], errors='coerce').fillna(0)
    s = pd.to_numeric(df_clean["Sets"], errors='coerce').fillna(0)
    r = pd.to_numeric(df_clean["Reps"], errors='coerce').fillna(0)
    total_tonnage = (w * s * r).sum()

    records = {}

    # Peak Sprint Speed
    sprint_df = df_clean[df_clean["Activity Type"] == "Sprint"].copy()
    if not sprint_df.empty:
        sprint_speeds = pd.to_numeric(sprint_df["Top Speed (km/h)"], errors='coerce').dropna()
        if not sprint_speeds.empty:
            max_idx = sprint_speeds.idxmax()
            best_row = sprint_df.loc[max_idx]
            records["sprint"] = {
                "val": f"{sprint_speeds.loc[max_idx]:.1f} KM/H",
                "athlete": str(best_row["Clean_Name"]).upper(),
                "label": "TOP SPRINT VELOCITY",
                "icon": "⚡"
            }

    # Max Endurance Distance
    end_df = df_clean[df_clean["Activity Type"] == "Endurance"].copy()
    if not end_df.empty:
        end_dists = pd.to_numeric(end_df["Distance (km)"], errors='coerce').dropna()
        if not end_dists.empty:
            max_idx = end_dists.idxmax()
            best_row = end_df.loc[max_idx]
            records["endurance"] = {
                "val": f"{end_dists.loc[max_idx]:.1f} KM",
                "athlete": str(best_row["Clean_Name"]).upper(),
                "label": "MAX RUN DISTANCE",
                "icon": "🏃"
            }

    # Max Strength Lift
    str_df = df_clean[df_clean["Activity Type"] == "Strength"].copy()
    if not str_df.empty:
        str_w = pd.to_numeric(str_df["Weight (kg)"], errors='coerce').dropna()
        if not str_w.empty:
            max_idx = str_w.idxmax()
            best_row = str_df.loc[max_idx]
            exercise = str(best_row.get("Exercise Name", "Lift")).upper()
            records["strength"] = {
                "val": f"{str_w.loc[max_idx]:.1f} KG",
                "athlete": str(best_row["Clean_Name"]).upper(),
                "exercise": exercise,
                "label": "HEAVIEST RECORDED LIFT",
                "icon": "🏋️"
            }

    recent_logs = []
    for _, row in df_clean.tail(3).iloc[::-1].iterrows():
        act = str(row.get("Activity Type", "Sprint"))
        if act == "Sprint":
            spd = row.get("Top Speed (km/h)")
            dur = row.get("Duration (min)")
            metric = f"{spd} km/h" if pd.notna(spd) else "-"
            if pd.notna(dur):
                metric += f" ({dur}m)"
        elif act == "Endurance":
            dst = row.get("Distance (km)")
            dur = row.get("Duration (min)")
            metric = f"{dst} km" if pd.notna(dst) else "-"
            if pd.notna(dur):
                metric += f" ({dur}m)"
        else:
            w_val = row.get("Weight (kg)")
            ex = row.get("Exercise Name", "Lift")
            metric = f"{w_val} kg {ex}" if pd.notna(w_val) else "-"

        hr_str = f"{float(row['Average Heart Rate']):.0f} BPM" if pd.notna(row.get('Average Heart Rate')) and str(row.get('Average Heart Rate')).strip() != '' else "N/A"

        recent_logs.append({
            "athlete": str(row["Clean_Name"]).upper(),
            "activity": act.upper(),
            "metric": metric,
            "timestamp": str(row.get("Timestamp", ""))[:16],
            "hr": hr_str
        })

    return {
        "total_sessions": total_sessions,
        "athletes": athletes,
        "athletes_count": athletes_count,
        "total_distance": f"{total_distance:.1f}",
        "total_tonnage": f"{total_tonnage:,.0f}",
        "records": records,
        "recent_logs": recent_logs
    }
# ==========================================
# FLASK ROUTES
# ==========================================
@app.route('/register', methods=['GET', 'POST'])
def register():
    error = None
    if request.method == 'POST':
        # Username: strip whitespace. Password: do NOT strip (user chose those chars).
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        confirm  = request.form.get('confirm_password', '')

        if not username or not password:
            error = "Both username and password are required."
        elif len(username) < 3 or len(username) > 50:
            error = "Username must be between 3 and 50 characters."
        elif not username.replace('_', '').replace('-', '').isalnum():
            error = "Username may only contain letters, numbers, underscores, and hyphens."
        elif len(password) < 8:
            error = "Password must be at least 8 characters long."
        elif password != confirm:
            error = "Passwords do not match. Please try again."
        else:
            hashed_pw = generate_password_hash(password)
            try:
                conn = sqlite3.connect(DB_FILE)
                cursor = conn.cursor()
                cursor.execute('INSERT INTO users (username, password) VALUES (?, ?)', (username, hashed_pw))
                conn.commit()
                conn.close()
                return redirect(url_for('login'))
            except sqlite3.IntegrityError:
                error = "That username is already taken. Choose another."

    return render_template('register.html', error=error)

@app.route('/login', methods=['GET', 'POST'])
def login():
    # Already logged in? Skip to dashboard.
    if current_user.is_authenticated:
        return redirect(url_for('index'))

    error = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        # Do NOT strip the password — preserve the exact characters the user typed.
        password = request.form.get('password', '')

        if not username or not password:
            error = "Both username and password are required."
        else:
            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("SELECT id, username, password FROM users WHERE username = ?", (username,))
            user_data = cursor.fetchone()
            conn.close()

            if user_data and check_password_hash(user_data[2], password):
                user = User(id=user_data[0], username=user_data[1])
                login_user(user, remember=False)
                # Honour the 'next' param but validate it to prevent open-redirect
                next_page = request.args.get('next', '')
                if next_page and next_page.startswith('/'):
                    return redirect(next_page)
                return redirect(url_for('index'))
            else:
                error = "Invalid username or password."

    return render_template('login.html', error=error)

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/', methods=['GET'])
@login_required
def index():
    telemetry = get_global_telemetry()
    return render_template(
        'index.html', 
        summary=None, 
        chart_url=None, 
        errors={}, 
        dash_error=None,
        form_data={},
        dash_name='',
        dash_activity='Sprint',
        dash_range='5',
        telemetry=telemetry,
        kpis=None,
        history=None
    )

@app.route('/log', methods=['POST'])
@login_required
def log_session():
    errors = {}
    form_data = request.form.to_dict()
    
    name = request.form.get('name', '').strip()
    if not name: errors['name'] = "Athlete's name is required."
    
    activity = request.form.get('activity', 'Sprint').strip()
    hr_raw = request.form.get('hr', '').strip()
    rpe_raw = request.form.get('rpe', '').strip()

    hr = None
    try:
        hr = float(hr_raw)
        if hr < 30 or hr > 250: errors['hr'] = "Heart rate must be between 30 and 250 BPM."
    except ValueError:
        errors['hr'] = "Please enter a valid numeric heart rate."

    rpe = None
    try:
        rpe = float(rpe_raw)
        if rpe < 1.0 or rpe > 10.0: errors['rpe'] = "RPE must be a score between 1.0 and 10.0."
    except ValueError:
        errors['rpe'] = "Please enter a valid numeric RPE score."

    data = {col: "" for col in COLUMNS}
    data["Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    data["Athlete's Name"] = name
    data["Activity Type"] = activity
    data["Average Heart Rate"] = hr if hr is not None else ""
    data["RPE"] = rpe if rpe is not None else ""

    summary_text = ""

    if activity == "Sprint":
        speed_raw = request.form.get('speed', '').strip()
        dur_raw = (request.form.get('duration') or '').strip()

        speed = None
        try:
            speed = float(speed_raw)
            if speed <= 0: errors['speed'] = "Top speed must be greater than 0 km/h."
        except ValueError:
            errors['speed'] = "Please enter a valid numeric speed."

        duration = None
        try:
            duration = float(dur_raw)
            if duration <= 0: errors['duration'] = "Duration must be greater than 0 minutes."
        except ValueError:
            errors['duration'] = "Please enter a valid numeric duration."

        if not errors:
            data["Top Speed (km/h)"] = speed
            data["Duration (min)"] = duration
            speed_zone = "Elite Zone (>28 km/h)" if speed >= 28.0 else "Standard Pace"
            hr_strain = "Zone 5 (Max)" if hr >= 180 else "Zone 4 (High)" if hr >= 160 else "Zone 3 (Moderate)"
            load_score = rpe * duration

            summary_text = (
                f"SESSION REPORT: {name.upper()} | {activity.upper()}\n"
                f"Logged At: {data['Timestamp']}\n"
                f"{'='*40}\n"
                f"- Top Speed: {speed:.1f} km/h ({speed_zone})\n"
                f"- Cardio Strain: {hr_strain} ({hr:.0f} BPM)\n"
                f"- Session Load Score: {load_score:.1f} a.u."
            )

    elif activity == "Endurance":
        dist_raw = request.form.get('distance', '').strip()
        dur_raw = (request.form.get('duration') or request.form.get('duration_endurance') or '').strip()

        distance = None
        try:
            distance = float(dist_raw)
            if distance <= 0: errors['distance'] = "Distance must be greater than 0 km."
        except ValueError:
            errors['distance'] = "Please enter a valid numeric distance."

        duration = None
        try:
            duration = float(dur_raw)
            if duration <= 0: errors['duration'] = "Duration must be greater than 0 minutes."
        except ValueError:
            errors['duration'] = "Please enter a valid numeric duration."

        if not errors:
            data["Distance (km)"] = distance
            data["Duration (min)"] = duration
            raw_pace = duration / distance if distance > 0 else 0
            pace_mins = int(raw_pace)
            pace_secs = int((raw_pace - pace_mins) * 60)
            load_score = rpe * duration

            summary_text = (
                f"SESSION REPORT: {name.upper()} | {activity.upper()}\n"
                f"Logged At: {data['Timestamp']}\n"
                f"{'='*40}\n"
                f"- Average Pace: {pace_mins}:{pace_secs:02d} min/km\n"
                f"- Distance Covered: {distance:.2f} km\n"
                f"- Aerobic Strain: {hr:.0f} BPM\n"
                f"- Session Load Score: {load_score:.1f} a.u."
            )

    elif activity == "Strength":
        exercise = request.form.get('exercise', '').strip()
        weight_raw = request.form.get('weight', '').strip()
        sets_raw = request.form.get('sets', '').strip()
        reps_raw = request.form.get('reps', '').strip()

        if not exercise: errors['exercise'] = "Exercise name is required."

        weight = None
        try:
            weight = float(weight_raw)
            if weight <= 0: errors['weight'] = "Weight must be greater than 0 kg."
        except ValueError:
            errors['weight'] = "Please enter a valid numeric weight."

        sets = None
        try:
            sets = int(sets_raw)
            if sets <= 0: errors['sets'] = "Sets must be at least 1."
        except ValueError:
            errors['sets'] = "Please enter a valid integer for sets."

        reps = None
        try:
            reps = int(reps_raw)
            if reps <= 0: errors['reps'] = "Reps must be at least 1."
        except ValueError:
            errors['reps'] = "Please enter a valid integer for reps."

        if not errors:
            data["Exercise Name"] = exercise
            data["Weight (kg)"] = weight
            data["Sets"] = sets
            data["Reps"] = reps

            volume_load = weight * sets * reps
            est_1rm = weight * (1 + (reps / 30))
            zone = "Pure Strength" if reps <= 5 else "Hypertrophy" if reps <= 12 else "Muscular Endurance"

            summary_text = (
                f"SESSION REPORT: {name.upper()} | {activity.upper()}\n"
                f"Logged At: {data['Timestamp']}\n"
                f"{'='*40}\n"
                f"- Exercise: {exercise}\n"
                f"- Adaptation Focus: {zone} ({reps} reps/set)\n"
                f"- Total Volume Load: {volume_load:.1f} kg\n"
                f"- Estimated 1RM: {est_1rm:.1f} kg"
            )

    if errors:
        return render_template(
            'index.html', 
            summary=None, 
            chart_url=None, 
            errors=errors, 
            dash_error=None,
            form_data=form_data,
            dash_name='',
            dash_activity='Sprint',
            dash_range='5',
            telemetry=get_global_telemetry(),
            kpis=None,
            history=None
        )

    # --- SAVE TO SQLITE DATABASE INSTEAD OF CSV ---
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute('''
        INSERT INTO workouts (
            timestamp, athlete_name, activity_type, duration, rpe, avg_hr, 
            top_speed, distance, exercise_name, weight, sets, reps, user_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get("Timestamp"),
        name,
        data.get("Activity Type"),
        data.get("Duration (min)") or None,
        data.get("RPE") or None,
        data.get("Average Heart Rate") or None,
        data.get("Top Speed (km/h)") or None,
        data.get("Distance (km)") or None,
        data.get("Exercise Name") or None,
        data.get("Weight (kg)") or None,
        data.get("Sets") or None,
        data.get("Reps") or None,
        current_user.id  
    ))
    
    conn.commit()
    conn.close()
    # -------------------------------

    return render_template(
        'index.html', 
        summary=summary_text, 
        chart_url=None, 
        errors={}, 
        dash_error=None,
        form_data={},
        dash_name='',
        dash_activity='Sprint',
        dash_range='5',
        telemetry=get_global_telemetry(),
        kpis=None,
        history=None
    )

@app.route('/dashboard', methods=['POST'])
@login_required
def dashboard():
    athlete_name = request.form.get('dash_name', '').strip()
    activity = request.form.get('dash_activity', 'Sprint').strip()
    dash_range = request.form.get('dash_range', '5').strip()
    if dash_range not in ['5', '10', 'all']:
        dash_range = '5'

    telemetry = get_global_telemetry()

    if not athlete_name:
        return render_template(
            'index.html', 
            summary=None, 
            chart_url=None, 
            errors={}, 
            dash_error="Please enter an athlete name to render analytics.",
            form_data={},
            dash_name=athlete_name,
            dash_activity=activity,
            dash_range=dash_range,
            telemetry=telemetry,
            kpis=None,
            history=None
        )

    df = get_dataframe()

    if df.empty or "Athlete's Name" not in df.columns:
        return render_template(
            'index.html', 
            summary=None, 
            chart_url=None, 
            errors={}, 
            dash_error="No dataset found. Log a session first.",
            form_data={},
            dash_name=athlete_name,
            dash_activity=activity,
            dash_range=dash_range,
            telemetry=telemetry,
            kpis=None,
            history=None
        )

    all_filtered = df[
        (df["Athlete's Name"].astype(str).str.lower() == athlete_name.lower()) & 
        (df["Activity Type"] == activity)
    ]

    if all_filtered.empty:
        return render_template(
            'index.html', 
            summary=None, 
            chart_url=None, 
            errors={}, 
            dash_error=f"No {activity} sessions logged for athlete '{athlete_name}'.",
            form_data={},
            dash_name=athlete_name,
            dash_activity=activity,
            dash_range=dash_range,
            telemetry=telemetry,
            kpis=None,
            history=None
        )

    # --- 1. CALCULATE HIGH-IMPACT KPI METRICS ---
    kpi_metrics = {}
    total_discipline_sessions = len(all_filtered)
    kpi_metrics["total_sessions"] = f"{total_discipline_sessions}"

    hrs_all = pd.to_numeric(all_filtered["Average Heart Rate"], errors='coerce').dropna()
    if not hrs_all.empty:
        avg_hr_val = hrs_all.mean()
        if avg_hr_val >= 180: hr_zone = "ZONE 5 (MAX STRAIN)"
        elif avg_hr_val >= 160: hr_zone = "ZONE 4 (ANAEROBIC)"
        elif avg_hr_val >= 140: hr_zone = "ZONE 3 (AEROBIC)"
        else: hr_zone = "ZONE 2 (ACTIVE RECOVERY)"
        kpi_metrics["avg_hr"] = f"{avg_hr_val:.0f} BPM"
        kpi_metrics["hr_zone"] = hr_zone
    else:
        kpi_metrics["avg_hr"] = "N/A"
        kpi_metrics["hr_zone"] = "NO DATA"

    if activity == "Sprint":
        speeds = pd.to_numeric(all_filtered["Top Speed (km/h)"], errors='coerce').fillna(0)
        durations = pd.to_numeric(all_filtered["Duration (min)"], errors='coerce').fillna(0)
        rpes = pd.to_numeric(all_filtered["RPE"], errors='coerce').fillna(0)

        pr_speed = speeds.max()
        kpi_metrics["pr_value"] = f"{pr_speed:.1f} KM/H"
        kpi_metrics["pr_label"] = "ALL-TIME TOP SPEED"
        kpi_metrics["pr_sub"] = "ELITE TIER (>28)" if pr_speed >= 28.0 else "STANDARD TIER"

        kpi_metrics["volume_value"] = f"{durations.sum():.0f} MIN"
        kpi_metrics["volume_label"] = "TOTAL SPRINT TIME"
        kpi_metrics["volume_sub"] = f"{total_discipline_sessions} SESSIONS LOGGED"

        load_scores = rpes * durations
        avg_load = load_scores.mean() if not load_scores.empty else 0
        kpi_metrics["workload_value"] = f"{avg_load:.0f} A.U."
        kpi_metrics["workload_label"] = "AVG SESSION LOAD"
        kpi_metrics["workload_sub"] = "HIGH STRAIN" if avg_load >= 350 else "BALANCED LOAD"

    elif activity == "Endurance":
        distances = pd.to_numeric(all_filtered["Distance (km)"], errors='coerce').fillna(0)
        durations = pd.to_numeric(all_filtered["Duration (min)"], errors='coerce').fillna(0)
        rpes = pd.to_numeric(all_filtered["RPE"], errors='coerce').fillna(0)

        pr_dist = distances.max()
        kpi_metrics["pr_value"] = f"{pr_dist:.1f} KM"
        kpi_metrics["pr_label"] = "LONGEST RUN (PR)"
        kpi_metrics["pr_sub"] = "PEAK ENDURANCE"

        kpi_metrics["volume_value"] = f"{distances.sum():.1f} KM"
        kpi_metrics["volume_label"] = "CUMULATIVE DISTANCE"
        kpi_metrics["volume_sub"] = f"{total_discipline_sessions} SESSIONS LOGGED"

        load_scores = rpes * durations
        avg_load = load_scores.mean() if not load_scores.empty else 0
        kpi_metrics["workload_value"] = f"{avg_load:.0f} A.U."
        kpi_metrics["workload_label"] = "AVG AEROBIC WORKLOAD"
        kpi_metrics["workload_sub"] = "HIGH VOLUME" if avg_load >= 400 else "OPTIMAL BASE"

    elif activity == "Strength":
        weights = pd.to_numeric(all_filtered["Weight (kg)"], errors='coerce').fillna(0)
        sets = pd.to_numeric(all_filtered["Sets"], errors='coerce').fillna(0)
        reps = pd.to_numeric(all_filtered["Reps"], errors='coerce').fillna(0)
        rpes = pd.to_numeric(all_filtered["RPE"], errors='coerce').fillna(0)

        tonnage = weights * sets * reps
        est_1rms = weights * (1 + (reps / 30))

        pr_weight = weights.max()
        peak_1rm = est_1rms.max()
        kpi_metrics["pr_value"] = f"{pr_weight:.1f} KG"
        kpi_metrics["pr_label"] = "MAX LOAD RECORDED"
        kpi_metrics["pr_sub"] = f"PEAK 1RM: {peak_1rm:.1f} KG"

        kpi_metrics["volume_value"] = f"{tonnage.sum():,.0f} KG"
        kpi_metrics["volume_label"] = "TOTAL TONNAGE MOVED"
        kpi_metrics["volume_sub"] = f"{total_discipline_sessions} SESSIONS LOGGED"

        avg_rpe = rpes.mean() if not rpes.empty else 0
        kpi_metrics["workload_value"] = f"{avg_rpe:.1f} / 10"
        kpi_metrics["workload_label"] = "AVG INTENSITY (RPE)"
        kpi_metrics["workload_sub"] = "HEAVY HYPERTROPHY" if avg_rpe >= 8.0 else "SUBMAXIMAL WORK"

    # --- 2. SLICE BY SESSION WINDOW ---
    if dash_range == '10': target_df = all_filtered.tail(10)
    elif dash_range == 'all': target_df = all_filtered
    else: target_df = all_filtered.tail(5)

    # --- 3. BUILD SESSION HISTORY BREAKDOWN TABLE ---
    history_sessions = []
    for _, row in target_df.iloc[::-1].iterrows():
        ts = str(row.get("Timestamp", ""))[:16]
        hr_raw_val = row.get("Average Heart Rate")
        hr_val = f"{float(hr_raw_val):.0f} BPM" if pd.notna(hr_raw_val) and str(hr_raw_val).strip() != '' else "N/A"
        rpe_raw_val = row.get("RPE")
        rpe_val = f"{float(rpe_raw_val):.1f}" if pd.notna(rpe_raw_val) and str(rpe_raw_val).strip() != '' else "N/A"

        if activity == "Sprint":
            spd, dur = row.get("Top Speed (km/h)", 0), row.get("Duration (min)", 0)
            output_main, output_detail = f"{spd} km/h", f"{dur} min sprint"
            try: metric_extra = f"{float(rpe_raw_val) * float(dur):.0f} a.u. load"
            except: metric_extra = "-"
        elif activity == "Endurance":
            dst, dur = row.get("Distance (km)", 0), row.get("Duration (min)", 0)
            output_main, output_detail = f"{dst} km", f"{dur} min duration"
            try: metric_extra = f"{float(rpe_raw_val) * float(dur):.0f} a.u. load"
            except: metric_extra = "-"
        else:
            w_val, s_val, r_val = row.get("Weight (kg)", 0), row.get("Sets", 0), row.get("Reps", 0)
            output_main = f"{w_val} kg"
            output_detail = f"{row.get('Exercise Name', 'Lift')} | {s_val}s × {r_val}r"
            try: metric_extra = f"{float(w_val) * (1 + (float(r_val) / 30)):.1f} kg 1RM"
            except: metric_extra = "-"

        history_sessions.append({"timestamp": ts, "activity": activity.upper(), "output_main": output_main, "output_detail": output_detail, "hr": hr_val, "rpe": rpe_val, "metric_extra": metric_extra})

    # --- 4. MATPLOTLIB DARK AESTHETIC RENDERING ---
    plt.style.use('dark_background')
    fig_w = max(10, min(14, len(target_df) * 1.1)) if len(target_df) > 5 else 10
    fig = Figure(figsize=(fig_w, 4.5), dpi=100, facecolor='#050505')
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)

    x_labels = [f"S{i+1}" for i in range(len(target_df))]
    for ax in (ax1, ax2):
        ax.set_facecolor('#111111')
        ax.tick_params(colors='#888888', labelsize=9)
        for spine in ax.spines.values(): spine.set_color('#333333')
        ax.xaxis.label.set_color('#ffffff')
        ax.yaxis.label.set_color('#ffffff')
        ax.title.set_color('#ccff00')

    if activity == "Sprint":
        ax1.plot(x_labels, pd.to_numeric(target_df["Top Speed (km/h)"], errors='coerce').fillna(0), marker='o', color='#ccff00', linewidth=2.5)
        ax1.set_title("TOP SPEED TREND")
        ax1.set_ylabel("Speed (km/h)")
        ax1.grid(True, color="#222222", linestyle="--")

        ax2.bar(x_labels, pd.to_numeric(target_df["Average Heart Rate"], errors='coerce').fillna(0), color='#ffffff', alpha=0.8, width=0.45)
        ax2.set_title("CARDIO STRAIN")
        ax2.set_ylabel("Heart Rate (BPM)")
        ax2.axhline(y=160, color='#ccff00', linestyle="--")

    elif activity == "Endurance":
        ax1.plot(x_labels, pd.to_numeric(target_df["Distance (km)"], errors='coerce').fillna(0), marker='s', color='#ccff00', linewidth=2.5)
        ax1.set_title("VOLUME PROGRESSION")
        ax1.set_ylabel("Distance (km)")
        ax1.grid(True, color="#222222", linestyle="--")

        ax2.bar(x_labels, pd.to_numeric(target_df["Average Heart Rate"], errors='coerce').fillna(0), color='#ffffff', alpha=0.8, width=0.45)
        ax2.set_title("AEROBIC HEART RATE")
        ax2.set_ylabel("BPM")

    elif activity == "Strength":
        weights = pd.to_numeric(target_df["Weight (kg)"], errors='coerce').fillna(0)
        reps = pd.to_numeric(target_df["Reps"], errors='coerce').fillna(0)
        ax1.bar(x_labels, weights * pd.to_numeric(target_df["Sets"], errors='coerce').fillna(0) * reps, color='#ffffff', alpha=0.8, width=0.45)
        ax1.set_title("TOTAL VOLUME TONNAGE")
        ax1.set_ylabel("Volume (kg)")

        ax2.plot(x_labels, weights * (1 + (reps / 30)), marker='^', color='#ccff00', linewidth=2.5)
        ax2.set_title("ESTIMATED 1RM TREND")
        ax2.set_ylabel("1RM (kg)")
        ax2.grid(True, color="#222222", linestyle="--")

    fig.tight_layout(pad=3.0)
    img = io.BytesIO()
    fig.savefig(img, format='png', bbox_inches='tight')
    img.seek(0)
    chart_url = base64.b64encode(img.getvalue()).decode('utf8')

    return render_template(
        'index.html', summary=None, chart_url=chart_url, errors={}, dash_error=None,
        form_data={}, dash_name=athlete_name, dash_activity=activity, dash_range=dash_range,
        telemetry=telemetry, kpis=kpi_metrics, history=history_sessions
    )

@app.route('/export/<file_type>', methods=['POST'])
@login_required
def export_report(file_type):
    athlete_name = request.form.get('dash_name', '').strip()
    activity = request.form.get('dash_activity', 'Sprint').strip()
    if not athlete_name: return "Athlete name required", 400

    df = get_dataframe()
    filtered = df[(df["Athlete's Name"].astype(str).str.lower() == athlete_name.lower()) & (df["Activity Type"] == activity)]
    if filtered.empty: return "No data found", 404

    total_sessions = len(filtered)
    avg_hr_num = pd.to_numeric(filtered['Average Heart Rate'], errors='coerce').mean()
    avg_hr_text = f"{avg_hr_num:.0f} BPM" if pd.notna(avg_hr_num) else "N/A"
    
    if activity == "Sprint":
        peak_val = pd.to_numeric(filtered['Top Speed (km/h)'], errors='coerce').max()
        peak_output = f"{peak_val:.1f} km/h" if pd.notna(peak_val) else "N/A"
    elif activity == "Endurance":
        peak_val = pd.to_numeric(filtered['Distance (km)'], errors='coerce').max()
        peak_output = f"{peak_val:.2f} km (Max Dist)" if pd.notna(peak_val) else "N/A"
    else:
        peak_val = pd.to_numeric(filtered['Weight (kg)'], errors='coerce').max()
        peak_output = f"{peak_val:.1f} kg (Max Weight)" if pd.notna(peak_val) else "N/A"

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    plt.style.use('dark_background')
    fig = Figure(figsize=(8, 3.5), dpi=100, facecolor='#050505')
    ax1 = fig.add_subplot(121)
    ax2 = fig.add_subplot(122)
    recent_5 = filtered.tail(5)
    x_labels = [f"S{i+1}" for i in range(len(recent_5))]

    for ax in (ax1, ax2):
        ax.set_facecolor('#111111')
        ax.tick_params(colors='#888888', labelsize=8)
        for spine in ax.spines.values(): spine.set_color('#333333')
        ax.title.set_color('#ccff00')

    hrs = pd.to_numeric(recent_5["Average Heart Rate"], errors='coerce').fillna(0)
    if activity == "Sprint":
        ax1.plot(x_labels, pd.to_numeric(recent_5["Top Speed (km/h)"], errors='coerce').fillna(0), marker='o', color='#ccff00')
        ax1.set_title("TOP SPEED TREND")
        ax2.bar(x_labels, hrs, color='#ffffff')
        ax2.set_title("CARDIO STRAIN")
    elif activity == "Endurance":
        ax1.plot(x_labels, pd.to_numeric(recent_5["Distance (km)"], errors='coerce').fillna(0), marker='s', color='#ccff00')
        ax1.set_title("VOLUME PROGRESSION")
        ax2.bar(x_labels, hrs, color='#ffffff')
        ax2.set_title("AEROBIC HEART RATE")
    else:
        w, r = pd.to_numeric(recent_5["Weight (kg)"], errors='coerce').fillna(0), pd.to_numeric(recent_5["Reps"], errors='coerce').fillna(0)
        ax1.bar(x_labels, w * pd.to_numeric(recent_5["Sets"], errors='coerce').fillna(0) * r, color='#ffffff')
        ax1.set_title("VOLUME TONNAGE")
        ax2.plot(x_labels, w * (1 + (r / 30)), marker='^', color='#ccff00')
        ax2.set_title("EST. 1RM TREND")

    fig.tight_layout()
    graph_io = io.BytesIO()
    fig.savefig(graph_io, format='png', bbox_inches='tight')
    graph_io.seek(0)
    buffer = io.BytesIO()

    if file_type == 'pdf':
        pdf = FPDF()
        pdf.add_page()
        pdf.set_fill_color(5, 5, 5)
        pdf.rect(0, 0, 210, 297, 'F')

        font_path = os.path.join(app.root_path, "static", "fonts", "Bebas.ttf")
        if os.path.exists(font_path):
            pdf.add_font("Bebas", "", font_path)
            title_font = "Bebas"
        else: title_font = "Helvetica"

        # Turn off auto-page-break to strictly enforce 1 single page
        pdf.set_auto_page_break(auto=False)

        # Place logo on the left
        logo_path = os.path.join(app.root_path, "static", "app logo design.jpg")
        if os.path.exists(logo_path): 
            pdf.image(logo_path, x=10, y=10, w=14)

        # Move text cursor right to sit next to the logo
        pdf.set_xy(28, 10)
        pdf.set_font(title_font, "", 24)
        pdf.set_text_color(204, 255, 0)
        pdf.cell(0, 10, "OBSIDIAN(R) PERFORMANCE LABS", ln=True)
        
        # Align subheaders identically with the title
        pdf.set_x(28)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(0, 6, f"ATHLETE DOSSIER: {athlete_name.upper()}", ln=True)
        pdf.set_x(28)
        pdf.set_text_color(136, 136, 136)
        pdf.cell(0, 6, f"DATE GENERATED: {timestamp}", ln=True)
        pdf.ln(5) 
        
        # Reset X back to the normal left margin for the rest of the document
        pdf.set_x(10)

        pdf.set_text_color(204, 255, 0)
        pdf.set_font(title_font, "", 16)
        pdf.cell(0, 8, "[1] AGGREGATE METRICS (LIFETIME)", ln=True)
        
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "", 10)
        pdf.cell(0, 6, f"- Primary Discipline: {activity.upper()}", ln=True)
        pdf.cell(0, 6, f"- Total Sessions Logged: {total_sessions}", ln=True)
        pdf.cell(0, 6, f"- Peak Output Recorded: {peak_output}", ln=True)
        pdf.cell(0, 6, f"- Average Cardio Strain: {avg_hr_text}", ln=True)
        pdf.ln(8)

        pdf.set_text_color(204, 255, 0)
        pdf.set_font(title_font, "", 16)
        pdf.cell(0, 8, "[2] RECENT TRAINING LOG (LAST 5 SESSIONS)", ln=True)
        
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(40, 8, "DATE", border=1)
        pdf.cell(25, 8, "TYPE", border=1)
        pdf.cell(30, 8, "STRAIN (HR)", border=1)
        pdf.cell(20, 8, "RPE", border=1)
        pdf.cell(60, 8, "OUTPUT", border=1)
        pdf.ln()

        pdf.set_font("Helvetica", "", 9)
        for _, row in recent_5.iloc[::-1].iterrows():
            pdf.cell(40, 8, str(row['Timestamp'])[:16], border=1)
            pdf.cell(25, 8, activity.upper(), border=1)
            pdf.cell(30, 8, f"{float(row['Average Heart Rate']):.0f} BPM" if pd.notna(row['Average Heart Rate']) and str(row['Average Heart Rate']).strip() != '' else "N/A", border=1)
            pdf.cell(20, 8, str(row['RPE']) if pd.notna(row['RPE']) else "N/A", border=1)
            if activity == "Sprint": output_str = f"{row['Top Speed (km/h)']} km/h, {row['Duration (min)']} min"
            elif activity == "Endurance": output_str = f"{row['Distance (km)']} km, {row['Duration (min)']} min"
            else: output_str = f"{row['Weight (kg)']} kg, {row['Sets']}x{row['Reps']}"
            pdf.cell(60, 8, output_str, border=1)
            pdf.ln()
        pdf.ln(8)

        pdf.set_text_color(204, 255, 0)
        pdf.set_font(title_font, "", 16)
        pdf.cell(0, 8, "[3] STATISTICAL DATA", ln=True)
        # Scaled down graph slightly to guarantee fit
        pdf.image(graph_io, w=145) 
        pdf.ln(2) 

        pdf.set_text_color(204, 255, 0)
        pdf.set_font(title_font, "", 16)
        pdf.cell(0, 8, "[4] COACH'S NOTES", ln=True)
        pdf.set_fill_color(20, 20, 20)
        # Reduced notes box height slightly
        pdf.cell(0, 16, "", border=1, ln=True, fill=True)

        # Pull footer up to Y=275 to absolutely prevent overflow
        pdf.set_y(275)
        pdf.set_draw_color(34, 34, 34)
        pdf.set_line_width(0.3)
        pdf.line(10, 274, 200, 274)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(204, 255, 0)
        pdf.cell(0, 4, "OBSIDIAN(R) PERFORMANCE LABS  |  BUILT FOR THE UNKNOWN", ln=True, align="C")
        pdf.set_font("Helvetica", "", 7)
        pdf.set_text_color(120, 120, 120)
        pdf.cell(0, 4, "(C) OBSIDIAN(R) PERFORMANCE LABS 2026. ALL RIGHTS RESERVED. CONFIDENTIAL BIOMETRIC REPORT.", ln=True, align="C")
        pdf_bytes = bytes(pdf.output())
        buffer.write(pdf_bytes)
        buffer.seek(0)
        return send_file(buffer, download_name=f"{athlete_name}_OBSIDIAN_REPORT.pdf", as_attachment=True, mimetype='application/pdf')

    elif file_type == 'doc':
        doc = Document()
        logo_path = os.path.join(app.root_path, "static", "app logo design.jpg")
        if os.path.exists(logo_path): doc.add_picture(logo_path, width=Inches(1.2))

        doc.add_heading('OBSIDIAN® PERFORMANCE LABS', 0)
        doc.add_paragraph(f"ATHLETE DOSSIER: {athlete_name.upper()}\nDATE GENERATED: {timestamp}")
        doc.add_heading('[1] AGGREGATE METRICS (LIFETIME)', level=1)
        doc.add_paragraph(f"- Primary Discipline: {activity.upper()}\n- Total Sessions Logged: {total_sessions}\n- Peak Output Recorded: {peak_output}\n- Average Cardio Strain: {avg_hr_text}")
        doc.add_heading('[2] RECENT TRAINING LOG (LAST 5 SESSIONS)', level=1)
        
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text, hdr_cells[1].text, hdr_cells[2].text, hdr_cells[3].text, hdr_cells[4].text = 'DATE', 'TYPE', 'STRAIN (HR)', 'RPE', 'OUTPUT'
        
        for _, row in recent_5.iloc[::-1].iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Timestamp'])[:16]
            row_cells[1].text = activity.upper()
            row_cells[2].text = f"{float(row['Average Heart Rate']):.0f} BPM" if pd.notna(row['Average Heart Rate']) and str(row['Average Heart Rate']).strip() != '' else "N/A"
            row_cells[3].text = str(row['RPE']) if pd.notna(row['RPE']) else "N/A"
            if activity == "Sprint": row_cells[4].text = f"{row['Top Speed (km/h)']} km/h, {row['Duration (min)']} min"
            elif activity == "Endurance": row_cells[4].text = f"{row['Distance (km)']} km, {row['Duration (min)']} min"
            else: row_cells[4].text = f"{row['Weight (kg)']} kg, {row['Sets']}x{row['Reps']}"

        doc.add_heading('[3] STATISTICAL DATA', level=1)
        doc.add_picture(graph_io, width=doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin)
        doc.add_heading('[4] COACH\'S NOTES', level=1)
        doc.add_paragraph("\n\n_________________________________________________________")
        doc.add_paragraph("© OBSIDIAN® PERFORMANCE LABS 2026. ALL RIGHTS RESERVED. CONFIDENTIAL BIOMETRIC REPORT.")

        doc.save(buffer)
        buffer.seek(0)
        return send_file(buffer, download_name=f"{athlete_name}_OBSIDIAN_REPORT.docx", as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')

if __name__ == '__main__':
    app.run(debug=True)