from flask import Flask, render_template, jsonify 
import pandas as pd 
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

def resolve_data_path():
    preferred = os.path.join(DATA_DIR, "Query Test.xlsx")
    if os.path.exists(preferred):
        return preferred
    if os.path.exists(DATA_DIR):
        for fname in os.listdir(DATA_DIR):
            if fname.endswith(".xlsx") and not fname.startswith("~$"):
                return os.path.join(DATA_DIR, fname)
    return preferred

@app.route("/api/data", methods=["GET"]) 
def get_data():
    file_path = resolve_data_path()
    if not os.path.exists(file_path):
        return jsonify({"error": f"File not found at {file_path}"}), 404

    df = pd.read_excel(file_path, sheet_name=0)
    df.columns = [str(col).strip() for col in df.columns]

    # 1. DEDUPLICATE INVOICED SALES
    inv_cols = [
        'Invoice', 'Part', 'Quantity', 
        'Customer Unit Price', 'Customer Ext. Price', 
        'ClassID', 'Group', 'Description'
    ]
    present_inv_cols = [c for c in inv_cols if c in df.columns]
    df_inv = df[present_inv_cols].drop_duplicates()

    df_inv = df_inv[(df_inv['Quantity'] > 0) & (df_inv['Customer Ext. Price'] > 0)].copy()
    df_inv['Quantity'] = pd.to_numeric(df_inv['Quantity'], errors='coerce').fillna(0)
    df_inv['Customer Ext. Price'] = pd.to_numeric(df_inv['Customer Ext. Price'], errors='coerce').fillna(0)

    # 2. DEDUPLICATE JOBS & ELIMINATE ZERO-COST ANOMALIES
    job_part_col = 'Part.1' if 'Part.1' in df.columns else 'Part'
    job_cols = [
        'Job', job_part_col, 'Completed Qty',
        'This Level Actual Material Cost', 'Lower Level Actual Material Cost',
        'This Level Actual Labor Cost', 'Lower Level Actual Labor Cost',
        'This Level Actual Burden Cost', 'Lower Level Actual Burden Cost',
        'This Level Actual Material Burden Cost'
    ]
    present_job_cols = [c for c in job_cols if c in df.columns]
    df_jobs = df[present_job_cols].drop_duplicates(subset=['Job']).copy()

    for c in present_job_cols:
        if c not in ['Job', job_part_col]:
            df_jobs[c] = pd.to_numeric(df_jobs[c], errors='coerce').fillna(0)

    df_jobs['Job_Mat'] = (
        df_jobs.get('This Level Actual Material Cost', 0) +
        df_jobs.get('Lower Level Actual Material Cost', 0) +
        df_jobs.get('This Level Actual Material Burden Cost', 0)
    )
    df_jobs['Job_Lab'] = (
        df_jobs.get('This Level Actual Labor Cost', 0) +
        df_jobs.get('Lower Level Actual Labor Cost', 0)
    )
    df_jobs['Job_Bur'] = (
        df_jobs.get('This Level Actual Burden Cost', 0) +
        df_jobs.get('Lower Level Actual Burden Cost', 0)
    )
    df_jobs['Job_Total'] = df_jobs['Job_Mat'] + df_jobs['Job_Lab'] + df_jobs['Job_Bur']

    # Filter out empty/uncosted production runs (Completed Qty > 0 and Actual Spend > 0)
    df_valid_jobs = df_jobs[(df_jobs['Completed Qty'] > 0) & (df_jobs['Job_Total'] > 0)].copy()

    # 3. AGGREGATE PER PART
    part_jobs = df_valid_jobs.groupby(job_part_col).agg(
        total_job_qty=('Completed Qty', 'sum'),
        total_job_mat=('Job_Mat', 'sum'),
        total_job_lab=('Job_Lab', 'sum'),
        total_job_bur=('Job_Bur', 'sum')
    ).reset_index()

    part_inv = df_inv.groupby('Part').agg(
        invoiced_qty=('Quantity', 'sum'),
        total_revenue=('Customer Ext. Price', 'sum')
    ).reset_index()

    parts_merged = pd.merge(part_inv, part_jobs, left_on='Part', right_on=job_part_col, how='inner')

    meta = df[['Part', 'Description', 'ClassID', 'Group']].drop_duplicates(subset=['Part']).copy()
    parts_merged = pd.merge(parts_merged, meta, on='Part', how='left')

    def assign_group_key(row):
        cid = str(row.get('ClassID', '')).replace('.0', '').strip()
        grp = str(row.get('Group', '')).strip()
        if cid and cid != 'nan':
            return cid
        if grp and grp != 'nan':
            return grp
        return 'OTHER'

    parts_merged['GroupKey'] = parts_merged.apply(assign_group_key, axis=1)

    # 4. ROLL UP BY CLASS
    grouped = parts_merged.groupby('GroupKey').agg(
        display_name=('Description', 'first'),
        invoiced_qty=('invoiced_qty', 'sum'),
        total_revenue=('total_revenue', 'sum'),
        job_qty=('total_job_qty', 'sum'),
        total_mat=('total_job_mat', 'sum'),
        total_lab=('total_job_lab', 'sum'),
        total_bur=('total_job_bur', 'sum')
    ).reset_index()

    metrics = {}
    for _, row in grouped.iterrows():
        key = str(row['GroupKey']).lower().replace(" ", "_")
        inv_qty = float(row['invoiced_qty'])
        job_qty = float(row['job_qty'])

        if inv_qty <= 0 or job_qty <= 0:
            continue

        price = round(float(row['total_revenue']) / inv_qty, 2)
        unit_mat = round(float(row['total_mat']) / job_qty, 2)
        unit_lab = round(float(row['total_lab']) / job_qty, 2)
        unit_bur = round(float(row['total_bur']) / job_qty, 2)

        # Baseline engineering model components
        scrap_cost = round(unit_mat * 0.035, 2)
        rework_cost = round(unit_lab * 0.08, 2)
        warranty_cost = round((unit_mat + unit_lab + unit_bur) * 0.015, 2)

        metrics[key] = {
            "name": f"Class {row['GroupKey']} | {row['display_name']}",
            "price": price,
            "volume": int(inv_qty),
            "materials": unit_mat,
            "labor": unit_lab,
            "overhead": unit_bur,
            "scrap": scrap_cost,
            "rework": rework_cost,
            "warranty": warranty_cost
        }

    return jsonify(metrics)

@app.after_request
def allow_mevisio_iframe(response):
    response.headers["X-Frame-Options"] = "ALLOWALL"
    response.headers["Content-Security-Policy"] = "frame-ancestors *"
    return response

@app.route("/")
def index():
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True, port=5000)
