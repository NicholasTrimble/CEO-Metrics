from flask import Flask, render_template, jsonify 
import pandas as pd 
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data", "Query Test (1).xlsx")

# Executive clean names for your equipment classes 
CLASS_NAMES = {
    "1370": "Energy Recovery Wheel Systems",
    "1500": "Water Wash Air Systems",
    "4160": "Wheel Cassette and Hub Assemblies",
    "ZZPK": "Remote Panels and Hood Assemblies",
    "4080": "Subassembly Panels and Structural Units",
    "ZZDR": "Specialty Door Assemblies",
    "1400": "Ductwork and Frame Assemblies",
    "4100": "Latch Plates and Aluminum Components",
    "4490": "Wheel Spokes and Aluminum Hubs",
    "1306": "Replacement Parts and Repair Units",
    "SPPRT": "Replacement Parts and Repair Units"
}

# 100% Real Audited Unit Costs from JobAsmbl Query Analysis 
AUDITED_PRODUCT_COSTS = {
    "1370": {"materials": 3798.36, "labor": 1118.18, "overhead": 1600.05, "scrap": 132.94, "rework": 89.45, "warranty": 97.75},
    "1500": {"materials": 336.67, "labor": 84.93, "overhead": 102.38, "scrap": 11.78, "rework": 6.79, "warranty": 7.86},
    "4160": {"materials": 153.66, "labor": 109.65, "overhead": 147.29, "scrap": 5.38, "rework": 8.77, "warranty": 6.16},
    "ZZPK": {"materials": 40.08, "labor": 18.75, "overhead": 25.80, "scrap": 1.40, "rework": 1.50, "warranty": 1.27},
    "4080": {"materials": 311.52, "labor": 34.68, "overhead": 37.53, "scrap": 10.90, "rework": 2.77, "warranty": 5.76},
    "ZZDR": {"materials": 31.99, "labor": 11.05, "overhead": 15.04, "scrap": 1.12, "rework": 0.88, "warranty": 0.87},
    "1400": {"materials": 2.48, "labor": 0.65, "overhead": 0.85, "scrap": 0.09, "rework": 0.05, "warranty": 0.06},
    "4100": {"materials": 7.02, "labor": 2.92, "overhead": 4.08, "scrap": 0.25, "rework": 0.23, "warranty": 0.21},
    "4490": {"materials": 5.01, "labor": 1.88, "overhead": 2.43, "scrap": 0.18, "rework": 0.15, "warranty": 0.14} }

def process_raw_epicor_export():
    if not os.path.exists(DATA_PATH):
        # Fallback if sales query is not loaded yet
        return {
            "1370": {"name": "Energy Recovery Wheel Systems", "price": 8200.0, "volume": 238, "materials": 3798.36, "labor": 1118.18, "overhead": 1600.05, "scrap": 132.94, "rework": 89.45, "warranty": 97.75},
            "1500": {"name": "Water Wash Air Systems", "price": 2850.0, "volume": 620, "materials": 336.67, "labor": 84.93, "overhead": 102.38, "scrap": 11.78, "rework": 6.79, "warranty": 7.86},
            "4160": {"name": "Wheel Cassette and Hub Assemblies", "price": 650.0, "volume": 6878, "materials": 153.66, "labor": 109.65, "overhead": 147.29, "scrap": 5.38, "rework": 8.77, "warranty": 6.16}
        }

    df = pd.read_excel(DATA_PATH, sheet_name=0)
    df.columns = [str(col).strip() for col in df.columns]

    df["Quantity"] = pd.to_numeric(df.get("Quantity", 0), errors="coerce").fillna(0)
    df["Customer Unit Price"] = pd.to_numeric(df.get("Customer Unit Price", 0), errors="coerce").fillna(0)
    df["Customer Ext. Price"] = pd.to_numeric(df.get("Customer Ext. Price", 0), errors="coerce").fillna(0)

    mask_calc_ext = (df["Customer Ext. Price"] == 0) & (df["Quantity"] > 0) & (df["Customer Unit Price"] > 0)
    df.loc[mask_calc_ext, "Customer Ext. Price"] = df["Quantity"] * df["Customer Unit Price"]

    active_df = df[(df["Quantity"] > 0) & (df["Customer Ext. Price"] > 0)].copy()

    def determine_group(row):
        cid = str(row.get("ClassID", "")).strip()
        grp = str(row.get("Group", "")).strip()
        if cid.endswith(".0"):
            cid = cid[:-2]
        if cid and cid != "nan" and cid != "":
            return cid
        if grp and grp != "nan" and grp != "":
            return grp
        part = str(row.get("Part", "")).strip()
        return part[:4] if part else "OTHER"

    active_df["FamilyKey"] = active_df.apply(determine_group, axis=1)

    aggregated = active_df.groupby("FamilyKey").agg(
        total_qty=("Quantity", "sum"),
        total_revenue=("Customer Ext. Price", "sum")
    ).reset_index()

    metrics = {}
    for _, row in aggregated.iterrows():
        raw_key = str(row["FamilyKey"])
        key = raw_key.lower().replace(" ", "_")
        qty = int(row["total_qty"])
        rev = float(row["total_revenue"])

        if qty <= 0:
            continue

        avg_price = round(rev / qty, 2)
        if avg_price <= 0:
            continue

        # Use 100% Real Audited Job Costs if available
        if raw_key in AUDITED_PRODUCT_COSTS:
            c = AUDITED_PRODUCT_COSTS[raw_key]
            mat_cost = c["materials"]
            labor_cost = c["labor"]
            burden_cost = c["overhead"]
            scrap_cost = c["scrap"]
            rework_cost = c["rework"]
            warranty_cost = c["warranty"]
        else:
            # Fallback benchmark
            mat_cost = round(avg_price * 0.48, 2)
            labor_cost = round(avg_price * 0.18, 2)
            burden_cost = round(avg_price * 0.08, 2)
            scrap_cost = round(mat_cost * 0.035, 2)
            rework_cost = round(labor_cost * 0.08, 2)
            warranty_cost = round(avg_price * 0.012, 2)

        display_name = CLASS_NAMES.get(raw_key, f"Equipment Group {raw_key}")

        metrics[key] = {
            "name": display_name,
            "price": avg_price,
            "volume": qty,
            "materials": mat_cost,
            "labor": labor_cost,
            "overhead": burden_cost,
            "scrap": scrap_cost,
            "rework": rework_cost,
            "warranty": warranty_cost
        }

    return metrics

@app.after_request
def allow_mevisio_iframe(response):
    response.headers["X-Frame-Options"] = "ALLOWALL"
    response.headers["Content-Security-Policy"] = "frame-ancestors *"
    return response

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/data", methods=["GET"]) 
def get_data():
    return jsonify(process_raw_epicor_export())

if __name__ == "__main__":
    app.run(debug=True, port=5000)
