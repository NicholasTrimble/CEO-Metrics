from flask import Flask, render_template, jsonify 
import pandas as pd 
import os

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data", "Query Test (1).xlsx")

# Simple, non technical names for product groups 
CLASS_NAMES = {
    "1500": "Water Wash Air Systems",
    "WU": "Water Wash Air Systems",
    "WUUWC": "Water Wash Air Systems",
    "1300": "Spare Parts and Extra Equipment",
    "1306": "Replacement Parts and Repair Units",
    "SPPRT": "Replacement Parts and Repair Units",
    "1400": "Ductwork and Frame Parts",
    "4340": "Air Filters and Clean Air Systems",
    "4390": "Electric Motors and Power Belts",
    "4490": "Sensors and Automatic Controls",
    "4250": "Pulleys and Small Hardware",
    "4310": "Control Panels and Indicator Lights"
}

def process_raw_epicor_export():
    if not os.path.exists(DATA_PATH):
        return {
            "uwc": {"name": "Water Wash Air Systems", "price": 2850.0, "volume": 120, "materials": 1350.0, "labor": 520.0, "overhead": 210.0, "scrap": 85.0, "rework": 65.0, "warranty": 40.0},
            "service": {"name": "Replacement Parts and Repair Units", "price": 6023.0, "volume": 45, "materials": 2700.0, "labor": 1100.0, "overhead": 480.0, "scrap": 140.0, "rework": 95.0, "warranty": 60.0},
            "custom": {"name": "Ductwork and Frame Parts", "price": 257.0, "volume": 320, "materials": 115.0, "labor": 55.0, "overhead": 22.0, "scrap": 12.0, "rework": 8.0, "warranty": 4.0}
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
        
        # Remove trailing decimal points
        if cid.endswith(".0"):
            cid = cid[:-2]
            
        if cid and cid != "nan" and cid != "":
            return cid
        if grp and grp != "nan" and grp != "":
            return grp
        part = str(row.get("Part", "")).strip()
        return part[:2] if part else "OTHER"

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

        # Standard business cost estimates
        mat_cost = round(avg_price * 0.48, 2)       # 48% Raw materials
        labor_cost = round(avg_price * 0.18, 2)     # 18% Worker pay
        burden_cost = round(avg_price * 0.08, 2)    # 8% Building and machine costs
        scrap_cost = round(mat_cost * 0.035, 2)     # 3.5% Wasted metal and parts
        rework_cost = round(labor_cost * 0.08, 2)   # 8% Time spent fixing mistakes
        warranty_cost = round(avg_price * 0.012, 2) # 1.2% Fixing issues after customer gets it

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
