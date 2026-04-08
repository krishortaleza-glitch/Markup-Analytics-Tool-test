import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Wholesale Markup Analytics", layout="wide")
st.title("💰 Wholesale Markup Analytics Tool")

# ==============================
# LOAD FILES
# ==============================
@st.cache_data
def load_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

st.header("Upload Files")

inv_file = st.file_uploader("Invoices")
prod_file = st.file_uploader("Products File")
front_file = st.file_uploader("Frontline")
tax_file = st.file_uploader("Taxes")
store_file = st.file_uploader("Storelist")

# ==============================
# FORMULA DISPLAY
# ==============================
st.markdown("### 📊 Markup % Formula")
st.info("Markup % = (Invoice Cost - (Frontline + Tax)) / (Frontline + Tax)")

# ==============================
# RUN LOGIC
# ==============================
if inv_file and prod_file and front_file and tax_file and store_file:

    inv = load_file(inv_file)
    prod = load_file(prod_file)
    front = load_file(front_file)
    tax = load_file(tax_file)
    store = load_file(store_file)

    st.success("Files loaded")

    # ==============================
    # FIXED COLUMN MAPPING
    # ==============================

    # Invoice
    inv_store = "store"
    inv_product = "productId"
    inv_cost = "price"

    # Product
    prod_id = "ProductId"
    prod_family = "Family"
    prod_type = "Type"

    # Frontline
    front_family = "Family"
    front_cost = "CasePrice"
    front_start = "Start"
    front_end = "End"

    # Store
    store_store = "uniqueId"
    store_state = "stateAbbrev"

    # Tax (ONLY selector kept)
    tax_state = st.selectbox("Tax State", tax.columns)
    tax_value = st.selectbox("Tax Value", tax.columns)

    if st.button("🚀 Run Analysis"):

        progress = st.progress(0)

        # ==============================
        # CLEAN KEYS (IMPORTANT)
        # ==============================
        def clean_id(x):
            return str(x).strip().lstrip("0")

        inv["ProductID"] = inv[inv_product].apply(clean_id)
        prod["ProductID"] = prod[prod_id].apply(clean_id)

        prod["Family"] = prod[prod_family].astype(str).str.strip().str.upper()
        front["Family"] = front[front_family].astype(str).str.strip().str.upper()

        inv["Store"] = inv[inv_store].astype(str).str.strip()
        store["Store"] = store[store_store].astype(str).str.strip()

        store["State"] = store[store_state].astype(str).str.strip()
        tax["State"] = tax[tax_state].astype(str).str.strip()

        progress.progress(10)

        # ==============================
        # MERGE PRODUCT (KEY STEP)
        # ==============================
        merged = inv.merge(
            prod[["ProductID", "Family", "Type"]],
            on="ProductID",
            how="left"
        )

        progress.progress(25)

        # ==============================
        # ACTIVE FRONTLINE
        # ==============================
        today = pd.Timestamp.today().normalize()

        front[front_start] = pd.to_datetime(front[front_start], errors="coerce")
        front[front_end] = pd.to_datetime(front[front_end], errors="coerce")
        front[front_end] = front[front_end].fillna(pd.Timestamp.max)

        active_front = front[
            (front[front_start] <= today) &
            (front[front_end] >= today)
        ].copy()

        active_front = (
            active_front
            .sort_values(front_start, ascending=False)
            .groupby("Family", as_index=False)
            .first()
        )

        progress.progress(45)

        # ==============================
        # MERGE FRONTLINE
        # ==============================
        merged = merged.merge(
            active_front[["Family", front_cost]],
            on="Family",
            how="left"
        )

        progress.progress(60)

        # ==============================
        # STORE → STATE
        # ==============================
        merged = merged.merge(
            store[["Store", "State"]],
            on="Store",
            how="left"
        )

        progress.progress(75)

        # ==============================
        # TAX
        # ==============================
        merged = merged.merge(
            tax[["State", tax_value]],
            on="State",
            how="left"
        )

        progress.progress(85)

        # ==============================
        # CALCULATIONS
        # ==============================
        merged["Invoice Cost"] = pd.to_numeric(merged[inv_cost], errors="coerce")
        merged["Frontline"] = pd.to_numeric(merged[front_cost], errors="coerce")
        merged["Tax"] = pd.to_numeric(merged[tax_value], errors="coerce")

        merged["Frontline"] = merged["Frontline"].fillna(0)
        merged["Tax"] = merged["Tax"].fillna(0)

        merged["Total Cost"] = merged["Frontline"] + merged["Tax"]
        merged["Markup"] = merged["Invoice Cost"] - merged["Total Cost"]
        merged["Markup %"] = merged["Markup"] / merged["Total Cost"]

        merged["Markup %"] = merged["Markup %"].replace([float("inf"), -float("inf")], 0)

        # FORMAT
        merged["Total Cost"] = merged["Total Cost"].round(2)
        merged["Markup"] = merged["Markup"].round(2)
        merged["Markup %"] = merged["Markup %"].round(3)

        progress.progress(90)

        # ==============================
        # FREQUENCY
        # ==============================
        freq = (
            merged
            .groupby(["State", "Family", "Invoice Cost"])
            .size()
            .reset_index(name="Frequency")
        )

        freq["Top"] = (
            freq.groupby(["State", "Family"])["Frequency"]
            .transform("max") == freq["Frequency"]
        )

        merged = merged.merge(freq, on=["State", "Family", "Invoice Cost"], how="left")

        # ==============================
        # FINAL OUTPUT (Type = Column C)
        # ==============================
        final = merged[[
            "State",
            "Family",
            "Type",  # ✅ from Product via ProductID
            "Invoice Cost",
            "Frontline",
            "Tax",
            "Total Cost",
            "Markup",
            "Markup %",
            "Frequency",
            "Top"
        ]].drop_duplicates()

        # ==============================
        # EXPORT
        # ==============================
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final.to_excel(writer, sheet_name="Analysis", index=False)
            merged.to_excel(writer, sheet_name="Full Output", index=False)

        output.seek(0)

        wb = load_workbook(output)
        ws = wb["Analysis"]

        green = PatternFill(start_color="C6EFCE", fill_type="solid")
        top_col = list(final.columns).index("Top") + 1

        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row=row_idx, column=top_col).value:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = green

        final_output = BytesIO()
        wb.save(final_output)
        final_output.seek(0)

        progress.progress(100)

        st.download_button(
            "📥 Download Analysis",
            data=final_output,
            file_name=f"markup_analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        )
