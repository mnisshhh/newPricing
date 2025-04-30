# --- Full Streamlit App with Integrated Logic ---

import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Price Calculator", layout="wide")
st.title("NEW ITEM PRICE - GDIN")

# Sidebar settings
st.sidebar.header("🔧 Settings")
custom_rate = st.sidebar.number_input("Custom Charge (%)", value=6.0, step=0.1)
international_rate = st.sidebar.number_input("International Courier Rate (₹/kg)", value=170.0, step=1.0)

# Provide download button for the template
with open("GDIN_TEMPELATE.xlsx", "rb") as template_file:
    st.download_button(
        label="📤 Download Excel Template",
        data=template_file,
        file_name="GDIN_TEMPELATE.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Domestic slabs
st.sidebar.subheader("Domestic Courier Slabs")
domestic_slabs = [
    (500, st.sidebar.number_input("≤ 500g", value=0, step=1)),
    (1000, st.sidebar.number_input("≤ 1000g", value=36, step=1)),
    (3000, st.sidebar.number_input("≤ 3000g", value=53, step=1)),
    (5000, st.sidebar.number_input("≤ 5000g", value=78, step=1)),
    (999999, st.sidebar.number_input("> 5000g", value=103, step=1)),
]

# Profit slabs
st.sidebar.header("📊 Profit Slabs")
profit_slabs = []
default_ranges = [
    (199, 100), (399, 130), (799, 160), (1599, 180),
    (3199, 240), (6399, 310), (12799, 500), (25599, 1300),
    (51199, 1500), (102399, 1800), (204799, 2200), (409599, 2400),
    (99999999, 5000)
]
for i, (limit, default_profit) in enumerate(default_ranges):
    value = st.sidebar.number_input(f"Profit if price ≤ {limit}", value=default_profit, key=f"profit_{i}")
    profit_slabs.append((limit, value))

# Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file and st.button("🚀 Process File"):
    with st.spinner("Processing..."):
        df = pd.read_excel(uploaded_file)

        # Drop unwanted columns
        df.drop(df.columns[[9, 10, 11, 12]], axis=1, inplace=True)

        # Calculate Discounted Price
        df['DISCOUNTED PRICE'] = df.iloc[:, 6] - (df.iloc[:, 6] * df.iloc[:, 8] / 100)

        # Profit calculation
        def get_profit(discounted_price):
            for limit, profit in profit_slabs:
                if discounted_price <= limit:
                    return max(profit, discounted_price * 0.10)
            return max(profit_slabs[-1][1], discounted_price * 0.10)

        df['PROFIT'] = df['DISCOUNTED PRICE'].apply(get_profit)

        # Courier fees
        df['INTERNATIONAL COURIER'] = (df.iloc[:, 7] / 1000) * international_rate

        def get_domestic(weight):
            for limit, fee in domestic_slabs:
                if weight <= limit:
                    return fee
            return domestic_slabs[-1][1]

        df['DOMESTIC COURIER'] = df.iloc[:, 7].apply(get_domestic)

        # SUM ALL
        df['SUM ALL'] = df['DISCOUNTED PRICE'] + df['PROFIT'] + df['INTERNATIONAL COURIER'] + df['DOMESTIC COURIER']

        # Initial Fee Calculators
        def initial_ref(sum_all):
            if sum_all <= 250: return sum_all * 0.03
            elif sum_all <= 500: return sum_all * 0.045
            elif sum_all <= 1000: return sum_all * 0.09
            else: return sum_all * 0.135

        def initial_closing(sum_all):
            if sum_all <= 250: return 7
            elif sum_all <= 500: return 20
            elif sum_all <= 1000: return 41
            else: return 80

        # Final Fee Calculators based on INTERNAL SELLING PRICE
        def final_ref(sum_all, internal_sp):
            if internal_sp <= 250: return sum_all * 0.03
            elif internal_sp <= 500: return sum_all * 0.045
            elif internal_sp <= 1000: return sum_all * 0.09
            else: return sum_all * 0.135

        def final_closing(internal_sp):
            if internal_sp <= 250: return 7
            elif internal_sp <= 500: return 20
            elif internal_sp <= 1000: return 41
            else: return 80

        # Calculation Function
        def compute(row):
            sum_all = row['SUM ALL']
            disc = row['DISCOUNTED PRICE']

            # Internal
            temp_rf = initial_ref(sum_all)
            temp_cf = initial_closing(sum_all)
            gst_rf = temp_rf * 0.18
            gst_cf = temp_cf * 0.18
            internal_sp = sum_all + temp_rf + temp_cf + gst_rf + gst_cf

            # Final based on internal_sp
            final_rf = final_ref(sum_all, internal_sp)
            final_cf = final_closing(internal_sp)
            gst_final_rf = final_rf * 0.18
            gst_final_cf = final_cf * 0.18
            custom_charge = disc * (custom_rate / 100)

            final_sp = sum_all + final_rf + final_cf + gst_final_rf + gst_final_cf + custom_charge

            return pd.Series([
                final_rf, final_cf, gst_final_rf, gst_final_cf, custom_charge, final_sp
            ])

        df[['REFERAL FEE', 'CLOSING FEE', 'GST ON REFERAL FEE', 'GST ON CLOSING FEE', 'CUSTOM CHARGE', 'SELLING PRICE']] = (
            df.apply(compute, axis=1)
        )

        st.success("✅ Done!")
        st.dataframe(df, use_container_width=True)

        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Processed')
        st.download_button("📥 Download Excel", data=buffer.getvalue(), file_name="processed_output.xlsx")
