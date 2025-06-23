import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Order Balance", layout="wide")
st.title("Order Balance - Export Data от SAP файл")

uploaded_file = st.file_uploader("Качи SAP Excel файл (export.XLSX)", type=["xlsx"])

def process_data(df):
    # 1. Преобразуване на типовете (по M-кода)
    df = df.copy()
    # Преобразуване на числови колони
    num_cols = [
        "Material", "Material Doc.Item", "Item", "Plant", "Storage Location", "Movement Type", "Goods recipient", "Sales Order Item", "Order", "Material Document", "Amount in LC", "Qty in order unit", "Last Characters", "Cost Center"
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    if "Posting Date" in df.columns:
        df["Posting Date"] = pd.to_datetime(df["Posting Date"], errors='coerce')
    if "Entry Date" in df.columns:
        df["Entry Date"] = pd.to_datetime(df["Entry Date"], errors='coerce')
    if "Document Date" in df.columns:
        df["Document Date"] = pd.to_datetime(df["Document Date"], errors='coerce')

    # 2. Филтриране на редове с празен Material
    df = df[df["Material"].notnull()]

    # 3. Избор на нужните колони
    cols = ["Material", "Material Description", "Movement Type", "Batch", "Quantity", "Goods recipient", "Order"]
    df = df[cols]
    df["Order"] = df["Order"].astype(str)

    # 4. Добавяне на multiplier
    def get_multiplier(mt):
        if mt == 101:
            return 1
        elif mt == 102:
            return -1
        elif mt == 261:
            return -1
        elif mt == 262:
            return 1
        elif mt == 531:
            return 1
        elif mt == 532:
            return -1
        else:
            return np.nan
    df["multiplier"] = df["Movement Type"].apply(get_multiplier)

    # 5. Pieces = Goods recipient * multiplier
    df["Pieces"] = df["Goods recipient"] * df["multiplier"]
    df = df.drop(["Goods recipient", "multiplier"], axis=1)

    # --- Materials By Order ---
    df_mbo = df.copy()
    df_mbo = df_mbo[~df_mbo["Movement Type"].isin([531, 532])]
    df_mbo["Quantity"] = df_mbo["Quantity"].abs()
    df_mbo["Pieces"] = df_mbo["Pieces"].abs()
    df_mbo["Division"] = df_mbo["Quantity"] / df_mbo["Pieces"]
    df_mbo["Movement Type"] = df_mbo["Movement Type"].astype(str).str[0]
    df_mbo = df_mbo.groupby(["Material", "Material Description", "Movement Type", "Batch", "Order"], as_index=False)["Division"].mean()
    df_mbo = df_mbo[df_mbo["Division"] > 0]
    df_mbo["PC kg"] = df_mbo["Division"].round(3)
    df_mbo = df_mbo.drop("Division", axis=1)
    # length
    df_mbo["length"] = df_mbo["Batch"].astype(str).str[-3:].astype(float) * 10 + 5
    df_mbo["Movement Type"] = df_mbo["Movement Type"].astype(int)
    df_mbo["Key"] = df_mbo["Movement Type"].apply(lambda x: "production" if x == 1 else ("consumption" if x == 2 else None))

    # --- Prod ---
    df_prod = df[~df["Movement Type"].isin([261, 262])].groupby(["Order"], as_index=False)["Pieces"].sum()
    df_prod = df_prod.rename(columns={"Pieces": "PC"})

    # --- Cons ---
    df_cons = df[df["Movement Type"].isin([261, 262])].groupby(["Order"], as_index=False)["Pieces"].sum()
    df_cons = df_cons.rename(columns={"Pieces": "PC"})

    # --- Orders Cons ---
    orders_cons = df_mbo[df_mbo["Key"] == "consumption"]
    # --- Orders Prod ---
    orders_prod = df_mbo[df_mbo["Key"] == "production"]
    # Добавяме Coef
    orders_prod = orders_prod.merge(orders_cons[["Order", "length"]], on="Order", how="left", suffixes=("", ".1"))
    orders_prod["Coef"] = (orders_prod["length.1"] // orders_prod["length"]).astype('Int64')

    # --- Real Cons ---
    real_cons = df[~df["Movement Type"].isin([261, 262])].groupby(["Order"], as_index=False)["Pieces"].sum()
    real_cons = real_cons.rename(columns={"Pieces": "PC"})
    real_cons = real_cons.merge(orders_prod[["Order", "Coef"]], on="Order", how="left")
    real_cons["PC consumption (REAL)"] = (real_cons["PC"] // real_cons["Coef"]).astype('Int64')
    real_cons = real_cons.drop(["PC", "Coef"], axis=1)

    # --- Export Data (ново) ---
    export = real_cons.merge(df_cons, on="Order", how="left")
    export = export.merge(df_prod, on="Order", how="left", suffixes=(" (SAP)", " production"))
    export = export.rename(columns={"PC (SAP)": "PC consumption (SAP)", "PC production": "PC production"})
    export = export.merge(orders_cons[["Order", "Material", "Material Description", "Batch", "PC kg"]], on="Order", how="left")
    export = export.merge(orders_prod[["Order", "Material", "Material Description", "Batch", "PC kg"]], on="Order", how="left", suffixes=("", ".1"))
    export["PC consumption (Difference)"] = export["PC consumption (REAL)"] + export["PC consumption (SAP)"]
    export["KG consumption (Difference)"] = (export["PC consumption (Difference)"] * export["PC kg"]).round(3)
    # Подреждане на колоните
    export = export[[
        "Order", "Material", "Material Description", "Batch", "PC kg", "Material.1", "Material Description.1", "Batch.1", "PC kg.1", "PC production", "PC consumption (REAL)", "PC consumption (SAP)", "PC consumption (Difference)", "KG consumption (Difference)"
    ]]
    # Филтриране на NaN/null редове за KG consumption (Difference)
    valid_mask = (
        export["KG consumption (Difference)"].notnull() &
        ~export["KG consumption (Difference)"].isna() &
        ~np.isnan(export["KG consumption (Difference)"])
    )
    anomaly_mask = valid_mask & (export["KG consumption (Difference)"].abs() > 1e9)
    export_clean = export[valid_mask & ~anomaly_mask].copy()

    # --- ERROR ORDERS ---
    error_mask = ~valid_mask | anomaly_mask
    error_orders = export[error_mask].copy()

    return export_clean, error_orders

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    st.write("## Преглед на данните от файла:")
    st.dataframe(df.head())
    try:
        export_data, error_orders = process_data(df)
        st.success("Обработката е успешна! Резултатът е Export Data:")
        if not isinstance(export_data, pd.DataFrame):
            export_data = pd.DataFrame(export_data)
        st.dataframe(export_data)
        if not export_data.empty:
            buffer = BytesIO()
            export_data.to_excel(buffer, index=False, engine='openpyxl')  # type: ignore
            buffer.seek(0)
            st.download_button(
                label="Изтегли Export Data като Excel",
                data=buffer,
                file_name="export_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.write("## ERROR ORDERS (Поръчки с грешки):")
        if not isinstance(error_orders, pd.DataFrame):
            error_orders = pd.DataFrame(error_orders)
        st.dataframe(error_orders)
        if not error_orders.empty:
            buffer2 = BytesIO()
            error_orders.to_excel(buffer2, index=False, engine='openpyxl')  # type: ignore
            buffer2.seek(0)
            st.download_button(
                label="Изтегли ERROR ORDERS като Excel",
                data=buffer2,
                file_name="error_orders.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        # --- Допълнителна таблица: само редове с KG consumption (Difference) != 0 и избрани колони ---
        if not isinstance(export_data, pd.DataFrame):
            export_data = pd.DataFrame(export_data)
        export_filtered = export_data[(export_data["KG consumption (Difference)"] != 0) & export_data["KG consumption (Difference)"].notnull() & ~export_data["KG consumption (Difference)"].isna()]
        export_filtered = export_filtered[[
            "Order", "Material", "Material Description", "Batch", "KG consumption (Difference)", "PC consumption (Difference)"
        ]].copy()
        if not isinstance(export_filtered, pd.DataFrame):
            export_filtered = pd.DataFrame(export_filtered)
        # Добавяме колоната 'Mvnt'
        export_filtered["Mvnt"] = export_filtered["PC consumption (Difference)"].apply(lambda x: 261 if x > 0 else (262 if x < 0 else None))
        st.write("## Export Data (само редове с KG consumption (Difference) ≠ 0):")
        st.dataframe(export_filtered)
        if not export_filtered.empty:
            buffer3 = BytesIO()
            export_filtered.to_excel(buffer3, index=False, engine='openpyxl')  # type: ignore
            buffer3.seek(0)
            st.download_button(
                label="Изтегли Export Data (≠ 0) като Excel",
                data=buffer3,
                file_name="export_data_nonzero.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Възникна грешка при обработката: {e}")
else:
    st.info("Моля, качете SAP Excel файл за обработка.") 
