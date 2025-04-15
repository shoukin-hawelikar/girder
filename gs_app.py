import streamlit as st
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline
from io import BytesIO

st.set_page_config(page_title="Plate Girder Predictor", layout="centered")
st.title("🔧 Optimum Plate Girder Design Predictor")

# File Upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df['length'] = df['length'].ffill()

    features = ['length', 'Load (kN/m)']
    targets = [
        'Depth (mm)', 'Web Thick. (mm)', 'Stiffener Spacing (mm)',
        'd/a Ratio', 'Flange Width (mm)', 'Flange Thick. (mm)',
        'Moment Capacity (kNm)', 'Shear Capacity (kN)', 'Deflection (mm)',
        'Stiffener Count', 'Weight (kg)'
    ]

    models = {}
    for target in targets:
        model = make_pipeline(
            PolynomialFeatures(degree=2),
            LinearRegression()
        )
        model.fit(df[features], df[target])
        models[target] = model

    # Input Section
    st.subheader("📥 Enter Girder Parameters")
    length = st.number_input("Length (m)", min_value=1.0, step=0.5)
    load = st.number_input("Load (kN/m)", min_value=1.0, step=10.0)

    if st.button("Generate Report"):
        input_data = pd.DataFrame([[length, load]], columns=features)
        predictions = {'Length': length, 'Load (kN/m)': load}

        for target in targets:
            pred = models[target].predict(input_data)[0]
            predictions[target] = max(round(pred, 2), 0)

        report_df = pd.DataFrame([
            (param, predictions[param]) for param in [
                'Length', 'Load (kN/m)', 'Depth (mm)', 'Web Thick. (mm)',
                'Stiffener Spacing (mm)', 'd/a Ratio', 'Flange Width (mm)',
                'Flange Thick. (mm)', 'Moment Capacity (kNm)',
                'Shear Capacity (kN)', 'Deflection (mm)',
                'Stiffener Count', 'Weight (kg)'
            ]
        ], columns=["Parameter", "Value"])

        st.success("✅ Prediction Complete")
        st.dataframe(report_df)

        # Download as Excel
        output = BytesIO()
        report_df.to_excel(output, index=False)
        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name=f"Opt_Girder_{length:.1f}m_{load:.1f}kN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
