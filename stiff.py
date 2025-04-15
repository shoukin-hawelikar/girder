# Step 1: Install required libraries
!pip install openpyxl

# Step 2: Import necessary libraries
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline
from google.colab import files

# Step 3: Upload your Excel file
uploaded = files.upload()

# Step 4: Load and prepare data
filename = list(uploaded.keys())[0]
df = pd.read_excel(filename, engine='openpyxl')
df['length'] = df['length'].ffill()

# Define features and targets
features = ['length', 'Load (kN/m)']
targets = [
    'Depth (mm)', 'Web Thick. (mm)', 'Stiffener Spacing (mm)',
    'd/a Ratio', 'Flange Width (mm)', 'Flange Thick. (mm)',
    'Moment Capacity (kNm)', 'Shear Capacity (kN)', 'Deflection (mm)',
    'Stiffener Count', 'Weight (kg)'
]

# Create and train models
models = {}
for target in targets:
    model = make_pipeline(
        PolynomialFeatures(degree=2),
        LinearRegression()
    )
    model.fit(df[features], df[target])
    models[target] = model

# Function to generate and download results
def generate_girder_report(length, load):
    # Create predictions
    input_data = pd.DataFrame([[length, load]], columns=features)
    predictions = {'Length': length, 'Load (kN/m)': load}
    
    for target in targets:
        pred = models[target].predict(input_data)[0]
        predictions[target] = max(round(pred, 2), 0)

    # Create output in two-column format
    parameter_order = [
        'Length', 'Load (kN/m)', 'Depth (mm)', 'Web Thick. (mm)',
        'Stiffener Spacing (mm)', 'd/a Ratio', 'Flange Width (mm)',
        'Flange Thick. (mm)', 'Moment Capacity (kNm)', 'Shear Capacity (kN)',
        'Deflection (mm)', 'Stiffener Count', 'Weight (kg)'
    ]
    
    report_data = [(param, predictions[param]) for param in parameter_order]
    report_df = pd.DataFrame(report_data, columns=['Parameter', 'Value'])

    # Format numeric values
    report_df['Value'] = report_df['Value'].apply(lambda x: f"{x:.2f}" if isinstance(x, (int, float)) else x)

    # Create filename
    safe_length = f"{length:.1f}".replace('.', '_')
    safe_load = f"{load:.1f}".replace('.', '_')
    output_filename = f"Opt_Girder_{safe_length}m_{safe_load}kN.xlsx"
    
    # Save to Excel
    report_df.to_excel(output_filename, index=False)
    
    # Download file
    files.download(output_filename)
    print(f"\n✅ Report downloaded as: {output_filename}")

# Get user input
print("Enter girder parameters:")
user_length = float(input("Length (meters): "))
user_load = float(input("Load (kN/m): "))

# Generate and download report
generate_girder_report(user_length, user_load)
