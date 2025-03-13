import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.genmod.families.family import NegativeBinomial
import tkinter as tk
from tkinter import filedialog
from scipy import stats

def select_file(title):
    """Allow user to select a file"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    try:
        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
    finally:
        root.destroy()
    
    return file_path if file_path else None

# Allow user to select input file
print("Please select the input Excel file...")
file_path = select_file("Select Excel Data File")

if not file_path:
    print("No file selected. Exiting.")
    exit()

# Import the data
print(f"Loading data from: {file_path}")
df = pd.read_excel(file_path, sheet_name="Stata")

# Setup
pd.set_option('display.max_columns', None)

# Encode categorical variables if not already encoded
categorical_vars = ['sex', 'marital_status', 'employment_status', 'purpose', 'accomd_type', 'us_state']
encoded_vars = {}

for var in categorical_vars:
    if var in df.columns:
        # Check if variable is already numeric
        if not pd.api.types.is_numeric_dtype(df[var]):
            new_var = f"{var}_enc"
            df[new_var] = pd.Categorical(df[var]).codes
            encoded_vars[var] = new_var
        else:
            encoded_vars[var] = var

# Set the truncation point for los (assuming truncation at 0)
df['los_trunc'] = df['los'].copy()
df.loc[df['los_trunc'] <= 0, 'los_trunc'] = np.nan

# Check for missing data
print("\nMissing data summary:")
print(df.isnull().sum())

print("\nMissing data patterns:")
missing_patterns = df.isnull().sum(axis=1)
print(missing_patterns.value_counts().sort_index())

# Visualize los distribution
plt.figure(figsize=(10, 6))
sns.histplot(df['los_trunc'], discrete=True)
plt.title('Histogram of Length of Stay')
plt.savefig('los_histogram.png')
plt.close()

# Summarize los by purpose
if 'purpose_enc' in df.columns:
    print("\nLength of stay by purpose:")
    purpose_stats = df.groupby('purpose_enc')['los_trunc'].agg(['count', 'mean', 'median', 'min', 'max', 'std'])
    print(purpose_stats)

# Detailed summary of los_trunc
print("\nDetailed summary of length of stay:")
los_describe = df['los_trunc'].describe(percentiles=[.25, .5, .75, .90, .95, .99])
print(los_describe)

# Cleaning process
# Step 1: Drop missing datapoints for key variables
key_vars = ['los', 'immigrant_population', 'import_from_slu', 'age', 
            encoded_vars.get('sex', 'sex_enc'), 
            encoded_vars.get('marital_status', 'marital_status_enc'), 
            encoded_vars.get('employment_status', 'employment_status_enc'), 
            'distance_miles', 
            encoded_vars.get('purpose', 'purpose_enc'), 
            encoded_vars.get('accomd_type', 'accomd_type_enc'), 
            'month_travel', 'state_percapita_income', 'state_unemployment']

# Count missing values per row for key variables
df['missing'] = df[key_vars].isnull().sum(axis=1)
print("\nNumber of missing values per observation:")
print(df['missing'].value_counts().sort_index())

# Drop observations with missing values in key variables
df_clean = df[df['missing'] == 0].drop('missing', axis=1)
print(f"\nRemaining observations after dropping missing values: {len(df_clean)}")

# Step 2: Drop outliers in length of stay
los_p95 = np.percentile(df_clean['los_trunc'].dropna(), 95)
df_clean['los_capped'] = df_clean['los_trunc'].copy()
df_clean.loc[df_clean['los_capped'] > los_p95, 'los_capped'] = los_p95

# Visualize the capped los distribution
plt.figure(figsize=(10, 6))
sns.histplot(df_clean['los_capped'], discrete=True)
plt.title('Histogram of Capped Length of Stay')
plt.savefig('los_capped_histogram.png')
plt.close()

# Step 3: Clean up the purpose of trip column
# Create a new simplified purpose variable
purpose_mapping = {
    1: 1,  # BUSINESS/MEETING -> Business
    2: 1,  # CONVENTION -> Business
    3: 1,  # CREW -> Business
    5: 2,  # EVENT -> Events
    6: 2,  # EVENTS -> Events
    7: 4,  # HONEYMOON -> Pleasure
    8: 5,  # INTRANSIT PASSEN -> Other
    9: 5,  # OTHER -> Other
    10: 4, # PLEASURE/HOLIDAY -> Pleasure
    11: 5, # RESIDENT -> Other
    12: 2, # SAINT LUCIA CARN -> Events
    13: 2, # SAINT LUCIA JAZZ -> Events
    14: 5, # SPORTS -> Other
    15: 5, # STUDY -> Other
    16: 5, # VISITING FRIENDS -> Other
    17: 3, # WEDDING -> Wedding
    18: 4, # pLEASURE/HOLIDAY -> Pleasure
    4: 5,  # CRICKET -> Other
}

purpose_labels = {
    1: "Business",
    2: "Events",
    3: "Wedding",
    4: "Pleasure",
    5: "Other"
}

# Add the simplified purpose variable
purpose_enc_col = encoded_vars.get('purpose', 'purpose_enc')
df_clean['purpose_simple'] = df_clean[purpose_enc_col].map(purpose_mapping)

# Check the new variable
print("\nPurpose simple distribution:")
purpose_counts = df_clean['purpose_simple'].value_counts().sort_index()
for code, count in purpose_counts.items():
    print(f"{code} ({purpose_labels.get(code, 'Unknown')}): {count}")

# Fit simple negative binomial regression
print("\nFitting simple negative binomial regression model...")
formula = ('los_capped ~ immigrant_population + import_from_slu + age + '
           f'C({encoded_vars.get("sex", "sex_enc")}) + '
           f'C({encoded_vars.get("marital_status", "marital_status_enc")}) + '
           f'C({encoded_vars.get("employment_status", "employment_status_enc")}) + '
           f'distance_miles + C(purpose_simple) + '
           f'C({encoded_vars.get("accomd_type", "accomd_type_enc")}) + '
           f'C(month_travel) + state_percapita_income + state_unemployment')

nb_model = smf.glm(formula=formula, 
                  data=df_clean, 
                  family=sm.families.NegativeBinomial(link=sm.families.links.log()))

try:
    nb_results = nb_model.fit()
    print("\nNegative Binomial Regression Results:")
    print(nb_results.summary())
    
    # Convert coefficients to incident rate ratios (IRR)
    print("\nIncident Rate Ratios (IRR):")
    irr = np.exp(nb_results.params)
    irr_conf = np.exp(nb_results.conf_int())
    irr_df = pd.DataFrame({'IRR': irr, 'Lower CI': irr_conf[0], 'Upper CI': irr_conf[1], 
                          'P-value': nb_results.pvalues})
    print(irr_df)
    
    # Multilevel model with random effects for states
    # In Python, we can approximate this using a mixed effects model with statsmodels
    # For a proper multilevel negative binomial model, consider using PyMC3 or other Bayesian frameworks
    print("\nNote: For a true multilevel negative binomial model with random effects,")
    print("consider using PyMC3, R (glmer.nb), or other specialized packages.")
    print("The simple negative binomial model here provides a starting point for analysis.")
    
    # Predictions and diagnostics
    df_clean['predicted'] = nb_results.predict()
    df_clean['residuals'] = df_clean['los_capped'] - df_clean['predicted']
    
    # Plot residuals
    plt.figure(figsize=(10, 6))
    plt.scatter(df_clean['predicted'], df_clean['residuals'], alpha=0.5)
    plt.axhline(y=0, color='r', linestyle='-')
    plt.xlabel('Predicted Values')
    plt.ylabel('Residuals')
    plt.title('Residual Plot')
    plt.savefig('residuals_plot.png')
    plt.close()
    
    # Check for heterogeneity across states if us_state is in the data
    if 'us_state_enc' in df_clean.columns or 'us_state' in df_clean.columns:
        state_var = 'us_state_enc' if 'us_state_enc' in df_clean.columns else 'us_state'
        state_means = df_clean.groupby(state_var)['los_capped'].mean().sort_values()
        
        plt.figure(figsize=(12, 8))
        state_means.plot(kind='bar')
        plt.xlabel('State')
        plt.ylabel('Average Length of Stay')
        plt.title('Mean Length of Stay by State')
        plt.xticks(rotation=90)
        plt.tight_layout()
        plt.savefig('los_by_state.png')
        plt.close()
        
except Exception as e:
    print(f"\nError in model fitting: {str(e)}")
    print("You may need to check your data or consider using a different modeling approach.")

print("\nAnalysis complete. Check the output figures for visualizations.")
