import pandas as pd

### INPUT VARIABLES ###
Brand_name = "Ponds"
TOM_tvc = 60
TOM_tvc_ica = 72
Spont_Brand_tvc = 75  # Same as BR Unaided - TVC
Spont_Brand_tvc_ica = 80  # Same as BR Unaided - TVC+ICA
Aided_Brand_tvc = 94
Aided_Brand_tvc_ica = 92 
Creative_type = 'F(TVC) + F(ICA)'



### LOAD AND CLEAN DATA ###

# Replace 'Global Campaign Tracker.xlsx' with the path to your file
file_path = "Global Campaign Tracker.xlsx"

# Load the specific sheet named 'India' using the 'openpyxl' engine
try:
    campaign_data_india = pd.read_excel(file_path, sheet_name='INDIA', engine='openpyxl')
    # print("Data from 'India' Sheet Loaded Successfully!")
    # print("Columns in Dataset:", campaign_data_india.columns)
    # print("First Few Rows:")
    # print(campaign_data_india.head())
except Exception as e:
    print(f"Error loading data from 'India' sheet: {e}")


# List of columns to keep; Remove unused columns for managability
columns_to_keep = [
    'Year', 'SECTOR', 'CATEGORY', 'ADVERTISER', 'BRAND', 'TARGET AUDIENCE',
    'MARKET', 'CAMPAIGN FORMAT', 'TOM - TVC', 'TOM - TVC+ICA', 
    'TOM Uplift (TVC vs TVC + ICA)', 'BR Unaided - TVC', 
    'BR Unaied - TVC+ICA', 'BR Unaided Uplift (TVC vs TVC + ICA)', 
    'Type of TVC (F/E/M)', 'Type of ICA (F/E/M)'
]

# Filter the data to include only the specified columns
campaign_data_india = campaign_data_india[columns_to_keep]

# Calculate the Spont Brand Uplift percentage for each record and add it as a new column
campaign_data_india['Spont Brand Uplift (%)'] = (
    (campaign_data_india['BR Unaied - TVC+ICA'] - campaign_data_india['BR Unaided - TVC']) /
    campaign_data_india['BR Unaided - TVC']
) * 100


# remove rows that do not have TVC-ICA scores
filtered_data = campaign_data_india.dropna(subset=['BR Unaided - TVC', 'BR Unaied - TVC+ICA'])

# Filter to include only records where 'BR Unaided - TVC' is greater than 25
filtered_data = filtered_data[filtered_data['BR Unaided - TVC'] > 25]

### APPLY SEARCH FILTERS ###

# Ensure that values in target audience column are treated as strings
filtered_data['TARGET AUDIENCE'] = filtered_data['TARGET AUDIENCE'].astype(str)
# Filter rows where Target Audience is Female
filtered_data = filtered_data[filtered_data['TARGET AUDIENCE'].str[0] == 'F']


### DEFINE BRAND SIZE THRESHOLDS ###

# using BR Unaided - TVC to define the thresholds for small, medium, and large brands.
br_unaided_percentiles = filtered_data['BR Unaided - TVC'].quantile([0.40, 0.69])


# Categorize brand sizes
def categorize_brand_size(br_unaided_score):
    if br_unaided_score <= br_unaided_percentiles[0.40]:
        return 'Small'
    elif br_unaided_score <= br_unaided_percentiles[0.69]:
        return 'Medium'
    else:
        return 'Large'

filtered_data['Brand Size'] = filtered_data['BR Unaided - TVC'].apply(categorize_brand_size)


# Calculate the average Spont Brand uplift percentage for each brand size
average_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].mean()
count_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].size()

# Print the averages to verify
print("Average Spont Brand Uplift by Brand Size:")
print(average_spont_brand_uplifts)
print(count_spont_brand_uplifts)

# Determine brand size of current ad
current_brand_size = categorize_brand_size(Spont_Brand_tvc)

# Calculate the Spont Brand uplift for the current brand
current_spont_brand_uplift = (Spont_Brand_tvc_ica - Spont_Brand_tvc) / Spont_Brand_tvc * 100  # Percentage uplift

# Retrieve the average Spont Brand uplift for the current brand size
average_spont_uplift_for_size = average_spont_brand_uplifts[current_brand_size]

# Compare the current brand's Spont Brand uplift to the average
print(f"\n--- Spont Brand (BR Unaided) Comparison ---")
print(f"Current Brand Spont Brand Uplift: {current_spont_brand_uplift:.2f}%")
print(f"Average Spont Brand Uplift for {current_brand_size} Brands: {average_spont_uplift_for_size:.2f}%")

if current_spont_brand_uplift > average_spont_uplift_for_size:
    print(f"The current ad for {Brand_name} shows a **significant improvement** in Spont Brand uplift compared to the average for {current_brand_size} brands.")
else:
    print(f"The current ad for {Brand_name} does **not show a significant improvement** in Spont Brand uplift compared to the average for {current_brand_size} brands.")



### TYPE OF TVC vs TYPE OF ICA CALCULATIONS ###

# Filter out rows where 'Type of TVC' or 'Type of ICA' are null
filtered_type_data = filtered_data.dropna(subset=['Type of TVC (F/E/M)', 'Type of ICA (F/E/M)'])

# Further filter where 'Type of ICA' is 'F'
filtered_ica_f_data = filtered_type_data[filtered_type_data['Type of ICA (F/E/M)'] == 'F']


# Step 1: Filter data to include only the specified combinations and target audience starting with 'F'
filtered_combinations_data = filtered_type_data[
    (((filtered_type_data['Type of TVC (F/E/M)'] == 'E') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
     ((filtered_type_data['Type of TVC (F/E/M)'] == 'F') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
     ((filtered_type_data['Type of TVC (F/E/M)'] == 'M') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F'))) &
    (filtered_type_data['TARGET AUDIENCE'].str[0] == 'F')  # Additional condition for target audience
]

# Step 2: Define the combinations and calculate metrics
combinations = {
    "E(TVC) + F(ICA)": {'TVC': 'E', 'ICA': 'F'},
    "F(TVC) + F(ICA)": {'TVC': 'F', 'ICA': 'F'},
    "M(TVC) + F(ICA)": {'TVC': 'M', 'ICA': 'F'}
}

# Initialize a dictionary to store results
combination_metrics = {}

# Loop through each combination
for combo_name, combo_values in combinations.items():
    # Filter the data for the current combination
    combo_data = filtered_combinations_data[
        (filtered_combinations_data['Type of TVC (F/E/M)'] == combo_values['TVC']) &
        (filtered_combinations_data['Type of ICA (F/E/M)'] == combo_values['ICA'])
    ]
    
    # Calculate metrics using the new percentage column
    avg_spont_brand_uplift = combo_data['Spont Brand Uplift (%)'].mean()  # Average percentage uplift
    record_count = combo_data.shape[0]  # Number of records
    
    # Store the results
    combination_metrics[combo_name] = {
        "Average Spont Brand Uplift (%)": avg_spont_brand_uplift,
        "Record Count": record_count
    }

# Step 3: Print the results
print("Metrics for Each Combination (Target Audience: Female):")
for combo, metrics in combination_metrics.items():
    print(f"\nCombination: {combo}")
    print(f"Average Spont Brand Uplift (%): {metrics['Average Spont Brand Uplift (%)']:.2f}")
    print(f"Record Count: {metrics['Record Count']}")


# Retrieve the average uplift for the current creative type
average_uplift_for_current_type = combination_metrics[Creative_type]["Average Spont Brand Uplift (%)"]

# Compare the current ad's uplift to the average
print(f"\n--- Comparison for Creative Type: {Creative_type} ---")
print(f"Current Ad Spont Brand Uplift: {current_spont_brand_uplift:.2f}%")
print(f"Average Spont Brand Uplift for {Creative_type}: {average_uplift_for_current_type:.2f}%")

if current_spont_brand_uplift > average_uplift_for_current_type:
    print(f"The current ad shows a **significant improvement** compared to the average for the same creative type.")
else:
    print(f"The current ad does **not show a significant improvement** compared to the average for the same creative type.")

