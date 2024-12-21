import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, PhotoImage

# Define the application class
class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analysis Application")
        self.root.geometry("800x600")  # Set a larger window size

        # Apply ttk theme
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Banner Frame
        banner_frame = tk.Frame(root, bg="#4CAF50")
        banner_frame.pack(fill="x")
        banner_label = tk.Label(
            banner_frame,
            text="Data Analysis Application",
            bg="#4CAF50",
            fg="white",
            font=("Helvetica", 16, "bold")
        )
        banner_label.pack(pady=10)

        # Input Fields
        self.inputs_frame = ttk.Labelframe(root, text="Input Parameters", padding=(10, 10))
        self.inputs_frame.pack(padx=20, pady=10, fill="x")

        # First Row: Brand Name and Creative Type
        ttk.Label(self.inputs_frame, text="Brand Name:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.brand_name_var = tk.StringVar(value="Ponds")
        ttk.Entry(self.inputs_frame, textvariable=self.brand_name_var).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(self.inputs_frame, text="Creative Type:").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        creative_type_options = ["F(TVC) + F(ICA)", "E(TVC) + F(ICA)", "M(TVC) + F(ICA)"]
        self.creative_type_var = tk.StringVar(value=creative_type_options[0])
        ttk.Combobox(self.inputs_frame, textvariable=self.creative_type_var, values=creative_type_options).grid(row=0, column=3, padx=5, pady=5)

        # Additional rows with improved spacing and padding
        self.tom_tvc_var = self.create_double_var(60)
        self.tom_tvc_ica_var = self.create_double_var(72)
        self.spont_brand_tvc_var = self.create_double_var(75)
        self.spont_brand_tvc_ica_var = self.create_double_var(80)
        self.mr_tvc_var = self.create_double_var(61)
        self.mr_tvc_ica_var = self.create_double_var(71)
        self.lower_percentile_var = self.create_double_var(40)
        self.upper_percentile_var = self.create_double_var(69)

        self.add_input_row("TOM TVC:", self.inputs_frame, 1, self.tom_tvc_var)
        self.add_input_row("TOM TVC+ICA:", self.inputs_frame, 1, self.tom_tvc_ica_var, col_offset=2)
        self.add_input_row("Spont Brand TVC:", self.inputs_frame, 2, self.spont_brand_tvc_var)
        self.add_input_row("Spont Brand TVC+ICA:", self.inputs_frame, 2, self.spont_brand_tvc_ica_var, col_offset=2)
        self.add_input_row("MR TVC:", self.inputs_frame, 3, self.mr_tvc_var)
        self.add_input_row("MR TVC+ICA:", self.inputs_frame, 3, self.mr_tvc_ica_var, col_offset=2)
        self.add_input_row("Lower Percentile:", self.inputs_frame, 4, self.lower_percentile_var)
        self.add_input_row("Upper Percentile:", self.inputs_frame, 4, self.upper_percentile_var, col_offset=2)

        # File Selection
        ttk.Label(self.inputs_frame, text="Excel File:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(self.inputs_frame, textvariable=self.file_path_var, width=35).grid(row=5, column=1, columnspan=2, padx=5, pady=5)
        ttk.Button(self.inputs_frame, text="Browse", command=self.browse_file).grid(row=5, column=3, padx=5, pady=5)

        # Target Audience Dropdown
        ttk.Label(self.inputs_frame, text="Target Audience:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        target_audience_options = ["None", "Male", "Female"]
        self.target_audience_var = tk.StringVar(value="None")
        ttk.Combobox(self.inputs_frame, textvariable=self.target_audience_var, values=target_audience_options).grid(row=6, column=1, padx=5, pady=5)

        # Run Button
        ttk.Button(root, text="Run Analysis", command=self.run_analysis).pack(pady=10)

        # Output Display
        self.output_frame = ttk.Labelframe(root, text="Analysis Results", padding=(10, 10))
        self.output_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create the Notebook for the tabs
        self.output_notebook = ttk.Notebook(self.output_frame, padding=(10, 10))
        self.output_notebook.pack(fill="both", expand=True, padx=20, pady=10)

        # Tab 1: Analysis Type 1
        self.tab1_frame = ttk.Frame(self.output_notebook)
        self.output_notebook.add(self.tab1_frame, text="Analysis Type 1")

        self.output_text1 = tk.Text(self.tab1_frame, wrap="word", height=15, width=80)
        self.output_text1.pack(side="left", fill="both", expand=True, padx=(0, 10))

        scrollbar1 = ttk.Scrollbar(self.tab1_frame, command=self.output_text1.yview)
        scrollbar1.pack(side="right", fill="y")
        self.output_text1.config(yscrollcommand=scrollbar1.set)

        # Tab 2: Analysis Type 2
        self.tab2_frame = ttk.Frame(self.output_notebook)
        self.output_notebook.add(self.tab2_frame, text="Analysis Type 2")

        self.output_text2 = tk.Text(self.tab2_frame, wrap="word", height=15, width=80)
        self.output_text2.pack(side="left", fill="both", expand=True, padx=(0, 10))

        scrollbar2 = ttk.Scrollbar(self.tab2_frame, command=self.output_text2.yview)
        scrollbar2.pack(side="right", fill="y")
        self.output_text2.config(yscrollcommand=scrollbar2.set)

        # Tab 3: Analysis Type 3
        self.tab3_frame = ttk.Frame(self.output_notebook)
        self.output_notebook.add(self.tab3_frame, text="Analysis Type 3")

        self.output_text3 = tk.Text(self.tab3_frame, wrap="word", height=15, width=80)
        self.output_text3.pack(side="left", fill="both", expand=True, padx=(0, 10))

        scrollbar3 = ttk.Scrollbar(self.tab3_frame, command=self.output_text3.yview)
        scrollbar3.pack(side="right", fill="y")
        self.output_text3.config(yscrollcommand=scrollbar3.set)


    def create_double_var(self, default):
        var = tk.DoubleVar(value=default)
        return var

    def add_input_row(self, label, frame, row, variable, col_offset=0):
        ttk.Label(frame, text=label).grid(row=row, column=0 + col_offset, sticky="e", padx=5, pady=5)
        ttk.Entry(frame, textvariable=variable).grid(row=row, column=1 + col_offset, padx=5, pady=5)

    def browse_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if filepath:
            self.file_path_var.set(filepath)

    def run_analysis(self):
        try:
            # Load inputs
            brand_name = self.brand_name_var.get()
            tom_tvc = self.tom_tvc_var.get()
            tom_tvc_ica = self.tom_tvc_ica_var.get()
            spont_brand_tvc = self.spont_brand_tvc_var.get()
            spont_brand_tvc_ica = self.spont_brand_tvc_ica_var.get()
            mr_tvc = self.mr_tvc_var.get()  # MR TVC
            mr_tvc_ica = self.mr_tvc_ica_var.get()  # MR TVC + ICA
            creative_type = self.creative_type_var.get()
            target_audience = self.target_audience_var.get()
            file_path = self.file_path_var.get()
            lower_percentile = self.lower_percentile_var.get()
            upper_percentile = self.upper_percentile_var.get()

            if not file_path:
                raise ValueError("Please select a valid Excel file.")

            # Load the dataset
            campaign_data_india = pd.read_excel(file_path, sheet_name='INDIA', engine='openpyxl')

            # List of columns to keep; Remove unused columns for manageability
            columns_to_keep = [
                'Year', 'SECTOR', 'CATEGORY', 'ADVERTISER', 'BRAND', 'TARGET AUDIENCE',
                'MARKET', 'CAMPAIGN FORMAT', 'TOM - TVC', 'TOM - TVC+ICA', 
                'TOM Uplift (TVC vs TVC + ICA)', 'BR Unaided - TVC', 
                'BR Unaied - TVC+ICA', 'BR Unaided Uplift (TVC vs TVC + ICA)', 
                'Type of TVC (F/E/M)', 'Type of ICA (F/E/M)', 'MR - TVC', 'MR - TVC+ICA'
            ]

            # Filter the data to include only the specified columns
            campaign_data_india = campaign_data_india[columns_to_keep]

            # Perform the analysis

            # Create a new column for % Uplift of Spont Brand
            campaign_data_india['Spont Brand Uplift (%)'] = (
                (campaign_data_india['BR Unaied - TVC+ICA'] - campaign_data_india['BR Unaided - TVC']) /
                campaign_data_india['BR Unaided - TVC']
            ) * 100

            # Create a new column for MR Uplift %
            campaign_data_india['MR Uplift (%)'] = (
                (campaign_data_india['MR - TVC+ICA'] - campaign_data_india['MR - TVC']) /
                campaign_data_india['MR - TVC']
            ) * 100

            filtered_data = campaign_data_india.dropna(subset=['BR Unaided - TVC', 'BR Unaied - TVC+ICA'])

            # Exclude low outliers
            filtered_data = filtered_data[filtered_data['BR Unaided - TVC'] > 25]

            # Filter for records where the target audience is female/male
            filtered_data['TARGET AUDIENCE'] = filtered_data['TARGET AUDIENCE'].astype(str)
            
            if target_audience == "Female":
                filtered_data = filtered_data[filtered_data['TARGET AUDIENCE'].str[0] == 'F']
            elif target_audience == "Male":
                filtered_data = filtered_data[filtered_data['TARGET AUDIENCE'].str[0] == 'M']


            br_unaided_percentiles = filtered_data['BR Unaided - TVC'].quantile([lower_percentile / 100, upper_percentile / 100])

            def categorize_brand_size(br_unaided_score):
                if br_unaided_score <= br_unaided_percentiles[lower_percentile / 100]:
                    return 'Small'
                elif br_unaided_score <= br_unaided_percentiles[upper_percentile / 100]:
                    return 'Medium'
                else:
                    return 'Large'

            filtered_data['Brand Size'] = filtered_data['BR Unaided - TVC'].apply(categorize_brand_size)

            # Average Spont Brand Uplift by Brand Size
            average_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].mean()
            count_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].size()

            # Average MR Uplift by Brand Size
            average_mr_uplifts = filtered_data.groupby('Brand Size')['MR Uplift (%)'].mean()
            count_mr_uplifts = filtered_data.groupby('Brand Size')['MR Uplift (%)'].size()

            current_spont_brand_uplift = (spont_brand_tvc_ica - spont_brand_tvc) / spont_brand_tvc * 100
            current_mr_uplift = (mr_tvc_ica - mr_tvc) / mr_tvc * 100

            current_brand_size = categorize_brand_size(spont_brand_tvc)
            average_spont_uplift_for_size = average_spont_brand_uplifts[current_brand_size]
            average_mr_uplift_for_size = average_mr_uplifts[current_brand_size]

            # Type of TVC vs Type of ICA Calculations
            filtered_type_data = filtered_data.dropna(subset=['Type of TVC (F/E/M)', 'Type of ICA (F/E/M)'])
            filtered_combinations_data = filtered_type_data[(
                ((filtered_type_data['Type of TVC (F/E/M)'] == 'E') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
                ((filtered_type_data['Type of TVC (F/E/M)'] == 'F') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
                ((filtered_type_data['Type of TVC (F/E/M)'] == 'M') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F'))
            )]

            combinations = {
                "E(TVC) + F(ICA)": {'TVC': 'E', 'ICA': 'F'},
                "F(TVC) + F(ICA)": {'TVC': 'F', 'ICA': 'F'},
                "M(TVC) + F(ICA)": {'TVC': 'M', 'ICA': 'F'}
            }

            combination_metrics = {}
            combination_metrics_mr = {}
            for combo_name, combo_values in combinations.items():
                combo_data = filtered_combinations_data[(
                    (filtered_combinations_data['Type of TVC (F/E/M)'] == combo_values['TVC']) &
                    (filtered_combinations_data['Type of ICA (F/E/M)'] == combo_values['ICA'])
                )]
                avg_spont_brand_uplift = combo_data['Spont Brand Uplift (%)'].mean()
                avg_mr_uplift = combo_data['MR Uplift (%)'].mean()
                record_count = combo_data.shape[0]
                combination_metrics[combo_name] = {
                    "Average Spont Brand Uplift (%)": avg_spont_brand_uplift,
                    "Record Count": record_count
                }

                combination_metrics_mr[combo_name] = {
                    "Average MR Uplift (%)": avg_mr_uplift,
                    "Record Count": record_count
                }

            average_uplift_for_current_type = combination_metrics[creative_type]["Average Spont Brand Uplift (%)"]
            average_mr_uplift_for_current_type = combination_metrics_mr[creative_type]["Average MR Uplift (%)"]

            ### Print out Results
            result1 = f"--- Analysis Results ---\n"
            result1 += f"Current Brand: {brand_name}\n"
            result1 += f"Current Spont Brand Uplift: {current_spont_brand_uplift:.2f}%\n"
            result1 += f"Average for {current_brand_size} Brands: {average_spont_uplift_for_size:.2f}%\n"

            if current_spont_brand_uplift > average_spont_uplift_for_size:
                result1 += f"The Spontaneous Brand Uplift % of this brand is above par for {current_brand_size} brands\n"
            else:
                result1 += f"The Spontaneous Brand Uplift % of this brand is below par for {current_brand_size} brands\n"

            result1 += "\nAverage Spont Brand Uplift by Brand Size:\n"
            for size, avg_uplift in average_spont_brand_uplifts.items():
                result1 += f"  {size}: {avg_uplift:.2f}%\n"

            result1 += "\nNumber of Brands by Size:\n"
            for size, count in count_spont_brand_uplifts.items():
                result1 += f"  {size}: {count}\n"

            result1 += "\n--- Type of TVC vs Type of ICA Analysis ---\n"
            for combo, metrics in combination_metrics.items():
                result1 += f"\nCombination: {combo}\n"
                result1 += f"  Average Spont Brand Uplift (%): {metrics['Average Spont Brand Uplift (%)']:.2f}\n"
                result1 += f"  Number of Studies: {metrics['Record Count']}\n"

            result1 += f"\n--- Comparison for Creative Type: {creative_type} ---\n"
            result1 += f"Current Ad Spont Brand Uplift: {current_spont_brand_uplift:.2f}%\n"
            result1 += f"Average Spont Brand Uplift for {creative_type}: {average_uplift_for_current_type:.2f}%\n"

            if current_spont_brand_uplift > average_uplift_for_current_type:
                result1 += f"The current ad shows a **significant improvement** compared to the average for the same creative type.\n"
            else:
                result1 += f"The current ad does **not show a significant improvement** compared to the average for the same creative type.\n"

            ## Output for MR analysis
            result2 = f"--- Analysis Results ---\n"
            result2 += f"\nCurrent Message Recall Uplift: {current_mr_uplift:.2f}%\n"
            result2 += f"Average for {current_brand_size} Brands: {average_mr_uplift_for_size:.2f}%\n"

            if current_mr_uplift > average_mr_uplift_for_size:
                result2 += f"The current ad shows a **significant improvement**.\n"
            else:
                result2 += f"The current ad does **not show a significant improvement**.\n"

            result2 += "\nAverage Message Recall Uplift by Brand Size:\n"
            for size, avg_uplift in average_mr_uplifts.items():
                result2 += f"  {size}: {avg_uplift:.2f}%\n"

            result2 += "\nNumber of Brands by Size:\n"
            for size, count in count_mr_uplifts.items():
                result2 += f"  {size}: {count}\n"

            result2 += "\n--- Type of TVC vs Type of ICA Analysis For MR---\n"
            for combo, metrics in combination_metrics_mr.items():
                result2 += f"\nCombination: {combo}\n"
                result2 += f"  Average MR Uplift (%): {metrics['Average MR Uplift (%)']:.2f}\n"
                result2 += f"  Number of studies: {metrics['Record Count']}\n"

            result2 += f"\n--- Comparison for Creative Type: {creative_type} ---\n"
            result2 += f"Current Ad MR Uplift: {current_mr_uplift:.2f}%\n"
            result2 += f"Average Spont Brand Uplift for {creative_type}: {average_mr_uplift_for_current_type:.2f}%\n"

            if current_mr_uplift > average_mr_uplift_for_current_type:
                result2 += f"The current ad shows a **significant improvement** compared to the average for the same creative type.\n"
            else:
                result2 += f"The current ad does **not show a significant improvement** compared to the average for the same creative type.\n"

            
            self.output_text1.delete("1.0", tk.END)
            self.output_text1.insert(tk.END, result1)

            self.output_text2.delete("1.0", tk.END)
            self.output_text2.insert(tk.END, result2)

            self.output_text3.delete("1.0", tk.END)
            self.output_text3.insert(tk.END, "hello")

        except Exception as e:
            messagebox.showerror("Error", str(e))

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
