import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Define the application class
class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analysis Application")

        # Input Fields
        self.inputs_frame = tk.Frame(root)
        self.inputs_frame.pack(pady=10)

        tk.Label(self.inputs_frame, text="Brand Name:").grid(row=0, column=0, sticky="e")
        self.brand_name_var = tk.StringVar(value="Ponds")
        tk.Entry(self.inputs_frame, textvariable=self.brand_name_var).grid(row=0, column=1)

        tk.Label(self.inputs_frame, text="TOM TVC:").grid(row=1, column=0, sticky="e")
        self.tom_tvc_var = tk.DoubleVar(value=60)
        tk.Entry(self.inputs_frame, textvariable=self.tom_tvc_var).grid(row=1, column=1)

        tk.Label(self.inputs_frame, text="TOM TVC+ICA:").grid(row=2, column=0, sticky="e")
        self.tom_tvc_ica_var = tk.DoubleVar(value=72)
        tk.Entry(self.inputs_frame, textvariable=self.tom_tvc_ica_var).grid(row=2, column=1)

        tk.Label(self.inputs_frame, text="Spont Brand TVC:").grid(row=3, column=0, sticky="e")
        self.spont_brand_tvc_var = tk.DoubleVar(value=75)
        tk.Entry(self.inputs_frame, textvariable=self.spont_brand_tvc_var).grid(row=3, column=1)

        tk.Label(self.inputs_frame, text="Spont Brand TVC+ICA:").grid(row=4, column=0, sticky="e")
        self.spont_brand_tvc_ica_var = tk.DoubleVar(value=80)
        tk.Entry(self.inputs_frame, textvariable=self.spont_brand_tvc_ica_var).grid(row=4, column=1)

        tk.Label(self.inputs_frame, text="Creative Type:").grid(row=5, column=0, sticky="e")
        self.creative_type_var = tk.StringVar(value="F(TVC) + F(ICA)")
        tk.Entry(self.inputs_frame, textvariable=self.creative_type_var).grid(row=5, column=1)

        # File selection
        tk.Label(self.inputs_frame, text="Excel File:").grid(row=6, column=0, sticky="e")
        self.file_path_var = tk.StringVar()
        tk.Entry(self.inputs_frame, textvariable=self.file_path_var, width=40).grid(row=6, column=1)
        tk.Button(self.inputs_frame, text="Browse", command=self.browse_file).grid(row=6, column=2)

        # Run Button
        tk.Button(root, text="Run Analysis", command=self.run_analysis).pack(pady=10)

        # Output Display
        self.output_frame = tk.Frame(root)
        self.output_frame.pack(fill="both", expand=True, pady=10)

        self.output_text = tk.Text(self.output_frame, wrap="word", height=20, width=80)
        self.output_text.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(self.output_frame, command=self.output_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.output_text.config(yscrollcommand=scrollbar.set)

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
            creative_type = self.creative_type_var.get()
            file_path = self.file_path_var.get()

            if not file_path:
                raise ValueError("Please select a valid Excel file.")

            # Load the dataset
            campaign_data_india = pd.read_excel(file_path, sheet_name='INDIA', engine='openpyxl')

            # Perform the analysis (simplified for GUI context)
            campaign_data_india['Spont Brand Uplift (%)'] = (
                (campaign_data_india['BR Unaied - TVC+ICA'] - campaign_data_india['BR Unaided - TVC']) /
                campaign_data_india['BR Unaided - TVC']
            ) * 100

            filtered_data = campaign_data_india.dropna(subset=['BR Unaided - TVC', 'BR Unaied - TVC+ICA'])
            filtered_data = filtered_data[filtered_data['BR Unaided - TVC'] > 25]

            # Filter for records where the target audience is female
            filtered_data['TARGET AUDIENCE'] = filtered_data['TARGET AUDIENCE'].astype(str)
            filtered_data = filtered_data[filtered_data['TARGET AUDIENCE'].str[0] == 'F']

            br_unaided_percentiles = filtered_data['BR Unaided - TVC'].quantile([0.40, 0.69])

            def categorize_brand_size(br_unaided_score):
                if br_unaided_score <= br_unaided_percentiles[0.40]:
                    return 'Small'
                elif br_unaided_score <= br_unaided_percentiles[0.69]:
                    return 'Medium'
                else:
                    return 'Large'

            filtered_data['Brand Size'] = filtered_data['BR Unaided - TVC'].apply(categorize_brand_size)
            average_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].mean()
            count_spont_brand_uplifts = filtered_data.groupby('Brand Size')['Spont Brand Uplift (%)'].size()

            current_spont_brand_uplift = (spont_brand_tvc_ica - spont_brand_tvc) / spont_brand_tvc * 100
            current_brand_size = categorize_brand_size(spont_brand_tvc)
            average_spont_uplift_for_size = average_spont_brand_uplifts[current_brand_size]

            # Type of TVC vs Type of ICA Calculations
            filtered_type_data = filtered_data.dropna(subset=['Type of TVC (F/E/M)', 'Type of ICA (F/E/M)'])
            filtered_combinations_data = filtered_type_data[
                (((filtered_type_data['Type of TVC (F/E/M)'] == 'E') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
                 ((filtered_type_data['Type of TVC (F/E/M)'] == 'F') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')) |
                 ((filtered_type_data['Type of TVC (F/E/M)'] == 'M') & (filtered_type_data['Type of ICA (F/E/M)'] == 'F')))
            ]

            combinations = {
                "E(TVC) + F(ICA)": {'TVC': 'E', 'ICA': 'F'},
                "F(TVC) + F(ICA)": {'TVC': 'F', 'ICA': 'F'},
                "M(TVC) + F(ICA)": {'TVC': 'M', 'ICA': 'F'}
            }

            combination_metrics = {}
            for combo_name, combo_values in combinations.items():
                combo_data = filtered_combinations_data[
                    (filtered_combinations_data['Type of TVC (F/E/M)'] == combo_values['TVC']) &
                    (filtered_combinations_data['Type of ICA (F/E/M)'] == combo_values['ICA'])
                ]
                avg_spont_brand_uplift = combo_data['Spont Brand Uplift (%)'].mean()
                record_count = combo_data.shape[0]
                combination_metrics[combo_name] = {
                    "Average Spont Brand Uplift (%)": avg_spont_brand_uplift,
                    "Record Count": record_count
                }

            average_uplift_for_current_type = combination_metrics[creative_type]["Average Spont Brand Uplift (%)"]

            result = f"--- Analysis Results ---\n"
            result += f"Current Brand: {brand_name}\n"
            result += f"Current Spont Brand Uplift: {current_spont_brand_uplift:.2f}%\n"
            result += f"Average for {current_brand_size} Brands: {average_spont_uplift_for_size:.2f}%\n"

            if current_spont_brand_uplift > average_spont_uplift_for_size:
                result += f"The current ad shows a **significant improvement**.\n"
            else:
                result += f"The current ad does **not show a significant improvement**.\n"

            result += "\nAverage Spont Brand Uplift by Brand Size:\n"
            for size, avg_uplift in average_spont_brand_uplifts.items():
                result += f"  {size}: {avg_uplift:.2f}%\n"

            result += "\nNumber of Brands by Size:\n"
            for size, count in count_spont_brand_uplifts.items():
                result += f"  {size}: {count}\n"

            result += "\n--- Type of TVC vs Type of ICA Analysis ---\n"
            for combo, metrics in combination_metrics.items():
                result += f"\nCombination: {combo}\n"
                result += f"  Average Spont Brand Uplift (%): {metrics['Average Spont Brand Uplift (%)']:.2f}\n"
                result += f"  Record Count: {metrics['Record Count']}\n"

            result += f"\n--- Comparison for Creative Type: {creative_type} ---\n"
            result += f"Current Ad Spont Brand Uplift: {current_spont_brand_uplift:.2f}%\n"
            result += f"Average Spont Brand Uplift for {creative_type}: {average_uplift_for_current_type:.2f}%\n"

            if current_spont_brand_uplift > average_uplift_for_current_type:
                result += f"The current ad shows a **significant improvement** compared to the average for the same creative type.\n"
            else:
                result += f"The current ad does **not show a significant improvement** compared to the average for the same creative type.\n"

            self.output_text.delete("1.0", tk.END)
            self.output_text.insert(tk.END, result)

        except Exception as e:
            messagebox.showerror("Error", str(e))

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
