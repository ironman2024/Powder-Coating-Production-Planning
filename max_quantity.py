import pandas as pd
import openpyxl
import re
from colormath.color_objects import LabColor
from colormath.color_diff import delta_e_cie2000

# Load the workbook to access cell formatting and data
file_path = 'database2024.xlsx'
workbook = openpyxl.load_workbook(file_path, data_only=True)

# Define patterns to identify Line and Volume columns
line_pattern = re.compile(r'Line\s*-\s*(\d+)', re.IGNORECASE)
volume_pattern = re.compile(r'Volume(\.\d+)?', re.IGNORECASE)

# Initialize dictionaries to store unique product data with additional fields
mto_product_data = {}
mts_product_data = {}

# Helper function to determine valid lines based on shade
def get_valid_lines(lab_values):
    l_value = lab_values[0]
    # Dark shades can only run on Lines 1 and 4
    if l_value < 50:  # Threshold for dark/light determination
        return [1, 4]
    # Light shades can only run on Lines 2 and 3
    return [2, 3]

# Process each sheet in the workbook
for sheet_name in workbook.sheetnames:
    df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=None)
    df = df.reset_index(drop=True)
    
    # Add columns for family, resin, and LAB values if not present
    required_cols = ['Family', 'Resin', 'L_value', 'a_value', 'b_value']
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
    
    line_volume_pairs = {}
    for i, col in enumerate(df.columns):
        line_match = line_pattern.match(str(col))
        if line_match and i + 1 < len(df.columns):
            line_num = line_match.group(1)
            next_col = df.columns[i + 1]
            if volume_pattern.match(str(next_col)):
                line_volume_pairs[f"Line-{line_num}"] = {
                    'product_col': col,
                    'volume_col': next_col
                }

    for line, cols in line_volume_pairs.items():
        line_data = df[[cols['product_col'], cols['volume_col'], 'Class', 'Family', 'Resin', 
                       'L_value', 'a_value', 'b_value']].copy()
        line_data = line_data.dropna(subset=[cols['product_col'], 'Class'])
        line_data[cols['volume_col']] = pd.to_numeric(line_data[cols['volume_col']], errors='coerce')
        
        line_data_for_freq = line_data.dropna(subset=[cols['volume_col']])
        line_data_for_max = line_data[line_data[cols['volume_col']] > 0].dropna(subset=[cols['volume_col']])

        grouped_freq = line_data_for_freq.groupby([cols['product_col'], 'Class']).size().reset_index(name='count')
        grouped_max = line_data_for_max.groupby([cols['product_col'], 'Class'])[cols['volume_col']].max().reset_index()

        grouped_data = pd.merge(grouped_freq, grouped_max, on=[cols['product_col'], 'Class'], how='outer')

        for _, row in grouped_data.iterrows():
            product_name = str(row[cols['product_col']])
            product_class = str(row['Class'])
            max_volume = row[cols['volume_col']] if not pd.isna(row[cols['volume_col']]) else 0
            frequency = row['count']

            target_dict = mto_product_data if product_class == 'MTO' else mts_product_data
            if product_name not in target_dict:
                target_dict[product_name] = {
                    'Product': product_name,
                    'Line-1 freq': 0, 'Line-1 Max Volume': 0,
                    'Line-2 freq': 0, 'Line-2 Max Volume': 0,
                    'Line-3 freq': 0, 'Line-3 Max Volume': 0,
                    'Line-4 freq': 0, 'Line-4 Max Volume': 0
                }

            target_dict[product_name][f"{line} freq"] += frequency
            target_dict[product_name][f"{line} Max Volume"] = max(
                target_dict[product_name][f"{line} Max Volume"],
                max_volume
            )

# Resolve duplicates between MTO and MTS based on total frequency
duplicate_products = set(mto_product_data.keys()).intersection(set(mts_product_data.keys()))
for product in duplicate_products:
    mto_total_freq = sum(mto_product_data[product][f'Line-{i} freq'] for i in range(1, 5))
    mts_total_freq = sum(mts_product_data[product][f'Line-{i} freq'] for i in range(1, 5))
    
    if mto_total_freq >= mts_total_freq:
        mts_product_data.pop(product)
    else:
        mto_product_data.pop(product)

# Combine MTO and MTS data
all_products = {**mto_product_data, **mts_product_data}

# Create final summary DataFrame
summary_data = []
for product, data in all_products.items():
    # Get frequencies and volumes for all lines
    line_stats = [(i, data[f'Line-{i} freq'], data[f'Line-{i} Max Volume']) 
                  for i in range(1, 5)]
    
    # Sort by frequency (descending), then by volume (descending)
    line_stats.sort(key=lambda x: (-x[1], -x[2]))
    
    # Get top two lines
    priority1_line = f"Line {line_stats[0][0]}" if line_stats[0][1] > 0 else "N/A"
    volume1 = line_stats[0][2] if line_stats[0][1] > 0 else 0
    
    priority2_line = f"Line {line_stats[1][0]}" if len(line_stats) > 1 and line_stats[1][1] > 0 else "N/A"
    volume2 = line_stats[1][2] if len(line_stats) > 1 and line_stats[1][1] > 0 else 0
    
    summary_data.append({
        'Product': product,
        'Priority 1': priority1_line,
        'Volume 1': volume1,
        'Priority 2': priority2_line,
        'Volume 2': volume2
    })

# Create final DataFrame and sort by Product
summary_df = pd.DataFrame(summary_data)
summary_df = summary_df.sort_values('Product')

# Save to Excel
output_file = 'Product_Summary.xlsx'
with pd.ExcelWriter(output_file) as writer:
    summary_df.to_excel(writer, sheet_name='Product Summary', index=False)

print(f"\nOutput saved to {output_file}")
