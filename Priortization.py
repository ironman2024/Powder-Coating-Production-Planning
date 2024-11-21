import pandas as pd
import numpy as np

file_path = 'Week Production Plan as of 2024-09-20 (2).xlsx'

sheet_names = pd.ExcelFile(file_path).sheet_names
print("Available sheets:", sheet_names)

output_file = 'prioritized_plan_combined.xlsx'
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:

    # Define a function to process each sheet
    def process_sheet(sheet_name):
        # Load the specified sheet using the correct row for headers (row 1)
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)

        # Normalize column names by stripping spaces and converting to lowercase
        df.columns = df.columns.str.strip().str.lower()

        # Check for the required columns depending on the sheet
        if sheet_name == 'MTO':
            required_columns = ['site', 'ageing', 'total balance to produce', 'mfg pro']
        elif sheet_name == 'MTS':
            required_columns = ['site', 'coverage', 'type', 'production plan', 'mfg pro', 'mts/rts', 'dio']
        else:
            raise ValueError(f"Unknown sheet name '{sheet_name}'. Only 'MTO' and 'MTS' are supported.")

        # Check if the required columns exist
        for col in required_columns:
            if col not in df.columns:
                raise KeyError(f"Column '{col}' is missing from the Excel file. Available columns: {df.columns.tolist()}")

        # Processing for 'MTO' sheet
        if sheet_name == 'MTO':
            # Step 1: Filter products made in Q0CP
            df_filtered = df[df['site'].str.lower() == 'q0cp']

            # Step 2: Sort by Ageing in descending order, and Total Balance to Produce in case of a tie
            df_sorted = df_filtered.sort_values(by=['ageing', 'total balance to produce'], ascending=[False, False])

            # Step 3: Select relevant columns for the final output
            final_columns = ['mfg pro', 'ageing', 'total balance to produce']
            df_final = df_sorted[final_columns]

        # Processing for 'MTS' sheet
        elif sheet_name == 'MTS':
            # Step 1: Filter products made in Thane
            df_filtered = df[df['site'].str.lower() == 'thane']

            # Step 2: Remove rows where the production plan is 0
            df_filtered = df_filtered[df_filtered['production plan'] != 0]

            # Step 3: Round DIO to the nearest integer
            df_filtered['dio'] = np.round(df_filtered['dio']).astype(int)

            # Step 4: Filter out rows where DIO is above 30
            df_filtered = df_filtered[df_filtered['dio'] <= 30]

            # Step 5: Assign a secondary priority based on MTS/RTS
            # Higher priority for MTO/RTS
            df_filtered['priority_mts'] = df_filtered['mts/rts'].apply(lambda x: 0 if x in ['MTO', 'RTS'] else 1)

            # Step 6: Sort by DIO and MTS/RTS priority, then by production plan
            df_sorted = df_filtered.sort_values(by=['dio', 'priority_mts', 'production plan'], ascending=[True, True, False])

            # Step 7: Custom sorting condition for DIO gaps and production plan
            for i in range(len(df_sorted) - 1):
                current_dio = df_sorted.iloc[i]['dio']
                next_dio = df_sorted.iloc[i + 1]['dio']
                current_prod_plan = df_sorted.iloc[i]['production plan']
                next_prod_plan = df_sorted.iloc[i + 1]['production plan']

                # Check if DIO gap is less than 1 and production plan difference is greater than 10,000
                if abs(current_dio - next_dio) < 1 and abs(current_prod_plan - next_prod_plan) > 10000:
                    # Swap the two rows if the next product has a higher production plan
                    if next_prod_plan > current_prod_plan:
                        df_sorted.iloc[[i, i + 1]] = df_sorted.iloc[[i + 1, i]]

            # Step 8: Select relevant columns for the final output
            final_columns = ['mfg pro', 'dio', 'mts/rts', 'production plan']
            df_final = df_sorted[final_columns]

        # Write the final DataFrame to the Excel writer under a new sheet
        df_final.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"Prioritization plan for '{sheet_name}' processed and added to the output file.")

    # Process both MTO and MTS sheets
    for sheet in ['MTO', 'MTS']:
        if sheet in sheet_names:
            process_sheet(sheet)
        else:
            print(f"Sheet '{sheet}' not found in the Excel file.")

# Output the result
print(f"Combined prioritization plans saved to {output_file}")
