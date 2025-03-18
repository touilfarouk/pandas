"# pandas" 

import pandas as pd

# Load CSV files pip install panda
questionnaire = pd.read_csv("questionnaire.csv", low_memory=False)
utilisation_du_sol = pd.read_csv("utilisation_du_sol.csv")

# Merge data on 'id_questionnaire' using a left join
merged_data = pd.merge(utilisation_du_sol, questionnaire, on="id_questionnaire", how="left")

# Pivot transformation
grouped_data = merged_data.groupby("id_questionnaire").agg(lambda x: list(x))

# Reformat data: Convert lists to dynamic columns
transformed_rows = []
for idx, row in grouped_data.iterrows():
    new_row = {"id_questionnaire": idx, "exploitant_cle_unique": row["exploitant_cle_unique"][0]}
    for i, (culture, superficie) in enumerate(zip(row["code_culture"], row["superficie_hec"]), start=1):
        new_row[f"code_culture{i}"] = culture
        new_row[f"superficie_hec{i}"] = superficie
    transformed_rows.append(new_row)

# Convert back to DataFrame
final_df = pd.DataFrame(transformed_rows)

# Export to Excel
output_file = "questionnaire_transformed.xlsx"
final_df.to_excel(output_file, index=False)

print(f"Transformation successful! File saved as {output_file}")





# import pandas as pd

# # Load CSV files
# questionnaire = pd.read_csv("questionnaire.csv", low_memory=False)
# utilisation_du_sol = pd.read_csv("utilisation_du_sol.csv")

# # Merge data on 'id_questionnaire' using a left join
# merged_data = pd.merge(utilisation_du_sol, questionnaire, on="id_questionnaire", how="left")

# # Export to Excel
# output_file = "questionnaire_report.xlsx"
# merged_data.to_excel(output_file, index=False)

# print(f"Merge successful! File saved as {output_file}")
