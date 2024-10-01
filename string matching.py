import pandas as pd
from fuzzywuzzy import fuzz, process

df_details = pd.read_excel(r'data.xlsx')  
df_simplified = pd.read_excel(r'data file.xlsx') 

detailed_descriptions = df_details[['Description', 'Quantity', 'Rate']]
simplified_descriptions = df_simplified[['Description', 'Quantity', 'Rate']]

detailed_descriptions['Description'] = detailed_descriptions['Description'].astype(str).fillna('')
simplified_descriptions['Description'] = simplified_descriptions['Description'].astype(str).fillna('')


def find_best_match(detailed_desc):
    try:
        
        result = process.extractOne(detailed_desc, simplified_descriptions['Description'], scorer=fuzz.partial_ratio)
        if result:
            match_description = result[0]
            score = result[1]
            
            matched_row = simplified_descriptions[simplified_descriptions['Description'] == match_description].iloc[0]
            return match_description, score, matched_row
        else:
            return 'No match', 0, None
    except Exception as e:
        print(f"Error finding match for description: {detailed_desc}")
        print(f"Exception: {e}")
        return 'Error', 0, None

matches = [find_best_match(desc) for desc in detailed_descriptions['Description']]


results = []
for index, desc in detailed_descriptions.iterrows():
    match_description, score, matched_row = matches[index]
    quantity_match = (desc['Quantity'] == matched_row['Quantity']) if matched_row is not None else False
    rate_match = (desc['Rate'] == matched_row['Rate']) if matched_row is not None else False
    results.append({
        'Detailed Description': desc['Description'],
        'Best Match': match_description,
        'Match Score': score,
        'Quantity Match': quantity_match,
        'Rate Match': rate_match
    })

results_df = pd.DataFrame(results)


results_df.to_excel('matched_results.xlsx', index=False)

print("Matching completed. Results saved to 'matched_results.xlsx'.")