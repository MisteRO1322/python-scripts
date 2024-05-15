import pandas as pd

# Load the Excel file
input_excel_file = 'C:/Users/mowens/OneDrive - CARACOLE, INC/Desktop/ExcelTestFolder/At-Home Test Kit Order Form.xlsx'
df = pd.read_excel(input_excel_file)

# Define groups of zip codes
hc_zip_codes = [45011, 45002, 45030, 45033, 45041, 45051, 45052, 45111, 45174, 45201, 45202, 45203, 45204, 45205, 45206, 45207, 45208, 45209, 45211, 45212, 45213, 45214, 45215,
                45216, 45217, 45218, 45219, 45220, 45221, 45222, 45223, 45224, 45225, 45226, 45227, 45228, 45229, 45230, 45231, 45232, 45233, 45234, 45235, 45236, 45237, 45238,
                45239, 45240, 45241, 45242, 45243, 45244, 45246, 45247, 45248, 45249, 45250, 45251, 45252, 45253, 45254, 45255, 45258, 45262, 45263, 45264, 45267, 45268, 45269,
                45270, 45271, 45273, 45274, 45275, 45277, 45280, 45296, 45298, 45299, 45999]
butler_zip_codes = [45003, 45004, 45011, 45012, 45013, 45104, 45105, 45018, 45042, 45044, 45050, 45053, 45055, 45056, 45061, 45062, 45063, 45064, 45067, 45069, 45071]
clermont_zip_codes = [45102, 45103, 45106, 45112, 45120, 45122, 45140, 45147, 45150, 45153, 45156, 45157, 45158 ,45160, 45176, 45245]
nky_zip_codes = [45102, 45103, 45106, 45112, 45120, 45122, 45140, 45147, 45150, 45153, 45156, 45157, 45158 ,45160, 45176, 45245, 41011, 41012, 41014,
                 41015, 41016, 41017, 41018, 41019, 41025, 41051, 41053, 41063, 41001, 41007, 41059, 41071, 41072, 41073, 41074, 41075, 41076, 41085,
                 41099, 41005, 41021, 41022, 41042, 41048, 41080, 41091, 41092, 41094, 45105, 45144, 45616, 45618, 45650, 45660, 45679, 45684, 45693,
                 45697, 41501, 41502, 41503, 41512, 41513, 41514, 41519, 41520, 41522, 41524, 41526, 41527, 41528, 41531, 41534, 41535, 41538, 41539,
                 41540, 41542, 41543, 41544, 41547, 41548, 41549, 41553, 41554, 41555, 41557, 41558, 41559, 41560, 41561, 41562, 41563, 41564, 41566,
                 41567, 41568, 41571, 41572, 41010, 41030, 41035, 41052, 41054, 41097]

# Filter rows for hc
hc_df = df[df['ZIP Code:'].isin(hc_zip_codes)]

# Save hc to a new Excel file
output_hc_file = 'C:/Users/mowens/OneDrive - CARACOLE, INC/Desktop/ExcelTestFolder/hc_filtered_rows.xlsx'
hc_df.to_excel(output_hc_file, index=False)

# Filter rows for butler
butler_df = df[df['ZIP Code:'].isin(butler_zip_codes)]

# Save butler to a new Excel file
output_butler_file = 'C:/Users/mowens/OneDrive - CARACOLE, INC/Desktop/ExcelTestFolder/butler_filtered_rows.xlsx'
butler_df.to_excel(output_butler_file, index=False)

# Filter rows for clermont
clermont_df = df[df['ZIP Code:'].isin(clermont_zip_codes)]

# Save clermont to a new Excel file
output_clermont_file = 'C:/Users/mowens/OneDrive - CARACOLE, INC/Desktop/ExcelTestFolder/clermont_filtered_rows.xlsx'
clermont_df.to_excel(output_clermont_file, index=False)

# Filter rows for nky
nky_df = df[df['ZIP Code:'].isin(nky_zip_codes)]

# Save nky to a new Excel file
output_nky_file = 'C:/Users/mowens/OneDrive - CARACOLE, INC/Desktop/ExcelTestFolder/nky_filtered_rows.xlsx'
nky_df.to_excel(output_nky_file, index=False)

print("Script executed successfully.")
