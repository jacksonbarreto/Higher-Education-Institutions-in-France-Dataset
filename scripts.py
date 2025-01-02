import pandas as pd

# Import eter-export-2021-FR.xlsx
df = pd.read_excel('eter-export-2021-FR.xlsx')

# Replace column 'ETER ID' column
df.rename(columns={'ETER ID': 'ETER_ID'}, inplace=True)

# Replace column 'Institution Name' column
df.rename(columns={'Institution Name': 'Name'}, inplace=True)

# Replace column 'Legal status' column
df.rename(columns={'Legal status': 'Category'}, inplace=True)

# Replace column 'Institutional website' column
df.rename(columns={'Institutional website': 'Url'}, inplace=True)

# Replace column 'Institution Category standardized' column
df.rename(columns={'Institution Category standardized': 'Institution_Category_Standardized'}, inplace=True)

# Replace column 'Member of European University alliance' column
df.rename(columns={'Member of European University alliance': 'Member_of_European_University_alliance'}, inplace=True)

# Replace column 'Region of establishment (NUTS 2)' column
df.rename(columns={'Region of establishment (NUTS 2)': 'NUTS2'}, inplace=True)

# Replace column 'Region of establishment (NUTS 3)' column
df.rename(columns={'Region of establishment (NUTS 3)': 'NUTS3'}, inplace=True)

# Replace values ​​in 'Category' column
category_replaces = {0: 'public', 1: 'private', 2: 'public'}
df['Category'] = df['Category'].replace(category_replaces)

# Replace null values for specific ETER IDs and columns
for eter_id, category, institutions_category_standardized_replaces in [
    ('FR0343', 'public', 2),
    ('FR0344', 'private', 2),
    ('FR0364', 'public', 1),
    ('FR0365', 'public', 1),
    ('FR0366', 'public', 1),
    ('FR0367', 'public', 1),
    ('FR0369', 'public', 1),
    ('FR0463', 'public', 2),
    ('FR0466', 'public', 2),
    ('FR0468', 'public', 1),
    ('FR0864', 'public', 2),
    ('FR0866', 'public', 2),
    ('FR0867', 'public', 1),
    ('FR0877', 'public', 2),
    ('FR0945', 'public', 1),
]:
    df.loc[df['ETER_ID'] == eter_id, 'Category'] = category
    df.loc[df['ETER_ID'] == eter_id, 'Institution_Category_Standardized'] = institutions_category_standardized_replaces

# Replace the values ​​in the 'Institution Category standardized' column
institutions_category_replaces = {0: 'Other', 1: 'University', 2: 'University of applied sciences'}
df['Institution_Category_Standardized'] = df['Institution_Category_Standardized'].replace(institutions_category_replaces)

# Replace the value in the 'Member of European University alliance' column
member_of_European_University_alliance_replaces = {0: 'No', 1: 'Yes'}
df['Member_of_European_University_alliance'] = df['Member_of_European_University_alliance'].replace(member_of_European_University_alliance_replaces)

# Import NUTS2013-NUTS2016.xlsx and select the right sheet
dfNuts16Raw = pd.read_excel('NUTS2013-NUTS2016.xlsx', sheet_name='NUTS2013-NUTS2016', header=1)

# Create NUTS2 mapping DataFrame for 2016
dfNuts2_2016 = dfNuts16Raw[['Code 2016', 'NUTS level 2']].copy()
dfNuts2_2016.rename(columns={
    'Code 2016': 'NUTS2',
    'NUTS level 2': 'NUTS2_Label_2016'
}, inplace=True)

# Create NUTS3 mapping DataFrame for 2016
dfNuts3_2016 = dfNuts16Raw[['Code 2016', 'NUTS level 3']].copy()
dfNuts3_2016.rename(columns={
    'Code 2016': 'NUTS3',
    'NUTS level 3': 'NUTS3_Label_2016'
}, inplace=True)

# Merge df with NUTS2 and NUTS3 for 2016
df = pd.merge(df, dfNuts2_2016, on='NUTS2', how='left')
df = pd.merge(df, dfNuts3_2016, on='NUTS3', how='left')

# Import NUTS2021.xlsx
dfNuts21Raw = pd.read_excel('NUTS2021.xlsx', sheet_name='NUTS & SR 2021')

# Create NUTS2 mapping DataFrame for 2021
dfNuts2_2021 = dfNuts21Raw[['Code 2021', 'NUTS level 2']].copy()
dfNuts2_2021.rename(columns={
    'Code 2021': 'NUTS2',
    'NUTS level 2': 'NUTS2_Label_2021'
}, inplace=True)

# Create NUTS3 mapping DataFrame for 2021
dfNuts3_2021 = dfNuts21Raw[['Code 2021', 'NUTS level 3']].copy()
dfNuts3_2021.rename(columns={
    'Code 2021': 'NUTS3',
    'NUTS level 3': 'NUTS3_Label_2021'
}, inplace=True)

# Merge df with NUTS2 and NUTS3 for 2021
df = pd.merge(df, dfNuts2_2021, on='NUTS2', how='left')
df = pd.merge(df, dfNuts3_2021, on='NUTS3', how='left')

# Arrange NUTS related columns in desired order
nuts_columns = [
    'NUTS2', 'NUTS2_Label_2016', 'NUTS2_Label_2021',
    'NUTS3', 'NUTS3_Label_2016', 'NUTS3_Label_2021'
]

# Find columns that already exist in the DataFrame
existing_nuts_columns = [col for col in nuts_columns if col in df.columns]

# Determine the position where NUTS columns should be inserted (after 'Url')
insertion_point = df.columns.get_loc('Url') + 1

# Get remaining columns without duplicating or changing original order
remaining_columns = [col for col in df.columns if col not in existing_nuts_columns]

# Insert NUTS columns in desired position
new_column_order = (
    remaining_columns[:insertion_point]
    + existing_nuts_columns
    + remaining_columns[insertion_point:]
)

# Reorganize DataFrame with new column order
df = df[new_column_order]

##### Sanatize #####

# Remove White spaces
df['Url'] = df['Url'].str.strip()

# Remove the http and https from url
df['Url'] = df['Url'].str.replace(r'^https?://', '', regex=True)

# Remove the third bar from the url
df['Url'] = df['Url'].str.replace(r'/.*', '', regex=True)

##### Enrichment #####

# Remove FR0129 Institut polytechnique LaSalle Beauvais
# Because this polytechnique belongs to University LaSalle
df = df[df['ETER_ID'] != 'FR0129']



# Remove FR0944 École nationale des impôts
# Because was not possible to find an url
df = df[df['ETER_ID'] != 'FR0944']



# Remove FR0513 Institut supérieur européen de gestion Lyon
# Because it's the same url and school, so remains only the main campus in Paris
df = df[df['ETER_ID'] != 'FR0513']



# Remove FR0138 VetAgro Sup Lempdes
# Because it's the same url and school, so remains only the main campus in  Marcy-l'Étoile
df = df[df['ETER_ID'] != 'FR0138']



# Remove FR0816 Comue Université Paris-Lumière
# Because was closed in 12/2023
df = df[df['ETER_ID'] != 'FR0816']



# Remove FR0235 Institut supérieur de l'électronique et du numérique Toulon
# Because it's the same url of Institut supérieur de l'électronique et du numérique Lille, so remains only the main campus
df = df[df['ETER_ID'] != 'FR0235']

df = df.reset_index(drop=True)

# FORBIDEN * CHECK LATER
# Remove FR0106 École spéciale militaire de Saint-Cyr
# Because it's the url returns 403 forbiden
# df = df[df['ETER_ID'] != 'FR0106']
# Reset the index
# df = df.reset_index(drop=True)

# Remove FR0107 École militaire interarmes
# Because it's the same url and school and the same campus and city, so remains only the École spéciale militaire de Saint-Cyr
# df = df[df['ETER_ID'] != 'FR0107']
# Reset the index
# df = df.reset_index(drop=True)

# DNS PROBLEM * CHECK LATER
# Remove FR0883 Institut national des sciences et techniques nucléaires
# Because it's the url returns DNS_PROBE_POSSIBLE
# df = df[df['ETER_ID'] != 'FR0106']
# Reset the index
# df = df.reset_index(drop=True)

# Change URL of FR0333 École catholique d'arts et métiers Strasbourg-Europe
# The Past url was wrong
df.loc[df['ETER_ID'] == 'FR0333', 'Url'] = 'www.icam.fr'

# Change URL of FR0906 Ecole pratique du service social
# The Past url was wrong
df.loc[df['ETER_ID'] == 'FR0906', 'Url'] = 'epss.fr'

# Change URL of FR0104 Escola superior nacional darquitetura da Nancy
# The Past url was wrong
df.loc[df['ETER_ID'] == 'FR0104', 'Url'] = 'www.ensa-nancy.fr'

# Change URL of FR0466 Institut national polytechnique Clermont-Auvergne
# The Past url was wrong
df.loc[df['ETER_ID'] == 'FR0466', 'Url'] = 'www.clermont-auvergne-inp.fr'

# Change URL of FR0907 Ecole Nationale d'Administration
# This Ecole change your name and url, because in 31/12/2021 the
# L'École nationale d'administration (ENA) was close
# Changed for Institut national du service public (INSP) in 01/01/2022
df.loc[df['ETER_ID'] == 'FR0907', ['Url', 'Name']] = ['insp.gouv.fr', 'Institut national du service public']

# Invalid HTTPS
# CHECK LATER FR0970 École nationale de la meteorologie

# Save data to CSV file
df.to_csv('france-heis.csv', index=False, encoding='utf-8')
