import pandas as pd
import scrubadub
import re

# load the data
df = pd.read_excel('data/FAKE-datasetPII.xlsx', header=1)

# rename columns for easier handling
df.columns = df.columns.str.strip().str.upper()

# extract ZIP from ADDRESS before dropping
def extract_zip(address):
    # look for a 5-digit ZIP code in the address before 'USA' or at the end of the string
    match = re.search(r'(\d{5})(?=\s*USA|$)', str(address))
    return match.group(1) if match else None
df['ZIP'] = df['ADDRESS'].apply(extract_zip)

# PII columns to be removed
pii_columns = [
    'UNNAMED: 0',        # Full Name
    'UNNAMED: 1',        # SSN
    'VISA MC AMEX',      # Credit Card #
    'ADDRESS',           # Full address (ZIP now extracted)
    'DOB',
    'PHONE'
]
df.drop(columns=pii_columns, inplace=True, errors='ignore')

# scrub the data for PII
scrubber = scrubadub.Scrubber()
for column in df.select_dtypes(include=['object']):
    df[column] = df[column].apply(lambda x: scrubber.clean(str(x)) if pd.notna(x) else x)

# drop columns with known PII
df.drop(columns=['First and Last Name', 'SSN', 'ADDRESS', 'DOB', 'PHONE'], inplace=True, errors='ignore')

# save the cleaned data to output folder
df.to_csv('output/cleaned_data.csv', index=False)
print("Data cleaned and saved to output/cleaned_data.csv")