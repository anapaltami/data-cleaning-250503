import pandas as pd
import scrubadub

# load the data
df = pd.read_excel('data/FAKE-datasetPII.xlsx')

# scrub the data for PII
scrubber = scrubadub.Scrubber()
for column in df.select_dtypes(include=['object']):
    df[column] = df[column].apply(lambda x: scrubber.clean(str(x)) if pd.notna(x) else x)

# drop columns with known PII
df.drop(columns=['First and Last Name', 'SSN', 'ADDRESS', 'DOB', 'PHONE'], inplace=True, errors='ignore')

# save the cleaned data to output folder
df.to_csv('output/cleaned_data.csv', index=False)
print("Data cleaned and saved to output/cleaned_data.csv")