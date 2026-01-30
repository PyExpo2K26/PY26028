import pandas as pd

def clean_excel(input_file, output_file):
    df = pd.read_excel(input_file)

    print("Before Cleaning:")
    print(df)

    df.columns = df.columns.str.strip().str.lower()
    df = df.drop_duplicates()

    if 'age' in df.columns:
        df['age'] = df['age'].fillna(df['age'].mean())

    if 'city' in df.columns:
        df['city'] = df['city'].fillna("Unknown")

    df.to_excel(output_file, index=False)

    print("After Cleaning:")
    print(df)
