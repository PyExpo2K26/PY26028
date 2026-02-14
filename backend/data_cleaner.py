import pandas as pd
import os 

def clean_excel(input_file, output_file):
    df = pd.read_excel(input_file)
    
    #cleaning column names
    df.columns = df.columns.astype(str).str.strip()
    
    #removing empty rows and columns
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    
    #removing duplicates
    df = df.drop_duplicates()

    #filling missing values
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(df[col].mean())
        else:
            df[col] = df[col].fillna("")

    #ensuring output directory exists
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)
    return output_file
