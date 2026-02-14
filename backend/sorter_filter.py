import pandas as pd
import os

def sort_and_filter(
    input_file,
    output_file,
    sort_column=None,
    ascending=True,
    filter_column=None,
    filter_value=None
):
    df = pd.read_excel(input_file)
    df.columns = df.columns.str.strip().str.lower()
    
    if sort_column:
        sort_column = sort_column.strip().lower()
    if filter_column:
        filter_column = filter_column.strip().lower()
        
    if filter_column and filter_value:
        if filter_column not in df.columns:
            raise ValueError(f"Filter column '{filter_column}' not found in Excel file")

        df = df[
            df[filter_column]
            .astype(str)
            .str.lower()
            == str(filter_value).lower()
        ]
        
    if sort_column:
        if sort_column not in df.columns:
            raise ValueError(f"Sort column '{sort_column}' not found in Excel file")

        df = df.sort_values(by=sort_column, ascending=ascending)
        
    

    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    df.to_excel(output_file, index=False)
    return output_file
    
