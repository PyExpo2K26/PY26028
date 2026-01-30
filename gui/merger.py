import pandas as pd

def merge_excel(
    file1,
    file2,
    output_file,
    common_column,
    how="inner"
):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # Clean column names
    df1.columns = df1.columns.str.strip().str.lower()
    df2.columns = df2.columns.str.strip().str.lower()

    common_column = common_column.lower()

    if common_column not in df1.columns or common_column not in df2.columns:
        raise ValueError("Common column not found in both files")

    merged_df = pd.merge(
        df1,
        df2,
        on=common_column,
        how=how
    )

    merged_df.to_excel(output_file, index=False)
