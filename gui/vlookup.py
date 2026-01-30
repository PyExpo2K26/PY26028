import pandas as pd

def vlookup_excel(
    main_file,
    lookup_file,
    output_file,
    key_column,
    lookup_column
):
    df_main = pd.read_excel(main_file)
    df_lookup = pd.read_excel(lookup_file)

    # normalize column names
    df_main.columns = df_main.columns.str.strip().str.lower()
    df_lookup.columns = df_lookup.columns.str.strip().str.lower()

    key_column = key_column.lower()
    lookup_column = lookup_column.lower()

    # perform lookup using merge
    merged = pd.merge(
        df_main,
        df_lookup[[key_column, lookup_column]],
        on=key_column,
        how="left"
    )

    merged.to_excel(output_file, index=False)
