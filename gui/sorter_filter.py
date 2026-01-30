import pandas as pd


def sort_and_filter(input_file, output_file, sort_column, city_filter=None):
    df = pd.read_excel(input_file)

    # Clean column names
    df.columns = df.columns.str.strip().str.lower()

    # Sort data
    if sort_column.lower() in df.columns:
        df = df.sort_values(by=sort_column.lower())

    # Filter by city if provided
    if city_filter and 'city' in df.columns:
        df = df[df['city'].str.lower() == city_filter.lower()]

    df.to_excel(output_file, index=False)


if __name__ == "__main__":
    # Example usage when running directly
    sort_and_filter(
        input_file="input.xlsx",
        output_file="output.xlsx",
        sort_column="name",
        city_filter="New York"
    )
