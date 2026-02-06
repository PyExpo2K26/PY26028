import pandas as pd
import os


def validate_rule(df, selected_columns, operation):
    if not isinstance(selected_columns, list) or len(selected_columns) < 2:
        raise ValueError("At least two columns must be selected")

    for idx in selected_columns:
        if not isinstance(idx, int):
            raise ValueError("Column index must be an integer")
        if idx < 0 or idx >= len(df.columns):
            raise ValueError(f"Column index {idx} does not exist")
        if not pd.api.types.is_numeric_dtype(df.iloc[:, idx]):
            raise ValueError(f"Column '{df.columns[idx]}' is not numeric")

    if operation in ("subtract", "divide") and len(selected_columns) != 2:
        raise ValueError(f"Operation '{operation}' requires exactly two columns")


def apply_calculated_column(df, selected_columns, operation, new_column_name):
    validate_rule(df, selected_columns, operation)

    data = df.iloc[:, selected_columns]

    if operation == "add":
        df[new_column_name] = data.sum(axis=1)
    elif operation == "multiply":
        df[new_column_name] = data.prod(axis=1)
    elif operation == "subtract":
        df[new_column_name] = data.iloc[:, 0] - data.iloc[:, 1]
    elif operation == "divide":
        df[new_column_name] = data.iloc[:, 0] / data.iloc[:, 1]
    else:
        raise ValueError(f"Unsupported operation: {operation}")

    return df


def process_excel_file(
    input_file_path,
    output_dir,
    selected_columns,
    operation,
    new_column_name
):
    if not os.path.exists(input_file_path):
        raise FileNotFoundError("Input Excel file not found")

    os.makedirs(output_dir, exist_ok=True)

    df = pd.read_excel(input_file_path)

    df = apply_calculated_column(
        df=df,
        selected_columns=selected_columns,
        operation=operation,
        new_column_name=new_column_name
    )

    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    output_file_path = os.path.join(
        output_dir,
        f"{base_name}_processed.xlsx"
    )

    df.to_excel(output_file_path, index=False)

    return output_file_path
