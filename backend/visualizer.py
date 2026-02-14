import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def get_data(file_path):
    df = pd.read_excel(file_path)
    numeric_cols = df.select_dtypes(include=["int64", "float64"]).columns.tolist()
    categorical_cols = df.select_dtypes(include=["object"]).columns.tolist()
    return df, numeric_cols, categorical_cols

# Individual Chart Functions
def show_histogram(df, column):
    plt.figure(figsize=(8, 4))
    sns.histplot(df[column], kde=True)
    plt.title(f'Distribution of {column}')
    plt.show()

def show_pie_chart(df, column):
    plt.figure(figsize=(8, 8))
    data = df[column].value_counts().head(10)
    plt.pie(data, labels=data.index, autopct='%1.1f%%', startangle=140)
    plt.title(f'{column} Distribution')
    plt.show()

def show_scatter(df, col1, col2):
    plt.figure(figsize=(8, 4))
    sns.scatterplot(data=df, x=col1, y=col2)
    plt.title(f'{col1} vs {col2}')
    plt.show()

def show_correlation_heatmap(df, numeric_cols):
   
    if len(numeric_cols) < 2:
        print("Error: Correlation heatmap requires at least 2 numeric columns.")
        return

    plt.figure(figsize=(10, 8)) 
    corr_matrix = df[numeric_cols].corr()
    sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".2f", linewidths=0.5)
    plt.title('Correlation Heatmap')
    plt.show()