import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import warnings
warnings.filterwarnings('ignore')

def analyze_dataset(file_path, output_format='html'):
    """
    Comprehensive analysis of the dataset including missing values, data types,
    duplicates, constant columns, outliers, and distributions.
    
    Parameters:
    -----------
    file_path : str
        Path to the Excel file
    output_format : str
        Output format ('html', 'pdf', or 'word')
    """
    # Read the data
    print("Reading the dataset...")
    df = pd.read_excel(file_path)
    
    # Store original shape
    original_shape = df.shape
    
    # 1. Missing Values Analysis
    print("\n1. Missing Values Analysis")
    missing_vals = df.isnull().sum()
    missing_vals = missing_vals[missing_vals > 0]
    print("\nColumns with missing values:")
    print(missing_vals)
    
    # 2. Data Type Categorization
    print("\n2. Data Type Analysis")
    numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns
    
    print("\nNumeric columns:")
    print(list(numeric_cols))
    print("\nCategorical columns:")
    print(list(categorical_cols))
    
    # 3. Duplicates Analysis
    print("\n3. Duplicates Analysis")
    duplicated_cols = []
    for col in df.columns:
        if df[col].duplicated().any():
            duplicated_cols.append(col)
    
    print("\nColumns with duplicates:")
    print(duplicated_cols)
    
    print(f"\nShape before removing duplicates: {df.shape}")
    df_no_duplicates = df.drop_duplicates()
    print(f"Shape after removing duplicates: {df_no_duplicates.shape}")
    
    # 4. Constant Columns Analysis
    print("\n4. Constant Columns Analysis")
    constant_cols = []
    for col in df.columns:
        if df[col].nunique() == 1:
            constant_cols.append(col)
    
    print("\nConstant columns:")
    print(constant_cols)
    
    # Remove constant columns
    df_no_constant = df.drop(columns=constant_cols)
    print(f"\nShape before removing constant columns: {df.shape}")
    print(f"Shape after removing constant columns: {df_no_constant.shape}")
    
    # 5. Create visualizations
    print("\n5. Creating visualizations...")
    
    # Box plots for numeric columns
    plt.figure(figsize=(15, 10))
    for i, col in enumerate(numeric_cols[:6], 1):  # Limit to first 6 numeric columns
        plt.subplot(2, 3, i)
        sns.boxplot(y=df[col])
        plt.title(f'Box Plot of {col}')
        plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('boxplots.png')
    plt.close()
    
    # Distribution plots
    plt.figure(figsize=(15, 10))
    for i, col in enumerate(numeric_cols[:6], 1):  # First 6 numeric columns
        plt.subplot(2, 3, i)
        sns.histplot(df[col], kde=True)
        plt.title(f'Distribution of {col}')
        plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('distributions.png')
    plt.close()
    
    # Generate HTML report
    html_content = f"""
    <h1>Data Analysis Report</h1>
    
    <h2>1. Missing Values Analysis</h2>
    {missing_vals.to_frame().to_html()}
    
    <h2>2. Data Type Analysis</h2>
    <h3>Numeric Columns:</h3>
    <p>{', '.join(numeric_cols)}</p>
    <h3>Categorical Columns:</h3>
    <p>{', '.join(categorical_cols)}</p>
    
    <h2>3. Duplicates Analysis</h2>
    <p>Duplicated columns: {', '.join(duplicated_cols)}</p>
    <p>Original shape: {original_shape}</p>
    <p>Shape after removing duplicates: {df_no_duplicates.shape}</p>
    
    <h2>4. Constant Columns Analysis</h2>
    <p>Constant columns: {', '.join(constant_cols)}</p>
    <p>Shape after removing constant columns: {df_no_constant.shape}</p>
    
    <h2>5. Visualizations</h2>
    <h3>Box Plots</h3>
    <img src='boxplots.png' alt='Box Plots'>
    <h3>Distributions</h3>
    <img src='distributions.png' alt='Distributions'>
    """
    
    # Save the report
    with open('data_analysis_report.html', 'w') as f:
        f.write(html_content)
    
    print("\nAnalysis complete! Check data_analysis_report.html for the full report.")

if __name__ == "__main__":
    # Replace with your Excel file path
    file_path = "DS_Python_Assignment.xlsx"
    analyze_dataset(file_path)
