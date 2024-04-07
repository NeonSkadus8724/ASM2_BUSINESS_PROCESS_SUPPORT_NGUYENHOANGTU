import pandas as pd

def clean_special_characters(text):
    """Remove special characters from text."""
    if isinstance(text, str):  
        return ''.join(e for e in text if e.isalnum() or e.isspace())
    else:
        return str(text)  

def clean_dataframe(df):
    """Clean DataFrame by applying clean_special_characters to each cell."""
    cleaned_df = df.applymap(clean_special_characters)
    cleaned_df.drop_duplicates(inplace=True)
    return cleaned_df

def main():
    excel_path = "data.xlsx"
    df = pd.read_excel(excel_path)
    cleaned_df = clean_dataframe(df)
    cleaned_excel_path = "cleaned_data.xlsx"
    cleaned_df.to_excel(cleaned_excel_path, index=False)
    print("Cleaned data has been exported to", cleaned_excel_path)
if __name__ == "__main__":
    main()