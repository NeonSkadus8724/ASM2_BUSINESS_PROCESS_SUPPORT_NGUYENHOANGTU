import pandas as pd

# Function to remove special characters
def clean_special_characters(text):
    # Define special characters
    special_characters = "!@#$%^&*()_+{}:\"<>?[];',./\\`~"
    # Remove special characters
    cleaned_text = ''.join(char for char in text if char.isalnum() or char.isspace() or char == '-')
    return cleaned_text

# Function to add spaces before capitals
def add_space_before_capitals(text):
    modified_text = ""
    # Iterate through each character in the text
    for i, char in enumerate(text):
        # If the character is a capital letter and the previous character is not a space, add a space before it
        if char.isupper() and i > 0 and text[i - 1] != ' ':
            modified_text += " " + char
        else:
            modified_text += char
    return modified_text.strip()  # Strip leading and trailing spaces

# Specify the path to the Excel file
excel_path = "data.xlsx"

# Read the Excel file into a DataFrame
df = pd.read_excel(excel_path)

# Clean special characters in all columns
df = df.applymap(lambda x: clean_special_characters(str(x)))

# Add spaces before capitals in specific columns
columns_to_modify = ['Product Name', 'Supplier', 'Category']
for column in columns_to_modify:
    df[column] = df[column].apply(add_space_before_capitals)

# Convert columns to title case
df['Supplier'] = df['Supplier'].str.title()
df['Category'] = df['Category'].str.title()
df['Product Name'] = df['Product Name'].str.title()

# Sort the DataFrame by Product ID and remove duplicates
df = df.astype({'Product ID': int}).sort_values(by='Product ID').drop_duplicates().reset_index(drop=True)

# Display the cleaned DataFrame
print(df)

# Specify the path for the cleaned Excel file
cleaned_excel_path = "cleaned_data.xlsx"

# Export the cleaned DataFrame to a new Excel file
df.to_excel(cleaned_excel_path, index=False)

print("Cleaned data has been exported to", cleaned_excel_path)
