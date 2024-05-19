import pandas as pd
import re
# You should still manually check titles that being with acronyms or abbreviations and titles with Greek letters
# Also check for possesive s: 's might get translated to 'S in the output file

# List of words to exclude from capitalization
exceptions = ['a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'if', 'in', 'nor', 'of', 'on', 'or', 'so', 'the', 'to',
              'up', 'yet']

def ap_title_case(text):
    def preserve_word(word):
        # Regular expressions to match:
        # Acronyms: all uppercase letters
        # Abbreviations: digits followed by uppercase letters or mixed uppercase and lowercase
        # Compound words separated by a slash or hyphen
        # Patterns with parentheses or colons followed by uppercase letters or digits
        # Greek letters like 'μ'
        if re.match(r'^[A-Z]{2,}$', word) or \
                re.match(r'^[0-9]+[A-Z]+$', word) or \
                re.match(r'^[A-Z0-9]+$', word) or \
                re.match(r'^[A-Z]+/[A-Z]+$', word) or \
                re.match(r'^[0-9]+[A-Z]+-[0-9]+[A-Z]+$', word) or \
                re.match(r'^\([A-Z0-9]+\)$', word) or \
                re.match(r'^[A-Z0-9]+:$', word) or \
                re.match(r'^[μ]+-', word) or \
                re.match(r'^μ+-', word) or \
                re.match(r'^[A-Za-z]+-[0-9]+$', word) or \
                re.match(r'^[A-Za-z]+-[A-Za-z]+$', word) or \
                re.match(r'^[a-zA-Z]+-[A-Za-z]+$', word):  # Ensure mixed-case words with hyphens are preserved
            return word
        else:
            return word.lower()

    words = re.split(r'(\W+)', text)  # Split by non-word characters while preserving them

    new_words = []
    for i, word in enumerate(words):
        # Capitalize the first and last word, and preserve other specific patterns
        if i == 0 or i == len(words) - 1 or word.lower() not in exceptions:
            new_words.append(word.capitalize() if word.islower() else word)
        else:
            new_words.append(preserve_word(word))

    return ''.join(new_words)

# Load the Excel file
file_path = 'Titles.xlsx'
file_path2 = 'Titles_updated.xlsx'

df = pd.read_excel(file_path)

# Check if 'Title' column exists
if 'Title' in df.columns:
    # Apply titlecase function to each title in the 'Title' column
    df['Title_Corrected'] = df['Title'].apply(ap_title_case)

    # Save the result back to the Excel file in Column K
    df['Column K'] = df['Title_Corrected']
    df.drop(columns=['Title_Corrected'], inplace=True)  # Clean up the temporary column

    # Save the updated DataFrame back to the Excel file
    df.to_excel(file_path2, index=False)

    print("Titles have been converted to AP Title Case and saved to Column K.")
else:
    print("The column 'Title' does not exist in the provided file.") # Add the word 'Title' in the first cell of the
                                                                     # column if you get this error

