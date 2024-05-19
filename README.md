# Title Case Converter

This repository contains a Python script that converts titles in an Excel file to AP Title Case. The script preserves acronyms, abbreviations, compound words, hyphenated expressions, and specific patterns.

## Description

The `convert_titles.py` script reads an Excel file named `Titles.xlsx` and converts the titles in Column A (with the header 'Title') to AP Title Case. The properly capitalized titles are then saved in Column K. The script handles special cases to preserve:
- Acronyms (e.g., 'NASA')
- Abbreviations with digits (e.g., '3D', '2D')
- Compound words separated by slashes or hyphens (e.g., 'YAP/TAZ', 'Wnt-11')
- Greek letters (e.g., 'μ-Fluidic')
- Specific patterns with parentheses or colons followed by uppercase letters or digits (e.g., '(IL-1β)', 'Title: Subtitle')

## Files

- `convert_titles.py`: The Python script to perform the title case conversion.
- `instructions_mac.md`: Instructions for Mac users to set up and run the script.
- `instructions_pc.md`: Instructions for PC users to set up and run the script.

## Requirements

- Python 3.x
- `pandas`, `titlecase`, `openpyxl` modules

## Usage

### For Mac Users

Follow the instructions in [instructions_mac.md](instructions_mac.md) to set up and run the script on a Mac.

### For PC Users

Follow the instructions in [instructions_pc.md](instructions_pc.md) to set up and run the script on a PC.

## How to Run the Script

1. Ensure that `Titles.xlsx` is placed in the same directory as `convert_titles.py`.
2. Run the script by executing `python convert_titles.py` in your terminal or command prompt.
3. The script will read the titles from Column A of `Titles.xlsx`, convert them to AP Title Case, and save the results in Column D.

## License

This project is licensed under the MIT License.
