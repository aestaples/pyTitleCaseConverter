# Instructions for Mac Users

## Step-by-Step Instructions

### 1. Open Terminal
1. Click on the **Spotlight** (magnifying glass) in the top right corner of your screen.
2. Type **Terminal** and press **Enter**.

### 2. Navigate to the Project Directory
Use the `cd` command to change to the directory where you want to save your project. For example, if you want to save it in the `Documents` folder:
```sh
cd ~/Documents

# 3. Create a Virtual Environment
# Create a virtual environment named venv:

python3 -m venv venv

# 4. Activate the Virtual Environment

source venv/bin/activate

# You should see (venv) at the beginning of your terminal prompt, indicating that the virtual environment is active.

# 5. Install Required Modules
# Install the required Python modules (pandas, titlecase, and openpyxl) using pip:

pip install pandas titlecase openpyxl

# 6. Place Your Excel File in the Same Directory
# Make sure your Titles.xlsx file is in the same directory where the convert_titles.py script is located.

# 7. Run the Python Script
# Run the script using Python:

python convert_titles.py

# This will execute your script, convert the titles in your Excel file according to the custom AP title case rules, and save the changes in Column K.

# 8. Deactivate the Virtual Environment (Optional)
# Once you are done, you can deactivate the virtual environment by running:

deactivate
```

## Troubleshooting

- Ensure that you have Python 3 installed on your system. You can check this by running

  `python3 --version`
  
  in your terminal.
- If you encounter permission issues, you might need to prepend some commands with sudo\, especially when installing packages.
- Make sure your Titles.xlsx file is correctly formatted and located in the same directory as your script.



