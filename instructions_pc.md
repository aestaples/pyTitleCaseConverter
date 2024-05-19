# Instructions for PC Users

## Step-by-Step Instructions

### 1. Open Command Prompt
1. Press `Win + R` to open the Run dialog.
2. Type `cmd` and press **Enter**.

### 2. Navigate to the Project Directory
Use the `cd` command to change to the directory where you want to save your project. For example, if you want to save it in the `Documents` folder:
```sh
cd %HOMEPATH%\Documents

# 3. Create a Virtual Environment
# Create a virtual environment named venv:

python -m venv venv

# 4. Activate the Virtual Environment
# Activate the virtual environment:

venv\Scripts\activate

# You should see (venv) at the beginning of your command prompt, indicating that the virtual environment is active.

# 5. Install Required Modules
# Install the required Python modules (pandas, titlecase, and openpyxl) using pip:

pip install pandas titlecase openpyxl

# 6. Place Your Excel File in the Same Directory
# Make sure your Titles.xlsx file is in the same directory where you saved the convert_titles.py script.

# 7. Run the Python Script
# Run the script using Python:

python convert_titles.py

# 8. Deactivate the Virtual Environment (Optional)
# Once you are done, you can deactivate the virtual environment by running:

deactivate
```

## Troubleshooting
* Ensure that you have Python 3 installed on your system. You can check this by running the following in your command prompt
```sh
  python --version 
```
  
* If you encounter permission issues, you might need to run the Command Prompt as an administrator.
* Make sure your Titles.xlsx file is correctly formatted and located in the same directory as your script.

