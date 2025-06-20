# AFAD Excel Helper: 
### Automating Fraud Detection for Property Restoration Applications
Excel Helper is a Python-based application designed to simplify the management of Excel sheets and automate the process of identifying inconsistencies in data related to property restoration applications. This tool is specifically aimed at cross-checking IDs in tables to detect fraud in cases where applicants are falsely claiming property ownership or making multiple applications for the same property slot.

## Features
1. Reformat Excel Sheets

2. Clean up and reformat existing Excel files to a consistent structure that can be more easily analyzed and processed.

3. Cross-Check IDs for Fraud Detection

4. Cross-reference IDs within the application to identify potential fraud, such as duplicate applications for the same property slot or misclassified property ownership (e.g., landowners incorrectly claiming property rights).

## Purpose
The main goal of Excel Helper is to automate the fraud detection process for applications seeking to restore property after an earthquake. By automating these checks, officials can identify inconsistencies more efficiently and ensure that the restoration process is fair and accurate.
## **Installation Process**
Follow these steps to set up the application:

### **1. Install Python 3.13+**
Ensure that Python 3.13 or later is installed on your system. If you don't have Python installed, download it from the official Python website:
- [Download Python](https://www.python.org/downloads/)

### **2. Set up a Virtual Environment (Optional but Recommended)**

We highly recommend using a virtual environment to manage project dependencies.

1. Open your Command Prompt or PowerShell.
2. Navigate to the project directory.
3. Run the following command to create a virtual environment:

   ```bash
   py -m venv myenv
   ```

4. Run the Application
Once the dependencies are installed, you can start the application. To run the main script:

```bash
pip install -r requirements.txt
python main.py
```
The application will load and start an interactive loop where you can choose to either resize cells in Excel files or check for inconsistencies in the data.

## Usage
The Excel Helper app operates in a simple loop, with the following options:

Resize Cells in Excel Sheets
- Resize the cells to a defined size for better readability and consistency.

Cross-Check IDs for Inconsistencies

- Select an Excel file that contains application data.
- The app will cross-reference the IDs across the sheet and catch inconsistencies, such as duplicate applications for the same property slot or mismatched data indicating fraudulent claims.

To stop the program at any time, simply enter son.


Folder Structure
The project folder has the following structure:

```bash
/excel-helper
  /build         # Build files (for packaging the application)
  /dist          # Distribution folder (contains packaged .exe file)
  main.py        # Main application script
  reader.py      # Script to read and process Excel files
  requirements.txt  # List of Python dependencies
```

# Contributing
If you'd like to contribute to the development of Excel Helper, feel free to fork the repository and submit a pull request. Any bug fixes, features, or improvements are welcome!

# License
This project is licensed under the MIT License - see the LICENSE file for details.

# Contact
For any questions or issues related to the project, feel free to reach out to the project maintainers. You can also raise issues or submit feature requests through the repository’s Issues tab on GitHub.

# Additional Notes
Compatibility: The application should work on most recent versions of Windows. However, it was developed and tested using Python 3.13 on Windows 10.

Data Formatting: Ensure that your Excel files follow a consistent format, as the program might not handle inconsistencies in sheet structure or missing data gracefully.
