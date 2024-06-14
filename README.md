# LandAndMorgageRegisterAutomation

This repository contains a script that checks the status of land and morgage registers by visiting the appropriate website and determining whether the register contains certain data. The results are saved in an Excel file.

## Files in the Repository

- `input_file.xlsx`: Excel file containing the register numbers to be checked.
- `output_file.xlsx`: Excel file where the results will be saved.
- `real_estate_checker.py`: Python script that performs the checking process.

## Prerequisites

Ensure you have the following installed:

- Python 3.x
- Pandas library
- Playwright library

You can install the required libraries using pip:

``
pip install pandas playwright
``

Additionally, you need to install the Playwright browsers:

``
python -m playwright install
``

## Usage
Place your input Excel file (input_file.xlsx) in the repository directory. Ensure it contains a column named RegisterNumber with the register numbers to be checked.

Run the script:

``
python real_estate_checker.py
``

The script will generate an output file named output_file.xlsx with the results.

## Script Details
The script performs the following steps:

1. Reads the register numbers from input_file.xlsx.
2. Launches a Playwright browser to navigate to the real estate register website.
3. For each register number, it searches for the register and checks for specific data fields.
4. Saves the results (YES/NO/ERROR) to output_file.xlsx.

## Example
Input File (input_file.xlsx)  
RegisterNumber  
BB1Z/00042835/6  
BB1Z/00042836/3  

Output File (output_file.xlsx)  
RegisterNumber	Status  
BB1Z/00042835/6  NO  
BB1Z/00042836/3  YES  

## License
This project is licensed under the MIT License. See the LICENSE file for details.

## Contact
For any questions or issues, please open an issue in the repository.
