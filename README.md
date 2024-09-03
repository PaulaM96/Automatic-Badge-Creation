# Automatic-Badge-Creation
This script automates the creation of badges based on an Excel spreadsheet and a predefined layout. It reads the input data, generates QR codes, and customizes PowerPoint slides to create ready-to-print badges.

## Table of Contents

- [Objective](#objective)
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Objective

The main goal of this project is to streamline the process of creating badges by automatically reading data from a spreadsheet and using it to generate PowerPoint slides that follow a specified badge layout. The script also generates QR codes for each badge.

## Features

- Automatically copies the required Excel files from a source folder.
- Reads and maps necessary columns from the Excel files.
- Removes filters from the spreadsheets to avoid data errors.
- Generates QR codes based on employee IDs.
- Customizes a PowerPoint template to create badges with employee data.
- Saves the generated badges in a designated output folder.
- Provides a user-friendly interface with a start button.

## Prerequisites

Ensure you have the following software and Python libraries installed:

- **Python 3.x**
- **pip** (Python package installer)

Required Python packages:
- `tkinter` (usually included with Python)
- `openpyxl`
- `pandas`
- `python-pptx`
- `qrcode[pil]`

To install the required packages, run:
pip install openpyxl pandas python-pptx qrcode[pil]

## Installation
Clone the repository or download the script files.

Make sure the Python environment is set up with the necessary dependencies.

Adjust the file paths in the script to point to your layout templates and input/output directories.

**Usage**

- Ensure that the Excel files with employee data are placed in the source directory, which is dynamically generated based on the current year and month.

- Modify the ppt_directory path to point to the folder containing the PowerPoint badge template (PADRAO CRACHA.pptx).

- Run the script using Python or execute the generated executable (if compiled).

- Click the INICIAR CRACHÁS button in the graphical interface to start the process.

- The generated badges will be saved in the output directory specified in the script.

## Important Paths in the Script:
Source Directory: \\caminho\{year}\{month}-{year}

Template Directory: \\caminho\CRACHAS

Destination Directory: \\caminho\CRACHAS\CRACHÁS FEITOS\{year}\{month}-{year}\CRACHAS AUTOMATICOS

## Troubleshooting

- Ensure all paths are accessible and correct, especially when working in network environments.

- Make sure the PowerPoint template (PADRAO CRACHA.pptx) is correctly formatted and located in the specified directory.

- If encountering missing dependencies, re-run the installation command for the required packages.

## Contributing
Feel free to open issues or submit pull requests to improve the script.
