# Project Name
Optimizing Workforce Allocation in Amazon warehouse.
## Description
This project aims to optimize workforce allocation in an Amazon warehouse based on the incoming shipment data and operational requirements. 
It calculates the required number of unloaders, injectors, facers, and dumper operators for a specified time range and for the whole day. 
The output is generated in an Excel file with detailed workforce allocation information.

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Installation
Clone the repository to your local machine: git clone https://github.com/Vishnuvarthanaadhi/Workforce.git

Navigate to the project directory:  cd Workforce

Install the required dependencies:  pip install pandas openpyxl
## Usage
Prepare your input data in an Excel file with the required format. The input file should contain Required Date and timing range.[Download Input Sheet](path/to/input_sheet.xlsx)
Update the input_file_path variable in the main() function of workforce.py with the path to your input Excel file.
Update the output_file_path variable in the main() function with the desired path for the output Excel file.[Download Output Sheet](path/to/output_sheet.xlsx)

Run the script:python workforce.py

The output Excel file will be generated with workforce allocation details.

## Contributing

Contributions are welcome! If you encounter any bugs, have suggestions for improvements, or want to contribute new features, feel free to open an issue or submit a pull request.

## License
This project is licensed under the MIT License.

## Code Structure

workforce.py: Main Python script containing functions for reading input data, preprocessing, filtering data, calculating workforce, writing output to Excel, and calculating workforce for the whole day.

Input.xlsx: Sample input Excel file containing Date and Timing range,Which is required to calculate.

MainData.xlsx: Sample output Excel file with workforce allocation details for a specified time range.

README.md: Project documentation.

## Dependencies
pandas
openpyxl

## Credits
The project utilizes the pandas and openpyxl libraries for data manipulation and working with Excel files in Python.
