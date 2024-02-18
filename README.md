# Project Name
Optimizing Workforce Allocation in XXX warehouse.
## Description
This project aims to optimize workforce allocation in an XXX warehouse based on the incoming shipment data and operational requirements. 
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
Prepare your input data in an Excel file with the required format. The input file should contain Required Date and timing range.[Download Input Sheet](https://github.com/Vishnuvarthanaadhi/Workforce/blob/5c87cc692ae603f218f57c08137f134910a7fc9f/Input.xlsx)
Update the input_file_path variable in the main() function of workforce.py with the path to your input Excel file.
Update the output_file_path variable in the main() function with the desired path for the output Excel file.[Download Output Sheet](https://github.com/Vishnuvarthanaadhi/Workforce/blob/a41ae7c4b591a5132cfe8e429946168737d69354/MainData.xlsx).

Run the script:python workforce.py

The output Excel file will be generated with workforce allocation details.

## Contributing

Contributions to this project are welcome. To contribute, follow these steps:

Fork the repository.

Create a new branch for your feature or bug fix.

Make your changes and commit them with descriptive commit messages.

Push your changes to your fork.

Submit a pull request to the main repository.

## License
This project is licensed under the MIT License.

## Code Structure

The code is organized into several functions:

read_input_data: Reads input data from an Excel file.

![Workflow](https://github.com/Vishnuvarthanaadhi/Workforce/blob/a6975e0c6c78cac52ec0962fb46b5bf98c0bc08e/Input.png)

preprocess_input_data: Preprocesses input data, including converting date and time columns.

filter_data: Filters data based on a specified time range.

calculate_workforce: Calculates workforce for a given time interval.

write_output_to_excel: Writes the calculated workforce to an output Excel file.So For the above Input, Output sheet contains the total workforce needed for the given Time Range.

![Workflow](https://github.com/Vishnuvarthanaadhi/Workforce/blob/5f3c3892291eebb2735fce5a9b52360eee10b965/Output.png)

calculate_workforce_whole_day: Calculates workforce for the whole day in hourly intervals.

![Workflow](https://github.com/Vishnuvarthanaadhi/Workforce/blob/eefe6485267c3c9ea7967b4fb2036d02dc919c77/Wholeday.png)
## Dependencies
pandas
openpyxl

## Credits
The project utilizes the pandas and openpyxl libraries for data manipulation and working with Excel files in Python.
