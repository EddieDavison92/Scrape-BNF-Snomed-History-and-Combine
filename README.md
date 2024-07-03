# BNF Snomed Mapping Data Scraper and Combiner

This project scrapes BNF Snomed mapping data from the NHS BSA website, converts the data from .xlsx to .csv format, combines the data into a single .csv file, and then converts the combined .csv file to .xlsx format.

## Project Structure

- `main.py`: The main script that performs the scraping, conversion, and combination tasks.
- `convert_to_csv.vbs`: A VBScript that handles the conversion of .xlsx files to .csv format.
- `bnf_snomed_mapping_data/`: Generated directory structure containing downloaded, processed, and output files.
  - `zip_files/`: Contains the downloaded .zip files.
  - `xlsx_files/`: Contains the extracted .xlsx files.
  - `csv_files/`: Contains the converted .csv files.
  - `latest/`: Contains the latest .xlsx file based on the date in the filename.
  - `output/`: Contains the final combined .csv and .xlsx files.
- `requirements.txt`: Lists the dependencies required to run the script.

## Prerequisites

- Python 3.6 or higher
- `pip` (Python package installer)

## Installation

1. Clone the repository:

    ```sh
    git clone (https://github.com/EddieDavison92/Scrape-BNF-Snomed-History-and-Combine/)
    ```

2. Set up a virtual environment (optional but recommended):

    ```sh
    python -m venv venv
    venv\Scripts\activate
    ```

3. Install the required packages:

    ```sh
    pip install -r requirements.txt
    ```

## Usage

1. Ensure you have `convert_to_csv.vbs` in the same directory as `main.py`.

2. Run the script:

    ```sh
    python main.py
    ```

The script will:
- Navigate to the NHS BSA website and find all .zip files containing BNF Snomed mapping data.
- Download and extract the .zip files.
- Convert the .xlsx files to .csv format using the VBScript.
- Combine the data from all .csv files into a single DataFrame, ensuring no duplicate rows.
- Save the combined data to a single .csv file in the `output` directory.
- Convert the combined .csv file to .xlsx format, ensuring that the `SNOMED Code` column is formatted as text, and set the table style to `LightStyle8`.

## Logging

The script provides detailed logging information, which is helpful for monitoring the progress and troubleshooting any issues.

## Output

The final combined files will be saved in the `bnf_snomed_mapping_data/output/` directory:
- `combined_bnf_snomed_mapping_data.csv`
- `combined_bnf_snomed_mapping_data.xlsx`

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
