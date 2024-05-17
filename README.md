# CBSE Result Scrapper

CBSE Result Scrapper is a Python-based tool to automate the process of retrieving CBSE Class XII results from the official website and storing the data in an Excel file. 

## Features

- **Automated Data Retrieval**: Uses Selenium to automate the process of logging in and scraping results from the CBSE website.
- **Data Processing**: Calculates the Best 5 and Core 5 percentages from the retrieved marks.
- **Excel Integration**: Writes the processed data to an Excel file with proper formatting and alignment.

## Setup

1. **Clone the Repository**:

    ```sh
    git clone https://github.com/Vitthal-Jauhari/CBSE-Result-Scrapper.git
    cd CBSE-Result-Scrapper
    ```

2. **Install Dependencies**:

    Ensure you have Python installed. Install the required Python packages:

    ```sh
    pip install selenium openpyxl
    ```

3. **Download WebDriver**:

    Download the Chrome WebDriver matching your Chrome browser version and place it in the specified path or update the path in the code:

    ```python
    browser = Browser('path_to_chromedriver')
    ```

## Usage

1. **Prepare the Input Excel File**:

    Ensure you have an input Excel file (e.g., `Class XII A details.xlsx`) with the necessary student information in the specified columns.

2. **Run the Script**:

    Execute the main script to start the scraping and data processing:

    ```sh
    python main.py
    ```

3. **Check the Output**:

    The script will create or update an Excel file (`Marks.xlsx`) with the retrieved and processed results.

## Notes

- Ensure that the CBSE results website URL used in the script is up-to-date and accessible.
- Customize the paths and column references in the script according to your specific requirements.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request or open an Issue for any improvements or suggestions.
