# Excel Data Updater

This Python script is designed to update an Excel workbook with data from a CSV file. It utilizes the pandas library for data manipulation and the openpyxl library for interacting with Excel workbooks.

## Usage

1. Make sure you have Python and the required libraries installed.

2. Clone this repository or download the script:

   ```bash
   git clone https://github.com/your-username/excel-data-updater.git
   cd excel-data-updater
   ```

3. Install the required Python packages by running the following command:

   ```bash
   pip install pandas openpyxl
   ```

4. Prepare your input data CSV file (`New_V2 Product Lot & SN release date.csv`) and Excel template file (`COC_Template.xlsx`).

5. Open the script in a code editor and modify the file paths for `COC_Template.xlsx` and `New_V2 Product Lot & SN release date.csv` to match your actual file paths.

6. Run the script:

   ```bash
   python excel_data_updater.py
   ```

7. The script will read data from the CSV file and update the Excel template file with the corresponding values. It will save the modified Excel files with filenames based on the serial numbers (SN).

## Important Note

- Ensure that the input CSV file columns `'SN'` and `'DATE'` match the column names used in the script. Modify the script if your columns have different names.

## Contributing

Contributions are welcome! If you'd like to contribute to this project, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Make your changes and test thoroughly.
4. Commit your changes and push to your fork.
5. Submit a pull request to the main repository.


## Credits

- This project was developed by [Vasanttan](https://github.com/Vasant98).
- The data manipulation is powered by the pandas library.
- The Excel interaction is provided by the openpyxl library.

## Contact

If you have any questions or suggestions, please feel free to contact me.
