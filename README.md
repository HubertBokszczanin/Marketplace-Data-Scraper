# Marketplace Data Scraper

This Python script is designed to scrape product data from two different marketplaces: Kaufland and eBay, based on the provided EAN (European Article Number) codes. It retrieves product information such as prices, titles, delivery times, reviews, and more, and saves the data to an Excel file for further analysis.

## Prerequisites
Before using this script, make sure you have the following installed:
- Python 3.x
- Required Python packages (you can install them using pip):
  - requests
  - bs4 (Beautiful Soup)
  - openpyxl
  - tkinter (for file selection dialog)
  - tqdm (for progress tracking)

## How to Use
Follow these steps to use the script:

1. Clone or download the script to your local machine.

2. Open a terminal or command prompt and navigate to the directory containing the script.

3. Run the script by executing the following command:

4. You will be prompted to select a marketplace:
- Type `1` for Kaufland
- Type `2` for eBay

5. Depending on your choice, follow the additional prompts to select the country (for eBay) and the input Excel file containing the EAN codes to scrape.

6. The script will begin processing the EAN codes and scraping data. It will display a progress bar to track the progress.

7. Once the scraping is complete, the data will be saved to an Excel file with a unique name in the same directory as the script.

8. You can find the scraped data in the Excel file, which will include columns such as EAN, URL, Price, Title, Delivery Time, Reviews, and more.

## Output Files
The script will create one or more Excel files, depending on the number of sets processed for each EAN:
- For Kaufland: The script creates separate files for each set of data processed, with filenames like `Kaufland_scraped_data_1.xlsx`, `Kaufland_scraped_data_2.xlsx`, and so on.
- For eBay: The script creates separate files for each set of data processed, including the country code in the filename, e.g., `Ebay.de_scraped_data_1.xlsx`, `Ebay.fr_scraped_data_1.xlsx`, and so on.

## Important Notes
- The script prevents processing more than two sets of data for the same EAN to avoid excessive scraping.
- The "Data Analizy" column in the output file contains a timestamp indicating when the data was scraped.

## Troubleshooting
- If you encounter issues or errors while running the script, ensure that you have the required Python packages installed and that your internet connection is stable.
- Make sure the input Excel file follows the format specified in the script.

## Disclaimer
This script is provided for educational and informational purposes. It is your responsibility to ensure compliance with the terms of service and policies of the respective marketplaces when using this script for scraping data.

**Happy scraping!**
