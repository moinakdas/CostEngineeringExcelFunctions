# Cost Engineering Library

## Overview

The Cost Engineering Library is a powerful Excel-based tool developed during my time at TC Electric to facilitate cost engineering tasks. It provides various functionalities to analyze cost data from contract S-48012 and perform essential calculations, such as summing up money spent on cost codes, calculating procurement and delivery percentages, and filtering data based on specific criteria.

## Functions

| Function Name                     | Description                                                                                      |
|-----------------------------------|--------------------------------------------------------------------------------------------------|
| `FindTotalDelta(SHEET)` | Calculates the total money spent on a cost code up until the calling cell of the specified sheet. Returns the calculated total. |
| `TotalSpent(COST_CODE)`   | Iterates through the specified worksheets, calculating the total sum of money spent on the specified cost code. Returns the total sum. |
| `ProcurementPercentage(SHEET)` | Calculates the procurement percentage on a specific worksheet. Calculates the percentage of non-empty cells in column N (starting from row 3) and returns the result. |
| `DeliveryPercentage(SHEET)` | Calculates the procurement percentage on a specific worksheet. Returns percentage of procured items over total items |
| `BECalc(BETYPE)`         | Calculates the total cost from the specified disadvantaged business type (BETYPE). Finds the columns corresponding to "Req #," "Vendor/Cert," and "Total Cost" and iterates through the rows, searching for the BETYPE Vendor/Cert column. If a match is found, it adds the corresponding line item cost to the `TotalM` variable, which is then returned.

## Usage

To use the Cost Engineering Library, you can import the VBA code provided in the `CostFunctions.vba` file into your Excel project. Once imported, you can call any of the functions mentioned above within your Excel worksheets or macros to perform various cost-related calculations.

Please note that this library was specifically designed for analyzing cost data from contract S-48012 at TC Electric. You may need to customize the library to fit the data structure and requirements of your specific projects.

## License

This library is provided under the [GNU Affero General Public License](https://www.gnu.org/licenses/agpl-3.0.en.html), allowing you to freely use, modify, and distribute the code. However, please note that it comes with no warranty, and you should use it at your own risk.

## Contributions

Contributions to this project are welcome! If you find any issues or have suggestions for improvements, feel free to submit a pull request.

## Contact

For any questions or inquiries related to this library, you can reach me at [Moinak.Das@Stonybrook.com](mailto:Moinak.Das@Stonybrook.edu).
