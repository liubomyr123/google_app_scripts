# Google Sheets Normal Distribution Calculator

This script provides a custom menu and functionality in Google Sheets to calculate the Normal Distribution using the `NORM.DIST` function. It allows users to quickly calculate the normal distribution for given values of `x`, mean, standard deviation, and cumulative distribution.

## Features
- Custom menu for quick access.
- Supports vertical and horizontal cell ranges.
- Validates input values for correct data types.
- Calculates normal distribution and displays the result in the selected cell.

## How to Use
1. Open your Google Sheets document.
2. Click on the **"Каразіна"** menu that appears after the script is added.
3. Select **"Use normal distribution"** from the dropdown menu.
4. Highlight the necessary cells with values:
   - **x**: The value to calculate the distribution for.
   - **mean**: The mean of the distribution.
   - **stdDev**: The standard deviation of the distribution.
   - **cumulative**: Whether to calculate the cumulative distribution (TRUE/FALSE).
5. The result will appear in the selected cell.

## Requirements
- Google Sheets (obviously).
- The script should be added via the Apps Script editor.

## Example Use Case
For example, if you want to calculate the cumulative normal distribution for `x = 10`, `mean = 5`, `stdDev = 2`, and set the cumulative value to TRUE, the script will automatically populate the formula `=NORM.DIST(10, 5, 2, TRUE)` in the selected cell.
