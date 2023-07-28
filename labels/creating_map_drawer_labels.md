**Note:** This specific workflow assumes the initial data comes in the form of a file with four fields: 'call_number', 'description', 'sheet_range', and 'scale'. The workflow (primarily the script) will have to be modified to accommodate data in other formats.

### Reformatting the data

Since the preferred format for the physical label is to have three columns, the 'sheet_range' and 'scale' fields will need to be combined into one. To accomplish this, run the 'excel_spreadsheet_to_labels.ipynb' Jupyter Notebook script. This script will combine these two fields, joining their values with a newline feed. It also adds 'spacer' rows and columns to the data to assist with additional formatting in the output Excel file. The script must be run in the same directory as the initial map-drawer-labels spreadsheet. The reformatted output will be sent to this same directory.

### Sizing label components

<!-- The dimensions of map-drawer labels are 5.25” x 1.75” or 133mm x 45mm. To change the measurement unit in Excel 2016, go to File > Options > Advanced. In the 'Display' section, select 'Millimeters' in the 'Ruler units' dropdown.

**NOTE:** Some of the below steps will be accomplished through running the script. However, they may be relevant in scenarios where the script is not used. -->

After running the script, open the new file in Excel. Ensure column widths are:

| Column | Width |
| ------ | ----- |
| A      | 1     |
| B      | 19    |
| C      | 20    |
| D      | 19    |
| E      | 2     |
| F      | 2     |
| G      | 19    |
| H      | 20    |
| I      | 19    |
| J      | 2     |

Ensure row heights are:

| Row | Height |
| --- | ------ |
| 1   | 5      |
| 2   | 70     |
| 3   | 10     |
| 4   | 10     |
| 5   | 70     |

<!-- **Steps - columns:**

- Select View > Page Layout to view the spreadsheet in print format
- Select Page Layout > Size > 11 x 17
- Select Page Layout > Margins, and set margins to 0
- Select columns B, D, and F, and wrap the text
  - Increase the font size to 14
  - Align the text to the middle, vertically
  - Align the text to the middle, horizontally
- Select columns A, C, and E, and set the column width to 5 (Column A is set to 5, rather than 10, to account for the additional space added when printing, since margin settings of 0 still result in some physical margin being present.)
- Select columns B and F, and set the column width to 30
- Select column D, and set the column width to 43
- Select column G, and set the column width to 10

**Steps - rows:**

- Select the first row (the 'header_spacer' row), and set its height to 5 (Row 1 is set to 5, rather than 10, to account for the additional space added when printing, since margin settings of 0 still result in some physical margin being present.)
- Place a filter on column H, and filter by 'text'
  - Select all 'text' rows, and set the row height to 25
- Filter column H by 'horizontal_spacer_top' and 'horizontal_spacer_bottom'
  - Select all filtered rows, and set the row height to 10 -->

<!-- ### Insert borders to help with cutting labels

- Select column G > Format Cells... > Border, and apply a line style of your preference to the right edge
- Filter column H by 'horizontal_spacer_bottom'
- Select all rows where column H equals 'horizontal_spacer_bottom' and columns A-G
- Select Format Cells... > Border, and apply a line style of your preference to the bottom edge (Leave the exisitng border from a previous step on the right edge.)
- For any label that is cutoff at the bottom of a page, click on the top-left cell of the label
  - Select Page Layout > Breaks > Insert Page Break -->

Apply to the call numbers **Bold** formatting.

Review the cell content to make sure it fits within the cells.

The labels should now be ready for printing. Printing should be done on 11" x 17" card stock.
