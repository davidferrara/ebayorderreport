# ebayorderreport
This is a custom Google Apps Script I created for a google sheet.

The purpose of this Google Sheet is to be able to import an Active Order Report from an ebay seller account of all the active orders, add additional data to the sheet, spit the imported sheet into separate sheets categorized by the "Ship By Date" and prepare it for printing.  This data is used to pick items from a warehouse to be shipped out.

To use: Import the ebay order report by replacing the "Imported Sheet" with the downloaded csv file.  In the custom Menu there are two options, "Process Impoted Sheet", and "Delete Sheets".

<b>Process Imported Sheet:</b>
First this deletes all the unnecessary columns and rows of data.  Next it sorts the "Imported Sheet" by the "Ship By Date" column.  A new column titled "Warehouse Location" is added and formatted to the sheet.  This column is then populated with data from the addLocation() function.  This function crossreferences two different sheets looking for matching SKUs from other tracking spreadsheets in order to find the location that cooresponds to the SKU on the imported sheet.  Next the separateRanges() creates an array of ranges separated by "Ship By Date" column.  New sheets are then created with each sheet cooresponding to the ship by date and the ranges are copied over to the new sheets.
