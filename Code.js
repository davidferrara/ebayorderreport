function onOpen() {
  const ui = SpreadsheetApp.getUi()
  
  ui.createMenu('Menu')
    .addItem('Process Imported Sheet', 'main')
    .addSeparator()
    .addItem('Delete Sheets', 'removeSheets')
    .addToUi();
}

function main() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadSheet.getSheets() //can be truncated with: const importedSheet = spreadSheet.getSheets()[0]
  const importedSheet = sheets[0]
  const template = spreadSheet.getSheetByName('Template')

  // Delete the extra rows and columns.
  deleteRowsCols(importedSheet)

  // Sort the sheet by date.
  const sortRange = importedSheet.getRange(2, 1 ,importedSheet.getLastRow() - 1, importedSheet.getLastColumn())
  sortRange.sort({column: 6, ascending: true})

  // Insert the Warehouse Location column.
  importedSheet.insertColumnAfter(4);
  importedSheet.getRange(1, 5).setValue('Warehouse Location')
  importedSheet.getRange(1, 1, importedSheet.getMaxRows(), importedSheet.getMaxColumns()).setNumberFormat("@")

  // Find the Warehouse Locations and add them.
  addLocation(importedSheet)

  let ranges = separateRanges(importedSheet)
  createSheets(template, importedSheet, spreadSheet, ranges)
}

function deleteRowsCols(importedSheet) {
  importedSheet.deleteColumns(52, 26)
  importedSheet.deleteColumns(28, 23)
  importedSheet.deleteColumn(26)
  importedSheet.deleteColumns(3, 20)
  importedSheet.deleteColumn(1)

  importedSheet.deleteRows(importedSheet.getMaxRows() - 4, 5)
  importedSheet.deleteRow(3)
  importedSheet.deleteRow(1)
}

function addLocation(importedSheet) {
  const ebaySheet = SpreadsheetApp.openById('1y_NlGxgcUBM2u_RiuouIkpwVzki4x1830ykEZUJjgjg').getSheetByName('Tracking');
  const archiveSheet = SpreadsheetApp.openById('1iR3WcwUame5vJpEDjNM2013QHpcD6V_XxXa80-uQnlA').getSheetByName('Archive Tracking');

  // Get the SKU ranges from the imported order sheet, the ebay tracking sheet, and the ebay archive sheet.
  let ebaySKURange = ebaySheet.getRange(3, 11, ebaySheet.getLastRow() - 2);
  let archiveSKURange = archiveSheet.getRange(2, 11, archiveSheet.getLastRow() - 1);
  let orderSKURange = importedSheet.getRange(2, 4, importedSheet.getLastRow() - 1);

  let ebaySKUValues = ebaySKURange.getValues();
  let arhciveSKUValues = archiveSKURange.getValues();
  let orderSKUValues = orderSKURange.getValues();

  // Get the Location ranges from the imported order sheet, the ebay tracking sheet, and the ebay archive sheet.
  let ebayLocationRange = ebaySheet.getRange(3, 9, ebaySheet.getLastRow() - 2);
  let archiveLocationRange = archiveSheet.getRange(2, 9, archiveSheet.getLastRow() - 1);
  let orderLocationRange = importedSheet.getRange(2, 5, importedSheet.getLastRow() - 1);

  let ebayLocationValues = ebayLocationRange.getValues();
  let arhciveLocationValues = archiveLocationRange.getValues();
  let orderLocationValues = orderLocationRange.getValues();

  // Iterate through the order sheet skus looking for matches on other sheets.
  let index, index2;
  for(let i = 0; i < orderSKUValues.length; i++) {
    if(orderSKUValues[i][0] != "") {
      index = ebaySKUValues.findIndex((x) => x == orderSKUValues[i][0]);
      if(index != -1) {
        orderLocationValues[i][0] = ebayLocationValues[index][0];
      } else {
        index2 = arhciveSKUValues.findIndex((x) => x == orderSKUValues[i][0]);
        if(index2 != -1) {
          orderLocationValues[i][0] = arhciveLocationValues[index2][0];
        }
      }
    }
  }
  
  orderLocationRange.setValues(orderLocationValues);
}

function separateRanges(importedSheet) {
  let result = []
  let dateRange = importedSheet.getRange(2, 7, importedSheet.getLastRow() - 1) // Get the date range.
  // dateRange.setNumberFormat("@") // Format the date range to plain text.
  let dateRangeValues = dateRange.getValues() // Get the date range values.

  // Extract the dates from the importSheet.
  let dates = []
  for(let i = 0; i < dateRangeValues.length; i++) {
    if(i == 0) {
      dates.push(dateRangeValues[i][0])
    } else {
      if((dateRangeValues[i][0] != dateRangeValues[i-1][0])&&(dateRangeValues[i][0] != '')) {
        dates.push(dateRangeValues[i][0])
      }
    }
  }

  // Extract the date ranges based on the extracted dates.
  let ranges = []
  let startRow, endRow, textFinder, finderRanges
  for(let j = 0; j < dates.length; j++) {
    textFinder = importedSheet.createTextFinder(dates[j])
    finderRanges = textFinder.findAll()
    startRow = finderRanges.shift().getRow()
    endRow = finderRanges.pop().getRow()
    ranges.push(importedSheet.getRange(startRow, 1, endRow - startRow + 1, 7))
  }

  for(let k = 0; k < dates.length; k++) {
    result.push([dates[k], ranges[k]])
    Logger.log(result[k][0].toString() + ' : ' + result[k][1].getA1Notation().toString())
  }

  return result
}

function createSheets(template, importedSheet, spreadSheet, ranges) {
  let newSheet, newDataRange, data
  for(let i = 0; i < ranges.length; i++) {
    data = importedSheet.getRange(ranges[i][1].getA1Notation().toString()).getValues()

    newSheet = template.copyTo(spreadSheet).setName(ranges[i][0])
    newDataRange = newSheet.getRange(2, 1, data.length, 7).setValues(data)

    newDataRange.sort({column: 5, ascending: true})

    if(i == 0) {
      spreadSheet.setActiveSheet(newSheet)
      spreadSheet.moveActiveSheet(2)
    }
    newSheet.setFrozenRows(1)
    newSheet.showSheet()
  }
}

function removeSheets() {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadSheet.getSheets()

  for(let i = 1; i < sheets.length; i++) {
    if(sheets[i].getName() != 'Template') {
      spreadSheet.deleteSheet(sheets[i])
    }
  }
}


























