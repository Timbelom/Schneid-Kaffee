function processSheet() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getSheetByName('Form responses 1');
    const data = sourceSheet.getDataRange().getValues();
  
    // Start from the second row, assuming the first row is headers
    for (let i = 1; i < data.length; i++) {
      let row = data[i];
  
      // Filter out empty cells and specific text
      const filteredRow = row.filter(cell => cell !== '' && cell !== 'Einen weiteren Artikel hinzufügen' && cell !== 'Abschicken');
      
      if (filteredRow.length > 0) {
        // Extract and format the date from the first cell
        const date = new Date(filteredRow[0]);
        const formattedDate = Utilities.formatDate(date, spreadsheet.getSpreadsheetTimeZone(), 'dd/MM/yyyy');
  
        // Create a new sheet for each processed row with the formatted date as its name
        const newSheetName = formattedDate ? formattedDate : 'Invalid Date ' + i;
        const newSheet = spreadsheet.insertSheet(newSheetName);
        const firstCellContent = filteredRow.shift(); // Get the first cell content
        const finalRows = [];
  
        // Subdivide the row into groups of 4 cells and concatenate B to D
        while (filteredRow.length > 0) {
          let smallRow = filteredRow.splice(0, 4);
          let concatenated = smallRow.slice(1, 3).join(' '); // Joining values from B to D
          finalRows.push([smallRow[0], concatenated,smallRow[3]]); // Keep the first cell and concatenated cell
        }
  
        // Write the data to the new sheet
        newSheet.getRange(2, 1, finalRows.length, 3).setValues(finalRows); // Write final rows
  
        // Sort the rows starting from the second row
        newSheet.getRange(2, 1, finalRows.length, 2).sort({column: 1, ascending: true});
  
        // Find the row with data only in the first column after sorting
        let lengthyText = '';
        for (let j = 2; j <= finalRows.length + 1; j++) {
          let rowRange = newSheet.getRange(j, 1, 1, 2);
          let rowData = rowRange.getValues()[0];
          if (rowData[1] === '') {
            // Store the lengthy text and clear the original cell
            lengthyText = rowData[0];
            newSheet.getRange(j, 1).clearContent();
            newSheet.deleteRow(j);
            break;
          }
        }
  
          // Auto resize columns and then add additional spaces for readability
        [1, 2].forEach(column => {
          newSheet.autoResizeColumn(column);
          let newWidth = newSheet.getColumnWidth(column) + 20; // Add 20 pixels for padding
          newSheet.setColumnWidth(column, newWidth);
        });
  
        // Insert two new rows after the first row
        newSheet.insertRowAfter(1);
        newSheet.insertRowAfter(2);
  
        // Place the lengthy text in the new second row if it exists
        if (lengthyText !== '') {
          newSheet.getRange(2, 1).setValue(lengthyText);
          // Merge cells A2:C3
          newSheet.getRange('A2:D3').merge();
        }
  
        // Move original A1 content to C1 and add specific text to A1
        newSheet.getRange('D1').setValue(firstCellContent);
        newSheet.getRange('A1').setValue("TestCompany \n test street 1 \n +12345678");
  
        // Insert a new row at the top and merge cells A1:C1
        newSheet.insertRowBefore(1);
        newSheet.getRange('A1:D1').merge();
        // Insert "Bestellung" into B1, center the text, and set font properties
        let headerCell = newSheet.getRange('A1');
        headerCell.setValue("Bestellung");
        headerCell.setHorizontalAlignment("center");
        headerCell.setFontWeight("bold");
        headerCell.setFontSize(20);
  
        var lastRow = newSheet.getLastRow();
  
        newSheet.getRange(lastRow + 2, 1).setValue("Im Packtisch  [ ]");
        newSheet.getRange(lastRow + 2, 3).setValue("Bearbeitet von __________");
        newSheet.getRange(lastRow + 4, 1).setValue("Lieferung per __________");
        newSheet.getRange(lastRow + 4, 3).setValue("Abholung am __________");
        newSheet.getRange(lastRow + 6, 1).setValue("Bestellannhme:__________");
        newSheet.getRange(lastRow + 6, 3).setValue("Erledigt __________");
      }
    }
  }
  