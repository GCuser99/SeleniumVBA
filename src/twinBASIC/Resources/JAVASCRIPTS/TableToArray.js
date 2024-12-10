function htmlTableToArray(table, skipHeader, skipFooter, createSpanData) {
    // This function recursively converts [nested] table[s] into an array [vba collection]
    
    let maxCols = 0;
    let v = [];

    // Remove the header and footer rows if necessary
    if (skipHeader) {
        table.deleteTHead();
    }
    if (skipFooter) {
        table.deleteTFoot();
    }

    // Handle row spans if createSpanData is true
    if (createSpanData) {
        for (let row of table.rows) {
            for (let cell of row.cells) {
                let rowSpan = cell.rowSpan;
                if (rowSpan > 1) {
                    for (let i = row.rowIndex + 1; i < row.rowIndex + rowSpan; i++) {
                        let targetRow = table.rows[i];
                        targetRow.insertCell(cell.cellIndex).innerText = cell.innerText;
                    }
                }
            }
        }

        // Handle column spans
        for (let row of table.rows) {
            for (let cell of row.cells) {
                let colSpan = cell.colSpan;
                if (colSpan > 1) {
                    for (let i = 1; i < colSpan; i++) {
                        row.insertCell(row.cells.length).innerText = cell.innerText;
                    }
                }
            }
        }
    }

    // Calculate the maximum number of columns needed for the array
    for (let row of table.rows) {
        if (row.cells.length > maxCols) {
            maxCols = row.cells.length;
        }
    }

    // Initialize the output array
    for (let i = 0; i < table.rows.length; i++) {
        v[i] = new Array(maxCols);
    }

    // Extract cell data from each row and store it in the array
    for (let row of table.rows) {
        for (let cell of row.cells) {
            let foundTable = false;

            // Check if the cell contains a table, process only the first table if found
            if (cell.children.length > 0) {
                for (let cellChild of cell.children) {
                    if (cellChild.tagName.toUpperCase() === 'TABLE') {
                        // We have a nested table...
                        v[row.rowIndex][cell.cellIndex] = htmlTableToArray(cellChild, skipHeader, skipFooter, createSpanData);
                        foundTable = true;
                        break;
                    }
                }
            }

            // If no table found, store the text content of the cell
            if (!foundTable) {
                v[row.rowIndex][cell.cellIndex] = cell.innerText;
            }
        }
    }

    return v;
}

// Create an empty table element
const htmlTable = document.createElement('table');
let table = arguments[0];

// Assign html of empty table element to incoming element's html
// Handle cases if incoming element's tag is table vs tbody
if (table.tagName.toUpperCase() === 'TABLE') {
    htmlTable.innerHTML = table.innerHTML;}
else if (table.tagName.toUpperCase() === 'TBODY'){
    htmlTable.innerHTML = table.outerHTML;}
else {
    return false;
}
// Convert table element to an array (vba collection)
return htmlTableToArray(htmlTable, arguments[1], arguments[2], arguments[3]);
