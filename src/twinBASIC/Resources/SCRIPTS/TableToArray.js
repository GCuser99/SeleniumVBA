function htmlTableToArray(table, skipHeader, skipFooter, createSpanData, ignoreCellFormatting) {
    // This function recursively converts [nested] table[s] into an array [vba collection]
        
    let maxCols = 0;
    // v is output array
    let v = [];
    // vrs is array to keep track of row span count if createSpanData = true
    let vrs = [];

    // Calculate the maximum number of columns needed for the arrays
    for (let row of table.rows) {
        if (row.cells.length > maxCols) {
            maxCols = row.cells.length;
        }
    }

    // Initialize the v and vrs arrays
    for (let i = 0; i < table.rows.length; i++) {
        v[i] = new Array(maxCols);
        if (createSpanData) {vrs[i] = new Array(maxCols).fill(1);}
    }

    // Extract cell data from each row and store it in the array
    for (let row of table.rows) {
        let colIdx = 0;
        for (let cell of row.cells) {
            let foundTable = false;

            // Check if the cell contains a table, process only the first table if found
            if (cell.children.length > 0) {
                for (let cellChild of cell.children) {
                    if (cellChild.tagName.toUpperCase() === 'TABLE') {
                        // We have a nested table...
                        v[row.rowIndex][colIdx] = htmlTableToArray(cellChild, skipHeader, skipFooter, createSpanData, ignoreCellFormatting);
                        foundTable = true;
                        break;
                    }
                }
            }

            // If no table found, store the text of the cell
            if (!foundTable) {
                // Store the contents of the cell in output array
                if (ignoreCellFormatting) {
                    // Store the complete text content, including hidden text
                    v[row.rowIndex][colIdx] = cell.textContent.replace(/\xA0/g,' ');}
                else {
                    // Store the visible text content, including <br>'s and other white space formatting
                    v[row.rowIndex][colIdx] = cell.innerText.replace(/\xA0/g,' ');
                }
                
                // Handle col spans if createSpanData is true
                if (createSpanData) {
                    // Store row span data for use later
                    vrs[row.rowIndex][colIdx] = cell.rowSpan;
                    let colSpan = cell.colSpan;
                    if (colSpan > 1) {
                        // Propogate column span data
                        for (let i = 1; i < colSpan; i++) {
                            v[row.rowIndex][colIdx + i] = v[row.rowIndex][colIdx];
                        }
                        colIdx += colSpan - 1;
                    }
                }
            }
            colIdx += 1;
        }
    }

    // Handle row spans if createSpanData is true
    if (createSpanData) {
        // Propogate row span data if needed
        for (let i = 0; i < v.length; i++) {
            for (let j = 0; j < v[i].length; j++) {
                let rowSpan = vrs[i][j];
                if (rowSpan > 1) {
                    for (let k = i + 1; k < i + rowSpan; k++) {
                        // Insert a copy of cell data from row with span
                        v[k].splice(j, 0, v[i][j]);
                        v[k].pop();
                        vrs[k].splice(j, 0, 1);
                        vrs[k].pop();
                    }
                }
            }
        }
    }

    // Remove the header and footer rows if called for
    if (skipHeader) {
        if (table.querySelector('thead') !== null) {
            v.shift();
        }
    }
    if (skipFooter) {
        if (table.querySelector('tfoot') !== null) {
            v.pop();
        }
    }
    return v;
}

const elem = arguments[0];

// Handle case if incoming element's tag is tbody
if (elem.tagName.toUpperCase() === 'TABLE') {
    var table = elem;}
else if (elem.tagName.toUpperCase() === 'TBODY'){
    var table = elem.parentElement;}
else {
    return false;
}
// Convert table element to an array (vba collection)
return htmlTableToArray(table, arguments[1], arguments[2], arguments[3], arguments[4]);