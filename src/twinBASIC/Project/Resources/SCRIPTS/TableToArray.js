function htmlTableToArray(table, skipHeader, skipFooter, createSpanData, ignoreCellFormatting) {
// This function recursively converts [nested] table[s] into an array [vba collection]
// Cells are placed on a grid: a cell skips columns already filled by row spans above it,
// and (if createSpanData = true) its value is copied over its whole rowspan x colspan area
const rows = table.rows;
let v = [];
for (let r = 0; r < rows.length; r++) {
    v[r] = [];
}

for (let r = 0; r < rows.length; r++) {
    let colIdx = 0;
    for (let cell of rows[r].cells) {
        // Skip columns already filled by a row span from a row above
        if (createSpanData) {
            while (v[r][colIdx] !== undefined) {
                colIdx++;
            }
        }

        // Check if the cell contains a table, process only the first table if found
        let value = null;
        let foundTable = false;
        for (let cellChild of cell.children) {
            if (cellChild.tagName.toUpperCase() === 'TABLE') {
                // We have a nested table...
                value = htmlTableToArray(cellChild, skipHeader, skipFooter, createSpanData, ignoreCellFormatting);
                foundTable = true;
                break;
            }
        }

        // If no table found, store the text of the cell
        if (!foundTable) {
            if (ignoreCellFormatting) {
                // Store the complete text content, including hidden text
                value = cell.textContent.replace(/\xA0/g, ' ');
            } else {
                // Store the visible text content, including <br>'s and other white space formatting
                value = cell.innerText.replace(/\xA0/g, ' ');
            }
        }

        if (createSpanData) {
            // Copy the value over the whole rowspan x colspan area
            const rowSpan = Math.max(1, cell.rowSpan);
            const colSpan = Math.max(1, cell.colSpan);
            for (let i = 0; i < rowSpan && r + i < rows.length; i++) {
                for (let j = 0; j < colSpan; j++) {
                    v[r + i][colIdx + j] = value;
                }
            }
            colIdx += colSpan;
        } else {
            v[r][colIdx] = value;
            colIdx++;
        }
    }
}

// Make every row the same width
let maxCols = 0;
for (let row of v) {
    if (row.length > maxCols) {
        maxCols = row.length;
    }
}
for (let row of v) {
    for (let j = 0; j < maxCols; j++) {
        if (row[j] === undefined) {
            row[j] = null;
        }
    }
}

// Remove the header and footer rows if called for
if (skipHeader && table.tHead) {
    v.splice(0, table.tHead.rows.length);
}
if (skipFooter && table.tFoot) {
    v.splice(v.length - table.tFoot.rows.length);
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