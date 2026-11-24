function toDataURI(str) {
    //remove unneeded whitespace before encoding
    const lines = str.replace(/\r\n/g, '\n').split('\n');
    const trimmedLines = lines.map(line => line.trim());
    const cleanedString = trimmedLines.join('\n');
    //now do the encoding and build the uri
    return 'data:text/html,' + encodeURIComponent(cleanedString.toWellFormed());
}
return toDataURI(arguments[0]);