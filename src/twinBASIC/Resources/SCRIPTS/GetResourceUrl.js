function getAbsolutePath(src) {
// Resolve relative URLs against the current page location
return new URL(src, window.location.href).href;
}

const elem = arguments[0]; 
const attr = arguments[1];
const tagName = elem.tagName;

// make sure attribute exists for element
if (!elem.hasAttribute(attr)) {
return 'Error in DownloadResource method: could not find Attribute ' + attr;
}

let resUrl;

// if srcset attribute for IMG tag then find the highest quality image
if (tagName.toUpperCase() === 'IMG' && attr.toLowerCase() === 'srcset') {
    
const srcset = elem.getAttribute(attr);
    
// Parse the srcset string into {url, descriptor} objects
const sources = srcset.split(',').map(source => {
    // split on whitespace; descriptor may be missing
    const [url, descriptor = "0w"] = source.trim().split(/\s+/);
    return { url, descriptor };
});    

// Initialize with the first candidate
let highestQuality = sources[0]; 

// Loop through candidates to find the highest quality
for (const source of sources) {
    if (source.descriptor.endsWith('x')) {
        // Handle pixel density descriptors (e.g. "2x")
        const pixelRatio = parseFloat(source.descriptor);
        if (parseFloat(highestQuality.descriptor) < pixelRatio) {
            highestQuality = source;
        }
    } else if (source.descriptor.endsWith('w')) {
        // Handle width descriptors (e.g. "1200w")
        const width = parseInt(source.descriptor, 10);
        if (parseInt(highestQuality.descriptor, 10) < width) {
            highestQuality = source;
        }
    }
}
resUrl = highestQuality.url;
}
else {
// Otherwise just use the attribute value directly (e.g. "src")
resUrl = elem.getAttribute(attr);
}

// Resolve the URL in case it is relative
const absolutePath = getAbsolutePath(resUrl);
return absolutePath;