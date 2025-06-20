function getAbsolutePath(src) {
    // return new URL(src, window.location.href).href;
    return URL.parse(src, window.location.href).href;
}

const elem = arguments[0]; 
const attr = arguments[1];
const tagName = elem.tagName;

// make sure attribute exists for element
if (!elem.hasAttribute(attr)) {return 'Error in DownloadResource method: could not find Attribute ';}
    
let resUrl;

// if srcset attribute for IMG tag then find the highest quality image
if (tagName.toUpperCase() === 'IMG' && attr.toLowerCase() === 'srcset') {
    
    const srcset = elem.getAttribute(attr);
    
    // Parse the srcset string:
    const sources = srcset.split(',').map(source => {
        const [url, descriptor] = source.trim().split(' ');
        return { url, descriptor };
    });    

    // Determine the highest quality image:
    let highestQuality = sources[0]; 

    for (const source of sources) {
        if (source.descriptor.endsWith('x')) {
            const pixelRatio = parseFloat(source.descriptor);
            if (pixelRatio > parseFloat(highestQuality.descriptor)) {
                highestQuality = source;
            }
        } else if (source.descriptor.endsWith('w')) {
            const width = parseInt(source.descriptor);
            if (width > parseInt(highestQuality.descriptor)) {
                highestQuality = source;
            }
        }
    }
    resUrl = highestQuality.url;
}
else {
    resUrl = elem.getAttribute(attr);
}

// Resolve the URL in case it is relative
const absolutePath = getAbsolutePath(resUrl);

return absolutePath;