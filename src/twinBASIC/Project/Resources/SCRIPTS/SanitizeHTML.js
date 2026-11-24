function sanitizeHTML(str) {
    // This function renders html source 'offline' by disabling dynamic content, while leaving the DOM tree intact.
    // The source is loaded into a simulated document fragment using a div element as the root.
    // This way the source can be parsed using DOM while disconnected from a html document root.
    
    // Rename html, head, and body tags to dummy names
    str = str.replace(/<html/gi, '<html_').replace(/<\/html/gi, '</html_');
    str = str.replace(/<head/gi, '<head_').replace(/<\/head/gi, '</head_');
    str = str.replace(/<body/gi, '<body_').replace(/<\/body/gi, '</body_');

    // Create a new DOM parser and a dummy div element to simulate document fragment
    const parser = new DOMParser();
    const doc = parser.parseFromString('<div>' + str + '</div>', 'text/html');
    const parentElem = doc.querySelector('div');

    // Disable script elements
    const scripts = parentElem.querySelectorAll('script');
    scripts.forEach(script => {
        script.innerText = ''; // Remove script content
    });

    // Disable meta refresh redirect
    const metas = parentElem.querySelectorAll('meta[http-equiv="refresh"]');
    metas.forEach(meta => {
        meta.setAttribute('content', ''); // Disable refresh attribute
    });

    // Disable src attributes
    const srcElems = parentElem.querySelectorAll('[src]');
    srcElems.forEach(elem => {
        elem.setAttribute('src', ''); // Remove src attribute
    });

    // Disable srcset attributes
    const srcsetElems = parentElem.querySelectorAll('[srcset]');
    srcsetElems.forEach(elem => {
        elem.setAttribute('srcset', ''); // Remove srcset attribute
    });

    // Disable href attributes, except about:blank ones
    const hrefElems = parentElem.querySelectorAll('[href]');
    hrefElems.forEach(elem => {
        const href = elem.getAttribute('href');
        if (!href.startsWith('about:blank')) {
            elem.setAttribute('href', ''); // Remove href attribute
        }
    });

    // Disable rel attributes
    const relElems = parentElem.querySelectorAll('[rel]');
    relElems.forEach(elem => {
        elem.setAttribute('rel', ''); // Remove rel attribute
    });

    // Disable url() in style elements (e.g., background-image URLs)
    const styleElems = parentElem.querySelectorAll('style');
    styleElems.forEach(style => {
        let css = style.innerHTML.toLowerCase();
        let regex = /url\((.*?)\)/g;
        css = css.replace(regex, 'url()'); // Remove url() references
        style.innerHTML = css; // Update the style content
    });

    // Disable url() in inline style attributes
    const inlineStyleElems = parentElem.querySelectorAll('[style*="url("]');
    inlineStyleElems.forEach(elem => {
        let style = elem.getAttribute('style');
        let regex = /url\((.*?)\)/g;
        style = style.replace(regex, 'url()'); // Remove url() references
        elem.setAttribute('style', style); // Update the style attribute
    });

    // Disable action attributes in form elements
    const formElems = parentElem.querySelectorAll('form[action]');
    formElems.forEach(form => {
        form.setAttribute('action', ''); // Remove action attribute
    });

    // Disable 'as' attribute in link elements with value 'script'
    const linkElems = parentElem.querySelectorAll('link[as="script"]');
    linkElems.forEach(link => {
        link.setAttribute('as', ''); // Remove as='script'
    });

    // Disable data attributes in object elements
    const objectElems = parentElem.querySelectorAll('object[data]');
    objectElems.forEach(obj => {
        obj.setAttribute('data', ''); // Remove data attribute
    });

    // Disable event attributes (onclick, onload, etc.)
    const eventAttrs = [
        'onclick', 'onload', 'onoffline', 'ononline', 'onunload', 'onpageshow',
        'onpagehide', 'onerror', 'onresize', 'onhashchange', 'onbeforeunload',
        'onpopstate', 'onmessage'
    ];
    eventAttrs.forEach(eventAttr => {
        const elems = parentElem.querySelectorAll(`[${eventAttr}]`);
        elems.forEach(elem => {
            elem.removeAttribute(eventAttr); // Remove event handler attributes
            // Add dummy attributes
            elem.setAttribute(eventAttr + '___xxx123', ''); // Adding a dummy attribute
        });
    });

    // Get the sanitized HTML as a string
    let strOut = parentElem.innerHTML;

    // Rename dummy tags back to original tag names
    strOut = strOut.replace(/<html_/g, '<html').replace(/<\/html_/g, '</html');
    strOut = strOut.replace(/<head_/g, '<head').replace(/<\/head_/g, '</head');
    strOut = strOut.replace(/<body_/g, '<body').replace(/<\/body_/g, '</body');
    strOut = strOut.replace(/___xxx123/g, ''); // Remove dummy attribute

    return strOut;
}
 
return sanitizeHTML(arguments[0]);