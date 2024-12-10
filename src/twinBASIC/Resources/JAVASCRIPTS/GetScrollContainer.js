// Determine through an element's css properties if it is scrollable (overflow content)
function isScrollable(element){
    const style = window.getComputedStyle(element);
    return (element.scrollHeight > element.clientHeight || element.scrollWidth > element.clientWidth)
        && (style.overflow === 'auto' || style.overflow === 'scroll' || style.overflowY === 'auto' ||
        style.overflowY === 'scroll' || style.overflowX === 'auto' || style.overflowX === 'scroll');
}

// Recursively find an element's scroll container (scrollable parent)
function getScrollContainer(element){
    if (!element) return document.documentElement;
    if (isScrollable(element)) {return element;}
    else {return getScrollContainer(element.parentElement);}
}

return getScrollContainer(arguments[0]);
