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

// Determine if the element is the document root
function isDocumentElement(obj){return (obj === document.documentElement);}

// Get the value of the container element's scroll behavior ('auto' or 'smooth')
function getScrollBehavior(container){return getComputedStyle(container).getPropertyValue('scroll-behavior');}

// This function calls the scrollIntoView method on the element for a smooth scroll 
// and listens for the scrollend event to resolve a promise  
const scrollIntoView = function(elem, container, options = {}) {
    return new Promise(function(resolve, _) {
        container.addEventListener('scrollend', (e) => {
            resolve('animation call');
        }, {once: true});
        elem.scrollIntoView(options);
    });
};

// Set a timeout to create a fallback in case the scroll operation takes longer than 1500 ms.
// This should only trigger if scrollend event does not trigger due to user trying to scroll in a
// non-scrollable container (animated scrolls in Chromium and FF are time-limited and should finish < 1500 ms).
let timeoutId;
const timeoutPromise = new Promise((resolve, _) => {
    timeoutId = setTimeout(() => resolve('animation call with timeout'), 1500);
});

// processing starts here...
var elem = arguments[0];
const options = JSON.parse(arguments[1]);

// prepare to manage animation if smooth
if (options.behavior === 'smooth') {var container = getScrollContainer(elem);}

// if desired scroll is set to auto (default), then see if we can resolve by finding container's scroll-behavior
if (options.behavior === 'auto') {
    var container = getScrollContainer(elem);
    if (getScrollBehavior(container) === 'auto') {options.behavior = 'instant';}
}

// if instant, then make the standard call, otherwise manage animation
if (options.behavior === 'instant') {
    elem.scrollIntoView(options); return 'standard call';}
else {
    // container cannot be documentElement for event listener
    if (isDocumentElement(container)) {container = document;}
    // wait for animation to complete or non-error throwing timeout, which ever one happens first
    return Promise.race ([scrollIntoView(elem, container, options), timeoutPromise]).finally(() => {
        clearTimeout (timeoutId); // clear the timeout when the Promise race finishes
    });
}
