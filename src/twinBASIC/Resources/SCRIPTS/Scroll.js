// Determine if the element is the document root
function isDocumentElement(obj){return (obj === document.documentElement);}

// Get the value of the container element's scroll behavior ('auto' or 'smooth')
function getScrollBehavior(container){return getComputedStyle(container).getPropertyValue('scroll-behavior');}

// This function calls the scrollTo method on the element for a smooth scroll
// and listens for the scrollend event to resolve a promise
const scroll = function(elem, container, scrollType, options = {}) {
    return new Promise(function(resolve, _) {
        container.addEventListener('scrollend', (e) => {
            resolve('animation call');
        }, {once: true});
        if (scrollType === 'to') {elem.scrollTo(options);} else {elem.scrollBy(options);}
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
if (typeof elem === 'string') {var elem =  document.scrollingElement;};
const scrollType = arguments[1];
const options = JSON.parse(arguments[2]);
var container = elem;

// if desired scroll is set to auto (default), then see if we can resolve by finding container's scroll-behavior
if (options.behavior === 'auto') {
    if (getScrollBehavior(container) === 'auto') {options.behavior = 'instant';}
}

// if instant, then make the standard call, otherwise manage animation
if (options.behavior === 'instant') {
    if (scrollType === 'to') {elem.scrollTo(options);} else {elem.scrollBy(options);} 
    return 'standard call';}
else {
    // container cannot be documentElement for event listener
    if (isDocumentElement(container)) {container = document;}
    // wait for animation to complete ot timeout, which ever one happens first
    return Promise.race ([scroll(elem, container, scrollType, options), timeoutPromise]).finally(() => {
        clearTimeout (timeoutId); // clear the timeout when the Promise race finishes
    });
}