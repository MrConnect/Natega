/* =========================================
   Security & Anti-DevTools Scripts
   ========================================= */

// 1. Disable Right Click (Context Menu)
document.addEventListener('contextmenu', function (e) {
    e.preventDefault();
});

// 2. Disable Common Developer Tools Keyboard Shortcuts
document.onkeydown = function (e) {
    // Disable F12
    if (e.keyCode === 123) {
        return false;
    }

    // Disable Ctrl+Shift+I (Chrome, Firefox DevTools)
    // Disable Ctrl+Shift+J (Chrome Console)
    // Disable Ctrl+Shift+C (Inspect Element)
    if (e.ctrlKey && e.shiftKey && (e.keyCode === 73 || e.keyCode === 74 || e.keyCode === 67)) {
        return false;
    }

    // Disable Ctrl+U (View Source)
    if (e.ctrlKey && e.keyCode === 85) {
        return false;
    }
};

// 3. Basic DevTools detection (Debugger trap)
// This will pause execution if DevTools is open, making it annoying to inspect
setInterval(function () {
    (function () {
        return false;
    })['constructor']('debugger')();
}, 2000);
