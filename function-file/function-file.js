// The initialize function must be run each time a new page is loaded.
(function () {
  Office.initialize = function (reason) {
      // If you need to initialize something you can do so here.
  };
})();

// Your function must be in the global namespace.
function writeText(event) {
  // Implement your custom code here. The following code is a simple example.
  
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}