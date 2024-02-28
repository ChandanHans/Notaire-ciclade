(function() {
  window.addEventListener('message', function(event) {
    if (event.data.type !== 'capsolverCallback')return;
    console.log("solved");
  })
})();