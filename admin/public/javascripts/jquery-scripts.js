$(document).ready( function() {
  $("tr.hover").hover(
    function () {
      $(this).find("td").each(function() {
        $(this).addClass("hover")
      });
    }, 
    function () {
      $(this).find("td").each(function() {
        $(this).removeClass("hover")
      });
    }
  );
});
