$(document).ready( function() {
  // open external link in new tab/window
  // use rel="external" instead of target="_blank"
  $('a[rel="external"]').click( function() {
      this.target = "_blank";
  });

  //CLICKABLE LI
  //if you have a list item and you need the each of the whole li's to link to a url, use this.
  //This will grab the first <a> tag (inside the selected li) and use it as the location to open (on li click).
  $("ul.clickable_li li").each(function() {
    var link = $(this).find("a").attr("href")

    if ( link ) {
      $(this).addClass("pointer")
      $(this).bind('click', function(){window.location = link});
    }
  });

  //EMAIL REPLACEMENT
  $("span.safemail").each(function(){
    exp = $(this).text().search(/\((.*?)\)/) != -1 ? new RegExp(/(.*?) \((.*?)\)/) : new RegExp(/.*/);
    match = exp.exec($(this).text());
    addr = match[1] ? match[1].replace(/ at /,"@").replace(/ dot /g,".") : match[0].replace(/ at /,"@").replace(/ dot /g,".");
    emaillink = match[2] ? match[2] : addr;
    subject = $(this).attr('title') ? "?subject="+$(this).attr('title').replace(/ /g,"%20") : "";
     $(this).after('<a href="mailto:'+addr+subject+'">'+ emaillink + '</a>');
    $(this).remove();
  });
  $("input.safemail").each(function(){
    $(this).val($(this).val().replace(/ at /,"@").replace(/ dot /g,"."));
  });

  //VALIDATE
  $(".validate_form").each(function() {
      $(this).validate();
  });

  //TEXT REPLACEMENT
  $("input.replace_text").each(function() {
      val = $(this).val()
      $(this).val('')

     $(this).input_replacement({text: val});
  });
});