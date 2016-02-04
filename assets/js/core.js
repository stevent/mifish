$(document).ready( function() {
  //EMAIL REPLACEMENT
  $("span.safemail, div.safemail").each(function(){
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
  $("option.safemail").each(function(){
    $(this).val($(this).val().replace(/ at /,"@").replace(/ dot /g,"."));
  });

  //COLORBOX
  $(".popup_image").each(function() {
    $('.popup_image').colorbox({ maxWidth:"80%", maxHeight:"80%", transition: "elastic", scalePhotos: true });
  });
  $(".popup_page").each(function() {
    var current = ""
    var rel     = $(this).attr("rel")

    if ( !current.length ) {
      current = '{current} of {total} ' + rel
    }

    $('.popup_page').colorbox({ width:"80%", height:"80%", iframe:true, current: current });
  });
});
