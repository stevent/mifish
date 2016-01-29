$(document).ready( function() {
  // open external link in new tab/window
  // use rel="external" instead of target="_blank"
  $('a[rel="external"]').click( function() {
      this.target = "_blank";
  });

  $("ul#nav a").each(function() {
    setActiveNavItem($(this),/^nav-(.*)$/)
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

  //JQUERY CYCLE
  $(".cycle_image ul").cycle();

  //COLORBOX
  $(".colorbox_image").each(function() {
    $('.colorbox_image').colorbox({ maxWidth:"80%", maxHeight:"80%", transition: "elastic", scalePhotos: true });
  });
  $(".colorbox_page").each(function() {
    var current = ""
    var rel     = $(this).attr("rel")

    if ( !current.length ) {
      current = '{current} of {total} ' + rel
    }

    $('.colorbox_page').colorbox({ width:"80%", height:"80%", iframe:true, current: current });
  });

  //VALIDATE
  $(".validate_form").each(function() {
    $(this).validate({
      errorPlacement: function(error, element) {
        error.appendTo( element.parent("li") )
      },
      showErrors: function (errorMap, errorList) {
        this.defaultShowErrors();
        $.each(errorList, function (i, error)
        {
            $(error.element).parent().find("label.error").css("display", "inline");
        });
      }
    });
  });

  //TEXT REPLACEMENT
  $("input.replace_text").each(function() {
      val = $(this).val()
      $(this).val('')

     $(this).input_replacement({text: val});
  });
});

//set the current active navigation item
//inputs - link - object we are checking
//inputs - pattern - the pattern we are checking for (e.g. /^nav-(.*)$/ id prefix followed by anything, it returns the anything
function setActiveNavItem(link,pattern) {
  if ( link.attr("id").length > 0 ) {
    if ( (link.attr("id").match(pattern) != null) ) {
      active_class = link.attr("id").match(pattern)[1]
      if ( $("body").hasClass(active_class) ) {
        link.addClass('active')
      }
    }
  }
}

// jquery.input_replacement.js by Dana Woodman
//Replaces default input text.
(function($) {
    $.fn.input_replacement = function(options) {
        // Compile default options and user specified options.
        var opts = $.extend({}, $.fn.input_replacement.defaults, options);
        return $(this).each(function() {
            var obj = $(this);
            // Build element specific options.
            obj.o = $.meta ? $.extend({}, opts, $this.data()) : opts;
            // If field is empty, append text, classes, etc...
            if (obj.val() == '') {
                obj.val(obj.o.text);
                if (obj.o.empty_class) {
                    obj.addClass(obj.o.empty_class);
                };
                // Focus on the input has occurred.
                obj.bind('focus', function() {
                    if (obj.val() == obj.o.text) {
                        obj.val('');
                    };
                    if (obj.o.empty_class) {
                        obj.removeClass(obj.o.empty_class);
                    };
                });
                // Focus has been lost.
                obj.bind('blur', function() {
                    if (obj.val() == '') {
                        obj.val(obj.o.text);
                        if (obj.o.empty_class) {
                            obj.addClass(obj.o.empty_class);
                        };
                    };
                });
                // Clear out the values on reload so they arent loaded after refresh.
                $(window).unload(function() {
                   if (obj.val() == obj.o.text) {
                       obj.val('');
                   }; 
                });
                // If nothing was entered, make sure the "text" is not submitted by removing it.
                var form = obj.parents('form'); //.map(function () { return this.tagName; }).get().join(", ");
                if (form) {
                    form.find("input[type=submit]").each(function() {
                        $(this).bind('click', function() {
                            if (obj.val() == obj.o.text) {
                                obj.val('');
                            };
                        });
                    });
                };
            };
        });
    };

    $.fn.input_replacement.defaults = {
        text: '', // The text to put in the empty input field.
        empty_class: '' // A class to be applied to empty input field. Gets removed after 'focus'.
    };
})(jQuery);