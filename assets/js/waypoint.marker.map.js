var markerClusterer = null;
var map             = null;
var imageUrl        = 'http://chart.apis.google.com/chart?cht=mm&chs=24x32&' + 'chco=FFFFFF,008CFF,000000&ext=.png';
var markers         = [];

//Function to initialize map
function initialize_map() {
  //create the map
  map = new google.maps.Map(document.getElementById('map'), {
    zoom: 2,
    center: new google.maps.LatLng(39.91, 116.38),
    mapTypeId: google.maps.MapTypeId.ROADMAP
  });

  //set the markers
  set_markers()

  //add lsitener to bounds change
  google.maps.event.addListener(map, 'bounds_changed', bounds_change);

  //generate the marker list
  generateMarkerViewList(markers)
}

//function for bounds change callback. changes the marker list on change.
//only lists the markers that are shown on the map
function bounds_change() {
  var marks = [];
  for (var i = 0; i < markers.length; i++) {
    if (map.getBounds().contains(markers[i].getPosition())) {
      marks.push(markers[i])
    }  else  {

    }
  }

  generateMarkerViewList(marks)
}

//function to set the markers
function set_markers() {
  var markerImage = new google.maps.MarkerImage(imageUrl, new google.maps.Size(24, 32));

  if ( data.items.length > 0 ) {
    var latlngbounds = new google.maps.LatLngBounds();

    for (var i = 0; i < data.items.length; ++i) {
      var latLng = new google.maps.LatLng(data.items[i].lat,data.items[i].lon)

      latlngbounds.extend(latLng);

      var marker = createMarker(data.items[i],latLng,markerImage)

      markers.push(marker);
    }

    map.fitBounds(latlngbounds);
  }

  //var size = 5

  //if ( markers.length > 0 ) {
  //  markerClusterer = new MarkerClusterer(map, markers, {
  //    gridSize: size
  //  });
  //}
}

//a function that creates a specific marker and adds a click loistener
function createMarker(marker,latLng,markerImage) {
  var mark = new google.maps.Marker({
    position: latLng,
    draggable: false,
    title: marker.title,
    map: map,
    icon: markerImage
  });

  mark.waypoint_id = marker.id

  mark.addListener('click', function() {
    if (mark.getAnimation() !== null) {
      mark.setAnimation(null);
    } else {
      //mark.setAnimation(google.maps.Animation.BOUNCE);
    }
    return false;
  });
  return mark
}

//a function that generates a marker list
function generateMarkerViewList(markers) {
  var $markerlist = $("#marker-list").find("ul")
  var list          = ""

  for (var i = 0; i < markers.length; i++) {
    li = '<li data-waypoint-id="' + markers[i].waypoint_id + '">'
    li += ' <span class="title">' +  markers[i].title + '</span>'
    li += ' <a class="view_info" href="member-waypoint-info.asp?id=' + markers[i].waypoint_id + '" target="_blank">info</a>'
    li += ' <a class="view_on_map" href="#">view</a>'
    li += ' <a class="toggle_visibility" href="#">hide</a>'
    li += '</li>'
    list += li
  }

  $markerlist.html(list)
}

function focus_on_mark(waypoint_id) {
  var zoom = 12;

  for (var i = 0; i < markers.length; i++) {
    mark = markers[i];

    if ( parseInt(mark.waypoint_id) == parseInt(waypoint_id) ) {
      if ( parseInt(map.getZoom()) <= parseInt(zoom) ) {
        map.setZoom(12);
      }

      map.panTo(mark.position);
      mark.setAnimation(google.maps.Animation.BOUNCE);
      i = markers.length;
    }
  }
}

function togle_mark_visibility(waypoint_id,override) {

  for (var i = 0; i < markers.length; i++) {
    mark = markers[i];

    if ( parseInt(mark.waypoint_id) == parseInt(waypoint_id) || parseInt(waypoint_id) == 0 ) {
      if ( override.length == 0 ) {
        if ( mark.visible ) {
          mark.setVisible(false);
          $('#marker-list li[data-waypoint-id="' + mark.waypoint_id + '"] .toggle_visibility').html("show");
        } else {
          mark.setVisible(true);
          $('#marker-list li[data-waypoint-id="' + mark.waypoint_id + '"] .toggle_visibility').html("hide");
        }
      } else {
        if ( override == "hide" ) {
          mark.setVisible(false);
          $('#marker-list li[data-waypoint-id="' + mark.waypoint_id + '"] .toggle_visibility').html("show");
        } else {
          mark.setVisible(true);
          $('#marker-list li[data-waypoint-id="' + mark.waypoint_id + '"] .toggle_visibility').html("hide");
        }
      }
    }
  }
}

$('#marker-list').on('click', '.view_info', function() {
  $.colorbox({href:$(this).attr('href'), width:"80%", height:"80%", iframe:true, current: '{current} of {total} waypoints'});
  return false;
});

$('#marker-list').on('click', '.view_on_map', function() {
  waypoint_id = $(this).parents("li").attr("data-waypoint-id")
  focus_on_mark(waypoint_id)
  return false;
});

$('#marker-list').on('click', '.toggle_visibility', function() {
  waypoint_id = $(this).parents("li").attr("data-waypoint-id")
  togle_mark_visibility(waypoint_id,"")
  return false;
});

$(document).ready(function(){
  $("#map-toolbar .hide_all").click(function(){
    togle_mark_visibility(0,"hide");
    return false;
  });
  $("#map-toolbar .show_all").click(function(){
    togle_mark_visibility(0,"show");
    return false;
  });
});

//initialize the map on load
google.maps.event.addDomListener(window, 'load', initialize_map);
