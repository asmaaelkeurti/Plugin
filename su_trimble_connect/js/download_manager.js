// Copyright 2014 Trimble Navigation Ltd.


var options = {
  base_url : "" // Base URL to prepend relative URLs for thumbnails.
};


$(document).ready(function() {

  $("#cancel").on("click", function() {
    callback("cancel");
  });

});


function set_base_url(url) {
  options.base_url = url;
}


function list_files(files) {
  for (var i = 0; i < files.length; ++i) {
    var file = files[i];
    var $parent = $("#file_list");
    add_file($parent, file);
  }
}


function add_file($parent, file) {
  // Build the HTML element.
  var id = html_file_id(file);
  var image_id = 'image_' + id;
  var thumbnail = get_thumbnail(file);
  var filesize = get_filesize(file);
  var $item = $("\
    <div class='item' id='" + id + "'>\
      <div class='content'>\
        <span class='thumbnail'>\
          <img src='" + loading_image() + "' id='" + image_id + "'>\
        </span>\
        <span class='label'>" + file.name + "</span>\
        <span class='status'>\
          <img src='../images/ajax-loader-small.gif'>\
        </span>\
        <span class='filesize'>" + filesize + "</span>\
      </div>\
    </div>\
  ");

  // Store the data with the element.
  $item.data("data", file);

  $parent.append($item);
  load_image(image_id, thumbnail);
  return $item;
}


// Generate a valid HTML id given an id for the folder item. This is used to
// look up the element later.
function html_file_id(file) {
  return "item_" + file.id;
}


function update_info(result, message) {
  var file = result.file;
  update_status(file, message);
}


function update_status(file, message) {
  var id = html_file_id(file);
  $status = $("#" + id + " .status");
  $status.text(message);
}


// TODO(thomthom): Make generic version to share with FileDialog.
function get_thumbnail(file) {
  var url = "";
  if ("thumbnailUrl" in file) {
    if (file.thumbnailUrl instanceof Array) {
      url = file.thumbnailUrl[0];
    } else {
      return missing_image();
    }
  } else {
    return missing_image();
  }
  if (url.substring(0, 4) == "http") {
    return url;
  } else {
    return options.base_url + url;
  }
}


function get_filesize(file) {
  if ("size" in file && is_a_string(file.size)) {
    return file.size;
  } else {
    // TODO(thomthom): Localize.
    return l10n("Unknown size")
  }
}


function all_downloads_complete() {
  $(".item .status img").hide();
  $("#cancel").text(l10n("Close"));
  // All downloads have completed. Make sure we don't display any load
  // indicators.
  var $images = $(".item img[src='" + loading_image() + "']");
  $images.attr("src", missing_image());
}
