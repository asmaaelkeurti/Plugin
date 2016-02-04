// Copyright 2014 Trimble Navigation Ltd.


var KEYCODE_ENTER = 13;


// A hash with strings used for localization.
var l10n_strings = {};

var access_token;


// Utility method to call back to Ruby, taking an optional JSON object as
// payload.
// TODO(thomthom): Not sure how big objects we can pass. Probably a limit.
// A safer solution would be to store the JSON data in a hidden input element
// and use UI::WebDialog.get_element_value.
function callback(name, data) {
  // Defer with a timer in order to allow the UI to update.
  setTimeout(function() {
    var $bridge = $("#SU_BRIDGE");
    $bridge.text("");
    if (data !== undefined) {
      var json = JSON.stringify(data);
      $bridge.text(json);
    }
    window.location = "skp:callback@" + name;
  }, 0);
}


$(document).ready(function() {

  create_bridge();
  disable_context_menu();
  disable_select();
  hook_up_default_button();

  callback("html_ready");

});


//noinspection JSUnusedGlobalSymbols
function log_error(error) {
  if (error.name == 'SyntaxError') {
    debugger; // TODO(thomthom): Don't call this in release.
  }
  //noinspection JSUnresolvedVariable
  var data = {
    'code'      : error.number,
    'filename'  : error.fileName,     // Missing in IE.
    'name'      : error.name,
    'line'      : error.lineNumber,   // Missing in IE.
    'col'       : error.columnNumber, // Missing in IE.
    'stack'     : error.stack,
    'string'    : error.toString()
  };
  callback("log_error", data);
}


// Creates a hidden textarea element used to pass data from JavaScript to
// Ruby. Ruby calls UI::WebDialog.get_element_value to fetch the content and
// parse it as JSON.
// This avoids many issues in regard to transferring data:
// * Data can be any size.
// * Avoid string encoding issues.
// * Avoid evaluation bug in SketchUp when the string has a single quote.
function create_bridge() {
  var $bridge = $("<textarea id='SU_BRIDGE'></textarea>");
  $bridge.hide();
  $("body").append($bridge);
}


function hook_up_default_button() {
  var $default_button = $("button[type=submit]");
  if ($default_button.length == 1) {

    $(document).keypress(function (event) {
      if (event.which == KEYCODE_ENTER) {
        $default_button.trigger('click');
        event.preventDefault();
        event.stopPropagation();
        return false;
      }
    });

  }
}


/* Disables text selection on elements other than input type elements where
 * it makes sense to allow selections. This mimics native windows.
 */
function disable_select() {
  $(document).on('mousedown selectstart', function(e) {
    return $(e.target).is('input, textarea, select, option');
  });
}


/* Disables the context menu with the exception for textboxes in order to
 * mimic native windows.
 */
function disable_context_menu() {
  $(document).on('contextmenu', function(e) {
    return $(e.target).is('input[type=text], input[type=email], input[type=password], textarea');
  });
}


// Gotto love JavaScript...
// http://stackoverflow.com/a/9436948/486990
function is_a_string(object) {
  return object instanceof String || typeof object == 'string';
}


function loading_image() {
  return '../images/ajax-loader.gif';
}

function missing_image() {
  return '../images/missing-image.svg';
}

function default_item_image(item) {
  switch (item.type) {
    case 'FOLDER':
      return '../images/folder_60.svg';
    case 'FILE':
      return '../images/file_60.svg';
    default:
      return missing_image();
  }
}


function debug(string) {
  callback("debug", { 'message' : string });
}

//noinspection JSUnusedGlobalSymbols
function set_access_token(token) {
  access_token = token;
}


var request_id = 0; // For debugging.
// @param [String] image_id The name of the IMG element to receive the data.
// @param [String] image_url
function load_image(image_id, image_url) {
  debug('load_image(' + image_id + ')');
  debug('> ' + image_url);
  if (image_url.substring(0, 4) != "http") {
    debug('> Not an absolute URI.');
    $('#' + image_id).attr('src', image_url);
    return;
  }

  var id = request_id + 1;
  request_id++;
  debug('download (' + id + ')');
  var xhr = new XMLHttpRequest();
  // Need to strip away the anchor part of the URI, otherwise the request fail.
  // What we end up stripping away is: #?1414516600000
  // Assuming this is a way to avoid caching? Might we need to do something else
  // to counter this?
  var base_url = image_url.split('#')[0];
  debug(image_url);
  debug(base_url);
  // https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest/Sending_and_Receiving_Binary_Data
  xhr.open("GET", base_url, true);
  // Check if we can use arraybuffer - fall back to not loading the image.
  try {
    xhr.responseType = 'arraybuffer';
    if (xhr.responseType !== 'arraybuffer') {
      //noinspection ExceptionCaughtLocallyJS
      throw 'Array buffer not supported.';
    }
    // IE9 doesn't seem to fail the above test, even though it doesn't support
    // arraybuffer.
    if (typeof(Uint8Array) === 'undefined') {
      //noinspection ExceptionCaughtLocallyJS
      throw 'Uint8Array not supported.';
    }
  } catch (error) {
    set_failed_image(image_id);
    debug(error.message);
    return;
  }
  // Pass in the access token the Trimble Connect API require.
  xhr.setRequestHeader("Authorization", "Bearer " + access_token);
  xhr.onreadystatechange = function() {
    if (xhr.readyState == 4)
    {
      debug('> ready (' + id + ')');
      debug(image_url);
      debug(image_id);
      debug(xhr.status);
      if (xhr.status == 200)
      {
        debug('Success!');
        try {
          set_binary_image_data(xhr, image_id);
        } catch(error) {
          set_failed_image(image_id);
          debug(error.message);
        }
      } else {
        set_failed_image(image_id);
        debug('Failure!');
        var text = String.fromCharCode.apply(null, new Uint8Array(xhr.response));
        debug(text);
      }
    }
  };
  xhr.send();
}

function set_failed_image(image_id) {
  var $img = $('#' + image_id);
  $img.attr('src', missing_image());
}

function set_binary_image_data(xhr, image_id) {
  var content_type = xhr.getResponseHeader('content-type');
  var base64_data = base64ArrayBuffer(xhr.response);
  var image_data = 'data:' + content_type + ';base64,' + base64_data;
  var $img = $('#' + image_id);
  $img.attr('src', image_data);
}


// http://stackoverflow.com/a/18650828/486990
function formatBytes(bytes, decimals) {
  if (bytes == 0) return '0 Bytes';
  if (bytes == 1) return '1 Byte';
  var k = 1000;
  var dm = decimals + 1 || 3;
  var sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
  var i = Math.floor(Math.log(bytes) / Math.log(k));
  return (bytes / Math.pow(k, i)).toPrecision(dm) + ' ' + sizes[i];
}


// Returns a localized string if such exist.
function l10n(string) {
  var result = l10n_strings[string];
  if (result === undefined) {
    return string;
  } else {
    return result;
  }
}


// Call this method from WebDialog.
// This collects all the strings in the HTML that needs to be localized.
//noinspection JSUnusedGlobalSymbols
function localize(strings) {
  l10n_strings = strings;
  $(".localize").each(function() {
    var $this = $(this);
    var type = $this.prop('tagName').toLowerCase();
    var input, output = '';

    switch(type) {

      case "input":
        input = $this.attr("placeholder");
        output = l10n(input);
        $this.attr("placeholder", output);
        break;

      default:
        input = $this.text();
        output = l10n(input);
        $this.text(output);
    }

  });
}
