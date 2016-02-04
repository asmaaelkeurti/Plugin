// Copyright 2014 Trimble Navigation Ltd.


var options = {
  open_dialog       : true,  // Set to true to pick only existing files.
  select_folder     : false, // Set to true to pick folders instead of files.
  allowed_filetypes : [],    // Arroy of allowed file types to open/save.
  base_url          : ""
};


$(document).ready(function() {
  // Hide the footer.
  toggle_footer(false);

  $("#ok").on("click", function() {
    action_ok();
  });

  $("#cancel").on("click", function() {
    callback("close_window");
  });

  $file_browser = $("#file_browser");
  show_loading_placeholder($file_browser, l10n("Loading projects..."));
});


function invalid_pick() {
  $('body').css('cursor', 'default');
}


function show_loading_placeholder($parent, message) {
  var $div = $("\
    <div class='loading_message'>\
      <img src='../images/ajax-loader.gif'>\
      <span>" + message + "</span>\
    </div>\
  ");
  $parent.append($div);
}


function hide_loading_placeholder() {
  // If the folder fetched is empty we don't have a handle for the parent. So
  // We just remove all the loadeders. There should be only one anyway.
  $(".loading_message").detach();
}


function set_loading_message(message) {
  $(".loading_message span").text(message);
}


function set_options(user_options) {
  if ("open_dialog" in user_options) {
    options.open_dialog = user_options.open_dialog;
    // Don't let the user enter filenames if this is an OpenDialog.
    $("#filename").attr("disabled", options.open_dialog);
    if (options.open_dialog) {
      $("#filename").hide();
    }
  }

  if ("filename" in user_options) {
    $("#filename input").val(user_options.filename);
  }

  if ("select_folder" in user_options) {
    options.select_folder = user_options.select_folder;
  }

  if ("allowed_filetypes" in user_options) {
    options.allowed_filetypes = user_options.allowed_filetypes;
  }

  if ("base_url" in user_options) {
    options.base_url = user_options.base_url;
  }

  if ("ok_label" in user_options) {
    $("#ok").text(user_options.ok_label);
  }

  if ("cancel_label" in user_options) {
    $("#cancel").text(user_options.cancel_label);
  }
}


function list_projects(projects) {
  var $parent = $("#file_browser");
  for (var i = 0; i < projects.length; ++i) {
    var project = projects[i];
    // Create a data item to match the folder structure.
    var item = {
      id           : project.rootId, // Folder ID
      name         : project.name,
      type         : "FOLDER",
      thumbnailUrl : [project.thumbnail], // Why is this an array?
      projectId    : project.id,
      parent_id    : null,
      version_id   : null,
      size         : project.size
    }
    add_folder_item($parent, item, project);
  }
  hide_loading_placeholder();
}


function list_folder_content(content) {
  for (var i = 0; i < content.length; ++i) {
    var item = content[i];
    var $parent = get_folder_element(item.parentId);
    var $children = $parent.children(".children");
    if ($children.length == 0) {
      $children = $("<div class='children'/>");
      $parent.append($children);
    }
    add_folder_item($children, item, item);
  }
  hide_loading_placeholder();
}


function fetch_folder_content($folder) {
  callback("get_folder_content", $folder.data("data"));
}


/* <div class="item folder" id="folder_XXXXXX">
 *   <div class="content">
 *     <span class="thumbnail">
 *       <img src="...">
 *     </span>
 *     <span class="label">Folder Name</span>
 *     <span class="filesize">123 Kb</span>
 *   </div>
 *   <div class="children">
 *     ...
 *   </div>
 * </div>
 *
 * <div class="item file" id="folder_XXXXXX">
 *   <div class="content">
 *     <span class="thumbnail">
 *       <img src="...">
 *     </span>
 *     <span class="label">FileName.skp</span>
 *   </div>
 * </div>
 *
 * @param [jQuery] $parent Parent jQuery element to append to.
 * @param [Object] item Folder data.
 * @param [Object] data Raw data object to attach to the element.
 *
 * @return [jQuery]
 */
function add_folder_item($parent, item, data) {
  // Build the HTML element.
  var id = html_folder_item_id(item.id);
  var image_id = 'image_' + id;
  var type = item.type.toLowerCase(); // .type can be "FOLDER" or "FILE".
  var size = formatBytes(item.size);
  var thumbnail = get_item_thumbnail(item);
  var $item = $("\
    <div class='item " + type + "' id='" + id + "'>\
      <div class='content'>\
        <span class='thumbnail'>\
          <img src='" + loading_image() + "' id='" + image_id + "'>\
        </span>\
        <span class='label'>" + item.name + "</span>\
        <span class='filesize'>" + size + "</span>\
      </div>\
    </div>\
  ");

  // Store the data with the element.
  $item.data("data", data);

  if (!is_valid_selection($item)) {
    $item.addClass("invalid_selection");
  }

  // Hook up events.
  var $content = $item.children(".content");
  $content.on("click", function() {
    // Mark item as selected.
    $(".file_browser .selected").removeClass("selected");
    $item.addClass("selected");
    // Update the filename textbox.
    if (options.select_folder) {
      if ($item.data("data").type == "FOLDER") {
        var filename = $item.data("data").name;
        $("#filename input").val(filename);
      }
    } else {
      if ($item.data("data").type == "FILE") {
        var filename = $item.data("data").name;
        $("#filename input").val(filename);
      }
    }
    // Display the footer if the selection is valid.
    toggle_footer($item);
    return false;
  });
  if (type == 'folder') {
    $content.on("dblclick", function() {
      var $item = $(this).parent();
      $children = $item.children(".children");
      if ($children.length > 0) {
        // Close the folder.
        $children.detach();
        $item.removeClass("open");
      } else {
        // Open the folder/
        $item.addClass("open");
        show_loading_placeholder($item, l10n("Loading..."));
        fetch_folder_content($item);
      }
      return false;
    });
  } else {
    $content.on("dblclick", function() {
      action_ok();
      return false;
    });
  }

  $parent.append($item);
  load_image(image_id, thumbnail);
  return $item;
}


function toggle_footer($item) {
  var $footer = $("#footer");
  if ($item !== false && is_valid_selection($item)) {
    $footer.show();
    $("body").removeClass("footer_hidden");
  } else {
    $footer.hide();
    $("body").addClass("footer_hidden");
  }
}


function is_valid_selection($item) {
  // Verify something was selected.
  if ($item.length == 0) {
    return false;
  }

  // When selecting folders, make sure a folder was selected.
  if (options.select_folder) {
    if (!$item.hasClass("folder")) {
      return false;
    }
  }

  // When opening files, make sure a file was selected.
  if (options.open_dialog && !options.select_folder) {
    if (!$item.hasClass("file")) {
      return false;
    }

    if (options.allowed_filetypes.length > 0) {
      var data = $item.data("data");
      var filename = data.name.toLowerCase();
      var file_extension = filename.split(".").pop();
      if ($.inArray(file_extension, options.allowed_filetypes) == -1) {
        return false;
      }
    }
  }

  return true;
}


// Generate a valid HTML id given an id for the folder item. This is used to
// look up the element later.
function html_folder_item_id(item_id) {
  return "item_" + item_id;
}


// Returns the jQuery element given a folder item id.
function get_folder_element(item_id) {
  var id = html_folder_item_id(item_id)
  return $("#" + id);
}


// Composite the thumbnail URL to be complete as some URLs for files that
// haven't got their own thumbnail will have a generic icon which is relative.
function get_thumbnail(url) {
  if (url.substring(0, 4) == "http") {
    return url;
  } else {
    // TODO(thomthom): Is this still valid for API v2.0?
    return options.base_url + url;
  }
}


function get_item_thumbnail(item) {

  if ('thumbnailUrl' in item) {
    var url = item.thumbnailUrl;
    if ($.isArray(url)) {
      url = url[0];
    }
    return get_thumbnail(url); // TODO(thomthom): Rename to get_full_path.
  } else {
    return default_item_image(item);
  }
}


// The action to perform when the user picks a file item. Triggered by the Ok
// button or double click.
function action_ok() {
  var $item = $(".selected");

  // Verify something was selected.
  if ($item.length == 0) {
    alert(l10n("Nothing selected."));
    return false;
  }

  var data = $item.data("data");

  // When selecting folders, make sure a folder was selected.
  if (options.select_folder) {
    if (!$item.hasClass("folder")) {
      alert(l10n("Please select a folder."));
      return false;
    }
  }

  // When opening files, make sure a file was selected.
  if (options.open_dialog && !options.select_folder) {
    if (!$item.hasClass("file")) {
      alert(l10n("Please select a file."));
      return false;
    }

    if (options.allowed_filetypes.length > 0) {
      var filename = data.name.toLowerCase();
      var file_extension = filename.split(".").pop();
      if ($.inArray(file_extension, options.allowed_filetypes) == -1) {
        var valid_file_types = options.allowed_filetypes.join(", ");
        alert(l10n("Please select a valid file type: ") + valid_file_types);
        return false;
      }
    }
  }

  // Save file dialog.
  if (!options.open_dialog && !options.select_folder) {
    // When this is a save dialog we inject the filename the user want for the
    // uploade file.
    // TODO(thomthom): Might want to indicate if we overwrite a file.
    var save_filename = $("#filename input").val();
    if (save_filename.length == 0) {
      alert(l10n("The chosen filename is not valid."));
      return false;
    }
    data["save_filename"] = save_filename
  }

  $('body').css('cursor', 'wait');
  callback("pick_item", data);
  return true;
}
