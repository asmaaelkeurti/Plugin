// Copyright 2014 Trimble Navigation Ltd.


$(document).ready(function() {

  $("#login").on("click", function() {
    var email = $("#email input").val();
    var password = $("#password input").val();

    disable_form();
    show_load_screen(l10n("Signing in..."));

    data = {
      "email"    : email,
      "password" : password
    }
    callback("login_user", data);

  });

  $("#cancel").on("click", function() {
    callback("close_window");
  });

});


function show_load_screen(message) {
  var $div = $("\
    <div id='load_screen'>\
      <div>\
        <img src='../images/ajax-loader.gif'>\
        <span>" + message + "</span>\
      </div>\
    </div>\
  ");
  $("body").append($div);
}

function close_load_screen() {
  $("#load_screen").detach();
}


function disable_form() {
  $("input, button").prop("disabled", true);
}

function enable_form() {
  $("input, button").prop("disabled", false);
}


function update_defaults(email) {
  if (email !== null) {
    $("#email").val(email);
  }
}


function notify_error(message) {
  alert(message);
  enable_form();
  close_load_screen();
}
