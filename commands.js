/* Ribbon button handler for Nine Four Format. */
function notify(message, isError) {
  var item = Office.context.mailbox.item;
  var opts = isError
    ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message }
    : {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: message,
        icon: "Icon.16x16",
        persistent: false
      };
  item.notificationMessages.replaceAsync("ninefourFormat", opts);
}

function formatEmail(event) {
  var item = Office.context.mailbox.item;
  var finish = function (msg, isError) {
    try { notify(msg, isError); } catch (e) {}
    event.completed();
  };

  item.body.getTypeAsync(function (typeRes) {
    if (typeRes.status === Office.AsyncResultStatus.Succeeded && typeRes.value === Office.CoercionType.Text) {
      finish("This draft is in plain text. Click Options > Switch to HTML, then Format.", true);
      return;
    }
    item.body.getAsync(Office.CoercionType.Html, function (res) {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        finish("Couldn't read the draft: " + res.error.message, true);
        return;
      }
      var out;
      try {
        out = NineFourFormat.transform(res.value, function (h) {
          return new DOMParser().parseFromString(h, "text/html");
        });
      } catch (e) {
        finish("Formatting failed: " + e.message, true);
        return;
      }
      item.body.setAsync(out, { coercionType: Office.CoercionType.Html }, function (setRes) {
        if (setRes.status !== Office.AsyncResultStatus.Succeeded) {
          finish("Couldn't update the draft: " + setRes.error.message, true);
        } else {
          finish("Formatted: Aptos 11 with signature.", false);
        }
      });
    });
  });
}

// Register for both newer (associate) and older (global function) hosts.
window.formatEmail = formatEmail;
Office.onReady(function () {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("formatEmail", formatEmail);
  }
});
