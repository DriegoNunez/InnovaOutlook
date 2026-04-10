(function () {
  registerActions();

  if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
    Office.onReady(function () {
      registerActions();
    });
  }

  function registerActions() {
    if (typeof Office === "undefined") {
      return;
    }

    if (!Office.actions || typeof Office.actions.associate !== "function") {
      return;
    }

    Office.actions.associate("addAttachmentList", addAttachmentList);
  }

  function addAttachmentList(event) {
    var item = Office.context.mailbox.item;

    if (!item || typeof item.getAttachmentsAsync !== "function") {
      reportError(event, "This command is only available while composing a message.");
      return;
    }

    item.getAttachmentsAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(event, "Unable to read the message attachments.");
        return;
      }

      var attachmentNames = (result.value || []).filter(function (attachment) {
        if (attachment.isInline) {
          return false;
        }

        if (!attachment.name) {
          return false;
        }

        return true;
      }).map(function (attachment) {
        return attachment.name;
      });

      if (attachmentNames.length) {
        insertAttachmentList(item, attachmentNames, event);
        return;
      }

      getSharedFileLinkNames(item, function (sharedLinkNames) {
        if (!sharedLinkNames.length) {
          notify("No file attachments were found. If you used 'Upload and share' in New Outlook, try 'Attach as copy' or keep the file name visible in the shared link.");
          event.completed();
          return;
        }

        insertAttachmentList(item, sharedLinkNames, event);
      });
    });
  }

  function insertAttachmentList(item, fileNames, event) {
    var uniqueFileNames = dedupeFileNames(fileNames);
    var attachmentCount = uniqueFileNames.length;
    var title = attachmentCount === 1
      ? "Encl. (1 file)"
      : "Encl. (" + attachmentCount + " files)";

    item.body.getTypeAsync(function (bodyTypeResult) {
      if (bodyTypeResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(event, "Unable to determine the message body format.");
        return;
      }

      var isHtmlBody = bodyTypeResult.value === Office.MailboxEnums.BodyType.Html;
      var content = buildContent(title, uniqueFileNames, isHtmlBody);
      var options = {
        coercionType: isHtmlBody ? Office.CoercionType.Html : Office.CoercionType.Text
      };

      item.body.setSelectedDataAsync(content, options, function (insertResult) {
        if (insertResult.status !== Office.AsyncResultStatus.Succeeded) {
          reportError(event, "Unable to insert the attachment list at the current cursor position.");
          return;
        }

        event.completed();
      });
    });
  }

  function notify(message) {
    Office.context.mailbox.item.notificationMessages.replaceAsync("attachment-list-status", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      persistent: false
    });
  }

  function reportError(event, message) {
    notify(message);
    event.completed();
  }

  function getSharedFileLinkNames(item, callback) {
    item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {
      if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
        callback(extractSharedFileLinkNamesFromHtml(bodyResult.value));
        return;
      }

      item.body.getAsync(Office.CoercionType.Text, function (textResult) {
        if (textResult.status !== Office.AsyncResultStatus.Succeeded) {
          callback([]);
          return;
        }

        callback(extractSharedFileLinkNamesFromText(textResult.value));
      });
    });
  }

  function extractSharedFileLinkNamesFromHtml(html) {
    if (!html) {
      return [];
    }

    var parser = new DOMParser();
    var documentRoot = parser.parseFromString(html, "text/html");
    var anchors = Array.prototype.slice.call(documentRoot.querySelectorAll("a[href]"));

    return dedupeFileNames(anchors.map(function (anchor) {
      var href = (anchor.getAttribute("href") || "").trim();
      var text = normalizeWhitespace(anchor.textContent || "");

      if (!isSharedFileUrl(href)) {
        return "";
      }

      if (looksLikeFileName(text)) {
        return text;
      }

      return getFileNameFromUrl(href);
    }).filter(Boolean));
  }

  function extractSharedFileLinkNamesFromText(text) {
    if (!text) {
      return [];
    }

    var matches = String(text).match(/\bhttps?:\/\/\S+/gi) || [];

    return dedupeFileNames(matches.map(function (url) {
      if (!isSharedFileUrl(url)) {
        return "";
      }

      return getFileNameFromUrl(url);
    }).filter(Boolean));
  }

  function isSharedFileUrl(url) {
    var normalizedUrl = String(url || "").toLowerCase();

    return normalizedUrl.indexOf("sharepoint.com") >= 0
      || normalizedUrl.indexOf("onedrive.live.com") >= 0
      || normalizedUrl.indexOf("1drv.ms") >= 0;
  }

  function looksLikeFileName(value) {
    return /\.[a-z0-9]{2,8}$/i.test(String(value || "").trim());
  }

  function getFileNameFromUrl(url) {
    var match = String(url || "").match(/\/([^\/?#]+\.[a-z0-9]{2,8})(?:[?#]|$)/i);
    return match ? decodeURIComponent(match[1]) : "";
  }

  function dedupeFileNames(fileNames) {
    var seen = {};

    return (fileNames || []).map(function (name) {
      return normalizeWhitespace(name);
    }).filter(function (name) {
      if (!name) {
        return false;
      }

      var key = name.toLowerCase();
      if (seen[key]) {
        return false;
      }

      seen[key] = true;
      return true;
    });
  }

  function normalizeWhitespace(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function buildContent(title, fileNames, isHtmlBody) {
    if (!isHtmlBody) {
      return [title].concat(fileNames).join("\n");
    }

    var lines = fileNames.map(function (name) {
      return escapeHtml(name);
    });

    return [
      "<div style=\"font-size:9pt;font-style:italic;\">",
      "<div>", escapeHtml(title), "</div>",
      "<div>", lines.join("</div><div>"), "</div>",
      "</div>"
    ].join("");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }
}());
