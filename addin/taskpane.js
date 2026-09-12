(function () {
  var config = {
    headerName: "X-AssignLetters-Id",
    missingHeaderMessage:
      "This message does not include the required assignment header, so AssignLetters cannot open the form.",
    apiUrl: "",
  };
  var headerValue = "";
  var userEmail = "";
  var subject = "";
  var itemId = "";

  function $(id) {
    return document.getElementById(id);
  }

  function apiBase() {
    return (config.apiUrl || window.location.origin || "").replace(/\/$/, "");
  }

  function setStatus(message, kind) {
    var el = $("status");
    el.hidden = !message;
    el.textContent = message || "";
    el.className = "status" + (kind ? " " + kind : "");
  }

  function parseInternetHeader(raw, name) {
    if (!raw || !name) {
      return "";
    }
    var unfolded = String(raw)
      .replace(/\r\n/g, "\n")
      .replace(/\n[ \t]+/g, " ");
    var escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    var match = unfolded.match(new RegExp("^" + escaped + "\\s*:\\s*(.*)$", "im"));
    return match ? match[1].trim() : "";
  }

  function queryParam(name) {
    try {
      return new URLSearchParams(window.location.search).get(name);
    } catch (ignore) {
      return null;
    }
  }

  function isMockMode() {
    return queryParam("mock") === "1" || queryParam("mock") === "true";
  }

  function fillStaffOptions(staff) {
    var select = $("staff");
    select.innerHTML = '<option value="">Select staff</option>';
    (staff || []).forEach(function (row) {
      var option = document.createElement("option");
      option.value = JSON.stringify({
        StaffName: row.StaffName,
        Department: row.Department,
      });
      option.textContent = row.StaffName + " (" + row.Department + ")";
      select.appendChild(option);
    });
    if (!staff || !staff.length) {
      $("staff-error").textContent =
        "No staff rows match this mailbox. Update data/staff.json so Department contains your email.";
    }
  }

  function showForm() {
    $("form").hidden = false;
    $("subtitle").textContent =
      "Header " + config.headerName + " = " + headerValue;
  }

  function showMissingHeader() {
    $("form").hidden = true;
    setStatus(config.missingHeaderMessage, "info");
  }

  function loadStaffThenForm() {
    var url = apiBase() + "/api/staff";
    if (userEmail) {
      url += "?email=" + encodeURIComponent(userEmail);
    }
    return fetch(url, { method: "GET" })
      .then(function (res) {
        if (!res.ok) {
          throw new Error("Staff list request failed (" + res.status + ")");
        }
        return res.json();
      })
      .then(function (data) {
        fillStaffOptions(data.staff || []);
        showForm();
      });
  }

  function readInternetHeaders(item) {
    return new Promise(function (resolve, reject) {
      if (item && typeof item.getAllInternetHeadersAsync === "function") {
        item.getAllInternetHeadersAsync(function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve(result.value || "");
            return;
          }
          reject(
            new Error(
              (result.error && result.error.message) ||
                "Could not read Internet headers."
            )
          );
        });
        return;
      }
      if (
        item &&
        item.internetHeaders &&
        typeof item.internetHeaders.getAsync === "function"
      ) {
        item.internetHeaders.getAsync([config.headerName], function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            var headers = result.value || {};
            resolve(
              config.headerName + ": " + (headers[config.headerName] || "")
            );
            return;
          }
          reject(
            new Error(
              (result.error && result.error.message) ||
                "Could not read Internet headers."
            )
          );
        });
        return;
      }
      reject(new Error("This Outlook host cannot read Internet headers."));
    });
  }

  function afterConfig() {
    if (isMockMode()) {
      userEmail = queryParam("email") || "ada@contoso.com";
      subject = queryParam("subject") || "Mock assignment message";
      itemId = "mock-item";
      var mockHeader = queryParam("header");
      if (mockHeader === null) {
        headerValue = "AL-1001";
      } else {
        headerValue = mockHeader;
      }
      $("subtitle").textContent = "Browser mock (not Outlook)";
      if (!headerValue) {
        showMissingHeader();
        return;
      }
      loadStaffThenForm().catch(function (err) {
        setStatus("Failed to load staff: " + err.message, "fail");
      });
      return;
    }

    if (!window.Office || !Office.onReady) {
      setStatus("Office.js did not load. Open this page from Outlook on the web.", "fail");
      return;
    }

    Office.onReady(function (info) {
      if (!info || info.host !== Office.HostType.Outlook) {
        setStatus(
          "Open this add-in from Outlook on the web (Read mode), or use /taskpane.html?mock=1 to preview.",
          "info"
        );
        return;
      }

      var mailbox = Office.context.mailbox;
      var item = mailbox && mailbox.item;
      userEmail =
        (mailbox && mailbox.userProfile && mailbox.userProfile.emailAddress) ||
        "";
      subject = (item && item.subject) || "";
      itemId = (item && (item.itemId || item.internetMessageId)) || "";

      readInternetHeaders(item)
        .then(function (raw) {
          headerValue = parseInternetHeader(raw, config.headerName);
          if (!headerValue) {
            showMissingHeader();
            return;
          }
          return loadStaffThenForm();
        })
        .catch(function (err) {
          setStatus(err.message || "Failed", "fail");
        });
    });
  }

  function loadConfig() {
    var base = window.location.origin || "";
    return fetch(base + "/api/config")
      .then(function (res) {
        if (!res.ok) {
          throw new Error("Config request failed (" + res.status + ")");
        }
        return res.json();
      })
      .then(function (data) {
        config.headerName = data.headerName || config.headerName;
        config.missingHeaderMessage =
          data.missingHeaderMessage || config.missingHeaderMessage;
        config.apiUrl = data.apiUrl || base;
      })
      .catch(function () {
        config.apiUrl = base;
      })
      .then(afterConfig);
  }

  $("form").addEventListener("submit", function (event) {
    event.preventDefault();
    $("staff-error").textContent = "";
    $("deadline-error").textContent = "";
    setStatus("", "");

    var rawStaff = $("staff").value;
    var deadline = $("deadline").value;
    var valid = true;
    if (!rawStaff) {
      $("staff-error").textContent = "Choose a staff member.";
      valid = false;
    }
    if (!deadline) {
      $("deadline-error").textContent = "Choose a deadline date.";
      valid = false;
    }
    if (!valid) {
      return;
    }

    var staff;
    try {
      staff = JSON.parse(rawStaff);
    } catch (err) {
      $("staff-error").textContent = "Invalid staff selection.";
      return;
    }

    var saveButton = $("save");
    saveButton.disabled = true;
    setStatus("Saving…", "info");

    fetch(apiBase() + "/api/save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        staffName: staff.StaffName,
        department: staff.Department,
        deadlineDate: deadline,
        internetHeaderName: config.headerName,
        internetHeaderValue: headerValue,
        userEmail: userEmail,
        subject: subject,
        itemId: itemId,
      }),
    })
      .then(function (res) {
        return res.json().then(function (body) {
          return { ok: res.ok && body && body.ok, body: body };
        });
      })
      .then(function (result) {
        if (result.ok) {
          setStatus("Success", "ok");
          return;
        }
        setStatus("Failed", "fail");
      })
      .catch(function () {
        setStatus("Failed", "fail");
      })
      .then(function () {
        saveButton.disabled = false;
      });
  });

  loadConfig();
})();
