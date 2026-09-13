(function () {
  var appBase = (window.location.origin || "").replace(/\/$/, "");
  var config = {
    headerName: "X-AssignLetters-Id",
    missingHeaderMessage:
      "This message does not include the required assignment header, so AssignLetters cannot open the form.",
    saveUrl: "",
  };
  var headerValue = "";
  var userEmail = "";
  var subject = "";
  var itemId = "";

  function $(id) {
    return document.getElementById(id);
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
    var url = appBase + "/api/staff";
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
    return fetch(appBase + "/api/config")
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
        config.saveUrl = data.saveUrl || data.apiUrl || "";
      })
      .catch(function () {
        config.saveUrl = "";
      })
      .then(afterConfig);
  }

  function logPushStatus(assignment, result) {
    return fetch(appBase + "/api/save", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        staffName: assignment.staffName,
        department: assignment.department,
        deadlineDate: assignment.deadlineDate,
        internetHeaderName: assignment.internetHeaderName,
        internetHeaderValue: assignment.internetHeaderValue,
        userEmail: assignment.userEmail,
        subject: assignment.subject,
        itemId: assignment.itemId,
        pushUrl: result.url || config.saveUrl || "",
        pushOk: result.ok,
        pushHttpStatus: result.status || null,
        pushError: result.error || null,
      }),
    }).catch(function () {
      return null;
    });
  }

  function pushAssignment(url, assignment) {
    return fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(assignment),
    }).then(function (res) {
      return res
        .json()
        .catch(function () {
          return {};
        })
        .then(function (body) {
          var ok = res.ok && !(body && body.ok === false);
          return {
            ok: ok,
            status: res.status,
            url: url,
            error: ok ? null : "Remote API returned HTTP " + res.status,
          };
        });
    });
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

    var assignment = {
      staffName: staff.StaffName,
      department: staff.Department,
      deadlineDate: deadline,
      internetHeaderName: config.headerName,
      internetHeaderValue: headerValue,
      userEmail: userEmail,
      subject: subject,
      itemId: itemId,
    };

    var saveButton = $("save");
    saveButton.disabled = true;
    setStatus("Saving…", "info");

    var dest = (config.saveUrl || "").replace(/\/$/, "");
    var push;
    if (!dest) {
      push = Promise.resolve({
        ok: false,
        status: 0,
        url: "",
        error: "ASSIGNLETTERS_API_URL is not configured.",
      });
    } else {
      push = pushAssignment(dest, assignment);
    }

    push
      .catch(function (err) {
        return {
          ok: false,
          status: 0,
          url: dest,
          error: (err && err.message) || "Network error",
        };
      })
      .then(function (result) {
        setStatus(result.ok ? "Success" : "Failed", result.ok ? "ok" : "fail");
        return logPushStatus(assignment, result);
      })
      .then(function () {
        saveButton.disabled = false;
      });
  });

  loadConfig();
})();
