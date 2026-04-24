/* Courses, Reminders, Preferences — /api/user/* */
(function () {
  function csrf() {
    return $("meta[name=csrf-token]").attr("content");
  }
  $.ajaxSetup({
    beforeSend: function (xhr, settings) {
      if (!/^(GET|HEAD|OPTIONS|TRACE)$/i.test(settings.type) && !settings.crossDomain) {
        xhr.setRequestHeader("X-CSRFToken", csrf());
      }
    },
  });

  if ($("#coursesRoot").length) {
    function showE(msg) {
      $("#courseErr").text(msg).removeClass("d-none");
      setTimeout(function () {
        $("#courseErr").addClass("d-none");
      }, 4000);
    }
    function load() {
      $.getJSON("/api/user/courses")
        .done(function (d) {
          const $u = $("#courseList").empty();
          (d.courses || []).forEach(function (c) {
            const $li = $(
              "<li class='list-group-item d-flex justify-content-between align-items-center' style='background:rgba(255,255,255,0.04);border-color:var(--border);color:var(--text);'></li>"
            );
            $li.append(
              $("<span></span>").html(
                "<strong>" + $("<div/>").text(c.code).html() + "</strong> · " + $("<div/>").text(c.title).html()
              )
            );
            $li.append(
              $("<button type='button' class='btn btn-sm btn-outline-danger'>Remove</button>").on("click", function () {
                if (!confirm("Remove?")) return;
                $.ajax({ url: "/api/user/courses/" + c.id, method: "DELETE" })
                  .done(load)
                  .fail(function (xhr) {
                    showE((xhr.responseJSON && xhr.responseJSON.error) || "Failed");
                  });
              })
            );
            $u.append($li);
          });
        })
        .fail(function () {
          showE("Could not load.");
        });
    }
    $("#btnAddCourse").on("click", function () {
      const code = ($("#cCode").val() || "").trim();
      const title = ($("#cTitle").val() || "").trim();
      if (!code || !title) {
        showE("Enter code and title.");
        return;
      }
      $.ajax({
        url: "/api/user/courses",
        method: "POST",
        contentType: "application/json",
        data: JSON.stringify({ code: code, title: title }),
      })
        .done(function () {
          $("#cCode").val("");
          $("#cTitle").val("");
          load();
        })
        .fail(function (xhr) {
          showE((xhr.responseJSON && xhr.responseJSON.error) || "Failed");
        });
    });
    load();
    return;
  }

  if ($("#remindersRoot").length) {
    function showR(msg) {
      $("#remErr").text(msg).removeClass("d-none");
      setTimeout(function () {
        $("#remErr").addClass("d-none");
      }, 4000);
    }
    function load() {
      $.getJSON("/api/user/reminders")
        .done(function (d) {
          const $u = $("#remList").empty();
          (d.reminders || []).forEach(function (m) {
            const due = m.due_at ? String(m.due_at).replace("T", " ").slice(0, 16) : "—";
            const $li = $(
              "<li class='list-group-item d-flex flex-wrap justify-content-between align-items-center gap-2' style='background:rgba(255,255,255,0.04);border-color:var(--border);color:var(--text);'></li>"
            );
            $li.append(
              $("<span></span>")
                .toggleClass("text-decoration-line-through text-muted", m.is_done)
                .text(m.title + " · " + due)
            );
            const $b = $("<div class='d-flex gap-1'></div>");
            $b.append(
              $("<button type='button' class='btn btn-sm btn-outline-secondary'>Toggle</button>").on("click", function () {
                $.ajax({
                  url: "/api/user/reminders/" + m.id,
                  method: "PATCH",
                  contentType: "application/json",
                  data: JSON.stringify({ is_done: !m.is_done }),
                })
                  .done(load)
                  .fail(function (xhr) {
                    showR((xhr.responseJSON && xhr.responseJSON.error) || "Failed");
                  });
              })
            );
            $b.append(
              $("<button type='button' class='btn btn-sm btn-outline-danger'>Delete</button>").on("click", function () {
                if (!confirm("Delete?")) return;
                $.ajax({ url: "/api/user/reminders/" + m.id, method: "DELETE" })
                  .done(load)
                  .fail(function (xhr) {
                    showR((xhr.responseJSON && xhr.responseJSON.error) || "Failed");
                  });
              })
            );
            $li.append($b);
            $u.append($li);
          });
        })
        .fail(function () {
          showR("Could not load.");
        });
    }
    $("#btnAddRem").on("click", function () {
      const title = ($("#rTitle").val() || "").trim();
      if (!title) {
        showR("Enter a title.");
        return;
      }
      const due = ($("#rDue").val() || "").trim();
      const body = { title: title };
      if (due) body.due_at = due.length === 16 ? due + ":00" : due;
      $.ajax({
        url: "/api/user/reminders",
        method: "POST",
        contentType: "application/json",
        data: JSON.stringify(body),
      })
        .done(function () {
          $("#rTitle").val("");
          $("#rDue").val("");
          load();
        })
        .fail(function (xhr) {
          showR((xhr.responseJSON && xhr.responseJSON.error) || "Failed");
        });
    });
    load();
    return;
  }

  if ($("#prefRoot").length) {
    function showP(msg) {
      $("#prefErr").text(msg).removeClass("d-none");
      $("#prefOk").addClass("d-none");
    }
    $.getJSON("/api/user/preferences")
      .done(function (d) {
        const pr = d.preferences || {};
        $("#prefTz").val(pr.timezone || "UTC");
        $("#prefWeek").val(pr.week_starts_on != null ? pr.week_starts_on : 0);
      })
      .fail(function () {
        showP("Could not load.");
      });
    $("#btnSavePref").on("click", function () {
      $("#prefErr").addClass("d-none");
      $.ajax({
        url: "/api/user/preferences",
        method: "PUT",
        contentType: "application/json",
        data: JSON.stringify({
          timezone: ($("#prefTz").val() || "UTC").trim() || "UTC",
          week_starts_on: parseInt($("#prefWeek").val(), 10) || 0,
        }),
      })
        .done(function () {
          $("#prefOk").removeClass("d-none");
          setTimeout(function () {
            $("#prefOk").addClass("d-none");
          }, 2000);
        })
        .fail(function (xhr) {
          showP((xhr.responseJSON && xhr.responseJSON.error) || "Save failed.");
        });
    });
  }
})();
