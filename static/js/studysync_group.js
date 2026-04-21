/* My Group — /api/groups */
(function () {
  const API = "/api/groups";

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

  let weekMonday = mondayOfLocal(new Date());
  let currentGroupId = null;
  let currentJoinCode = "";
  let lastMembers = [];

  const AVATAR_STYLES = [
    "linear-gradient(135deg,var(--accent),var(--accent2))",
    "linear-gradient(135deg,#ff6b35,#ffc94a)",
    "linear-gradient(135deg,#e040fb,#7c5cfc)",
    "linear-gradient(135deg,#26d07c,#4f8ef7)",
    "linear-gradient(135deg,#ff5555,#ffc94a)",
  ];

  const TYPE_ABBR = {
    lecture: "Lec",
    lab: "Lab",
    tutorial: "Tut",
    exam: "Exam",
    assignment: "Asg",
    workshop: "Wsh",
    other: "Evt",
  };

  const TYPE_BG = {
    lecture: "rgba(79,142,247,0.22)",
    lab: "rgba(255,107,53,0.22)",
    tutorial: "rgba(224,64,251,0.22)",
    exam: "rgba(255,85,85,0.22)",
    assignment: "rgba(38,208,124,0.2)",
    workshop: "rgba(255,201,74,0.2)",
    other: "rgba(255,255,255,0.08)",
  };

  const TYPE_DOT = {
    lecture: "var(--accent)",
    lab: "var(--orange)",
    tutorial: "var(--pink)",
    exam: "#ff5555",
    assignment: "var(--green)",
    workshop: "var(--yellow)",
    other: "var(--muted)",
  };

  function startOfDay(d) {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function mondayOfLocal(d) {
    const c = startOfDay(d);
    const diff = (c.getDay() + 6) % 7;
    c.setDate(c.getDate() - diff);
    return c;
  }

  function addDaysLocal(d, n) {
    const x = new Date(d.getTime());
    x.setDate(x.getDate() + n);
    return x;
  }

  function isoDateLocal(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  function hourLabel(h) {
    if (h === 12) return "12 PM";
    if (h > 12) return h - 12 + " PM";
    return h + " AM";
  }

  function showErr(msg) {
    $("#groupError").text(msg).removeClass("d-none");
    setTimeout(function () {
      $("#groupError").addClass("d-none");
    }, 6000);
  }

  function isoWeekParam() {
    return isoDateLocal(weekMonday);
  }

  function updateWeekTitle() {
    const fri = addDaysLocal(weekMonday, 4);
    const a = weekMonday.toLocaleDateString(undefined, { month: "short", day: "numeric" });
    const b = fri.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
    $("#mergedWeekTitle").text("Group timetable — " + a + " – " + b);
  }

  function markTodayColumns() {
    const today = startOfDay(new Date()).getTime();
    for (let i = 0; i < 5; i++) {
      const dd = addDaysLocal(weekMonday, i);
      const isT = startOfDay(dd).getTime() === today;
      $('.grid-header .grid-day[data-dow="' + i + '"]').toggleClass("today", isT);
    }
  }

  function eventInCell(ev, weekMon, wd, hourStart) {
    const d0 = startOfDay(addDaysLocal(weekMon, wd));
    const cs = new Date(d0.getFullYear(), d0.getMonth(), d0.getDate(), hourStart, 0, 0, 0);
    const ce = new Date(cs.getTime() + 3600000);
    const es = new Date(String(ev.start_at).replace(" ", "T"));
    const en = new Date(String(ev.end_at).replace(" ", "T"));
    if (Number.isNaN(es.getTime()) || Number.isNaN(en.getTime())) return false;
    return es < ce && en > cs;
  }

  function esc(s) {
    return $("<div/>").text(s).html();
  }

  function buildMergedGrid(events) {
    const H0 = 8;
    const H1 = 18;
    let html = "";
    for (let h = H0; h < H1; h++) {
      html += '<div class="grid-row"><div class="time-cell">' + hourLabel(h) + "</div>";
      for (let wd = 0; wd < 5; wd++) {
        const cellEvents = events.filter(function (ev) {
          return eventInCell(ev, weekMonday, wd, h);
        });
        const allFree = cellEvents.length === 0;
        html += '<div class="day-cell' + (allFree ? " free-all" : "") + '">';
        cellEvents.forEach(function (ev) {
          const ab = TYPE_ABBR[ev.event_type] || "Evt";
          const bg = TYPE_BG[ev.event_type] || TYPE_BG.other;
          const dot = TYPE_DOT[ev.event_type] || TYPE_DOT.other;
          const initials = ev.member_initials || "?";
          const label = initials + " · " + ab;
          html +=
            '<div class="mini-event" style="background:' +
            bg +
            ';color:var(--text)"><span class="mini-dot" style="background:' +
            dot +
            '"></span>' +
            esc(label) +
            "</div>";
        });
        html += "</div>";
      }
      html += "</div>";
    }
    $("#mergedGridBody").html(html);
  }

  function formatSlotPill(s) {
    const a = String(s.start).replace("T", " ");
    const b = String(s.end).replace("T", " ");
    return s.day + " · " + a.split(" ")[1] + " – " + b.split(" ")[1];
  }

  function renderFreeSlots(slots) {
    const $p = $("#freeSlotPills");
    $p.empty();
    if (!slots || !slots.length) {
      $p.append($("<span>", { class: "text-muted small", text: "No common 30-minute gaps this week in 8:00–20:00 (or not enough members)." }));
      return;
    }
    slots.forEach(function (s) {
      const $pill = $('<div class="slot-pill"></div>');
      $pill.append($("<i>", { class: "fas fa-clock" }));
      $pill.append(document.createTextNode(" " + formatSlotPill(s)));
      $p.append($pill);
    });
  }

  function renderMemberStack(members) {
    const $s = $("#memberStack");
    $s.empty();
    members.forEach(function (m, i) {
      $s.append(
        $("<div>", {
          class: "sm-avatar",
          text: m.initials,
          style: "background:" + AVATAR_STYLES[i % AVATAR_STYLES.length],
        })
      );
    });
  }

  function renderMemberList(members) {
    const $l = $("#memberList");
    $l.empty();
    members.forEach(function (m, i) {
      const $card = $('<div class="member-card"></div>');
      const $top = $('<div class="member-top"></div>');
      $top.append(
        $("<div>", {
          class: "mem-avatar",
          text: m.initials,
          style: "background:" + AVATAR_STYLES[i % AVATAR_STYLES.length],
        })
      );
      const $info = $("<div></div>");
      const $name = $('<div class="mem-name"></div>');
      $name.append(document.createTextNode(m.full_name));
      if (m.is_you) {
        $name.append(
          $('<span style="font-size:0.7rem;color:var(--accent);font-weight:600;margin-left:4px"></span>').text("(You)")
        );
      }
      $info.append($name);
      $info.append($("<div>", { class: "mem-id", text: m.student_id || String(m.user_id) }));
      $top.append($info);
      $top.append($('<div class="mem-meta">Member</div>'));
      $card.append($top);
      $card.append(
        $("<div>", {
          class: "mem-meta",
          text: "Calendar events this week: " + m.events_this_week,
        })
      );
      $l.append($card);
    });
  }

  function renderTasks(tasks) {
    const $l = $("#taskList");
    $l.empty();
    if (!tasks.length) {
      $l.append($("<div>", { class: "text-muted small", text: "No tasks yet." }));
      return;
    }
    tasks.forEach(function (t) {
      const $row = $('<div class="task-row"></div>');
      const ini = t.assignee_initials || "—";
      const grad = AVATAR_STYLES[(t.assignee_user_id || 0) % AVATAR_STYLES.length];
      $row.append(
        $("<div>", {
          class: "task-assignee",
          text: ini.length > 3 ? ini.slice(0, 3) : ini,
          style: "background:" + (t.assignee_user_id ? grad : "rgba(255,255,255,0.12)"),
        })
      );
      const $mid = $("<div></div>");
      $mid.append($("<div>", { class: "task-name", text: t.title }));
      const sub = t.assignee_name ? "Assigned to " + t.assignee_name : "Unassigned";
      $mid.append($("<div>", { class: "small text-muted", text: sub }));
      $row.append($mid);
      $row.append($("<div>", { class: "task-due", text: t.due_date ? "Due " + t.due_date : "" }));
      const $sel = $('<select class="form-select form-select-sm task-status-select"></select>');
      $sel.attr("data-task-id", t.id);
      ["not_started", "in_progress", "done"].forEach(function (st) {
        const $o = $("<option>", { value: st, text: st.replace("_", " ") });
        if (t.status === st) $o.prop("selected", true);
        $sel.append($o);
      });
      $row.append($sel);
      const $del = $('<button type="button" class="btn btn-sm btn-link text-danger p-0" title="Delete">&times;</button>');
      $del.on("click", function () {
        if (!confirm("Delete this task?")) return;
        $.ajax({ url: API + "/" + currentGroupId + "/tasks/" + t.id, method: "DELETE" })
          .done(function () {
            loadGroupData();
          })
          .fail(function (xhr) {
            showErr((xhr.responseJSON && xhr.responseJSON.error) || "Delete failed.");
          });
      });
      $row.append($del);
      $l.append($row);
    });

    $(".task-status-select")
      .off("change")
      .on("change", function () {
        const tid = $(this).data("task-id");
        const st = $(this).val();
        $.ajax({
          url: API + "/" + currentGroupId + "/tasks/" + tid,
          method: "PATCH",
          contentType: "application/json",
          data: JSON.stringify({ status: st }),
        }).fail(function (xhr) {
          showErr((xhr.responseJSON && xhr.responseJSON.error) || "Update failed.");
        });
      });
  }

  function fillAssigneeSelect(members) {
    const $s = $("#selectTaskAssignee");
    $s.empty();
    $s.append($("<option>", { value: "", text: "— Unassigned —" }));
    members.forEach(function (m) {
      $s.append($("<option>", { value: m.user_id, text: m.full_name }));
    });
  }

  function loadGroupData() {
    if (!currentGroupId) return;
    const ws = isoWeekParam();
    updateWeekTitle();
    markTodayColumns();
    $.when(
      $.getJSON(API + "/" + currentGroupId + "?week_start=" + encodeURIComponent(ws)),
      $.getJSON(API + "/" + currentGroupId + "/merged-timetable?week_start=" + encodeURIComponent(ws)),
      $.getJSON(API + "/" + currentGroupId + "/free-slots?week_start=" + encodeURIComponent(ws)),
      $.getJSON(API + "/" + currentGroupId + "/tasks")
    ).done(function (a, b, c, d) {
      const det = a[0];
      const merged = b[0];
      const free = c[0];
      const tasks = d[0];
      const g = det.group;
      currentJoinCode = g.join_code || "";
      $("#dispGroupName").text(g.name);
      $("#dispJoinCode").text(currentJoinCode);
      $("#dispMemberCount").text(String((det.members || []).length));
      lastMembers = det.members || [];
      renderMemberStack(lastMembers);
      renderMemberList(lastMembers);
      renderFreeSlots(free.slots || []);
      buildMergedGrid(merged.events || []);
      renderTasks(tasks.tasks || []);
      fillAssigneeSelect(lastMembers);
    }).fail(function (xhr) {
      showErr((xhr.responseJSON && xhr.responseJSON.error) || "Could not load group.");
    });
  }

  function showGroupUi(hasGroup) {
    $("#panelNoGroup").toggleClass("d-none", hasGroup);
    $("#panelGroup").toggleClass("d-none", !hasGroup);
    $("#btnInviteCode, #btnLeaveGroup").toggleClass("d-none", !hasGroup);
  }

  function applyMine(groups) {
    const $sel = $("#groupSelect");
    $sel.empty();
    if (!groups.length) {
      $sel.addClass("d-none");
      currentGroupId = null;
      showGroupUi(false);
      return;
    }
    showGroupUi(true);
    if (groups.length > 1) {
      $sel.removeClass("d-none");
      groups.forEach(function (g) {
        $sel.append($("<option>", { value: g.id, text: g.name }));
      });
      if (!currentGroupId || !groups.some(function (g) { return g.id === currentGroupId; })) {
        currentGroupId = groups[0].id;
      }
      $sel.val(String(currentGroupId));
    } else {
      $sel.addClass("d-none");
      currentGroupId = groups[0].id;
    }
    loadGroupData();
  }

  function loadMine() {
    $.getJSON(API + "/mine")
      .done(function (data) {
        applyMine(data.groups || []);
      })
      .fail(function (xhr) {
        showErr((xhr.responseJSON && xhr.responseJSON.error) || "Could not load groups.");
      });
  }

  $("#groupSelect").on("change", function () {
    currentGroupId = parseInt($(this).val(), 10);
    loadGroupData();
  });

  $("#btnPrevWeekG").on("click", function () {
    weekMonday = addDaysLocal(weekMonday, -7);
    loadGroupData();
  });
  $("#btnNextWeekG").on("click", function () {
    weekMonday = addDaysLocal(weekMonday, 7);
    loadGroupData();
  });

  $("#btnInviteCode").on("click", function () {
    if (!currentJoinCode) return;
    navigator.clipboard.writeText(currentJoinCode).then(
      function () {
        alert("Join code copied: " + currentJoinCode);
      },
      function () {
        prompt("Copy this join code:", currentJoinCode);
      }
    );
  });

  $("#btnLeaveGroup").on("click", function () {
    if (!currentGroupId || !confirm("Leave this study group?")) return;
    $.ajax({ url: API + "/" + currentGroupId + "/leave", method: "POST", contentType: "application/json", data: "{}" })
      .done(function () {
        currentGroupId = null;
        loadMine();
      })
      .fail(function (xhr) {
        showErr((xhr.responseJSON && xhr.responseJSON.error) || "Leave failed.");
      });
  });

  $("#btnCreateGroupSubmit").on("click", function () {
    const name = ($("#inputCreateName").val() || "").trim();
    $("#createGroupErr").addClass("d-none").text("");
    if (!name) {
      $("#createGroupErr").text("Name is required.").removeClass("d-none");
      return;
    }
    $.ajax({
      url: API + "/",
      method: "POST",
      contentType: "application/json",
      data: JSON.stringify({ name: name }),
    })
      .done(function () {
        bootstrap.Modal.getInstance(document.getElementById("modalCreateGroup")).hide();
        $("#inputCreateName").val("");
        loadMine();
      })
      .fail(function (xhr) {
        $("#createGroupErr")
          .text((xhr.responseJSON && xhr.responseJSON.error) || "Create failed.")
          .removeClass("d-none");
      });
  });

  $("#btnJoinGroupSubmit").on("click", function () {
    const code = ($("#inputJoinCode").val() || "").trim().toUpperCase();
    $("#joinGroupErr").addClass("d-none").text("");
    if (!code) {
      $("#joinGroupErr").text("Enter a join code.").removeClass("d-none");
      return;
    }
    $.ajax({
      url: API + "/join",
      method: "POST",
      contentType: "application/json",
      data: JSON.stringify({ join_code: code }),
    })
      .done(function () {
        bootstrap.Modal.getInstance(document.getElementById("modalJoinGroup")).hide();
        $("#inputJoinCode").val("");
        loadMine();
      })
      .fail(function (xhr) {
        $("#joinGroupErr")
          .text((xhr.responseJSON && xhr.responseJSON.error) || "Join failed.")
          .removeClass("d-none");
      });
  });

  $("#btnOpenAddTask").on("click", function () {
    $("#inputTaskTitle").val("");
    $("#inputTaskDue").val("");
    $("#selectTaskAssignee").val("");
    $("#addTaskErr").addClass("d-none").text("");
    fillAssigneeSelect(lastMembers);
    new bootstrap.Modal("#modalAddTask").show();
  });

  $("#btnAddTaskSubmit").on("click", function () {
    const title = ($("#inputTaskTitle").val() || "").trim();
    $("#addTaskErr").addClass("d-none").text("");
    if (!title) {
      $("#addTaskErr").text("Title is required.").removeClass("d-none");
      return;
    }
    const payload = { title: title };
    const asg = $("#selectTaskAssignee").val();
    if (asg) payload.assignee_user_id = parseInt(asg, 10);
    const due = $("#inputTaskDue").val();
    if (due) payload.due_date = due;
    $.ajax({
      url: API + "/" + currentGroupId + "/tasks",
      method: "POST",
      contentType: "application/json",
      data: JSON.stringify(payload),
    })
      .done(function () {
        bootstrap.Modal.getInstance(document.getElementById("modalAddTask")).hide();
        loadGroupData();
      })
      .fail(function (xhr) {
        $("#addTaskErr")
          .text((xhr.responseJSON && xhr.responseJSON.error) || "Save failed.")
          .removeClass("d-none");
      });
  });

  loadMine();
})();
