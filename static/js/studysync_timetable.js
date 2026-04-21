/* StudySync timetable: week view + AJAX against /api/timetable/events */
(function () {
  const START_HOUR = 8;
  const END_HOUR = 16; // exclusive (last slot starts at 3 PM when END=16)
  const SLOT_MIN = 80;
  const HOURS = [];
  for (let h = START_HOUR; h < END_HOUR; h++) HOURS.push(h);

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

  function toDatetimeLocal(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return y + "-" + m + "-" + day + "T" + hh + ":" + mm;
  }

  function toDatetimeLocalFromIso(iso) {
    if (!iso) return "";
    const s = String(iso).replace(" ", "T");
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return "";
    return toDatetimeLocal(d);
  }

  function fromDatetimeLocalValue(v) {
    if (!v) return null;
    return v.length === 16 ? v + ":00" : v;
  }

  function dayIndexInWeek(evStart, weekMon) {
    const a = startOfDay(evStart).getTime() - startOfDay(weekMon).getTime();
    return Math.round(a / 86400000);
  }

  function typeLabel(t) {
    const map = {
      lecture: "Lecture",
      lab: "Laboratory",
      tutorial: "Tutorial",
      exam: "Exam",
      assignment: "Assignment",
      workshop: "Workshop",
      other: "Event",
    };
    return map[t] || "Event";
  }

  function evClass(t) {
    if (["lecture", "lab", "tutorial", "exam", "assignment", "workshop", "other"].includes(t)) return "ev-" + t;
    return "ev-other";
  }

  let currentMonday = mondayOfLocal(new Date());
  let lastEvents = [];
  const hiddenTypes = new Set();

  function weekRangeLabel() {
    const fri = addDaysLocal(currentMonday, 4);
    const opts = { month: "short", day: "numeric" };
    const a = currentMonday.toLocaleDateString(undefined, opts);
    const b = fri.toLocaleDateString(undefined, { ...opts, year: "numeric" });
    $("#tt-date-range").text(a + " – " + b);
  }

  function buildWeekHeader() {
    let html = '<div class="day-header" style="background:rgba(255,255,255,0.02);border-bottom:1px solid var(--border);"></div>';
    const today = startOfDay(new Date()).getTime();
    for (let i = 0; i < 5; i++) {
      const d = addDaysLocal(currentMonday, i);
      const isToday = startOfDay(d).getTime() === today;
      const cls = isToday ? "day-header today" : "day-header";
      html +=
        '<div class="' +
        cls +
        '"><span>' +
        ["Mon", "Tue", "Wed", "Thu", "Fri"][i] +
        '</span><span class="date-num">' +
        d.getDate() +
        "</span></div>";
    }
    $("#tt-week-header").html(html);
  }

  function buildRulerAndPanels() {
    let ruler = "";
    HOURS.forEach(function (h) {
      ruler += '<div class="tt-hour-line">' + hourLabel(h) + "</div>";
    });
    $("#tt-time-ruler").html(ruler);

    let panels = "";
    for (let d = 0; d < 5; d++) {
      panels +=
        '<div class="tt-day-panel" data-day="' +
        d +
        '"><div class="tt-events-layer" id="tt-layer-' +
        d +
        '"></div></div>';
    }
    $("#tt-day-panels").html(panels);
  }

  function minutesSinceStartHour(dt) {
    const base = new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), START_HOUR, 0, 0, 0);
    return (dt - base) / 60000;
  }

  function clearLayers() {
    for (let d = 0; d < 5; d++) $("#tt-layer-" + d).empty();
  }

  function applyFilters() {
    $(".tt-events-layer .event").each(function () {
      const t = $(this).data("etype");
      if (hiddenTypes.has(t)) $(this).hide();
      else $(this).show();
    });
  }

  function placeEvent(ev) {
    const start = new Date(String(ev.start_at).replace(" ", "T"));
    const end = new Date(String(ev.end_at).replace(" ", "T"));
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return;

    const idx = dayIndexInWeek(start, currentMonday);
    if (idx < 0 || idx > 4) return;

    const mins = minutesSinceStartHour(start);
    const durMin = Math.max(5, (end - start) / 60000);
    const top = (mins / 60) * SLOT_MIN;
    const height = (durMin / 60) * SLOT_MIN;
    const maxTop = HOURS.length * SLOT_MIN;
    const clampedTop = Math.max(0, Math.min(top, maxTop - 10));
    const clampedHeight = Math.max(24, Math.min(height, maxTop - clampedTop));

    const loc = ev.location
      ? '<div class="event-loc"><i class="fas fa-map-marker-alt"></i> ' + $("<div/>").text(ev.location).html() + "</div>"
      : "";

    const el = $(
      '<div class="event ' +
        evClass(ev.event_type) +
        '" data-id="' +
        ev.id +
        '" data-etype="' +
        ev.event_type +
        '" style="top:' +
        clampedTop +
        "px;height:" +
        clampedHeight +
        'px;">' +
        '<div class="event-title">' +
        $("<div/>").text(ev.title).html() +
        "</div>" +
        '<div class="event-sub">' +
        $("<div/>").text(typeLabel(ev.event_type)).html() +
        "</div>" +
        loc +
        "</div>"
    );
    el.on("click", function () {
      openEdit(ev);
    });
    $("#tt-layer-" + idx).append(el);
  }

  function renderEvents(events) {
    lastEvents = events || [];
    clearLayers();
    lastEvents.forEach(placeEvent);
    applyFilters();
  }

  function showLoadError(msg) {
    $("#loadError").removeClass("d-none").text(msg);
  }

  function hideLoadError() {
    $("#loadError").addClass("d-none").text("");
  }

  function loadWeek() {
    hideLoadError();
    weekRangeLabel();
    buildWeekHeader();
    buildRulerAndPanels();
    const ws = isoDateLocal(currentMonday);
    $.getJSON("/api/timetable/events", { week_start: ws })
      .done(function (data) {
        if (data.week_start) {
          const p = String(data.week_start).split("-");
          currentMonday = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 0, 0, 0, 0);
          weekRangeLabel();
          buildWeekHeader();
          buildRulerAndPanels();
        }
        renderEvents(data.events || []);
      })
      .fail(function (xhr) {
        let msg = "Could not load events.";
        if (xhr.responseJSON && xhr.responseJSON.error) msg = xhr.responseJSON.error;
        showLoadError(msg);
      });
  }

  const modalEl = document.getElementById("eventModal");
  const bsModal = modalEl ? new bootstrap.Modal(modalEl) : null;

  function openCreate() {
    $("#eventModalTitle").text("Create event");
    $("#evId").val("");
    $("#evTitle").val("");
    $("#evType").val("lecture");
    $("#evLocation").val("");
    $("#evNotes").val("");
    const d0 = addDaysLocal(currentMonday, 0);
    d0.setHours(9, 0, 0, 0);
    const d1 = addDaysLocal(currentMonday, 0);
    d1.setHours(10, 0, 0, 0);
    $("#evStart").val(toDatetimeLocal(d0));
    $("#evEnd").val(toDatetimeLocal(d1));
    $("#btnDeleteEvent").addClass("d-none");
    bsModal.show();
  }

  function openEdit(ev) {
    $("#eventModalTitle").text("Edit event");
    $("#evId").val(String(ev.id));
    $("#evTitle").val(ev.title);
    $("#evType").val(ev.event_type);
    $("#evStart").val(toDatetimeLocalFromIso(ev.start_at));
    $("#evEnd").val(toDatetimeLocalFromIso(ev.end_at));
    $("#evLocation").val(ev.location || "");
    $("#evNotes").val(ev.notes || "");
    $("#btnDeleteEvent").removeClass("d-none");
    bsModal.show();
  }

  function saveEvent() {
    const id = $("#evId").val();
    const payload = {
      title: $("#evTitle").val().trim(),
      event_type: $("#evType").val(),
      start_at: fromDatetimeLocalValue($("#evStart").val()),
      end_at: fromDatetimeLocalValue($("#evEnd").val()),
      location: $("#evLocation").val().trim(),
      notes: $("#evNotes").val().trim(),
    };
    if (!payload.title) {
      alert("Title is required.");
      return;
    }
    const url = id ? "/api/timetable/events/" + id : "/api/timetable/events";
    const method = id ? "PATCH" : "POST";
    $.ajax({ url: url, method: method, contentType: "application/json", data: JSON.stringify(payload) })
      .done(function () {
        bsModal.hide();
        loadWeek();
      })
      .fail(function (xhr) {
        let msg = "Save failed.";
        if (xhr.responseJSON && xhr.responseJSON.error) msg = xhr.responseJSON.error;
        alert(msg);
      });
  }

  function deleteEvent() {
    const id = $("#evId").val();
    if (!id || !confirm("Delete this event?")) return;
    $.ajax({ url: "/api/timetable/events/" + id, method: "DELETE" })
      .done(function () {
        bsModal.hide();
        loadWeek();
      })
      .fail(function (xhr) {
        let msg = "Delete failed.";
        if (xhr.responseJSON && xhr.responseJSON.error) msg = xhr.responseJSON.error;
        alert(msg);
      });
  }

  $("#btnPrevWeek").on("click", function () {
    currentMonday = addDaysLocal(currentMonday, -7);
    loadWeek();
  });
  $("#btnNextWeek").on("click", function () {
    currentMonday = addDaysLocal(currentMonday, 7);
    loadWeek();
  });
  $("#btnCreateEvent").on("click", openCreate);
  $("#btnSaveEvent").on("click", saveEvent);
  $("#btnDeleteEvent").on("click", deleteEvent);

  $(".filter-tag").on("click", function () {
    const t = $(this).data("type");
    if (!t) return;
    if (hiddenTypes.has(t)) {
      hiddenTypes.delete(t);
      $(this).removeClass("dim");
    } else {
      hiddenTypes.add(t);
      $(this).addClass("dim");
    }
    applyFilters();
  });

  loadWeek();
})();
