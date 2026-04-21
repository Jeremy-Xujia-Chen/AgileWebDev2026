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

  function fmtRange(s) {
    if (!s.starts_at || !s.ends_at) return "";
    return s.starts_at.replace("T", " ") + " → " + s.ends_at.replace("T", " ");
  }

  function render(sessions) {
    const $g = $("#examGrid");
    const $empty = $("#examEmpty");
    $g.empty();
    if (!sessions.length) {
      $empty.removeClass("d-none");
      return;
    }
    $empty.addClass("d-none");
    sessions.forEach(function (s) {
      const href = "/exams/" + s.id;
      const $a = $("<a>", { href: href, class: "exam-card" });
      $a.append($("<div>", { class: "exam-card-title", text: s.title }));
      const $meta = $("<div>", { class: "exam-card-meta" });
      if (s.course_code) {
        $meta.append($("<span>", { class: "text-muted", text: s.course_code }));
        $meta.append(document.createTextNode(" · "));
      }
      $meta.append(document.createTextNode(fmtRange(s)));
      $a.append($meta);
      $g.append($a);
    });
  }

  function load() {
    $("#examLoadError").addClass("d-none");
    $.getJSON("/api/exams/sessions")
      .done(function (data) {
        render(data.sessions || []);
      })
      .fail(function (xhr) {
        const msg = (xhr.responseJSON && xhr.responseJSON.error) || "Could not load exams.";
        $("#examLoadError").text(msg).removeClass("d-none");
      });
  }

  function localToIso(val) {
    if (!val) return null;
    const d = new Date(val);
    if (Number.isNaN(d.getTime())) return null;
    const pad = (n) => String(n).padStart(2, "0");
    return (
      d.getFullYear() +
      "-" +
      pad(d.getMonth() + 1) +
      "-" +
      pad(d.getDate()) +
      "T" +
      pad(d.getHours()) +
      ":" +
      pad(d.getMinutes()) +
      ":00"
    );
  }

  $("#btnNewExam").on("click", function () {
    $("#examCreateFormError").addClass("d-none").text("");
    $("#examCreateForm")[0].reset();
    new bootstrap.Modal("#examCreateModal").show();
  });

  $("#btnSaveExam").on("click", function () {
    const $f = $("#examCreateForm");
    const title = ($f.find('[name="title"]').val() || "").trim();
    const starts = localToIso($f.find('[name="starts_at"]').val());
    const ends = localToIso($f.find('[name="ends_at"]').val());
    if (!title || !starts || !ends) {
      $("#examCreateFormError").text("Title, start and end are required.").removeClass("d-none");
      return;
    }
    const payload = {
      title: title,
      course_code: ($f.find('[name="course_code"]').val() || "").trim(),
      starts_at: starts,
      ends_at: ends,
      location: ($f.find('[name="location"]').val() || "").trim(),
      notes: ($f.find('[name="notes"]').val() || "").trim(),
    };
    const wt = $f.find('[name="weight_percent"]').val();
    if (wt !== "") payload.weight_percent = parseFloat(wt);
    $("#examCreateFormError").addClass("d-none").text("");
    $.ajax({ url: "/api/exams/sessions", method: "POST", contentType: "application/json", data: JSON.stringify(payload) })
      .done(function (res) {
        bootstrap.Modal.getInstance(document.getElementById("examCreateModal")).hide();
        window.location.href = "/exams/" + res.session.id;
      })
      .fail(function (xhr) {
        const msg = (xhr.responseJSON && xhr.responseJSON.error) || "Save failed.";
        $("#examCreateFormError").text(msg).removeClass("d-none");
      });
  });

  $(load);
})();
