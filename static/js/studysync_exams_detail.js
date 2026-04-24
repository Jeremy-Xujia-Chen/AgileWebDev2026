(function () {
  function csrf() {
    return $("meta[name=csrf-token]").attr("content");
  }

  $.ajaxSetup({
    beforeSend: function (xhr, settings) {
      if (
        !/^(GET|HEAD|OPTIONS|TRACE)$/i.test(settings.type) &&
        !settings.crossDomain
      ) {
        xhr.setRequestHeader("X-CSRFToken", csrf());
      }
    },
  });

  const $root = $("#examDetailRoot");
  const examId = $root.data("exam-id");
  const startsRaw = $root.data("starts-at");

  function showErr(msg) {
    $("#examDetailError").text(msg).removeClass("d-none");
    setTimeout(function () {
      $("#examDetailError").addClass("d-none");
    }, 5000);
  }

  function updateCountdown() {
    if (!startsRaw) return;
    const target = new Date(startsRaw);
    const now = new Date();
    let ms = target.getTime() - now.getTime();
    if (ms < 0) ms = 0;
    const sec = Math.floor(ms / 1000);
    const days = Math.floor(sec / 86400);
    const hours = Math.floor((sec % 86400) / 3600);
    const mins = Math.floor((sec % 3600) / 60);
    $("#cdDays").text(days);
    $("#cdHours").text(hours);
    $("#cdMins").text(mins);
  }

  function patchTopic(topicId, body) {
    return $.ajax({
      url: "/api/exams/topics/" + topicId,
      method: "PATCH",
      contentType: "application/json",
      data: JSON.stringify(body),
    });
  }

  $(".checklist-item .check-box").on("click keypress", function (e) {
    if (e.type === "keypress" && e.which !== 13 && e.which !== 32) return;
    e.preventDefault();
    const $row = $(this).closest(".checklist-item");
    const id = $row.data("topic-id");
    const cur = $row.data("status");
    const next = cur === "done" ? "not_started" : "done";
    patchTopic(id, { status: next })
      .done(function (res) {
        const st = res.topic.status;
        $row.data("status", st);
        $row.toggleClass("done", st === "done");
        const $cb = $row.find(".check-box");
        $cb.attr("aria-checked", st === "done" ? "true" : "false");
        $cb.html(st === "done" ? '<i class="fas fa-check"></i>' : "");
      })
      .fail(function (xhr) {
        showErr(
          (xhr.responseJSON && xhr.responseJSON.error) ||
            "Could not update topic.",
        );
      });
  });

  let notesTimer;
  $("#examNotes").on("input", function () {
    clearTimeout(notesTimer);
    const val = $(this).val();
    notesTimer = setTimeout(function () {
      $.ajax({
        url: "/api/exams/sessions/" + examId,
        method: "PATCH",
        contentType: "application/json",
        data: JSON.stringify({ notes: val }),
      }).fail(function (xhr) {
        showErr(
          (xhr.responseJSON && xhr.responseJSON.error) ||
            "Could not save notes.",
        );
      });
    }, 800);
  });

  $(".topic-progress").on("change", function () {
    const $row = $(this).closest(".checklist-item");
    const id = $row.data("topic-id");
    const v = parseInt($(this).val(), 10);
    patchTopic(id, { progress_percent: v }).fail(function (xhr) {
      showErr(
        (xhr.responseJSON && xhr.responseJSON.error) ||
          "Could not update progress.",
      );
    });
  });

  $("#btnAddTopic").on("click", function () {
    const label = window.prompt("Topic label");
    if (!label || !label.trim()) return;
    $.ajax({
      url: "/api/exams/sessions/" + examId + "/topics",
      method: "POST",
      contentType: "application/json",
      data: JSON.stringify({ label: label.trim() }),
    })
      .done(function () {
        window.location.reload();
      })
      .fail(function (xhr) {
        showErr(
          (xhr.responseJSON && xhr.responseJSON.error) ||
            "Could not add topic.",
        );
      });
  });

  $("#btnDeleteExam").on("click", function () {
    if (!window.confirm("Delete this exam and all revision topics?")) return;
    $.ajax({ url: "/api/exams/sessions/" + examId, method: "DELETE" })
      .done(function () {
        window.location.href = "/exams";
      })
      .fail(function (xhr) {
        showErr(
          (xhr.responseJSON && xhr.responseJSON.error) || "Delete failed.",
        );
      });
  });

  updateCountdown();
  setInterval(updateCountdown, 60000);
})();
