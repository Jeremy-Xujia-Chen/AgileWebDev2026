/* StudySync AI Planner — POST /api/planner/chat (stub or real backend) */
(function () {
  const $root = $("#aiPlannerRoot");
  const chatUrl = $root.data("chat-url");
  const userInitials = String($root.data("user-initials") || "U").slice(0, 4);
  const userName = String($root.data("user-name") || "User");

  const PLAN_COLORS = ["#4f8ef7", "#ffc94a", "#7c5cfc", "#26d07c", "#ff5555"];

  let conversationId = null;
  const HIST_KEY = "studysync_ai_roundtrip";

  function readHistory() {
    try {
      const s = sessionStorage.getItem(HIST_KEY);
      return s ? JSON.parse(s) : [];
    } catch (e) {
      return [];
    }
  }

  function saveHistoryItem(role, text) {
    const arr = readHistory();
    arr.push({
      at: new Date().toISOString(),
      role: role,
      text: String(text).slice(0, 2000),
    });
    while (arr.length > 50) arr.shift();
    try {
      sessionStorage.setItem(HIST_KEY, JSON.stringify(arr));
    } catch (e) {}
  }

  $("#aiHistoryModal").on("show.bs.modal", function () {
    const arr = readHistory();
    $("#aiHistoryPre").text(
      arr.length
        ? arr.map(function (x) {
            return "[" + x.at + "] " + x.role + ": " + x.text;
          }).join("\n\n")
        : "(No messages yet this session.)"
    );
  });

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

  function escapeHtml(s) {
    return $("<div>").text(s).html();
  }

  function formatNowLabel() {
    const d = new Date();
    return d.toLocaleString(undefined, {
      hour: "numeric",
      minute: "2-digit",
      month: "short",
      day: "numeric",
    });
  }

  function scrollToBottom() {
    const el = document.getElementById("chatBody");
    if (el) el.scrollTop = el.scrollHeight;
  }

  function buildPlanCard(blocks) {
    if (!blocks || !blocks.length) return null;
    const $card = $("<div>", { class: "plan-card" });
    $card.append(
      $("<div>", { class: "plan-card-header" }).append(
        $("<i>", { class: "fas fa-calendar-check" }),
        document.createTextNode(" Suggested plan"),
      ),
    );
    blocks.forEach(function (b, i) {
      const start = b.start_label != null ? String(b.start_label) : "";
      const end = b.end_label != null ? String(b.end_label) : "";
      const timeLabel = start && end ? start + " – " + end : start || end || "";
      const title = b.title != null ? String(b.title) : "";
      const dm = b.duration_minutes;
      const durLabel = dm != null && dm !== "" ? String(dm) + " min" : "";

      const $row = $("<div>", { class: "plan-row" });
      $row.append(
        $("<div>", { class: "plan-dot" }).css(
          "background",
          PLAN_COLORS[i % PLAN_COLORS.length],
        ),
      );
      $row.append($("<div>", { class: "plan-time", text: timeLabel }));
      $row.append($("<div>", { class: "plan-subject", text: title }));
      $row.append($("<div>", { class: "plan-duration", text: durLabel }));
      $card.append($row);
    });
    return $card;
  }

  function appendUserMessage(text) {
    const $msg = $("<div>", { class: "msg user" });
    $msg.append(
      $("<div>", {
        class: "msg-avatar",
        style: "background:linear-gradient(135deg,#ff6b35,#ffc94a)",
        text: userInitials,
      }),
    );
    const $col = $("<div>");
    $col.append($("<div>", { class: "msg-bubble", text: text }));
    $col.append($("<div>", { class: "msg-time", text: formatNowLabel() }));
    $msg.append($col);
    $("#typingIndicator").before($msg);
  }

  function appendAssistantMessage(replyText, planBlocks) {
    const $msg = $("<div>", { class: "msg" });
    $msg.append(
      $("<div>", {
        class: "msg-avatar",
        style:
          "background:linear-gradient(135deg,var(--accent),var(--accent2))",
      }).append($("<i>", { class: "fas fa-robot" })),
    );
    const $col = $("<div>");
    const $bubble = $("<div>", { class: "msg-bubble" });
    const safe = escapeHtml(replyText || "").replace(/\n/g, "<br/>");
    $bubble.append($("<div>", { class: "ai-msg-text" }).html(safe));
    const $plan = buildPlanCard(planBlocks);
    if ($plan) $bubble.append($plan);
    $col.append($bubble);
    $col.append($("<div>", { class: "msg-time", text: formatNowLabel() }));
    $msg.append($col);
    $("#typingIndicator").before($msg);
  }

  function appendErrorMessage(msg) {
    appendAssistantMessage(
      "Something went wrong: " +
        (msg || "Unknown error") +
        "\n\nPlease try again.",
      [],
    );
  }

  function setTyping(on) {
    $("#typingIndicator").css("display", on ? "flex" : "none");
    scrollToBottom();
  }

  function sendChat(text) {
    const trimmed = (text || "").trim();
    if (!trimmed) return;

    saveHistoryItem("you", trimmed);
    appendUserMessage(trimmed);
    $("#chatInput").val("").css("height", "auto");
    setTyping(true);
    $("#btnSendChat").prop("disabled", true);

    $.ajax({
      url: chatUrl,
      method: "POST",
      contentType: "application/json",
      data: JSON.stringify({
        message: trimmed,
        conversation_id: conversationId,
        context: {
          user_name: userName,
          locale: navigator.language || "en",
        },
      }),
    })
      .done(function (data) {
        setTyping(false);
        if (!data || data.ok === false) {
          appendErrorMessage((data && data.error) || "Request failed.");
          return;
        }
        if (data.conversation_id != null) conversationId = data.conversation_id;
        saveHistoryItem("assistant", (data && data.reply_text) || "");
        appendAssistantMessage(data.reply_text || "", data.plan_blocks || []);
        scrollToBottom();
      })
      .fail(function (xhr) {
        setTyping(false);
        const j = xhr.responseJSON;
        appendErrorMessage(
          (j && (j.error || j.message)) || xhr.statusText || "Network error",
        );
        scrollToBottom();
      })
      .always(function () {
        $("#btnSendChat").prop("disabled", false);
        scrollToBottom();
      });
  }

  function clearLocalChat() {
    conversationId = null;
    try {
      sessionStorage.removeItem(HIST_KEY);
    } catch (e) {}
    $("#chatBody .msg").each(function () {
      const id = this.id;
      if (id !== "typingIndicator") $(this).remove();
    });
    setTyping(false);
    const $welcome = $(`
      <div class="msg" id="aiWelcomeMsg">
        <div class="msg-avatar" style="background:linear-gradient(135deg,var(--accent),var(--accent2))"><i class="fas fa-robot"></i></div>
        <div>
          <div class="msg-bubble">Chat cleared. Ask anything when you are ready — I still use the stub backend until your teammate connects a model.</div>
          <div class="msg-time">Just now</div>
        </div>
      </div>`);
    $("#typingIndicator").before($welcome);
    scrollToBottom();
  }

  $("#btnSendChat").on("click", function () {
    sendChat($("#chatInput").val());
  });

  $("#chatInput").on("keydown", function (e) {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendChat($(this).val());
    }
  });

  $("#chatInput").on("input", function () {
    this.style.height = "auto";
    this.style.height = this.scrollHeight + "px";
  });

  $(".quick-prompts .quick-btn").on("click", function () {
    const q = $(this).data("quick");
    if (q) sendChat(String(q));
  });

  $("#btnAiClear").on("click", clearLocalChat);

  scrollToBottom();
})();
