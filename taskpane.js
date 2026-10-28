/* Subject Helper – Outlook glue code */
(function () {
  "use strict";
  const C = window.SubjectCore;
  const $ = (id) => document.getElementById(id);
  const DEFAULT_MODELS = { claude: "claude-haiku-4-5", openai: "gpt-4.1-mini" };
  const STORE_KEY = "subjectHelper.settings";

  // ---------- settings (stored only on this computer) ----------
  function loadSettings() {
    try {
      return Object.assign({ provider: "rules", apiKey: "", model: "", extraWords: "" },
        JSON.parse(localStorage.getItem(STORE_KEY) || "{}"));
    } catch (e) {
      return { provider: "rules", apiKey: "", model: "", extraWords: "" };
    }
  }
  function saveSettings(s) {
    try { localStorage.setItem(STORE_KEY, JSON.stringify(s)); return true; } catch (e) { return false; }
  }
  let settings = loadSettings();

  function fillSettingsForm() {
    $("provider").value = settings.provider;
    $("apiKey").value = settings.apiKey;
    $("model").value = settings.model || DEFAULT_MODELS[settings.provider] || "";
    $("extraWords").value = settings.extraWords;
    $("aiFields").hidden = settings.provider === "rules";
    updateModeLabel();
  }
  function updateModeLabel() {
    const names = { rules: "Rules only", claude: "Claude", openai: "OpenAI" };
    const noKey = settings.provider !== "rules" && !settings.apiKey;
    $("modeLabel").textContent = "Engine: " + names[settings.provider] +
      (noKey ? " (no key saved, using rules only)" : "");
  }

  // ---------- Office helpers (promise wrappers) ----------
  function getAsyncP(fn) {
    return new Promise((resolve, reject) =>
      fn((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))));
  }
  const item = () => Office.context.mailbox.item;
  const getBody = () => getAsyncP((cb) => item().body.getAsync(Office.CoercionType.Text, cb));
  const getSubject = () => getAsyncP((cb) => item().subject.getAsync(cb));
  const setSubject = (s) => getAsyncP((cb) => item().subject.setAsync(s, cb));
  const getRecips = (field) =>
    item()[field] ? getAsyncP((cb) => item()[field].getAsync(cb)).catch(() => []) : Promise.resolve([]);

  // ---------- AI calls ----------
  async function askClaude(prompt) {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": settings.apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: settings.model || DEFAULT_MODELS.claude,
        max_tokens: 200,
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error((data.error && data.error.message) || "Claude API error " + res.status);
    return (data.content || []).map((c) => c.text || "").join("");
  }

  async function askOpenAI(prompt) {
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: { "content-type": "application/json", authorization: "Bearer " + settings.apiKey },
      body: JSON.stringify({
        model: settings.model || DEFAULT_MODELS.openai,
        max_tokens: 200,
        messages: [{ role: "user", content: prompt }],
      }),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error((data.error && data.error.message) || "OpenAI API error " + res.status);
    return data.choices && data.choices[0] && data.choices[0].message ? data.choices[0].message.content : "";
  }

  // ---------- UI ----------
  function setStatus(msg, isError) {
    const el = $("status");
    el.textContent = msg || "";
    el.classList.toggle("error", !!isError);
  }

  function showResults(list, note) {
    const ul = $("results");
    ul.innerHTML = "";
    list.forEach((s) => {
      const li = document.createElement("li");
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = s;
      const tag = document.createElement("span");
      tag.className = "tag";
      tag.textContent = "Click to use";
      b.appendChild(tag);
      b.addEventListener("click", async () => {
        try {
          await setSubject(s);
          ul.querySelectorAll("button").forEach((x) => {
            x.classList.remove("used");
            x.querySelector(".tag").textContent = "Click to use";
          });
          b.classList.add("used");
          tag.textContent = "Added as subject ✓";
        } catch (e) {
          setStatus("Couldn't set the subject: " + (e.message || e), true);
        }
      });
      li.appendChild(b);
      ul.appendChild(li);
    });
    setStatus(note || "");
  }

  async function suggest() {
    const btn = $("suggestBtn");
    btn.disabled = true;
    $("results").innerHTML = "";
    $("previewBox").hidden = true;
    setStatus("Reading your email…");
    try {
      const me = Office.context.mailbox.userProfile || {};
      const [body, subject, to, cc, bcc] = await Promise.all([
        getBody(), getSubject().catch(() => ""), getRecips("to"), getRecips("cc"), getRecips("bcc"),
      ]);

      const cleaned = C.cleanBody(body, me.displayName);
      if (cleaned.replace(/\s/g, "").length < 15) {
        setStatus("Write a sentence or two first, then click again.");
        return;
      }

      const names = [me.displayName, me.emailAddress];
      [...to, ...cc, ...bcc].forEach((r) => names.push(r.displayName, r.emailAddress));
      const extra = (settings.extraWords || "").split(",").map((w) => w.trim()).filter(Boolean);
      const safe = C.redact(cleaned, { names, extraWords: extra });
      const isReply = /^\s*(re|fw|fwd)\s*:/i.test(subject || "");

      const useAI = settings.provider !== "rules" && settings.apiKey;
      if (!useAI) {
        showResults(C.rulesSubjects(safe), "Suggestions from rules (no AI).");
        return;
      }

      setStatus("Getting suggestions…");
      $("previewText").textContent = safe;
      $("previewBox").hidden = false;
      const raw = settings.provider === "claude"
        ? await askClaude(C.buildPrompt(safe, isReply))
        : await askOpenAI(C.buildPrompt(safe, isReply));
      let list = C.parseSubjects(raw);
      if (isReply) list = list.map((s) => (/^re:/i.test(s) ? s : "Re: " + s));
      if (!list.length) throw new Error("The AI didn't return any subjects.");
      showResults(list);
    } catch (e) {
      // fall back to rules so the button always gives something
      setStatus("AI didn't respond (" + (e.message || e) + "). Showing rule-based ideas instead.", true);
      try {
        const body = await getBody();
        const me = Office.context.mailbox.userProfile || {};
        const safe = C.redact(C.cleanBody(body, me.displayName), { names: [me.displayName] });
        const ul = $("results");
        if (!ul.children.length) {
          const note = $("status").textContent;
          showResults(C.rulesSubjects(safe), note);
          $("status").classList.add("error");
        }
      } catch (_) { /* nothing else to do */ }
    } finally {
      btn.disabled = false;
    }
  }

  // ---------- wire up ----------
  Office.onReady(() => {
    fillSettingsForm();

    $("suggestBtn").addEventListener("click", suggest);

    $("settingsToggle").addEventListener("click", () => {
      const open = $("settingsView").hidden;
      $("settingsView").hidden = !open;
      $("mainView").hidden = open;
      $("settingsToggle").textContent = open ? "Back" : "Settings";
      $("settingsToggle").setAttribute("aria-expanded", String(open));
      $("settingsStatus").textContent = "";
    });

    $("provider").addEventListener("change", () => {
      const p = $("provider").value;
      $("aiFields").hidden = p === "rules";
      if (!$("model").value || Object.values(DEFAULT_MODELS).includes($("model").value)) {
        $("model").value = DEFAULT_MODELS[p] || "";
      }
    });

    $("saveSettings").addEventListener("click", () => {
      settings = {
        provider: $("provider").value,
        apiKey: $("apiKey").value.trim(),
        model: $("model").value.trim(),
        extraWords: $("extraWords").value,
      };
      const ok = saveSettings(settings);
      updateModeLabel();
      $("settingsStatus").textContent = ok ? "Saved." : "Couldn't save on this computer (storage blocked).";
    });
  });
})();
