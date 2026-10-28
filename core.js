/* Subject Helper – pure logic (no Office APIs). Works in the browser and in Node for testing. */
(function (root) {
  "use strict";

  // ---------- 1. Clean the body: drop quoted history + signature ----------
  function cleanBody(text, senderName) {
    if (!text) return "";
    let t = text.replace(/\r\n?/g, "\n");

    // Cut quoted reply/forward history
    const historyMarkers = [
      /\n_{10,}\s*\n\s*From:/i,
      /\n-{2,}\s*Original Message\s*-{2,}/i,
      /\nFrom:\s.+\n\s*Sent:\s/i,
      /\nOn .{5,120} wrote:\s*\n/i,
    ];
    for (const re of historyMarkers) {
      const m = t.search(re);
      if (m >= 0) t = t.slice(0, m);
    }

    // Cut signature / confidentiality notice
    const sigMarkers = [/\nCONFIDENTIALITY NOTICE/i, /\n--\s*\n/, /\n_{10,}/];
    if (senderName) {
      const esc = senderName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      sigMarkers.unshift(new RegExp("\\n\\s*" + esc + "\\s*\\n", "i"));
    }
    for (const re of sigMarkers) {
      const m = ("\n" + t).search(re);
      if (m >= 0) t = ("\n" + t).slice(0, m).replace(/^\n/, "");
    }

    return t.replace(/\n{3,}/g, "\n\n").trim();
  }

  // ---------- 2. Redact names, numbers and other identifiers ----------
  const STOP_NAME_WORDS = new Set(["the", "and", "for", "you", "all", "team", "staff", "office", "health", "center"]);

  function nameTokens(names) {
    const out = new Set();
    (names || []).forEach((n) => {
      if (!n) return;
      // display names and the local part of email addresses
      String(n)
        .replace(/@.*$/, "")
        .split(/[\s,._\-;:'"()<>]+/)
        .forEach((w) => {
          if (w.length >= 2 && !/^\d+$/.test(w) && !STOP_NAME_WORDS.has(w.toLowerCase())) out.add(w);
        });
    });
    return [...out].sort((a, b) => b.length - a.length);
  }

  function redact(text, opts) {
    opts = opts || {};
    let t = text;

    // Emails, links
    t = t.replace(/[\w.+\-]+@[\w\-]+\.[\w.\-]+/g, "[email]");
    t = t.replace(/\bhttps?:\/\/\S+|\bwww\.\S+/gi, "[link]");

    // SSN-style, phone numbers
    t = t.replace(/\b\d{3}-\d{2}-\d{4}\b/g, "[id]");
    t = t.replace(/(\+?1[\s.\-]?)?\(?\b\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}\b/g, "[phone]");

    // Dates: 01/02/1980, 1-2-80, 1980-01-02, Jan 2 1980, January 2, 1980
    t = t.replace(/\b\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}\b/g, "[date]");
    t = t.replace(/\b\d{4}-\d{2}-\d{2}\b/g, "[date]");
    t = t.replace(
      /\b(Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|June?|July?|Aug(ust)?|Sep(t(ember)?)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)\.?\s+\d{1,2}(st|nd|rd|th)?,?\s+\d{4}\b/gi,
      "[date]"
    );
    t = t.replace(/\b(DOB|D\.O\.B\.?|date of birth)\s*[:\-]?\s*\S+/gi, "DOB [date]");

    // Street addresses: "672 Parkside Ave", "12 E 4th Street Apt 3"
    t = t.replace(
      /\b\d{1,6}\s+(?:[A-Z0-9][\w.]*\s+){1,4}(St|Street|Ave|Avenue|Rd|Road|Blvd|Boulevard|Pl|Place|Ln|Lane|Dr|Drive|Ct|Court|Pkwy|Parkway|Way|Terrace|Ter)\b\.?(\s*(Apt|Unit|#|Fl|Floor)\.?\s*\w+)?/gi,
      "[address]"
    );

    // Any remaining 5+ digit runs (MRN, Medicaid IDs, account numbers…) and alphanumeric IDs like AB12345C
    t = t.replace(/\b[A-Z]{1,3}\d{5,}[A-Z]?\b/g, "[id]");
    t = t.replace(/\b\d[\d\s\-]{4,}\d\b/g, "[number]");

    // Titles: Mr. Cohen, Dr Jane Smith
    t = t.replace(/\b(Mr|Mrs|Ms|Miss|Dr|Rabbi|Rev)\.?\s+[A-Z][a-zA-Z'\-]+(\s+[A-Z][a-zA-Z'\-]+)?/g, "$1 [name]");

    // Greeting names: "Hi Sarah,", "Good morning Moshe and Leah"
    t = t.replace(
      /^(\s*(?:hi|hello|hey|dear|good (?:morning|afternoon|evening)|thanks|thank you)),?\s+([A-Z][a-zA-Z'\-]+(?:\s+(?:and|&)\s+[A-Z][a-zA-Z'\-]+|\s+[A-Z][a-zA-Z'\-]+)*)/gim,
      "$1 [name]"
    );

    // "patient/client/member X Y", "the Waldman case/family"
    t = t.replace(
      /\b([Pp]atient|[Cc]lient|[Mm]ember|[Rr]esident|[Pp]articipant|[Cc]ustomer|[Aa]pplicant|[Ee]nrollee|[Cc]aller|[Mm]other|[Ff]ather|[Ss]on|[Dd]aughter)\s+(?!ID\b)[A-Z][a-zA-Z'\-]+(\s+[A-Z][a-zA-Z'\-]+)?/g,
      "$1 [name]"
    );
    t = t.replace(/\b[A-Z][a-zA-Z'\-]+(?:'s)?\s+(case|family|household|file|chart|account|application|referral)\b/g, (m, w) => {
      const first = m.split(/\s+/)[0].replace(/'s$/, "");
      return COMMON_CAPS.has(first.toLowerCase()) ? m : "[name] " + w;
    });

    // Known names: recipients, sender, extra words from settings
    const tokens = nameTokens([...(opts.names || []), ...(opts.extraWords || [])]);
    tokens.forEach((w) => {
      const esc = w.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      t = t.replace(new RegExp("\\b" + esc + "('s)?\\b", "gi"), "[name]");
    });

    // Collapse repeats
    t = t.replace(/(\[name\]\s*){2,}/g, "[name] ");
    return t;
  }

  const COMMON_CAPS = new Set([
    "the", "this", "that", "my", "your", "our", "their", "his", "her", "new", "open", "closed", "same", "next",
    "last", "a", "an", "each", "every", "one", "any", "no", "medicaid", "medicare", "snap", "unite", "zoho",
  ]);

  // ---------- 3. Rules-only subjects (no AI) ----------
  const STOPWORDS = new Set(
    ("a about above after again against all am an and any are as at be because been before being below between both but by " +
      "can could did do does doing down during each few for from further had has have having he her here hers herself him himself " +
      "his how i if in into is it its itself just let me more most my myself no nor not now of off on once only or other our ours " +
      "ourselves out over own same she should so some such than that the their theirs them themselves then there these they this " +
      "those through to too under until up very was we were what when where which while who whom why will with would you your " +
      "yours yourself yourselves hi hello hey dear thanks thank please regards best also get got know like need want see let's lets " +
      "im i'm i've ive i'll ill you're youre don't dont can't cant it's its we're were just send sent today tomorrow yesterday " +
      "email emails reply attached attachment name number phone date link id address make sure going go well still back one two")
      .split(/\s+/)
  );

  function titleCase(s) {
    return s.replace(/(^|[\s\-–])([a-z])/g, (m, a, b) => a + b.toUpperCase());
  }

  function keyPhrases(text, max) {
    const body = text
      .replace(/\[[a-z]+\]/g, " ")
      .replace(/^\s*(hi|hello|hey|dear|good (morning|afternoon|evening))\b.*$/gim, " ")
      .replace(/^\s*(thanks|thank you|best|regards|sincerely)[\s\S]{0,40}$/gim, " ");
    const orig = body.match(/[A-Za-z][A-Za-z'\-]{2,}/g) || [];
    const words = orig.map((w) => w.toLowerCase());
    const display = new Map();
    orig.forEach((w) => { const k = w.toLowerCase(); if (!display.has(k) || /[A-Z]{2,}/.test(w)) display.set(k, w); });
    const counts = new Map();
    const firstPos = new Map();
    words.forEach((w, i) => {
      if (STOPWORDS.has(w)) return;
      counts.set(w, (counts.get(w) || 0) + 1);
      if (!firstPos.has(w)) firstPos.set(w, i);
    });
    // bigrams of non-stopwords
    const bigrams = new Map();
    for (let i = 0; i < words.length - 1; i++) {
      const a = words[i], b = words[i + 1];
      if (STOPWORDS.has(a) || STOPWORDS.has(b)) continue;
      const k = a + " " + b;
      bigrams.set(k, (bigrams.get(k) || 0) + 1);
    }
    const scored = [...counts.entries()]
      .map(([w, c]) => [w, c * 2 + w.length / 6 - (firstPos.get(w) || 0) / 200])
      .sort((x, y) => y[1] - x[1]);
    const topBigram = [...bigrams.entries()].sort((x, y) => y[1] - x[1])[0];
    const out = [];
    if (topBigram && topBigram[1] >= 1) out.push(topBigram[0]);
    for (const [w] of scored) {
      if (out.length >= max) break;
      if (!out.some((p) => p.split(" ").includes(w))) out.push(w);
    }
    return out.slice(0, max).map((p) => p.split(" ").map((w) => display.get(w) || w).join(" "));
  }

  function rulesSubjects(text) {
    const lower = text.toLowerCase();
    const topic = titleCase(keyPhrases(text, 2).join(" – ") || "Quick Note");
    const main = titleCase(keyPhrases(text, 1)[0] || "Quick Note");
    const has = (re) => re.test(lower);
    const out = [];

    if (has(/\b(urgent|asap|immediately|right away|time.sensitive)\b/)) out.push("Urgent: " + topic);
    if (has(/\b(follow(ing)?[\s\-]?up|checking in|any update|status)\b/)) out.push("Follow-Up: " + main);
    if (has(/\b(meet(ing)?|call|schedule|available|availability|zoom|teams)\b/)) out.push("Scheduling: " + main);
    if (has(/\b(attached|attaching|enclosed|see the file|find the)\b/)) out.push(main + " – Attached");
    if (has(/\b(can you|could you|would you|please|request|need you to)\b/)) out.push("Request: " + main);
    if (has(/\?/)) out.push("Question About " + main);
    if (has(/\b(reminder|don't forget|remember to)\b/)) out.push("Reminder: " + main);
    if (has(/\b(update|fyi|heads up|letting you know)\b/)) out.push("Update: " + main);
    if (has(/\b(thank you|thanks so much|appreciate)\b/)) out.push("Thank You – " + main);

    out.push(topic);
    out.push(main + " – Next Steps");

    return [...new Set(out.map((s) => s.replace(/\s+/g, " ").trim()))].slice(0, 3);
  }

  // ---------- 4. AI prompt + response parsing ----------
  function buildPrompt(text, isReply) {
    return (
      "Suggest 3 different subject lines for the email below.\n" +
      "Rules:\n" +
      "- Each subject is 3 to 8 words, clear and specific, professional tone.\n" +
      "- Never include personal names, phone numbers, dates of birth or ID numbers.\n" +
      "- Placeholders like [name], [phone], [date] mark removed private details; do not repeat them.\n" +
      (isReply ? "- This is a reply; do not start with 'Re:'.\n" : "") +
      "- Reply with ONLY a JSON array of 3 strings, nothing else.\n\n" +
      "Email:\n\"\"\"\n" +
      text.slice(0, 6000) +
      "\n\"\"\""
    );
  }

  function parseSubjects(raw) {
    if (!raw) return [];
    const m = raw.match(/\[[\s\S]*\]/);
    if (m) {
      try {
        const arr = JSON.parse(m[0]);
        if (Array.isArray(arr)) return arr.map((s) => String(s).trim()).filter(Boolean).slice(0, 3);
      } catch (e) { /* fall through */ }
    }
    return raw
      .split("\n")
      .map((l) => l.replace(/^\s*(\d+[.)]|[-*•])\s*/, "").replace(/^["']|["']$/g, "").trim())
      .filter((l) => l && l.length < 120)
      .slice(0, 3);
  }

  const api = { cleanBody, redact, rulesSubjects, keyPhrases, buildPrompt, parseSubjects, nameTokens };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.SubjectCore = api;
})(typeof window !== "undefined" ? window : globalThis);
