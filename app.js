/*
 * app.js — Giao diện add-in "Chính Tả Việt" cho Microsoft Word
 *   - Kiểm tra chính tả nhanh (luật chính tả, offline)
 *   - Kiểm tra chính tả bằng AI (Claude API)
 *   - Rà soát căn cứ pháp lý (kiểm tra hình thức offline + AI tra cứu web)
 * Viết bằng JavaScript ES5 để chạy được trên cả Word 2016/2019 và Microsoft 365.
 */
(function () {
  "use strict";

  var P = window.Promise || (window.OfficeExtension && window.OfficeExtension.Promise);
  var KEYS = { apiKey: "vnspell.apiKey", model: "vnspell.model", legalModel: "vnspell.legalModel",
    dict: "vnspell.userDict", skip: "vnspell.skipForeign", format: "vnspell.format" };
  var COLORS = { word: "Yellow", phrase: "Yellow", ai: "Turquoise", legal: "Pink", legalai: "Pink", format: "Yellow" };
  var TAGS = { word: "Chính tả", phrase: "Từ hay sai", ai: "AI chính tả", legal: "Kiểm tra hình thức", legalai: "AI tra cứu", format: "Thể thức" };
  var GROUPS = { word: 0, phrase: 0, ai: 0, format: 1, legal: 2, legalai: 2 };
  var GROUP_TITLES = { 1: "Thể thức văn bản", 2: "Căn cứ pháp lý" };
  var MM = 2.8346, CM = 28.3465;
  var STATUS = {
    sai: { label: "✗ Sai", cls: "st-bad", order: 0 },
    het_hieu_luc: { label: "⌛ Hết hiệu lực", cls: "st-old", order: 1 },
    canh_bao: { label: "⚠ Cần xem lại", cls: "st-warn", order: 2 },
    can_kiem_tra: { label: "? Chưa xác minh được", cls: "st-unk", order: 3 },
    luu_y: { label: "ℹ Lưu ý", cls: "st-info", order: 4 },
    dung: { label: "✓ Đúng", cls: "st-ok", order: 5 }
  };
  var SPELL_CHUNK = 3500;
  var LEGAL_MAX_SENTENCES = 40;

  var SPELL_PROMPT =
    "Bạn là chuyên gia soát lỗi chính tả tiếng Việt. Hãy tìm trong văn bản người dùng gửi: lỗi chính tả, " +
    "lỗi gõ phím, sai dấu thanh, dùng nhầm từ (ví dụ: dành/giành, sát nhập/sáp nhập, chuẩn đoán/chẩn đoán), " +
    "nhầm phụ âm (s/x, ch/tr, d/gi/r, l/n), nhầm vần (n/ng, c/t).\n" +
    "KHÔNG sửa văn phong, KHÔNG viết lại câu, KHÔNG sửa tên riêng, từ nước ngoài, thuật ngữ chuyên ngành, " +
    "cách viết i/y (lý/lí) hay vị trí dấu (hòa/hoà) vì đều được chấp nhận. Chỉ báo lỗi bạn chắc chắn.\n" +
    "Trả về DUY NHẤT một đối tượng JSON, không viết gì thêm, theo mẫu:\n" +
    "{\"loi\":[{\"sai\":\"...\",\"sua\":\"...\",\"giai_thich\":\"...\"}]}\n" +
    "- \"sai\": chép NGUYÊN VĂN đoạn chứa lỗi từ văn bản (giữ đúng chữ hoa/thường và dấu), dài 2 đến 5 từ, " +
    "gồm cả từ đứng cạnh để không nhầm với chỗ viết đúng.\n" +
    "- \"sua\": đoạn đó sau khi sửa (chỉ thay phần sai, giữ nguyên phần còn lại).\n" +
    "- \"giai_thich\": giải thích ngắn bằng tiếng Việt, dưới 20 từ.\n" +
    "Nếu không có lỗi, trả về {\"loi\":[]}";

  function legalPrompt() {
    var d = new Date();
    var today = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
    return "Bạn là chuyên viên pháp chế Việt Nam, chuyên rà soát việc viện dẫn văn bản quy phạm pháp luật. Hôm nay là ngày " + today + ".\n" +
      "Người dùng gửi các câu được đánh số [1], [2]... trích từ một văn bản. Với MỖI trích dẫn văn bản pháp luật trong các câu đó " +
      "(tên văn bản, số hiệu, ngày ban hành, điều, khoản, điểm), hãy DÙNG CÔNG CỤ TÌM KIẾM WEB để xác minh, ưu tiên nguồn chính thống: " +
      "vbpl.vn, congbao.chinhphu.vn, chinhphu.vn, quochoi.vn, moj.gov.vn, sau đó mới tới thuvienphapluat.vn, luatvietnam.vn.\n" +
      "Cần kiểm tra:\n" +
      "1. Văn bản có tồn tại không; tên, loại văn bản, số hiệu, ngày ban hành, cơ quan ban hành có đúng không.\n" +
      "2. Văn bản còn hiệu lực tại ngày hôm nay không; nếu đã bị thay thế hoặc sửa đổi thì bởi văn bản nào, từ ngày nào.\n" +
      "3. Điều/khoản/điểm được dẫn có tồn tại không, và nội dung có đúng với ý mà câu văn đang viện dẫn không.\n" +
      "Quy tắc: không dựa vào trí nhớ khi có thể tra cứu; không bịa số hiệu hay nội dung. Không tìm được nguồn xác thực thì đánh giá " +
      "\"can_kiem_tra\". Nếu câu văn nói rõ sự việc xảy ra khi luật cũ còn hiệu lực thì dẫn luật cũ không phải là lỗi. " +
      "Trích dẫn tới chính văn bản đang soạn (\"Luật này\", \"Nghị định này\") thì bỏ qua.\n" +
      "Sau khi tra cứu xong, KẾT THÚC câu trả lời bằng đúng một khối sau (JSON hợp lệ):\n" +
      "<ketqua>{\"ket_qua\":[{\"cau\":1,\"trich_dan\":\"...\",\"danh_gia\":\"dung\",\"nhan_xet\":\"...\",\"goi_y\":\"\",\"nguon\":[\"https://...\"]}]}</ketqua>\n" +
      "- cau: số thứ tự câu.\n" +
      "- trich_dan: chép NGUYÊN VĂN phần trích dẫn trong câu, đúng từng ký tự (ví dụ \"khoản 2 Điều 468 Bộ luật Dân sự 2015\").\n" +
      "- danh_gia: một trong \"dung\", \"sai\", \"het_hieu_luc\", \"can_kiem_tra\".\n" +
      "- nhan_xet: tiếng Việt, dưới 60 từ, nói rõ đúng/sai ở điểm nào, văn bản nào thay thế (nếu có).\n" +
      "- goi_y: đoạn viết đúng để thay cho trich_dan; để \"\" nếu trích dẫn đúng hoặc chưa chắc.\n" +
      "- nguon: 1–3 đường link đã dùng để xác minh.\n" +
      "Mọi trích dẫn đều phải có mặt trong ket_qua, kể cả trích dẫn đúng.";
  }

  var state = { issues: [], busy: false, adapter: null, nextId: 1, canComment: false, inWord: false };

  // ------------------------------------------------------------------
  // Tiện ích
  // ------------------------------------------------------------------
  function $(id) { return document.getElementById(id); }
  function store(key, value) {
    try {
      if (value === undefined) return window.localStorage.getItem(key);
      if (value === null) window.localStorage.removeItem(key);
      else window.localStorage.setItem(key, value);
    } catch (e) { /* bộ nhớ bị chặn */ }
    return null;
  }
  function loadDict() {
    try { return JSON.parse(store(KEYS.dict) || "{}") || {}; } catch (e) { return {}; }
  }
  function saveDict(d) { store(KEYS.dict, JSON.stringify(d)); }
  function skipForeign() { return store(KEYS.skip) !== "0"; }
  function trim(s) { return String(s == null ? "" : s).replace(/^\s+|\s+$/g, ""); }

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }
  function button(label, cls, onClick) {
    var b = el("button", cls, label);
    b.onclick = onClick;
    return b;
  }
  function setStatus(msg, kind) {
    var s = $("status");
    s.textContent = msg || "";
    s.className = kind || "";
    s.style.display = msg ? "block" : "none";
  }
  function setBusy(b) {
    state.busy = b;
    var ids = ["btnQuick", "btnAI", "btnLegal", "btnFmtCheck", "btnFmtApply", "btnFixAll", "btnClear"];
    for (var i = 0; i < ids.length; i++) $(ids[i]).disabled = b;
  }
  function errText(e) { return (e && (e.message || e.toString())) || "Lỗi không rõ"; }
  function isLetter(ch) { return new RegExp("[" + VnChecker.LETTER_CLASS + "]").test(ch || ""); }
  function openIssues() { return state.issues.filter(function (i) { return !i.done; }); }
  function isLegal(is) { return is.kind === "legal" || is.kind === "legalai" || is.kind === "format"; }
  function fmtSettings() {
    var d = FormatChecker.DEFAULTS, saved = {};
    try { saved = JSON.parse(store(KEYS.format) || "{}") || {}; } catch (e) { saved = {}; }
    var st = JSON.parse(JSON.stringify(d));
    Object.keys(saved).forEach(function (k) { st[k] = saved[k]; });
    return st;
  }
  function supports(set, ver) {
    try { return !!(window.Office && Office.context && Office.context.requirements.isSetSupported(set, ver)); } catch (e) { return false; }
  }
  function needsMark(is) { return !!is.search && !(is.status === "dung"); }

  function sequence(items, fn) {
    var p = P.resolve();
    items.forEach(function (item, idx) { p = p.then(function () { return fn(item, idx); }); });
    return p;
  }

  function openLink(url) {
    try {
      if (window.Office && Office.context && Office.context.ui && Office.context.ui.openBrowserWindow) {
        Office.context.ui.openBrowserWindow(url);
        return;
      }
    } catch (e) { /* bỏ qua */ }
    window.open(url, "_blank");
  }

  // ------------------------------------------------------------------
  // Làm việc với văn bản Word
  // ------------------------------------------------------------------
  function escapeSearch(s) { return s.replace(/\^/g, "^^"); }
  function searchOptions(issue) { return { matchCase: !!issue.matchCase, matchWholeWord: issue.wholeWord !== false }; }
  function sameText(a, b, matchCase) {
    a = VnChecker.nfc(a); b = VnChecker.nfc(b);
    return matchCase ? a === b : a.toLowerCase() === b.toLowerCase();
  }
  function noSearch() { return P.reject(new Error("Không xác định được vị trí trích dẫn trong văn bản.")); }

  var WordDoc = {
    getText: function () {
      return Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var body = ctx.document.body;
        sel.load("text");
        body.load("text");
        return ctx.sync().then(function () {
          var s = sel.text || "";
          if (s.replace(/\s/g, "").length > 1) return { text: s, scope: "đoạn đang chọn" };
          return { text: body.text || "", scope: "toàn bộ văn bản" };
        });
      });
    },
    getParagraphs: function () {
      var nested = supports("WordApi", "1.3");
      return Word.run(function (ctx) {
        var ps = ctx.document.body.paragraphs;
        ps.load(nested ? "text,tableNestingLevel,font/name,font/size" : "text,font/name,font/size");
        return ctx.sync().then(function () {
          return ps.items.map(function (p) {
            return { text: p.text || "", inTable: nested ? p.tableNestingLevel > 0 : false, fontName: p.font.name, fontSize: p.font.size };
          });
        });
      });
    },
    _each: function (issues, apply) {
      issues = issues.filter(function (is) { return is.search && is.search.length <= 250; });
      if (!issues.length) return P.resolve();
      return Word.run(function (ctx) {
        var body = ctx.document.body, found = [];
        issues.forEach(function (is) {
          var r = body.search(escapeSearch(is.search), searchOptions(is));
          r.load("text");
          found.push({ is: is, r: r });
        });
        return ctx.sync().then(function () {
          found.forEach(function (f) { apply(f.is, f.r.items); });
          return ctx.sync();
        });
      });
    },
    highlight: function (issues) {
      return this._each(issues.filter(needsMark), function (is, items) {
        if (!isLegal(is)) is.count = items.length || is.count;
        items.forEach(function (it) { it.font.highlightColor = COLORS[is.kind]; });
      });
    },
    clearHighlight: function (issues) {
      return this._each(issues, function (is, items) {
        items.forEach(function (it) { it.font.highlightColor = null; });
      });
    },
    replace: function (issue, replacement) {
      if (!issue.search) return noSearch();
      return Word.run(function (ctx) {
        var r = ctx.document.body.search(escapeSearch(issue.search), searchOptions(issue));
        r.load("text");
        return ctx.sync().then(function () {
          var n = 0;
          r.items.forEach(function (it) {
            if (!sameText(it.text, issue.search, issue.matchCase)) return;
            var text = issue.matchCase ? replacement : VnChecker.applyCase(VnChecker.nfc(it.text), replacement.toLowerCase());
            var nr = it.insertText(text, "Replace");
            nr.font.highlightColor = null;
            n++;
          });
          return ctx.sync().then(function () { return n; });
        });
      });
    },
    select: function (issue) {
      if (!issue.search) return noSearch();
      return Word.run(function (ctx) {
        var r = ctx.document.body.search(escapeSearch(issue.search), searchOptions(issue));
        r.load("text");
        return ctx.sync().then(function () {
          if (!r.items.length) return 0;
          issue.cursor = (issue.cursor || 0) % r.items.length;
          r.items[issue.cursor].select();
          issue.cursor++;
          return ctx.sync().then(function () { return r.items.length; });
        });
      });
    },
    comment: function (issue, text) {
      if (!issue.search) return noSearch();
      return Word.run(function (ctx) {
        var r = ctx.document.body.search(escapeSearch(issue.search), searchOptions(issue));
        r.load("text");
        return ctx.sync().then(function () {
          if (!r.items.length) return 0;
          r.items[0].insertComment(text);
          return ctx.sync().then(function () { return 1; });
        });
      });
    }
  };

  // Chế độ thử: dùng ô văn bản trên trang thay cho Word
  var DemoDoc = {
    getParagraphs: function () {
      return P.resolve($("demoText").value.split(/\n/).map(function (t) { return { text: t, inTable: false }; }));
    },
    _re: function (issue) {
      var esc = issue.search.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      var L = VnChecker.LETTER_CLASS;
      var core = issue.wholeWord === false ? "(" + esc + ")" : "(" + esc + ")(?=$|[^" + L + "])";
      var pre = issue.wholeWord === false ? "()" : "(^|[^" + L + "])";
      return new RegExp(pre + core, issue.matchCase ? "g" : "gi");
    },
    getText: function () {
      var t = $("demoText"), s = t.value.substring(t.selectionStart, t.selectionEnd);
      if (s.replace(/\s/g, "").length > 1) return P.resolve({ text: s, scope: "đoạn đang chọn" });
      return P.resolve({ text: t.value, scope: "toàn bộ ô văn bản" });
    },
    highlight: function () { return P.resolve(); },
    clearHighlight: function () { return P.resolve(); },
    replace: function (issue, replacement) {
      if (!issue.search) return noSearch();
      var t = $("demoText"), n = 0;
      t.value = VnChecker.nfc(t.value).replace(this._re(issue), function (m, pre, w) {
        n++;
        return pre + (issue.matchCase ? replacement : VnChecker.applyCase(w, replacement.toLowerCase()));
      });
      return P.resolve(n);
    },
    select: function (issue) {
      if (!issue.search) return noSearch();
      var t = $("demoText"), re = this._re(issue), m, hits = [];
      while ((m = re.exec(t.value)) !== null) { hits.push(m.index + m[1].length); if (!m[0].length) re.lastIndex++; }
      if (!hits.length) return P.resolve(0);
      issue.cursor = (issue.cursor || 0) % hits.length;
      var start = hits[issue.cursor++];
      t.focus();
      t.setSelectionRange(start, start + issue.search.length);
      return P.resolve(hits.length);
    },
    comment: function () { return P.reject(new Error("Ghi chú chỉ dùng được trong Word.")); }
  };

  // Thay nhóm lỗi cũ (cùng loại) bằng nhóm lỗi mới
  function replaceKinds(kinds, fresh) {
    var old = state.issues.filter(function (i) { return kinds.indexOf(i.kind) >= 0 && !i.done; });
    return state.adapter.clearHighlight(old).then(function () {
      state.issues = state.issues.filter(function (i) { return kinds.indexOf(i.kind) < 0 && !i.done; }).concat(fresh);
      return state.adapter.highlight(fresh);
    });
  }

  // ------------------------------------------------------------------
  // Danh sách lỗi
  // ------------------------------------------------------------------
  function sortIssues(list) {
    return list.slice().sort(function (a, b) {
      var la = GROUPS[a.kind] || 0, lb = GROUPS[b.kind] || 0;
      if (la !== lb) return la - lb;
      if (la) {
        var oa = (STATUS[a.status] || STATUS.can_kiem_tra).order, ob = (STATUS[b.status] || STATUS.can_kiem_tra).order;
        if (oa !== ob) return oa - ob;
      }
      return (a.pos || 0) - (b.pos || 0);
    });
  }

  function renderSources(card, sources) {
    if (!sources || !sources.length) return;
    var box = el("div", "sources", "Nguồn: ");
    sources.forEach(function (url, i) {
      var host = url.replace(/^https?:\/\//, "").split("/")[0];
      var a = el("a", "", host);
      a.href = url;
      a.title = url;
      a.onclick = function (e) { e.preventDefault(); openLink(url); };
      if (i) box.appendChild(document.createTextNode(" · "));
      box.appendChild(a);
    });
    card.appendChild(box);
  }

  function render() {
    var list = $("issues");
    list.innerHTML = "";
    var open = sortIssues(openIssues());
    var fixable = open.filter(function (i) { return !isLegal(i); });
    $("toolbar").style.display = open.length ? "flex" : "none";
    $("btnFixAll").style.display = fixable.length ? "" : "none";

    var lastGroup = 0;
    open.forEach(function (is) {
      var g = GROUPS[is.kind] || 0;
      if (g !== lastGroup && GROUP_TITLES[g]) list.appendChild(el("div", "section", GROUP_TITLES[g]));
      lastGroup = g;
      var st = isLegal(is) ? (STATUS[is.status] || STATUS.can_kiem_tra) : null;
      var card = el("div", "card " + is.kind + (st ? " " + st.cls : ""));
      var head = el("div", "card-head");
      head.appendChild(el("span", "tag", TAGS[is.kind]));
      if (st) head.appendChild(el("span", "badge " + st.cls, st.label));
      card.appendChild(head);
      card.appendChild(el("div", st ? "cite" : "wrong", is.display));
      if (!st && is.count > 1) card.appendChild(el("span", "count", "× " + is.count + " chỗ"));
      if (is.reason) card.appendChild(el("div", "reason", is.reason));
      if (is.kind === "legalai" && !is.search && is.context) {
        card.appendChild(el("div", "muted", "Trong câu: " + is.context.slice(0, 160)));
      }

      var sugg = el("div", "sugg");
      if (is.suggestions.length) {
        if (st) sugg.appendChild(el("span", "muted", "Gợi ý: "));
        is.suggestions.forEach(function (s) {
          var b = button(s, "chip", function () { fixIssue(is, s); });
          b.title = "Bấm để sửa thành: " + s;
          b.disabled = !is.search;
          sugg.appendChild(b);
        });
        card.appendChild(sugg);
      } else if (!st) {
        sugg.appendChild(el("span", "muted", "Chưa có gợi ý — bấm \"Tìm\" rồi tự sửa trong văn bản."));
        card.appendChild(sugg);
      }
      renderSources(card, is.sources);

      var act = el("div", "actions");
      if (is.search) act.appendChild(button("Tìm", "", function () { selectIssue(is); }));
      if (st && is.kind !== "format" && state.canComment && is.search && is.status !== "dung") {
        act.appendChild(button("Ghi chú vào văn bản", "", function () { commentIssue(is); }));
      }
      act.appendChild(button(is.status === "dung" ? "Ẩn" : "Bỏ qua", "", function () { ignoreIssue(is); }));
      if (is.kind === "word") act.appendChild(button("Thêm vào từ điển", "", function () { addToDict(is); }));
      card.appendChild(act);
      list.appendChild(card);
    });
  }

  function makeIssue(kind, fields) {
    fields.id = state.nextId++;
    fields.kind = kind;
    fields.done = false;
    fields.suggestions = fields.suggestions || [];
    return fields;
  }

  function summary(scope) {
    var open = openIssues().filter(function (i) { return !isLegal(i); });
    if (!open.length) return setStatus("Không phát hiện lỗi chính tả trong " + scope + ". 🎉", "ok");
    var total = 0;
    open.forEach(function (i) { total += i.count || 1; });
    setStatus("Tìm thấy " + open.length + " lỗi (" + total + " chỗ) trong " + scope + ". Bấm vào ô màu xanh để sửa.", "");
  }

  // ------------------------------------------------------------------
  // Gọi Claude API
  // ------------------------------------------------------------------
  function apiError(status, data) {
    var msg = data && data.error && data.error.message ? data.error.message : "";
    if (status === 401) return "API key không đúng. Vào Cài đặt để kiểm tra lại.";
    if (/web.?search/i.test(msg)) return "Tài khoản Claude API chưa bật tra cứu web. Vào platform.claude.com → Settings → Privacy, bật Web search rồi thử lại.";
    if (status === 403) return "API key không có quyền dùng mô hình này. " + msg;
    if (status === 404) return "Không tìm thấy mô hình AI. Thử chọn mô hình khác trong Cài đặt. " + msg;
    if (status === 429) return "Gửi quá nhiều yêu cầu hoặc tài khoản hết hạn mức. Thử lại sau ít phút. " + msg;
    if (status === 400 && /credit|balance/i.test(msg)) return "Tài khoản Claude API hết tiền. Hãy nạp thêm ở mục Billing.";
    if (status >= 500) return "Máy chủ Claude đang bận (mã " + status + "). Thử lại sau.";
    return "Lỗi từ Claude API (mã " + status + "): " + msg;
  }

  function postMessages(payload, key, timeoutMs) {
    return new P(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open("POST", "https://api.anthropic.com/v1/messages", true);
      xhr.setRequestHeader("content-type", "application/json");
      xhr.setRequestHeader("x-api-key", key);
      xhr.setRequestHeader("anthropic-version", "2023-06-01");
      xhr.setRequestHeader("anthropic-dangerous-direct-browser-access", "true");
      xhr.timeout = timeoutMs || 120000;
      xhr.onload = function () {
        var data = null;
        try { data = JSON.parse(xhr.responseText); } catch (e) { /* không phải JSON */ }
        if (xhr.status >= 200 && xhr.status < 300 && data) resolve(data);
        else reject(new Error(apiError(xhr.status, data)));
      };
      xhr.onerror = function () { reject(new Error("Không kết nối được tới Claude. Kiểm tra mạng Internet.")); };
      xhr.ontimeout = function () { reject(new Error("Claude phản hồi quá lâu. Thử lại với đoạn văn ngắn hơn.")); };
      xhr.send(JSON.stringify(payload));
    });
  }

  // Gọi API; nếu Claude tạm dừng giữa chừng (pause_turn) thì gửi tiếp để nó làm xong. Trả về toàn bộ chữ.
  function callClaude(opts) {
    var messages = [{ role: "user", content: opts.user }];
    var text = "", rounds = 0;
    function step() {
      var payload = { model: opts.model, max_tokens: opts.maxTokens || 4096, system: opts.system, messages: messages };
      if (opts.tools) payload.tools = opts.tools;
      return postMessages(payload, opts.key, opts.timeout).then(function (data) {
        (data.content || []).forEach(function (b) { if (b.type === "text") text += b.text; });
        if (data.stop_reason === "pause_turn" && rounds++ < 4) {
          messages = messages.concat([{ role: "assistant", content: data.content }]);
          return step();
        }
        return text;
      });
    }
    return step();
  }

  function requireKey() {
    var key = store(KEYS.apiKey);
    if (!key) {
      toggleSettings(true);
      setStatus("Hãy dán API key của Claude vào phần Cài đặt bên dưới rồi bấm Lưu.", "err");
    }
    return key;
  }

  // ------------------------------------------------------------------
  // Kiểm tra nhanh (luật chính tả)
  // ------------------------------------------------------------------
  function runQuick() {
    if (state.busy) return;
    setBusy(true);
    setStatus("Đang kiểm tra…");
    var scope;
    state.adapter.getText()
      .then(function (res) {
        scope = res.scope;
        var found = VnChecker.checkText(res.text, { userDict: loadDict(), skipForeign: skipForeign() });
        return replaceKinds(["word", "phrase"], found.map(function (f) {
          return makeIssue(f.kind, {
            display: f.word, search: f.search, reason: f.reason, suggestions: f.suggestions,
            replacement: f.replacement, confident: f.confident, count: f.count, pos: f.pos,
            matchCase: false, wholeWord: true
          });
        }));
      })
      .then(function () { render(); summary(scope); }, function (e) { setStatus("Có lỗi: " + errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  // ------------------------------------------------------------------
  // Kiểm tra chính tả bằng AI
  // ------------------------------------------------------------------
  function splitChunks(text) {
    var paras = text.split(/(\r\n|\r|\n)/), chunks = [], cur = "";
    paras.forEach(function (p) {
      if ((cur + p).length > SPELL_CHUNK && cur.replace(/\s/g, "")) { chunks.push(cur); cur = ""; }
      while (p.length > SPELL_CHUNK) {
        var cut = p.lastIndexOf(". ", SPELL_CHUNK);
        if (cut < SPELL_CHUNK / 2) cut = SPELL_CHUNK;
        chunks.push(p.slice(0, cut + 1)); p = p.slice(cut + 1);
      }
      cur += p;
    });
    if (cur.replace(/\s/g, "")) chunks.push(cur);
    return chunks;
  }

  function extractJSON(out, tag) {
    var body = out;
    if (tag) {
      var a = out.lastIndexOf("<" + tag + ">"), b = out.lastIndexOf("</" + tag + ">");
      if (a >= 0 && b > a) body = out.slice(a + tag.length + 2, b);
    }
    var s = body.indexOf("{"), e = body.lastIndexOf("}");
    if (s < 0 || e <= s) return null;
    try { return JSON.parse(body.slice(s, e + 1)); } catch (err) { return null; }
  }

  function runAI() {
    if (state.busy) return;
    var key = requireKey();
    if (!key) return;
    var model = store(KEYS.model) || "claude-haiku-5-5";
    setBusy(true);
    setStatus("Đang lấy văn bản…");
    var scope, found = [], seen = {};

    state.adapter.getText()
      .then(function (res) {
        scope = res.scope;
        var chunks = splitChunks(VnChecker.nfc(res.text));
        if (!chunks.length) throw new Error("Văn bản đang trống.");
        return sequence(chunks, function (chunk, i) {
          setStatus("AI đang đọc phần " + (i + 1) + "/" + chunks.length + "…");
          return callClaude({ key: key, model: model, system: SPELL_PROMPT,
            user: "Văn bản cần soát lỗi:\n<<<\n" + chunk + "\n>>>" }).then(function (out) {
            var obj = extractJSON(out);
            var list = (obj && (obj.loi || obj.errors)) || [];
            list.forEach(function (x) {
              if (!x) return;
              var sai = trim(VnChecker.nfc(x.sai)), sua = trim(VnChecker.nfc(x.sua));
              if (!sai || !sua || sai === sua || sai.length > 250 || chunk.indexOf(sai) < 0 || seen[sai]) return;
              seen[sai] = true;
              found.push(makeIssue("ai", {
                display: sai, search: sai, reason: x.giai_thich || "AI đề xuất sửa",
                suggestions: [sua], confident: true, count: 1, pos: chunk.indexOf(sai),
                matchCase: true, wholeWord: isLetter(sai.charAt(0)) && isLetter(sai.charAt(sai.length - 1))
              }));
            });
          });
        });
      })
      .then(function () { return replaceKinds(["ai"], found); })
      .then(function () { render(); summary(scope); }, function (e) { setStatus(errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  // ------------------------------------------------------------------
  // Rà soát căn cứ pháp lý
  // ------------------------------------------------------------------
  function batchSentences(sentences) {
    var batches = [], cur = [], len = 0;
    sentences.forEach(function (s) {
      if (cur.length && (cur.length >= 4 || len + s.length > 1800)) { batches.push(cur); cur = []; len = 0; }
      cur.push(s); len += s.length;
    });
    if (cur.length) batches.push(cur);
    return batches;
  }

  function legalSummary(scope, aiDone, aiNote) {
    var list = openIssues().filter(function (i) { return i.kind === "legal" || i.kind === "legalai"; }), count = {};
    list.forEach(function (i) { count[i.status] = (count[i.status] || 0) + 1; });
    var parts = [];
    Object.keys(STATUS).forEach(function (k) { if (count[k]) parts.push(count[k] + " " + STATUS[k].label.replace(/^\S+\s/, "").toLowerCase()); });
    var msg = list.length ? "Kết quả rà soát " + scope + ": " + parts.join(", ") + "." : "Không phát hiện vấn đề về hình thức trích dẫn trong " + scope + ".";
    if (aiNote) msg += " " + aiNote;
    setStatus(msg, aiDone && !count.sai && !count.het_hieu_luc ? "ok" : "");
  }

  function runLegal() {
    if (state.busy) return;
    setBusy(true);
    setStatus("Đang kiểm tra hình thức trích dẫn…");
    var key = store(KEYS.apiKey);
    var model = store(KEYS.legalModel) || "claude-sonnet-5-5";
    var scope, sentences = [];

    state.adapter.getText()
      .then(function (res) {
        scope = res.scope;
        var local = LegalChecker.check(res.text).map(function (f) {
          return makeIssue("legal", { status: f.status, display: f.display, search: f.search, pos: f.pos,
            reason: f.reason, suggestions: f.suggestions, matchCase: true, wholeWord: false });
        });
        sentences = LegalChecker.citationSentences(res.text, LEGAL_MAX_SENTENCES);
        return replaceKinds(["legal", "legalai"], local);
      })
      .then(function () {
        render();
        if (!key) {
          legalSummary(scope, false, "Muốn AI tra cứu nội dung điều luật và hiệu lực văn bản, hãy nhập API key trong Cài đặt.");
          return;
        }
        if (!sentences.length) { legalSummary(scope, true, "Không thấy câu nào có trích dẫn văn bản để AI tra cứu."); return; }
        var batches = batchSentences(sentences);
        var failed = 0;
        return sequence(batches, function (batch, bi) {
          setStatus("AI đang tra cứu nhóm " + (bi + 1) + "/" + batches.length + " (" + sentences.length +
            " câu có trích dẫn). Mỗi nhóm có thể mất 1–2 phút…");
          var user = "Các câu cần rà soát:\n" + batch.map(function (s, i) { return "[" + (i + 1) + "] " + s; }).join("\n");
          return callClaude({ key: key, model: model, system: legalPrompt(), user: user, maxTokens: 8000, timeout: 300000,
            tools: [{ type: "web_search_20250305", name: "web_search", max_uses: Math.min(12, batch.length * 3) }] })
            .then(function (out) {
              var obj = extractJSON(out, "ketqua");
              if (!obj) { failed++; return; }
              var fresh = [];
              (obj.ket_qua || []).forEach(function (x) {
                if (!x) return;
                var sentence = batch[(parseInt(x.cau, 10) || 1) - 1] || "";
                var cite = trim(VnChecker.nfc(x.trich_dan));
                if (!cite) return;
                var status = STATUS[x.danh_gia] ? x.danh_gia : "can_kiem_tra";
                var goiY = trim(VnChecker.nfc(x.goi_y));
                var inText = sentence.indexOf(cite) >= 0 && cite.length <= 250;
                fresh.push(makeIssue("legalai", {
                  status: status, display: cite, search: inText ? cite : null, context: sentence,
                  reason: trim(x.nhan_xet), suggestions: goiY && goiY !== cite && status !== "dung" ? [goiY] : [],
                  sources: (x.nguon || []).filter(function (u) { return /^https?:\/\//.test(u); }).slice(0, 3),
                  matchCase: true, wholeWord: false, pos: 100000 + bi * 10 + (parseInt(x.cau, 10) || 0)
                }));
              });
              // Mục kiểm tra hình thức trùng với kết quả AI (cùng báo có vấn đề) thì ẩn bớt cho gọn
              fresh.forEach(function (f) {
                if (f.status === "dung" || f.status === "can_kiem_tra") return;
                state.issues.forEach(function (l) {
                  if (l.kind === "legal" && !l.done && (f.display.indexOf(l.display) >= 0 || l.display.indexOf(f.display) >= 0)) l.done = true;
                });
              });
              state.issues = state.issues.concat(fresh);
              return state.adapter.highlight(fresh).then(render);
            }, function (e) {
              if (/API key|hết tiền|chưa bật|kết nối/.test(errText(e))) throw e;
              failed++;
            });
        }).then(function () {
          var note = failed ? "Có " + failed + " nhóm AI không trả kết quả, hãy thử lại." : "";
          if (sentences.length >= LEGAL_MAX_SENTENCES) note += " Chỉ rà soát " + LEGAL_MAX_SENTENCES + " câu đầu; bôi đen từng phần để rà soát tiếp.";
          legalSummary(scope, true, note + " Kết quả AI chỉ để tham khảo — hãy mở nguồn để đối chiếu.");
        });
      })
      .then(null, function (e) { render(); setStatus(errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  // ------------------------------------------------------------------
  // Thể thức văn bản
  // ------------------------------------------------------------------
  var BODY_ROLES = { body: 1, cancu: 1, dieu: 1 };

  function fontIssues(paras, roles, st) {
    var out = [], wrongFont = 0, wrongSize = 0, i;
    for (i = 0; i < paras.length; i++) {
      var p = paras[i];
      if (!trim(p.text) || p.fontName === undefined) continue;
      if (p.fontName !== st.font) wrongFont++;
      if (BODY_ROLES[roles[i]] && !p.inTable && (p.fontSize === null || p.fontSize < 13 || p.fontSize > 14)) wrongSize++;
    }
    if (wrongFont) out.push({ kind: "format", status: "sai", display: wrongFont + " đoạn không dùng phông " + st.font, search: null, pos: -2,
      reason: "Văn bản hành chính dùng phông " + st.font + " (bộ mã Unicode), màu đen. Bấm \"Tự chỉnh thể thức\" để sửa.", suggestions: [] });
    if (wrongSize) out.push({ kind: "format", status: "sai", display: wrongSize + " đoạn nội dung có cỡ chữ ngoài 13–14", search: null, pos: -1,
      reason: "Nội dung văn bản dùng cỡ chữ 13 hoặc 14. Bấm \"Tự chỉnh thể thức\" để sửa.", suggestions: [] });
    return out;
  }

  function marginIssues() {
    if (!state.inWord || !supports("WordApiDesktop", "1.3")) return P.resolve([]);
    return Word.run(function (ctx) {
      var secs = ctx.document.sections;
      secs.load("items");
      return ctx.sync().then(function () {
        if (!secs.items.length) return [];
        var ps = secs.items[0].pageSetup;
        ps.load("topMargin,bottomMargin,leftMargin,rightMargin");
        return ctx.sync().then(function () {
          var bad = [];
          function mm(v) { return Math.round(v / MM); }
          if (mm(ps.topMargin) < 20 || mm(ps.topMargin) > 25) bad.push("trên " + mm(ps.topMargin) + " mm (chuẩn 20–25)");
          if (mm(ps.bottomMargin) < 20 || mm(ps.bottomMargin) > 25) bad.push("dưới " + mm(ps.bottomMargin) + " mm (chuẩn 20–25)");
          if (mm(ps.leftMargin) < 30 || mm(ps.leftMargin) > 35) bad.push("trái " + mm(ps.leftMargin) + " mm (chuẩn 30–35)");
          if (mm(ps.rightMargin) < 15 || mm(ps.rightMargin) > 20) bad.push("phải " + mm(ps.rightMargin) + " mm (chuẩn 15–20)");
          return bad.length ? [{ kind: "format", status: "sai", display: "Lề trang chưa đúng", search: null, pos: -3,
            reason: "Lề " + bad.join("; ") + ". Bấm \"Tự chỉnh thể thức\" để sửa.", suggestions: [] }] : [];
        });
      });
    }).then(null, function () { return []; });
  }

  function runFormatCheck(silentScope) {
    if (state.busy && !silentScope) return P.resolve();
    setBusy(true);
    setStatus("Đang kiểm tra thể thức…");
    var st = fmtSettings(), found = [];
    return state.adapter.getParagraphs()
      .then(function (paras) {
        var roles = FormatChecker.classify(paras);
        found = FormatChecker.check(paras, roles).concat(fontIssues(paras, roles, st));
        return marginIssues();
      })
      .then(function (m) {
        found = found.concat(m).map(function (f) {
          return makeIssue("format", { status: f.status, display: f.display, search: f.search, pos: f.pos, reason: f.reason,
            suggestions: f.suggestions, matchCase: true, wholeWord: false });
        });
        return replaceKinds(["format"], found);
      })
      .then(function () {
        render();
        var n = found.length;
        if (silentScope) setStatus(silentScope + (n ? " Còn " + n + " điểm cần xem lại bên dưới (chủ yếu về chữ, cần bạn xác nhận)." : " Không còn lỗi thể thức nào phát hiện được."), n ? "" : "ok");
        else setStatus(n ? "Phát hiện " + n + " điểm về thể thức. Lỗi về chữ: bấm ô xanh để sửa; lỗi về định dạng: bấm \"Tự chỉnh thể thức\"." :
          "Không phát hiện lỗi thể thức. 🎉", n ? "" : "ok");
      }, function (e) { setStatus("Có lỗi: " + errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  function pageNumberOoxml(size) {
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" ' +
      'pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
      '<w:sz w:val="' + (size * 2) + '"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
  }

  function runFormatApply() {
    if (state.busy) return;
    if (!state.inWord) { setStatus("Tự chỉnh thể thức chỉ chạy được trong Word. Ở chế độ thử, hãy dùng \"Kiểm tra thể thức\".", "err"); return; }
    setBusy(true);
    setStatus("Đang chỉnh thể thức…");
    var st = fmtSettings(), notes = [], counts = 0;
    var nested = supports("WordApi", "1.3");

    Word.run(function (ctx) {
      var ps = ctx.document.body.paragraphs;
      ps.load(nested ? "text,tableNestingLevel" : "text");
      return ctx.sync().then(function () {
        var paras = ps.items.map(function (p) { return { text: p.text || "", inTable: nested ? p.tableNestingLevel > 0 : false }; });
        var roles = FormatChecker.classify(paras);
        ps.items.forEach(function (p, i) {
          var role = roles[i], inTable = paras[i].inTable;
          p.font.name = st.font;
          p.font.color = "#000000";
          if (inTable && (role === "body" || role === "empty")) return;   // bảng số liệu: chỉ đổi phông
          var f = FormatChecker.formatFor(role, inTable, st);
          if (f.size) p.font.size = f.size;
          if (f.bold !== undefined) p.font.bold = f.bold;
          if (f.italic !== undefined) p.font.italic = f.italic;
          if (role === "empty") return;
          if (f.align) p.alignment = f.align;
          if (f.indent !== undefined) p.firstLineIndent = f.indent * CM;
          if (f.left !== undefined) p.leftIndent = f.left;
          if (f.after !== undefined) p.spaceAfter = f.after;
          p.spaceBefore = f.before || 0;
          if (f.line) p.lineSpacing = f.line * 12;
          counts++;
        });
        return ctx.sync();
      }).then(function () {
        // Lề trang (Word Microsoft 365 trên máy tính)
        if (!supports("WordApiDesktop", "1.3")) {
          notes.push("Word này chưa cho add-in chỉnh lề: vào Layout → Margins → Custom Margins, đặt Trên " + st.margins.top + ", Dưới " + st.margins.bottom +
            ", Trái " + st.margins.left + ", Phải " + st.margins.right + " mm.");
          return;
        }
        var secs = ctx.document.sections;
        secs.load("items");
        return ctx.sync().then(function () {
          secs.items.forEach(function (sec) {
            sec.pageSetup.topMargin = st.margins.top * MM;
            sec.pageSetup.bottomMargin = st.margins.bottom * MM;
            sec.pageSetup.leftMargin = st.margins.left * MM;
            sec.pageSetup.rightMargin = st.margins.right * MM;
          });
          return ctx.sync();
        }).then(null, function () { notes.push("Không chỉnh được lề trang tự động, hãy chỉnh tay trong Layout → Margins."); });
      }).then(function () {
        if (!st.pageNumbers) return;
        var secs = ctx.document.sections;
        secs.load("items");
        return ctx.sync().then(function () {
          var header = secs.items[0].getHeader("Primary");
          header.load("text");
          return ctx.sync().then(function () {
            if (trim(header.text)) { notes.push("Đầu trang đã có nội dung nên không chèn số trang."); return; }
            header.insertOoxml(pageNumberOoxml(st.bodySize), "Start");
            notes.push("Đã chèn số trang ở giữa lề trên. Để ẩn số ở trang đầu: bấm đúp vào đầu trang → tick \"Different First Page\".");
            return ctx.sync();
          });
        }).then(null, function () { notes.push("Không chèn được số trang tự động."); });
      });
    })
      .then(function () {
        setBusy(false);
        return runFormatCheck("Đã chỉnh định dạng " + counts + " đoạn (Ctrl + Z để hoàn tác). " + notes.join(" "));
      }, function (e) { setBusy(false); setStatus("Không chỉnh được: " + errText(e), "err"); });
  }

  function saveFormatSettings() {
    var st = fmtSettings();
    st.bodySize = parseInt($("fBody").value, 10) || 14;
    st.lineSpacing = parseFloat($("fLine").value) || 1.15;
    st.firstIndent = parseFloat($("fIndent").value) || 1.27;
    st.margins = {
      top: parseFloat($("fTop").value) || 20, bottom: parseFloat($("fBottom").value) || 20,
      left: parseFloat($("fLeft").value) || 30, right: parseFloat($("fRight").value) || 15
    };
    st.pageNumbers = $("fPage").checked;
    store(KEYS.format, JSON.stringify(st));
  }
  function loadFormatSettings() {
    var st = fmtSettings();
    $("fBody").value = String(st.bodySize);
    $("fLine").value = String(st.lineSpacing);
    $("fIndent").value = String(st.firstIndent);
    $("fTop").value = st.margins.top; $("fBottom").value = st.margins.bottom;
    $("fLeft").value = st.margins.left; $("fRight").value = st.margins.right;
    $("fPage").checked = !!st.pageNumbers;
  }

  // ------------------------------------------------------------------
  // Các thao tác trên từng mục
  // ------------------------------------------------------------------
  function fixIssue(is, replacement) {
    if (state.busy) return;
    setBusy(true);
    state.adapter.replace(is, replacement)
      .then(function (n) {
        is.done = true;
        render();
        if (n) setStatus("Đã sửa " + n + " chỗ: \"" + is.display + "\" → \"" + replacement + "\".", "ok");
        else setStatus("Không còn thấy \"" + is.display + "\" trong văn bản (có thể đã được sửa).", "");
      }, function (e) { setStatus("Không sửa được: " + errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  function fixAll() {
    if (state.busy) return;
    var targets = openIssues().filter(function (i) { return !isLegal(i) && i.confident && i.suggestions.length; });
    if (!targets.length) { setStatus("Không có lỗi nào chắc chắn để sửa tự động. Hãy chọn gợi ý cho từng lỗi.", ""); return; }
    setBusy(true);
    var total = 0;
    sequence(targets, function (is) {
      var rep = is.kind === "word" ? is.suggestions[0].toLowerCase() : (is.replacement || is.suggestions[0]);
      return state.adapter.replace(is, rep).then(function (n) { total += n; is.done = true; });
    }).then(function () {
      render();
      setStatus("Đã sửa " + total + " chỗ. Các lỗi còn lại cần bạn tự chọn cách sửa.", "ok");
    }, function (e) { render(); setStatus("Có lỗi khi sửa: " + errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  function selectIssue(is) {
    state.adapter.select(is).then(function (n) {
      if (!n) setStatus("Không còn thấy \"" + is.display + "\" trong văn bản.", "");
    }, function (e) { setStatus(errText(e), "err"); });
  }

  function commentIssue(is) {
    var st = STATUS[is.status] || STATUS.can_kiem_tra;
    var text = "[Rà soát căn cứ – " + st.label + "] " + (is.reason || "");
    if (is.suggestions.length) text += " Gợi ý: " + is.suggestions[0] + ".";
    if (is.sources && is.sources.length) text += " Nguồn: " + is.sources.join(" ; ");
    state.adapter.comment(is, text).then(function (n) {
      setStatus(n ? "Đã thêm ghi chú vào văn bản." : "Không tìm thấy đoạn trích dẫn trong văn bản.", n ? "ok" : "");
    }, function (e) { setStatus(errText(e), "err"); });
  }

  function ignoreIssue(is) {
    is.done = true;
    render();
    state.adapter.clearHighlight([is]).then(null, function () {});
  }

  function addToDict(is) {
    var d = loadDict();
    d[is.search.toLowerCase()] = true;
    saveDict(d);
    ignoreIssue(is);
    updateDictInfo();
    setStatus("Đã thêm \"" + is.display + "\" vào từ điển riêng. Lần sau sẽ không báo lỗi từ này.", "ok");
  }

  function clearHighlights() {
    if (state.busy) return;
    setBusy(true);
    state.adapter.clearHighlight(state.issues)
      .then(function () { state.issues = []; render(); setStatus("Đã xóa tô màu và danh sách.", "ok"); },
        function (e) { setStatus(errText(e), "err"); })
      .then(function () { setBusy(false); });
  }

  // ------------------------------------------------------------------
  // Cài đặt
  // ------------------------------------------------------------------
  function toggleSettings(force) {
    var box = $("settings");
    var show = force === true || (force !== false && box.style.display !== "block");
    box.style.display = show ? "block" : "none";
    if (show) {
      $("apiKey").value = store(KEYS.apiKey) || "";
      $("model").value = store(KEYS.model) || "claude-haiku-5-5";
      $("legalModel").value = store(KEYS.legalModel) || "claude-sonnet-5-5";
      $("skipForeign").checked = skipForeign();
      loadFormatSettings();
      updateDictInfo();
    }
  }
  function updateDictInfo() {
    var n = Object.keys(loadDict()).length;
    $("dictInfo").textContent = "Từ điển riêng của bạn có " + n + " từ.";
  }
  function saveSettings() {
    var key = $("apiKey").value.replace(/\s/g, "");
    store(KEYS.apiKey, key || null);
    store(KEYS.model, $("model").value);
    store(KEYS.legalModel, $("legalModel").value);
    store(KEYS.skip, $("skipForeign").checked ? "1" : "0");
    saveFormatSettings();
    toggleSettings(false);
    setStatus("Đã lưu cài đặt.", "ok");
  }
  function clearDict() {
    saveDict({});
    updateDictInfo();
    setStatus("Đã xóa từ điển riêng.", "ok");
  }

  // ------------------------------------------------------------------
  // Khởi động
  // ------------------------------------------------------------------
  function start(inWord) {
    state.inWord = inWord;
    state.adapter = inWord ? WordDoc : DemoDoc;
    try {
      state.canComment = inWord && Office.context.requirements.isSetSupported("WordApi", "1.4");
    } catch (e) { state.canComment = false; }
    if (!inWord) {
      $("demo").style.display = "block";
      if (window.ManifestTool) ManifestTool.init();
    }
    $("btnQuick").onclick = runQuick;
    $("btnAI").onclick = runAI;
    $("btnLegal").onclick = runLegal;
    $("btnFmtCheck").onclick = function () { runFormatCheck(); };
    $("btnFmtApply").onclick = runFormatApply;
    $("btnFixAll").onclick = fixAll;
    $("btnClear").onclick = clearHighlights;
    $("btnSettings").onclick = function () { toggleSettings(); };
    $("btnSave").onclick = saveSettings;
    $("btnClearDict").onclick = clearDict;
    render();
  }

  if (window.Office && Office.onReady) {
    Office.onReady(function (info) { start(!!(info && info.host === Office.HostType.Word)); });
  } else {
    // office.js không tải được (đang mở ngoài Word hoặc mất mạng) -> chế độ thử
    if (!P) { document.body.innerHTML = "Trình duyệt quá cũ, hãy mở bằng Chrome hoặc Edge."; return; }
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", function () { start(false); });
    else start(false);
  }
})();
