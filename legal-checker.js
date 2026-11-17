/*
 * legal-checker.js — Kiểm tra HÌNH THỨC các trích dẫn văn bản pháp luật (chạy offline, miễn phí)
 *
 * Kiểm tra được:
 *   - Số hiệu văn bản: năm ở tương lai, khóa Quốc hội không khớp năm (VD: 45/2019/QH13 -> phải là QH14)
 *   - Số hiệu không khớp tên luật (VD: "Bộ luật Dân sự số 92/2015/QH13" — 92 là Bộ luật Tố tụng dân sự)
 *   - Năm trong tên không khớp năm trong số hiệu / ngày ban hành
 *   - Dẫn luật đã hết hiệu lực (VD: Luật Đất đai 2013) hoặc năm không tồn tại (VD: Bộ luật Dân sự 2016)
 *   - Ngày tháng không có thật (VD: ngày 30 tháng 02)
 *   - Tên luật chưa ghi năm ban hành hoặc số hiệu
 *
 * Nội dung điều, khoản và hiệu lực của văn bản khác được kiểm tra bằng AI (xem app.js).
 * Bảng dữ liệu bên dưới cập nhật đến tháng 11/2026 — có thể tự bổ sung.
 */
(function (root) {
  "use strict";

  function nfc(s) {
    s = String(s == null ? "" : s);
    return s.normalize ? s.normalize("NFC") : s;
  }
  var LETTER = /[A-Za-zÀ-ɏẠ-ỹ̀-ͯ]/;
  function isLetter(ch) { return !!ch && LETTER.test(ch); }

  // ------------------------------------------------------------------
  // 1. Dữ liệu: các luật hay được viện dẫn
  //    versions: các năm đã ban hành; complete = true nghĩa là danh sách đầy đủ
  // ------------------------------------------------------------------
  var LAWS = {
    "bộ luật dân sự": { name: "Bộ luật Dân sự", versions: [1995, 2005, 2015], complete: true },
    "bộ luật hình sự": { name: "Bộ luật Hình sự", versions: [1985, 1999, 2015], complete: true },
    "bộ luật tố tụng dân sự": { name: "Bộ luật Tố tụng dân sự", versions: [2004, 2015], complete: true },
    "bộ luật tố tụng hình sự": { name: "Bộ luật Tố tụng hình sự", versions: [1988, 2003, 2015], complete: true },
    "bộ luật lao động": { name: "Bộ luật Lao động", versions: [1994, 2012, 2019], complete: true },
    "luật hôn nhân và gia đình": { name: "Luật Hôn nhân và gia đình", versions: [1959, 1986, 2000, 2014], complete: true },
    "luật đất đai": { name: "Luật Đất đai", versions: [1987, 1993, 2003, 2013, 2024], complete: true },
    "luật nhà ở": { name: "Luật Nhà ở", versions: [2005, 2014, 2023] },
    "luật kinh doanh bất động sản": { name: "Luật Kinh doanh bất động sản", versions: [2006, 2014, 2023] },
    "luật doanh nghiệp": { name: "Luật Doanh nghiệp", versions: [1999, 2005, 2014, 2020], not: /^\s+nhà nước/ },
    "luật đầu tư": { name: "Luật Đầu tư", versions: [2005, 2014, 2020, 2025], not: /^\s+(công|theo phương thức)/ },
    "luật công chứng": { name: "Luật Công chứng", versions: [2006, 2014, 2024] },
    "luật ban hành văn bản quy phạm pháp luật": { name: "Luật Ban hành văn bản quy phạm pháp luật", versions: [2008, 2015, 2025], not: /^\s+của/ },
    "luật cư trú": { name: "Luật Cư trú", versions: [2006, 2020] },
    "luật căn cước công dân": { name: "Luật Căn cước công dân", versions: [2014] },
    "luật căn cước": { name: "Luật Căn cước", versions: [2023], not: /^\s+công dân/ },
    "luật tố cáo": { name: "Luật Tố cáo", versions: [2011, 2018] },
    "luật tố tụng hành chính": { name: "Luật Tố tụng hành chính", versions: [2010, 2015] },
    "luật hộ tịch": { name: "Luật Hộ tịch", versions: [2014] },
    "luật xử lý vi phạm hành chính": { name: "Luật Xử lý vi phạm hành chính", versions: [2012] },
    "luật thương mại": { name: "Luật Thương mại", versions: [1997, 2005], not: /^\s+điện tử/ },
    "luật tổ chức quốc hội": { name: "Luật Tổ chức Quốc hội", versions: [2014] }
  };

  // Văn bản đã bị thay thế: "tên|năm" -> văn bản thay thế và ngày có hiệu lực của văn bản mới
  var REPLACED = {
    "bộ luật dân sự|1995": { by: "bộ luật dân sự|2005", from: "01/01/2006" },
    "bộ luật dân sự|2005": { by: "bộ luật dân sự|2015", from: "01/01/2017" },
    "bộ luật hình sự|1999": { by: "bộ luật hình sự|2015", from: "01/01/2018" },
    "bộ luật tố tụng dân sự|2004": { by: "bộ luật tố tụng dân sự|2015", from: "01/7/2016" },
    "bộ luật tố tụng hình sự|2003": { by: "bộ luật tố tụng hình sự|2015", from: "01/01/2018" },
    "bộ luật lao động|2012": { by: "bộ luật lao động|2019", from: "01/01/2021" },
    "luật hôn nhân và gia đình|2000": { by: "luật hôn nhân và gia đình|2014", from: "01/01/2015" },
    "luật đất đai|2003": { by: "luật đất đai|2013", from: "01/7/2014" },
    "luật đất đai|2013": { by: "luật đất đai|2024", from: "01/8/2024" },
    "luật nhà ở|2014": { by: "luật nhà ở|2023", from: "01/8/2024" },
    "luật kinh doanh bất động sản|2014": { by: "luật kinh doanh bất động sản|2023", from: "01/8/2024" },
    "luật doanh nghiệp|2014": { by: "luật doanh nghiệp|2020", from: "01/01/2021" },
    "luật đầu tư|2014": { by: "luật đầu tư|2020", from: "01/01/2021" },
    "luật đầu tư|2020": { by: "luật đầu tư|2025", from: "01/3/2026" },
    "luật công chứng|2014": { by: "luật công chứng|2024", from: "01/7/2025" },
    "luật ban hành văn bản quy phạm pháp luật|2015": { by: "luật ban hành văn bản quy phạm pháp luật|2025", from: "01/4/2025" },
    "luật cư trú|2006": { by: "luật cư trú|2020", from: "01/7/2021" },
    "luật căn cước công dân|2014": { by: "luật căn cước|2023", from: "01/7/2024" },
    "luật tố cáo|2011": { by: "luật tố cáo|2018", from: "01/01/2019" }
  };

  // Số hiệu đã biết -> "tên|năm"
  var NUMBERS = {
    "33/2005/QH11": "bộ luật dân sự|2005", "91/2015/QH13": "bộ luật dân sự|2015",
    "100/2015/QH13": "bộ luật hình sự|2015", "92/2015/QH13": "bộ luật tố tụng dân sự|2015",
    "101/2015/QH13": "bộ luật tố tụng hình sự|2015", "10/2012/QH13": "bộ luật lao động|2012",
    "45/2019/QH14": "bộ luật lao động|2019", "52/2014/QH13": "luật hôn nhân và gia đình|2014",
    "45/2013/QH13": "luật đất đai|2013", "31/2024/QH15": "luật đất đai|2024",
    "65/2014/QH13": "luật nhà ở|2014", "27/2023/QH15": "luật nhà ở|2023",
    "66/2014/QH13": "luật kinh doanh bất động sản|2014", "29/2023/QH15": "luật kinh doanh bất động sản|2023",
    "68/2014/QH13": "luật doanh nghiệp|2014", "59/2020/QH14": "luật doanh nghiệp|2020",
    "67/2014/QH13": "luật đầu tư|2014", "61/2020/QH14": "luật đầu tư|2020", "143/2025/QH15": "luật đầu tư|2025",
    "53/2014/QH13": "luật công chứng|2014", "46/2024/QH15": "luật công chứng|2024",
    "80/2015/QH13": "luật ban hành văn bản quy phạm pháp luật|2015",
    "64/2025/QH15": "luật ban hành văn bản quy phạm pháp luật|2025",
    "68/2020/QH14": "luật cư trú|2020", "59/2014/QH13": "luật căn cước công dân|2014",
    "26/2023/QH15": "luật căn cước|2023", "3/2011/QH13": "luật tố cáo|2011", "25/2018/QH14": "luật tố cáo|2018",
    "93/2015/QH13": "luật tố tụng hành chính|2015", "60/2014/QH13": "luật hộ tịch|2014",
    "15/2012/QH13": "luật xử lý vi phạm hành chính|2012", "36/2005/QH11": "luật thương mại|2005",
    "57/2014/QH13": "luật tổ chức quốc hội|2014"
  };
  var NUMBER_OF = {};
  Object.keys(NUMBERS).forEach(function (n) { NUMBER_OF[NUMBERS[n]] = n.replace(/^(\d)\//, "0$1/"); });

  // Nhiệm kỳ Quốc hội các khóa
  var KHOA = { 10: [1997, 2002], 11: [2002, 2007], 12: [2007, 2011], 13: [2011, 2016],
    14: [2016, 2021], 15: [2021, 2026], 16: [2026, 2031] };
  var ROMAN = { 10: "X", 11: "XI", 12: "XII", 13: "XIII", 14: "XIV", 15: "XV", 16: "XVI" };

  function lawLabel(key) {
    var p = key.split("|");
    var law = LAWS[p[0]];
    var label = (law ? law.name : p[0]) + " " + p[1];
    if (NUMBER_OF[key]) label += " (số " + NUMBER_OF[key] + ")";
    return label;
  }
  function finalReplacement(key) {
    var chain = [], cur = key, guard = 0;
    while (REPLACED[cur] && guard++ < 10) { chain.push({ key: REPLACED[cur].by, from: REPLACED[cur].from }); cur = REPLACED[cur].by; }
    return chain;
  }
  function normNumber(n) {
    var p = n.split("/");
    return parseInt(p[0], 10) + "/" + p[1] + "/" + p.slice(2).join("/");
  }
  function khoaFor(year) {
    var out = [];
    Object.keys(KHOA).forEach(function (k) { if (year >= KHOA[k][0] && year <= KHOA[k][1]) out.push(parseInt(k, 10)); });
    return out;
  }
  function daysInMonth(m, y) { return [31, (y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0)) ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]; }

  // ------------------------------------------------------------------
  // 2. Kiểm tra
  // ------------------------------------------------------------------
  var NUM_RE = /(\d{1,4}(?:\.\d+)?)\/(\d{4})\/([A-ZĐ][A-ZĐ0-9]*(?:-[A-ZĐ0-9]+)*)/g;

  function check(text, opts) {
    opts = opts || {};
    var thisYear = opts.year || new Date().getFullYear();
    text = nfc(text);
    var lower = text.toLowerCase();
    var issues = [];
    function add(status, start, end, reason, suggestions) {
      var span = text.slice(start, end);
      for (var i = 0; i < issues.length; i++) if (issues[i].display === span && issues[i].reason === reason) return;
      issues.push({ kind: "legal", status: status, display: span, search: span, pos: start,
        reason: reason, suggestions: suggestions || [] });
    }
    var m;

    // --- 2a. Số hiệu văn bản ---
    NUM_RE.lastIndex = 0;
    while ((m = NUM_RE.exec(text)) !== null) {
      var start = m.index, end = start + m[0].length;
      if (isLetter(text.charAt(end)) || /[\d\/]/.test(text.charAt(start - 1) || "")) continue;
      var year = parseInt(m[2], 10), suffix = m[3];
      if (year > thisYear) { add("sai", start, end, "Năm ban hành " + year + " trong số hiệu ở tương lai."); continue; }
      if (year < 1945) { add("sai", start, end, "Năm " + year + " không hợp lý cho văn bản pháp luật Việt Nam."); continue; }

      var qh = /(?:^|-)(?:UBTV)?QH(\d+)$/.exec(suffix);
      if (qh) {
        var k = parseInt(qh[1], 10), ok = khoaFor(year);
        if (KHOA[k] && ok.length && ok.indexOf(k) < 0) {
          var fix = ok.length === 1 ? [m[0].replace(/QH\d+$/, "QH" + ok[0])] : [];
          add("sai", start, end, "Năm " + year + " thuộc Quốc hội khóa " + ok.map(function (x) { return ROMAN[x]; }).join("/") +
            ", nhưng số hiệu ghi QH" + k + " (khóa " + ROMAN[k] + ", " + KHOA[k][0] + "–" + KHOA[k][1] + ").", fix);
          continue;
        }
      }

      // Đoạn chữ ngay trước số hiệu (để biết đang nói luật nào)
      var winStart = Math.max(0, start - 140);
      var win = lower.slice(winStart, start);
      var cut = Math.max(win.lastIndexOf(";"), win.lastIndexOf("\n"), win.lastIndexOf("\r"), win.lastIndexOf(". "));
      if (cut >= 0) win = win.slice(cut + 1);
      var kw = /(bộ luật|luật|pháp lệnh|nghị định|thông tư|nghị quyết|quyết định)\s+(?!số[\s:])/g, km, nameText = "";
      while ((km = kw.exec(win)) !== null) nameText = win.slice(km.index);

      var known = NUMBERS[normNumber(m[0])];
      if (known && nameText && /^(bộ luật|luật)\s/.test(nameText)) {
        var title = known.split("|")[0];
        if (nameText.indexOf(title) < 0) {
          var sugg = [];
          Object.keys(LAWS).forEach(function (t) {
            if (nameText.indexOf(t) === 0 && t !== title) {
              var cand = t + "|" + year;
              if (NUMBER_OF[cand]) sugg.push(NUMBER_OF[cand]);
            }
          });
          add("sai", start, end, "Số hiệu " + m[0] + " là của " + LAWS[title].name + " " + known.split("|")[1] + ", không khớp với tên văn bản đang ghi.", sugg);
          continue;
        }
      }
      // Năm trong tên / ngày ban hành khác năm trong số hiệu
      var ny = /(?:năm|\s)\s*(\d{4})\s*(?:,\s*)?(?:số\s*:?\s*)?$/.exec(nameText);
      if (ny && parseInt(ny[1], 10) !== year) {
        add("sai", start, end, "Năm ghi trong tên (" + ny[1] + ") khác năm trong số hiệu (" + year + ").");
        continue;
      }
      var after = text.slice(end, end + 60);
      var dy = /^\s*,?\s*ngày\s+\d{1,2}\s*(?:tháng\s*\d{1,2}\s*năm\s*|\/\s*\d{1,2}\s*\/\s*)(\d{4})/i.exec(after);
      if (dy && parseInt(dy[1], 10) !== year) {
        add("canh_bao", start, end + dy[0].length, "Năm ban hành theo ngày (" + dy[1] + ") khác năm trong số hiệu (" + year + "). Kiểm tra lại.");
      }
    }

    // --- 2b. Tên luật + năm ---
    var titles = Object.keys(LAWS).sort(function (a, b) { return b.length - a.length; });
    var used = [], mentions = {};
    function overlaps(s, e) { for (var i = 0; i < used.length; i++) if (s < used[i][1] && e > used[i][0]) return true; return false; }
    titles.forEach(function (t) {
      var from = 0, idx;
      while ((idx = lower.indexOf(t, from)) >= 0) {
        from = idx + t.length;
        var e = idx + t.length;
        if (isLetter(text.charAt(idx - 1)) || isLetter(text.charAt(e)) || overlaps(idx, e)) continue;
        var law = LAWS[t];
        var rest = lower.slice(e, e + 90);
        if (law.not && law.not.test(rest)) continue;
        used.push([idx, e]);
        var year = null, extra = 0, mm;
        if ((mm = /^\s*(?:,\s*)?(?:năm\s+)?(\d{4})(?![\d\/])/.exec(rest))) { year = +mm[1]; extra = mm[0].length; }
        else if ((mm = /^\s*(?:,\s*)?(?:số\s*:?\s*)?(\d{1,4}\/(\d{4})\/[a-zđ0-9-]+)/.exec(rest))) { year = +mm[2]; extra = mm[0].length; }
        else if ((mm = /^\s*(?:,\s*)?(?:được\s+(?:quốc hội\s+)?thông qua\s+)?ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+(\d{4})/.exec(rest))) { year = +mm[1]; extra = mm[0].length; }
        if (!mentions[t]) mentions[t] = { withYear: false, first: idx, firstEnd: e };
        if (year) mentions[t].withYear = true;
        if (!year) continue;

        var key = t + "|" + year;
        if (REPLACED[key]) {
          var chain = finalReplacement(key), last = chain[chain.length - 1];
          var reason = law.name + " " + year + " đã hết hiệu lực từ " + chain[0].from + ", được thay thế bởi " + lawLabel(chain[0].key) + ".";
          for (var c = 1; c < chain.length; c++) reason += " Văn bản này lại được thay thế bởi " + lawLabel(chain[c].key) + " từ " + chain[c].from + ".";
          reason += " Nếu sự việc xảy ra trước thời điểm thay thế thì việc dẫn luật cũ có thể vẫn đúng.";
          var span = text.slice(idx, e + extra);
          var newTitle = last.key.split("|")[0], newYear = last.key.split("|")[1];
          var sug = [];
          if (newTitle === t && /\d{4}$/.test(span) && !/\//.test(span)) sug.push(span.replace(/\d{4}$/, newYear));
          else if (newTitle !== t && !/\//.test(span)) sug.push(LAWS[newTitle].name + " " + newYear);
          add("het_hieu_luc", idx, e + extra, reason, sug);
        } else if (law.complete && law.versions.indexOf(year) < 0) {
          var prev = null;
          law.versions.forEach(function (v) { if (v < year) prev = v; });
          var r2 = "Không có " + law.name + " năm " + year + ". Các năm đã ban hành: " + law.versions.join(", ") + ".";
          if (prev) r2 += " Nếu ý bạn là văn bản sửa đổi năm " + year + ", nên ghi \"" + law.name + " " + prev + " (sửa đổi, bổ sung năm " + year + ")\".";
          var span2 = text.slice(idx, e + extra);
          add("sai", idx, e + extra, r2, prev && !/\//.test(span2) ? [span2.replace(/\d{4}$/, String(prev))] : []);
        }
      }
    });
    Object.keys(mentions).forEach(function (t) {
      var mt = mentions[t];
      if (!mt.withYear) {
        add("luu_y", mt.first, mt.firstEnd, "Lần đầu nhắc tới " + LAWS[t].name + " nên ghi rõ năm ban hành hoặc số hiệu để tránh hiểu nhầm phiên bản.");
      }
    });

    // --- 2c. Ngày tháng không có thật ---
    var DATE_RE = /ngày\s+(\d{1,2})\s*(?:tháng\s*(\d{1,2})\s*năm\s*|\/\s*(\d{1,2})\s*\/\s*)(\d{4})/gi;
    while ((m = DATE_RE.exec(text)) !== null) {
      var d = +m[1], mo = +(m[2] || m[3]), y = +m[4];
      if (mo < 1 || mo > 12 || d < 1 || d > daysInMonth(mo, y)) {
        add("sai", m.index, m.index + m[0].length, "Ngày tháng không có thật: " + m[0] + ".");
      }
    }

    issues.sort(function (a, b) { return a.pos - b.pos; });
    return issues;
  }

  // ------------------------------------------------------------------
  // 3. Tìm các câu có trích dẫn để gửi cho AI rà soát nội dung
  // ------------------------------------------------------------------
  var DOC_KW = /(Hiến pháp|Bộ luật|Luật|Pháp lệnh|Nghị định|Nghị quyết|Thông tư|Quyết định)(?!\s+(này|sư\b))/;
  function citationSentences(text, max) {
    text = nfc(text);
    max = max || 40;
    var out = [], seen = {};
    var lines = text.split(/\r\n|\r|\n|\u000b/);
    lines.forEach(function (line) {
      var parts = line.replace(/([.;!?])\s+(?=[A-ZÀ-ỸĐ"“(\[])/g, "$1\u0001").split("\u0001");
      parts.forEach(function (s) {
        s = s.replace(/^\s+|\s+$/g, "");
        if (s.length < 8) return;
        var hasNumber = /\d+\/\d{4}\/[A-ZĐ]/.test(s);
        var canCu = /^căn cứ\s/i.test(s);
        var article = /(Điều|điều)\s+\d+/.test(s) && DOC_KW.test(s);
        var knownTitle = /(Bộ luật|Luật|Pháp lệnh)\s+[^\d;.]{2,80}?\d{4}/.test(s) && DOC_KW.test(s);
        if (!(hasNumber || canCu || article || knownTitle)) return;
        if (s.length > 700) s = s.slice(0, 700);
        if (seen[s]) return;
        seen[s] = true;
        if (out.length < max) out.push(s);
      });
    });
    return out;
  }

  var api = { check: check, citationSentences: citationSentences, LAWS: LAWS, REPLACED: REPLACED, NUMBERS: NUMBERS };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.LegalChecker = api;
})(this);
