/*
 * format-checker.js — Nhận diện các thành phần thể thức và kiểm tra cách trình bày văn bản hành chính
 *
 * Thông số mặc định theo Phụ lục I Nghị định 30/2020/NĐ-CP về công tác văn thư
 * (Công an nhân dân hiện thực hiện công tác văn thư theo Thông tư 26/2019/TT-BCA, sửa đổi bởi
 * Thông tư 35/2021/TT-BCA, trên nền Nghị định 30/2020/NĐ-CP). Mọi thông số đều chỉnh được trong DEFAULTS.
 */
(function (root) {
  "use strict";

  function nfc(s) {
    s = String(s == null ? "" : s);
    return s.normalize ? s.normalize("NFC") : s;
  }
  function clean(s) { return nfc(s).replace(/[\u00a0\t]+/g, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, ""); }
  function isUpper(s) { return /[A-ZĐÀ-Ỹ]/.test(s) && s === s.toUpperCase() && s !== s.toLowerCase(); }
  function normUpper(s) {
    return clean(s).toUpperCase().replace(/HOÀ/g, "HÒA").replace(/OẢ/g, "ỎA").replace(/UỶ/g, "ỦY").replace(/OẠ/g, "ỌA").replace(/:$/, "");
  }

  // ------------------------------------------------------------------
  // Thông số trình bày (đơn vị: cỡ chữ = pt, lề = mm, thụt dòng = cm)
  // ------------------------------------------------------------------
  var DEFAULTS = {
    font: "Times New Roman",
    bodySize: 14,          // nội dung: 13–14
    headerSize: 13,        // Quốc hiệu, tên cơ quan, số ký hiệu: 12–13
    lineSpacing: 1.15,     // từ 1 (dòng đơn) đến 1,5
    firstIndent: 1.27,     // 1 cm hoặc 1,27 cm
    spaceAfter: 6,         // tối thiểu 6pt
    margins: { top: 20, bottom: 20, left: 30, right: 15 },  // trên/dưới 20–25, trái 30–35, phải 15–20
    pageNumbers: true
  };

  var DOC_TYPES = ["NGHỊ QUYẾT", "QUYẾT ĐỊNH", "CHỈ THỊ", "QUY CHẾ", "QUY ĐỊNH", "THÔNG CÁO", "THÔNG BÁO",
    "HƯỚNG DẪN", "CHƯƠNG TRÌNH", "KẾ HOẠCH", "PHƯƠNG ÁN", "ĐỀ ÁN", "DỰ ÁN", "BÁO CÁO", "BIÊN BẢN", "TỜ TRÌNH",
    "HỢP ĐỒNG", "CÔNG ĐIỆN", "BẢN GHI NHỚ", "BẢN THỎA THUẬN", "GIẤY ỦY QUYỀN", "GIẤY MỜI", "GIẤY GIỚI THIỆU",
    "GIẤY NGHỈ PHÉP", "PHIẾU GỬI", "PHIẾU CHUYỂN", "PHIẾU BÁO", "THƯ CÔNG", "KẾT LUẬN", "LỆNH", "BẢN CAM KẾT",
    "ĐƠN", "BIÊN BẢN GIAO NHẬN"];
  var TYPE_SET = {};
  DOC_TYPES.forEach(function (t) { TYPE_SET[t] = true; });

  var RE = {
    quocHieu: /^CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM$/,
    tieuNgu: /^độc\s*lập\s*[-–—]?\s*tự\s*do\s*[-–—]?\s*hạnh\s*phúc$/i,
    so: /^số\s*:?\s*/i,
    diaDanh: /,?\s*ngày\s+(\.{2,}|…|\d{1,2}|\s*)\s*tháng\s+(\.{2,}|…|\d{1,2}|\s*)\s*năm\s+\d{4}\s*$/i,
    vv: /^v\s*\/\s*v\b/i,
    deco: /^[-_–—=.*~\s]+$/,
    canCu: /^(căn cứ|theo đề nghị|xét đề nghị|thực hiện)\s/i,
    phan: /^(phần|chương)\s+(thứ\s+\S+|[IVXLC]+)\b\.?$/i,
    muc: /^(mục|tiểu mục)\s+\d+\.?$/i,
    dieu: /^điều\s+\d+[a-zđ]?\./i,
    noiNhan: /^nơi nhận\s*:?$/i,
    kinhGui: /^kính gửi\s*:/i,
    quyen: /^(TM|KT|TL|TUQ|Q)\.\s*/,
    hocHam: /^(GS|PGS|TS|TSKH|ThS|CN|KS|BS)\.?\s+/,
    rank: /^(Đại tướng|Thượng tướng|Trung tướng|Thiếu tướng|Đại tá|Thượng tá|Trung tá|Thiếu tá|Đại úy|Thượng úy|Trung úy|Thiếu úy)\s+/,
    daKy: /^\(?(đã ký|đã kí|ký|kí)[^)]*\)?$/i
  };

  function isTitleCaseName(s) {
    var t = s.replace(RE.rank, "").replace(RE.hocHam, "");
    var words = t.split(" ");
    if (words.length < 2 || words.length > 6 || /[.;:,!?]$/.test(t) || /\d/.test(t)) return false;
    for (var i = 0; i < words.length; i++) {
      var c = words[i].charAt(0);
      if (c !== c.toUpperCase() || c === c.toLowerCase()) return false;
    }
    return !isUpper(t);
  }

  // ------------------------------------------------------------------
  // Nhận diện vai trò từng đoạn
  //   paras: [{ text, inTable }]
  // ------------------------------------------------------------------
  function classify(paras) {
    var n = paras.length, roles = [], texts = [], i, j;
    for (i = 0; i < n; i++) { texts.push(clean(paras[i].text)); roles.push(texts[i] ? "body" : "empty"); }

    // Tên loại văn bản
    var typeIdx = -1;
    for (i = 0; i < Math.min(n, 45); i++) {
      if (texts[i] && TYPE_SET[normUpper(texts[i])]) { typeIdx = i; break; }
    }

    // Phần đầu văn bản (trước tên loại; với công văn là tới dòng V/v / địa danh)
    var headerEnd = typeIdx;
    if (headerEnd < 0) {
      headerEnd = 0;
      for (i = 0; i < Math.min(n, 18); i++) {
        var tx = texts[i];
        if (RE.quocHieu.test(normUpper(tx)) || RE.tieuNgu.test(tx) || RE.so.test(tx) || RE.diaDanh.test(tx) || RE.vv.test(tx)) headerEnd = i + 1;
      }
      // V/v có thể dài 2 dòng
      while (headerEnd < n && headerEnd < 20 && texts[headerEnd] && !RE.kinhGui.test(texts[headerEnd]) &&
             paras[headerEnd].inTable && paras[headerEnd - 1] && roles[headerEnd - 1] !== "empty" && RE.vv.test(texts[headerEnd - 1] || "")) headerEnd++;
    }

    var orgs = [];
    for (i = 0; i < headerEnd; i++) {
      var t = texts[i], u = normUpper(t);
      if (!t) continue;
      if (RE.deco.test(t)) roles[i] = "deco";
      else if (RE.quocHieu.test(u)) roles[i] = "quochieu";
      else if (RE.tieuNgu.test(t)) roles[i] = "tieungu";
      else if (RE.so.test(t) && /\//.test(t)) roles[i] = "so";
      else if (RE.diaDanh.test(t)) roles[i] = "diadanh";
      else if (RE.vv.test(t)) roles[i] = "vv";
      else if (i > 0 && roles[i - 1] === "vv" && !isUpper(t)) roles[i] = "vv";
      else if (isUpper(t)) { roles[i] = "coquan"; orgs.push(i); }
      else roles[i] = "headerother";
    }
    // Nhiều dòng tên cơ quan: các dòng trước là cơ quan chủ quản, dòng cuối là cơ quan ban hành
    if (orgs.length > 1) for (j = 0; j < orgs.length - 1; j++) roles[orgs[j]] = "chuquan";

    var start = headerEnd;
    if (typeIdx >= 0) {
      roles[typeIdx] = "tenloai";
      start = typeIdx + 1;
      // Trích yếu: 1–2 đoạn ngay sau tên loại
      var k = typeIdx + 1, got = 0;
      while (k < n && got < 2 && texts[k] && !RE.canCu.test(texts[k]) && !RE.kinhGui.test(texts[k]) &&
             !isUpper(texts[k]) && texts[k].length < 250 && !RE.deco.test(texts[k])) {
        roles[k] = "trichyeu"; got++; k++;
      }
      while (k < n && (!texts[k] || RE.deco.test(texts[k]))) { if (texts[k]) roles[k] = "deco"; k++; }
      // Thẩm quyền ban hành (VD: "GIÁM ĐỐC CÔNG AN TỈNH ...") đứng trước căn cứ
      if (k < n && isUpper(texts[k]) && texts[k].length < 160) {
        var m = k + 1;
        while (m < n && !texts[m]) m++;
        if (m < n && RE.canCu.test(texts[m])) { roles[k] = "thamquyen"; k = m; }
      }
      start = k;
    }

    // Thân văn bản
    var inNoiNhan = false, sigZone = false, lastSig = -1;
    for (i = start; i < n; i++) {
      var s = texts[i];
      if (!s) continue;
      if (RE.deco.test(s)) { roles[i] = "deco"; continue; }
      if (RE.noiNhan.test(s)) { roles[i] = "noinhan"; inNoiNhan = true; sigZone = true; continue; }
      if (inNoiNhan && /^[-+•]/.test(s)) { roles[i] = "noinhan_item"; continue; }
      inNoiNhan = false;
      if (RE.kinhGui.test(s)) { roles[i] = "kinhgui"; continue; }
      if (RE.canCu.test(s) && !sigZone) { roles[i] = "cancu"; continue; }
      if (RE.phan.test(s)) { roles[i] = "phan"; markTitle(i); continue; }
      if (RE.muc.test(s)) { roles[i] = "muc"; markTitle(i); continue; }
      if (RE.dieu.test(s)) { roles[i] = "dieu"; continue; }
      if (/^[A-ZĐÀ-Ỹ ]+:$/.test(s) && i > start && roles.slice(start, i).indexOf("cancu") >= 0 && !sigZone) { roles[i] = "label"; continue; }
      // Khối chữ ký: thường ở cuối, sau nơi nhận hoặc trong bảng
      var nearEnd = i >= n - 18;
      if (RE.quyen.test(s) && s.length < 120) { roles[i] = "ky_quyen"; sigZone = true; lastSig = i; continue; }
      if ((sigZone || nearEnd) && isUpper(s) && s.split(" ").length <= 12 && !/[.;]$/.test(s) && (paras[i].inTable || sigZone)) {
        roles[i] = "ky_quyen"; sigZone = true; lastSig = i; continue;
      }
      if (lastSig >= 0 && i - lastSig <= 8 && (RE.daKy.test(s) || /^\[.*\]$/.test(s))) { roles[i] = "ky_other"; continue; }
      if (lastSig >= 0 && i - lastSig <= 8 && isTitleCaseName(s)) { roles[i] = "ky_ten"; lastSig = -1; continue; }
    }

    function markTitle(idx) {
      for (var q = idx + 1; q < n && q <= idx + 3; q++) {
        if (!texts[q]) continue;
        if (!RE.dieu.test(texts[q]) && texts[q].length < 200) roles[q] = "tieude";
        return;
      }
    }
    return roles;
  }

  // ------------------------------------------------------------------
  // Định dạng cho từng vai trò
  //   align: Left | Centered | Right | Justified ; indent: thụt dòng đầu (cm)
  // ------------------------------------------------------------------
  function formatFor(role, inTable, st) {
    st = st || DEFAULTS;
    var B = st.bodySize, H = st.headerSize, I = st.firstIndent;
    var body = { size: B, align: "Justified", indent: inTable ? 0 : I, after: st.spaceAfter, line: st.lineSpacing };
    switch (role) {
      case "quochieu": return { size: H, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "tieungu": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "chuquan": return { size: H, bold: false, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "coquan": return { size: H, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "so": return { size: 13, bold: false, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "diadanh": return { size: B, bold: false, italic: true, align: inTable ? "Centered" : "Right", indent: 0, after: 0, line: 1 };
      case "vv": return { size: H, bold: false, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "deco": return { align: "Centered", indent: 0, after: 0 };
      case "headerother": return { size: H, align: "Centered", indent: 0, after: 0, line: 1 };
      case "tenloai": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1, before: 6 };
      case "trichyeu": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "thamquyen": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 6, line: 1, before: 12 };
      case "label": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 6, line: 1 };
      case "kinhgui": return { size: B, bold: false, italic: false, align: "Centered", indent: 0, after: 6, line: 1, before: 6 };
      case "cancu": return { size: B, bold: false, italic: true, align: "Justified", indent: I, after: st.spaceAfter, line: st.lineSpacing };
      case "phan": case "muc": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1, before: 6 };
      case "tieude": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 6, line: 1 };
      case "dieu": return { size: B, bold: true, italic: false, align: "Justified", indent: I, after: st.spaceAfter, line: st.lineSpacing };
      case "noinhan": return { size: 12, bold: true, italic: true, align: "Left", indent: 0, after: 0, line: 1, left: 0 };
      case "noinhan_item": return { size: 11, bold: false, italic: false, align: "Left", indent: 0, after: 0, line: 1, left: 0 };
      case "ky_quyen": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "ky_ten": return { size: B, bold: true, italic: false, align: "Centered", indent: 0, after: 0, line: 1 };
      case "ky_other": return { size: B, align: "Centered", indent: 0, after: 0 };
      case "empty": return { size: B, indent: 0 };
      default: return body;
    }
  }

  // ------------------------------------------------------------------
  // Kiểm tra lỗi trình bày có thể phát hiện qua chữ
  // ------------------------------------------------------------------
  function check(paras, roles) {
    roles = roles || classify(paras);
    var issues = [], i;
    function add(status, idx, display, reason, suggestions) {
      issues.push({ kind: "format", status: status, display: display, search: display && display.length <= 250 ? display : null,
        pos: idx, reason: reason, suggestions: suggestions || [] });
    }
    var texts = paras.map(function (p) { return clean(p.text); });
    var raw = paras.map(function (p) { return nfc(p.text).replace(/[\r\n\u000b]+$/, ""); });
    var hasRole = function (r) { return roles.indexOf(r) >= 0; };

    // Quốc hiệu
    var qhFound = false;
    for (i = 0; i < Math.min(texts.length, 20); i++) {
      if (/^cộng hòa xã hội chủ nghĩa việt nam$/i.test(normUpper(texts[i]).toLowerCase())) {
        qhFound = true;
        if (!RE.quocHieu.test(clean(texts[i]))) {
          add("sai", i, raw[i].trim(), "Quốc hiệu phải viết bằng chữ in hoa: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM.", ["CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"]);
        }
      }
    }
    if (!qhFound && (hasRole("tenloai") || hasRole("so"))) add("luu_y", 0, null, "Chưa thấy Quốc hiệu \"CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\" ở đầu văn bản.");

    for (i = 0; i < texts.length; i++) {
      var t = texts[i], r = roles[i], line = raw[i].trim();
      if (!t) continue;

      if (r === "tieungu" && t !== "Độc lập - Tự do - Hạnh phúc") {
        add("sai", i, line, "Tiêu ngữ viết đúng là \"Độc lập - Tự do - Hạnh phúc\": chữ in thường, giữa các cụm từ có gạch nối (-) và cách chữ.", ["Độc lập - Tự do - Hạnh phúc"]);
      }

      if (r === "so") {
        var ms = /^(số)\s*(:?)\s*(\d+)(\s*\/)/i.exec(t);
        if (ms) {
          var fixed = "Số: " + (ms[3].length === 1 ? "0" + ms[3] : ms[3]) + "/";
          var orig = /^số\s*:?\s*\d+\s*\//i.exec(line);
          if (orig && orig[0] !== fixed) {
            var why = [];
            if (ms[1] !== "Số" || !ms[2]) why.push("viết \"Số:\" (chữ S hoa, có dấu hai chấm)");
            if (ms[3].length === 1) why.push("số nhỏ hơn 10 phải có số 0 phía trước");
            if (why.length) add("sai", i, orig[0], "Số văn bản: " + why.join("; ") + ".", [fixed]);
          }
        }
        var sym = /\/\s*([A-Za-zĐđ0-9-]+(?:\s*-\s*[A-Za-zĐđ0-9]+)*)\s*$/.exec(t);
        if (sym && (sym[1] !== sym[1].toUpperCase() || /\s/.test(sym[1]))) {
          add("canh_bao", i, sym[1], "Ký hiệu văn bản viết chữ in hoa, giữa các nhóm chữ có gạch nối (-), không cách chữ.", [sym[1].toUpperCase().replace(/\s+/g, "")]);
        }
      }

      if (r === "diadanh") {
        var md = /ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/i.exec(line);
        if (md) {
          var d = md[1], mo = md[2];
          var nd = (parseInt(d, 10) < 10 && d.length === 1) ? "0" + d : d;
          var nmo = mo, mi = parseInt(mo, 10);
          if ((mi === 1 || mi === 2) && mo.length === 1) nmo = "0" + mo;
          if (mi >= 3 && mo.length === 2 && mo.charAt(0) === "0") nmo = String(mi);
          if (nd !== d || nmo !== mo) {
            add("sai", i, md[0], "Ngày nhỏ hơn 10 và tháng 1, 2 phải có số 0 phía trước; các tháng còn lại không thêm số 0.",
              ["ngày " + nd + " tháng " + nmo + " năm " + md[3]]);
          }
        }
        var mc = /^([^,]*?\S)\s+ngày\s/.exec(line);
        if (mc && !/,$/.test(mc[1])) add("sai", i, mc[0], "Sau địa danh phải có dấu phẩy.", [mc[1] + ", ngày "]);
      }

      if (r === "vv") {
        var mv = /^v\s*\/\s*v\s*:?\s*/i.exec(line);
        if (mv && mv[0].replace(/\s+$/, "") !== "V/v") add("sai", i, mv[0].replace(/\s+$/, ""), "Trích yếu công văn bắt đầu bằng \"V/v\" (viết đúng chữ hoa/thường, không có dấu hai chấm).", ["V/v"]);
      }

      if (r === "cancu") {
        var ml = /^(căn cứ\s+(?:bộ luật|luật|pháp lệnh)\s[^;]*?)(\s+số\s*:?\s*\d+\/\d{4}\/[A-ZĐ0-9-]+)/i.exec(line);
        if (ml) add("luu_y", i, ml[2].replace(/^\s+/, ""), "Căn cứ là Luật, Pháp lệnh thì chỉ ghi tên và ngày thông qua, không ghi số, ký hiệu (Phụ lục I Nghị định 30/2020/NĐ-CP).");
      }

      if (r === "ky_ten" && RE.hocHam.test(t)) {
        var mh = RE.hocHam.exec(line.trim());
        add("sai", i, mh[0].replace(/\s+$/, ""), "Không ghi học hàm, học vị trước họ tên người ký.");
      }
    }

    // Dấu câu cuối mỗi căn cứ: chấm phẩy, căn cứ cuối cùng: dấu chấm
    var groups = [], cur = [];
    for (i = 0; i < roles.length; i++) {
      if (roles[i] === "cancu") cur.push(i);
      else if (roles[i] !== "empty" && cur.length) { groups.push(cur); cur = []; }
    }
    if (cur.length) groups.push(cur);
    groups.forEach(function (g) {
      g.forEach(function (idx, k) {
        var line = raw[idx].replace(/\s+$/, "");
        var want = k === g.length - 1 ? "." : ";";
        var last = line.charAt(line.length - 1);
        if (last === want) return;
        var tail = line.slice(-Math.min(40, line.length));
        var fixedTail = /[.;,:]$/.test(tail) ? tail.slice(0, -1) + want : tail + want;
        add("sai", idx, tail, want === ";" ? "Cuối mỗi căn cứ phải có dấu chấm phẩy (;)." : "Căn cứ cuối cùng kết thúc bằng dấu chấm (.).", [fixedTail]);
      });
    });

    // Nơi nhận
    var nn = roles.indexOf("noinhan");
    if (nn < 0 && (hasRole("tenloai") || hasRole("vv"))) {
      add("luu_y", texts.length, null, "Chưa thấy phần \"Nơi nhận:\" ở cuối văn bản.");
    } else if (nn >= 0) {
      var items = [];
      for (i = nn + 1; i < roles.length; i++) { if (roles[i] === "noinhan_item") items.push(i); else if (roles[i] !== "empty") break; }
      if (items.length) {
        var lastItem = items[items.length - 1];
        if (!/^[-+•]\s*Lưu\s*:/i.test(texts[lastItem])) add("luu_y", lastItem, null, "Dòng cuối của Nơi nhận nên là \"- Lưu: VT, ...\".");
        items.forEach(function (idx, k) {
          var ln = raw[idx].replace(/\s+$/, ""), want = k === items.length - 1 ? "." : ";";
          if (ln.charAt(ln.length - 1) !== want) {
            var tl = ln.slice(-Math.min(30, ln.length));
            add("canh_bao", idx, tl, want === ";" ? "Mỗi nơi nhận kết thúc bằng dấu chấm phẩy (;)." : "Dòng cuối của Nơi nhận kết thúc bằng dấu chấm (.).",
              [/[.;,:]$/.test(tl) ? tl.slice(0, -1) + want : tl + want]);
          }
        });
      }
    }
    issues.sort(function (a, b) { return a.pos - b.pos; });
    return issues;
  }

  var api = { classify: classify, formatFor: formatFor, check: check, DEFAULTS: DEFAULTS, clean: clean };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FormatChecker = api;
})(this);
