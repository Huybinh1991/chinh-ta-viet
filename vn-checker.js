/*
 * vn-checker.js — Bộ kiểm tra chính tả tiếng Việt (chạy offline, miễn phí)
 *
 * Cách hoạt động:
 *   1. Tách văn bản thành từng tiếng (âm tiết).
 *   2. Kiểm tra mỗi tiếng có ghép đúng luật tiếng Việt không:
 *      phụ âm đầu + vần hợp lệ, luật c/k, g/gh, ng/ngh, dấu thanh với vần p/t/c/ch,
 *      không có hai dấu thanh trong một tiếng.
 *   3. Phát hiện lỗi gõ Telex chưa chuyển dấu (VD: "vieejt" -> "việt").
 *   4. Dò các cụm từ hay viết sai (VD: "chia sẽ" -> "chia sẻ").
 *
 * Viết bằng JavaScript đời cũ (ES5) để chạy được cả trên Word 2016/2019.
 */
(function (root) {
  "use strict";

  // ---------------------------------------------------------------
  // 1. Bảng dấu thanh: 0 ngang, 1 huyền, 2 sắc, 3 hỏi, 4 ngã, 5 nặng
  // ---------------------------------------------------------------
  var TONE_TABLE = {
    "a": "aàáảãạ", "ă": "ăằắẳẵặ", "â": "âầấẩẫậ",
    "e": "eèéẻẽẹ", "ê": "êềếểễệ", "i": "iìíỉĩị",
    "o": "oòóỏõọ", "ô": "ôồốổỗộ", "ơ": "ơờớởỡợ",
    "u": "uùúủũụ", "ư": "ưừứửữự", "y": "yỳýỷỹỵ"
  };
  var TONE_NAMES = ["không dấu", "huyền", "sắc", "hỏi", "ngã", "nặng"];

  var CHAR_INFO = {};
  Object.keys(TONE_TABLE).forEach(function (base) {
    var s = TONE_TABLE[base];
    for (var t = 0; t < 6; t++) CHAR_INFO[s.charAt(t)] = { base: base, tone: t };
  });

  var VOWELS = "aăâeêioôơuưy";
  var MARKED_VOWELS = "ăâêôơư"; // nguyên âm có mũ/móc
  var VN_SPECIAL = /[ăâêôơưđàáảãạằắẳẵặầấẩẫậèéẻẽẹềếểễệìíỉĩịòóỏõọồốổỗộờớởỡợùúủũụừứửữựỳýỷỹỵ]/;

  // Các chữ cái dùng để tách từ (Latin + tiếng Việt + dấu tổ hợp)
  var LETTER_CLASS = "A-Za-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u017F\\u01A0\\u01A1\\u01AF\\u01B0\\u0300-\\u036F\\u1EA0-\\u1EF9";
  var WORD_RE_SRC = "[" + LETTER_CLASS + "]+";

  function nfc(s) {
    s = String(s == null ? "" : s);
    return s.normalize ? s.normalize("NFC") : s;
  }
  function isVowel(c) { return VOWELS.indexOf(c) >= 0; }

  // Tách dấu thanh: "việt" -> { base: "viêt", tone: 5 }
  function stripTone(word) {
    var base = "", tone = 0, tones = [];
    for (var i = 0; i < word.length; i++) {
      var c = word.charAt(i), info = CHAR_INFO[c];
      if (info) {
        base += info.base;
        if (info.tone) { tone = info.tone; tones.push(info.tone); }
      } else {
        base += c;
      }
    }
    return { base: base, tone: tone, tones: tones };
  }

  // ---------------------------------------------------------------
  // 2. Phụ âm đầu và vần hợp lệ
  // ---------------------------------------------------------------
  var ONSETS = ["ngh", "ng", "gh", "gi", "kh", "nh", "ph", "th", "tr", "ch",
    "b", "c", "d", "đ", "g", "h", "k", "l", "m", "n", "p", "r", "s", "t", "v", "x", ""];

  var RHYME_LIST = (
    "a ai ao au ay ac ach am an ang anh ap at " +
    "ăc ăm ăn ăng ăp ăt " +
    "âu ây âc âm ân âng âp ât " +
    "e eo ec em en eng ep et " +
    "ê êu êch êm ên ênh êp êt " +
    "i ia iu ich im in inh ip it " +
    "iêc iêm iên iêng iêp iêt iêu " +
    "yêm yên yêt yêu " +
    "o oi oc om on ong op ot ooc oong " +
    "ô ôi ôc ôm ôn ông ôp ôt " +
    "ơ ơi ơm ơn ơp ơt " +
    "u ua ui uc um un ung up ut " +
    "uôc uôi uôm uôn uông uôt " +
    "ư ưa ưi ưu ưc ưm ưng ưt " +
    "ươc ươi ươm ươn ương ươp ươt ươu " +
    "oa oac oach oai oam oan oang oanh oao oap oat oay " +
    "oăc oăm oăn oăng oăt " +
    "oe oen oeo oet " +
    "uân uâng uât uây " +
    "uê uêch uênh " +
    "uy uya uych uyên uyêt uyn uynh uyp uyt uyu " +
    "uơ y"
  ).split(" ");
  var RHYMES = {};
  RHYME_LIST.forEach(function (r) { if (r) RHYMES[r] = true; });

  // Tìm mọi cách tách (phụ âm đầu, vần) hợp lệ
  function splitSyllable(base) {
    var parses = [];
    if (base.indexOf("qu") === 0) {
      var rq = base.slice(1); // giữ "u" làm âm đệm: "quốc" -> q + "uôc"
      if (rq.length > 1 && RHYMES[rq]) parses.push({ onset: "qu", rhyme: rq });
      return parses;
    }
    if (base.charAt(0) === "q") return parses; // q luôn phải đi với u
    for (var i = 0; i < ONSETS.length; i++) {
      var on = ONSETS[i];
      if (base.indexOf(on) !== 0) continue;
      var rest = base.slice(on.length);
      if (!rest) continue;
      if (on === "gi") {
        if (RHYMES[rest]) parses.push({ onset: "gi", rhyme: rest });
        else if (RHYMES["i" + rest]) parses.push({ onset: "gi", rhyme: "i" + rest }); // "giết" = gi + iêt
        continue;
      }
      if (RHYMES[rest]) parses.push({ onset: on, rhyme: rest });
    }
    return parses;
  }

  // Luật chính tả cho một cách tách
  function ruleError(p, tone) {
    var r0 = p.rhyme.charAt(0);
    var front = "ieê".indexOf(r0) >= 0;
    if (p.onset === "c" && (front || r0 === "y")) return { code: "c_k", onset: "k" };
    if (p.onset === "k" && !(front || r0 === "y")) return { code: "k_c", onset: "c" };
    if (p.onset === "g" && (r0 === "e" || r0 === "ê")) return { code: "g_gh", onset: "gh" };
    if (p.onset === "gh" && !front) return { code: "gh_g", onset: "g" };
    if (p.onset === "ng" && front) return { code: "ng_ngh", onset: "ngh" };
    if (p.onset === "ngh" && !front) return { code: "ngh_ng", onset: "ng" };
    if (/(c|ch|p|t)$/.test(p.rhyme) && tone !== 2 && tone !== 5) return { code: "stop_tone" };
    return null;
  }

  // Kiểm tra một tiếng (đã viết thường)
  function validate(lower) {
    var st = stripTone(lower);
    if (st.tones.length > 1) return { valid: false, code: "two_tones", st: st };
    var parses = splitSyllable(st.base);
    if (!parses.length) return { valid: false, code: "structure", st: st };
    var first = null;
    for (var i = 0; i < parses.length; i++) {
      var e = ruleError(parses[i], st.tone);
      if (!e) return { valid: true, st: st };
      if (!first) { first = e; first.parse = parses[i]; }
    }
    return { valid: false, code: first.code, fixOnset: first.onset, parse: first.parse, st: st };
  }
  function isValidSyllable(lower) { return validate(lower).valid; }

  // ---------------------------------------------------------------
  // 3. Đặt dấu thanh vào đúng chỗ (kiểu cũ: hòa, thủy, của)
  // ---------------------------------------------------------------
  function placeTone(base, tone) {
    if (!tone) return base;
    var chars = base.split("");
    var start = -1, end = -1, i;
    for (i = 0; i < chars.length; i++) {
      if (isVowel(chars[i])) { if (start < 0) start = i; end = i; }
      else if (start >= 0) break;
    }
    if (start < 0) return base;
    if (start > 0 && end > start &&
        ((chars[start - 1] === "q" && chars[start] === "u") ||
         (chars[start - 1] === "g" && chars[start] === "i"))) start++;
    var pos = -1;
    for (i = start; i <= end; i++) if (MARKED_VOWELS.indexOf(chars[i]) >= 0) pos = i;
    if (pos < 0) {
      var len = end - start + 1, hasFinal = end < chars.length - 1;
      if (len === 1) pos = start;
      else if (hasFinal) pos = end;
      else if (len === 3) pos = start + 1;
      else pos = start;
    }
    chars[pos] = TONE_TABLE[chars[pos]].charAt(tone);
    return chars.join("");
  }

  // ---------------------------------------------------------------
  // 4. Đoán lỗi gõ Telex chưa chuyển dấu: "vieejt" -> "việt"
  // ---------------------------------------------------------------
  var TELEX_TONES = { s: 2, f: 1, r: 3, x: 4, j: 5 };
  function telexDecode(lower) {
    var st = stripTone(lower), s = st.base, tone = st.tone, key = "";
    var out = "", seenVowel = false;
    for (var i = 0; i < s.length; i++) {
      var c = s.charAt(i);
      if (seenVowel && TELEX_TONES.hasOwnProperty(c)) { tone = TELEX_TONES[c]; key = c; continue; }
      if (isVowel(c)) seenVowel = true;
      out += c;
    }
    var m = out.replace(/^dd/, "đ")
      .replace(/uwow/g, "ươ").replace(/uow/g, "ươ")
      .replace(/ow/g, "ơ").replace(/uw/g, "ư").replace(/aw/g, "ă")
      .replace(/aa/g, "â").replace(/ee/g, "ê").replace(/oo/g, "ô")
      .replace(/w/g, "ư");
    var modified = m !== out;
    if (/[fjwz]/.test(m)) return null;
    if (!modified && !key) return null;
    var result = placeTone(m, tone);
    if (result === lower || !isValidSyllable(result)) return null;
    return { result: result, key: key, modified: modified };
  }

  // ---------------------------------------------------------------
  // 5. Sinh gợi ý sửa
  // ---------------------------------------------------------------
  var FAMILY = { a: "aăâ", "ă": "aăâ", "â": "aăâ", e: "eê", "ê": "eê",
    o: "oôơ", "ô": "oôơ", "ơ": "oôơ", u: "uư", "ư": "uư", d: "dđ", "đ": "dđ", i: "iy", y: "iy" };

  function editCandidates(lower, st) {
    var base = st.base, tone = st.tone, out = [];
    function add(b, t) {
      var w = placeTone(b, t);
      if (w !== lower && out.indexOf(w) < 0 && isValidSyllable(w)) out.push(w);
    }
    var i, j;
    for (i = 0; i < base.length; i++) {           // đổi chữ cùng họ: a/ă/â, o/ô/ơ...
      var fam = FAMILY[base.charAt(i)];
      if (!fam) continue;
      for (j = 0; j < fam.length; j++) {
        if (fam.charAt(j) !== base.charAt(i)) add(base.slice(0, i) + fam.charAt(j) + base.slice(i + 1), tone);
      }
    }
    for (i = 0; i < base.length; i++) add(base.slice(0, i) + base.slice(i + 1), tone);  // thừa 1 chữ
    for (i = 0; i < base.length - 1; i++) {                                              // gõ ngược 2 chữ
      add(base.slice(0, i) + base.charAt(i + 1) + base.charAt(i) + base.slice(i + 2), tone);
    }
    return out;
  }

  function applyCase(orig, s) {
    if (!orig) return s;
    var up = orig.toUpperCase(), low = orig.toLowerCase();
    if (orig.length > 1 && orig === up && up !== low) return s.toUpperCase();
    var f = orig.charAt(0);
    if (f !== f.toLowerCase()) return s.charAt(0).toUpperCase() + s.slice(1);
    return s;
  }

  // Từ nước ngoài hay gặp — không báo lỗi
  var WHITELIST = {};
  ("ok oke okay email online offline internet website web video facebook zalo youtube tiktok google " +
   "app file pdf word excel wifi laptop game show idol fan marketing livestream link code admin fanpage " +
   "inbox comment like share post page covid virus vaccine test check team group chat shop sale ship " +
   "shipper teen how now low bow see too soon room seen door boom moon noon bee dawn bus max mix box tax " +
   "car bar per cv ceo kpi").split(" ").forEach(function (w) { if (w) WHITELIST[w] = true; });

  var RULE_TEXT = {
    c_k: "Trước i, e, ê, y phải viết \"k\" chứ không viết \"c\"",
    k_c: "\"k\" chỉ đứng trước i, e, ê, y — ở đây phải viết \"c\"",
    g_gh: "Trước e, ê phải viết \"gh\"",
    gh_g: "\"gh\" chỉ đứng trước i, e, ê — ở đây phải viết \"g\"",
    ng_ngh: "Trước i, e, ê phải viết \"ngh\"",
    ngh_ng: "\"ngh\" chỉ đứng trước i, e, ê — ở đây phải viết \"ng\"",
    stop_tone: "Tiếng kết thúc bằng c, ch, p, t chỉ đi với dấu sắc hoặc nặng",
    two_tones: "Một tiếng không thể có hai dấu thanh",
    structure: "Không ghép thành tiếng Việt hợp lệ — có thể gõ sai"
  };

  // Kiểm tra một từ. Trả về null nếu đúng (hoặc bỏ qua), ngược lại trả về lỗi.
  function checkWord(word, opts) {
    opts = opts || {};
    var w = nfc(word);
    var lower = w.toLowerCase();
    if (!lower) return null;
    if (WHITELIST[lower] || (opts.userDict && opts.userDict[lower])) return null;

    var v = validate(lower);
    if (v.valid) return null;
    if (!/[aăâeêioôơuưy]/.test(v.st.base)) return null; // ký hiệu, chữ cái đơn lẻ

    var hasVn = VN_SPECIAL.test(lower);
    var tel = telexDecode(lower);
    var strongTelex = !!tel && (hasVn || tel.modified || tel.key === "f" || tel.key === "j");

    if (!hasVn && !strongTelex) {
      if (w.length >= 2 && w === w.toUpperCase()) return null;   // chữ viết tắt: UBND, HCM
      if (/^[A-Z]/.test(w)) return null;                         // tên riêng nước ngoài
      if (/[fjwz]/.test(lower)) return null;                     // chữ cái không có trong tiếng Việt
      if (opts.skipForeign !== false) return null;               // từ không dấu, có thể là tiếng Anh
    }

    var res = { word: w, code: v.code, reason: "", suggestions: [], confident: false };
    if (strongTelex) {
      res.code = "telex";
      res.reason = "Có vẻ gõ Telex nhưng chưa chuyển thành chữ có dấu";
      res.suggestions.push(tel.result);
      res.confident = true;
    } else {
      res.reason = RULE_TEXT[v.code] || RULE_TEXT.structure;
      var st = v.st, i;
      if (v.code === "two_tones") {
        for (i = 0; i < st.tones.length; i++) res.suggestions.push(placeTone(st.base, st.tones[i]));
      } else if (v.code === "stop_tone") {
        res.suggestions.push(placeTone(st.base, 2), placeTone(st.base, 5));
        res.reason += st.tone ? " (đang là dấu " + TONE_NAMES[st.tone] + ")" : " (đang không có dấu)";
      } else if (v.fixOnset) {
        var b = v.fixOnset + st.base.slice(v.parse.onset.length);
        if (v.parse.onset === "gi") b = v.fixOnset + st.base.slice(2);
        var fixed = placeTone(b, st.tone);
        if (isValidSyllable(fixed)) { res.suggestions.push(fixed); res.confident = true; }
      }
      var edits = editCandidates(lower, st);
      for (i = 0; i < edits.length; i++) if (res.suggestions.indexOf(edits[i]) < 0) res.suggestions.push(edits[i]);
    }
    var clean = [];
    for (var k = 0; k < res.suggestions.length && clean.length < 5; k++) {
      var s = res.suggestions[k];
      if (s && s !== lower && clean.indexOf(s) < 0 && isValidSyllable(s)) clean.push(s);
    }
    res.suggestions = clean.map(function (s) { return applyCase(w, s); });
    if (!res.suggestions.length) res.confident = false;
    return res;
  }

  // ---------------------------------------------------------------
  // 6. Cụm từ hay viết sai: [sai, đúng, ghi chú]
  // ---------------------------------------------------------------
  var COMMON_ERRORS = [
    ["xuất xắc", "xuất sắc"], ["suất sắc", "xuất sắc"], ["chuẩn đoán", "chẩn đoán"],
    ["sáng lạng", "xán lạn"], ["xán lạng", "xán lạn"],
    ["tranh dành", "tranh giành"], ["dành giật", "giành giật"], ["giành dụm", "dành dụm"],
    ["dành chiến thắng", "giành chiến thắng"], ["giành riêng", "dành riêng"], ["giành thời gian", "dành thời gian"],
    ["cọ sát", "cọ xát"], ["sát nhập", "sáp nhập"], ["xát nhập", "sáp nhập"],
    ["chia sẽ", "chia sẻ"], ["san sẽ", "san sẻ"], ["rò rĩ", "rò rỉ"], ["bổ xung", "bổ sung"],
    ["sơ xuất", "sơ suất"], ["năng xuất", "năng suất"], ["lãng mạng", "lãng mạn"], ["dè bỉu", "dè bĩu"],
    ["chau chuốt", "trau chuốt"], ["trau giồi", "trau dồi"], ["chau dồi", "trau dồi"],
    ["đọc giả", "độc giả"], ["vô hình chung", "vô hình trung"], ["nổ lực", "nỗ lực"],
    ["bứt xúc", "bức xúc"], ["bức phá", "bứt phá"], ["dấu diếm", "giấu giếm"], ["giấu diếm", "giấu giếm"],
    ["sử lý", "xử lý"], ["sử lí", "xử lí"], ["xắp xếp", "sắp xếp"], ["sắp sếp", "sắp xếp"],
    ["chỉnh chu", "chỉn chu"], ["trân thành", "chân thành"], ["chân trọng", "trân trọng"],
    ["xem sét", "xem xét"], ["xắc xảo", "sắc sảo"], ["chính sác", "chính xác"],
    ["dải thưởng", "giải thưởng"], ["dõ ràng", "rõ ràng"], ["rỏ ràng", "rõ ràng"],
    ["sẳn sàng", "sẵn sàng"], ["rảnh rổi", "rảnh rỗi"], ["chẳng lẻ", "chẳng lẽ"], ["đẹp đẻ", "đẹp đẽ"],
    ["suy nghỉ", "suy nghĩ"], ["nghĩ ngơi", "nghỉ ngơi"], ["sữa chữa", "sửa chữa"], ["sửa chửa", "sửa chữa"],
    ["giữ dìn", "giữ gìn"], ["khuyếch trương", "khuếch trương"], ["khuyếch đại", "khuếch đại"],
    ["mải mải", "mãi mãi"], ["lổi lầm", "lỗi lầm"], ["hẵn hoi", "hẳn hoi"], ["dư giả", "dư dả"],
    ["giở dang", "dở dang"], ["kĩ lưởng", "kĩ lưỡng"], ["kỹ lưởng", "kỹ lưỡng"], ["giải bày", "giãi bày"],
    ["xuất sứ", "xuất xứ"], ["suất xứ", "xuất xứ"], ["sứ sở", "xứ sở"],
    ["chuyền đạt", "truyền đạt"], ["chuyền thống", "truyền thống"], ["chuyền hình", "truyền hình"],
    ["trinh phục", "chinh phục"], ["xinh viên", "sinh viên"], ["xinh nhật", "sinh nhật"],
    ["sinh đẹp", "xinh đẹp"], ["xử dụng", "sử dụng"], ["xác xuất", "xác suất"],
    ["chiến lượt", "chiến lược"], ["lượt bỏ", "lược bỏ"], ["tóm lượt", "tóm lược"],
    ["giữ liệu", "dữ liệu"], ["giữ dội", "dữ dội"], ["dàng buộc", "ràng buộc"], ["dành mạch", "rành mạch"],
    ["dức khoát", "dứt khoát"], ["bàng hoàn", "bàng hoàng"], ["ngoằn nghèo", "ngoằn ngoèo"],
    ["bánh trưng", "bánh chưng"], ["chưng bày", "trưng bày"], ["chả lời", "trả lời"],
    ["chách nhiệm", "trách nhiệm"], ["truẩn bị", "chuẩn bị"], ["dải quyết", "giải quyết"],
    ["thẫm quyền", "thẩm quyền"], ["thẫm phán", "thẩm phán"], ["dấy tờ", "giấy tờ"],
    ["thừa kết", "thừa kế"], ["di trúc", "di chúc"], ["sỡ hữu", "sở hữu"], ["tranh trấp", "tranh chấp"],
    ["trấp hành", "chấp hành"], ["sét xử", "xét xử"], ["xét sử", "xét xử"], ["nghỉa vụ", "nghĩa vụ"],
    ["khiếu lại", "khiếu nại"], ["bảng án", "bản án"], ["sử phạt", "xử phạt"], ["dám hộ", "giám hộ"],
    ["dao dịch", "giao dịch"], ["chình bày", "trình bày"]
  ];

  // ---------------------------------------------------------------
  // 7. Tách từ và kiểm tra cả đoạn văn
  // ---------------------------------------------------------------
  function tokenize(text) {
    var re = new RegExp(WORD_RE_SRC, "g"), m, tokens = [], last = 0;
    while ((m = re.exec(text)) !== null) {
      tokens.push({ word: m[0], lower: m[0].toLowerCase(), index: m.index, sep: text.slice(last, m.index) });
      last = m.index + m[0].length;
    }
    return tokens;
  }

  function checkText(text, opts) {
    opts = opts || {};
    text = nfc(text);
    var tokens = tokenize(text);
    var issues = [], byKey = {}, cache = {};
    var i, k;

    // Cụm từ hay sai
    COMMON_ERRORS.forEach(function (ce) {
      var parts = ce[0].split(" ");
      for (i = 0; i + parts.length <= tokens.length; i++) {
        var ok = true;
        for (k = 0; k < parts.length; k++) {
          var t = tokens[i + k];
          if (t.lower !== parts[k] || (k > 0 && !/^\s+$/.test(t.sep))) { ok = false; break; }
        }
        if (!ok) continue;
        var key = "p:" + ce[0];
        if (!byKey[key]) {
          var last = tokens[i + parts.length - 1];
          var found = text.slice(tokens[i].index, last.index + last.word.length);
          byKey[key] = { kind: "phrase", word: found, search: ce[0], pos: tokens[i].index,
            reason: ce[2] || "Cụm từ này hay bị viết sai", suggestions: [applyCase(found, ce[1])],
            replacement: ce[1], confident: true, count: 0 };
          issues.push(byKey[key]);
        }
        byKey[key].count++;
      }
    });

    // Từng tiếng
    for (i = 0; i < tokens.length; i++) {
      var tk = tokens[i], wkey = "w:" + tk.lower;
      if (byKey[wkey]) { byKey[wkey].count++; continue; }
      if (cache.hasOwnProperty(tk.word)) continue;
      var r = checkWord(tk.word, opts);
      if (!r) { cache[tk.word] = false; continue; }
      byKey[wkey] = { kind: "word", word: tk.word, search: tk.lower, pos: tk.index, reason: r.reason,
        suggestions: r.suggestions, confident: r.confident, code: r.code, count: 1 };
      issues.push(byKey[wkey]);
    }
    issues.sort(function (a, b) { return a.pos - b.pos; });
    return issues;
  }

  var api = {
    checkText: checkText, checkWord: checkWord, isValidSyllable: isValidSyllable,
    placeTone: placeTone, stripTone: stripTone, telexDecode: telexDecode,
    applyCase: applyCase, nfc: nfc, tokenize: tokenize,
    LETTER_CLASS: LETTER_CLASS, COMMON_ERRORS: COMMON_ERRORS
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.VnChecker = api;
})(this);
