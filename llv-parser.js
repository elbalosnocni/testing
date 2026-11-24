/* llv-parser.js — đọc lịch ca từ Excel của từng bộ phận (bố cục tự nhận dạng).
 * Chạy được trên trình duyệt (window.LLV) và Node (module.exports) để kiểm thử.
 * Đầu vào là ma trận ô (mảng 2 chiều) do SheetJS trả về với cellDates:false, raw:true:
 *   - ô ngày/giờ là SỐ (serial Excel), ô chữ là chuỗi. Ô "1/2/3" bị Excel đổi sang định dạng ngày vẫn là số 1/2/3.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.LLV = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  var EPOCH = Date.UTC(1899, 11, 30);

  function strip(s) {
    return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd').replace(/Đ/g, 'D').toLowerCase().replace(/\s+/g, ' ').trim();
  }
  function clean(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function daysIn(y, m) { return new Date(y, m, 0).getDate(); }

  /* ---------- Ký hiệu ca ---------- */
  function timeFromFraction(x) {
    var mins = Math.round((x - Math.floor(x)) * 1440);
    return Math.floor(mins / 60) + ':' + pad(mins % 60);
  }
  /** Chuẩn hóa ô ca thành "khóa" so sánh được (không phân biệt hoa/thường, khoảng trắng). */
  function normSym(v) {
    if (v == null) return '';
    if (typeof v === 'number') {
      if (!isFinite(v)) return '';
      if (Number.isInteger(v)) return String(v);
      if (v > 0 && v < 1) return timeFromFraction(v);
      return String(v);
    }
    if (typeof v === 'boolean') return '';
    var s = String(v).normalize('NFC').replace(/\s+/g, ' ').trim();
    if (!s) return '';
    s = s.replace(/\s*([-+*\/])\s*/g, '$1').replace(/^0(\d:)/, '$1').toUpperCase();
    return s.replace(/([A-Z])[\/.](?=[A-Z])/g, '$1');
  }
  /** Nhãn hiển thị gốc của ô (để người dùng nhận ra ký hiệu của mình). */
  function rawLabel(v) {
    if (v == null) return '';
    if (typeof v === 'number') return Number.isInteger(v) ? String(v) : (v > 0 && v < 1 ? timeFromFraction(v) : String(v));
    return clean(v);
  }

  /* ---------- Nhận dạng ngày ở dòng tiêu đề ---------- */
  function dayFromCell(v) {
    if (typeof v === 'number') {
      if (Number.isInteger(v) && v >= 1 && v <= 31) return v;
      if (v > 40000 && v < 80000) return new Date(EPOCH + Math.floor(v) * 864e5).getUTCDate();
      return null;
    }
    if (typeof v === 'string') {
      var s = v.trim(), m;
      if ((m = s.match(/^0?(\d{1,2})$/))) { var n = +m[1]; return n >= 1 && n <= 31 ? n : null; }
      if ((m = s.match(/^(\d{4})-(\d{2})-(\d{2})/))) return +m[3];
      if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-]\d{2,4})?$/))) return +m[2] <= 12 ? +m[1] : null;
    }
    return null;
  }

  function hasNamesLeft(rows, r, c0) {
    var n = 0;
    for (var i = r + 1; i < Math.min(rows.length, r + 10); i++) {
      var row = rows[i] || [];
      for (var c = 0; c < c0; c++) {
        var v = row[c];
        if (typeof v === 'string' && /[A-Za-zÀ-ỹ]{2}/.test(v) && v.trim().length >= 4) { n++; break; }
      }
    }
    return n >= 2;
  }
  /** Tìm dòng tiêu đề có dãy 1,2,3…≥28 cột liên tiếp và có tên nhân viên ở bên trái. */
  function detectGrid(rows, opt) {
    opt = opt || {};
    var minRun = opt.minRun || 28, want = opt.daysInMonth || 28, limit = Math.min(rows.length, opt.scanRows || 45), cands = [];
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [], d = [];
      for (var c = 0; c < row.length; c++) d[c] = dayFromCell(row[c]);
      for (var c0 = 0; c0 < row.length; c0++) {
        if (d[c0] !== 1) continue;
        var k = 1;
        while (c0 + k < row.length && d[c0 + k] === k + 1) k++;
        if (k >= minRun) { cands.push({ headerRow: r, firstCol: c0, len: k, ok: hasNamesLeft(rows, r, c0) }); break; }
      }
    }
    var pool = cands.filter(function (x) { return x.ok; });
    if (!pool.length) return null;
    var best = pool.filter(function (x) { return x.len >= want; })[0] || pool.sort(function (a, b) { return b.len - a.len; })[0];
    var dayCols = {};
    for (var i = 0; i < Math.min(best.len, 31); i++) dayCols[i + 1] = best.firstCol + i;
    return { headerRow: best.headerRow, firstCol: best.firstCol, dayCols: dayCols, nDays: Math.min(best.len, 31), daysInMonth: opt.daysInMonth || Math.min(best.len, 31) };
  }

  /* ---------- Nhận dạng cột thông tin ---------- */
  var RE = {
    name: /(^|\b)(ho ?(&|va|\/)? ?ten|full ?name|employee name|name|ten nv|nhan vien)(\b|$)/,
    stt: /^(stt|no|tt|so|#|sn|n0)\.?$/,
    id: /(ma ?nv|msnv|ma so nv|employee ?(id|code)|id ?code|^id$|^code$|ma nhan vien|ma so|^qc$|^ma$)/,
    clock: /(cham cong|clock|attendance|van tay|^mvt$|finger)/,
    pos: /(position|chuc vu|chuc danh|^job$|vi tri|^pos\.?$)/,
    group: /(bo phan|phong ban|^dept|^section|^to$|^team|^nhom)/
  };
  function colStats(rows, startRow, col, maxRows) {
    var n = 0, text = 0, words = 0, codes = 0, seq = 0, prev = null, num = 0;
    for (var r = startRow; r < Math.min(rows.length, startRow + (maxRows || 60)); r++) {
      var v = (rows[r] || [])[col]; if (v == null || v === '') continue; n++;
      if (typeof v === 'number') { num++; if (prev != null && v === prev + 1) seq++; prev = v; if (v < 1e7) codes++; }
      else {
        var s = clean(v);
        if (/^[A-Za-z]{0,4}\s?\d{1,7}$/.test(s)) codes++;
        else if (s.split(' ').length >= 2 && /[A-Za-zÀ-ỹ]{2}/.test(s) && s.length >= 5) { text++; words++; }
        else if (/[A-Za-zÀ-ỹ]{3}/.test(s)) text++;
        prev = null;
      }
    }
    return { n: n, text: text, words: words, codes: codes, seq: seq, num: num };
  }
  function headerLabels(rows, grid) {
    var out = [], from = Math.max(0, grid.headerRow - 4), to = grid.headerRow;
    for (var c = 0; c < grid.firstCol; c++) {
      var parts = [];
      for (var r = from; r <= to; r++) { var v = (rows[r] || [])[c]; if (typeof v === 'string' && clean(v)) parts.push(clean(v)); }
      out[c] = parts.join(' / ');
    }
    return out;
  }
  function guessCols(rows, grid) {
    var labels = headerLabels(rows, grid), start = grid.headerRow + 1, res = { name: -1, id: -1, clock: -1, pos: -1, group: -1 }, used = {};
    var L = labels.map(strip);
    function pick(key, re, skip) {
      for (var c = 0; c < L.length; c++) {
        if (used[c] || (skip && skip(c))) continue;
        var parts = L[c].split(' / ');
        for (var p = 0; p < parts.length; p++) if (re.test(parts[p])) { res[key] = c; used[c] = 1; return; }
      }
    }
    pick('name', RE.name);
    pick('clock', RE.clock);
    pick('id', RE.id);
    pick('pos', RE.pos);
    pick('group', RE.group, function (c) { return c === res.name; });
    if (res.name < 0) { // dự phòng: cột có nhiều chữ nhất
      var bestC = -1, bestW = 0;
      for (var c = 0; c < grid.firstCol; c++) { var st = colStats(rows, start, c); if (st.words > bestW) { bestW = st.words; bestC = c; } }
      res.name = bestC; if (bestC >= 0) used[bestC] = 1;
    }
    if (res.id < 0) { // dự phòng: cột toàn mã ngắn, không phải STT tăng dần
      for (var c2 = 0; c2 < grid.firstCol; c2++) {
        if (used[c2]) continue;
        var s2 = colStats(rows, start, c2);
        if (s2.n >= 3 && s2.codes / s2.n > 0.7 && s2.seq / s2.n < 0.5 && !RE.stt.test(L[c2] || '')) { res.id = c2; used[c2] = 1; break; }
      }
    }
    return { cols: res, labels: labels };
  }

  /* ---------- Chú thích ký hiệu dưới bảng ---------- */
  function timeTokens(s) {
    var out = [], re = /(\d{1,2})\s*(?:[:h]\s*(\d{2})?)?\s*(am|pm|sa|ch)?(?![\d])/gi, m;
    while ((m = re.exec(s))) {
      var h = +m[1], mi = m[2] ? +m[2] : 0, ap = (m[3] || '').toLowerCase();
      if (/[:h]/i.test(m[0]) === false && !ap) continue; // số trần không phải giờ
      if (ap === 'pm' || ap === 'ch') { if (h < 12) h += 12; } else if ((ap === 'am' || ap === 'sa') && h === 12) h = 0;
      if (h > 24 || mi > 59) continue;
      out.push(pad(h % 24 === 0 && h === 24 ? 0 : h) + ':' + pad(mi));
    }
    return out;
  }
  function hoursFromSymbolKey(key) {
    var m = String(key).match(/^(\d{1,2})(?::(\d{2}))?-(\d{1,2})(?::(\d{2}))?$/);
    if (!m) return '';
    return pad(+m[1]) + ':' + (m[2] || '00') + '-' + pad(+m[3]) + ':' + (m[4] || '00');
  }
  function parseLegend(rows, startRow) {
    var leg = {};
    for (var r = startRow; r < Math.min(rows.length, startRow + 400); r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        var v = row[c]; if (typeof v !== 'string') continue;
        var s = clean(v), m = s.match(/^([^:=]{1,10}?)\s*[:=]\s*(.+)$/);
        if (m) {
          var key = normSym(m[1].replace(/[:=\s]+$/, '')), tk = timeTokens(m[2].replace(/^[:=\s]+/, ''));
          if (/^\d{1,2}$/.test(key) && /^[0-5]\d(?!\s*h)\s*(am|pm|sa|ch|->|-|–|\s|$)/i.test(m[2])) continue; // "6 : 00 PM ..." là giờ, không phải ký hiệu
          if (!key || /^(NOTED|APPROVED|PREPARED|REMARKS|DATE|NOTE|GHI CHU)/.test(key)) continue;
          if (tk.length >= 2) leg[key] = { hours: tk[0] + '-' + tk[1], text: m[2] };
          else if (!tk.length && key.length <= 4 && m[2].length >= 3 && m[2].length <= 40 && /[A-Za-zÀ-ỹ]{3}/.test(m[2]) && !/^\d+$/.test(key)) leg[key] = leg[key] || { hours: '', text: m[2] };
        } else if (s.length > 8) {
          var tk2 = timeTokens(s), left = row[c - 1];
          if (tk2.length >= 2 && typeof left === 'string' && clean(left).length <= 6 && clean(left)) leg[normSym(left)] = { hours: tk2[0] + '-' + tk2[1], text: s };
        }
      }
    }
    return leg;
  }

  /* ---------- Trích nhân viên ---------- */
  var WEEKDAY = /^(MON|TUE|WED|THU|FRI|SAT|SUN|CN|T[2-7]|THU?\s?[2-7]|CHU NHAT)$/;
  var FOOTER = /^(name|date|ho ten|ho & ten|noted|approved|prepared|approval|expat|remarks|ghi chu|nguoi lap|giam doc|truong bo phan|tong|total)/;
  function extract(rows, grid, cols, opt) {
    opt = opt || {};
    var nd = grid.daysInMonth, emps = [], group = '', symbols = {}, start = grid.headerRow + 1, lastRow = start;
    function cell(r, c) { return c >= 0 ? (rows[r] || [])[c] : null; }
    for (var r = start; r < rows.length; r++) {
      var nm = cell(r, cols.name);
      if (nm == null || nm === '' || typeof nm === 'number') continue;
      nm = clean(nm); if (!nm) continue;
      var cells = {}, has = 0, seqHits = 0, wk = 0, tot = 0;
      for (var d0 = 1; d0 <= nd; d0++) {
        var c0x = grid.dayCols[d0]; if (c0x == null) continue;
        var v0 = (rows[r] || [])[c0x]; if (v0 == null || v0 === '') continue;
        tot++; if (dayFromCell(v0) === d0) seqHits++;
        if (WEEKDAY.test(strip(v0).toUpperCase())) wk++;
      }
      if (tot >= 15 && seqHits >= tot * 0.8 || tot >= 8 && wk >= tot * 0.6) continue; // dòng tiêu đề lặp / dòng thứ trong tuần
      for (var d = 1; d <= nd; d++) {
        var c = grid.dayCols[d]; if (c == null) continue;
        var v = (rows[r] || [])[c], k = normSym(v);
        if (!k) continue;
        if (typeof v === 'string' && (v.trim().length > 10 || (/[:=]/.test(v) && v.length > 5) || /^(noted|approved|prepared|approval)/i.test(v.trim()))) continue; // chú thích/chữ ký lẫn trong bảng
        has++; cells[d] = k;
        if (!symbols[k]) symbols[k] = { raw: rawLabel(v), count: 0 };
        symbols[k].count++;
      }
      if (!has) { if (!FOOTER.test(strip(nm))) group = nm; continue; }
      // dòng nhiều chữ chú thích bị lẫn vào (ví dụ "A : 7:00AM - 4:00PM") không phải nhân viên
      if (/[:=].*\d/.test(nm) && nm.length > 20) continue;
      var rowGroup = cols.group >= 0 ? clean(cell(r, cols.group)) : '';
      emps.push({
        row: r + 1, name: nm,
        id: cols.id >= 0 ? clean(cell(r, cols.id)) : '',
        clock: cols.clock >= 0 ? clean(cell(r, cols.clock)) : '',
        position: cols.pos >= 0 ? clean(cell(r, cols.pos)) : '',
        group: rowGroup || group, cells: cells
      });
      lastRow = r;
    }
    var legend = parseLegend(rows, grid.headerRow + 1);
    return { employees: emps, symbols: symbols, legend: legend };
  }

  /* ---------- Chọn sheet theo tháng ---------- */
  var MON = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  function monthTokens(name) {
    var s = strip(name), out = [], m;
    var re1 = /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\W*(20\d{2}|\d{2})?(?!\d)/g;
    while ((m = re1.exec(s))) out.push({ m: MON[m[1]], y: m[2] ? (m[2].length === 2 ? 2000 + +m[2] : +m[2]) : null });
    var re2 = /(?:^|[^\d])(1[0-2]|0?[1-9])\s*[-./_ ]?\s*(20\d{2}|\d{2})(?!\d)/g;
    while ((m = re2.exec(s))) out.push({ m: +m[1], y: m[2].length === 2 ? 2000 + +m[2] : +m[2] });
    var re3 = /(?:thang|th|t)\s*(1[0-2]|0?[1-9])(?!\d)(?:\s*[-./_ ]?\s*(20\d{2}|\d{2})(?!\d))?/g;
    while ((m = re3.exec(s))) out.push({ m: +m[1], y: m[2] ? (m[2].length === 2 ? 2000 + +m[2] : +m[2]) : null });
    var re4 = /(?:^|[^\d])(0?[1-9]|1[0-2])\s*$/; // tên sheet chỉ là "10"
    if (!out.length && (m = s.match(re4)) && /^\D*\d{1,2}$/.test(s)) out.push({ m: +m[1], y: null });
    return out;
  }
  /** score: 3 = đúng tháng+năm, 2 = đúng tháng (không ghi năm), 0 = không khớp */
  function sheetScore(name, year, month) {
    var best = 0, toks = monthTokens(name);
    for (var i = 0; i < toks.length; i++) {
      var t = toks[i];
      if (t.m !== month) continue;
      if (t.y === year) best = Math.max(best, 3);
      else if (t.y == null) best = Math.max(best, 2);
    }
    return best;
  }

  /* ---------- Gợi ý quy đổi ký hiệu → mã ca chung ---------- */
  var SYNONYMS = {
    OFF: ['OFF', 'DO', 'DAY OFF', 'NGHI', 'NGHỈ', 'REST'],
    AL: ['AL', 'VL', 'PHEP', 'PHÉP', 'NGHI PHEP', 'NGHỈ PHÉP', 'FL'],
    HC: ['HC', 'HÀNH CHÍNH', 'HANH CHINH'],
    LE: ['H', 'LE', 'LỄ', 'HOLIDAY', 'HOL'],
    TS: ['TS', 'NTS', 'THAI SAN', 'THAI SẢN'],
    WFH: ['WFH']
  };
  function suggestMapping(symbols, legend, catalog) {
    var byCode = {}, res = {};
    catalog.forEach(function (s) { if (s.active !== false) { byCode[normSym(s.code)] = s; byCode[normSym(s.name)] = s; } });
    function byHours(h, startOnly) {
      var p = h.split('-'), hit = null;
      catalog.forEach(function (s) {
        if (s.active === false || !s.start || !s.end || s.start === '00:00' && s.end === '00:00') return;
        var a = pad(+String(s.start).split(':')[0]) + ':' + String(s.start).split(':')[1], b = pad(+String(s.end).split(':')[0]) + ':' + String(s.end).split(':')[1];
        if (a === p[0] && (startOnly || b === p[1])) hit = hit || s;
      });
      return hit;
    }
    Object.keys(symbols).forEach(function (k) {
      var hit = byCode[k];
      if (hit) { res[k] = { code: hit.code, why: 'trùng mã/tên ca chung', conf: 'high' }; return; }
      for (var code in SYNONYMS) {
        if (SYNONYMS[code].some(function (x) { return normSym(x) === k; }) && catalog.some(function (s) { return s.code === code && s.active !== false; })) {
          res[k] = { code: code, why: 'ký hiệu thông dụng', conf: 'med' }; return;
        }
      }
      var lg = legend[k];
      if (lg && lg.hours) {
        var h1 = byHours(lg.hours);
        if (h1) { res[k] = { code: h1.code, why: 'chú thích trong file: ' + lg.hours, conf: 'med' }; return; }
        res[k] = { code: '__NEW__', newShift: lg.hours, why: 'chú thích trong file: ' + lg.hours + ' (chưa có ca chung này)', conf: 'med' }; return;
      }
      var hs = hoursFromSymbolKey(k);
      if (hs) {
        var h2 = byHours(hs);
        res[k] = h2 ? { code: h2.code, why: 'ký hiệu là giờ ' + hs, conf: 'med' } : { code: '__NEW__', newShift: hs, why: 'ký hiệu là giờ ' + hs + ' (chưa có ca chung này)', conf: 'med' };
        return;
      }
      var tm = k.match(/^(\d{1,2}):(\d{2})$/);
      if (tm) { var h3 = byHours(pad(+tm[1]) + ':' + tm[2] + '-', true); if (h3) { res[k] = { code: h3.code, why: 'giờ vào ' + k + ' (đoán, cần kiểm tra)', conf: 'low' }; return; } }
      var dg = k.match(/^([123])$/);
      if (dg && byCode['C' + dg[1]]) { res[k] = { code: 'C' + dg[1], why: 'số ca (đoán, cần kiểm tra)', conf: 'low' }; return; }
    });
    return res;
  }

  /** Dạng tên so khớp: bỏ dấu, chữ thường, bỏ khoảng trắng thừa. */
  function nameKey(s) { return strip(s).replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim(); }
  /** Cảnh báo tên dùng font VNI/TCVN cũ (hiện ký tự lạ như "NGUYEÃN THÒ"). */
  function looksLegacyFont(names) {
    var bad = 0;
    names.forEach(function (n) { if (/[AEIOUY][ÃÕÙÛÖØÏÑÒ]/.test(n) && n === n.toUpperCase()) bad++; });
    return names.length > 3 && bad / names.length > 0.2;
  }

  /** Phân tích 1 sheet: trả về null nếu không giống bảng lịch. */
  function parseSheet(rows, year, month, override) {
    var grid = detectGrid(rows, { daysInMonth: daysIn(year, month) });
    if (!grid) return null;
    grid.daysInMonth = Math.min(daysIn(year, month), grid.nDays);
    var g = guessCols(rows, grid), cols = Object.assign({}, g.cols, override || {});
    if (cols.name < 0) return null;
    var ex = extract(rows, grid, cols);
    return { grid: grid, cols: cols, labels: g.labels, employees: ex.employees, symbols: ex.symbols, legend: ex.legend };
  }

  return {
    normSym: normSym, rawLabel: rawLabel, strip: strip, nameKey: nameKey, dayFromCell: dayFromCell,
    detectGrid: detectGrid, guessCols: guessCols, extract: extract, parseLegend: parseLegend, timeTokens: timeTokens,
    sheetScore: sheetScore, monthTokens: monthTokens, suggestMapping: suggestMapping, parseSheet: parseSheet,
    looksLegacyFont: looksLegacyFont, hoursFromSymbolKey: hoursFromSymbolKey, daysIn: daysIn
  };
});
