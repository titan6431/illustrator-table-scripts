/*
  Parametric Bar Chart v1.0 (Bar-only, ES3-safe)
  Illustrator ExtendScript

  반영 사항(핵심)
  - 값/라벨 텍스트: K90 (진하게)
  - 축선(X) + 그리드(가로선): K80 (둘 다 통일)
  - AutoScale + Headroom tick 옵션: 최대값이 꼭대기면 +1 step
  - Values Only 업데이트 시 값 라벨이 우측으로 밀리는 버그 수정(센터 정렬 + bounds 재센터링)
  - 그룹 간격 = 막대 간격 * 2.5 (하드 룰)
  - 데이터 많아도 plot 폭 밖으로 절대 안 나가게 강제 fit
  - 데이터 적으면 barWidth 자동 확대 허용(상한 2.0배)
  - 양끝 여백은 “조금” 남기되 과하게 넓어지지 않게(outer cap)
  - Y 라벨 위치: left / right / both / none
  - Y 축선 위치: left / right / both / none
  - 범례 스와치 크기 3.5mm
  - UI 정리: 폰트 UI 삭제, thousand separators UI 삭제(항상 ON), dashed 삭제, grid opacity UI 삭제, empty cell UI 삭제(항상 0 처리)
  - Legend 탭: Text lift 기본값 2.7mm로 변경
  - UI 창: 리사이즈 가능 + 세로 스크롤바 제공(모니터 작아도 하단 안 잘리게)
*/

#target illustrator

(function () {

  // -----------------------------
  // Polyfills (old engine)
  // -----------------------------
  if (!Array.prototype.map) {
    Array.prototype.map = function (cb) {
      var out = [];
      for (var i = 0; i < this.length; i++) out.push(cb(this[i], i, this));
      return out;
    };
  }
  if (!Array.prototype.filter) {
    Array.prototype.filter = function (cb) {
      var out = [];
      for (var i = 0; i < this.length; i++) if (cb(this[i], i, this)) out.push(this[i]);
      return out;
    };
  }
  if (typeof String.prototype.trim !== "function") {
    String.prototype.trim = function () { return this.replace(/^\s+|\s+$/g, ""); };
  }

  if (app.documents.length === 0) {
    alert("열려 있는 문서가 없습니다.");
    return;
  }

  var doc = app.activeDocument;

  // -----------------------------
  // Units / Helpers
  // -----------------------------
  var MM_TO_PT = 2.8346456693;
  function mm(v) { return Number(v) * MM_TO_PT; }
  function clamp(n, a, b) { return Math.max(a, Math.min(b, n)); }

  function getActiveArtboardRect(_doc) {
    var idx = _doc.artboards.getActiveArtboardIndex();
    return _doc.artboards[idx].artboardRect; // [left, top, right, bottom]
  }

  function clearGroup(g) {
    while (g.pageItems.length > 0) {
      try { g.pageItems[0].remove(); } catch (e) { break; }
    }
  }

  function isChartGroup(item) {
    if (!item || item.typename !== "GroupItem") return false;
    if (!item.note) return false;
    return item.note.indexOf('"type":"ParamBarChart"') !== -1;
  }

  function safeJSONParse(s) {
    try { if (typeof JSON === "object" && JSON.parse) return JSON.parse(s); } catch (e1) {}
    try { return eval("(" + s + ")"); } catch (e2) { return null; }
  }

  function safeJSONStringify(obj) {
    try { if (typeof JSON === "object" && JSON.stringify) return JSON.stringify(obj); } catch (e0) {}
    var parts = [];
    for (var k in obj) {
      if (!obj.hasOwnProperty(k)) continue;
      var v = obj[k];
      if (typeof v === "string") {
        parts.push('"' + k + '":"' + v.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\n/g, "\\n").replace(/\t/g, "\\t") + '"');
      } else if (typeof v === "number" || typeof v === "boolean") {
        parts.push('"' + k + '":' + String(v));
      }
    }
    return "{" + parts.join(",") + "}";
  }

  function findChartGroupFromSelection(item) {
    var cur = item, guard = 0;
    while (cur && guard < 80) {
      if (cur.typename === "GroupItem" && isChartGroup(cur)) return cur;
      cur = cur.parent; guard++;
    }
    return null;
  }

  function findChildGroupByName(root, name) {
    if (!root || !root.groupItems) return null;
    for (var i = 0; i < root.groupItems.length; i++) {
      if (root.groupItems[i].name === name) return root.groupItems[i];
    }
    return null;
  }

  // -----------------------------
  // Fonts (UI 삭제, 내부 name 우선)
  // -----------------------------
  function getFontByNameTry(name) {
    try { return app.textFonts.getByName(name); } catch (e) { return null; }
  }

  function resolveFontSmart(preferredName, fallbackName, preferredFamily, preferredStyle) {
    // 1) exact internal name
    var f = getFontByNameTry(preferredName);
    if (f) return f;

    // 2) family+style match
    if (preferredFamily && preferredStyle) {
      try {
        for (var i = 0; i < app.textFonts.length; i++) {
          var tf = app.textFonts[i];
          if (tf.family === preferredFamily && tf.style === preferredStyle) return tf;
        }
      } catch (e2) {}
    }

    // 3) fallback name
    var ff = getFontByNameTry(fallbackName);
    if (ff) return ff;

    // 4) last: return null (AI default)
    return null;
  }

  // -----------------------------
  // Number formatting
  // -----------------------------
  function addThousandSep(str) {
    var neg = false;
    if (str.charAt(0) === "-") { neg = true; str = str.substring(1); }
    var parts = str.split(".");
    var i = parts[0];
    var out = "";
    while (i.length > 3) {
      out = "," + i.substring(i.length - 3) + out;
      i = i.substring(0, i.length - 3);
    }
    out = i + out;
    if (parts.length > 1) out += "." + parts[1];
    return (neg ? "-" : "") + out;
  }

  function formatValue(v, decimals, suffix) {
    var s = Number(v).toFixed(decimals);
    s = addThousandSep(s); // 항상 ON
    if (suffix && String(suffix).length > 0) s += String(suffix);
    return s;
  }

  // -----------------------------
  // Color (K 기준: 0=White, 100=Black 느낌으로 사용)
  // -----------------------------
  function makeK(k) {
    var c = new GrayColor();
    c.gray = clamp(Number(k), 0, 100);
    return c;
  }

  function axisStrokeDefault() { return makeK(80); }   // K80
  function gridStrokeDefault() { return makeK(80); }   // K80
  function valueTextDefault() { return makeK(90); }    // K90
  function labelTextDefault() { return makeK(90); }    // K90

  // -----------------------------
  // Data parsing
  // -----------------------------
  function detectDelimiter(line) {
    if (line.indexOf("\t") !== -1) return "\t";
    if (line.indexOf(",") !== -1) return ",";
    return "\t";
  }

  function parseDataSingle(raw) {
    raw = (raw || "").replace(/\r/g, "").trim();
    if (!raw) return null;
    var lines = raw.split("\n").filter(function (l) { return l.trim().length > 0; });
    if (lines.length < 1) return null;
    var delim = detectDelimiter(lines[0]);

    var categories = [];
    var seriesNames = ["Value"];
    var values = [];

    for (var i = 0; i < lines.length; i++) {
      var parts = lines[i].split(delim);
      if (parts.length === 0) continue;

      var cat = (parts[0] || "").toString().trim();
      if (!cat) cat = "Cat" + (i + 1);

      var vStr = "";
      if (parts.length >= 2) vStr = (parts[1] || "").toString().trim();
      else vStr = (parts[0] || "").toString().trim();

      var num = Number(vStr);
      if (parts.length === 1) {
        if (isNaN(num)) continue;
        cat = String(i + 1);
      }

      categories.push(cat);

      var vv = 0; // empty cell => 0 (항상)
      if (vStr === "") vv = 0;
      else vv = (isNaN(num) ? 0 : num);

      values.push([vv]);
    }

    if (categories.length === 0) return null;

    var minV = 0, maxV = 0, hasAny = false;
    for (var r = 0; r < values.length; r++) {
      var v = values[r][0];
      hasAny = true;
      if (v > maxV) maxV = v;
      if (v < minV) minV = v;
    }
    if (!hasAny) { minV = 0; maxV = 1; }

    return {
      delim: delim,
      categories: categories,
      seriesNames: seriesNames,
      values: values,
      dataMin: minV,
      dataMax: maxV
    };
  }

  // 그룹 모드: 첫 줄 전체(첫 칸 포함)=시리즈명, 2번째 줄부터 첫 칸=카테고리
  function parseDataGroup(raw) {
    raw = (raw || "").replace(/\r/g, "").trim();
    if (!raw) return null;

    var lines = raw.split("\n").filter(function (l) { return l.trim().length > 0; });
    if (lines.length < 2) return null;

    var delim = detectDelimiter(lines[0]);

    var headParts = lines[0].split(delim).map(function (s) { return s.trim(); });
    var seriesNames = [];
    for (var h = 0; h < headParts.length; h++) {
      var nm = headParts[h];
      if (!nm) nm = "S" + (h + 1);
      seriesNames.push(nm);
    }
    if (seriesNames.length < 1) return null;

    var categories = [];
    var values = [];

    for (var i = 1; i < lines.length; i++) {
      var parts = lines[i].split(delim);
      if (parts.length < 1) continue;

      var cat = (parts[0] || "").toString().trim();
      if (!cat) cat = "Cat" + i;
      categories.push(cat);

      var row = [];
      for (var s = 0; s < seriesNames.length; s++) {
        var idx = s + 1;
        var vStr = (idx < parts.length) ? (parts[idx] || "").toString().trim() : "";
        if (vStr === "") { row.push(0); continue; }
        var num = Number(vStr);
        row.push(isNaN(num) ? 0 : num);
      }
      values.push(row);
    }

    if (categories.length === 0) return null;

    var minV = 0, maxV = 0, hasAny = false;
    for (var r = 0; r < values.length; r++) {
      for (var s2 = 0; s2 < seriesNames.length; s2++) {
        var vv = values[r][s2];
        hasAny = true;
        if (vv > maxV) maxV = vv;
        if (vv < minV) minV = vv;
      }
    }
    if (!hasAny) { minV = 0; maxV = 1; }

    return {
      delim: delim,
      categories: categories,
      seriesNames: seriesNames,
      values: values,
      dataMin: minV,
      dataMax: maxV
    };
  }

  function parseData(raw, groupMode) {
    return groupMode ? parseDataGroup(raw) : parseDataSingle(raw);
  }

  // -----------------------------
  // Autoscale (nice numbers)
  // -----------------------------
  function niceNumber(range, round) {
    var exponent = Math.floor(Math.log(range) / Math.LN10);
    var fraction = range / Math.pow(10, exponent);
    var niceFraction;

    if (round) {
      if (fraction < 1.5) niceFraction = 1;
      else if (fraction < 3) niceFraction = 2;
      else if (fraction < 7) niceFraction = 5;
      else niceFraction = 10;
    } else {
      if (fraction <= 1) niceFraction = 1;
      else if (fraction <= 2) niceFraction = 2;
      else if (fraction <= 5) niceFraction = 5;
      else niceFraction = 10;
    }
    return niceFraction * Math.pow(10, exponent);
  }

  function autoScale(minV, maxV, tickCount) {
    if (minV === maxV) maxV = minV + 1;
    var minA = (minV >= 0) ? 0 : minV;
    var range = niceNumber(maxV - minA, false);
    var step = niceNumber(range / (tickCount || 5), true);
    var niceMin = Math.floor(minA / step) * step;
    var niceMax = Math.ceil(maxV / step) * step;
    if (niceMax === niceMin) niceMax = niceMin + step;
    return { min: niceMin, max: niceMax, step: step };
  }

  // -----------------------------
  // Drawing primitives
  // -----------------------------
  function addPointText(group, x, y, text, sizePt, fillColor, justification, fontObj) {
    var tf = group.textFrames.add();
    tf.kind = TextType.POINTTEXT;
    tf.contents = String(text);
    tf.position = [x, y];
    try {
      var tr = tf.textRange;
      tr.characterAttributes.size = Number(sizePt);
      if (fillColor) tr.characterAttributes.fillColor = fillColor;
      if (fontObj) tr.characterAttributes.textFont = fontObj;
      if (justification) tr.paragraphAttributes.justification = justification;
    } catch (e) {}
    return tf;
  }

  function forceCenterJustification(tf) {
    try {
      tf.kind = TextType.POINTTEXT;
      tf.textRange.paragraphAttributes.justification = Justification.CENTER;
    } catch (e) {}
  }

  // position/justification 변경 후 X가 미세하게 흔들리는 문제를 bounds 기준으로 강제 해결
  function centerTextFrameAtX(tf, targetX) {
    try {
      var b = tf.visibleBounds; // [left, top, right, bottom]
      var cx = (b[0] + b[2]) / 2;
      tf.translate(targetX - cx, 0);
      return;
    } catch (e1) {}
    try {
      var b2 = tf.geometricBounds;
      var cx2 = (b2[0] + b2[2]) / 2;
      tf.translate(targetX - cx2, 0);
    } catch (e2) {}
  }

  function drawLine(group, x1, y1, x2, y2, strokeColor, strokePt, opacity) {
    var p = group.pathItems.add();
    p.setEntirePath([[x1, y1], [x2, y2]]);
    p.filled = false;
    p.stroked = true;
    p.strokeWidth = strokePt;
    if (strokeColor) p.strokeColor = strokeColor;
    try { if (opacity !== undefined) p.opacity = clamp(Number(opacity), 0, 100); } catch (e) {}
    return p;
  }

  // Top-only rounded rectangle as single path
  function addTopRoundedBar(group, topY, leftX, w, h, r, fillColor) {
    r = Math.max(0, Math.min(r, Math.min(w / 2, h)));
    if (r <= 0.01) {
      var rect = group.pathItems.rectangle(topY, leftX, w, h);
      rect.stroked = false; rect.filled = true; rect.fillColor = fillColor;
      return rect;
    }

    var k = 0.5522847498;
    var c = r * k;

    var x0 = leftX;
    var x1 = leftX + w;
    var yT = topY;
    var yB = topY - h;

    var ptBL = [x0, yB];
    var ptBR = [x1, yB];
    var ptR  = [x1, yT - r];
    var ptTR = [x1 - r, yT];
    var ptTL = [x0 + r, yT];
    var ptL  = [x0, yT - r];

    var p = group.pathItems.add();
    p.stroked = false;
    p.filled = true;
    p.fillColor = fillColor;
    p.closed = true;
    p.setEntirePath([ptBL, ptBR, ptR, ptTR, ptTL, ptL]);

    var pp = p.pathPoints;
    for (var i = 0; i < pp.length; i++) {
      pp[i].pointType = PointType.CORNER;
      pp[i].leftDirection = pp[i].anchor;
      pp[i].rightDirection = pp[i].anchor;
    }

    pp[2].pointType = PointType.SMOOTH;
    pp[2].rightDirection = [x1, (yT - r) + c];

    pp[3].pointType = PointType.SMOOTH;
    pp[3].leftDirection = [(x1 - r) + c, yT];

    pp[4].pointType = PointType.SMOOTH;
    pp[4].rightDirection = [(x0 + r) - c, yT];

    pp[5].pointType = PointType.SMOOTH;
    pp[5].leftDirection = [x0, (yT - r) + c];

    return p;
  }

  function mapY(val, minV, maxV, plotBottomY, plotH) {
    if (maxV === minV) return plotBottomY;
    var t = (val - minV) / (maxV - minV);
    t = clamp(t, 0, 1);
    return plotBottomY + (t * plotH);
  }

  // -----------------------------
  // Geometry: fit + expand
  // groupGap = barGap * 2.5 (hard)
  // -----------------------------
  function computeGroupW(seriesCount, bw, bg) {
    return seriesCount * bw + (seriesCount - 1) * bg;
  }

  function totalBarsWidth(catCount, seriesCount, bw, bg, gg) {
    var groupW = computeGroupW(seriesCount, bw, bg);
    return (catCount * groupW) + (catCount - 1) * gg;
  }

  function resolveBarGeometryAdaptive(plotW, catCount, seriesCount, bw, bg, gg, allowExpand) {
    var minBW = mm(0.8);
    var minBG = mm(0.2);

    function enforceRatio(_bg, _gg) {
      var minGG = _bg * 2.5;
      return Math.max(_gg, minGG);
    }

    gg = enforceRatio(bg, gg);

    var tw = totalBarsWidth(catCount, seriesCount, bw, bg, gg);
    if (tw > plotW) {
      var guard = 0;
      while (tw > plotW && guard < 2000) {
        var factor = plotW / tw;
        factor = clamp(factor, 0.7, 0.97);

        bg = Math.max(minBG, bg * factor);
        gg = enforceRatio(bg, gg * factor);

        bw = Math.max(minBW, bw * factor);

        tw = totalBarsWidth(catCount, seriesCount, bw, bg, gg);
        guard++;
        if (bw === minBW && bg === minBG) break;
      }

      if (tw > plotW) {
        var factor2 = plotW / tw;
        bw = Math.max(minBW, bw * factor2);
        bg = Math.max(minBG, bg * factor2);
        gg = enforceRatio(bg, gg * factor2);
      }

      return { bw: bw, bg: bg, gg: gg, mode: "shrink" };
    }

    if (allowExpand) {
      var baseTW = tw;
      var targetFill = plotW * 0.82;
      var maxFactor = 2.0;
      if (baseTW < targetFill && catCount > 0) {
        var desiredFactor = targetFill / baseTW;
        desiredFactor = Math.min(desiredFactor, maxFactor);
        if (desiredFactor > 1.0) {
          bw = bw * desiredFactor;
          gg = enforceRatio(bg, gg);
        }
      }
    }

    return { bw: bw, bg: bg, gg: gg, mode: "as-is/expand" };
  }

  // -----------------------------
  // Layout: outer margin + gap caps
  // -----------------------------
  function computeBarLayoutNice(plotLeft, plotW, catCount, seriesCount, bw, bg, gg, autoDistribute) {
    var xPos = [];
    var groupW = computeGroupW(seriesCount, bw, bg);

    var gapMin = Math.max(gg, bg * 2.5);

    var outerMin = mm(2);
    var outerMax = Math.min(mm(12), plotW * 0.12);

    var gapMax = Math.min(groupW * 1.2, mm(30));
    if (gapMax < gapMin) gapMax = gapMin;

    var barsOnly = catCount * groupW;
    var gapsOnly = (catCount > 1) ? (catCount - 1) * gapMin : 0;

    var remaining = plotW - (barsOnly + gapsOnly);

    var gapUsed = gapMin;
    var outer = outerMin;
    var startX = plotLeft;

    if (!autoDistribute) {
      var totalW = barsOnly + gapsOnly;
      startX = plotLeft + (plotW - totalW) / 2;
      outer = (plotW - totalW) / 2;
    } else {
      if (remaining < 0) {
        startX = plotLeft + outerMin;
        gapUsed = gapMin;
        outer = outerMin;
      } else {
        var outerGrowCap = outerMax - outerMin;
        var useOuter = Math.min(remaining, outerGrowCap * 2);
        outer = outerMin + (useOuter / 2);
        remaining -= useOuter;

        if (catCount > 1) {
          var gapGrowCap = gapMax - gapUsed;
          var useGap = Math.min(remaining, gapGrowCap * (catCount - 1));
          gapUsed = gapUsed + (useGap / (catCount - 1));
          remaining -= useGap;
        }

        startX = plotLeft + outer + (remaining / 2);
      }
    }

    for (var c = 0; c < catCount; c++) {
      var arr = [];
      var base = startX + c * (groupW + gapUsed);
      for (var s = 0; s < seriesCount; s++) arr.push(base + s * (bw + bg));
      xPos.push(arr);
    }

    return { xPos: xPos, groupW: groupW, startX: startX, gapUsed: gapUsed, outer: outer };
  }

  // -----------------------------
  // Legend (below, centered, wrap)
  // -----------------------------
  function drawLegend(gLegend, plotLeft, plotW, baseY, seriesNames, seriesFills, cfg, fontObj) {
    if (!cfg.legendOn) return;
    if (!seriesNames || seriesNames.length <= 1) return;

    var fontPt = Number(cfg.legendFontPt);
    if (isNaN(fontPt) || fontPt <= 0) fontPt = 8;

    var lift = mm(Number(cfg.legendTextLiftMm));
    if (isNaN(lift)) lift = mm(2.7); // 기본 2.7mm

    var sw = mm(3.5);
    var gapX = mm(6);
    var gapSwText = mm(2.5);
    var lineGap = mm(6);

    function approxTextW(txt) {
      return (fontPt * 0.55) * String(txt).length;
    }

    var items = [];
    for (var i = 0; i < seriesNames.length; i++) {
      var t = seriesNames[i];
      var w = sw + gapSwText + approxTextW(t);
      items.push({ text: t, w: w, fill: seriesFills[i] });
    }

    var lines = [];
    var cur = [];
    var curW = 0;
    for (var k = 0; k < items.length; k++) {
      var addW = items[k].w + (cur.length > 0 ? gapX : 0);
      if (cur.length > 0 && (curW + addW) > plotW) {
        lines.push({ items: cur, w: curW });
        cur = [];
        curW = 0;
      }
      cur.push(items[k]);
      curW += addW;
    }
    if (cur.length > 0) lines.push({ items: cur, w: curW });

    for (var li = 0; li < lines.length; li++) {
      var line = lines[li];
      var x = plotLeft + (plotW - line.w) / 2;
      var y = baseY - (li * lineGap);

      for (var ii = 0; ii < line.items.length; ii++) {
        var it = line.items[ii];

        var rectTop = y + (sw / 2);
        var rectLeft = x;
        var r = gLegend.pathItems.rectangle(rectTop, rectLeft, sw, sw);
        r.stroked = false;
        r.filled = true;
        r.fillColor = it.fill;

        var tx = x + sw + gapSwText;
        addPointText(gLegend, tx, y + lift, it.text, fontPt, labelTextDefault(), Justification.LEFT, fontObj);

        x += it.w + gapX;
      }
    }
  }

  // -----------------------------
  // Main draw (full rebuild)
  // -----------------------------
  function drawChartFull(_doc, cfg, targetGroup) {
    var parsed = parseData(cfg.rawData, cfg.groupMode);
    if (!parsed) {
      alert("데이터를 파싱할 수 없습니다.\n(단일: [카테고리\\t값] / 그룹: 첫 줄=시리즈명, 2줄부터 첫 칸=카테고리)");
      return null;
    }

    var ab = getActiveArtboardRect(_doc);
    var left = ab[0], top = ab[1], right = ab[2], bottom = ab[3];
    var abW = right - left, abH = top - bottom;

    var chartW = mm(Number(cfg.chartWidthMm));
    var chartH = mm(Number(cfg.chartHeightMm));
    var originX = left + (abW - chartW) / 2;
    var originY = top - (abH - chartH) / 2;

    var padL = mm(Number(cfg.padLeftMm));
    var padR = mm(Number(cfg.padRightMm));
    var padT = mm(Number(cfg.padTopMm));
    var padB = mm(Number(cfg.padBottomMm));

    var plotLeft = originX + padL;
    var plotTopY = originY - padT;
    var plotW = chartW - padL - padR;
    var plotH = chartH - padT - padB;
    var plotBottomY = plotTopY - plotH;

    var axisMin = Number(cfg.axisMin);
    var axisMax = Number(cfg.axisMax);
    var axisStep = Number(cfg.axisStep);

    if (cfg.axisAuto) {
      var s = autoScale(parsed.dataMin, parsed.dataMax, 5);
      axisMin = s.min; axisMax = s.max; axisStep = s.step;

      if (cfg.headroomTick) {
        if (parsed.dataMax >= axisMax - 1e-9) axisMax += axisStep;
      }
    }

    var g = targetGroup || _doc.groupItems.add();
    g.name = "ParamBarChart";
    if (targetGroup) clearGroup(g);

    cfg._resolvedAxisMin = axisMin;
    cfg._resolvedAxisMax = axisMax;
    cfg._resolvedAxisStep = axisStep;
    cfg._categoryCount = parsed.categories.length;
    cfg._seriesCount = parsed.seriesNames.length;
    g.note = safeJSONStringify(cfg);

    function sub(name) {
      var sg = g.groupItems.add();
      sg.name = name;
      return sg;
    }

    var gGrid = sub("Grid");
    var gSeries = sub("Series");
    var gAxes = sub("Axes");
    var gValues = sub("ValueLabels");
    var gCats = sub("CategoryLabels");
    var gAxisLabelsL = sub("AxisLabelsLeft");
    var gAxisLabelsR = sub("AxisLabelsRight");
    var gLegend = sub("Legend");

    var fontObj = resolveFontSmart(
      cfg.preferredFontPSName,
      cfg.fallbackFontPSName,
      cfg.preferredFontFamily,
      cfg.preferredFontStyle
    );

    var axisCol = axisStrokeDefault();
    var gridCol = gridStrokeDefault();
    var textCol = labelTextDefault();
    var valueCol = valueTextDefault();

    var gridStrokePt = Number(cfg.gridThicknessPt);
    var axisStrokePt = Number(cfg.axisThicknessPt);
    if (isNaN(gridStrokePt) || gridStrokePt <= 0) gridStrokePt = 0.3;
    if (isNaN(axisStrokePt) || axisStrokePt <= 0) axisStrokePt = 0.75;

    if (cfg.gridShow) {
      var v = axisMin;
      var guard = 0;
      while (v <= axisMax + 1e-9 && guard < 500) {
        var yy = mapY(v, axisMin, axisMax, plotBottomY, plotH);
        drawLine(gGrid, plotLeft, yy, plotLeft + plotW, yy, gridCol, gridStrokePt, cfg.gridOpacityFixed);
        v += axisStep;
        guard++;
      }
    }

    var bw = mm(Number(cfg.barWidthMm));
    var bg = mm(Number(cfg.barGapMm));
    if (isNaN(bg) || bg <= 0) bg = mm(4);

    var gg = bg * 2.5;

    var rad = mm(Number(cfg.barRadiusMm));
    var geo = resolveBarGeometryAdaptive(plotW, parsed.categories.length, parsed.seriesNames.length, bw, bg, gg, !!cfg.allowBarWidthExpand);
    bw = geo.bw; bg = geo.bg; gg = geo.gg;

    var layout = computeBarLayoutNice(plotLeft, plotW, parsed.categories.length, parsed.seriesNames.length, bw, bg, gg, !!cfg.autoDistributeGroups);

    function seriesFillGrayK(idx, count) {
      if (count <= 1) return makeK(15);
      var k0 = 12, k1 = 40;
      var t = (count === 1) ? 0 : (idx / (count - 1));
      return makeK(k0 + (k1 - k0) * t);
    }

    var showVals = !!cfg.showValueLabels;
    var valFont = Number(cfg.valueFontSizePt);
    if (isNaN(valFont) || valFont <= 0) valFont = 8;

    var valDec = Math.max(0, Math.min(6, Math.round(Number(cfg.valueDecimals))));
    if (isNaN(valDec)) valDec = 0;

    var vOff = mm(Number(cfg.valueLabelOffsetMm));
    if (isNaN(vOff)) vOff = mm(5);

    for (var r = 0; r < parsed.categories.length; r++) {
      for (var sIdx = 0; sIdx < parsed.seriesNames.length; sIdx++) {
        var val = parsed.values[r][sIdx];
        if (val === null || val === undefined) val = 0;

        var yTop = mapY(val, axisMin, axisMax, plotBottomY, plotH);
        var h = yTop - plotBottomY;
        if (h < 0) h = 0;

        var x = layout.xPos[r][sIdx];
        var fill = seriesFillGrayK(sIdx, parsed.seriesNames.length);

        var bar = addTopRoundedBar(gSeries, yTop, x, bw, h, rad, fill);
        bar.name = "bar_" + r + "_" + sIdx;

        if (showVals) {
          var lbl = formatValue(val, valDec, cfg.valueSuffix);
          var tx = x + (bw / 2);
          var tf = addPointText(gValues, tx, yTop + vOff, lbl, valFont, valueCol, Justification.CENTER, fontObj);
          tf.name = "val_" + r + "_" + sIdx;
          forceCenterJustification(tf);
          centerTextFrameAtX(tf, tx);
        }
      }
    }

    drawLine(gAxes, plotLeft, plotBottomY, plotLeft + plotW, plotBottomY, axisCol, axisStrokePt, 100).name = "xAxis";

    function drawYAxisAt(x, name) {
      var ln = drawLine(gAxes, x, plotTopY, x, plotBottomY, axisCol, axisStrokePt, 100);
      ln.name = name;
    }
    if (cfg.yAxisLineMode === "left" || cfg.yAxisLineMode === "both") drawYAxisAt(plotLeft, "yAxisL");
    if (cfg.yAxisLineMode === "right" || cfg.yAxisLineMode === "both") drawYAxisAt(plotLeft + plotW, "yAxisR");

    if (cfg.showAxisLabels && cfg.yLabelMode !== "none") {
      var axisFont = Number(cfg.axisFontSizePt);
      if (isNaN(axisFont) || axisFont <= 0) axisFont = 8;

      var yLift = mm(Number(cfg.axisLabelLiftMm));
      if (isNaN(yLift)) yLift = mm(2.5);

      var xOff = mm(Number(cfg.axisLabelOffsetMm));
      if (isNaN(xOff)) xOff = mm(2);

      var vv2 = axisMin;
      var guard2 = 0;
      while (vv2 <= axisMax + 1e-9 && guard2 < 500) {
        var ay = mapY(vv2, axisMin, axisMax, plotBottomY, plotH);
        var txt = String(vv2);

        if (cfg.yLabelMode === "left" || cfg.yLabelMode === "both") {
          addPointText(gAxisLabelsL, plotLeft - xOff, ay + yLift, txt, axisFont, textCol, Justification.RIGHT, fontObj).name = "yLblL_" + guard2;
        }
        if (cfg.yLabelMode === "right" || cfg.yLabelMode === "both") {
          addPointText(gAxisLabelsR, plotLeft + plotW + xOff, ay + yLift, txt, axisFont, textCol, Justification.LEFT, fontObj).name = "yLblR_" + guard2;
        }

        vv2 += axisStep;
        guard2++;
      }
    }

    if (cfg.showCategoryLabels) {
      var catFont = Number(cfg.categoryFontSizePt);
      if (isNaN(catFont) || catFont <= 0) catFont = 9;

      var catGapMm = Number(cfg.categoryLabelGapMm);
      if (isNaN(catGapMm)) catGapMm = 1.5;

      for (var cc = 0; cc < parsed.categories.length; cc++) {
        var first = layout.xPos[cc][0];
        var last = layout.xPos[cc][parsed.seriesNames.length - 1] + bw;
        var xCenter = (first + last) / 2;
        var yCat = plotBottomY - mm(catGapMm);
        addPointText(gCats, xCenter, yCat, parsed.categories[cc], catFont, textCol, Justification.CENTER, fontObj).name = "cat_" + cc;
      }
    }

    if (cfg.legendOn && parsed.seriesNames.length > 1) {
      var baseY = plotBottomY - mm(Number(cfg.categoryLabelGapMm || 1.5)) - mm(8);
      var fills = [];
      for (var iS = 0; iS < parsed.seriesNames.length; iS++) fills.push(seriesFillGrayK(iS, parsed.seriesNames.length));
      drawLegend(gLegend, plotLeft, plotW, baseY, parsed.seriesNames, fills, cfg, fontObj);
    }

    g.selected = true;
    return g;
  }

  // -----------------------------
  // Styles extraction (Values Only)
  // -----------------------------
  function getFirstTextStyle(g) {
    if (!g || !g.textFrames || g.textFrames.length === 0) return null;
    var tf = g.textFrames[0];
    try {
      var tr = tf.textRange;
      return {
        size: tr.characterAttributes.size,
        fillColor: tr.characterAttributes.fillColor,
        font: tr.characterAttributes.textFont,
        just: tr.paragraphAttributes.justification
      };
    } catch (e) { return null; }
  }

  function getFirstPathStyle(g) {
    if (!g) return null;
    for (var i = 0; i < g.pathItems.length; i++) {
      var p = g.pathItems[i];
      if (p.stroked) {
        return {
          strokeWidth: p.strokeWidth,
          strokeColor: p.strokeColor,
          opacity: (p.opacity !== undefined ? p.opacity : 100)
        };
      }
    }
    return null;
  }

  // -----------------------------
  // Values Only update
  // -----------------------------
  function updateValuesOnly(_doc, targetGroup, cfgNew) {
    var cfgOld = safeJSONParse(targetGroup.note);
    if (!cfgOld) return drawChartFull(_doc, cfgNew, targetGroup);

    var parsedNew = parseData(cfgNew.rawData, cfgNew.groupMode);
    if (!parsedNew) {
      alert("데이터 파싱 실패");
      return null;
    }

    if (cfgOld._categoryCount !== parsedNew.categories.length || cfgOld._seriesCount !== parsedNew.seriesNames.length) {
      alert("Values Only는 카테고리/시리즈 개수가 같을 때만 가능합니다.\n(개수가 바뀌었으니 Full Update로 다시 생성합니다.)");
      return drawChartFull(_doc, cfgNew, targetGroup);
    }

    var gGrid = findChildGroupByName(targetGroup, "Grid");
    var gSeries = findChildGroupByName(targetGroup, "Series");
    var gAxes = findChildGroupByName(targetGroup, "Axes");
    var gValues = findChildGroupByName(targetGroup, "ValueLabels");
    var gCats = findChildGroupByName(targetGroup, "CategoryLabels");
    var gAxisLabelsL = findChildGroupByName(targetGroup, "AxisLabelsLeft");
    var gAxisLabelsR = findChildGroupByName(targetGroup, "AxisLabelsRight");

    if (!gSeries || !gAxes) return drawChartFull(_doc, cfgNew, targetGroup);

    var cfg = cfgOld;
    cfg.rawData = cfgNew.rawData;
    cfg.groupMode = cfgNew.groupMode;

    var ab = getActiveArtboardRect(_doc);
    var left = ab[0], top = ab[1], right = ab[2], bottom = ab[3];
    var abW = right - left, abH = top - bottom;

    var chartW = mm(Number(cfg.chartWidthMm));
    var chartH = mm(Number(cfg.chartHeightMm));
    var originX = left + (abW - chartW) / 2;
    var originY = top - (abH - chartH) / 2;

    var padL = mm(Number(cfg.padLeftMm));
    var padR = mm(Number(cfg.padRightMm));
    var padT = mm(Number(cfg.padTopMm));
    var padB = mm(Number(cfg.padBottomMm));

    var plotLeft = originX + padL;
    var plotTopY = originY - padT;
    var plotW = chartW - padL - padR;
    var plotH = chartH - padT - padB;
    var plotBottomY = plotTopY - plotH;

    var axisStyle = getFirstPathStyle(gAxes) || { strokeWidth: 0.75, strokeColor: axisStrokeDefault(), opacity: 100 };
    var gridStyle = getFirstPathStyle(gGrid) || { strokeWidth: 0.3, strokeColor: gridStrokeDefault(), opacity: cfg.gridOpacityFixed };

    var valTextStyle = getFirstTextStyle(gValues);
    var catTextStyle = getFirstTextStyle(gCats);
    var yLTextStyle = getFirstTextStyle(gAxisLabelsL);
    var yRTextStyle = getFirstTextStyle(gAxisLabelsR);

    var fontObj = null;
    if (valTextStyle && valTextStyle.font) fontObj = valTextStyle.font;
    else if (catTextStyle && catTextStyle.font) fontObj = catTextStyle.font;
    else if (yLTextStyle && yLTextStyle.font) fontObj = yLTextStyle.font;
    else if (yRTextStyle && yRTextStyle.font) fontObj = yRTextStyle.font;
    else fontObj = resolveFontSmart(cfg.preferredFontPSName, cfg.fallbackFontPSName, cfg.preferredFontFamily, cfg.preferredFontStyle);

    var axisMin = Number(cfg.axisMin);
    var axisMax = Number(cfg.axisMax);
    var axisStep = Number(cfg.axisStep);

    if (cfg.axisAuto) {
      var s = autoScale(parsedNew.dataMin, parsedNew.dataMax, 5);
      axisMin = s.min; axisMax = s.max; axisStep = s.step;
      if (cfg.headroomTick) {
        if (parsedNew.dataMax >= axisMax - 1e-9) axisMax += axisStep;
      }
    }

    cfg._resolvedAxisMin = axisMin;
    cfg._resolvedAxisMax = axisMax;
    cfg._resolvedAxisStep = axisStep;

    var bw = mm(Number(cfg.barWidthMm));
    var bg = mm(Number(cfg.barGapMm));
    if (isNaN(bg) || bg <= 0) bg = mm(4);
    var gg = bg * 2.5;
    var rad = mm(Number(cfg.barRadiusMm));

    var geo = resolveBarGeometryAdaptive(plotW, parsedNew.categories.length, parsedNew.seriesNames.length, bw, bg, gg, !!cfg.allowBarWidthExpand);
    bw = geo.bw; bg = geo.bg; gg = geo.gg;

    var layout = computeBarLayoutNice(plotLeft, plotW, parsedNew.categories.length, parsedNew.seriesNames.length, bw, bg, gg, !!cfg.autoDistributeGroups);

    function findBarByName(name) {
      for (var i = 0; i < gSeries.pathItems.length; i++) if (gSeries.pathItems[i].name === name) return gSeries.pathItems[i];
      return null;
    }
    function findTextByName(group, name) {
      if (!group || !group.textFrames) return null;
      for (var i = 0; i < group.textFrames.length; i++) if (group.textFrames[i].name === name) return group.textFrames[i];
      return null;
    }

    if (cfg.showCategoryLabels && gCats) {
      for (var cc = 0; cc < parsedNew.categories.length; cc++) {
        var tcat = findTextByName(gCats, "cat_" + cc);
        if (tcat) tcat.contents = parsedNew.categories[cc];
      }
    }

    var showVals = !!cfg.showValueLabels;
    var valFont = (valTextStyle && valTextStyle.size) ? valTextStyle.size : Number(cfg.valueFontSizePt);
    if (isNaN(valFont) || valFont <= 0) valFont = 8;

    var vOff = mm(Number(cfg.valueLabelOffsetMm));
    if (isNaN(vOff)) vOff = mm(5);

    var valDec = Math.max(0, Math.min(6, Math.round(Number(cfg.valueDecimals))));
    if (isNaN(valDec)) valDec = 0;

    for (var r = 0; r < parsedNew.categories.length; r++) {
      for (var sIdx = 0; sIdx < parsedNew.seriesNames.length; sIdx++) {
        var nm = "bar_" + r + "_" + sIdx;
        var oldBar = findBarByName(nm);
        var fillKeep = null;
        if (oldBar && oldBar.filled) fillKeep = oldBar.fillColor;

        var val = parsedNew.values[r][sIdx];
        if (val === null || val === undefined) val = 0;

        if (oldBar) { try { oldBar.remove(); } catch (e) {} }

        var yTop = mapY(val, axisMin, axisMax, plotBottomY, plotH);
        var h = yTop - plotBottomY;
        if (h < 0) h = 0;

        var x = layout.xPos[r][sIdx];
        var bar = addTopRoundedBar(gSeries, yTop, x, bw, h, rad, fillKeep || makeK(15));
        bar.name = nm;

        if (showVals && gValues) {
          var tv = findTextByName(gValues, "val_" + r + "_" + sIdx);
          var lbl = formatValue(val, valDec, cfg.valueSuffix);
          var tx = x + (bw / 2);

          if (!tv) {
            tv = addPointText(gValues, tx, yTop + vOff, lbl, valFont, valueTextDefault(), Justification.CENTER, fontObj);
            tv.name = "val_" + r + "_" + sIdx;
            forceCenterJustification(tv);
            centerTextFrameAtX(tv, tx);
          } else {
            tv.contents = lbl;

            // 순서 중요: 정렬이 position의 기준점을 흔들 수 있음
            forceCenterJustification(tv);

            tv.position = [tx, yTop + vOff];

            // 마지막에 bounds 중심으로 강제 재센터링(우측으로 미세 이동 방지)
            centerTextFrameAtX(tv, tx);
          }
        }
      }
    }

    if (gGrid) clearGroup(gGrid);
    if (gAxes) clearGroup(gAxes);
    if (gAxisLabelsL) clearGroup(gAxisLabelsL);
    if (gAxisLabelsR) clearGroup(gAxisLabelsR);

    if (cfg.gridShow && gGrid) {
      var v = axisMin, guard = 0;
      while (v <= axisMax + 1e-9 && guard < 500) {
        var yy = mapY(v, axisMin, axisMax, plotBottomY, plotH);
        drawLine(gGrid, plotLeft, yy, plotLeft + plotW, yy, gridStyle.strokeColor, gridStyle.strokeWidth, cfg.gridOpacityFixed);
        v += axisStep;
        guard++;
      }
    }

    drawLine(gAxes, plotLeft, plotBottomY, plotLeft + plotW, plotBottomY, axisStyle.strokeColor, axisStyle.strokeWidth, 100).name = "xAxis";

    function drawYAxisAt(x, name) {
      var ln = drawLine(gAxes, x, plotTopY, x, plotBottomY, axisStyle.strokeColor, axisStyle.strokeWidth, 100);
      ln.name = name;
    }
    if (cfg.yAxisLineMode === "left" || cfg.yAxisLineMode === "both") drawYAxisAt(plotLeft, "yAxisL");
    if (cfg.yAxisLineMode === "right" || cfg.yAxisLineMode === "both") drawYAxisAt(plotLeft + plotW, "yAxisR");

    if (cfg.showAxisLabels && cfg.yLabelMode !== "none") {
      var axisFont = (yLTextStyle && yLTextStyle.size) ? yLTextStyle.size : Number(cfg.axisFontSizePt);
      if (isNaN(axisFont) || axisFont <= 0) axisFont = 8;

      var yLift = mm(Number(cfg.axisLabelLiftMm));
      if (isNaN(yLift)) yLift = mm(2.5);
      var xOff = mm(Number(cfg.axisLabelOffsetMm));
      if (isNaN(xOff)) xOff = mm(2);

      var col = labelTextDefault();

      var vv2 = axisMin, guard2 = 0;
      while (vv2 <= axisMax + 1e-9 && guard2 < 500) {
        var ay = mapY(vv2, axisMin, axisMax, plotBottomY, plotH);
        var txt = String(vv2);

        if (cfg.yLabelMode === "left" || cfg.yLabelMode === "both") {
          addPointText(gAxisLabelsL, plotLeft - xOff, ay + yLift, txt, axisFont, col, Justification.RIGHT, fontObj).name = "yLblL_" + guard2;
        }
        if (cfg.yLabelMode === "right" || cfg.yLabelMode === "both") {
          addPointText(gAxisLabelsR, plotLeft + plotW + xOff, ay + yLift, txt, axisFont, col, Justification.LEFT, fontObj).name = "yLblR_" + guard2;
        }

        vv2 += axisStep;
        guard2++;
      }
    }

    targetGroup.note = safeJSONStringify(cfg);
    targetGroup.selected = true;
    return targetGroup;
  }

  // -----------------------------
  // UI (리사이즈 + 스크롤)
  // -----------------------------
  var sel = doc.selection;
  var target = null;
  var existing = null;

  if (sel.length > 0) {
    var found = findChartGroupFromSelection(sel[0]);
    if (found) { target = found; existing = safeJSONParse(found.note); }
  }

  var cfg = existing || {
    type: "ParamBarChart",
    version: "1.0",

    rawData: "A\t20\nB\t12\nC\t5\nD\t10",
    groupMode: false,

    chartWidthMm: 160,
    chartHeightMm: 70,

    padLeftMm: 8,
    padRightMm: 5,
    padTopMm: 8,
    padBottomMm: 8,

    barWidthMm: 10,
    barGapMm: 4,
    barRadiusMm: 2,

    autoDistributeGroups: true,
    allowBarWidthExpand: true,

    axisAuto: true,
    headroomTick: true,

    axisMin: 0,
    axisMax: 10,
    axisStep: 2,

    showAxisLabels: true,
    yLabelMode: "left",
    yAxisLineMode: "none",

    axisFontSizePt: 8,
    axisLabelLiftMm: 2.5,
    axisLabelOffsetMm: 2,

    gridShow: true,
    gridOpacityFixed: 60,

    gridThicknessPt: 0.3,
    axisThicknessPt: 0.75,

    showValueLabels: true,
    valueFontSizePt: 8,
    valueLabelOffsetMm: 5,
    valueDecimals: 0,
    valueSuffix: "",

    showCategoryLabels: true,
    categoryFontSizePt: 9,
    categoryLabelGapMm: 1.5,

    legendOn: false,
    legendFontPt: 8,
    legendTextLiftMm: 2.7, // 기본값 변경: 2.7mm

    preferredFontPSName: "KoPubWorldDotumPL",
    preferredFontFamily: "KoPubWorld돋움체_Pro",
    preferredFontStyle: "Light",
    fallbackFontPSName: "KoPubWorldDotumPM"
  };

  var win = new Window(
    "dialog",
    target ? "Parametric Bar Chart (Update)" : "Parametric Bar Chart (Create)",
    undefined,
    { resizeable: true }
  );
  win.alignChildren = ["fill", "fill"];
  win.orientation = "column";
  win.preferredSize = [640, 640];
  win.minimumSize = [520, 420];

  // Scroll container (content scrolls, buttons stay visible)
  var scrollWrap = win.add("group");
  scrollWrap.orientation = "row";
  scrollWrap.alignChildren = ["fill", "fill"];
  scrollWrap.alignment = ["fill", "fill"];

  var view = scrollWrap.add("panel", undefined, "");
  view.alignChildren = ["fill", "top"];
  view.alignment = ["fill", "fill"];
  view.margins = 12;

  var sb = scrollWrap.add("scrollbar");
  sb.preferredSize.width = 16;
  sb.alignment = ["right", "fill"];
  sb.visible = true;

  var content = view.add("group");
  content.orientation = "column";
  content.alignChildren = ["fill", "top"];
  content.alignment = ["fill", "top"];

  function contentHeight() {
    try {
      var b = content.bounds; // [l,t,r,b]
      return (b[3] - b[1]);
    } catch (e) { return 0; }
  }

  function refreshScroll() {
    try {
      win.layout.layout(true);
      win.layout.resize();
    } catch (e0) {}

    var vh = 0;
    try { vh = view.size.height - 24; } catch (e1) { vh = 0; }
    if (vh < 80) vh = 80;

    var ch = contentHeight();
    if (ch < 0) ch = 0;

    var maxScroll = Math.max(0, ch - vh);
    sb.minvalue = 0;
    sb.maxvalue = maxScroll;

    if (sb.value > maxScroll) sb.value = maxScroll;
    content.location = [content.location[0], 0 - sb.value];

    // 내용이 충분히 짧으면 스크롤 숨기기(선택)
    // sb.visible = (maxScroll > 1);
  }

  sb.onChanging = function () {
    content.location = [content.location[0], 0 - sb.value];
  };

  // Mouse wheel scroll (가능한 엔진에서만 동작)
  try {
    view.addEventListener("mousewheel", function (e) {
      var step = 30;
      var nv = sb.value - (e.wheelDelta > 0 ? step : -step);
      sb.value = clamp(nv, sb.minvalue, sb.maxvalue);
      content.location = [content.location[0], 0 - sb.value];
    });
  } catch (eW) {}

  win.onResizing = win.onResize = function () {
    try { this.layout.resize(); } catch (e) {}
    refreshScroll();
  };

  // -----------------------------
  // Data
  // -----------------------------
  var pData = content.add("panel", undefined, "Data (Paste TSV/CSV)");
  pData.alignChildren = ["fill", "top"];
  var dataInput = pData.add("edittext", [0, 0, 600, 140], cfg.rawData, { multiline: true });

  var gMode = pData.add("group");
  var cbGroup = gMode.add("checkbox", undefined, "Group mode (first row = series/legend names)");
  cbGroup.value = !!cfg.groupMode;

  // Update options
  var updateMode = "full";
  var rbFull = null, rbVals = null;
  if (target) {
    var pUpd = content.add("panel", undefined, "Update");
    pUpd.alignChildren = ["left", "top"];
    rbFull = pUpd.add("radiobutton", undefined, "Full Update (rebuild)");
    rbVals = pUpd.add("radiobutton", undefined, "Values Only (keep style/layout)");
    rbVals.value = true;
    rbFull.value = false;
    updateMode = "values";
  }

  // Layout
  var pLayout = content.add("panel", undefined, "Layout");
  pLayout.alignChildren = ["left", "top"];

  var gSize = pLayout.add("group");
  gSize.add("statictext", undefined, "Chart W(mm):");
  var inW = gSize.add("edittext", undefined, String(cfg.chartWidthMm)); inW.characters = 6;
  gSize.add("statictext", undefined, "H(mm):");
  var inH = gSize.add("edittext", undefined, String(cfg.chartHeightMm)); inH.characters = 6;

  var gPad = pLayout.add("group");
  gPad.add("statictext", undefined, "Pad L:");
  var inPL = gPad.add("edittext", undefined, String(cfg.padLeftMm)); inPL.characters = 4;
  gPad.add("statictext", undefined, "R:");
  var inPR = gPad.add("edittext", undefined, String(cfg.padRightMm)); inPR.characters = 4;
  gPad.add("statictext", undefined, "T:");
  var inPT = gPad.add("edittext", undefined, String(cfg.padTopMm)); inPT.characters = 4;
  gPad.add("statictext", undefined, "B:");
  var inPB = gPad.add("edittext", undefined, String(cfg.padBottomMm)); inPB.characters = 4;

  var gBar = pLayout.add("group");
  gBar.add("statictext", undefined, "Bar W(mm):");
  var inBW = gBar.add("edittext", undefined, String(cfg.barWidthMm)); inBW.characters = 5;
  gBar.add("statictext", undefined, "Bar Gap(mm):");
  var inBG = gBar.add("edittext", undefined, String(cfg.barGapMm)); inBG.characters = 5;

  var gBar2 = pLayout.add("group");
  gBar2.add("statictext", undefined, "Top Radius(mm):");
  var inBR = gBar2.add("edittext", undefined, String(cfg.barRadiusMm)); inBR.characters = 5;

  var gAuto = pLayout.add("group");
  var cbAutoDist = gAuto.add("checkbox", undefined, "Auto distribute (balanced outer + gaps)");
  cbAutoDist.value = !!cfg.autoDistributeGroups;

  // Axis/Grid
  var pAxis = content.add("panel", undefined, "Axis/Grid");
  pAxis.alignChildren = ["left", "top"];

  var cbAutoScale = pAxis.add("checkbox", undefined, "Auto Scale Y");
  cbAutoScale.value = !!cfg.axisAuto;

  var cbHeadroom = pAxis.add("checkbox", undefined, "Headroom tick (+1 step when max hits top)");
  cbHeadroom.value = !!cfg.headroomTick;

  var gMan = pAxis.add("group");
  gMan.add("statictext", undefined, "Min:");
  var inMin = gMan.add("edittext", undefined, String(cfg.axisMin)); inMin.characters = 6;
  gMan.add("statictext", undefined, "Max:");
  var inMax = gMan.add("edittext", undefined, String(cfg.axisMax)); inMax.characters = 6;
  gMan.add("statictext", undefined, "Step:");
  var inStep = gMan.add("edittext", undefined, String(cfg.axisStep)); inStep.characters = 6;

  var cbYLbl = pAxis.add("checkbox", undefined, "Show Y Labels");
  cbYLbl.value = !!cfg.showAxisLabels;

  var gYMode = pAxis.add("group");
  gYMode.add("statictext", undefined, "Y label mode:");
  var ddYLblMode = gYMode.add("dropdownlist", undefined, ["left", "right", "both", "none"]);
  ddYLblMode.selection = 0;
  try { ddYLblMode.selection = ["left","right","both","none"].indexOf(cfg.yLabelMode); } catch(e) {}
  if (ddYLblMode.selection === null) ddYLblMode.selection = 0;

  gYMode.add("statictext", undefined, "Y axis line:");
  var ddYAxisLine = gYMode.add("dropdownlist", undefined, ["none", "left", "right", "both"]);
  ddYAxisLine.selection = 0;
  try { ddYAxisLine.selection = ["none","left","right","both"].indexOf(cfg.yAxisLineMode); } catch(e2) {}
  if (ddYAxisLine.selection === null) ddYAxisLine.selection = 0;

  var gY = pAxis.add("group");
  gY.add("statictext", undefined, "Font(pt):");
  var inYFont = gY.add("edittext", undefined, String(cfg.axisFontSizePt)); inYFont.characters = 4;
  gY.add("statictext", undefined, "Y Lift(mm):");
  var inYLift = gY.add("edittext", undefined, String(cfg.axisLabelLiftMm)); inYLift.characters = 4;
  gY.add("statictext", undefined, "Offset(mm):");
  var inYOff = gY.add("edittext", undefined, String(cfg.axisLabelOffsetMm)); inYOff.characters = 4;

  var cbGrid = pAxis.add("checkbox", undefined, "Show Grid");
  cbGrid.value = !!cfg.gridShow;

  // Labels
  var pLab = content.add("panel", undefined, "Labels");
  pLab.alignChildren = ["left", "top"];

  var cbVals = pLab.add("checkbox", undefined, "Show Value Labels");
  cbVals.value = !!cfg.showValueLabels;

  var gVal = pLab.add("group");
  gVal.add("statictext", undefined, "Value Font(pt):");
  var inValFont = gVal.add("edittext", undefined, String(cfg.valueFontSizePt)); inValFont.characters = 4;
  gVal.add("statictext", undefined, "Offset(mm, +up):");
  var inValOff = gVal.add("edittext", undefined, String(cfg.valueLabelOffsetMm)); inValOff.characters = 4;
  gVal.add("statictext", undefined, "Decimals:");
  var inValDec = gVal.add("edittext", undefined, String(cfg.valueDecimals)); inValDec.characters = 3;

  var gSuf = pLab.add("group");
  gSuf.add("statictext", undefined, "Suffix:");
  var inSuffix = gSuf.add("edittext", undefined, String(cfg.valueSuffix || "")); inSuffix.characters = 10;

  var cbCats = pLab.add("checkbox", undefined, "Show Category Labels");
  cbCats.value = !!cfg.showCategoryLabels;

  var gCat = pLab.add("group");
  gCat.add("statictext", undefined, "Cat Font(pt):");
  var inCatFont = gCat.add("edittext", undefined, String(cfg.categoryFontSizePt)); inCatFont.characters = 4;
  gCat.add("statictext", undefined, "Cat Gap(mm):");
  var inCatGap = gCat.add("edittext", undefined, String(cfg.categoryLabelGapMm)); inCatGap.characters = 4;

  // Legend
  var pLegend = content.add("panel", undefined, "Legend");
  pLegend.alignChildren = ["left", "top"];
  var cbLegend = pLegend.add("checkbox", undefined, "Legend ON (below, centered, wrap)");
  cbLegend.value = !!cfg.legendOn;

  var gLeg2 = pLegend.add("group");
  gLeg2.add("statictext", undefined, "Text Lift(mm):");
  var inLegLift = gLeg2.add("edittext", undefined, String(cfg.legendTextLiftMm)); inLegLift.characters = 4;
  gLeg2.add("statictext", undefined, "Font(pt):");
  var inLegFont = gLeg2.add("edittext", undefined, String(cfg.legendFontPt)); inLegFont.characters = 4;

  function refreshUI() {
    var manual = !cbAutoScale.value;
    inMin.enabled = manual;
    inMax.enabled = manual;
    inStep.enabled = manual;

    ddYLblMode.enabled = cbYLbl.value;
    ddYAxisLine.enabled = true;

    inLegLift.enabled = cbLegend.value;
    inLegFont.enabled = cbLegend.value;

    refreshScroll();
  }
  cbAutoScale.onClick = refreshUI;
  cbYLbl.onClick = refreshUI;
  cbLegend.onClick = refreshUI;

  // Buttons (고정)
  var btns = win.add("group");
  btns.alignment = "right";
  var bCancel = btns.add("button", undefined, "Cancel");
  var bOK = btns.add("button", undefined, target ? "Update" : "Create");

  bCancel.onClick = function () { win.close(0); };

  bOK.onClick = function () {
    cfg.type = "ParamBarChart";
    cfg.version = "1.0";

    cfg.rawData = dataInput.text;
    cfg.groupMode = cbGroup.value;

    cfg.chartWidthMm = Number(inW.text);
    cfg.chartHeightMm = Number(inH.text);

    cfg.padLeftMm = Number(inPL.text);
    cfg.padRightMm = Number(inPR.text);
    cfg.padTopMm = Number(inPT.text);
    cfg.padBottomMm = Number(inPB.text);

    cfg.barWidthMm = Number(inBW.text);
    cfg.barGapMm = Number(inBG.text);
    cfg.barRadiusMm = Number(inBR.text);

    cfg.autoDistributeGroups = cbAutoDist.value;
    cfg.allowBarWidthExpand = true;

    cfg.axisAuto = cbAutoScale.value;
    cfg.headroomTick = cbHeadroom.value;

    cfg.axisMin = Number(inMin.text);
    cfg.axisMax = Number(inMax.text);
    cfg.axisStep = Number(inStep.text);

    cfg.showAxisLabels = cbYLbl.value;
    cfg.yLabelMode = ddYLblMode.selection ? ddYLblMode.selection.text : "left";
    cfg.yAxisLineMode = ddYAxisLine.selection ? ddYAxisLine.selection.text : "none";

    cfg.axisFontSizePt = Number(inYFont.text);
    cfg.axisLabelLiftMm = Number(inYLift.text);
    cfg.axisLabelOffsetMm = Number(inYOff.text);

    cfg.gridShow = cbGrid.value;

    cfg.showValueLabels = cbVals.value;
    cfg.valueFontSizePt = Number(inValFont.text);
    cfg.valueLabelOffsetMm = Number(inValOff.text);
    cfg.valueDecimals = Number(inValDec.text);
    cfg.valueSuffix = inSuffix.text;

    cfg.showCategoryLabels = cbCats.value;
    cfg.categoryFontSizePt = Number(inCatFont.text);
    cfg.categoryLabelGapMm = Number(inCatGap.text);

    cfg.legendOn = cbLegend.value;
    cfg.legendTextLiftMm = Number(inLegLift.text);
    cfg.legendFontPt = Number(inLegFont.text);

    cfg.gridOpacityFixed = 60;
    cfg.gridThicknessPt = cfg.gridThicknessPt || 0.3;
    cfg.axisThicknessPt = cfg.axisThicknessPt || 0.75;

    if (!cfg.rawData || cfg.rawData.trim().length === 0) {
      alert("데이터가 비어 있습니다.");
      return;
    }
    if (isNaN(cfg.chartWidthMm) || cfg.chartWidthMm <= 0 || isNaN(cfg.chartHeightMm) || cfg.chartHeightMm <= 0) {
      alert("차트 크기(mm)를 확인해 주세요.");
      return;
    }
    if (!cfg.axisAuto) {
      if (isNaN(cfg.axisMin) || isNaN(cfg.axisMax) || cfg.axisMax === cfg.axisMin) {
        alert("Manual Axis: Min/Max를 확인해 주세요.");
        return;
      }
      if (isNaN(cfg.axisStep) || cfg.axisStep <= 0) {
        alert("Manual Axis: Step을 확인해 주세요.");
        return;
      }
    }

    win.close(1);
  };

  // 초기 레이아웃/스크롤 세팅
  win.onShow = function () {
    refreshUI();
    refreshScroll();
  };

  var res = win.show();
  if (res !== 1) return;

  if (target && rbFull && rbVals) {
    updateMode = rbVals.value ? "values" : "full";
  }

  try { app.executeMenuCommand("deselectall"); } catch (e) {}

  if (!target) {
    drawChartFull(doc, cfg, null);
  } else {
    if (updateMode === "values") updateValuesOnly(doc, target, cfg);
    else drawChartFull(doc, cfg, target);
  }

})();
