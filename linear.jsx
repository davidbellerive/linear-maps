#target illustrator

(function () {
  // =========================
  // CONFIG
  // =========================
  var CFG_DEFAULTS = {
    ARTBOARD_WIDTH: 1500,
    ARTBOARD_HEIGHT: 300,
    BASELINE_Y: 48,
    H_PADDING: 90,

    AUTO_PAD_TERMINALS: true,
    PAD_RIGHT_MARGIN: 8,
    PAD_LEFT_MARGIN: 8,
    PAD_MAX_EXTRA: 260,
    PAD_CHECK_LEFT: false,

    LINE_STROKE: 16,

    LINE_OUTLINE_ENABLED: false,
    LINE_OUTLINE_WIDTH: 3,
    LINE_OUTLINE_COLOR: { r: 0, g: 0, b: 0 },

    STATION_RADIUS: 10,
    STATION_STROKE_WIDTH: null, // null = auto from LINE_STROKE

    LABEL_MODE: "angled",
    LABEL_TILT: -20,
    LABEL_CLEARANCE: 10,
    LABEL_X_NUDGE: -6,

    // Fonts (independent, with fallback to ArialMT)
    FONT_LABEL_NAME: "ArialMT",
    FONT_TITLE_NAME: "ArialMT",
    FONT_SUBTITLE_NAME: "ArialMT",
    FONT_FOOTER_NAME: "ArialMT",

    FONT_SIZE: 18,

    STATION_OUTLINE: { r: 0, g: 0, b: 0 },

    // Footer
    FOOTER_TEXT: "Rail Fans Canada — 2026",
    FOOTER_FONT_SIZE: 11,
    FOOTER_COLOR: { r: 155, g: 155, b: 155 },
    FOOTER_BOTTOM_MARGIN: 15,

    // Title
    DRAW_TITLE: true,
    TITLE_TEMPLATE: "{name}{years_paren}", // tokens: {system} {id} {name} {years} {years_paren}
    TITLE_FONT_SIZE: 44,
    TITLE_COLOR: { r: 0, g: 0, b: 0 },
    TITLE_LEFT: 14,
    TITLE_TOP_FROM_TOP: 18,

    // Subtitle
    DRAW_SUBTITLE: false,
    SUBTITLE_TEMPLATE: "{system}", // tokens: {system} {id} {name} {years} {years_paren}
    SUBTITLE_FONT_SIZE: 18,
    SUBTITLE_COLOR: { r: 90, g: 90, b: 90 },
    SUBTITLE_LEFT: 14,         // used only when no title exists
    SUBTITLE_TOP_FROM_TOP: 70,

    // Dynamic height
    AUTO_HEIGHT: true,
    MIN_HEIGHT: 220,

    RESERVE_TITLE_BAND: true,
    TITLE_BAND_PADDING: 10,
    TITLE_BAND_LINE_GAP: 10,

    // Export
    EXPORT_SVG: true,

    // If empty: export next to lines.csv (legacy behavior)
    // If set:
    //   - absolute path: "C:/path/to/output" or "/Users/name/output"
    //   - relative path: "exports-svg" (relative to lines.csv folder)
    EXPORT_DESTINATION_FOLDER: "",

    // If true: create a subfolder per system (CSV col 1)
    EXPORT_GROUP_BY_SYSTEM: false,

    // Output file name (without extension). Default uses line_name (CSV field 3 => L.name)
    // Tokens: {system} {id} {name} {years} {years_paren}
    EXPORT_FILENAME_TEMPLATE: "{name}",

    OUTLINE_TEXT_FOR_SVG: true,
    CLOSE_AFTER_EXPORT: false
  };

  var CFG = null;

  // =========================
  // Helpers
  // =========================
  function trim(s) { return (s || "").replace(/^\s+|\s+$/g, ""); }

  function getScriptFolder() {
    try { return File($.fileName).parent; }
    catch (e) { return null; }
  }

  function readFile(file) {
    file.encoding = "UTF-8";
    if (!file.open("r")) throw new Error("Unable to open file: " + file.fsName);
    var txt = file.read();
    file.close();
    return txt;
  }

  function safeParseJSON(text) {
    try {
      if (typeof JSON !== "undefined" && JSON.parse) return JSON.parse(text);
    } catch (e1) {}
    try { return eval("(" + text + ")"); } catch (e2) { return null; }
  }

  function isPlainObject(x) {
    return x && (typeof x === "object") && !(x instanceof Array);
  }

  function mergeDeep(base, overrides) {
    for (var k in overrides) {
      if (!overrides.hasOwnProperty(k)) continue;
      var v = overrides[k];
      if (isPlainObject(v) && isPlainObject(base[k])) {
        mergeDeep(base[k], v);
      } else {
        base[k] = v;
      }
    }
    return base;
  }

  // Font getter: if requested font doesn't exist, fallback to ArialMT
  function getFontOrArial(fontName) {
    var name = trim(fontName || "");
    if (!name) name = "ArialMT";
    try { return app.textFonts.getByName(name); }
    catch (e) { return app.textFonts.getByName("ArialMT"); }
  }

  function rgb(o) {
    var c = new RGBColor();
    c.red = o.r; c.green = o.g; c.blue = o.b;
    return c;
  }

  function safeFileName(s) {
    return (s || "")
      .replace(/[\\\/\:\*\?\"\<\>\|]/g, "_")
      .replace(/\s+/g, " ")
      .replace(/^\s+|\s+$/g, "");
  }

  function loadConfigFromJSONFile(cfgFile) {
    var raw = trim(readFile(cfgFile));
    if (raw && raw.charCodeAt(0) === 0xFEFF) raw = raw.substring(1);

    var parsed = safeParseJSON(raw);
    if (parsed === null) throw new Error("Invalid JSON in: " + cfgFile.fsName);

    if (parsed.CFG && isPlainObject(parsed.CFG)) return parsed.CFG;
    return parsed;
  }

  function buildConfig() {
    var scriptFolder = getScriptFolder();
    var cfgFile = null;

    if (scriptFolder) {
      var candidate = File(scriptFolder.fsName + "/linear-config.json");
      if (candidate.exists) cfgFile = candidate;
    }

    if (!cfgFile) {
      cfgFile = File.openDialog("Select linear-config.json", "JSON:*.json");
      if (!cfgFile) throw new Error("No linear-config.json selected.");
    }

    var overrides = loadConfigFromJSONFile(cfgFile);

    var cfg = {};
    for (var k in CFG_DEFAULTS) cfg[k] = CFG_DEFAULTS[k];
    mergeDeep(cfg, overrides);

    // Ensure font defaults always exist (and always fallback to ArialMT)
    if (!cfg.FONT_LABEL_NAME) cfg.FONT_LABEL_NAME = "ArialMT";
    if (!cfg.FONT_TITLE_NAME) cfg.FONT_TITLE_NAME = "ArialMT";
    if (!cfg.FONT_SUBTITLE_NAME) cfg.FONT_SUBTITLE_NAME = "ArialMT";
    if (!cfg.FONT_FOOTER_NAME) cfg.FONT_FOOTER_NAME = "ArialMT";

    // Warn unknown keys (typo catcher)
    var unknown = [];
    for (var ok in overrides) {
      if (overrides.hasOwnProperty(ok) && !CFG_DEFAULTS.hasOwnProperty(ok)) unknown.push(ok);
    }
    if (unknown.length) {
      alert(
        "linear-config.json contains " + unknown.length + " unknown key(s).\n" +
        "They will still be applied, but double-check for typos:\n\n" +
        unknown.join(", ")
      );
    }

    cfg.__CONFIG_FILE__ = cfgFile;
    return cfg;
  }

  try { CFG = buildConfig(); }
  catch (cfgErr) {
    alert("Config load failed:\n" + cfgErr.message);
    return;
  }

  // CSV splitter that respects quotes and escaped quotes
  function splitCSVLine(line) {
    var out = [];
    var cur = "";
    var inQuotes = false;

    for (var i = 0; i < line.length; i++) {
      var ch = line.charAt(i);

      if (ch === '"') {
        if (inQuotes && i + 1 < line.length && line.charAt(i + 1) === '"') {
          cur += '"'; i++;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (ch === "," && !inQuotes) {
        out.push(cur); cur = "";
      } else {
        cur += ch;
      }
    }
    out.push(cur);
    return out;
  }

  function isIntLike(s) {
    s = trim(s);
    if (s === "") return false;
    return /^-?\d+$/.test(s);
  }

  function clamp255(n) {
    n = parseInt(n, 10);
    if (isNaN(n)) return 0;
    return Math.max(0, Math.min(255, n));
  }

  // system, id, name, years, then scan for first 3 integer cols = r,g,b
  // remaining cols are stations (usually one quoted field).
  function parseLinesCSV(text) {
    var rows = text.split(/\r\n|\n|\r/);
    var parsed = [];

    for (var i = 0; i < rows.length; i++) {
      var raw = trim(rows[i]);
      if (!raw) continue;
      if (raw.indexOf("#") === 0) continue;
      if (i === 0 && raw.toLowerCase().indexOf("system,") === 0) continue;

      var cols = splitCSVLine(raw);
      if (cols.length < 7) continue;

      var system = trim(cols[0]);
      var line_id = trim(cols[1]);
      var line_name = trim(cols[2]);
      var years = trim(cols[3]);

      var rgbStart = -1;
      for (var k = 4; k < cols.length; k++) {
        if (isIntLike(cols[k]) && (k + 2) < cols.length && isIntLike(cols[k + 1]) && isIntLike(cols[k + 2])) {
          rgbStart = k;
          break;
        }
      }
      if (rgbStart === -1) continue;

      var r = clamp255(cols[rgbStart]);
      var g = clamp255(cols[rgbStart + 1]);
      var b = clamp255(cols[rgbStart + 2]);

      var stationsRaw = "";
      for (var j = rgbStart + 3; j < cols.length; j++) {
        if (stationsRaw) stationsRaw += ",";
        stationsRaw += cols[j];
      }
      stationsRaw = trim(stationsRaw);

      if (stationsRaw.length >= 2 && stationsRaw.charAt(0) === '"' && stationsRaw.charAt(stationsRaw.length - 1) === '"') {
        stationsRaw = stationsRaw.substring(1, stationsRaw.length - 1);
      }

      var stations = stationsRaw.split("|");
      for (var s = 0; s < stations.length; s++) stations[s] = trim(stations[s]);
      if (stations.length < 2) continue;

      parsed.push({
        system: system,
        id: line_id,
        name: line_name,
        years: years,
        color: { r: r, g: g, b: b },
        stations: stations
      });
    }

    return parsed;
  }

  function getLabelRotation() {
    if (CFG.LABEL_MODE === "vertical") return 90;
    return 90 + CFG.LABEL_TILT;
  }

  function getArtboardHeight(doc) {
    var r = doc.artboards[0].artboardRect;
    return r[1] - r[3];
  }

  function addWhiteBackground(doc) {
    var H = getArtboardHeight(doc);
    var bg = doc.pathItems.rectangle(H, 0, CFG.ARTBOARD_WIDTH, H);
    bg.stroked = false;
    bg.filled = true;
    bg.fillColor = rgb({ r: 255, g: 255, b: 255 });
    bg.zOrder(ZOrderMethod.SENDTOBACK);
    bg.locked = true;
    bg.name = "bg-white";
    return bg;
  }

  function getMaxTopOfTextFrames(doc) {
    var maxTop = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {
      var tf = doc.textFrames[i];
      try {
        var gb = tf.geometricBounds; // [L,T,R,B]
        maxTop = Math.max(maxTop, gb[1]);
      } catch (e) {}
    }
    return maxTop;
  }

  function resizeArtboardHeight(doc, newH) {
    var ab = doc.artboards[0];
    ab.artboardRect = [0, newH, CFG.ARTBOARD_WIDTH, 0];

    // rebuild bg
    for (var i = doc.pathItems.length - 1; i >= 0; i--) {
      var p = doc.pathItems[i];
      if (p.name === "bg-white") {
        try { p.locked = false; } catch (e1) {}
        try { p.remove(); } catch (e2) {}
        break;
      }
    }
    addWhiteBackground(doc);
  }

  function estimateTitleBandHeight() {
    return (CFG.TITLE_FONT_SIZE * 1.35) + CFG.TITLE_BAND_PADDING;
  }

  function outlineAllText(doc) {
    for (var i = doc.textFrames.length - 1; i >= 0; i--) {
      var tf = doc.textFrames[i];
      try {
        tf.createOutline();
        tf.remove();
      } catch (e) {}
    }
  }

  function drawFooter(doc, text, fontObj) {
    if (!text) return;

    var footer = doc.textFrames.add();
    footer.contents = text;
    footer.textRange.characterAttributes.size = CFG.FOOTER_FONT_SIZE;
    footer.textRange.characterAttributes.textFont = fontObj;
    footer.textRange.paragraphAttributes.justification = Justification.CENTER;
    footer.textRange.characterAttributes.fillColor = rgb(CFG.FOOTER_COLOR);

    footer.left = 0;
    footer.top = CFG.FOOTER_BOTTOM_MARGIN;

    var gb = footer.geometricBounds;
    var w = gb[2] - gb[0];
    footer.left = (CFG.ARTBOARD_WIDTH - w) / 2;
  }

  // ---- template helpers ----
  function applyTemplate(tpl, L) {
    var years = trim(L.years || "");
    var yearsParen = years ? (" (" + years + ")") : "";

    var s = tpl || "";
    s = s.split("{system}").join(L.system || "");
    s = s.split("{id}").join(L.id || "");
    s = s.split("{name}").join(L.name || "");
    s = s.split("{years}").join(years);
    s = s.split("{years_paren}").join(yearsParen);

    return s;
  }

  // ----- OUTLINE-BASED LEFT ALIGNMENT (reliable) -----
  function getOutlinedLeft(tf) {
    // Duplicate → outline → measure → delete, so we don't destroy live text
    var dup = null;
    var g = null;
    try {
      dup = tf.duplicate();
      // keep it out of the way; does not need to be visible
      dup.hidden = true;

      // Illustrator sometimes needs a redraw before createOutline/bounds settle
      app.redraw();

      g = dup.createOutline(); // GroupItem
      app.redraw();

      var gb = g.geometricBounds; // [L,T,R,B] of outline geometry
      var left = gb[0];

      try { g.remove(); } catch (e1) {}
      try { dup.remove(); } catch (e2) {}

      return left;
    } catch (e) {
      try { if (g) g.remove(); } catch (e3) {}
      try { if (dup) dup.remove(); } catch (e4) {}
      // fallback to geometricBounds if outlining fails
      try { return tf.geometricBounds[0]; } catch (e5) { return tf.left; }
    }
  }

  function snapTextOutlineLeft(tf, targetX) {
    // move tf so its OUTLINED left edge equals targetX
    var leftNow = getOutlinedLeft(tf);
    var dx = targetX - leftNow;
    tf.left += dx;
    app.redraw();
    return getOutlinedLeft(tf);
  }

  function drawTitle(doc, L, fontObj) {
    if (!CFG.DRAW_TITLE) return null;

    var text = applyTemplate(CFG.TITLE_TEMPLATE, L);
    if (!trim(text)) return null;

    var H = getArtboardHeight(doc);

    var title = doc.textFrames.add();
    title.contents = text;
    title.textRange.characterAttributes.size = CFG.TITLE_FONT_SIZE;
    title.textRange.characterAttributes.textFont = fontObj;
    title.textRange.characterAttributes.fillColor = rgb(CFG.TITLE_COLOR);
    title.textRange.paragraphAttributes.justification = Justification.LEFT;

    title.left = CFG.TITLE_LEFT;
    title.top  = H - CFG.TITLE_TOP_FROM_TOP;

    // Make title's outlined-left EXACTLY TITLE_LEFT
    var titleOutlineLeft = snapTextOutlineLeft(title, CFG.TITLE_LEFT);

    return { frame: title, outlineLeft: titleOutlineLeft };
  }

  function drawSubtitle(doc, L, fontObj, alignOutlineLeftToX) {
    if (!CFG.DRAW_SUBTITLE) return null;

    var text = applyTemplate(CFG.SUBTITLE_TEMPLATE, L);
    if (!trim(text)) return null;

    var H = getArtboardHeight(doc);

    var sub = doc.textFrames.add();
    sub.contents = text;
    sub.textRange.characterAttributes.size = CFG.SUBTITLE_FONT_SIZE;
    sub.textRange.characterAttributes.textFont = fontObj;
    sub.textRange.characterAttributes.fillColor = rgb(CFG.SUBTITLE_COLOR);
    sub.textRange.paragraphAttributes.justification = Justification.LEFT;

    // initial placement
    sub.left = CFG.SUBTITLE_LEFT;
    sub.top  = H - CFG.SUBTITLE_TOP_FROM_TOP;

    // If a title outline-left is provided, match it; otherwise snap to SUBTITLE_LEFT
    var target = (alignOutlineLeftToX !== null && alignOutlineLeftToX !== undefined)
      ? alignOutlineLeftToX
      : CFG.SUBTITLE_LEFT;

    snapTextOutlineLeft(sub, target);

    return sub;
  }

  function ensureFolder(parentFolder, name) {
    var f = Folder(parentFolder.fsName + "/" + name);
    if (!f.exists) {
      if (!f.create()) throw new Error("Could not create folder: " + f.fsName);
    }
    return f;
  }

  function resolveExportBaseFolder(linesFile) {
    var dest = trim(CFG.EXPORT_DESTINATION_FOLDER || "");
    if (!dest) return linesFile.parent;

    var isAbs = (/^[A-Za-z]\:/.test(dest) || dest.charAt(0) === "/" || dest.charAt(0) === "\\");
    var f = isAbs ? Folder(dest) : Folder(linesFile.parent.fsName + "/" + dest);

    if (!f.exists) {
      if (!f.create()) throw new Error("Could not create export destination: " + f.fsName);
    }
    return f;
  }

  function exportAsSVG(doc, destFile) {
    var opts = new ExportOptionsSVG();
    opts.embedRasterImages = true;
    opts.fontSubsetting = SVGFontSubsetting.GLYPHSUSED;
    opts.coordinatePrecision = 2;
    opts.cssProperties = SVGCSSPropertyLocation.PRESENTATIONATTRIBUTES;
    doc.exportFile(destFile, ExportType.SVG, opts);
  }

  function measureLabelOverhang(doc, text, anchorX, anchorY, rotation, fontObj) {
    var tf = doc.textFrames.add();
    tf.contents = text;
    tf.textRange.characterAttributes.size = CFG.FONT_SIZE;
    tf.textRange.characterAttributes.textFont = fontObj;
    tf.textRange.paragraphAttributes.justification = Justification.LEFT;

    tf.left = anchorX;
    tf.top  = anchorY;
    tf.rotate(rotation);

    var gb = tf.geometricBounds;
    tf.left += (anchorX - gb[0]);
    tf.left += CFG.LABEL_X_NUDGE;

    gb = tf.geometricBounds;
    var minX = gb[0];
    var maxX = gb[2];

    tf.remove();

    return {
      leftOverhang: Math.max(0, anchorX - minX),
      rightOverhang: Math.max(0, maxX - anchorX)
    };
  }

  function computeTerminalExtraPadding(doc, stations, basePad, rotation, fontObj, baselineY) {
    var lineLeft = basePad;
    var lineRight = CFG.ARTBOARD_WIDTH - basePad;

    var firstX = lineLeft;
    var lastX  = lineRight;

    var extraLeft = 0;
    var extraRight = 0;

    var lastText = stations[stations.length - 1];
    var mLast = measureLabelOverhang(doc, lastText, lastX, baselineY, rotation, fontObj);

    extraRight = Math.min(
      CFG.PAD_MAX_EXTRA,
      Math.max(0, (mLast.rightOverhang + CFG.PAD_RIGHT_MARGIN) - basePad)
    );

    if (CFG.PAD_CHECK_LEFT) {
      var firstText = stations[0];
      var mFirst = measureLabelOverhang(doc, firstText, firstX, baselineY, rotation, fontObj);
      extraLeft = Math.min(
        CFG.PAD_MAX_EXTRA,
        Math.max(0, (mFirst.leftOverhang + CFG.PAD_LEFT_MARGIN) - basePad)
      );
    }

    return { extraLeft: extraLeft, extraRight: extraRight };
  }

  // =========================
  // Load lines CSV (lines.csv)
  // =========================
  var scriptFolder = getScriptFolder();
  var dataFile = null;

  if (scriptFolder) {
    var candidateCSV = File(scriptFolder.fsName + "/lines.csv");
    if (candidateCSV.exists) dataFile = candidateCSV;
  }

  if (!dataFile) {
    dataFile = File.openDialog("Select lines.csv", "CSV:*.csv");
    if (!dataFile) {
      alert("No lines.csv selected.");
      return;
    }
  }

  var csvText;
  try { csvText = readFile(dataFile); }
  catch (e) {
    alert("Failed to read CSV:\n" + e.message);
    return;
  }

  var LINES;
  try { LINES = parseLinesCSV(csvText); }
  catch (e2) {
    alert("Failed to parse CSV:\n" + e2.message);
    return;
  }

  if (!LINES.length) {
    alert("No valid rows found in CSV.");
    return;
  }

  // =========================
  // Export folder resolution
  // =========================
  var exportBaseFolder = null;
  if (CFG.EXPORT_SVG) {
    try {
      exportBaseFolder = resolveExportBaseFolder(dataFile);
    } catch (eDest) {
      alert(eDest.message);
      return;
    }
  }

  // =========================
  // Fonts (resolved once, all fallback to ArialMT)
  // =========================
  var labelFontObj = getFontOrArial(CFG.FONT_LABEL_NAME);
  var titleFontObj = getFontOrArial(CFG.FONT_TITLE_NAME);
  var subtitleFontObj = getFontOrArial(CFG.FONT_SUBTITLE_NAME);
  var footerFontObj = getFontOrArial(CFG.FONT_FOOTER_NAME);

  // Shared objects
  var stationStroke = rgb(CFG.STATION_OUTLINE);
  var labelRotation = getLabelRotation();

  // =========================
  // Generate docs
  // =========================
  for (var li = 0; li < LINES.length; li++) {
    var L = LINES[li];

    var doc = app.documents.add(DocumentColorSpace.RGB, CFG.ARTBOARD_WIDTH, CFG.ARTBOARD_HEIGHT);
    doc.rulerUnits = RulerUnits.Pixels;

    var ab = doc.artboards[0];
    ab.artboardRect = [0, CFG.ARTBOARD_HEIGHT, CFG.ARTBOARD_WIDTH, 0];

    addWhiteBackground(doc);

    var baselineY = CFG.BASELINE_Y;

    // padding
    var padLeft = CFG.H_PADDING;
    var padRight = CFG.H_PADDING;

    if (CFG.AUTO_PAD_TERMINALS) {
      var extras = computeTerminalExtraPadding(doc, L.stations, CFG.H_PADDING, labelRotation, labelFontObj, baselineY);
      padLeft  = CFG.H_PADDING + extras.extraLeft;
      padRight = CFG.H_PADDING + extras.extraRight;
    }

    var lineLeft = padLeft;
    var lineRight = CFG.ARTBOARD_WIDTH - padRight;

    if (lineRight - lineLeft < 200) {
      lineLeft = CFG.H_PADDING;
      lineRight = CFG.ARTBOARD_WIDTH - CFG.H_PADDING;
    }

    var spacing = (lineRight - lineLeft) / (L.stations.length - 1);

    // Optional line outline
    if (CFG.LINE_OUTLINE_ENABLED) {
      var outline = doc.pathItems.add();
      outline.stroked = true;
      outline.filled = false;
      outline.strokeWidth = CFG.LINE_STROKE + (CFG.LINE_OUTLINE_WIDTH * 2);
      outline.strokeCap = StrokeCap.ROUNDENDCAP;
      outline.strokeColor = rgb(CFG.LINE_OUTLINE_COLOR);
      outline.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);
    }

    // Main line
    var line = doc.pathItems.add();
    line.stroked = true;
    line.filled = false;
    line.strokeWidth = CFG.LINE_STROKE;
    line.strokeCap = StrokeCap.ROUNDENDCAP;
    line.strokeColor = rgb(L.color);
    line.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);

    var white = new RGBColor();
    white.red = 255; white.green = 255; white.blue = 255;

    var lineTopY = baselineY + (CFG.LINE_STROKE / 2);

    // Stations + labels
    for (var i = 0; i < L.stations.length; i++) {
      var x = lineLeft + i * spacing;

      var dot = doc.pathItems.ellipse(
        baselineY + CFG.STATION_RADIUS,
        x - CFG.STATION_RADIUS,
        CFG.STATION_RADIUS * 2,
        CFG.STATION_RADIUS * 2
      );
      dot.filled = true;
      dot.stroked = true;
      dot.fillColor = white;
      dot.strokeColor = stationStroke;

      var stationStrokeWidth =
        (CFG.STATION_STROKE_WIDTH !== null && CFG.STATION_STROKE_WIDTH !== undefined)
          ? CFG.STATION_STROKE_WIDTH
          : Math.max(2, Math.round(CFG.LINE_STROKE * 0.35));

      dot.strokeWidth = stationStrokeWidth;

      var label = doc.textFrames.add();
      label.contents = L.stations[i];
      label.textRange.characterAttributes.size = CFG.FONT_SIZE;
      label.textRange.characterAttributes.textFont = labelFontObj;
      label.textRange.paragraphAttributes.justification = Justification.LEFT;

      label.left = x;
      label.top = baselineY;

      label.rotate(labelRotation);

      var gb = label.geometricBounds;
      label.left += (x - gb[0]);
      label.left += CFG.LABEL_X_NUDGE;

      gb = label.geometricBounds;
      label.top += ((lineTopY + CFG.LABEL_CLEARANCE) - gb[3]);
    }

    // Footer first (so height calc includes it if needed)
    drawFooter(doc, CFG.FOOTER_TEXT, footerFontObj);

    // Dynamic height pass
    if (CFG.AUTO_HEIGHT) {
      var maxTextTop = getMaxTopOfTextFrames(doc);

      var reserve = 0;
      if (CFG.DRAW_TITLE && CFG.RESERVE_TITLE_BAND) {
        reserve = estimateTitleBandHeight() + CFG.TITLE_TOP_FROM_TOP + CFG.TITLE_BAND_LINE_GAP;
      }

      var neededH = Math.ceil(Math.max(CFG.MIN_HEIGHT, maxTextTop + reserve));
      resizeArtboardHeight(doc, neededH);
    }

    // Title + Subtitle after resizing (subtitle aligned to TITLE via outline-left)
    var titleInfo = drawTitle(doc, L, titleFontObj);
    var titleOutlineLeft = titleInfo ? titleInfo.outlineLeft : null;
    drawSubtitle(doc, L, subtitleFontObj, titleOutlineLeft);

    // =========================
    // Export SVG
    // =========================
    if (CFG.EXPORT_SVG && exportBaseFolder) {
      var targetFolder = exportBaseFolder;

      if (CFG.EXPORT_GROUP_BY_SYSTEM) {
        try {
          targetFolder = ensureFolder(exportBaseFolder, safeFileName(L.system));
        } catch (eSys) {
          alert("Could not create system folder:\n" + eSys.message);
          targetFolder = exportBaseFolder;
        }
      }

      var baseName = applyTemplate(CFG.EXPORT_FILENAME_TEMPLATE, L);
      baseName = safeFileName(baseName);
      if (!baseName) baseName = safeFileName(L.name || (L.system + "-" + L.id));

      var svgFile = File(targetFolder.fsName + "/" + baseName + ".svg");

      try {
        if (CFG.OUTLINE_TEXT_FOR_SVG) outlineAllText(doc);
        exportAsSVG(doc, svgFile);
      } catch (eSvg) {
        alert("SVG export failed for:\n" + baseName + "\n\n" + eSvg.message);
      }

      if (CFG.CLOSE_AFTER_EXPORT) {
        doc.close(SaveOptions.DONOTSAVECHANGES);
      }
    }
  }

  var exportMsg = "";
  if (CFG.EXPORT_SVG) {
    exportMsg =
      "\n\nExported SVGs to:\n" + exportBaseFolder.fsName +
      (CFG.EXPORT_GROUP_BY_SYSTEM ? "\n(With system subfolders)" : "");
  }

  alert("Generated " + LINES.length + " diagrams from:\n" + dataFile.fsName + exportMsg);
})();
