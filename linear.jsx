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
    LINE_OUTLINE_COLOR: "#000000",

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

    STATION_OUTLINE: "#000000",

    // Footer
    FOOTER_TEXT: "Rail Fans Canada — 2026",
    FOOTER_FONT_SIZE: 11,
    FOOTER_COLOR: "#9B9B9B",
    FOOTER_BOTTOM_MARGIN: 15,

    // Title
    DRAW_TITLE: true,
    TITLE_TEMPLATE: "{name}{years_paren}", // tokens: {system} {id} {name} {years} {years_paren}
    TITLE_FONT_SIZE: 44,
    TITLE_COLOR: "#000000",
    TITLE_LEFT: 14,
    TITLE_TOP_FROM_TOP: 18,

    // Subtitle
    DRAW_SUBTITLE: false,
    SUBTITLE_TEMPLATE: "{system}", // tokens: {system} {id} {name} {years} {years_paren}
    SUBTITLE_FONT_SIZE: 18,
    SUBTITLE_COLOR: "#5A5A5A",
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

    // If empty: export next to the selected lines folder
    // If set:
    //   - absolute path: "C:/path/to/output" or "/Users/name/output"
    //   - relative path: "exports-svg" (relative to the selected lines folder)
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

  function normalizeHex(hex) {
  var h = trim(hex || "");
  if (!h) throw new Error("Missing HEX color.");
  if (h.charAt(0) !== "#") h = "#" + h;
  if (!/^#([0-9a-fA-F]{6})$/.test(h)) throw new Error("Invalid HEX color: " + hex);
  return h.toUpperCase();
}

function hexToRgb255(hex) {
  var h = normalizeHex(hex).substring(1);
  return {
    r: parseInt(h.substring(0, 2), 16),
    g: parseInt(h.substring(2, 4), 16),
    b: parseInt(h.substring(4, 6), 16)
  };
}

function hexColor(hex) {
  var o = hexToRgb255(hex);
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
    bg.fillColor = hexColor("#FFFFFF");
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
    footer.textRange.characterAttributes.fillColor = hexColor(CFG.FOOTER_COLOR);

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
    title.textRange.characterAttributes.fillColor = hexColor(CFG.TITLE_COLOR);
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
    sub.textRange.characterAttributes.fillColor = hexColor(CFG.SUBTITLE_COLOR);
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

  function resolveExportBaseFolder(baseFolder) {
    var dest = trim(CFG.EXPORT_DESTINATION_FOLDER || "");
    if (!dest) return baseFolder;

    var isAbs = (/^[A-Za-z]\:/.test(dest) || dest.charAt(0) === "/" || dest.charAt(0) === "\\");
    var f = isAbs ? Folder(dest) : Folder(baseFolder.fsName + "/" + dest);

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
  // Load line JSON files (one file per line)
  // =========================
  function listJSONFiles(folder) {
    var files = folder.getFiles(function (f) {
      return (f instanceof File) && /\.json$/i.test(f.name);
    });
    return files;
  }

  function parseLineJSONFile(file) {
    var raw = trim(readFile(file));
    if (raw && raw.charCodeAt(0) === 0xFEFF) raw = raw.substring(1);

    var obj = safeParseJSON(raw);
    if (obj === null) throw new Error("Invalid JSON in: " + file.fsName);

    // Required fields (mirrors CSV semantics)
    // system -> operator/network name
    // region -> becomes L.id (token {id})
    // line_name -> becomes L.name (token {name})
    // years -> display years
    // color -> HEX string "#RRGGBB"
    // stations -> array of station names OR pipe-delimited string
    var system = trim(obj.system || "");
    var region = trim(obj.region || "");
    var lineName = trim(obj.line_name || "");
    var years = trim(obj.years || "");
    var color = trim(obj.color || "");
    var stations = obj.stations;

    if (!system) throw new Error("Missing required field: system (" + file.name + ")");
    if (!region) throw new Error("Missing required field: region (" + file.name + ")");
    if (!lineName) throw new Error("Missing required field: line_name (" + file.name + ")");
    if (!color) throw new Error("Missing required field: color (" + file.name + ")");

    color = normalizeHex(color);

    var stationList = [];
    if (stations instanceof Array) {
      for (var i = 0; i < stations.length; i++) stationList.push(trim(stations[i]));
    } else if (typeof stations === "string") {
      var parts = stations.split("|");
      for (var j = 0; j < parts.length; j++) stationList.push(trim(parts[j]));
    } else {
      throw new Error("stations must be an array or a pipe-delimited string (" + file.name + ")");
    }

    // remove empties
    var cleaned = [];
    for (var k = 0; k < stationList.length; k++) if (stationList[k]) cleaned.push(stationList[k]);
    if (cleaned.length < 2) throw new Error("stations must contain at least 2 entries (" + file.name + ")");

    return {
      system: system,
      id: region,
      name: lineName,
      years: years,
      color: color,      // HEX string
      stations: cleaned
    };
  }

  var scriptFolder = getScriptFolder();
  var linesFolder = null;

  if (scriptFolder) {
    var candidateFolder = Folder(scriptFolder.fsName + "/lines");
    if (candidateFolder.exists) linesFolder = candidateFolder;
  }

  if (!linesFolder) {
    linesFolder = Folder.selectDialog("Select folder containing line JSON files");
    if (!linesFolder) {
      alert("No lines folder selected.");
      return;
    }
  }

  var jsonFiles = listJSONFiles(linesFolder);
  if (!jsonFiles.length) {
    alert("No .json files found in:\n" + linesFolder.fsName);
    return;
  }

  var LINES = [];
  for (var f = 0; f < jsonFiles.length; f++) {
    try {
      LINES.push(parseLineJSONFile(jsonFiles[f]));
    } catch (eLine) {
      alert("Failed to parse line file:" + jsonFiles[f].fsName + "" + eLine.message);
      return;
    }
  }


  // =========================
  // Export folder resolution
  // =========================
  var exportBaseFolder = null;
  if (CFG.EXPORT_SVG) {
    try {
      exportBaseFolder = resolveExportBaseFolder(linesFolder.parent);
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
  var stationStroke = hexColor(CFG.STATION_OUTLINE);
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
      outline.strokeColor = hexColor(CFG.LINE_OUTLINE_COLOR);
      outline.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);
    }

    // Main line
    var line = doc.pathItems.add();
    line.stroked = true;
    line.filled = false;
    line.strokeWidth = CFG.LINE_STROKE;
    line.strokeCap = StrokeCap.ROUNDENDCAP;
    line.strokeColor = hexColor(L.color);
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

  alert("Generated " + LINES.length + " diagrams from:\n" + linesFolder.fsName + exportMsg);
})();
