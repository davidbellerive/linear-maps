#target illustrator

(function () {
  // =========================
  // CONFIG DEFAULTS
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

    // Fonts (independent, fallback to ArialMT)
    FONT_LABEL_NAME: "ArialMT",
    FONT_TITLE_NAME: "ArialMT",
    FONT_SUBTITLE_NAME: "ArialMT",
    FONT_FOOTER_NAME: "ArialMT",

    FONT_SIZE: 18,

    STATION_OUTLINE: "#000000",

    // Footer
    FOOTER_TEXT: "Rail Fans Canada \u2014 2026",
    FOOTER_FONT_SIZE: 11,
    FOOTER_COLOR: "#9B9B9B",
    FOOTER_BOTTOM_MARGIN: 15,

    // Title
    DRAW_TITLE: true,
    TITLE_TEMPLATE: "{name}{years_paren}", // tokens: {system} {region} {id} {name} {years} {years_paren}
    TITLE_FONT_SIZE: 44,
    TITLE_COLOR: "#000000",
    TITLE_LEFT: 14,
    TITLE_TOP_FROM_TOP: 18,

    // Subtitle
    DRAW_SUBTITLE: false,
    SUBTITLE_TEMPLATE: "{system} \u2022 {region} \u2022 {years}", // tokens: {system} {region} {id} {name} {years} {years_paren}
    SUBTITLE_FONT_SIZE: 18,
    SUBTITLE_COLOR: "#5A5A5A",
    SUBTITLE_LEFT: 14,
    SUBTITLE_TOP_FROM_TOP: 70,

    // Dynamic height
    AUTO_HEIGHT: true,
    MIN_HEIGHT: 220,

    RESERVE_TITLE_BAND: true,
    TITLE_BAND_PADDING: 10,
    TITLE_BAND_LINE_GAP: 10,

    // Connection marker pills (rendered below station dot)
    DRAW_CONNECTIONS: true,
    CONNECTION_PILL_HEIGHT: 12,    // height of the colored pill
    CONNECTION_PILL_MIN_WIDTH: 20, // minimum pill width
    CONNECTION_PILL_RADIUS: 3,     // corner radius
    CONNECTION_FONT_SIZE: 8,
    CONNECTION_Y_GAP: 4,           // gap between station dot bottom and first pill
    CONNECTION_Y_SPACING: 3,       // gap between stacked pills

    // Export
    EXPORT_SVG: true,
    // absolute path: "C:/path/to/output" or "/Users/name/output"
    // relative path: "exports-svg" (relative to the system folder)
    // empty: export next to the system folder
    EXPORT_DESTINATION_FOLDER: "",
    EXPORT_GROUP_BY_SYSTEM: false,
    // tokens: {system} {region} {id} {name} {years} {years_paren}
    EXPORT_FILENAME_TEMPLATE: "{name}",

    OUTLINE_TEXT_FOR_SVG: true,
    CLOSE_AFTER_EXPORT: false
  };

  var CFG = null;

  // =========================
  // HELPERS
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

  // Font getter: fallback to ArialMT if requested font is unavailable
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

  // =========================
  // CONFIG LOADING
  // =========================
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

    // Ensure font defaults always exist
    if (!cfg.FONT_LABEL_NAME)    cfg.FONT_LABEL_NAME    = "ArialMT";
    if (!cfg.FONT_TITLE_NAME)    cfg.FONT_TITLE_NAME    = "ArialMT";
    if (!cfg.FONT_SUBTITLE_NAME) cfg.FONT_SUBTITLE_NAME = "ArialMT";
    if (!cfg.FONT_FOOTER_NAME)   cfg.FONT_FOOTER_NAME   = "ArialMT";

    // Warn on unknown keys (typo catcher)
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

  // =========================
  // LAYOUT HELPERS
  // =========================
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

  // Measure tallest text frame, excluding the footer (which is repositioned after resize)
  function getMaxTopOfTextFrames(doc) {
    var maxTop = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {
      var tf = doc.textFrames[i];
      if (tf.name === "footer") continue; // footer is placed after height pass
      try {
        var gb = tf.geometricBounds; // [L, T, R, B]
        maxTop = Math.max(maxTop, gb[1]);
      } catch (e) {}
    }
    return maxTop;
  }

  function resizeArtboardHeight(doc, newH) {
    var ab = doc.artboards[0];
    ab.artboardRect = [0, newH, CFG.ARTBOARD_WIDTH, 0];

    // Rebuild background at new height
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

  // Outline all live text for SVG export. Collects and reports any failures.
  function outlineAllText(doc) {
    var failed = [];
    for (var i = doc.textFrames.length - 1; i >= 0; i--) {
      var tf = doc.textFrames[i];
      try {
        tf.createOutline();
      } catch (e) {
        failed.push(tf.contents ? tf.contents.substring(0, 30) : "(unnamed)");
      }
    }
    if (failed.length) {
      alert(
        failed.length + " text frame(s) could not be outlined and may embed fonts in the SVG:\n\n" +
        failed.join("\n")
      );
    }
  }

  // =========================
  // FOOTER
  // Footer is drawn AFTER height resize so it always sits at the bottom correctly.
  // =========================
  function drawFooter(doc, text, fontObj) {
    if (!text) return;

    var H = getArtboardHeight(doc);

    var footer = doc.textFrames.add();
    footer.name = "footer";
    footer.contents = text;
    footer.textRange.characterAttributes.size = CFG.FOOTER_FONT_SIZE;
    footer.textRange.characterAttributes.textFont = fontObj;
    footer.textRange.paragraphAttributes.justification = Justification.CENTER;
    footer.textRange.characterAttributes.fillColor = hexColor(CFG.FOOTER_COLOR);

    var gb = footer.geometricBounds;
    var w = gb[2] - gb[0];
    footer.left = (CFG.ARTBOARD_WIDTH - w) / 2;
    footer.top  = CFG.FOOTER_BOTTOM_MARGIN;
  }

  // =========================
  // TEMPLATE TOKENS
  // {system} {region} {id} {name} {years} {years_paren}
  // Note: {region} is the human-readable region name from stations.json.
  //       {id} is the line's own id field from the line file.
  // =========================
  function applyTemplate(tpl, L) {
    var years = trim(L.years || "");
    var yearsParen = years ? (" (" + years + ")") : "";

    var s = tpl || "";
    s = s.split("{system}").join(L.system || "");
    s = s.split("{region}").join(L.region || "");
    s = s.split("{id}").join(L.id || "");
    s = s.split("{name}").join(L.name || "");
    s = s.split("{years}").join(years);
    s = s.split("{years_paren}").join(yearsParen);

    return s;
  }

  // =========================
  // OUTLINE-BASED LEFT ALIGNMENT
  // Duplicate → outline → measure → delete, preserving live text.
  // =========================
  function getOutlinedLeft(tf) {
    var dup = null;
    var g = null;
    try {
      dup = tf.duplicate();
      dup.hidden = true;
      app.redraw();
      g = dup.createOutline();
      app.redraw();
      var left = g.geometricBounds[0];
      try { g.remove(); }   catch (e1) {}
      try { dup.remove(); } catch (e2) {}
      return left;
    } catch (e) {
      try { if (g)   g.remove(); }   catch (e3) {}
      try { if (dup) dup.remove(); } catch (e4) {}
      try { return tf.geometricBounds[0]; } catch (e5) { return tf.left; }
    }
  }

  function snapTextOutlineLeft(tf, targetX) {
    var leftNow = getOutlinedLeft(tf);
    tf.left += (targetX - leftNow);
    app.redraw();
    return getOutlinedLeft(tf);
  }

  // =========================
  // TITLE + SUBTITLE
  // =========================
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

    sub.left = CFG.SUBTITLE_LEFT;
    sub.top  = H - CFG.SUBTITLE_TOP_FROM_TOP;

    var target = (alignOutlineLeftToX !== null && alignOutlineLeftToX !== undefined)
      ? alignOutlineLeftToX
      : CFG.SUBTITLE_LEFT;

    snapTextOutlineLeft(sub, target);
    return sub;
  }

  // =========================
  // CONNECTION MARKERS
  // Draws colored pills below a station dot for each connection to another line.
  // The current line's own ID is excluded so only transfers are shown.
  // =========================
  function drawConnectionMarkers(doc, station, currentLineId, network, x, dotBottomY, fontObj) {
    if (!CFG.DRAW_CONNECTIONS) return 0;
    if (!station.connections || !station.connections.length) return 0;
    if (!network) return 0;

    var pillH = CFG.CONNECTION_PILL_HEIGHT;
    var pillMinW = CFG.CONNECTION_PILL_MIN_WIDTH;
    var yGap = CFG.CONNECTION_Y_GAP;
    var ySpacing = CFG.CONNECTION_Y_SPACING;
    var totalHeightUsed = 0;

    var currentY = dotBottomY - yGap; // Illustrator Y is top-of-artboard = high value

    for (var c = 0; c < station.connections.length; c++) {
      var connId = station.connections[c];

      // Skip the line currently being rendered — only show transfer targets
      if (connId === currentLineId) continue;

      var connDef = network[connId];
      if (!connDef) continue; // unknown line ID, skip silently

      var connColor = connDef.color ? hexColor(connDef.color) : hexColor("#888888");
      var connLabel = connDef.label || connId;

      // Measure label width with a temporary text frame
      var tmpTf = doc.textFrames.add();
      tmpTf.contents = connLabel;
      tmpTf.textRange.characterAttributes.size = CFG.CONNECTION_FONT_SIZE;
      tmpTf.textRange.characterAttributes.textFont = fontObj;
      var tmpGb = tmpTf.geometricBounds;
      var textW = tmpGb[2] - tmpGb[0];
      tmpTf.remove();

      var pillW = Math.max(pillMinW, textW + 8);
      var pillLeft = x - (pillW / 2);
      var pillTop  = currentY; // top of pill in Illustrator coords

      // Draw rounded rectangle pill
      var pill = doc.pathItems.roundedRectangle(
        pillTop, pillLeft, pillW, pillH,
        CFG.CONNECTION_PILL_RADIUS * 2,
        CFG.CONNECTION_PILL_RADIUS * 2
      );
      pill.filled = true;
      pill.stroked = false;
      pill.fillColor = connColor;

      // Draw label centered in pill
      var lbl = doc.textFrames.add();
      lbl.contents = connLabel;
      lbl.textRange.characterAttributes.size = CFG.CONNECTION_FONT_SIZE;
      lbl.textRange.characterAttributes.textFont = fontObj;
      lbl.textRange.paragraphAttributes.justification = Justification.CENTER;

      var white = new RGBColor();
      white.red = 255; white.green = 255; white.blue = 255;
      lbl.textRange.characterAttributes.fillColor = white;

      var lblGb = lbl.geometricBounds;
      var lblH   = lblGb[1] - lblGb[3];
      lbl.left = pillLeft + (pillW / 2) - ((lblGb[2] - lblGb[0]) / 2);
      lbl.top  = pillTop - (pillH / 2) + (lblH / 2);

      currentY = currentY - pillH - ySpacing;
      totalHeightUsed += pillH + ySpacing;
    }

    return totalHeightUsed;
  }

  // =========================
  // EXPORT HELPERS
  // =========================
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

  // =========================
  // LABEL OVERHANG / PADDING
  // =========================
  function measureLabelOverhang(doc, labelText, anchorX, anchorY, rotation, fontObj) {
    var tf = doc.textFrames.add();
    tf.contents = labelText;
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
      leftOverhang:  Math.max(0, anchorX - minX),
      rightOverhang: Math.max(0, maxX - anchorX)
    };
  }

  function computeTerminalExtraPadding(doc, stations, basePad, rotation, fontObj, baselineY) {
    var lineLeft  = basePad;
    var lineRight = CFG.ARTBOARD_WIDTH - basePad;
    var extraLeft  = 0;
    var extraRight = 0;

    // Right terminal — always checked
    var lastLabel = stations[stations.length - 1].label;
    var mLast = measureLabelOverhang(doc, lastLabel, lineRight, baselineY, rotation, fontObj);
    extraRight = Math.min(
      CFG.PAD_MAX_EXTRA,
      Math.max(0, (mLast.rightOverhang + CFG.PAD_RIGHT_MARGIN) - basePad)
    );

    // Left terminal — opt-in via PAD_CHECK_LEFT
    if (CFG.PAD_CHECK_LEFT) {
      var firstLabel = stations[0].label;
      var mFirst = measureLabelOverhang(doc, firstLabel, lineLeft, baselineY, rotation, fontObj);
      extraLeft = Math.min(
        CFG.PAD_MAX_EXTRA,
        Math.max(0, (mFirst.leftOverhang + CFG.PAD_LEFT_MARGIN) - basePad)
      );
    }

    return { extraLeft: extraLeft, extraRight: extraRight };
  }

  // =========================
  // STATION REGISTRY LOADING
  // Looks for stations.json in the same folder as the line files (system folder).
  // If absent, the script operates in legacy mode (stations array = plain strings).
  // =========================
  function loadStationRegistry(systemFolder) {
    var registryFile = File(systemFolder.fsName + "/stations.json");
    if (!registryFile.exists) return null;

    var raw = trim(readFile(registryFile));
    if (raw && raw.charCodeAt(0) === 0xFEFF) raw = raw.substring(1);

    var parsed = safeParseJSON(raw);
    if (!parsed) throw new Error("Invalid JSON in: " + registryFile.fsName);

    return {
      system:   trim(parsed.system  || ""),
      region:   trim(parsed.region  || ""),
      network:  parsed.network  || {},
      stations: parsed.stations || {}
    };
  }

  // =========================
  // LINE FILE PARSING
  // Resolves station IDs against the registry (if present).
  // Falls back to treating array entries as plain label strings (legacy behaviour).
  // =========================
  function listJSONFiles(folder) {
    return folder.getFiles(function (f) {
      return (f instanceof File) && /\.json$/i.test(f.name) && f.name !== "stations.json";
    });
  }

  function parseLineJSONFile(file, registry) {
    var raw = trim(readFile(file));
    if (raw && raw.charCodeAt(0) === 0xFEFF) raw = raw.substring(1);

    var obj = safeParseJSON(raw);
    if (obj === null) throw new Error("Invalid JSON in: " + file.fsName);

    var lineId   = trim(obj.id        || "");
    var lineName = trim(obj.line_name || "");
    var years    = trim(obj.years     || "");
    var color    = trim(obj.color     || "");
    var rawStations = obj.stations;

    if (!lineName) throw new Error("Missing required field: line_name (" + file.name + ")");
    if (!color)    throw new Error("Missing required field: color (" + file.name + ")");

    color = normalizeHex(color);

    // system + region: prefer registry values; fall back to fields in the line file itself
    var system = registry ? registry.system : trim(obj.system || "");
    var region = registry ? registry.region : trim(obj.region || "");
    var network = registry ? registry.network : {};

    if (!system) throw new Error("Missing system name. Define it in stations.json or in the line file (" + file.name + ")");
    if (!region) throw new Error("Missing region name. Define it in stations.json or in the line file (" + file.name + ")");

    // Resolve station list
    var stationIds = [];
    if (rawStations instanceof Array) {
      for (var i = 0; i < rawStations.length; i++) {
        var entry = rawStations[i];
        stationIds.push(typeof entry === "string" ? trim(entry) : null);
      }
    } else if (typeof rawStations === "string") {
      // Legacy pipe-delimited string
      var parts = rawStations.split("|");
      for (var j = 0; j < parts.length; j++) stationIds.push(trim(parts[j]));
    } else {
      throw new Error("stations must be an array or a pipe-delimited string (" + file.name + ")");
    }

    // Remove empties
    var cleaned = [];
    for (var k = 0; k < stationIds.length; k++) if (stationIds[k]) cleaned.push(stationIds[k]);
    if (cleaned.length < 2) throw new Error("stations must contain at least 2 entries (" + file.name + ")");

    // Resolve station objects
    // With registry: look up each ID and return enriched object.
    // Without registry (legacy): treat each entry as the display label directly.
    var resolvedStations = [];
    for (var s = 0; s < cleaned.length; s++) {
      var sid = cleaned[s];

      if (registry && registry.stations) {
        var staDef = registry.stations[sid];
        if (!staDef) {
          throw new Error(
            "Station ID \"" + sid + "\" not found in stations.json (" + file.name + ")"
          );
        }
        resolvedStations.push({
          id:          sid,
          label:       trim(staDef.label  || sid),
          label2:      trim(staDef.label2 || ""),
          connections: staDef.connections instanceof Array ? staDef.connections : [],
          icons:       staDef.icons       instanceof Array ? staDef.icons       : []
        });
      } else {
        // Legacy mode: entry is a plain label string
        resolvedStations.push({
          id:          sid,
          label:       sid,
          label2:      "",
          connections: [],
          icons:       []
        });
      }
    }

    return {
      id:       lineId,
      system:   system,
      region:   region,
      name:     lineName,
      years:    years,
      color:    color,
      network:  network,
      stations: resolvedStations
    };
  }

  // =========================
  // FOLDER DISCOVERY
  // Script walks up one level looking for system subfolders containing line files.
  // Falls back to a manual folder select if nothing is found automatically.
  // =========================
  var scriptFolder = getScriptFolder();
  var systemFolders = [];

  if (scriptFolder) {
    // Check for a direct "lines/" sibling folder
    var directLines = Folder(scriptFolder.fsName + "/lines");
    if (directLines.exists) {
      // Check if lines/ itself contains system subfolders
      var subFolders = directLines.getFiles(function (f) { return f instanceof Folder; });
      if (subFolders.length > 0) {
        // lines/ contains system subfolders (new structure)
        for (var sf = 0; sf < subFolders.length; sf++) systemFolders.push(subFolders[sf]);
      } else {
        // lines/ is itself a flat folder of line files (legacy structure)
        systemFolders.push(directLines);
      }
    }
  }

  if (!systemFolders.length) {
    var manualFolder = Folder.selectDialog("Select folder containing line JSON files (or a system subfolder)");
    if (!manualFolder) { alert("No folder selected."); return; }
    systemFolders.push(manualFolder);
  }

  // =========================
  // LOAD ALL LINES across all system folders
  // =========================
  var LINES = [];

  for (var sfi = 0; sfi < systemFolders.length; sfi++) {
    var sysFolder = systemFolders[sfi];

    var registry = null;
    try {
      registry = loadStationRegistry(sysFolder);
    } catch (eReg) {
      alert("Failed to load stations.json in:\n" + sysFolder.fsName + "\n\n" + eReg.message);
      return;
    }

    var jsonFiles = listJSONFiles(sysFolder);
    if (!jsonFiles.length) continue;

    for (var f = 0; f < jsonFiles.length; f++) {
      try {
        LINES.push(parseLineJSONFile(jsonFiles[f], registry));
      } catch (eLine) {
        alert("Failed to parse line file:\n" + jsonFiles[f].fsName + "\n\n" + eLine.message);
        return;
      }
    }
  }

  if (!LINES.length) {
    alert("No line files found.");
    return;
  }

  // =========================
  // EXPORT FOLDER RESOLUTION
  // Base is the parent of the first system folder (i.e. the lines/ folder)
  // =========================
  var exportBaseFolder = null;
  if (CFG.EXPORT_SVG) {
    try {
      exportBaseFolder = resolveExportBaseFolder(systemFolders[0].parent);
    } catch (eDest) {
      alert(eDest.message);
      return;
    }
  }

  // =========================
  // FONTS — resolved once, shared across all docs
  // =========================
  var labelFontObj    = getFontOrArial(CFG.FONT_LABEL_NAME);
  var titleFontObj    = getFontOrArial(CFG.FONT_TITLE_NAME);
  var subtitleFontObj = getFontOrArial(CFG.FONT_SUBTITLE_NAME);
  var footerFontObj   = getFontOrArial(CFG.FONT_FOOTER_NAME);
  var connFontObj     = getFontOrArial(CFG.FONT_LABEL_NAME);

  var stationStroke = hexColor(CFG.STATION_OUTLINE);
  var labelRotation = getLabelRotation();

  // =========================
  // GENERATE DOCUMENTS
  // =========================
  for (var li = 0; li < LINES.length; li++) {
    var L = LINES[li];

    var doc = app.documents.add(DocumentColorSpace.RGB, CFG.ARTBOARD_WIDTH, CFG.ARTBOARD_HEIGHT);
    doc.rulerUnits = RulerUnits.Pixels;

    var ab = doc.artboards[0];
    ab.artboardRect = [0, CFG.ARTBOARD_HEIGHT, CFG.ARTBOARD_WIDTH, 0];

    addWhiteBackground(doc);

    var baselineY = CFG.BASELINE_Y;

    // Compute terminal padding
    var padLeft  = CFG.H_PADDING;
    var padRight = CFG.H_PADDING;

    if (CFG.AUTO_PAD_TERMINALS) {
      var extras = computeTerminalExtraPadding(doc, L.stations, CFG.H_PADDING, labelRotation, labelFontObj, baselineY);
      padLeft  = CFG.H_PADDING + extras.extraLeft;
      padRight = CFG.H_PADDING + extras.extraRight;
    }

    var lineLeft  = padLeft;
    var lineRight = CFG.ARTBOARD_WIDTH - padRight;

    if (lineRight - lineLeft < 200) {
      lineLeft  = CFG.H_PADDING;
      lineRight = CFG.ARTBOARD_WIDTH - CFG.H_PADDING;
    }

    var spacing = (lineRight - lineLeft) / (L.stations.length - 1);

    // Optional line outline (decorative border behind the main stroke)
    if (CFG.LINE_OUTLINE_ENABLED) {
      var lineOutline = doc.pathItems.add();
      lineOutline.stroked = true;
      lineOutline.filled = false;
      lineOutline.strokeWidth = CFG.LINE_STROKE + (CFG.LINE_OUTLINE_WIDTH * 2);
      lineOutline.strokeCap = StrokeCap.ROUNDENDCAP;
      lineOutline.strokeColor = hexColor(CFG.LINE_OUTLINE_COLOR);
      lineOutline.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);
    }

    // Main line stroke
    var lineItem = doc.pathItems.add();
    lineItem.stroked = true;
    lineItem.filled = false;
    lineItem.strokeWidth = CFG.LINE_STROKE;
    lineItem.strokeCap = StrokeCap.ROUNDENDCAP;
    lineItem.strokeColor = hexColor(L.color);
    lineItem.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);

    var white = new RGBColor();
    white.red = 255; white.green = 255; white.blue = 255;

    var lineTopY    = baselineY + (CFG.LINE_STROKE / 2);
    var dotBottomY  = baselineY - CFG.STATION_RADIUS; // bottom of station circle in Illustrator Y

    // Stations, labels, connection markers
    for (var i = 0; i < L.stations.length; i++) {
      var station = L.stations[i];
      var x = lineLeft + i * spacing;

      // Station dot
      var stationStrokeWidth =
        (CFG.STATION_STROKE_WIDTH !== null && CFG.STATION_STROKE_WIDTH !== undefined)
          ? CFG.STATION_STROKE_WIDTH
          : Math.max(2, Math.round(CFG.LINE_STROKE * 0.35));

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
      dot.strokeWidth = stationStrokeWidth;

      // Station label
      // label2 is rendered as a second line using \r (native Illustrator line break in JSX)
      var labelText = station.label;
      if (station.label2) labelText = labelText + "\r" + station.label2;

      var label = doc.textFrames.add();
      label.contents = labelText;
      label.textRange.characterAttributes.size = CFG.FONT_SIZE;
      label.textRange.characterAttributes.textFont = labelFontObj;
      label.textRange.paragraphAttributes.justification = Justification.LEFT;

      label.left = x;
      label.top  = baselineY;
      label.rotate(labelRotation);

      var gb = label.geometricBounds;
      label.left += (x - gb[0]);
      label.left += CFG.LABEL_X_NUDGE;

      gb = label.geometricBounds;
      label.top += ((lineTopY + CFG.LABEL_CLEARANCE) - gb[3]);

      // Connection markers below station dot
      drawConnectionMarkers(doc, station, L.id, L.network, x, dotBottomY, connFontObj);
    }

    // Dynamic height pass — footer excluded from measurement, drawn after resize
    if (CFG.AUTO_HEIGHT) {
      var maxTextTop = getMaxTopOfTextFrames(doc);

      var reserve = 0;
      if (CFG.DRAW_TITLE && CFG.RESERVE_TITLE_BAND) {
        reserve = estimateTitleBandHeight() + CFG.TITLE_TOP_FROM_TOP + CFG.TITLE_BAND_LINE_GAP;
      }

      var neededH = Math.ceil(Math.max(CFG.MIN_HEIGHT, maxTextTop + reserve));
      resizeArtboardHeight(doc, neededH);
    }

    // Title + subtitle drawn after resize
    var titleInfo = drawTitle(doc, L, titleFontObj);
    var titleOutlineLeft = titleInfo ? titleInfo.outlineLeft : null;
    drawSubtitle(doc, L, subtitleFontObj, titleOutlineLeft);

    // Footer drawn last, after final canvas height is known
    drawFooter(doc, CFG.FOOTER_TEXT, footerFontObj);

    // =========================
    // SVG EXPORT
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
      (CFG.EXPORT_GROUP_BY_SYSTEM ? "\n(Grouped by system)" : "");
  }

  alert("Generated " + LINES.length + " diagram(s).\n\nSource: " + systemFolders[0].parent.fsName + exportMsg);
})();