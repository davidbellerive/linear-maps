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
    SUBTITLE_TEMPLATE: "{system} \u2022 {region} \u2022 {years}",
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

    // -------------------------------------------------
    // MARKER BLOCK GLOBAL SCALE + FOOTER CLEARANCE
    //
    // MARKER_SCALE: multiplier applied to all connection and icon
    //   sizing values at render time. 1.0 = default. 0.75 = smaller.
    //   1.5 = larger. Adjusting this one value scales the entire
    //   marker + icon block proportionally without touching individual
    //   size keys.
    //
    // MARKER_FOOTER_GAP: minimum vertical space (px) reserved between
    //   the bottom of the icon/marker block and the top of the footer
    //   text. AUTO_HEIGHT factors this in so the canvas is always tall
    //   enough to breathe between content and footer.
    // -------------------------------------------------
    MARKER_SCALE: 1.0,
    MARKER_FOOTER_GAP: 20,

    // -------------------------------------------------
    // CONNECTION MARKERS
    // Rendered below station dot. Shape auto-sizes:
    //   - circle if short_label fits (1-2 chars)
    //   - pill otherwise
    // Asset override: assets/connections/<id>.svg|png|jpg
    // All size values are multiplied by MARKER_SCALE at render time.
    // -------------------------------------------------
    DRAW_CONNECTIONS: true,
    CONNECTION_SIZE: 18,          // base diameter for circle / height for pill
    CONNECTION_PILL_MIN_WIDTH: 26,// base minimum pill width
    CONNECTION_PILL_RADIUS: 4,    // base corner radius for pill shape
    CONNECTION_FONT_SIZE: 9,      // not scaled — text size stays legible
    CONNECTION_Y_GAP: 6,          // base gap: station dot bottom → first marker top
    CONNECTION_Y_SPACING: 3,      // base gap between stacked markers

    // -------------------------------------------------
    // PICTOGRAM ICONS
    // Rendered below connection markers (or at the same
    // baseline if a station has no connections).
    // Asset source: assets/icons/<key>.svg|png|jpg
    // All size values are multiplied by MARKER_SCALE at render time.
    // -------------------------------------------------
    DRAW_ICONS: true,
    ICON_SIZE: 18,                // base rendered width = height (square)
    ICON_X_GAP: 4,                // base horizontal gap between icons in a row
    ICON_Y_GAP: 6,                // base gap: bottom of connection block → first icon row top
    ICON_Y_SPACING: 4,            // base gap between icon rows
    ICON_MAX_PER_ROW: 3,          // wrap after this many icons per row (not scaled)

    // Export
    EXPORT_SVG: true,
    EXPORT_DESTINATION_FOLDER: "",
    EXPORT_GROUP_BY_SYSTEM: false,
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
  // ASSET RESOLUTION
  // Looks for <key>.svg, then .png, then .jpg in a given folder.
  // Returns a File object if found, null otherwise.
  // =========================
  function findAsset(folder, key) {
    if (!folder || !folder.exists) return null;
    var exts = [".svg", ".png", ".jpg", ".jpeg"];
    for (var e = 0; e < exts.length; e++) {
      var f = File(folder.fsName + "/" + key + exts[e]);
      if (f.exists) return f;
    }
    return null;
  }

  // Place a raster or SVG asset file into a document, scaled to fit targetSize x targetSize.
  // Centered on (cx, cy) in Illustrator coordinates.
  // Returns the placed item, or null on failure.
  function placeAsset(doc, assetFile, cx, cy, targetSize) {
    try {
      var item = doc.placedItems.add();
      item.file = assetFile;
      item.embed(); // embed so SVG export is self-contained

      // Scale to fit targetSize
      var gb = item.geometricBounds; // [L, T, R, B]
      var w  = gb[2] - gb[0];
      var h  = gb[1] - gb[3];
      if (w <= 0 || h <= 0) { item.remove(); return null; }

      var scale = (targetSize / Math.max(w, h)) * 100;
      item.resize(scale, scale);

      // Re-read bounds after scale
      gb = item.geometricBounds;
      w  = gb[2] - gb[0];
      h  = gb[1] - gb[3];

      // Center on (cx, cy)
      item.left = cx - (w / 2);
      item.top  = cy + (h / 2);

      return item;
    } catch (e) {
      return null;
    }
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

    if (!cfg.FONT_LABEL_NAME)    cfg.FONT_LABEL_NAME    = "ArialMT";
    if (!cfg.FONT_TITLE_NAME)    cfg.FONT_TITLE_NAME    = "ArialMT";
    if (!cfg.FONT_SUBTITLE_NAME) cfg.FONT_SUBTITLE_NAME = "ArialMT";
    if (!cfg.FONT_FOOTER_NAME)   cfg.FONT_FOOTER_NAME   = "ArialMT";

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
  // MARKER SIZE RESOLUTION
  // All connection + icon size values are multiplied by MARKER_SCALE once
  // and stored in MS (marker sizes). Use MS.* throughout rendering so
  // changing MARKER_SCALE in the config scales the entire block uniformly.
  // CONNECTION_FONT_SIZE is intentionally not scaled — text legibility
  // should be tuned independently.
  // =========================
  var MS = (function () {
    var s = CFG.MARKER_SCALE || 1.0;
    return {
      CONNECTION_SIZE:           CFG.CONNECTION_SIZE           * s,
      CONNECTION_PILL_MIN_WIDTH: CFG.CONNECTION_PILL_MIN_WIDTH * s,
      CONNECTION_PILL_RADIUS:    CFG.CONNECTION_PILL_RADIUS    * s,
      CONNECTION_Y_GAP:          CFG.CONNECTION_Y_GAP          * s,
      CONNECTION_Y_SPACING:      CFG.CONNECTION_Y_SPACING      * s,
      ICON_SIZE:                 CFG.ICON_SIZE                 * s,
      ICON_X_GAP:                CFG.ICON_X_GAP                * s,
      ICON_Y_GAP:                CFG.ICON_Y_GAP                * s,
      ICON_Y_SPACING:            CFG.ICON_Y_SPACING            * s
    };
  })();

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

  // Footer excluded from height measurement — it is drawn after resize
  function getMaxTopOfTextFrames(doc) {
    var maxTop = 0;
    for (var i = 0; i < doc.textFrames.length; i++) {
      var tf = doc.textFrames[i];
      if (tf.name === "footer") continue;
      try {
        var gb = tf.geometricBounds;
        maxTop = Math.max(maxTop, gb[1]);
      } catch (e) {}
    }
    return maxTop;
  }

  // Also factor in the lowest point of any path or placed item (icons, pills)
  function getMaxTopOfAllItems(doc) {
    var maxTop = getMaxTopOfTextFrames(doc);

    for (var i = 0; i < doc.pathItems.length; i++) {
      var p = doc.pathItems[i];
      if (p.name === "bg-white") continue;
      try { maxTop = Math.max(maxTop, p.geometricBounds[1]); } catch (e) {}
    }
    for (var j = 0; j < doc.placedItems.length; j++) {
      try { maxTop = Math.max(maxTop, doc.placedItems[j].geometricBounds[1]); } catch (e) {}
    }
    for (var k = 0; k < doc.groupItems.length; k++) {
      try { maxTop = Math.max(maxTop, doc.groupItems[k].geometricBounds[1]); } catch (e) {}
    }
    return maxTop;
  }

  function resizeArtboardHeight(doc, newH) {
    var ab = doc.artboards[0];
    ab.artboardRect = [0, newH, CFG.ARTBOARD_WIDTH, 0];
    for (var i = doc.pathItems.length - 1; i >= 0; i--) {
      var p = doc.pathItems[i];
      if (p.name === "bg-white") {
        try { p.locked = false; } catch (e1) {}
        try { p.remove(); }      catch (e2) {}
        break;
      }
    }
    addWhiteBackground(doc);
  }

  function estimateTitleBandHeight() {
    return (CFG.TITLE_FONT_SIZE * 1.35) + CFG.TITLE_BAND_PADDING;
  }

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
  // Drawn after height resize so it always anchors to the final canvas bottom.
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
    footer.left = (CFG.ARTBOARD_WIDTH - (gb[2] - gb[0])) / 2;
    footer.top  = CFG.FOOTER_BOTTOM_MARGIN;
  }

  // =========================
  // TEMPLATE TOKENS
  // {system} {region} {id} {name} {years} {years_paren}
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
  // =========================
  function getOutlinedLeft(tf) {
    var dup = null;
    var g   = null;
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
    title.textRange.characterAttributes.size     = CFG.TITLE_FONT_SIZE;
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
    sub.textRange.characterAttributes.size      = CFG.SUBTITLE_FONT_SIZE;
    sub.textRange.characterAttributes.textFont  = fontObj;
    sub.textRange.characterAttributes.fillColor = hexColor(CFG.SUBTITLE_COLOR);
    sub.textRange.paragraphAttributes.justification = Justification.LEFT;
    sub.left = CFG.SUBTITLE_LEFT;
    sub.top  = H - CFG.SUBTITLE_TOP_FROM_TOP;
    var target = (alignOutlineLeftToX !== null && alignOutlineLeftToX !== undefined)
      ? alignOutlineLeftToX : CFG.SUBTITLE_LEFT;
    snapTextOutlineLeft(sub, target);
    return sub;
  }

  // =========================
  // CONNECTION MARKER GEOMETRY
  //
  // Shape logic:
  //   If connDef.short_label exists AND is <= 2 chars → circle (diameter = CONNECTION_SIZE)
  //   Otherwise → auto-width pill (height = CONNECTION_SIZE)
  //
  // Asset override:
  //   If assets/connections/<id>.svg|png|jpg exists → place asset, skip drawn shape
  //
  // Returns the pixel height consumed by one marker (circle diameter or pill height).
  // =========================

  // Measure how tall the full connection stack will be for one station (without drawing)
  function measureConnectionBlockHeight(station, currentLineId, network) {
    if (!CFG.DRAW_CONNECTIONS || !network) return 0;
    var count = 0;
    for (var c = 0; c < station.connections.length; c++) {
      var cid = station.connections[c];
      if (cid === currentLineId) continue;
      if (!network[cid]) continue;
      count++;
    }
    if (count === 0) return 0;
    return count * MS.CONNECTION_SIZE + (count - 1) * MS.CONNECTION_Y_SPACING;
  }

  // Draw the connection markers for one station.
  // topY = Illustrator Y of the TOP of where the first marker should sit.
  function drawConnectionMarkersAt(doc, station, currentLineId, network, x, topY, fontObj, connIconFolder) {
    if (!CFG.DRAW_CONNECTIONS || !network) return;

    for (var c = 0; c < station.connections.length; c++) {
      var connId = station.connections[c];
      if (connId === currentLineId) continue;
      var connDef = network[connId];
      if (!connDef) continue;

      var markerTop = topY - (c * (MS.CONNECTION_SIZE + MS.CONNECTION_Y_SPACING));
      // markerTop is the Illustrator Y of the top edge of this marker

      // Asset override check
      var assetFile = findAsset(connIconFolder, connId);
      if (assetFile) {
        var cy = markerTop - (MS.CONNECTION_SIZE / 2);
        placeAsset(doc, assetFile, x, cy, MS.CONNECTION_SIZE);
        continue;
      }

      // Drawn marker — circle or pill
      var connColor  = connDef.color ? hexColor(connDef.color) : hexColor("#888888");
      var shortLabel = trim(connDef.short_label || "");
      var fullLabel  = trim(connDef.label || connId);
      var useCircle  = shortLabel.length > 0 && shortLabel.length <= 2;
      var displayLabel = useCircle ? shortLabel : fullLabel;

      var markerW, markerH, cornerR;
      if (useCircle) {
        markerW = MS.CONNECTION_SIZE;
        markerH = MS.CONNECTION_SIZE;
        cornerR = MS.CONNECTION_SIZE; // fully rounded = circle
      } else {
        var tmpTf = doc.textFrames.add();
        tmpTf.contents = displayLabel;
        tmpTf.textRange.characterAttributes.size     = CFG.CONNECTION_FONT_SIZE;
        tmpTf.textRange.characterAttributes.textFont = fontObj;
        var tmpGb = tmpTf.geometricBounds;
        var textW = tmpGb[2] - tmpGb[0];
        tmpTf.remove();
        markerW = Math.max(MS.CONNECTION_PILL_MIN_WIDTH, textW + 10);
        markerH = MS.CONNECTION_SIZE;
        cornerR = MS.CONNECTION_PILL_RADIUS;
      }

      var markerLeft = x - (markerW / 2);

      var shape = doc.pathItems.roundedRectangle(
        markerTop, markerLeft, markerW, markerH,
        cornerR * 2, cornerR * 2
      );
      shape.filled  = true;
      shape.stroked = false;
      shape.fillColor = connColor;

      // Label inside marker
      var lbl = doc.textFrames.add();
      lbl.contents = displayLabel;
      lbl.textRange.characterAttributes.size     = CFG.CONNECTION_FONT_SIZE;
      lbl.textRange.characterAttributes.textFont = fontObj;
      lbl.textRange.paragraphAttributes.justification = Justification.CENTER;
      var white = new RGBColor();
      white.red = 255; white.green = 255; white.blue = 255;
      lbl.textRange.characterAttributes.fillColor = white;

      var lblGb = lbl.geometricBounds;
      var lblH  = lblGb[1] - lblGb[3];
      var lblW  = lblGb[2] - lblGb[0];
      lbl.left = x - (lblW / 2);
      lbl.top  = markerTop - (markerH / 2) + (lblH / 2);
    }
  }

  // =========================
  // ICON GRID
  // Icons are laid out in rows of up to ICON_MAX_PER_ROW.
  // Each row is centered on x.
  // The entire grid starts at topY (Illustrator Y of the top of the first row).
  // Asset source: assets/icons/<key>.svg|png|jpg
  // Missing assets are warned once per run and skipped.
  // All sizes use MS.* (scaled by MARKER_SCALE).
  // =========================
  function measureIconGridHeight(icons) {
    if (!CFG.DRAW_ICONS || !icons || !icons.length) return 0;
    var rows = Math.ceil(icons.length / CFG.ICON_MAX_PER_ROW);
    return rows * MS.ICON_SIZE + (rows - 1) * MS.ICON_Y_SPACING;
  }

  function drawIconGridAt(doc, icons, x, topY, iconFolder, missingIconLog) {
    if (!CFG.DRAW_ICONS || !icons || !icons.length) return;
    if (!iconFolder || !iconFolder.exists) return;

    for (var i = 0; i < icons.length; i++) {
      var key = trim(icons[i]);
      var row = Math.floor(i / CFG.ICON_MAX_PER_ROW);
      var col = i % CFG.ICON_MAX_PER_ROW;

      var rowCount  = Math.min(CFG.ICON_MAX_PER_ROW, icons.length - row * CFG.ICON_MAX_PER_ROW);
      var rowWidth  = rowCount * MS.ICON_SIZE + (rowCount - 1) * MS.ICON_X_GAP;
      var rowStartX = x - (rowWidth / 2);

      var iconCX = rowStartX + col * (MS.ICON_SIZE + MS.ICON_X_GAP) + (MS.ICON_SIZE / 2);
      var iconCY = topY - (row * (MS.ICON_SIZE + MS.ICON_Y_SPACING)) - (MS.ICON_SIZE / 2);

      var assetFile = findAsset(iconFolder, key);
      if (!assetFile) {
        missingIconLog[key] = true;
        continue;
      }

      placeAsset(doc, assetFile, iconCX, iconCY, MS.ICON_SIZE);
    }
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
    opts.embedRasterImages    = true;
    opts.fontSubsetting       = SVGFontSubsetting.GLYPHSUSED;
    opts.coordinatePrecision  = 2;
    opts.cssProperties        = SVGCSSPropertyLocation.PRESENTATIONATTRIBUTES;
    doc.exportFile(destFile, ExportType.SVG, opts);
  }

  // =========================
  // LABEL OVERHANG / PADDING
  // =========================
  function measureLabelOverhang(doc, labelText, anchorX, anchorY, rotation, fontObj) {
    var tf = doc.textFrames.add();
    tf.contents = labelText;
    tf.textRange.characterAttributes.size     = CFG.FONT_SIZE;
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
    var extraLeft = 0, extraRight = 0;

    var lastLabel = stations[stations.length - 1].label;
    var mLast = measureLabelOverhang(doc, lastLabel, lineRight, baselineY, rotation, fontObj);
    extraRight = Math.min(CFG.PAD_MAX_EXTRA, Math.max(0, (mLast.rightOverhang + CFG.PAD_RIGHT_MARGIN) - basePad));

    if (CFG.PAD_CHECK_LEFT) {
      var firstLabel = stations[0].label;
      var mFirst = measureLabelOverhang(doc, firstLabel, lineLeft, baselineY, rotation, fontObj);
      extraLeft = Math.min(CFG.PAD_MAX_EXTRA, Math.max(0, (mFirst.leftOverhang + CFG.PAD_LEFT_MARGIN) - basePad));
    }

    return { extraLeft: extraLeft, extraRight: extraRight };
  }

  // =========================
  // STATION REGISTRY LOADING
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

    var lineId      = trim(obj.id        || "");
    var lineName    = trim(obj.line_name || "");
    var years       = trim(obj.years     || "");
    var color       = trim(obj.color     || "");
    var rawStations = obj.stations;

    if (!lineName) throw new Error("Missing required field: line_name (" + file.name + ")");
    if (!color)    throw new Error("Missing required field: color ("     + file.name + ")");
    color = normalizeHex(color);

    var system  = registry ? registry.system  : trim(obj.system  || "");
    var region  = registry ? registry.region  : trim(obj.region  || "");
    var network = registry ? registry.network : {};

    if (!system) throw new Error("Missing system name. Define it in stations.json or the line file (" + file.name + ")");
    if (!region) throw new Error("Missing region name. Define it in stations.json or the line file (" + file.name + ")");

    var stationIds = [];
    if (rawStations instanceof Array) {
      for (var i = 0; i < rawStations.length; i++) {
        var entry = rawStations[i];
        stationIds.push(typeof entry === "string" ? trim(entry) : null);
      }
    } else if (typeof rawStations === "string") {
      var parts = rawStations.split("|");
      for (var j = 0; j < parts.length; j++) stationIds.push(trim(parts[j]));
    } else {
      throw new Error("stations must be an array or pipe-delimited string (" + file.name + ")");
    }

    var cleaned = [];
    for (var k = 0; k < stationIds.length; k++) if (stationIds[k]) cleaned.push(stationIds[k]);
    if (cleaned.length < 2) throw new Error("stations must contain at least 2 entries (" + file.name + ")");

    var resolvedStations = [];
    for (var s = 0; s < cleaned.length; s++) {
      var sid = cleaned[s];
      if (registry && registry.stations) {
        var staDef = registry.stations[sid];
        if (!staDef) throw new Error("Station ID \"" + sid + "\" not found in stations.json (" + file.name + ")");
        resolvedStations.push({
          id:          sid,
          label:       trim(staDef.label  || sid),
          label2:      trim(staDef.label2 || ""),
          connections: staDef.connections instanceof Array ? staDef.connections : [],
          icons:       staDef.icons       instanceof Array ? staDef.icons       : []
        });
      } else {
        resolvedStations.push({ id: sid, label: sid, label2: "", connections: [], icons: [] });
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
  // =========================
  var scriptFolder  = getScriptFolder();
  var systemFolders = [];

  // Asset folders — next to linear.jsx
  var connIconFolder = null;
  var iconFolder     = null;
  if (scriptFolder) {
    var assetsFolder = Folder(scriptFolder.fsName + "/assets");
    if (assetsFolder.exists) {
      var ci = Folder(assetsFolder.fsName + "/connections");
      var ii = Folder(assetsFolder.fsName + "/icons");
      if (ci.exists) connIconFolder = ci;
      if (ii.exists) iconFolder     = ii;
    }
  }

  if (scriptFolder) {
    var directLines = Folder(scriptFolder.fsName + "/lines");
    if (directLines.exists) {
      var subFolders = directLines.getFiles(function (f) { return f instanceof Folder; });
      if (subFolders.length > 0) {
        for (var sf = 0; sf < subFolders.length; sf++) systemFolders.push(subFolders[sf]);
      } else {
        systemFolders.push(directLines);
      }
    }
  }

  if (!systemFolders.length) {
    var manualFolder = Folder.selectDialog("Select folder containing line JSON files");
    if (!manualFolder) { alert("No folder selected."); return; }
    systemFolders.push(manualFolder);
  }

  // =========================
  // LOAD ALL LINES
  // =========================
  var LINES = [];

  for (var sfi = 0; sfi < systemFolders.length; sfi++) {
    var sysFolder = systemFolders[sfi];
    var registry  = null;
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

  if (!LINES.length) { alert("No line files found."); return; }

  // =========================
  // EXPORT FOLDER RESOLUTION
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
  // FONTS — resolved once
  // =========================
  var labelFontObj    = getFontOrArial(CFG.FONT_LABEL_NAME);
  var titleFontObj    = getFontOrArial(CFG.FONT_TITLE_NAME);
  var subtitleFontObj = getFontOrArial(CFG.FONT_SUBTITLE_NAME);
  var footerFontObj   = getFontOrArial(CFG.FONT_FOOTER_NAME);
  var connFontObj     = getFontOrArial(CFG.FONT_LABEL_NAME);

  var stationStroke = hexColor(CFG.STATION_OUTLINE);
  var labelRotation = getLabelRotation();

  // Accumulate missing icon warnings across all lines
  var missingIconLog = {};

  // =========================
  // GENERATE DOCUMENTS
  // =========================
  for (var li = 0; li < LINES.length; li++) {
    var L = LINES[li];

    var doc = app.documents.add(DocumentColorSpace.RGB, CFG.ARTBOARD_WIDTH, CFG.ARTBOARD_HEIGHT);
    doc.rulerUnits = RulerUnits.Pixels;
    doc.artboards[0].artboardRect = [0, CFG.ARTBOARD_HEIGHT, CFG.ARTBOARD_WIDTH, 0];
    addWhiteBackground(doc);

    var baselineY = CFG.BASELINE_Y;

    // Terminal padding
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

    // Optional line outline
    if (CFG.LINE_OUTLINE_ENABLED) {
      var lineOutline = doc.pathItems.add();
      lineOutline.stroked     = true;
      lineOutline.filled      = false;
      lineOutline.strokeWidth = CFG.LINE_STROKE + (CFG.LINE_OUTLINE_WIDTH * 2);
      lineOutline.strokeCap   = StrokeCap.ROUNDENDCAP;
      lineOutline.strokeColor = hexColor(CFG.LINE_OUTLINE_COLOR);
      lineOutline.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);
    }

    // Main line stroke
    var lineItem = doc.pathItems.add();
    lineItem.stroked     = true;
    lineItem.filled      = false;
    lineItem.strokeWidth = CFG.LINE_STROKE;
    lineItem.strokeCap   = StrokeCap.ROUNDENDCAP;
    lineItem.strokeColor = hexColor(L.color);
    lineItem.setEntirePath([[lineLeft, baselineY], [lineRight, baselineY]]);

    var white = new RGBColor();
    white.red = 255; white.green = 255; white.blue = 255;

    var lineTopY   = baselineY + (CFG.LINE_STROKE / 2);
    var dotBottomY = baselineY - CFG.STATION_RADIUS; // Illustrator Y at bottom of station dot

    // -------------------------------------------------------
    // PASS 1 — measure max connection block height across all
    // stations so we can set a shared marker/icon baseline
    // -------------------------------------------------------
    var maxConnBlockH = 0;
    for (var mi = 0; mi < L.stations.length; mi++) {
      var mh = measureConnectionBlockHeight(L.stations[mi], L.id, L.network);
      if (mh > maxConnBlockH) maxConnBlockH = mh;
    }

    // connBlockTopY: shared Illustrator Y of the TOP of the first connection marker.
    // All stations use this same Y so markers sit on a consistent horizontal band.
    var connBlockTopY    = dotBottomY - MS.CONNECTION_Y_GAP;
    var connBlockBottomY = connBlockTopY - maxConnBlockH;

    // iconTopY: shared Illustrator Y of the TOP of the first icon row.
    // If no station on this line has connections, icons sit directly below
    // the dot using CONNECTION_Y_GAP as the only offset (no double gap).
    var iconTopY = (maxConnBlockH > 0)
      ? connBlockBottomY - MS.ICON_Y_GAP
      : dotBottomY - MS.CONNECTION_Y_GAP;

    // -------------------------------------------------------
    // PASS 2 — draw stations, labels, connections, icons
    // -------------------------------------------------------
    for (var i = 0; i < L.stations.length; i++) {
      var station = L.stations[i];
      var x = lineLeft + i * spacing;

      // Station dot
      var stationStrokeWidth = (CFG.STATION_STROKE_WIDTH !== null && CFG.STATION_STROKE_WIDTH !== undefined)
        ? CFG.STATION_STROKE_WIDTH
        : Math.max(2, Math.round(CFG.LINE_STROKE * 0.35));

      var dot = doc.pathItems.ellipse(
        baselineY + CFG.STATION_RADIUS,
        x - CFG.STATION_RADIUS,
        CFG.STATION_RADIUS * 2,
        CFG.STATION_RADIUS * 2
      );
      dot.filled      = true;
      dot.stroked     = true;
      dot.fillColor   = white;
      dot.strokeColor = stationStroke;
      dot.strokeWidth = stationStrokeWidth;

      // Station label (label2 → second line via native \r line break)
      var labelText = station.label;
      if (station.label2) labelText = labelText + "\r" + station.label2;

      var label = doc.textFrames.add();
      label.contents = labelText;
      label.textRange.characterAttributes.size     = CFG.FONT_SIZE;
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

      // Connection markers — drawn from connBlockTopY downward
      drawConnectionMarkersAt(doc, station, L.id, L.network, x, connBlockTopY, connFontObj, connIconFolder);

      // Icon grid — drawn at shared iconTopY regardless of whether this station has connections
      drawIconGridAt(doc, station.icons, x, iconTopY, iconFolder, missingIconLog);
    }

    // Dynamic height
    // neededH accounts for:
    //   - tallest drawn content (labels, markers, icons)
    //   - title band reserve (if enabled)
    //   - MARKER_FOOTER_GAP between content bottom and footer top
    //   - footer text height
    //   - FOOTER_BOTTOM_MARGIN below footer
    if (CFG.AUTO_HEIGHT) {
      var maxTop = getMaxTopOfAllItems(doc);

      var titleReserve = 0;
      if (CFG.DRAW_TITLE && CFG.RESERVE_TITLE_BAND) {
        titleReserve = estimateTitleBandHeight() + CFG.TITLE_TOP_FROM_TOP + CFG.TITLE_BAND_LINE_GAP;
      }

      // Estimate footer height from font size (actual frame drawn after resize)
      var footerH = CFG.FOOTER_TEXT ? (CFG.FOOTER_FONT_SIZE * 1.2) : 0;

      var neededH = Math.ceil(Math.max(
        CFG.MIN_HEIGHT,
        maxTop + titleReserve + CFG.MARKER_FOOTER_GAP + footerH + CFG.FOOTER_BOTTOM_MARGIN
      ));
      resizeArtboardHeight(doc, neededH);
    }

    // Title + subtitle after resize
    var titleInfo        = drawTitle(doc, L, titleFontObj);
    var titleOutlineLeft = titleInfo ? titleInfo.outlineLeft : null;
    drawSubtitle(doc, L, subtitleFontObj, titleOutlineLeft);

    // Footer last — anchors to final canvas bottom
    drawFooter(doc, CFG.FOOTER_TEXT, footerFontObj);

    // SVG export
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

      var baseName = safeFileName(applyTemplate(CFG.EXPORT_FILENAME_TEMPLATE, L));
      if (!baseName) baseName = safeFileName(L.name || (L.system + "-" + L.id));

      var svgFile = File(targetFolder.fsName + "/" + baseName + ".svg");
      try {
        if (CFG.OUTLINE_TEXT_FOR_SVG) outlineAllText(doc);
        exportAsSVG(doc, svgFile);
      } catch (eSvg) {
        alert("SVG export failed for:\n" + baseName + "\n\n" + eSvg.message);
      }

      if (CFG.CLOSE_AFTER_EXPORT) doc.close(SaveOptions.DONOTSAVECHANGES);
    }
  }

  // Warn about any missing icon assets (once, after all lines processed)
  var missingKeys = [];
  for (var mk in missingIconLog) {
    if (missingIconLog.hasOwnProperty(mk)) missingKeys.push(mk);
  }
  if (missingKeys.length) {
    alert(
      missingKeys.length + " icon asset(s) were referenced but not found in assets/icons/:\n\n" +
      missingKeys.join("\n") +
      "\n\nExpected: assets/icons/<key>.svg  (or .png / .jpg)"
    );
  }

  var exportMsg = "";
  if (CFG.EXPORT_SVG) {
    exportMsg = "\n\nExported SVGs to:\n" + exportBaseFolder.fsName +
      (CFG.EXPORT_GROUP_BY_SYSTEM ? "\n(Grouped by system)" : "");
  }
  alert("Generated " + LINES.length + " diagram(s).\n\nSource: " + systemFolders[0].parent.fsName + exportMsg);
})();