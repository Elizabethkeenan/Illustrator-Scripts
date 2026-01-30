/*
CATZper-Lite CMYK Chart Generator (Illustrator JSX) - Swatch/Selection Center Color Update
NEW:
- Can pull center color from:
  (1) Selected object's fill color (preferred)
  (2) A named swatch in the Swatches panel
  (3) Manual CMYK entry fallback

Notes:
- Illustrator scripting can't reliably read "currently highlighted swatch" in the Swatches panel.
- Best workflow: keep a small square filled with the swatch you want, select it, run script.
*/

(function () {
  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;

  // ---- Utilities ----
  function clamp(v, min, max) {
    return Math.max(min, Math.min(max, v));
  }

  function toInt(s, fallback) {
    var n = parseInt(s, 10);
    return isNaN(n) ? fallback : n;
  }

  function toNum(s, fallback) {
    var n = parseFloat(s);
    return isNaN(n) ? fallback : n;
  }

  function makeCMYK(c, m, y, k) {
    var col = new CMYKColor();
    col.cyan = clamp(c, 0, 100);
    col.magenta = clamp(m, 0, 100);
    col.yellow = clamp(y, 0, 100);
    col.black = clamp(k, 0, 100);
    return col;
  }

  function colorToCMYK(aiColor) {
    // Returns {c,m,y,k} or null if cannot convert.
    // Handles: CMYKColor, SpotColor (with CMYK/RGB/Lab base), RGBColor, GrayColor.
    if (!aiColor || !aiColor.typename) return null;

    try {
      if (aiColor.typename === "CMYKColor") {
        return {
          c: aiColor.cyan,
          m: aiColor.magenta,
          y: aiColor.yellow,
          k: aiColor.black
        };
      }

      if (aiColor.typename === "GrayColor") {
        // Gray in Illustrator is 0..100 where 100=black
        // Convert to CMYK approx: K = gray, others 0
        return { c: 0, m: 0, y: 0, k: aiColor.gray };
      }

      if (aiColor.typename === "RGBColor") {
        // Convert RGB -> CMYK using Illustrator's convertSampleColor if available
        // Fallback: rough conversion
        if (app.convertSampleColor) {
          var out = app.convertSampleColor(ImageColorSpace.RGB, [aiColor.red, aiColor.green, aiColor.blue],
                                           ImageColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
          // out is [C,M,Y,K] in 0..100
          return { c: out[0], m: out[1], y: out[2], k: out[3] };
        } else {
          // Rough RGB->CMYK fallback
          var r = aiColor.red / 255, g = aiColor.green / 255, b = aiColor.blue / 255;
          var k = 1 - Math.max(r, g, b);
          if (k >= 1) return { c: 0, m: 0, y: 0, k: 100 };
          var c = (1 - r - k) / (1 - k);
          var m = (1 - g - k) / (1 - k);
          var y = (1 - b - k) / (1 - k);
          return { c: c * 100, m: m * 100, y: y * 100, k: k * 100 };
        }
      }

      if (aiColor.typename === "SpotColor") {
        // SpotColor has .spot.color which can be CMYK/RGB/Lab/Gray depending on swatch definition
        var base = aiColor.spot.color;
        return colorToCMYK(base);
      }

      if (aiColor.typename === "LabColor") {
        // Convert Lab -> CMYK if possible
        if (app.convertSampleColor) {
          var outLab = app.convertSampleColor(ImageColorSpace.LAB,
                                              [aiColor.l, aiColor.a, aiColor.b],
                                              ImageColorSpace.CMYK,
                                              ColorConvertPurpose.defaultpurpose);
          return { c: outLab[0], m: outLab[1], y: outLab[2], k: outLab[3] };
        }
        return null;
      }
    } catch (e) {
      return null;
    }

    return null;
  }

  function getFillColorFromSelection() {
    // Try to read a reasonable "fillColor" from selection.
    // Returns Illustrator Color object or null.
    if (!doc.selection || doc.selection.length === 0) return null;
    if (doc.selection.length !== 1) return null;

    var it = doc.selection[0];
    try {
      // PathItem
      if (it.typename === "PathItem") {
        if (it.filled) return it.fillColor;
        return null;
      }
      // CompoundPathItem
      if (it.typename === "CompoundPathItem") {
        if (it.pathItems.length > 0 && it.pathItems[0].filled) return it.pathItems[0].fillColor;
        return null;
      }
      // TextFrame
      if (it.typename === "TextFrame") {
        var tr = it.textRange;
        // If there's mixed formatting, this is imperfect; we just take first character attributes.
        return tr.characterAttributes.fillColor;
      }
      // GroupItem: try first filled path inside
      if (it.typename === "GroupItem") {
        for (var i = 0; i < it.pageItems.length; i++) {
          var child = it.pageItems[i];
          if (child.typename === "PathItem" && child.filled) return child.fillColor;
          if (child.typename === "CompoundPathItem" && child.pathItems.length > 0 && child.pathItems[0].filled) {
            return child.pathItems[0].fillColor;
          }
        }
        return null;
      }
    } catch (e) {
      return null;
    }

    return null;
  }

  function getColorFromSwatchName(name) {
    // Returns Illustrator Color object or null
    if (!name) return null;
    try {
      var sw = doc.swatches.getByName(name);
      // sw.color is a Color object
      return sw.color;
    } catch (e) {
      return null;
    }
  }

  // ---- UI ----
  var w = new Window("dialog", "CMYK Tolerance Chart (CATZper-Lite)");
  w.alignChildren = "fill";

  var p1 = w.add("panel", undefined, "Grid");
  p1.orientation = "column";
  p1.alignChildren = "left";

  var row1 = p1.add("group");
  row1.add("statictext", undefined, "Columns (odd):");
  var edtCols = row1.add("edittext", undefined, "7");
  edtCols.characters = 4;

  var row2 = p1.add("group");
  row2.add("statictext", undefined, "Rows (odd):");
  var edtRows = row2.add("edittext", undefined, "7");
  edtRows.characters = 4;

  var row3 = p1.add("group");
  row3.add("statictext", undefined, "Patch size (in):");
  var edtPatch = row3.add("edittext", undefined, "1.25");
  edtPatch.characters = 6;

  var row4 = p1.add("group");
  row4.add("statictext", undefined, "Gutter (in):");
  var edtGutter = row4.add("edittext", undefined, "0.15");
  edtGutter.characters = 6;

  var pCenter = w.add("panel", undefined, "Center Color Source");
  pCenter.orientation = "column";
  pCenter.alignChildren = "left";

  var chkUseSel = pCenter.add("checkbox", undefined, "Use selected object's fill color as center (recommended)");
  chkUseSel.value = true;

  var gSw = pCenter.add("group");
  gSw.add("statictext", undefined, "Or swatch name:");
  var edtSwatch = gSw.add("edittext", undefined, "");
  edtSwatch.characters = 24;

  var p2 = w.add("panel", undefined, "Manual Center Color (fallback) CMYK 0–100");
  p2.orientation = "column";
  p2.alignChildren = "left";

  function cmykRow(label, def) {
    var g = p2.add("group");
    g.add("statictext", undefined, label);
    var e = g.add("edittext", undefined, def);
    e.characters = 4;
    return e;
  }

  var edtC = cmykRow("C:", "30");
  var edtM = cmykRow("M:", "20");
  var edtY = cmykRow("Y:", "20");
  var edtK = cmykRow("K:", "10");

  var p3 = w.add("panel", undefined, "Steps per Cell (X axis +, Y axis +)");
  p3.orientation = "column";
  p3.alignChildren = "left";

  function stepRow(label, defX, defY) {
    var g = p3.add("group");
    g.add("statictext", undefined, label);
    g.add("statictext", undefined, "X:");
    var ex = g.add("edittext", undefined, defX);
    ex.characters = 5;
    g.add("statictext", undefined, "Y:");
    var ey = g.add("edittext", undefined, defY);
    ey.characters = 5;
    return { x: ex, y: ey };
  }

  var stC = stepRow("ΔC per cell", "2", "0");
  var stM = stepRow("ΔM per cell", "0", "0");
  var stY = stepRow("ΔY per cell", "0", "2");
  var stK = stepRow("ΔK per cell", "0", "3"); // default includes a darkening axis now

  var p4 = w.add("panel", undefined, "Labels");
  p4.orientation = "column";
  p4.alignChildren = "left";

  var gLbl = p4.add("group");
  var chkLabels = gLbl.add("checkbox", undefined, "Add labels under each patch");
  chkLabels.value = true;

  var gFont = p4.add("group");
  gFont.add("statictext", undefined, "Font size (pt):");
  var edtFont = gFont.add("edittext", undefined, "8");
  edtFont.characters = 4;

  var gTitle = p4.add("group");
  gTitle.add("statictext", undefined, "Chart title:");
  var edtTitle = gTitle.add("edittext", undefined, "GRAY_MATCH_01");
  edtTitle.characters = 20;

  var btns = w.add("group");
  btns.alignment = "right";
  var btnCancel = btns.add("button", undefined, "Cancel");
  var btnOK = btns.add("button", undefined, "Generate", { name: "ok" });

  btnCancel.onClick = function () { w.close(0); };

  btnOK.onClick = function () {
    var cols = toInt(edtCols.text, 7);
    var rows = toInt(edtRows.text, 7);
    if (cols < 1 || rows < 1) { alert("Rows/Columns must be >= 1."); return; }
    if (cols % 2 === 0 || rows % 2 === 0) { alert("Use odd numbers for rows/columns so there is a center patch."); return; }

    var patchIn = toNum(edtPatch.text, 1.25);
    var gutterIn = toNum(edtGutter.text, 0.15);
    if (patchIn <= 0) { alert("Patch size must be > 0."); return; }
    if (gutterIn < 0) { alert("Gutter must be >= 0."); return; }

    // Steps
    var dCx = toNum(stC.x.text, 0), dCy = toNum(stC.y.text, 0);
    var dMx = toNum(stM.x.text, 0), dMy = toNum(stM.y.text, 0);
    var dYx = toNum(stY.x.text, 0), dYy = toNum(stY.y.text, 0);
    var dKx = toNum(stK.x.text, 0), dKy = toNum(stK.y.text, 0);

    var fontSize = toNum(edtFont.text, 8);
    if (fontSize <= 0) fontSize = 8;

    var title = (edtTitle.text || "").replace(/^\s+|\s+$/g, "");
    if (!title) title = "CHART";

    // Determine base color
    var base = null;

    if (chkUseSel.value) {
      var selColor = getFillColorFromSelection();
      if (selColor) {
        base = colorToCMYK(selColor);
        if (!base) {
          alert("Selected object's fill color could not be converted to CMYK. Try a CMYK swatch/object, or use swatch name/manual entry.");
          return;
        }
      }
    }

    if (!base) {
      var swName = (edtSwatch.text || "").replace(/^\s+|\s+$/g, "");
      if (swName) {
        var swColor = getColorFromSwatchName(swName);
        if (!swColor) {
          alert("Swatch not found: \"" + swName + "\"\nCheck spelling exactly, or use selected object/manual entry.");
          return;
        }
        base = colorToCMYK(swColor);
        if (!base) {
          alert("Swatch color could not be converted to CMYK. Try a CMYK/Spot(CMYK) swatch, or use manual entry.");
          return;
        }
      }
    }

    if (!base) {
      // Manual fallback
      base = {
        c: toNum(edtC.text, 0),
        m: toNum(edtM.text, 0),
        y: toNum(edtY.text, 0),
        k: toNum(edtK.text, 0)
      };
    }

    w.close(1);

    app.executeMenuCommand("deselectall");

    var IN_TO_PT = 72;
    var patch = patchIn * IN_TO_PT;
    var gutter = gutterIn * IN_TO_PT;

    // Create group
    var layer = doc.activeLayer;
    var grp = layer.groupItems.add();
    grp.name = "CATZperLite_" + title;

    // Place on active artboard
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()];
    var r = ab.artboardRect; // [left, top, right, bottom]
    var left = r[0], top = r[1];

    var margin = 36; // 0.5"
    var startX = left + margin;
    var startY = top - margin;

    // Title
    var titleTf = grp.textFrames.add();
    titleTf.contents =
      title +
      "  |  Center CMYK(" +
      Math.round(base.c) + "," + Math.round(base.m) + "," + Math.round(base.y) + "," + Math.round(base.k) +
      ")  |  Steps X(" +
      dCx + "," + dMx + "," + dYx + "," + dKx +
      ") Y(" +
      dCy + "," + dMy + "," + dYy + "," + dKy + ")";
    titleTf.textRange.characterAttributes.size = fontSize + 2;
    titleTf.left = startX;
    titleTf.top = startY;
    titleTf.textRange.characterAttributes.fillColor = makeCMYK(0, 0, 0, 100);

    var chartTop = startY - (fontSize + 10);

    var cx = Math.floor(cols / 2);
    var cy = Math.floor(rows / 2);

    function setStroke(pi, on) {
      if (on) {
        pi.stroked = true;
        pi.strokeWidth = 0.5;
        pi.strokeColor = makeCMYK(0, 0, 0, 30);
      } else {
        pi.stroked = false;
      }
    }

    for (var row = 0; row < rows; row++) {
      for (var col = 0; col < cols; col++) {
        var ox = col - cx;
        var oy = row - cy;

        var c = base.c + ox * dCx + oy * dCy;
        var m = base.m + ox * dMx + oy * dMy;
        var yv = base.y + ox * dYx + oy * dYy;
        var k = base.k + ox * dKx + oy * dKy;

        c = clamp(c, 0, 100);
        m = clamp(m, 0, 100);
        yv = clamp(yv, 0, 100);
        k = clamp(k, 0, 100);

        var x = startX + col * (patch + gutter);
        var yTop = chartTop - row * (patch + gutter);

        var rect = grp.pathItems.rectangle(yTop, x, patch, patch);
        rect.filled = true;
        rect.fillColor = makeCMYK(c, m, yv, k);
        setStroke(rect, true);

        if (ox === 0 && oy === 0) {
          rect.strokeWidth = 2;
          rect.strokeColor = makeCMYK(0, 0, 0, 70);
        }

        if (chkLabels.value) {
          var lbl = grp.textFrames.add();
          lbl.textRange.characterAttributes.size = fontSize;
          lbl.textRange.characterAttributes.fillColor = makeCMYK(0, 0, 0, 100);

          var s =
            "ox " + ox + "  oy " + oy + "\n" +
            "C" + Math.round(c) + " M" + Math.round(m) + " Y" + Math.round(yv) + " K" + Math.round(k);

          lbl.contents = s;
          lbl.left = x;
          lbl.top = yTop - patch - 4;
        }
      }
    }

    var note = grp.textFrames.add();
    note.contents = "Tip: Circle best patch in DAYLIGHT and in SHOP lighting. For vehicles, daylight usually wins. If you need darker, increase ΔK on Y axis.";
    note.textRange.characterAttributes.size = fontSize;
    note.textRange.characterAttributes.fillColor = makeCMYK(0, 0, 0, 70);
    note.left = startX;
    note.top = chartTop - rows * (patch + gutter) - 20;

    alert("Chart generated: " + grp.name + "\nCenter pulled from: " +
      (chkUseSel.value && getFillColorFromSelection() ? "Selection" : (edtSwatch.text ? "Swatch name" : "Manual")));
  };

  w.show();
})();
