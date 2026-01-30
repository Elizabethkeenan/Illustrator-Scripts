/************************************************************
 * Swatch Grid from Palette (Illustrator-safe)
 * Creates a labeled grid of color swatches in the active doc
 * 
 * - Uses all solid color swatches (process + spot)
 * - Ignores [None], [Registration], patterns, gradients
 * - Units: in, mm, cm, pt, px (e.g. "0.75 in", "5 mm", "10 pt")
 ************************************************************/

if (app.documents.length === 0) {
    alert("Open a document with a swatch palette first.");
} else {
    main();
}

function main() {
    var doc = app.activeDocument;

    var usableSwatches = getUsableSwatches(doc);
    if (usableSwatches.length === 0) {
        alert("No solid color swatches found.\n(Skips [None], [Registration], patterns, gradients.)");
        return;
    }

    var options = showDialog(usableSwatches.length);
    if (!options) return; // user cancelled

    // Illustrator already treats the whole script as one undo step
    createSwatchGrid(doc, usableSwatches, options);
}

/**
 * Collect solid-color swatches: CMYK, RGB, Gray, Lab, Spot.
 * Skip special swatches and pattern/gradient types.
 */
function getUsableSwatches(doc) {
    var list = [];
    var swatches = doc.swatches;
    for (var i = 0; i < swatches.length; i++) {
        var sw = swatches[i];

        // Skip special ones like [None], [Registration], etc.
        if (sw.name && sw.name.charAt(0) === "[") continue;

        if (!sw.color) continue;

        var t = sw.color.typename;
        if (t === "CMYKColor" ||
            t === "RGBColor" ||
            t === "GrayColor" ||
            t === "LabColor" ||
            t === "SpotColor") {
            list.push(sw);
        }
        // patterns & gradients are ignored by design
    }
    return list;
}

/**
 * Simple unit parser: returns points (Illustrator's internal).
 * Accepts:
 *   "0.75 in", "5 mm", "1.5 cm", "12 pt", "12 px", "10"
 * If no unit, treats as points.
 */
function toPoints(value) {
    if (typeof value === "number") return value;

    var s = String(value).replace(/^\s+|\s+$/g, "").toLowerCase();
    if (!s.length) return 0;

    var unit = "";
    var num  = parseFloat(s);

    if (isNaN(num)) return 0;

    if (s.match(/in$/)) {
        unit = "in";
    } else if (s.match(/mm$/)) {
        unit = "mm";
    } else if (s.match(/cm$/)) {
        unit = "cm";
    } else if (s.match(/pt$/)) {
        unit = "pt";
    } else if (s.match(/px$/)) {
        unit = "px";
    } else {
        // no unit → treat as points
        unit = "pt";
    }

    switch (unit) {
        case "in":
            return num * 72;
        case "mm":
            return num * 72 / 25.4;
        case "cm":
            return num * 72 / 2.54;
        case "pt":
        case "px":
        default:
            return num;
    }
}

/**
 * Dialog for basic options.
 */
function showDialog(swatchCount) {
    var win = new Window("dialog", "Swatch Grid from Palette");

    win.orientation = "column";
    win.alignChildren = ["fill", "top"];

    var infoGroup = win.add("group");
    infoGroup.add("statictext", undefined, "Swatches found: " + swatchCount);

    // Size + columns
    var sizePanel = win.add("panel", undefined, "Chip & Grid");
    sizePanel.orientation = "column";
    sizePanel.alignChildren = ["fill", "top"];
    sizePanel.margins = 12;

    var g1 = sizePanel.add("group");
    g1.add("statictext", undefined, "Chip width:");
    var chipW = g1.add("edittext", undefined, "0.75 in");
    chipW.preferredSize.width = 80;

    var g2 = sizePanel.add("group");
    g2.add("statictext", undefined, "Chip height:");
    var chipH = g2.add("edittext", undefined, "0.5 in");
    chipH.preferredSize.width = 80;

    var g3 = sizePanel.add("group");
    g3.add("statictext", undefined, "Columns:");
    var cols = g3.add("edittext", undefined, "10");
    cols.preferredSize.width = 50;

    var gapRow = sizePanel.add("group");
    gapRow.add("statictext", undefined, "Horizontal gap:");
    var gapX = gapRow.add("edittext", undefined, "0.125 in");
    gapX.preferredSize.width = 80;
    gapRow.add("statictext", undefined, "Vertical gap:");
    var gapY = gapRow.add("edittext", undefined, "0.125 in");
    gapY.preferredSize.width = 80;

    // Margins
    var marginPanel = win.add("panel", undefined, "Margins from artboard edge");
    marginPanel.orientation = "column";
    marginPanel.alignChildren = ["fill", "top"];
    marginPanel.margins = 12;

    var mRow1 = marginPanel.add("group");
    mRow1.add("statictext", undefined, "Top:");
    var mTop = mRow1.add("edittext", undefined, "0.5 in");
    mTop.preferredSize.width = 80;
    mRow1.add("statictext", undefined, "Left:");
    var mLeft = mRow1.add("edittext", undefined, "0.5 in");
    mLeft.preferredSize.width = 80;

    // Labels
    var labelPanel = win.add("panel", undefined, "Labels");
    labelPanel.orientation = "column";
    labelPanel.alignChildren = ["fill", "top"];
    labelPanel.margins = 12;

    var l1 = labelPanel.add("group");
    l1.add("statictext", undefined, "Label font size (pt):");
    var labelSize = l1.add("edittext", undefined, "8");
    labelSize.preferredSize.width = 50;

    var l2 = labelPanel.add("group");
    l2.add("statictext", undefined, "Label gap below chip:");
    var labelGap = l2.add("edittext", undefined, "0.0625 in");
    labelGap.preferredSize.width = 80;

    // Buttons
    var btns = win.add("group");
    btns.alignment = "right";
    var cancel = btns.add("button", undefined, "Cancel");
    var ok     = btns.add("button", undefined, "OK");

    cancel.onClick = function () {
        win.close(0);
    };
    ok.onClick = function () {
        if (isNaN(parseInt(cols.text)) || parseInt(cols.text) < 1) {
            alert("Columns must be a positive integer.");
            return;
        }
        if (isNaN(parseFloat(labelSize.text)) || parseFloat(labelSize.text) <= 0) {
            alert("Label size must be a positive number.");
            return;
        }
        win.close(1);
    };

    var result = win.show();
    if (result !== 1) return null;

    return {
        chipWidth:  toPoints(chipW.text),
        chipHeight: toPoints(chipH.text),
        columns:    parseInt(cols.text),
        gapX:       toPoints(gapX.text),
        gapY:       toPoints(gapY.text),
        marginTop:  toPoints(mTop.text),
        marginLeft: toPoints(mLeft.text),
        labelSize:  parseFloat(labelSize.text),
        labelGap:   toPoints(labelGap.text)
    };
}

/**
 * Draw grid: rectangles + text labels.
 */
function createSwatchGrid(doc, swatches, opt) {
    var abIndex = doc.artboards.getActiveArtboardIndex();
    var ab      = doc.artboards[abIndex].artboardRect; // [left, top, right, bottom]

    var startX = ab[0] + opt.marginLeft;
    var startY = ab[1] - opt.marginTop;

    var cols = opt.columns;
    if (cols < 1) cols = 1;

    var chipW = opt.chipWidth;
    var chipH = opt.chipHeight;
    var gapX  = opt.gapX;
    var gapY  = opt.gapY;

    var labelSize = opt.labelSize;
    var labelGap  = opt.labelGap;

    // Build a simple black color for text
    var black = new CMYKColor();
    black.cyan    = 0;
    black.magenta = 0;
    black.yellow  = 0;
    black.black   = 100;

    var allGroup = doc.groupItems.add();
    allGroup.name = "Swatch Grid";

    for (var i = 0; i < swatches.length; i++) {
        var sw = swatches[i];
        var col = i % cols;
        var row = Math.floor(i / cols);

        var x = startX + col * (chipW + gapX);
        var y = startY - row * (chipH + gapY + labelGap + labelSize); // label + chip spacing baked into row height

        // Chip rectangle
        var chip = doc.pathItems.rectangle(y, x, chipW, chipH);
        chip.stroked = false;
        chip.filled  = true;
        chip.fillColor = sw.color;

        // Label text below chip
        var labelY = y - chipH - labelGap;
        var label = doc.textFrames.add();
        label.contents = sw.name;
        label.left = x;           // top-left of text box
        label.top  = labelY;      // top baseline position

        var tr = label.textRange;
        tr.characterAttributes.size = labelSize;
        tr.characterAttributes.fillColor = black;
        try {
            tr.paragraphAttributes.justification = Justification.LEFT;
        } catch (e) {}

        // Group chip + label
        var g = doc.groupItems.add();
        chip.move(g, ElementPlacement.PLACEATEND);
        label.move(g, ElementPlacement.PLACEATEND);
        g.move(allGroup, ElementPlacement.PLACEATEND);
    }
}
