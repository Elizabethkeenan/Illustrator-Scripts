// Label_Helper_PairedBars_LR_TB_Horizontal_preview.jsx
// Two labels per run: Left&Right OR Top&Bottom.
// Bars and text land on "Labels" layer; text always above bars.
// Adds/uses a "CutContour" 100% magenta spot stroke for the bars.
// Now with: saved settings (Harmonizer-style) + Preview.

(function () {
    // ---- guards ----
    if (app.documents.length === 0) { alert("Open a document first."); return; }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) { alert("Select one or more objects and run again."); return; }

    // ---- Harmonizer-style settings file ----
    var scriptName = 'Label_Bars_Paired',
        settingFile = {
            name: scriptName + '__setting.json',
            folder: Folder.myDocuments + '/LA_AI_Scripts/'
        };

    var isUndo = false; // for preview state

    function checkSettingFolder() {
        var f = new Folder(settingFile.folder);
        if (!f.exists) f.create();
    }

    function saveSettings(cfg) {
        try {
            checkSettingFolder();
            var f = new File(settingFile.folder + settingFile.name);
            var data = [
                cfg.scale,
                cfg.barWidthIn,
                cfg.gapIn,
                cfg.barHeightIn,
                cfg.label,
                cfg.flipSecond,
                cfg.fontName,
                cfg.placement
            ].toString();
            f.open('w');
            f.write(data);
            f.close();
        } catch (e) {
            // ignore
        }
    }

    function loadSettings() {
        checkSettingFolder();
        var f = new File(settingFile.folder + settingFile.name);
        if (!f.exists) return null;
        try {
            f.open('r');
            var data = f.read().split('\n'),
                main = data[0].split(',');
            f.close();
            return {
                scale:       parseFloat(main[0]) || 1,
                barWidthIn:  parseFloat(main[1]) || 1.5,
                gapIn:       parseFloat(main[2]) || 0.25,
                barHeightIn: parseFloat(main[3]) || 1.0,
                label:       main[4] || '',
                flipSecond:  (main[5] === 'true'),
                fontName:    main[6] || 'MyriadPro-Regular',
                placement:   main[7] || 'Top & Bottom'
            };
        } catch (e) {
            try { f.close(); } catch (ee) {}
            return null;
        }
    }

    // ---- core geometry + drawing helpers ----
    var PT = 72;
    function IN(v) { return v * PT; } // v is already full-size inches * cfg.scale

    function getCutContourSpotColor() {
        var spot;
        try { spot = doc.spots.getByName("CutContour"); } catch (e) {
            spot = doc.spots.add();
            spot.name = "CutContour";
            var c = new CMYKColor();
            c.cyan = 0; c.magenta = 100; c.yellow = 0; c.black = 0; // 100% M
            spot.color = c;
        }
        var sc = new SpotColor();
        sc.spot = spot;
        sc.tint = 100;
        return sc;
    }
    var cutContourSC = getCutContourSpotColor();

    function trimStr(s) {
        if (s === undefined || s === null) return "";
        s = "" + s;
        return s.replace(/^\s+/, "").replace(/\s+$/, "");
    }

    // clip-aware bounds helper
    function findClipPathBounds(item) {
        function clippingAncestor(it) {
            var cur = it;
            while (cur && cur.parent) {
                if (cur.typename === "GroupItem" && cur.clipped) return cur;
                cur = cur.parent;
            }
            return null;
        }

        var target = item;
        if (!(item.typename === "GroupItem" && item.clipped)) {
            var anc = clippingAncestor(item);
            if (anc) target = anc;
        }

        if (target && target.typename === "GroupItem" && target.clipped) {
            for (var i = 0; i < target.pageItems.length; i++) {
                var pi = target.pageItems[i];
                if (pi.typename === "PathItem" && pi.clipping) {
                    return pi.geometricBounds;
                }
                if (pi.typename === "CompoundPathItem" &&
                    pi.pathItems.length &&
                    pi.pathItems[0].clipping) {
                    return pi.geometricBounds;
                }
            }
        }
        return null;
    }

    function fallbackFitAndCenter(tf, x, top, w, h, textPad) {
        var maxW = Math.max(2, w - 2 * textPad);
        var maxH = Math.max(2, h - 2 * textPad);
        var tr = tf.textRange;
        var start = 500;
        tr.characterAttributes.size = start;
        tr.characterAttributes.leading = start * 1.08;

        var gb = tf.geometricBounds;
        var curW = Math.max(1, gb[2] - gb[0]);
        var curH = Math.max(1, gb[1] - gb[3]);
        var scale = Math.min(maxW / curW, maxH / curH);
        var finalSize = Math.max(1, start * scale);
        tr.characterAttributes.size = finalSize;
        tr.characterAttributes.leading = Math.max(1, finalSize * 1.08);

        gb = tf.geometricBounds;
        var barCX = x + w / 2;
        var barCY = top - h / 2;
        tf.translate(barCX - (gb[0] + gb[2]) / 2, barCY - (gb[1] + gb[3]) / 2);
        return tf;
    }

    function addTextMax(label, x, top, w, h, rotateDeg, layer, chosenFont, textPad) {
        if (!label) return null;

        var maxW = Math.max(2, w - 2 * textPad);
        var maxH = Math.max(2, h - 2 * textPad);

        // live text
        var tf = layer.textFrames.add();
        tf.contents = label;
        tf.kind = TextType.POINTTEXT;
        var tr = tf.textRange;
        if (chosenFont) tr.characterAttributes.textFont = chosenFont;
        tr.paragraphAttributes.justification = Justification.CENTER;

        var start = 800;
        tr.characterAttributes.size = start;
        tr.characterAttributes.leading = start * 1.08;

        // bar center
        var barCX = x + w / 2;
        var barCY = top - h / 2;

        // rough center before outlining
        var gb0 = tf.geometricBounds;
        var curCX0 = (gb0[0] + gb0[2]) / 2;
        var curCY0 = (gb0[1] + gb0[3]) / 2;
        tf.translate(barCX - curCX0, barCY - curCY0);

        // outlines for accurate fitting
        var dup = tf.duplicate();
        var outlines = null;
        try {
            outlines = dup.createOutline(); // dup replaced by outlines group
        } catch (e) {
            try { dup.remove(); } catch (ee) {}
            return fallbackFitAndCenter(tf, x, top, w, h, textPad);
        }

        // rotate outlines just for measuring, if needed
        if (rotateDeg) {
            try { outlines.rotate(rotateDeg, true, true, true, true, Transformation.CENTER); } catch (e2) {}
        }

        var ogb = outlines.visibleBounds;
        var oW = Math.max(1, ogb[2] - ogb[0]);
        var oH = Math.max(1, ogb[1] - ogb[3]);
        var scale = Math.min(maxW / oW, maxH / oH);
        var finalSize = Math.max(1, start * scale);

        // cleanup temp outlines
        try { outlines.remove(); } catch (e3) {}

        // apply final size
        tr.characterAttributes.size = finalSize;
        tr.characterAttributes.leading = Math.max(1, finalSize * 1.08);

        // recenter unrotated
        var gb = tf.geometricBounds;
        var curCX = (gb[0] + gb[2]) / 2;
        var curCY = (gb[1] + gb[3]) / 2;
        tf.translate(barCX - curCX, barCY - curCY);

        // rotate final text if requested
        if (rotateDeg) {
            try { tf.rotate(rotateDeg, true, true, true, true, Transformation.CENTER); } catch (e4) {}
        }

        // recenter AFTER rotation using visible bounds
        var vb = tf.visibleBounds;
        var vCX = (vb[0] + vb[2]) / 2;
        var vCY = (vb[1] + vb[3]) / 2;
        tf.translate(barCX - vCX, barCY - vCY);

        // safety shrink in case rounding overflowed
        var b2 = tf.visibleBounds;
        var curW = Math.max(1, b2[2] - b2[0]);
        var curH = Math.max(1, b2[1] - b2[3]);
        if (curW > maxW + 0.5 || curH > maxH + 0.5) {
            var shrink = Math.min(maxW / curW, maxH / curH) * 0.995;
            var newSize = Math.max(1, tr.characterAttributes.size * shrink);
            tr.characterAttributes.size = newSize;
            tr.characterAttributes.leading = Math.max(1, newSize * 1.08);

            // recenter again after shrink
            vb = tf.visibleBounds;
            vCX = (vb[0] + vb[2]) / 2;
            vCY = (vb[1] + vb[3]) / 2;
            tf.translate(barCX - vCX, barCY - vCY);
        }

        tf.zOrder(ZOrderMethod.BRINGTOFRONT);
        return tf;
    }

    function setupBar(bar, cfg) {
        bar.stroked = true;
        bar.strokeColor = cutContourSC;
        bar.strokeWidth = Math.max(0.1, 1 * cfg.scale);
        bar.filled = false;
        bar.zOrder(ZOrderMethod.BRINGTOFRONT);
    }

    // ---- core runner: uses cfg and current selection ----
    function runLabelMaker(cfg) {
        if (!cfg) return;

        var layer;
        try { layer = doc.layers.getByName("Labels"); }
        catch (e) { layer = doc.layers.add(); layer.name = "Labels"; }
        layer.locked = false; layer.visible = true;

        var barW    = IN(cfg.barWidthIn  * cfg.scale);
        var gap     = IN(cfg.gapIn       * cfg.scale);
        var barH    = IN(cfg.barHeightIn * cfg.scale);
        var textPad = IN(0.04 * cfg.scale);

        var chosenFont = null;
        try { if (cfg.fontName) { chosenFont = app.textFonts.getByName(cfg.fontName); } } catch (e) {}
        if (!chosenFont) { try { chosenFont = app.textFonts.getByName("MyriadPro-Regular"); } catch (e2) {} }

        var created = [];
        var sel = doc.selection;
        for (var i = 0; i < sel.length; i++) {
            var it = sel[i];

            var b = findClipPathBounds(it);
            if (!b) {
                try { b = it.geometricBounds; }
                catch (e1) {
                    try { b = it.visibleBounds; } catch (e2) { continue; }
                }
            }

            var L = b[0], T = b[1], R = b[2], B = b[3];
            var objH = T - B;

            if (cfg.placement === "Left & Right") {
                var leftX  = L - gap - barW;
                var rightX = R + gap;
                var barTopLR = T - ((objH - barH) / 2);

                var leftBar  = layer.pathItems.rectangle(barTopLR, leftX,  barW, barH);
                var rightBar = layer.pathItems.rectangle(barTopLR, rightX, barW, barH);
                setupBar(leftBar, cfg);
                setupBar(rightBar, cfg);
                created.push(leftBar, rightBar);

                // Vertical LR like ambulance numbers
                var angLeft  = -90;
                var angRight = 90;

                if (cfg.label) {
                    var ltf = addTextMax(cfg.label, leftX,  barTopLR, barW, barH, angLeft,  layer, chosenFont, textPad);
                    var rtf = addTextMax(cfg.label, rightX, barTopLR, barW, barH, angRight, layer, chosenFont, textPad);
                    if (ltf) created.push(ltf);
                    if (rtf) created.push(rtf);
                }

            } else { // Top & Bottom
                var objCX = (L + R) / 2;
                var barX  = objCX - barW / 2;

                var topBarTop    = T + gap + barH;
                var bottomBarTop = B - gap - barH;

                var topBar    = layer.pathItems.rectangle(topBarTop,    barX, barW, barH);
                var bottomBar = layer.pathItems.rectangle(bottomBarTop, barX, barW, barH);
                setupBar(topBar, cfg);
                setupBar(bottomBar, cfg);
                created.push(topBar, bottomBar);

                var angTop    = 0;
                var angBottom = 0;

                if (cfg.label) {
                    var ttf = addTextMax(cfg.label, barX, topBarTop,    barW, barH, angTop,    layer, chosenFont, textPad);
                    var btf = addTextMax(cfg.label, barX, bottomBarTop, barW, barH, angBottom, layer, chosenFont, textPad);
                    if (ttf) created.push(ttf);
                    if (btf) created.push(btf);
                }
            }
        }

        if (created.length) {
            doc.selection = created;
        }
        app.redraw();
    }

    // ---- dialog with Preview ----
    function showOptions() {
        var prefs = loadSettings() || {
            scale:       1,
            barWidthIn:  1.5,
            gapIn:       0.25,
            barHeightIn: 1.0,
            label:       '',
            flipSecond:  true,
            fontName:    'MyriadPro-Regular',
            placement:   'Top & Bottom'
        };

        var w = new Window('dialog', 'Label Bars (Paired) – Preview');
        w.orientation = 'column'; 
        w.alignChildren = ['fill', 'top']; 
        w.margins = 14;

        function numBox(parent, label, def, chars) {
            var g = parent.add('group'); 
            g.orientation = 'row';
            g.add('statictext', undefined, label);
            var e = g.add('edittext', undefined, def); 
            e.characters = chars || 8; 
            return e;
        }

        // ---- Scale ----
        var scaleP = w.add('panel', undefined, 'Scale');
        scaleP.orientation = 'row'; 
        scaleP.alignChildren = ['left', 'top']; 
        scaleP.margins = 12;
        var rFull   = scaleP.add('radiobutton', undefined, 'Full (1:1)');
        var rTenth  = scaleP.add('radiobutton', undefined, '1/10 (0.1)');
        var rCustom = scaleP.add('radiobutton', undefined, 'Custom:');
        var customE = scaleP.add('edittext', undefined, '1'); 
        customE.characters = 6; 
        customE.enabled = false;

        // preset from prefs.scale
        if (Math.abs(prefs.scale - 1) < 0.0001) {
            rFull.value = true;
        } else if (Math.abs(prefs.scale - 0.1) < 0.0001) {
            rTenth.value = true;
        } else {
            rCustom.value = true;
            customE.text = String(prefs.scale);
            customE.enabled = true;
        }
        function updScaleUI() { customE.enabled = (rCustom.value === true); }
        rFull.onClick = updScaleUI; 
        rTenth.onClick = updScaleUI; 
        rCustom.onClick = updScaleUI;
        updScaleUI();

        // ---- Sizes ----
        var sizeP = w.add('panel', undefined, 'Sizes (Full-Size Inches)');
        sizeP.orientation = 'column'; 
        sizeP.alignChildren = ['fill', 'top']; 
        sizeP.margins = 12;
        var wEdt  = numBox(sizeP, 'Bar width:',  String(prefs.barWidthIn), 8);
        var gEdt  = numBox(sizeP, 'Gap:',        String(prefs.gapIn), 8);
        var hEdt  = numBox(sizeP, 'Bar height:', String(prefs.barHeightIn), 8);

        // ---- Label ----
        var labP = w.add('panel', undefined, 'Label');
        labP.orientation = 'column'; 
        labP.alignChildren = ['left', 'top']; 
        labP.margins = 12;
        var rbGroup = labP.add('group'); 
        rbGroup.orientation = 'column'; 
        rbGroup.alignChildren = ['left', 'top'];
        function rb(t) { return rbGroup.add('radiobutton', undefined, t); }
        var rNone  = rb('None'),
            rCS    = rb('CS'),
            rSS    = rb('SS'),
            rRear  = rb('REAR'),
            rFront = rb('FRONT'),
            rCust  = rb('CUSTOM');
        var custG = labP.add('group'); 
        custG.add('statictext', undefined, 'Custom:');
        var custE = custG.add('edittext', undefined, ''); 
        custE.characters = 14; 
        custG.enabled = false;

        function updLab() { custG.enabled = (rCust.value === true); }

        function labelClick() { updLab(); previewStart(); }

rNone.onClick  = labelClick;
rCS.onClick    = labelClick;
rSS.onClick    = labelClick;
rRear.onClick  = labelClick;
rFront.onClick = labelClick;
rCust.onClick  = labelClick;

        // preset label radio from prefs.label
        if (!prefs.label) {
            rNone.value = true;
        } else if (prefs.label === 'CS') {
            rCS.value = true;
        } else if (prefs.label === 'SS') {
            rSS.value = true;
        } else if (prefs.label === 'REAR') {
            rRear.value = true;
        } else if (prefs.label === 'FRONT') {
            rFront.value = true;
        } else {
            rCust.value = true;
            custE.text = prefs.label;
        }
        updLab();

        // ---- Options ----
        var optP = w.add('panel', undefined, 'Options');
        optP.orientation = 'column'; 
        optP.alignChildren = ['fill', 'top']; 
        optP.margins = 12;
        var flipCB = optP.add('checkbox', undefined, 'Flip SECOND label (Right in LR / ignored in TB)');
        // we keep flipSecond here just for future expansion; not used in geometry yet
        flipCB.value = (prefs.flipSecond !== false);
        var fontGrp = optP.add('group'); 
        fontGrp.add('statictext', undefined, 'Font (PostScript name):');
        var fontEdt = fontGrp.add('edittext', undefined, prefs.fontName || 'MyriadPro-Regular'); 
        fontEdt.characters = 20;

        // ---- Placement ----
        var placeP = w.add('panel', undefined, 'Placement');
        placeP.orientation = 'row'; 
        placeP.alignChildren = ['left', 'center']; 
        placeP.margins = 12;
        var placeDD = placeP.add('dropdownlist', undefined, ['Left & Right', 'Top & Bottom']);
        placeDD.selection = (prefs.placement === 'Left & Right') ? 0 : 1;

        // ---- Preview ----
        var previewChk = w.add('checkbox', undefined, 'Preview');

        // ---- Buttons ----
        var btns = w.add('group'); 
        btns.alignment = 'right';
        var cancelBtn = btns.add('button', undefined, 'Cancel', { name: 'cancel' });
        var okBtn     = btns.add('button', undefined, 'OK',     { name: 'ok' });
        w.defaultElement = okBtn;

        // build cfg from current UI
        function num(v, d) {
            var n = parseFloat(v);
            return (isNaN(n) || n <= 0) ? d : n;
        }

        function pickLabel() {
            if (rNone.value)  return '';
            if (rCS.value)    return 'CS';
            if (rSS.value)    return 'SS';
            if (rRear.value)  return 'REAR';
            if (rFront.value) return 'FRONT';
            if (rCust.value)  return trimStr(custE.text);
            return '';
        }

        function currentCfg() {
            var scaleVal = 1;
            if (rFull.value) scaleVal = 1;
            else if (rTenth.value) scaleVal = 0.1;
            else if (rCustom.value) scaleVal = num(customE.text, 1);

            return {
                scale:        scaleVal,
                barWidthIn:   num(wEdt.text, 1.5),
                gapIn:        num(gEdt.text, 0.25),
                barHeightIn:  num(hEdt.text, 1.0),
                label:        pickLabel(),
                flipSecond:   (flipCB.value === true),
                fontName:     ('' + fontEdt.text).replace(/^\s+|\s+$/g, ''),
                placement:    placeDD.selection ? placeDD.selection.text : 'Top & Bottom'
            };
        }

        // ---- Preview logic (Harmonizer-style) ----
        function previewStart() {
            if (!previewChk.value) {
                if (isUndo) {
                    app.undo();
                    app.redraw();
                    isUndo = false;
                }
                return;
            }

            if (isUndo) {
                app.undo();
            } else {
                isUndo = true;
            }

            runLabelMaker(currentCfg());
            app.redraw();
        }

        // attach preview to controls
       var liveControls = [
    wEdt, gEdt, hEdt,
    customE, fontEdt,
    custE,              // <-- keep this so typing custom text updates preview
    rFull, rTenth, rCustom,
    flipCB
];



        for (var i = 0; i < liveControls.length; i++) {
            if (!liveControls[i]) continue;
            if (liveControls[i].onChanging !== undefined) {
                liveControls[i].onChanging = previewStart;
            } else {
                liveControls[i].onClick = previewStart;
            }
        }
        if (placeDD) placeDD.onChange = previewStart;
        previewChk.onClick = previewStart;

        // ensure preview is undone when dialog closes
        w.onClose = function () {
            if (isUndo) {
                app.undo();
                app.redraw();
                isUndo = false;
            }
            return true;
        };

        var finalCfg = null;
        okBtn.onClick = function () {
            finalCfg = currentCfg();
            saveSettings(finalCfg);
            w.close(1);
        };
        cancelBtn.onClick = function () { w.close(0); };

        if (w.show() !== 1) return null;
        return finalCfg;
    }

    // ---- run ----
    var cfg = showOptions();
    if (!cfg) return;

    // After dialog closes, any preview has been undone by w.onClose.
    // Now apply final result once.
    runLabelMaker(cfg);
})();
