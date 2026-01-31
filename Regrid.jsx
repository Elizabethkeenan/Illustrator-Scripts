#target illustrator

function main() {
    if (app.documents.length === 0) {
        alert("Open a document first.");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (!sel || sel.length === 0) {
        alert("Select the grid items (or a group containing them) first.");
        return;
    }

    // If a single group is selected, use its children as the grid items
    if (sel.length === 1 && sel[0].typename === "GroupItem") {
        sel = sel[0].pageItems;
    }

    // Collect items (treat each group as a single cell, so chip+label stay together)
    var items = [];
    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        if (it.typename === "GroupItem") {
            items.push(it);
        } else {
            items.push(it);
        }
    }

    if (items.length < 2) {
        alert("Need at least 2 items in the grid.");
        return;
    }

    // --- Measure centers and sizes ---
    var infos = [];
    var sumW = 0, sumH = 0;

    for (i = 0; i < items.length; i++) {
        var b = items[i].visibleBounds; // [L, T, R, B]
        var cx = (b[0] + b[2]) / 2;
        var cy = (b[1] + b[3]) / 2;
        var w = b[2] - b[0];
        var h = b[1] - b[3];

        sumW += w;
        sumH += h;

        infos.push({
            item: items[i],
            cx: cx,
            cy: cy,
            w: w,
            h: h
        });
    }

    var avgW = sumW / items.length;
    var avgH = sumH / items.length;

    // Tolerance to decide when a new column/row starts
    var colTol = avgW * 0.5;
    var rowTol = avgH * 0.5;

    // --- Detect columns (by X center) ---
    var tmp = infos.slice().sort(function (a, b) { return a.cx - b.cx; });
    var colCenters = [];

    for (i = 0; i < tmp.length; i++) {
        var x = tmp[i].cx;
        if (colCenters.length === 0 || Math.abs(x - colCenters[colCenters.length - 1]) > colTol) {
            colCenters.push(x);
        } else {
            // merge into existing cluster center
            var idx = colCenters.length - 1;
            colCenters[idx] = (colCenters[idx] + x) / 2;
        }
    }

    // Assign column index to each item
    for (i = 0; i < infos.length; i++) {
        var bestCol = -1;
        var bestDist = 1e9;
        for (var c = 0; c < colCenters.length; c++) {
            var d = Math.abs(infos[i].cx - colCenters[c]);
            if (d < bestDist) {
                bestDist = d;
                bestCol = c;
            }
        }
        infos[i].col = bestCol;
    }

    // --- Detect rows (by Y center, top to bottom) ---
    tmp = infos.slice().sort(function (a, b) { return b.cy - a.cy; });
    var rowCenters = [];

    for (i = 0; i < tmp.length; i++) {
        var y = tmp[i].cy;
        if (rowCenters.length === 0 || Math.abs(y - rowCenters[rowCenters.length - 1]) > rowTol) {
            rowCenters.push(y);
        } else {
            var idxR = rowCenters.length - 1;
            rowCenters[idxR] = (rowCenters[idxR] + y) / 2;
        }
    }

    // Assign row index
    for (i = 0; i < infos.length; i++) {
        var bestRow = -1;
        var bestRowDist = 1e9;
        for (var r = 0; r < rowCenters.length; r++) {
            var dr = Math.abs(infos[i].cy - rowCenters[r]);
            if (dr < bestRowDist) {
                bestRowDist = dr;
                bestRow = r;
            }
        }
        infos[i].row = bestRow;
    }

    var cols = colCenters.length;
    var rows = rowCenters.length;

    // --- Average spacing (center-to-center) ---
    var colSpacing = 0;
    var rowSpacing = 0;

    if (cols > 1) {
        var sumX = 0;
        for (i = 0; i < cols - 1; i++) {
            sumX += (colCenters[i + 1] - colCenters[i]);
        }
        colSpacing = sumX / (cols - 1);
    }

    if (rows > 1) {
        var sumY = 0;
        for (i = 0; i < rows - 1; i++) {
            // Y decreases downward, so subtract in this order
            sumY += (rowCenters[i] - rowCenters[i + 1]);
        }
        rowSpacing = sumY / (rows - 1);
    }

    var colIn = cols > 1 ? (colSpacing / 72).toFixed(3) : "0.000";
    var rowIn = rows > 1 ? (rowSpacing / 72).toFixed(3) : "0.000";

    // --- UI dialog ---
    var win = new Window("dialog", "Grid Analyzer / Grouper");
    win.orientation = "column";
    win.alignChildren = "left";

    win.add("statictext", undefined, "Detected grid:");
    win.add("statictext", undefined, "Columns: " + cols);
    win.add("statictext", undefined, "Rows: " + rows);
    win.add("statictext", undefined,
        "Column spacing (center–center): " + colIn + " in");
    win.add("statictext", undefined,
        "Row spacing (center–center): " + rowIn + " in");

    // Grouping options
    var panelGroup = win.add("panel", undefined, "Optional grouping");
    panelGroup.orientation = "column";
    panelGroup.alignChildren = "left";

    var rbNone = panelGroup.add("radiobutton", undefined, "No grouping (just report)");
    var rbCols = panelGroup.add("radiobutton", undefined, "Group by columns (vertical)");
    var rbRows = panelGroup.add("radiobutton", undefined, "Group by rows (horizontal)");
    rbNone.value = true;

    // Even spacing options
    var panelEven = win.add("panel", undefined, "Even spacing (optional)");
    panelEven.orientation = "column";
    panelEven.alignChildren = "left";

    var cbEvenCols = panelEven.add("checkbox", undefined, "Normalize column spacing");
    var cbEvenRows = panelEven.add("checkbox", undefined, "Normalize row spacing");

    var btns = win.add("group");
    btns.alignment = "right";
    var okBtn = btns.add("button", undefined, "OK");
    var cancelBtn = btns.add("button", undefined, "Cancel", { name: "cancel" });

    okBtn.onClick = function () { win.close(1); };
    cancelBtn.onClick = function () { win.close(0); };

    var result = win.show();
    if (result !== 1) return;

    // --- Apply grouping if requested ---
    var parent = items[0].parent;

    if (rbCols.value) {
        var colGroups = [];
        for (c = 0; c < cols; c++) {
            colGroups[c] = parent.groupItems.add();
        }
        for (i = 0; i < infos.length; i++) {
            var gCol = colGroups[infos[i].col];
            infos[i].item.moveToBeginning(gCol);
        }
    } else if (rbRows.value) {
        var rowGroups = [];
        for (r = 0; r < rows; r++) {
            rowGroups[r] = parent.groupItems.add();
        }
        for (i = 0; i < infos.length; i++) {
            var gRow = rowGroups[infos[i].row];
            infos[i].item.moveToBeginning(gRow);
        }
    }

    // --- Normalize spacing if requested ---
    // Re-read centers (positions may not have changed, but cheap to recompute)
    function refreshCenters() {
        for (var k = 0; k < infos.length; k++) {
            var bb = infos[k].item.visibleBounds;
            infos[k].cx = (bb[0] + bb[2]) / 2;
            infos[k].cy = (bb[1] + bb[3]) / 2;
        }
    }

    refreshCenters();

    if (cbEvenCols.value && cols > 1) {
        var startX = colCenters[0];
        for (c = 0; c < cols; c++) {
            var targetCx = startX + c * colSpacing;
            for (i = 0; i < infos.length; i++) {
                if (infos[i].col === c) {
                    var b2 = infos[i].item.visibleBounds;
                    var cx2 = (b2[0] + b2[2]) / 2;
                    var dx = targetCx - cx2;
                    infos[i].item.left += dx;
                }
            }
        }
        refreshCenters();
    }

    if (cbEvenRows.value && rows > 1) {
        var startY = rowCenters[0]; // topmost
        for (r = 0; r < rows; r++) {
            var targetCy = startY - r * rowSpacing;
            for (i = 0; i < infos.length; i++) {
                if (infos[i].row === r) {
                    var b3 = infos[i].item.visibleBounds;
                    var cy3 = (b3[1] + b3[3]) / 2;
                    var dy = targetCy - cy3;
                    infos[i].item.top += dy;
                }
            }
        }
    }
}

main();
