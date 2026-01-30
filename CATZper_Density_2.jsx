#target illustrator

(function () {
  if (app.documents.length === 0) {
    alert("Open a document first.");
    return;
  }

  var doc = app.activeDocument;
  var sel = doc.selection;
  if (!sel || sel.length === 0) {
    alert("Select your CATZper grid (group is OK), then run again.");
    return;
  }

  // ===================== SETTINGS =====================
  // "10% darker" (recommended): add density mostly via K + a little Y buffer
  var darkenK = 10;   // main density increase
  var darkenY = 3;    // warmth buffer
  var darkenC = 0;
  var darkenM = 0;

  // Offset for the duplicated grid (points)
  var xOffset = 500;  // move right
  var yOffset = 0;    // move up/down (usually 0)
  // ====================================================

  function clamp(v) { return Math.max(0, Math.min(100, v)); }

  function darkenCMYK(col) {
    var n = new CMYKColor();
    n.cyan    = clamp(col.cyan    + darkenC);
    n.magenta = clamp(col.magenta + darkenM);
    n.yellow  = clamp(col.yellow  + darkenY);
    n.black   = clamp(col.black   + darkenK);
    return n;
  }

  // Collect all filled CMYK pathItems recursively from selection
  var patches = [];

  function collect(item) {
    if (!item) return;

    // Groups
    if (item.typename === "GroupItem") {
      for (var i = 0; i < item.pageItems.length; i++) collect(item.pageItems[i]);
      return;
    }

    // Compound paths
    if (item.typename === "CompoundPathItem") {
      // compoundPathItem.pathItems are the actual shapes
      for (var j = 0; j < item.pathItems.length; j++) collect(item.pathItems[j]);
      return;
    }

    // Path items (rectangles, etc.)
    if (item.typename === "PathItem") {
      if (item.filled && item.fillColor && item.fillColor.typename === "CMYKColor") {
        patches.push(item);
      }
      return;
    }

    // Everything else (TextFrame, PlacedItem, RasterItem, etc.) ignored
  }

  for (var s = 0; s < sel.length; s++) collect(sel[s]);

  if (patches.length === 0) {
    alert("I couldn't find any filled CMYK patches in your selection.\n\nTip: Make sure the patch rectangles are CMYK fills (not RGB/Gray/Spot), and select the grid group before running.");
    return;
  }

  // Make a destination group on the active layer
  var dest = doc.activeLayer.groupItems.add();
  dest.name = "CATZper +Density Twin";

  var processed = 0;

  for (var p = 0; p < patches.length; p++) {
    var src = patches[p];

    // duplicate the patch
    var dup = src.duplicate(dest, ElementPlacement.PLACEATEND);

    // apply darker fill
    dup.fillColor = darkenCMYK(src.fillColor);

    // move it
    dup.translate(xOffset, yOffset);

    processed++;
  }

  alert("Done.\nDuplicated + darkened patches: " + processed + "\n(Placed in group: " + dest.name + ")");
})();
