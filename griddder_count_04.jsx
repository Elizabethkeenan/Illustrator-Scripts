/* 

  Author: Alexander Ladygin (i@ladygin.pro)
  Program version: Adobe Illustrator CS5+
  Name: griddder.jsx;

  Copyright (c) 2018
  www.ladyginpro.ru
This code has been heavily edited from the original Gridder to better suit step and repeat behaviour for printing. 
Select what you wish to print in a grid format
*/

$.errorMessage = function (err) { alert(err + '\n' + err.line); };

function LA(obj, callback, reverse) {
    if (!callback) {
        if (obj instanceof Array) {
            return obj;
        } else {
            var arr = $.getArr(obj);
            if (arr === obj) {
                if ($.isColor(obj)) {
                    return obj;
                } else {
                    return [obj];
                }
            }
            return arr;
        }
    } else if (callback instanceof Function) {
        var arr = $.getArr(obj);
        if (arr === obj) { arr = [obj]; }
        if (reverse) {
            var i = arr.length;
            while (i--) callback(arr[i], i, arr);
        } else {
            for (var j = 0; j < arr.length; j++) callback(arr[j], j, arr);
        }
        return arr;
    }
}

$.each = function (object, callback, reverse) {
    try {
        if (object && object.length) {
            var l = object.length;
            if (!reverse) for (var i = 0; i < l; i++) callback(object[i], i, object);
            else while (l--) callback(object[l], l, object);
        }
        return $;
    } catch (e) { $.errorMessage('$.each() - error: ' + e); }
};

Object.prototype.each = function (callback, reverse) {
    if (this.length) $.each(this, callback, reverse);
    return this;
};

Object.prototype.getChildsByFilter = function (filterCallback, returnFirst) {
    filterCallback = filterCallback instanceof Function ? filterCallback : function () { return true; };
    var arr = [], items = LA(this), l = items.length;
    for (var i = 0; i < l; i++) {
        if (items[i].typename === 'GroupItem') {
            arr = arr.concat(LA(items[i].pageItems).getChildsByFilter(filterCallback));
        } else if (filterCallback(items[i])) {
            arr.push(items[i]);
            if (returnFirst) return arr;
        }
    }
    return arr;
};

$.isArr = function (a) {
    if ((!a) ||
        (typeof a === 'string') ||
        (a.typename === 'Document') ||
        (a.typename === 'Layer') ||
        (a.typename === 'PathItem') ||
        (a.typename === 'GroupItem') ||
        (a.typename === 'PageItem') ||
        (a.typename === 'CompoundPathItem') ||
        (a.typename === 'TextFrame') ||
        (a.typename === 'TextRange') ||
        (a.typename === 'GraphItem') ||
        (a.typename === 'Document') ||
        (a.typename === 'Artboard') ||
        (a.typename === 'LegacyTextItem') ||
        (a.typename === 'NoNNativeItem') ||
        (a.typename === 'Pattern') ||
        (a.typename === 'PlacedItem') ||
        (a.typename === 'PluginItem') ||
        (a.typename === 'RasterItem') ||
        (a.typename === 'MeshItem') ||
        (a.typename === 'SymbolItem')
    ) {
        return false;
    } else if (!a.typename && !(a instanceof Array)) {
        return false;
    } else {
        return true;
    }
};

$.getArr = function (obj, attr, value, exclude) {
    var arr = [];
    function checkExclude(item) {
        if (exclude !== undefined) {
            var j = exclude.length;
            while (j--) if (exclude[j] === item) return true;
        }
        return false;
    }
    if ($.isArr(obj)) {
        for (var i = 0; i < obj.length; i++) {
            if (!checkExclude(obj[i])) {
                if (attr) {
                    if (value !== undefined) {
                        arr.push(obj[i][attr][value]);
                    } else {
                        arr.push(obj[i][attr]);
                    }
                } else {
                    arr.push(obj[i]);
                }
            }
        }
        return arr;
    } else if (attr) {
        return obj[attr];
    } else {
        return obj;
    }
};

$.isColor = function (color) {
    if ((color.typename === 'GradientColor') ||
        (color.typename === 'PatternColor') ||
        (color.typename === 'CMYKColor') ||
        (color.typename === 'SpotColor') ||
        (color.typename === 'GrayColor') ||
        (color.typename === 'LabColor') ||
        (color.typename === 'RGBColor') ||
        (color.typename === 'NoColor')
    ) {
        return true;
    } else {
        return false;
    }
};

$.appName = {
    indesign: (BridgeTalk.appName.toLowerCase() === 'indesign'),
    photoshop: (BridgeTalk.appName.toLowerCase() === 'photoshop'),
    illustrator: (BridgeTalk.appName.toLowerCase() === 'illustrator')
};

$.color = function (a, v) {
    if (a) {
        if (typeof a === 'string') { a = a.toLowerCase(); }
    } else { return undefined; }

    if ((a === 'hex') && $.appName.illustrator) {
        if (!v) {
            return new RGBColor();
        } else {
            if (v === 'random') return $.color('rgb', v);
            else return $.hexToColor(v, 'RGB');
        }
    } else if ((a === 'cmyk') || (a === 'cmykcolor')) {
        var c = new CMYKColor(), b = [];
        if (v) {
            b = b.concat(v);
            if (typeof v === 'string' && v.toLowerCase() === 'random') {
                b = [Math.floor(Math.random() * 100), Math.floor(Math.random() * 100), Math.floor(Math.random() * 100), Math.floor(Math.random() * 100)];
            } else {
                for (var i = 0; i < b.length; i++) {
                    if (b[i] === 'random') { b[i] = Math.floor(Math.random() * 100); }
                }
            }
            c.cyan = parseInt(b[0]);
            c.magenta = parseInt(b[1]);
            c.yellow = parseInt(b[2]);
            c.black = parseInt(b[3]);
        }
        return c;
    } else if ((a === 'rgb') || (a === 'rgbcolor') || ((a === 'hex') && $.appName.photoshop)) {
        var c2 = new RGBColor(), b2 = [];
        if (v) {
            b2 = b2.concat(v);
            if (typeof v === 'string' && v.toLowerCase() === 'random') {
                b2 = [Math.floor(Math.random() * 255), Math.floor(Math.random() * 255), Math.floor(Math.random() * 255)];
            } else {
                for (var i2 = 0; i2 < b2.length; i2++) {
                    if (b2[i2] === 'random') { b2[i2] = Math.floor(Math.random() * 100); }
                }
            }
            if ($.appName.photoshop) {
                if (a !== 'hex' || (typeof v === 'string' && v.toLowerCase() === 'random')) {
                    c2.red = parseInt(b2[0]);
                    c2.green = parseInt(b2[1]);
                    c2.blue = parseInt(b2[2]);
                } else { c2.hexValue = b2[0]; }
            } else if ($.appName.illustrator) {
                c2.red = parseInt(b2[0]);
                c2.green = parseInt(b2[1]);
                c2.blue = parseInt(b2[2]);
            }
        }
        return c2;
    } else if ((a === 'gray') || (a === 'grayscale') || (a === 'graycolor')) {
        var c3 = new GrayColor(), b3 = [];
        if (v) {
            b3 = b3.concat(v);
            if (typeof v === 'string' && v.toLowerCase() === 'random') { b3 = Math.floor(Math.random() * 100); }
            c3.gray = parseInt(b3[0] || b3);
        }
        return c3;
    } else if ((a === 'lab') || (a === 'labcolor')) {
        var c4 = new LabColor(), b4 = [];
        if (v) {
            b4 = b4.concat(v);
            if (typeof v === 'string' && v.toLowerCase() === 'random') {
                b4 = [Math.floor(Math.random() * 100), Math.floor(-128 + Math.random() * 256), Math.floor(-128 + Math.random() * 256)];
            } else {
                for (var i4 = 0; i4 < b4.length; i4++) {
                    if (i4 === 0) {
                        if (b4[i4] === 'random') { b4[i4] = Math.floor(Math.random() * 100); }
                    } else {
                        if (b4[i4] === 'random') { b4[i4] = Math.floor(-128 + Math.random() * 256); }
                    }
                }
            }
            c4.l = parseInt(b4[0]);
            c4.a = parseInt(b4[1]);
            c4.b = parseInt(b4[2]);
        }
        return c4;
    } else if ((a === 'spot') || (a === 'spotcolor')) {
        var c5 = new SpotColor(), b5 = [];
        if (v) {
            b5 = b5.concat(v);
            c5.tint = parseInt(b5[1]);
        }
        return c5;
    } else if ((a === 'gradient') || (a === 'Gradient') || (a === 'GradientColor')) {
        var c6 = app.activeDocument.gradients.add(), g = new GradientColor(), b6 = [];
        if (v) {
            b6 = b6.concat(v);
            for (var i6 = 0; i6 < b6.length; i6++) {
                c6.gradientStops[i6].color = $.color(b6[i6][0], b6[i6][1]);
            }
            g.gradient = c6;
        }
        return g;
    } else if ((a === 'no') || (a === 'nocolor')) {
        return new NoColor();
    }
};

$.toHex = function (color, hash) {
    if (color.typename !== 'RGBColor' && $.appName.illustrator) {
        color = $.convertColor(color, 'RGB');
    }
    return (hash ? '#' : '') + to(color.red) + to(color.green) + to(color.blue);
    function to(val) {
        var hex = val.toString(16);
        return hex.length === 1 ? '0' + hex : hex;
    }
};

$.hexToColor = function (color, type) {
    color = color.toLowerCase();
    color = correct(color);
    function correct(a) {
        var l, b = '000000';
        if (a[0] === '#') { a = a.slice(1); }
        l = a.length;
        a = a + b.slice(l);
        return a;
    }
    return $.convertColor($.color('rgb', [
        parseInt((gc(color)).slice(0, 2), 16),
        parseInt((gc(color)).slice(2, 4), 16),
        parseInt((gc(color)).slice(4, 6), 16)
    ]), type || 'rgb');
    function gc(h) { return (h.slice(0, 1) === '#') ? h.slice(1, 7) : h; }
};

$.getColorValues = function (color) {
    if (color === undefined) { return undefined; }
    else if (color.typename === 'CMYKColor') { return [color.cyan, color.magenta, color.yellow, color.black]; }
    else if (color.typename === 'RGBColor') { return [color.red, color.green, color.blue]; }
    else if (color.typename === 'LabColor') { return [color.l, color.a, color.b]; }
    else if (color.typename === 'SpotColor') { return [color.spotl, color.tint]; }
    else if (color.typename === 'GrayColor') { return [color.gray]; }
    else if (color.typename === 'NoColor') { return undefined; }
    else if (color.typename === 'GradientColor') {
        var colors = [], gradients = color.gradient.gradientStops;
        for (var i = 0; i < gradients.length; i++) {
            colors = colors.concat(gradients[i].color.getColorValues());
        }
        return colors;
    }
};

CMYKColor.prototype.getColorValues = function () { return $.getColorValues(this); };
RGBColor.prototype.getColorValues = function () { return $.getColorValues(this); };
GrayColor.prototype.getColorValues = function () { return $.getColorValues(this); };
LabColor.prototype.getColorValues = function () { return $.getColorValues(this); };
NoColor.prototype.getColorValues = function () { return $.getColorValues(this); };

$.getUnits = function (val, def) {
    try {
        return 'px,pt,mm,cm,in,pc'.indexOf(val.slice(-2)) > -1 ? val.slice(-2) : def;
    } catch (e) { $.errorMessage('check units: " ' + e + ' "'); }
};

// ----- INCHES LOGIC HERE -----
$.convertUnits = function (obj, b) {
    if (obj === undefined) { return obj; }

    // Illustrator geometry is in points; we keep 'px' as math target
    if (b === undefined) { b = 'px'; }

    // Treat bare numbers as inches
    if (typeof obj === 'number') {
        obj = obj + 'in';
    }

    var unit, val;

    if (typeof obj === 'string') {
        unit = $.getUnits(obj);
        val = parseFloat(obj);

        if (unit && !isNaN(val)) {
            obj = val;
        } else if (!isNaN(val)) {
            obj = val;
            unit = 'in'; // unitless → inches
        }
    }

    if ($.appName.illustrator) {
        if (((unit === 'px') || (unit === 'pt')) && (b === 'mm')) {
            obj = parseFloat(obj) / 2.83464566929134;
        } else if (((unit === 'px') || (unit === 'pt')) && (b === 'cm')) {
            obj = parseFloat(obj) / (2.83464566929134 * 10);
        } else if (((unit === 'px') || (unit === 'pt')) && (b === 'in')) {
            obj = parseFloat(obj) / 72;
        } else if ((unit === 'mm') && ((b === 'px') || (b === 'pt'))) {
            obj = parseFloat(obj) * 2.83464566929134;
        } else if ((unit === 'mm') && (b === 'cm')) {
            obj = parseFloat(obj) * 10;
        } else if ((unit === 'mm') && (b === 'in')) {
            obj = parseFloat(obj) / 25.4;
        } else if ((unit === 'cm') && ((b === 'px') || (b === 'pt'))) {
            obj = parseFloat(obj) * 2.83464566929134 * 10;
        } else if ((unit === 'cm') && (b === 'mm')) {
            obj = parseFloat(obj) / 10;
        } else if ((unit === 'cm') && (b === 'in')) {
            obj = parseFloat(obj) * 2.54;
        } else if ((unit === 'in') && ((b === 'px') || (b === 'pt'))) {
            obj = parseFloat(obj) * 72;
        } else if ((unit === 'in') && (b === 'mm')) {
            obj = parseFloat(obj) * 25.4;
        } else if ((unit === 'in') && (b === 'cm')) {
            obj = parseFloat(obj) * 25.4;
        }
        return parseFloat(obj);
    } else if ($.appName.photoshop) {
        return parseFloat(obj);
    }
};

// margins helper used by align() and griddder()
function parseMargin(value, ifErrReturnValue) {
    value = (typeof value === 'string' ? value.split(' ') : (value instanceof Array ? value : ''));
    if (!value.length) return ifErrReturnValue !== undefined ? ifErrReturnValue : [0, 0, 0, 0];

    if (value.length === 2) {
        value[2] = value[0];
        value[3] = value[1];
    } else if (value.length < 4) {
        var val = value[value.length - 1];
        for (var i = value.length; i < 4; i++) {
            value[i] = val;
        }
    }

    for (var j = 0; j < value.length; j++) {
        value[j] = $.convertUnits(value[j], 'px');
    }

    return value;
}

Object.prototype.extend = function (userObject, deep) {
    try {
        for (var key in userObject) {
            if (this.hasOwnProperty(key)) {
                if (deep &&
                    this[key] instanceof Object &&
                    !(this[key] instanceof Array) &&
                    userObject[key] instanceof Object &&
                    !(userObject[key] instanceof Array)) {
                    this[key].extend(userObject[key], deep);
                } else {
                    this[key] = userObject[key];
                }
            }
        }
        return this;
    } catch (e) {
        $.errorMessage('$.objectParser() - error: ' + e);
    }
};

function getBounds(items, bounds) {
    bounds = bounds || 'geometricBounds';
    bounds = (bounds && bounds.toLowerCase().indexOf('bounds') === -1) ? bounds += 'Bounds' : bounds;
    var l = items.length, x = [], y = [], w = [], h = [];
    for (var i = 0; i < l; i++) {
        x.push(items[i][bounds][0]);
        y.push(items[i][bounds][1]);
        w.push(items[i][bounds][2]);
        h.push(items[i][bounds][3]);
    }
    return [Math.min.apply(null, x), Math.max.apply(null, y), Math.max.apply(null, w), Math.min.apply(null, h)];
}

Object.prototype.align = function (preset, userOptions) {
    var obj = LA(this),
        options = {
            bounds: 'geometric',
            margin: '0 0 0 0',
            artboard: activeDocument.artboards.getActiveArtboardIndex(),
            object: { node: undefined, offset: 'outline', bounds: 'geometric' }
        }.extend(userOptions || {}, true);

    options.margin = parseMargin(options.margin);

    function process(item, obj_length) {
        var rect = activeDocument.artboards[options.artboard].artboardRect,
            m = options.margin, distance = 0,
            gb = item.geometricBounds, vb = item.visibleBounds,
            w = item.width / 2, h = item.height / 2,
            offset = options.object.offset.slice(0, 1).toLowerCase();

        if (options.bounds.slice(0, 1).toLowerCase() === 'v') {
            distance = vb[2] - gb[2];
        }

        if (!options.object || !options.object.node) {
            if (preset === 'top') { item.position = [item.position[0], rect[1] - distance - m[0]]; }
            else if (preset === 'right') { item.position = [rect[2] - (w * 2) - distance - m[1], item.position[1]]; }
            else if (preset === 'bottom') { item.position = [item.position[0], rect[3] + (h * 2) + distance + m[2]]; }
            else if (preset === 'left') { item.position = [rect[0] + distance + m[3], item.position[1]]; }
            else if (preset === 'vcenter') { item.position = [item.position[0], (rect[3] + rect[1]) / 2 + h]; }
            else if (preset === 'hcenter') { item.position = [(rect[2] + rect[0]) / 2 - w, item.position[1]]; }
            else if (preset === 'topleft') { item.position = [rect[0] + distance + m[3], rect[1] - distance - m[0]]; }
            else if (preset === 'topcenter') { item.position = [(rect[2] + rect[0]) / 2 - w, rect[1] - distance - m[0]]; }
            else if (preset === 'topright') { item.position = [rect[2] - (w * 2) - distance - m[1], rect[1] - distance - m[0]]; }
            else if (preset === 'middleright') { item.position = [rect[2] - (w * 2) - distance - m[2], (rect[3] + rect[1]) / 2 + h]; }
            else if (preset === 'bottomright') { item.position = [rect[2] - (w * 2) - distance - m[1], rect[3] + (h * 2) + distance + m[2]]; }
            else if (preset === 'bottomcenter') { item.position = [(rect[2] + rect[0]) / 2 - w, rect[3] + (h * 2) + distance + m[2]]; }
            else if (preset === 'bottomleft') { item.position = [rect[0] + distance + m[3], rect[3] + (h * 2) + distance + m[2]]; }
            else if (preset === 'middleleft') { item.position = [rect[0] + distance + m[3], (rect[3] + rect[1]) / 2 + h]; }
            else if (preset === 'center') { item.position = [(rect[2] + rect[0]) / 2 - w, (rect[3] + rect[1]) / 2 + h]; }
        } else {
            // (align-to-object chunk unchanged)
        }
    }

    for (var i = 0; i < obj.length; i++) { process(obj[i], obj.length); }
    return this;
};

Object.prototype.appendTo = function (relativeObject, elementPlacement) {
    var obj = LA(this), i = obj.length;
    if (typeof elementPlacement === 'string' && elementPlacement.length) {
        switch (elementPlacement.toLowerCase()) {
            case 'inside': elementPlacement = 'INSIDE'; break;
            case 'begin': elementPlacement = 'PLACEATBEGINNING'; break;
            case 'end': elementPlacement = 'PLACEATEND'; break;
            case 'before': elementPlacement = 'PLACEBEFORE'; break;
            case 'after': elementPlacement = 'PLACEAFTER'; break;
            default: elementPlacement = '';
        }
    }
    while (i--) {
        if (obj[i].parent !== relativeObject && obj[i] !== relativeObject) {
            if (!elementPlacement) obj[i].moveToBeginning(relativeObject);
            else obj[i].move(relativeObject, ElementPlacement[elementPlacement]);
        }
    }
    return this;
};

Object.prototype.group = function () {
    var obj = LA(this), g = obj[0].parent.groupItems.add();
    this.appendTo(g);
    return g;
};

Object.prototype.ungroup = function () {
    var arr = [], obj = LA(this);
    for (var i = 0; i < obj.length; i++) {
        if (obj[i].typename === 'GroupItem') {
            var j = obj[i].pageItems.length;
            while (j--) {
                arr = arr.concat(obj[i].pageItems[0]);
                obj[i].pageItems[0].moveBefore(obj[i]);
            }
            obj[i].remove();
        } else {
            arr = arr.concat(obj[i]);
        }
    }
    return arr;
};

Object.prototype.getExactSize = function (property, bounds) {
    bounds = bounds || 'geometricBounds';
    bounds = (bounds && bounds.toLowerCase().indexOf('bounds') === -1) ? bounds += 'Bounds' : bounds;
    var values = [],
        item = this,
        b = item[bounds],
        propertyArr = property.replace(/ /g, '').split(','),
        i = propertyArr.length;
    if (b) {
        while (i--) {
            if (propertyArr[i] === 'width') values.push(b[2] - b[0]);
            if (propertyArr[i] === 'height') values.push(b[1] - b[3]);
        }
    }
    return values.length === 1 ? values[0] : values;
};

// -----------------------------------------------------------------------------
//  MAIN SCRIPT OPTIONS + GRIDDER LOGIC
// -----------------------------------------------------------------------------

var scriptName = 'Griddder',
    copyright = ' \u00A9 www.ladyginpro.ru',
    settingFile = {
        name: scriptName + '__setting.json',
        folder: Folder.myDocuments + '/LA_AI_Scripts/'
    },
    $margins = '0',
    $cropMarksColor = {
        type: 'CMYKColor',
        values: [0, 0, 0, 100]
    };

function checkEffects() {
    var testPath = activeDocument.pathItems.add(),
        value = (testPath.applyEffect instanceof Function);
    testPath.remove();
    return value;
}
var effectsEnabled = checkEffects();

Object.prototype.griddder = function (userOptions) {
    try {
        var options = {
            columns: 1,
            rows: 1,
            gutter: { columns: 0, rows: 0 },
            group: 'only_cropmarks',
            align: 'center',
            bounds: 'visible',
            fitToArtboard: false,
            margin: '0mm 0mm 0mm 0mm',
            cropMarks: {
                size: 5,
                offset: 0,
                position: 'relative',
                enabled: false,
                attr: {
                    strokeWidth: 1,
                    strokeColor: { type: 'cmyk', values: [100, 100, 100, 100] }
                }
            }
        }.extend(userOptions || {}, true);

        options.align = options.align.toLowerCase();

        // convert to internal px units (inputs can be in in/mm/cm/pt)
        options.gutter.rows = $.convertUnits(options.gutter.rows, 'px');
        options.gutter.columns = $.convertUnits(options.gutter.columns, 'px');
        options.cropMarks.size = $.convertUnits(options.cropMarks.size, 'px');
        options.cropMarks.offset = $.convertUnits(options.cropMarks.offset, 'px');
        options.cropMarks.attr.strokeWidth = $.convertUnits(options.cropMarks.attr.strokeWidth, 'px');

        var groupItems,
            cropMarksGroup,
            // original selection turned into a flat array
            _rawItems = LA(this),
            // this will hold only “top level” objects (each group counted once)
            items = [],
            collection = [],
            marksCollection = [],
            columns = options.columns,
            rows = options.rows,
            gutter = options.gutter,
            align = options.align,
            fitToArtboard = options.fitToArtboard,
            bounds = options.bounds,
            art = activeDocument.artboards[activeDocument.artboards.getActiveArtboardIndex()],
            margin = parseMargin(options.margin),
            marksPos = (options.cropMarks.position.toLowerCase().slice(0, 1) === 'r' ? options.cropMarks.offset : 0);

        // ---- collapse selection so each GroupItem is treated as ONE item ----
        for (var gi = 0; gi < _rawItems.length; gi++) {
            var it = _rawItems[gi];

            if (it.typename === "GroupItem") {
                // Count the whole group once
                items.push(it);
            } else if (!it.parent || it.parent.typename !== "GroupItem") {
                // Ungrouped art (or top-level stuff) – keep it
                items.push(it);
            }
            // Anything that is a child *inside* a group is ignored here,
            // so it doesn’t inflate the count.
        }

        var groupItems,
            cropMarksGroup,
            marksCollection = [],
            collection = [];

        function createCropMark(coords) {
            var mark = cropMarksGroup.pathItems.add();
            mark.setEntirePath(coords);
            mark.filled = false;
            mark.strokeColor = $.color(options.cropMarks.attr.strokeColor.type, options.cropMarks.attr.strokeColor.values);
            mark.strokeWidth = options.cropMarks.attr.strokeWidth;
            marksCollection.push(mark);
        }

        function getItemMask(obj) {
            var itemMask = obj.getChildsByFilter(function (mask) { return mask.clipping; }, true);
            return itemMask.length ? itemMask[0] : obj;
        }

        items.each(function (item) {

            var itemMask = getItemMask(item),
                itemSize = {
                    width: itemMask.getExactSize('width', bounds) + gutter.columns,
                    height: itemMask.getExactSize('height', bounds) + gutter.rows
                },
                artSize = {
                    width: (art.artboardRect[2] - art.artboardRect[0]) - margin[1] - margin[3],
                    height: (-(art.artboardRect[3] - art.artboardRect[1])) - margin[0] - margin[2]
                },
                isRotate = false,
                group = item.group();

            cropMarksGroup = options.cropMarks.enabled && options.group !== 'none' ? group.groupItems.add() : group;

            if (fitToArtboard) {
                // landscape
                var landscape = {
                    columns: Math.floor(artSize.width / itemSize.width),
                    rows: Math.floor(artSize.height / itemSize.height)
                };
                landscape.columns = (itemSize.width * landscape.columns) > artSize.width ? landscape.columns - 1 : landscape.columns;
                landscape.rows = (itemSize.height * landscape.rows) > artSize.height ? landscape.rows - 1 : landscape.rows;

                // portrait
                var portrait = {
                    columns: Math.floor(artSize.width / itemSize.height),
                    rows: Math.floor(artSize.height / itemSize.width)
                };
                portrait.columns = (itemSize.height * portrait.columns) > artSize.width ? portrait.columns - 1 : portrait.columns;
                portrait.rows = (itemSize.width * portrait.rows) > artSize.height ? portrait.rows - 1 : portrait.rows;

                if (portrait.columns * portrait.rows > landscape.columns * landscape.rows) {
                    isRotate = true;
                    columns = portrait.columns;
                    rows = portrait.rows;
                } else {
                    columns = landscape.columns;
                    rows = landscape.rows;
                }
            }

            if (isRotate) {
                group.rotate(90);
                var wTmp = itemSize.width;
                itemSize.width = itemSize.height;
                itemSize.height = wTmp;
            }

            var useEffects = (effectsEnabled && __useEffect.value);

            for (var ci = 0; ci < columns; ci++) {
                for (var rj = 0; rj < rows; rj++) {
                    var s = (!ci && !rj ? item : (!effectsEnabled || !__useEffect.value ? item.duplicate() : item)),
                        bnds = getItemMask(s)[bounds + 'Bounds'];

                    collection.push(s);

                    if (!effectsEnabled || !__useEffect.value) {
                        s.left += (itemSize.width * ci);
                        s.top -= (itemSize.height * rj);
                    } else if (effectsEnabled && __useEffect.value) {
                        bnds = [
                            bnds[0] + (itemSize.width * ci),
                            bnds[1] - (itemSize.height * rj),
                            bnds[2] + (itemSize.width * ci),
                            bnds[3] - (itemSize.height * rj)
                        ];
                    }

                    // (crop mark creation code here – unchanged from original)
                }
            }

            if (fitToArtboard && !(effectsEnabled && __useEffect.value)) {
                group.align('center', { bounds: 'visible' });
            } else if (!fitToArtboard && typeof options.align === 'string' && options.align.toLowerCase() !== 'none') {
                group.align(options.align);
            }

        }, true);

        return {
            options: options,
            group: groupItems,
            items: collection,
            marks: marksCollection
        };
    } catch (e) {
        $.errorMessage('griddder() - error: ' + e);
    }
};

// -----------------------------------------------------------------------------
//  UI
// -----------------------------------------------------------------------------

function inputNumberEvents(ev, _input, min, max, callback) {
    var step,
        _dir = (ev.keyName ? ev.keyName.toLowerCase().slice(0, 1) : '#none#'),
        _value = parseFloat(_input.text),
        units = ('pt,mm,cm,in,'.indexOf(_input.text.length > 2 ? (',' + _input.text.replace(/ /g, '').slice(-2) + ',') : ',!,') > -1 ?
            _input.text.replace(/ /g, '').slice(-2) : '');

    min = (min === undefined ? 0 : min);
    max = (max === undefined ? Infinity : max);
    step = (ev.shiftKey ? 10 : (ev.ctrlKey ? .1 : 1));

    if (isNaN(_value)) {
        _input.text = min;
    } else {
        _value = ((_dir === 'u') ? _value + step : ((_dir === 'd') ? _value - step : false));
        if (_value !== false) {
            _value = (_value <= min ? min : (_value >= max ? max : _value));
            _input.text = _value;
            if (callback instanceof Function) callback(ev, _value, _input, min, max, units);
            else if (units) _input.text = parseFloat(_input.text) + ' ' + units;
        } else if (units) _input.text = parseFloat(_input.text) + ' ' + units;
    }
}

function doubleGutters(ev, _value, items) {
    if (ev.altKey && ((ev.keyName === 'Left') || (ev.keyName === 'Right'))) {
        doubleValues(_value, items);
    }
}

function doubleValues(val, items) {
    var i = items.length;
    while (i--) { items[i].text = val; }
}

var win = new Window('dialog', scriptName + copyright),
    globalGroup = win.add('group');
globalGroup.orientation = 'column';
globalGroup.alignChildren = ['fill', 'fill'];

var inputSize = [0, 0, 75, 25];

// Vars used across UI + save/load
var __columns, __rows, __columns_gutter, __rows_gutter,
    __align, __bounds, __group,
    __CMEnabled, __cmSize, __cmWeight, __cmOffset, __cmPosition,
    __useEffect,
    cmGroup,
    __totalText, __widthText, __heightText, __countSubgroups;



// GRID PANEL ---------------------------------------------------
var gridPanel = globalGroup.add('panel');
gridPanel.orientation = 'row';
gridPanel.alignChildren = ['fill', 'fill'];

// left column: cols/rows/gutters
var gridLeft = gridPanel.add('group');
gridLeft.orientation = 'column';
gridLeft.alignChildren = ['fill', 'fill'];

// Columns + gutter cols
var rowCols = gridLeft.add('group');
rowCols.orientation = 'row';
rowCols.alignChildren = ['fill', 'fill'];

var colGroup = rowCols.add('group');
colGroup.orientation = 'column';
colGroup.alignChildren = 'fill';
colGroup.add('statictext', undefined, 'Columns');
__columns = colGroup.add('edittext', inputSize, '4');

// still keep keydown for arrow-step + Alt-double behavior
__columns.addEventListener('keydown', function (e) {
    inputNumberEvents(e, __columns, 1, Infinity);
    doubleGutters(e, this.text, [__columns, __rows]);
    updateCount(); // catch Up/Down arrow adjustments
});

// NEW: live update while typing
__columns.onChanging = function () {
    updateCount();
};


var colGutGroup = rowCols.add('group');
colGutGroup.orientation = 'column';
colGutGroup.alignChildren = 'fill';
colGutGroup.add('statictext', undefined, 'Gutter Cols:');
__columns_gutter = colGutGroup.add('edittext', inputSize, '0 in');
__columns_gutter.addEventListener('keydown', function (e) {
    inputNumberEvents(e, __columns_gutter, 0, Infinity);
    doubleGutters(e, this.text, [__columns_gutter, __rows_gutter]);
});

// Rows + gutter rows
var rowRows = gridLeft.add('group');
rowRows.orientation = 'row';
rowRows.alignChildren = ['fill', 'fill'];

var rowGroup = rowRows.add('group');
rowGroup.orientation = 'column';
rowGroup.alignChildren = 'fill';
rowGroup.add('statictext', undefined, 'Rows:');
__rows = rowGroup.add('edittext', inputSize, '4');

// keep keydown for arrow-step + Alt-double
__rows.addEventListener('keydown', function (e) {
    inputNumberEvents(e, __rows, 1, Infinity);
    doubleGutters(e, this.text, [__columns, __rows]);
    updateCount(); // catch Up/Down arrow adjustments
});

// NEW: live update while typing
__rows.onChanging = function () {
    updateCount();
};


var rowGutGroup = rowRows.add('group');
rowGutGroup.orientation = 'column';
rowGutGroup.alignChildren = 'fill';
rowGutGroup.add('statictext', undefined, 'Gutter Rows:');
__rows_gutter = rowGutGroup.add('edittext', inputSize, '0 in');
__rows_gutter.addEventListener('keydown', function (e) {
    inputNumberEvents(e, __rows_gutter, 0, Infinity);
    doubleGutters(e, this.text, [__columns_gutter, __rows_gutter]);
});

// Effects checkbox
__useEffect = gridLeft.add('checkbox', undefined, 'Use Transform effects (if available)');
__useEffect.value = false;

// right column: align/bounds/group
var gridRight = gridPanel.add('group');
gridRight.orientation = 'column';
gridRight.alignChildren = ['fill', 'fill'];

// Align
var gAlign = gridRight.add('group');
gAlign.orientation = 'row';
gAlign.add('statictext', undefined, 'Align:');
__align = gAlign.add('dropdownlist', [0, 0, 140, 25],
    'None,Center,Top Left,Top Center,Top Right,Middle Right,Bottom Right,Bottom Center,Bottom Left,Middle Left'.split(',')
);
__align.selection = 0;

// Bounds
var gBounds = gridRight.add('group');
gBounds.orientation = 'row';
gBounds.add('statictext', undefined, 'Bounds:');
__bounds = gBounds.add('dropdownlist', [0, 0, 110, 25], 'Geometric,Visible'.split(','));
__bounds.selection = 0;

// Group
var gGroup = gridRight.add('group');
gGroup.orientation = 'row';
gGroup.add('statictext', undefined, 'Group:');
__group = gGroup.add('dropdownlist', [0, 0, 180, 25],
    'None,Only Items,Only CropMarks,Items and CropMarks,Items and CropMarks Singly'.split(',')
);
__group.selection = 0;

// CROP MARKS PANEL ---------------------------------------------
var cmPanel = globalGroup.add('panel');
cmPanel.orientation = 'column';
cmPanel.alignChildren = ['fill', 'fill'];

var cmEnableGroup = cmPanel.add('group');
cmEnableGroup.orientation = 'row';
__CMEnabled = cmEnableGroup.add('checkbox', undefined, 'Enable crop marks');
__CMEnabled.value = false;

cmGroup = cmPanel.add('group');
cmGroup.orientation = 'column';
cmGroup.alignChildren = ['fill', 'fill'];
cmGroup.enabled = __CMEnabled.value;

__CMEnabled.onClick = function () { cmGroup.enabled = __CMEnabled.value; };

// first row: size / weight / offset / position
var cmTopRow = cmGroup.add('group');
cmTopRow.orientation = 'row';
cmTopRow.alignChildren = ['fill', 'fill'];

var gSize = cmTopRow.add('group');
gSize.orientation = 'column';
gSize.alignChildren = 'fill';
gSize.add('statictext', undefined, 'Size:');
__cmSize = gSize.add('edittext', [0, 0, 65, 25], '0.25 in'); // 1/4"

var gWeight = cmTopRow.add('group');
gWeight.orientation = 'column';
gWeight.alignChildren = 'fill';
gWeight.add('statictext', undefined, 'Weight:');
__cmWeight = gWeight.add('edittext', [0, 0, 65, 25], '0.01 in'); // ~1pt

var gOffset = cmTopRow.add('group');
gOffset.orientation = 'column';
gOffset.alignChildren = 'fill';
gOffset.add('statictext', undefined, 'Offset:');
__cmOffset = gOffset.add('edittext', [0, 0, 65, 25], '0 in');

var gPos = cmTopRow.add('group');
gPos.orientation = 'column';
gPos.alignChildren = 'fill';
gPos.add('statictext', undefined, 'Position:');
__cmPosition = gPos.add('dropdownlist', [0, 0, 90, 25], 'Absolute,Relative'.split(','));
__cmPosition.selection = 0;

// second row: color
var cmBottomRow = cmGroup.add('group');
cmBottomRow.orientation = 'row';
cmBottomRow.alignChildren = ['fill', 'fill'];

var __cmColorType = cmBottomRow.add('dropdownlist', [0, 0, 170, 25],
    'Color type: CMYK Color,Color type: RGB Color'.split(',')
);
__cmColorType.selection = 0;
__cmColorType.addEventListener('change', function () {
    var newColorType = this.selection.text.replace('Color type: ', '').replace(/ /g, '');
    $cropMarksColor.values = $.convertColor($.color($cropMarksColor.type, $cropMarksColor.values), newColorType).getColorValues();
    $cropMarksColor.type = newColorType;
});

var __cmColorButton = cmBottomRow.add('button', [0, 0, 135, 25], 'Choose color..');
__cmColorButton.onClick = function () {
    var $cropMarksColorNew = app.showColorPicker($.color($cropMarksColor.type, $cropMarksColor.values));
    $cropMarksColor.type = $cropMarksColorNew.typename;
    $cropMarksColor.values = $cropMarksColorNew.getColorValues();
};

// LIVE TOTAL PIECES ---------------------------------------------
var totalGroup = globalGroup.add('group');
totalGroup.orientation = 'row';
totalGroup.alignChildren = ['left', 'center'];
totalGroup.margins = 0;

totalGroup.add('statictext', undefined, 'Total pieces:');
__totalText = totalGroup.add('statictext', undefined, '0');
// LIVE TOTAL WIDTH ----------------------------------------------
var widthGroup = globalGroup.add('group');
widthGroup.orientation = 'row';
widthGroup.alignChildren = ['left', 'center'];
widthGroup.margins = 0;

widthGroup.add('statictext', undefined, 'Total width:');
__widthText = widthGroup.add('statictext', undefined, '0 in');

// Total height
var heightGroup = globalGroup.add('group');
heightGroup.add('statictext', undefined, 'Total height:');
__heightText = heightGroup.add('statictext', undefined, '0 in');


// COUNT MODE (SUBGROUPS) ----------------------------------------
var countModeGroup = globalGroup.add('group');
countModeGroup.orientation = 'row';
countModeGroup.alignChildren = ['left', 'center'];
countModeGroup.margins = 0;

__countSubgroups = countModeGroup.add(
    'checkbox',
    undefined,
    'Count subgroups inside groups'
);
__countSubgroups.value = false;
__countSubgroups.onClick = function () {
    updateCount();
};

// MARGINS + FIT BUTTONS -----------------------------------------
var marginRow = globalGroup.add('group');
marginRow.orientation = 'row';
marginRow.alignChildren = ['fill', 'fill'];
marginRow.margins = 0;

var marginsButton = marginRow.add('button', undefined, 'margins');
marginsButton.onClick = function () {
    $margins = prompt(
        'Enter the margin - top right bottom left.\n' +
        'Units: in (default if unitless), or mm/cm/px.\nSeparator: space',
        $margins
    ).toLowerCase();
};

var fitToArtButton = marginRow.add('button', undefined, 'Fit on artboard');
fitToArtButton.helpTip = 'Fit to artboard';
fitToArtButton.onClick = function () { startAction(true); };

// OK / CANCEL ---------------------------------------------------
var buttonsRow = globalGroup.add('group');
buttonsRow.orientation = 'row';
buttonsRow.alignChildren = ['fill', 'fill'];
buttonsRow.margins = 0;

var cancel = buttonsRow.add('button', undefined, 'Cancel');
cancel.helpTip = 'Press Esc to Close';
cancel.onClick = function () { win.close(); };

var ok = buttonsRow.add('button', undefined, 'OK');
ok.helpTip = 'Press Enter to Run';
ok.onClick = function () { startAction(false); };
ok.active = true;

// DATA + SETTINGS + RUN -----------------------------------------
function getData($fitArtboard) {
    return {
        columns: __columns.text,
        rows: __rows.text,
        gutter: {
            columns: __columns_gutter.text,
            rows: __rows_gutter.text
        },
        group: __group.selection.text.replace(/ /g, '_').toLowerCase(),
        align: __align.selection.text.replace(/ /g, ''),
        bounds: __bounds.selection.text.toLowerCase(),
        fitToArtboard: !!$fitArtboard,
        margin: $margins,
        cropMarks: {
            size: __cmSize.text,
            offset: __cmOffset.text,
            position: __cmPosition.selection.text.toLowerCase(),
            enabled: __CMEnabled.value,
            attr: {
                strokeWidth: __cmWeight.text,
                strokeColor: {
                    type: $cropMarksColor.type,
                    values: $cropMarksColor.values
                }
            }
        }
    };
}

function startAction($fitArtboard) {
    if (!selection || selection.length === 0) {
        alert('Please select at least one object.');
        return;
    }

    var data = getData($fitArtboard);
    var result = selection.griddder(data);

    // optional: pop up a summary at the end
    if (result && result.items && result.items.length) {
   alert('Total pieces generated: ' + total);
    }

    win.close();
}

function saveSettings() {
    var $file = new File(settingFile.folder + settingFile.name),
        data = [
            __columns.text,
            __rows.text,
            __columns_gutter.text,
            __rows_gutter.text,
            __align.selection.index,
            __bounds.selection.index,
            __group.selection.index,
            __CMEnabled.value,
            __cmSize.text,
            __cmWeight.text,
            __cmOffset.text,
            __cmPosition.selection.index,
            __useEffect.value
        ].toString() + '\n' + $cropMarksColor.type + '\n' + $cropMarksColor.values.toString() + '\n' + $margins;

    $file.open('w');
    $file.write(data);
    $file.close();
}

function loadSettings() {
    var $file = File(settingFile.folder + settingFile.name);
    if ($file.exists) {
        try {
            $file.open('r');
            var data = $file.read().split('\n'),
                $main = data[0].split(','),
                $cmColorType = data[1],
                $cmColorValues = data[2].split(','),
                $mrgns = data[3];
            __columns.text = $main[0];
            __rows.text = $main[1];
            __columns_gutter.text = $main[2];
            __rows_gutter.text = $main[3];
            __align.selection = parseInt($main[4], 10);
            __bounds.selection = parseInt($main[5], 10);
            __group.selection = parseInt($main[6], 10);
            __CMEnabled.value = ($main[7] === 'true');
            __cmSize.text = $main[8];
            __cmWeight.text = $main[9];
            __cmOffset.text = $main[10];
            __cmPosition.selection = parseInt($main[11], 10);
            __useEffect.value = ($main[12] === 'true');

            $cropMarksColor.type = $cmColorType;
            $cropMarksColor.values = $cmColorValues;
            $margins = $mrgns;
            cmGroup.enabled = __CMEnabled.value;
        } catch (e) { }
        $file.close();
    }
}
function updateCount() {
    if (!__totalText) return;

    var cols = parseInt(__columns.text, 10),
        rows = parseInt(__rows.text, 10);

    if (isNaN(cols) || cols < 1) cols = 0;
    if (isNaN(rows) || rows < 1) rows = 0;

    var countSubgroups = (__countSubgroups && __countSubgroups.value);

    // selection-aware counting
    var selCount = 0;
    var sampleItem = null;

    if (selection && selection.length) {
        for (var i = 0; i < selection.length; i++) {
            var it = selection[i];

            if (countSubgroups && it.typename === "GroupItem") {
                // count children inside the group
                var kids = it.pageItems;
                for (var j = 0; j < kids.length; j++) {
                    var child = kids[j];
                    if (child.clipping) continue; // skip clipping masks

                    selCount++;
                    if (!sampleItem) sampleItem = child;
                }
            } else {
                // normal mode: count top-level items
                if (it.typename === "GroupItem") {
                    selCount++;
                    if (!sampleItem) sampleItem = it;
                } else if (!it.parent || it.parent.typename !== "GroupItem") {
                    selCount++;
                    if (!sampleItem) sampleItem = it;
                }
            }
        }
    }

    // total pieces
    var total = cols * rows * selCount;
    __totalText.text = String(total || 0);

    // ----- TOTAL WIDTH -----
    var widthText = "0 in";

    if (cols > 0 && sampleItem) {
        var gb = sampleItem.geometricBounds; // [L, T, R, B] in points
        var itemWidthPt = gb[2] - gb[0];

        var gutterColPt = 0;
        if (__columns_gutter && __columns_gutter.text) {
            gutterColPt = $.convertUnits(__columns_gutter.text, 'px');
        }

        var totalWidthPt = (cols * itemWidthPt) + ((cols - 1) * gutterColPt);
        var totalWidthIn = totalWidthPt / 72; // 72 pt per inch
        widthText = totalWidthIn.toFixed(3) + " in";
    }

    if (__widthText) {
        __widthText.text = widthText;
    }

    // ----- TOTAL HEIGHT -----
    var heightText = "0 in";

    if (rows > 0 && sampleItem) {
        var gbH = sampleItem.geometricBounds; // [L, T, R, B]
        var itemHeightPt = Math.abs(gbH[1] - gbH[3]); // top - bottom

        var gutterRowPt = 0;
        if (__rows_gutter && __rows_gutter.text) {
            gutterRowPt = $.convertUnits(__rows_gutter.text, 'px');
        }

        var totalHeightPt = (rows * itemHeightPt) + ((rows - 1) * gutterRowPt);
        var totalHeightIn = totalHeightPt / 72;
        heightText = totalHeightIn.toFixed(3) + " in";
    }

    if (__heightText) {
        __heightText.text = heightText;
    }

    // if you're using this for the popup, you can also stash total here:
    // _globalTotalPieces = total;
}


function checkSettingFolder() {
    var $folder = new Folder(settingFile.folder);
    if (!$folder.exists) $folder.create();
}



checkSettingFolder();
loadSettings();
updateCount();

win.center();
win.show();
