#target illustrator
(function () {
    // ---------- helpers ----------
    function trim(s){ return (s==null?"":String(s)).replace(/^\s+|\s+$/g,""); }
    function r6(n){ return Math.round(n*1e6)/1e6; }

    function toPts(val, unit){
        unit = (unit||"in").toLowerCase(); // default = inches
        if (unit=='"' || unit==="in") return val*72;
        if (unit==="pt") return val;
        if (unit==="mm") return val*72/25.4;
        if (unit==="cm") return val*72/2.54;
        if (unit==="px") return val;
        return val*72;
    }
    function fromPts(pts, unit){
        if (unit==="pt") return pts;
        if (unit==="in") return pts/72;
        if (unit==="mm") return pts*25.4/72;
        if (unit==="cm") return pts*2.54/72;
        return pts/72;
    }
    function unionBounds(items, useVisible){
        var b=null, bb;
        for(var i=0;i<items.length;i++){
            try{ bb = useVisible ? items[i].visibleBounds : items[i].geometricBounds; }catch(e){ bb=null; }
            if(!bb) continue;
            if(!b) b = bb.slice();
            else{
                b[0]=Math.min(b[0],bb[0]);
                b[1]=Math.max(b[1],bb[1]);
                b[2]=Math.max(b[2],bb[2]);
                b[3]=Math.min(b[3],bb[3]);
            }
        }
        return b;
    }
    function copyToClipboard(doc, text){
        var tf = doc.textFrames.add(); tf.contents = String(text);
        tf.left = 100000; tf.top = 100000;
        doc.selection = null; tf.selected = true;
        app.executeMenuCommand("copy");
        tf.remove();
    }

    // ---------- get items ----------
    function promoteToItem(node){
        var p=node, steps=0;
        while(p && steps<8){
            try{ if(p.typename && /Item$/.test(p.typename)) return p; p=p.parent; }catch(_){ p=null; }
            steps++;
        }
        return null;
    }
    function getItems(doc){
        var items=[];
        if (app.selection && app.selection.length>0){
            for (var i=0;i<app.selection.length;i++){
                var pi=promoteToItem(app.selection[i]);
                if(pi) items.push(pi);
            }
        }
        if (items.length===0){
            var all=doc.pageItems;
            for (var j=0;j<all.length;j++){
                var it=all[j];
                if (it.locked||it.hidden) continue;
                var lay=it.layer; if (lay && !lay.visible) continue;
                items.push(it);
            }
        }
        return items;
    }

    // ---------- guards ----------
    if (app.documents.length===0){ alert("Open a document."); return; }
    var doc=app.activeDocument;
    var items=getItems(doc);
    if (items.length===0){ alert("No objects to measure."); return; }

    var bb = unionBounds(items, false);
    if(!bb){ alert("Could not read bounds."); return; }
    var curW = bb[2]-bb[0], curH = bb[1]-bb[3];

    // ---------- detect oversized artboard ----------
    var ab=doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    var abW=ab[2]-ab[0], abH=ab[1]-ab[3];
    var defaultRatio = (abW>14400 || abH>14400) ? 10 : 1; // 200 in = 14400 pt

    // ---------- dialog ----------
    var d = new Window("dialog","Quick Scale (Smart Inches)");
    d.orientation="column"; d.alignChildren=["fill","top"]; d.margins=12; d.spacing=8;

    var cur = d.add("panel", undefined, "Current Size");
    cur.orientation="column"; cur.margins=8;
    cur.add("statictext", undefined, "Width:  " + r6(fromPts(curW,"in")) + " in  ("+r6(fromPts(curW,"mm"))+" mm)");
    cur.add("statictext", undefined, "Height: " + r6(fromPts(curH,"in")) + " in  ("+r6(fromPts(curH,"mm"))+" mm)");

    var tgt = d.add("panel", undefined, "Target");
    tgt.orientation="row"; tgt.margins=8; tgt.spacing=8;
    var et = tgt.add("edittext", undefined, "52 in"); et.characters=18;
    var axisGrp = tgt.add("group");
    var rbW = axisGrp.add("radiobutton", undefined, "Width"); rbW.value=true;
    var rbH = axisGrp.add("radiobutton", undefined, "Height");

    var preset = d.add("panel", undefined, "Scale Preset");
    preset.orientation="row"; preset.margins=8; preset.spacing=8;
    var rbFull=preset.add("radiobutton",undefined,"Full (1:1)");
    var rbTen=preset.add("radiobutton",undefined,"1/10 (0.1)");
    var rbCustom=preset.add("radiobutton",undefined,"Custom (1:x)");
    var etCustom=preset.add("edittext",undefined,"20"); etCustom.characters=5; etCustom.enabled=false;

    if (defaultRatio==1) rbFull.value=true;
    if (defaultRatio==10) rbTen.value=true;

    rbCustom.onClick=function(){ etCustom.enabled=true; };
    rbFull.onClick=function(){ etCustom.enabled=false; };
    rbTen.onClick=function(){ etCustom.enabled=false; };

    var btns = d.add("group"); btns.alignment="right";
    var ok = btns.add("button", undefined, "Apply", {name:"ok"});
    var justCopy = btns.add("button", undefined, "Just Copy");
    var cancel = btns.add("button", undefined, "Cancel", {name:"cancel"});

    // ---------- compute function ----------
    function compute(){
        var input=trim(et.text);
        if(!input){ throw Error("Enter a target like 52 in, 300 mm, *2, /2."); }

        var curSize = rbH.value ? curH : curW;
        var targetPts=null, factor=null;

        if (/^[\*\+\-\/]/.test(input)){
            var op=input.charAt(0), rest=trim(input.substring(1));
            if(op==='*'||op==='/'){
                var f=parseFloat(rest);
                if(!(f>0)) throw Error("Use *N or /N.");
                factor=(op==='*')?f:(1/f);
                targetPts=curSize*factor;
            } else {
                var m=rest.match(/^([0-9.]+)\s*([a-z"]+)?$/i);
                if(!m) throw Error("Use +N unit or -N unit.");
                var delta=toPts(parseFloat(m[1]),(m[2]||"in"));
                targetPts=(op==='+')?(curSize+delta):(curSize-delta);
                factor=targetPts/curSize;
            }
        } else {
            var m2=input.match(/^([0-9.]+)\s*([a-z"]+)?$/i);
            if(!m2) throw Error("Enter like 52, 52 in, 300 mm");
            var n=parseFloat(m2[1]), unit=(m2[2]||"in").toLowerCase();
            if(unit==='"') unit="in";
            targetPts=toPts(n,unit);
            factor=targetPts/curSize;
        }

        var ratio=1;
        if(rbTen.value) ratio=10;
        else if(rbCustom.value){
            var r=parseFloat(etCustom.text);
            if(!(r>0)) throw Error("Custom must be >0.");
            ratio=r;
        }
        var docTarget=targetPts/ratio;
        factor=docTarget/curSize;
        var percent=factor*100;

        return {factor:factor, percent:percent};
    }

    ok.onClick = function(){
        try{
            var res=compute();
            var g=doc.groupItems.add();
            for(var i=items.length-1;i>=0;i--){ try{ items[i].move(g,ElementPlacement.PLACEATBEGINNING); }catch(_){ } }
            doc.selection=null; g.selected=true;
            g.resize(res.percent,res.percent,true,true,true,true,true);
            app.executeMenuCommand("ungroup");

            var out="Percent: "+r6(res.percent)+"\nFactor: "+r6(res.factor);
            copyToClipboard(doc,out);
            alert("Scaled.\n"+out+"\n\n(Copied to clipboard.)");
            d.close();
        }catch(e){ alert(e.message); }
    };

    justCopy.onClick = function(){
        try{
            var res=compute();
            // Only copy the numeric percent
            var out = r6(res.percent);
            copyToClipboard(doc,out);
            alert("Copied only.\nScale: " + out + "%");
            d.close();
        }catch(e){ alert(e.message); }
    };

    d.show();
})();
