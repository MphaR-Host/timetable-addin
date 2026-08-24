/* Task pane controller: wires the Excel layer to the calendar renderer, and
   acts as the RPC host for the pop-out "Big view" dialog (a dialog cannot touch
   the Excel object model itself, so it asks the pane to do it). */
(function () {
  "use strict";

  var cfg = null, items = [], available = [], warning = "", booted = false;
  var srcSheets = [], bigDialog = null, refreshing = false, pending = 0;

  var FIELD_LABEL = { id:"ID", name:"Task", start:"Start", end:"End",
    status:"Status", color:"Color", key:"Key", async:"Async", notes:"Notes" };
  var GUESS = {
    id:["id","uid","ref","key","no","number","č.","cislo"],
    name:["task","name","title","milestone","activity","činnost","cinnost","úkol","ukol","description","deliverable","step"],
    start:["start","start date","from","begin","od","začátek","zacatek","planned start"],
    end:["end","end date","finish","to","due","deadline","do","konec","planned end","planned finish"],
    status:["status","state","stav","progress","phase"],
    color:["color","colour","barva","legend"],
    key:["key","milestone","key milestone","critical","important"],
    async:["async","asynchronous","weekend","continuous","calendar days"],
    notes:["notes","note","comment","comments","remark","poznámka","poznamka","detail"]
  };

  function $(id){ return document.getElementById(id); }
  var toastT=null;
  function toast(msg,bad){
    var t=$("toast"); t.textContent=msg; t.className="toast show"+(bad?" bad":"");
    clearTimeout(toastT); toastT=setTimeout(function(){ t.className="toast"+(bad?" bad":""); },3600);
  }
  function showErr(msg){
    var b=$("errBanner");
    if(msg){ b.style.display="block"; b.textContent="⚠ "+msg; }
    else b.style.display="none";
  }

  /* ---------- boot ---------- */
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      document.body.innerHTML = "<p class='pad'>This add-in runs in Excel.</p>";
      return;
    }
    cfg = Xl.loadConfig();
    $("density").value = cfg.density || "comfortable";
    $("showDates").checked = !!cfg.showDates;
    bind();
    if (!cfg.sheet || !cfg.mapping || !cfg.mapping.name) {
      $("setupBanner").style.display = "block";
      openCfg();
    } else {
      refresh().then(startWatching);
    }
    booted = true;
  });

  function startWatching(){
    if(!cfg.sheet) return;
    Xl.watch(cfg, function(){ if(!pending) refresh(); }).catch(function(){});
  }

  /* ---------- data ---------- */
  function refresh(){
    if(!cfg.sheet || !cfg.mapping.name) return Promise.resolve();
    if(refreshing) return Promise.resolve();
    refreshing = true;
    return Xl.read(cfg).then(function(res){
      items = res.items; available = res.available; warning = res.warning;
      showErr("");
      $("setupBanner").style.display = "none";
      render();
      pushToDialog();
    }).catch(function(e){
      showErr(e.message || String(e));
    }).then(function(){ refreshing = false; });
  }

  function act(p){
    pending++;
    return p.then(function(r){ pending--; return refresh().then(function(){ return r; }); })
            .catch(function(e){ pending--; toast(e.message||String(e), true); throw e; });
  }

  /* ---------- undo -----------------------------------------------------
     Office.js edits don't reliably land on Excel's Ctrl+Z stack, so we keep
     our own. Each mutation captures how to reverse itself, from the last-known
     item state, before applying. The inverse re-uses the same apply* helpers
     with record=false so undoing doesn't itself pile onto the stack. */
  var undoStack = [];
  function itemById(id){ for (var i=0;i<items.length;i++) if (items[i].id===id) return items[i]; return null; }
  function pushUndo(fn){ undoStack.push(fn); if (undoStack.length>40) undoStack.shift(); refreshUndoBtn(); }
  function refreshUndoBtn(){
    var b=$("undoBtn"); if(b) b.disabled = undoStack.length===0;
  }
  function doUndo(){ var fn=undoStack.pop(); refreshUndoBtn(); if(fn) fn(); else toast("Nothing to undo"); }

  function applyUpdate(id, field, value, record){
    if (record){ var it=itemById(id); var old = it ? it[field] : ""; pushUndo(function(){ applyUpdate(id, field, old, false); }); }
    act(field==="color" ? Xl.setColor(cfg, id, value) : Xl.update(cfg, id, field, value)).catch(function(){});
  }
  function applySetDates(id, start, end, record){
    if (record){ var it=itemById(id); var os=it?it.start:"", oe=it?(it.end||""):""; pushUndo(function(){ applySetDates(id, os, oe, false); }); }
    act(Xl.setDates(cfg, id, start, end)).catch(function(){});
  }
  function applyAdd(start, end, name, record){
    act(Xl.add(cfg, start, end, name)).then(function(r){
      if (record && r && r.id) pushUndo(function(){ act(Xl.remove(cfg, r.id)).catch(function(){}); });
    }).catch(function(){});
  }
  function applyDelete(id, record){
    if (record){ var it=itemById(id); if(it){ var snap=JSON.parse(JSON.stringify(it)); pushUndo(function(){ act(Xl.addFull(cfg, snap)).catch(function(){}); }); } }
    act(Xl.remove(cfg, id)).catch(function(){});
  }

  /* ---------- render ---------- */
  function render(){
    $("ttl").textContent = (cfg.sheet||"Timetable") + (cfg.table?(" · "+cfg.table):"");
    var wb=$("warnBanner");
    if(warning){ wb.style.display="block"; wb.textContent="⚠ "+warning; }
    else wb.style.display="none";
    $("skey").style.display = items.length ? "flex" : "none";
    if(items.length) Cal.statusKey($("skey"), items);
    renderTable();
    renderFilters();
    Cal.render($("calendar"), items, {
      hidden: cfg.hidden||[], density: cfg.density,
      perRow: "auto", showDates: cfg.showDates
    });
  }

  function renderTable(){
    var tb=$("mbody");
    if(tb.contains(document.activeElement)) return;   // never interrupt typing
    tb.innerHTML="";
    if(!items.length){
      tb.innerHTML='<p class="note">No rows yet. Click “+ Add task”, or type straight into the sheet.</p>';
      return;
    }
    var can=function(f){ return available.indexOf(f)>=0; };
    items.forEach(function(it){
      var cd = it.start ? countdown(it) : {txt:"",soon:false};
      var div=document.createElement("div");
      div.className="trow"+(it.status==="Done"?" done":"");
      div.dataset.id=it.id;
      var f=[];
      if(can("color")) f.push('<label class="fld"><span>Color</span><input type="color" data-f="color" value="'+it.color+'"></label>');
      f.push('<label class="fld"><span>Start</span><input type="date" data-f="start" value="'+(it.start||"")+'"></label>');
      if(can("end")) f.push('<label class="fld"><span>End</span><input type="date" data-f="end" value="'+(it.end||"")+'"></label>');
      if(can("status")) f.push('<label class="fld"><span>Status</span><select data-f="status">'+
        Cal.STATUSES.map(function(s){ return '<option '+(s===it.status?"selected":"")+">"+s+"</option>"; }).join("")+"</select></label>");
      if(can("key")) f.push('<label class="keybox"><input type="checkbox" data-f="key" '+(it.key?"checked":"")+">★ key</label>");
      if(can("async")) f.push('<label class="keybox"><input type="checkbox" data-f="async" '+(it.async?"checked":"")+">async</label>");
      div.innerHTML=
        '<div class="r1"><button class="xl" data-act="select" title="Select this row in Excel">⤢</button>'+
        '<input class="name" data-f="name" value="'+Cal.esc(it.name)+'" placeholder="Task / milestone…">'+
        '<span class="cd '+(cd.soon?"soon":"")+'">'+cd.txt+"</span>"+
        '<button class="btn sm danger" data-act="del" title="Delete row">✕</button></div>'+
        '<div class="r2">'+f.join("")+"</div>";
      tb.appendChild(div);
    });
    bindRows();
  }

  function countdown(it){
    var diff=Cal.daysBetween(Cal.today0(), Cal.iStart(it));
    if(it.status==="Done") return {txt:"Done",soon:false};
    return { txt: diff<0 ? (-diff)+"d ago" : diff===0 ? "today" : "in "+diff+"d",
             soon: diff>=0 && diff<=7 };
  }

  function bindRows(){
    var tb=$("mbody"), typeT=null;
    tb.querySelectorAll("input,select").forEach(function(el){
      var id=el.closest(".trow").dataset.id, f=el.dataset.f;
      function push(v){ applyUpdate(id,f,v,true); }
      if(el.type==="checkbox") el.addEventListener("change",function(){ push(el.checked); });
      else if(el.type==="color") el.addEventListener("change",function(){ push(el.value); });
      else if(el.classList.contains("name")){
        el.addEventListener("input",function(){ clearTimeout(typeT);
          typeT=setTimeout(function(){ push(el.value); },500); });
        el.addEventListener("blur",function(){ clearTimeout(typeT); push(el.value); });
      } else el.addEventListener("change",function(){ push(el.value); });
    });
    tb.querySelectorAll("[data-act]").forEach(function(b){
      b.addEventListener("click",function(e){
        var id=e.target.closest(".trow").dataset.id;
        if(e.target.dataset.act==="del") applyDelete(id,true);
        else Xl.selectRow(cfg,id).catch(function(){});
      });
    });
  }

  function renderFilters(){
    var sf=$("statusFilters"); sf.innerHTML="";
    Cal.STATUSES.forEach(function(s){
      var hidden=(cfg.hidden||[]).indexOf(s)>=0;
      var l=document.createElement("label"); l.className="chk";
      l.innerHTML='<input type="checkbox" '+(hidden?"":"checked")+">"+s;
      l.querySelector("input").addEventListener("change",function(ev){
        cfg.hidden = ev.target.checked
          ? (cfg.hidden||[]).filter(function(x){ return x!==s; })
          : (cfg.hidden||[]).concat(s);
        Xl.saveConfig(cfg); render(); pushToDialog();
      });
      sf.appendChild(l);
    });
  }

  /* clicking anything in the calendar selects that row in Excel — unless a drag
     just happened (move/resize), in which case swallow the trailing click. */
  document.addEventListener("click", function(e){
    if(Cal._suppressClick){ Cal._suppressClick=false; e.stopPropagation(); return; }
    var el=e.target.closest("[data-jump]");
    if(el && cfg && cfg.sheet){ e.stopPropagation(); Xl.selectRow(cfg,el.dataset.jump).catch(function(){}); }
  }, true);

  /* drag to move/resize; right-click a box to add/edit/delete its text */
  Cal.enableDrag($("calendar"), {
    canResize: function(){ return available.indexOf("end") >= 0; },
    onEdit:   function(id, start, end){ applySetDates(id, start, end, true); },
    onRename: function(id, name){ applyUpdate(id, "name", name, true); },
    onSelect: function(id){ Xl.selectRow(cfg, id).catch(function(){}); },
    onAdd:    function(date, name){ applyAdd(date, "", name, true); },
    onCreate: function(start, end){ applyAdd(start, end, "", true); },
    onDelete: function(id){ applyDelete(id, true); }
  });

  /* ---------- source & mapping ---------- */
  function openCfg(){
    $("cfgModal").classList.add("open");
    Xl.listSources().then(function(res){
      srcSheets=res.sheets;
      var sh=$("cfgSheet");
      sh.innerHTML=srcSheets.map(function(s,i){ return '<option value="'+i+'">'+Cal.esc(s.name)+"</option>"; }).join("");
      var pick=0;
      srcSheets.forEach(function(s,i){ if(s.name===cfg.sheet) pick=i; });
      sh.value=String(pick);
      fillTables();
    }).catch(function(e){ showErr(e.message||String(e)); });
  }
  function closeCfg(){ $("cfgModal").classList.remove("open"); }
  function curSheet(){ return srcSheets[+$("cfgSheet").value] || null; }

  function fillTables(){
    var s=curSheet(), tb=$("cfgTable");
    if(!s){ tb.innerHTML=""; return; }
    var opts=s.tables.map(function(t){
      return '<option value="'+Cal.esc(t.name)+'">Table “'+Cal.esc(t.name)+"” — "+
             t.headers.length+" cols, "+t.rows+" rows</option>"; });
    (s.anchors||[]).forEach(function(a){
      opts.push('<option value="'+Cal.esc(a.cell)+'">Range at '+Cal.esc(a.cell)+" — "+
        a.cols+" cols ("+Cal.esc(a.headers.slice(0,3).join(", "))+"…)</option>"); });
    tb.innerHTML=opts.join("") || '<option value="A1">Plain range from A1</option>';
    if(cfg.table && Array.prototype.some.call(tb.options,function(o){ return o.value===cfg.table; }))
      tb.value=cfg.table;
    fillMaps();
  }
  function curHeaders(){
    var s=curSheet(); if(!s) return [];
    var v=$("cfgTable").value;
    var t=s.tables.filter(function(x){ return x.name===v; })[0];
    if(t) return t.headers;
    var a=(s.anchors||[]).filter(function(x){ return x.cell===v; })[0];
    return a?a.headers:[];
  }
  function fillMaps(){
    var hdrs=curHeaders(), box=$("cfgMaps");
    var same = curSheet() && curSheet().name===cfg.sheet && $("cfgTable").value===cfg.table;
    box.innerHTML=Xl.FIELDS.map(function(f){
      var req=Xl.REQUIRED.indexOf(f)>=0, chosen="";
      if(same && cfg.mapping[f] && hdrs.indexOf(cfg.mapping[f])>=0) chosen=cfg.mapping[f];
      if(!chosen){
        var g=GUESS[f]||[];
        var hit=hdrs.filter(function(h){ return g.indexOf(String(h).toLowerCase().trim())>=0; })[0];
        if(!hit) hit=hdrs.filter(function(h){
          return g.some(function(x){ return String(h).toLowerCase().trim().indexOf(x)===0; }); })[0];
        if(hit) chosen=hit;
      }
      var opts=['<option value="">— none —</option>'].concat(hdrs.map(function(h){
        return '<option value="'+Cal.esc(h)+'" '+(h===chosen?"selected":"")+">"+Cal.esc(h)+"</option>"; }));
      return '<div class="maprow"><label>'+FIELD_LABEL[f]+(req?' <span class="req">*</span>':"")+
             '</label><select data-map="'+f+'">'+opts.join("")+"</select></div>";
    }).join("");
  }

  function saveCfg(){
    var s=curSheet();
    if(!s){ toast("Pick a sheet first",true); return; }
    var mapping={};
    document.querySelectorAll("#cfgMaps select").forEach(function(sel){
      mapping[sel.dataset.map]=sel.value; });
    for(var i=0;i<Xl.REQUIRED.length;i++){
      var f=Xl.REQUIRED[i];
      if(!mapping[f]){ toast(FIELD_LABEL[f]+" must be mapped to a column",true); return; }
    }
    cfg.sheet=s.name; cfg.table=$("cfgTable").value; cfg.mapping=mapping;
    Xl.saveConfig(cfg).then(function(){
      closeCfg(); toast("Source saved");
      return refresh().then(startWatching);
    }).catch(function(e){ toast(e.message||String(e),true); });
  }

  /* ---------- big view dialog (RPC host) ---------- */
  function openBig(){
    var url=window.location.href.replace(/taskpane\.html.*$/,"dialog.html");
    Office.context.ui.displayDialogAsync(url,{height:88,width:88,promptBeforeOpen:false},
      function(res){
        if(res.status!==Office.AsyncResultStatus.Succeeded){
          toast("Could not open the big view: "+(res.error?res.error.message:""),true); return;
        }
        bigDialog=res.value;
        bigDialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogMessage);
        bigDialog.addEventHandler(Office.EventType.DialogEventReceived, function(){ bigDialog=null; });
      });
  }
  function replyToDialog(obj){
    if(!bigDialog) return;
    try { bigDialog.messageChild(JSON.stringify(obj)); } catch(e){}
  }
  function pushToDialog(){
    var design=null; try{ design=Xl.loadDesign(cfg.sheet); }catch(e){}
    replyToDialog({ op:"data", items:items, available:available,
                    density:cfg.density, showDates:cfg.showDates,
                    hidden:cfg.hidden||[], title:cfg.sheet, design:design });
  }
  function onDialogMessage(arg){
    var m;
    try { m=JSON.parse(arg.message); } catch(e){ return; }
    if(m.op==="ready"){ pushToDialog(); return; }
    if(m.op==="select"){ Xl.selectRow(cfg,m.id).catch(function(){}); return; }
    if(m.op==="update"){ applyUpdate(m.id,m.field,m.value,true); return; }
    if(m.op==="setdates"){ applySetDates(m.id,m.start,m.end,true); return; }
    if(m.op==="savedesign"){ Xl.saveDesign(m.key, m.json)
        .then(function(){ replyToDialog({ op:"toast", msg:"Design saved to workbook" }); })
        .catch(function(e){ replyToDialog({ op:"toast", msg:"Save failed: "+(e.message||e) }); }); return; }
    if(m.op==="add"){ applyAdd(m.start,m.end||"",m.name,true); return; }
    if(m.op==="delete"){ applyDelete(m.id,true); return; }
    if(m.op==="undo"){ doUndo(); return; }
    if(m.op==="prefs"){
      if(m.density) cfg.density=m.density;
      if(m.showDates!==undefined) cfg.showDates=m.showDates;
      Xl.saveConfig(cfg); $("density").value=cfg.density;
      $("showDates").checked=!!cfg.showDates; render(); return;
    }
  }

  /* ---------- wiring ---------- */
  function bind(){
    $("cfgBtn").addEventListener("click",openCfg);
    $("cfgCancel").addEventListener("click",closeCfg);
    $("cfgSave").addEventListener("click",saveCfg);
    $("cfgSheet").addEventListener("change",fillTables);
    $("cfgTable").addEventListener("change",fillMaps);
    $("cfgModal").addEventListener("click",function(e){ if(e.target.id==="cfgModal") closeCfg(); });
    $("anchorBtn").addEventListener("click",function(){
      var cell=$("cfgAnchor").value.trim().toUpperCase(), s=curSheet();
      if(!cell||!s) return;
      if(!Xl.isCellRef(cell)){ toast("That is not a cell reference, e.g. B7",true); return; }
      Xl.headersAt(s.name,cell).then(function(r){
        if(!r.headers.length){ toast("No header text found at "+cell,true); return; }
        s.anchors=(s.anchors||[]).filter(function(a){ return a.cell!==cell; });
        s.anchors.push({cell:cell,headers:r.headers,cols:r.headers.length});
        fillTables(); $("cfgTable").value=cell; fillMaps();
        toast("Found "+r.headers.length+" columns at "+cell);
      }).catch(function(e){ toast(e.message||String(e),true); });
    });
    $("addBtn").addEventListener("click",function(){ applyAdd("", "", "", true); });
    $("undoBtn").addEventListener("click",function(){ doUndo(); });
    $("sortBtn").addEventListener("click",function(){
      act(Xl.sortByDate(cfg)).then(function(){ toast("Rows sorted by start date"); }).catch(function(){}); });
    $("reloadBtn").addEventListener("click",function(){ refresh(); });
    $("bigBtn").addEventListener("click",openBig);
    ["density","showDates"].forEach(function(id){
      $(id).addEventListener("change",function(){
        cfg.density=$("density").value; cfg.showDates=$("showDates").checked;
        Xl.saveConfig(cfg); render(); pushToDialog();
      });
    });
  }
})();
