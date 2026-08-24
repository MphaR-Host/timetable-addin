/* Annotated timeline view.
 *
 * A second way to look at the same items: compact month grids with solid
 * colour-per-task day cells, and the task descriptions living OUTSIDE the grid
 * as draggable/resizable label boxes tied to their days by colour arrows the
 * user can bend and re-anchor. Built for exporting a slide-ready PNG.
 *
 * It knows nothing about Excel: dragging a coloured run to reschedule calls
 * opts.onEdit(id, startISO, endISO) and the host persists it — the same
 * contract the calendar view uses. Classes are namespaced `an-`.
 */
(function (global) {
  "use strict";

  var MON = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var DOW = ["Mo","Tu","We","Th","Fr","Sa","Su"];
  var CW = 46, CH = 34;
  var HDR = {0:"#c47fb0",1:"#5aa0d6",2:"#7fb069",3:"#e0a13c",4:"#5aa0d6",5:"#7fb069",6:"#5aa0d6",7:"#d9736b",8:"#e0b93c",9:"#7fb069",10:"#b08bd6",11:"#5aa0d6"};
  var PALETTE = ["#2e86c1","#e67e22","#27ae60","#8e44ad","#c0392b","#16a085","#d4ac0d","#34495e","#d81b83","#2980b9","#af601a","#1abc9c"];
  // A wider, well-spaced palette used to disambiguate tasks that arrive with the
  // same colour (e.g. a sheet with no Colour column, where the host assigns a
  // small repeating palette by row) so every task in the view is distinct.
  var DISTINCT = ["#2e86c1","#e67e22","#27ae60","#8e44ad","#c0392b","#16a085","#d4ac0d","#34495e","#d81b83","#2980b9",
                  "#af601a","#1abc9c","#7f8c8d","#c0398b","#0e6655","#b7950b","#6c3483","#1f618d","#a04000","#117864"];
  /* colour per task id: keep each incoming colour when unique, otherwise take the
     next unused DISTINCT colour, so no two visible tasks share a fill. */
  function colorMap(items){
    var map={}, used={}, di=0;
    items.forEach(function(t){
      var ov=S.colorOverride[t.id];                 // an explicit right-click colour always wins
      if(ov){ ov=ov.toLowerCase(); used[ov]=1; map[t.id]=ov; return; }
      var c=(t.color||"").toLowerCase();
      if(!c || used[c]){ while(used[DISTINCT[di%DISTINCT.length].toLowerCase()] && di<DISTINCT.length*3) di++;
        c=DISTINCT[di%DISTINCT.length].toLowerCase(); di++; }
      used[c]=1; map[t.id]=c;
    });
    return map;
  }
  var PRESETS = { "16:9":[1600,900], "A4":[1754,1240] };

  function parse(s){ if(!s) return null; var p=String(s).split("-").map(Number); if(!p[0]) return null; return new Date(p[0],p[1]-1,p[2]); }
  function iso(d){ return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0"); }
  function todayISO(){ var t=new Date(); return iso(new Date(t.getFullYear(),t.getMonth(),t.getDate())); }
  function midx(d){ return (d.getDay()+6)%7; }
  function addD(d,n){ var x=new Date(d); x.setDate(x.getDate()+n); return x; }
  function daysBetween(a,b){ return Math.round((b-a)/86400000); }
  function esc(s){ return String(s==null?"":s).replace(/[&<>]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;"}[c];}); }
  function fmtNice(d){ return d.getDate()+" "+MON[d.getMonth()].slice(0,3); }
  function isMile(t){ return !t.end || t.end===t.start; }
  function tStart(t){ return parse(t.start); }
  function tEnd(t){ var e=t.end?parse(t.end):null, s=tStart(t); return (e&&e>s)?e:s; }
  function colr(t,i){ return t.color || PALETTE[i%PALETTE.length]; }
  function ink(hex){ var h=String(hex||"#2e86c1").replace("#",""); var r=parseInt(h.slice(0,2),16)||0,g=parseInt(h.slice(2,4),16)||0,b=parseInt(h.slice(4,6),16)||0;
    return (0.299*r+0.587*g+0.114*b)>150?"#22303f":"#ffffff"; }

  var S = { board:null, bar:null, items:[], opts:{}, mounted:false, storeKey:"default",
            placed:{}, widths:{}, arrow:{}, selMonths:null, cw:1500, ch:950,
            preset:"custom", perRow:"auto", title:"", lock:false, snap:true,
            monthPos:{}, monthMoved:{}, colorOverride:{}, extraArrows:{},
            userPlaced:{}, known:null, sel:null, selSet:{}, undo:[], ctx:null };
  function ast(id){ return S.arrow[id]||(S.arrow[id]={startOff:null,wp:null,end:null}); }
  function extras(id){ return S.extraArrows[id]||(S.extraArrows[id]=[]); }
  function arrowsOf(id){ return [ast(id)].concat(S.extraArrows[id]||[]); }   // [main, ...extras]
  function arrowRef(id,ai){ return ai<=0 ? ast(id) : (extras(id)[ai-1]||ast(id)); }

  /* ---- undo: snapshot the design fields before each change ---- */
  function snapshotDesign(){ return JSON.stringify({ placed:S.placed, widths:S.widths, arrow:S.arrow,
    extraArrows:S.extraArrows, monthPos:S.monthPos, monthMoved:S.monthMoved, colorOverride:S.colorOverride,
    cw:S.cw, ch:S.ch, title:S.title, selMonths:S.selMonths, perRow:S.perRow }); }
  function pushUndo(){ S.undo.push(snapshotDesign()); if(S.undo.length>40) S.undo.shift(); }
  function applyDesignObj(s){ ["placed","widths","arrow","extraArrows","monthPos","monthMoved","colorOverride",
    "cw","ch","title","selMonths","perRow"].forEach(function(k){ if(s[k]!==undefined) S[k]=s[k]; }); }
  function undo(){ if(!S.undo.length){ toast("Nothing to undo"); return; }
    applyDesignObj(JSON.parse(S.undo.pop())); persist(); buildBar(monthsAvail(S.items)); render(); }
  function toast(m){ if(S.opts.onToast) S.opts.onToast(m); }

  function persist(){ try{ localStorage.setItem("tt-anno:"+S.storeKey, JSON.stringify(
      { selMonths:S.selMonths, cw:S.cw, ch:S.ch, preset:S.preset, perRow:S.perRow, title:S.title,
        lock:S.lock, snap:S.snap, placed:S.placed, widths:S.widths, arrow:S.arrow,
        monthPos:S.monthPos, monthMoved:S.monthMoved, colorOverride:S.colorOverride, extraArrows:S.extraArrows,
        userPlaced:S.userPlaced, known:S.known })); }catch(e){} }
  function load(){ try{
      var raw=localStorage.getItem("tt-anno:"+S.storeKey); if(!raw) return; var d=JSON.parse(raw)||{};
      ["selMonths","cw","ch","preset","perRow","title","lock","snap","placed","widths","arrow","monthPos","monthMoved","colorOverride","extraArrows","userPlaced","known"]
        .forEach(function(k){ if(d[k]!==undefined && d[k]!==null) S[k]=d[k]; });
    }catch(e){} }

  function monthsAvail(items){
    var out=[];
    items.forEach(function(t){ var s=tStart(t); if(!s) return;
      [s, tEnd(t)].forEach(function(d){ var k=d.getFullYear()+"-"+d.getMonth();
        if(!out.some(function(m){return m.k===k;})) out.push({k:k,y:d.getFullYear(),m:d.getMonth()}); }); });
    out.sort(function(a,b){return a.y-b.y||a.m-b.m;});
    return out;
  }

  function mount(boardEl, barEl, opts){ S.board=boardEl; S.bar=barEl; S.opts=opts||{};
    S.storeKey=(opts&&opts.storeKey)||"default"; S.mounted=true; load(); }

  function update(items){
    S.items = (items||[]).filter(function(t){ return tStart(t); });
    var ids={}; S.items.forEach(function(t){ ids[t.id]=1; });
    // prune layout state for tasks that no longer exist
    [S.placed,S.widths,S.arrow,S.userPlaced].forEach(function(m){ Object.keys(m).forEach(function(k){ if(!ids[k]) delete m[k]; }); });
    // new-task detection: first time we ever see this timetable, seed silently
    if(S.known===null){ S.known={}; S.items.forEach(function(t){ S.known[t.id]=1; }); S._newIds={}; }
    else { S._newIds={}; S.items.forEach(function(t){ if(!S.known[t.id]){ S._newIds[t.id]=1; } S.known[t.id]=1; }); }
    var avail = monthsAvail(S.items);
    if(S.selMonths===null){ S.selMonths={}; avail.slice(0,2).forEach(function(m){ S.selMonths[m.k]=true; }); }
    persist(); buildBar(avail); render();
  }

  /* ---------------- toolbar ---------------- */
  function buildBar(avail){
    if(!S.bar) return; S.bar.innerHTML="";
    var b=S.bar;
    add(b,"strong","Months:");
    avail.forEach(function(m){
      var l=el("label","an-mo"); var cb=el("input"); cb.type="checkbox"; cb.checked=!!S.selMonths[m.k];
      cb.onchange=function(){ S.selMonths[m.k]=cb.checked; persist(); render(); };
      l.appendChild(cb); l.appendChild(document.createTextNode(MON[m.m].slice(0,3)+" "+String(m.y).slice(2))); b.appendChild(l);
    });
    // title
    var tg=el("span","an-grp"); tg.appendChild(document.createTextNode("Title "));
    var ti=el("input"); ti.type="text"; ti.placeholder="(optional)"; ti.value=S.title; ti.style.width="150px";
    ti.oninput=function(){ S.title=ti.value; persist(); render(); }; tg.appendChild(ti); b.appendChild(tg);
    // months per row
    var pg=el("span","an-grp"); pg.appendChild(document.createTextNode("Per row "));
    var ps=el("select"); ["auto","1","2","3","4"].forEach(function(v){ var o=el("option"); o.value=v; o.textContent=v; if(String(S.perRow)===v)o.selected=true; ps.appendChild(o); });
    ps.onchange=function(){ S.perRow=ps.value; persist(); render(); }; pg.appendChild(ps); b.appendChild(pg);
    // canvas preset + size
    var cg=el("span","an-grp"); cg.appendChild(document.createTextNode("Canvas "));
    var cs=el("select"); [["custom","Custom"],["16:9","16:9"],["A4","A4 land."],["fit","Fit"]].forEach(function(p){ var o=el("option"); o.value=p[0]; o.textContent=p[1]; if(S.preset===p[0])o.selected=true; cs.appendChild(o); });
    var wi=el("input"); wi.type="number"; wi.step="20"; wi.value=S.cw; wi.style.width="60px";
    var hi=el("input"); hi.type="number"; hi.step="20"; hi.value=S.ch; hi.style.width="60px";
    cs.onchange=function(){ S.preset=cs.value;
      if(PRESETS[S.preset]){ S.cw=PRESETS[S.preset][0]; S.ch=PRESETS[S.preset][1]; wi.value=S.cw; hi.value=S.ch; render(); }
      else if(S.preset==="fit"){ render(); fitToContent(); wi.value=S.cw; hi.value=S.ch; }
      else render(); persist(); };
    wi.onchange=function(){ S.cw=+wi.value||1500; S.preset="custom"; cs.value="custom"; persist(); render(); };
    hi.onchange=function(){ S.ch=+hi.value||950; S.preset="custom"; cs.value="custom"; persist(); render(); };
    cg.appendChild(cs); cg.appendChild(wi); cg.appendChild(document.createTextNode(" × ")); cg.appendChild(hi); b.appendChild(cg);
    // actions
    var ag=el("span","an-grp");
    var undoB=el("button","an-btn"); undoB.textContent="↩ Undo"; undoB.title="Undo the last design change"; undoB.onclick=undo; ag.appendChild(undoB);
    var tidy=el("button","an-btn"); tidy.textContent="Tidy"; tidy.title="Auto-arrange the labels"; tidy.onclick=function(){ pushUndo(); layoutDefaults(true); persist(); }; ag.appendChild(tidy);
    var reset=el("button","an-btn"); reset.textContent="Reset"; reset.title="Clear all manual arranging (labels and month positions)"; reset.onclick=function(){ S.placed={}; S.widths={}; S.arrow={}; S.userPlaced={}; S.monthPos={}; S.monthMoved={}; persist(); render(); }; ag.appendChild(reset);
    b.appendChild(ag);
    // toggles
    var lk=el("label","an-mo"); var lkc=el("input"); lkc.type="checkbox"; lkc.checked=S.lock;
    lkc.onchange=function(){ S.lock=lkc.checked; persist(); render(); }; lk.appendChild(lkc); lk.appendChild(document.createTextNode("🔒 Lock dates")); b.appendChild(lk);
    var sn=el("label","an-mo"); var snc=el("input"); snc.type="checkbox"; snc.checked=S.snap;
    snc.onchange=function(){ S.snap=snc.checked; persist(); }; sn.appendChild(snc); sn.appendChild(document.createTextNode("Snap")); b.appendChild(sn);
    // copy + save
    var cp=el("button","an-btn"); cp.textContent="⧉ Copy"; cp.title="Copy the board to the clipboard"; cp.onclick=copyToClipboard; b.appendChild(cp);
    var sv=el("button","an-btn an-btn-primary"); sv.textContent="💾 Save"; sv.title="Save this annotated design into the workbook"; sv.onclick=doSave; b.appendChild(sv);
    var hint=el("span","an-hint"); hint.textContent="Right-click a card to rename / recolour · drag on empty space to select several · drag a month header or the canvas edges."; b.appendChild(hint);
  }
  function el(tag,cls){ var e=document.createElement(tag); if(cls) e.className=cls; return e; }
  function add(p,tag,txt){ var e=el(tag); e.textContent=txt; if(tag==="strong") e.style.fontSize="12.5px"; p.appendChild(e); return e; }

  /* ---------------- render ---------------- */
  function render(){
    if(!S.board) return;
    var board=S.board; board.innerHTML=""; clearTip();
    board.style.setProperty("--an-cw",CW+"px"); board.style.setProperty("--an-ch",CH+"px");
    board.style.width=S.cw+"px"; board.style.height=S.ch+"px";

    var titleH = 0;
    if(S.title){ var tt=el("div","an-title"); tt.textContent=S.title; board.appendChild(tt); titleH=44; }

    var titleTop = titleH ? titleH+18 : 26;
    var avail=monthsAvail(S.items);
    var shown=avail.filter(function(m){ return S.selMonths[m.k]; });
    var idColor=colorMap(S.items);
    var tISO=todayISO();

    var taskCells={}, taskWrap={}, monthEls=[];
    shown.forEach(function(mm){
      var wrap=el("div","an-month"); wrap.dataset.mk=mm.k; wrap.style.position="absolute"; wrap.style.width=(7*CW+6*3)+"px";
      var hdr=el("div","an-mhdr"); hdr.style.background=HDR[mm.m]||"#4a6b8a"; hdr.textContent=MON[mm.m]+" "+mm.y; wrap.appendChild(hdr);
      var dow=el("div","an-dow"); DOW.forEach(function(d,i){var s=el("div");s.textContent=d; if(i>=5)s.style.color="#b3bcc7"; dow.appendChild(s);}); wrap.appendChild(dow);
      var grid=el("div","an-grid"); wrap.appendChild(grid);
      var first=new Date(mm.y,mm.m,1), start=addD(first,-midx(first)), last=new Date(mm.y,mm.m+1,0);
      var weeks=Math.ceil((midx(first)+last.getDate())/7);
      for(var i=0;i<weeks*7;i++){
        var d=addD(start,i), c=el("div","an-cell");
        if(d.getMonth()!==mm.m){ c.classList.add("empty"); grid.appendChild(c); continue; }
        var di=iso(d); c.dataset.date=di; c.innerHTML='<span class="an-num">'+d.getDate()+'</span>';
        var covers=S.items.filter(function(t){ return !isMile(t)&&di>=t.start&&di<=iso(tEnd(t)); })
                   .concat(S.items.filter(function(t){ return isMile(t)&&t.start===di; }));
        if(covers.length){
          var primary=covers[0];
          c.classList.add("tinted"); c.style.background=idColor[primary.id]; c.style.color=ink(idColor[primary.id]);
          c.dataset.pid=primary.id; c.dataset.jump=primary.id;
          if(isMile(primary)) c.classList.add("an-mile");
          if(covers.length>1){ var bars=el("div","an-cbars");
            covers.slice(1).forEach(function(t){ var bb=el("div","an-cbar"); bb.style.background=idColor[t.id]; bars.appendChild(bb); }); c.appendChild(bars); }
          covers.forEach(function(t){ (taskCells[t.id]=taskCells[t.id]||[]).push(c); taskWrap[t.id]=wrap; });
        } else { c.classList.add("plain"); if((i%7)>=5) c.classList.add("an-weekend"); }
        if(di===tISO) c.classList.add("an-today");
        grid.appendChild(c);
      }
      // round the ends of each same-task run within a week row, and close the gap between run cells
      var cells=[].slice.call(grid.children);
      cells.forEach(function(c,i){
        if(!c.dataset.pid) return; var col=i%7;
        var lft = col>0 && cells[i-1].dataset.pid===c.dataset.pid;
        var rgt = col<6 && cells[i+1] && cells[i+1].dataset.pid===c.dataset.pid;
        c.style.borderRadius=(lft?"0":"4px")+" "+(rgt?"0":"4px")+" "+(rgt?"0":"4px")+" "+(lft?"0":"4px");
        c.style.marginRight = rgt ? "-3px" : "";
        if(lft||rgt) c.style.zIndex="0";
      });
      board.appendChild(wrap); monthEls.push({wrap:wrap,key:mm.k}); makeMonthDrag(wrap,mm.k);
    });
    positionMonths(monthEls, titleTop);

    var svg=document.createElementNS("http://www.w3.org/2000/svg","svg"); svg.setAttribute("class","an-arrows"); board.appendChild(svg);
    svg.setAttribute("width",S.cw); svg.setAttribute("height",S.ch);

    var shownTasks=S.items.filter(function(t){ return taskCells[t.id]; }).sort(function(a,b){ return a.start<b.start?-1:1; });
    var labels={};
    shownTasks.forEach(function(t){
      var color=idColor[t.id], tc=ink(color);
      var lab=el("div","an-lbl"); lab.style.background=color; lab.style.color=tc; lab.style.borderColor=color;
      var dtc = tc==="#ffffff" ? "rgba(255,255,255,.8)" : "rgba(0,0,0,.55)";
      lab.innerHTML='<span class="an-dt" style="color:'+dtc+'">'+(isMile(t)?fmtNice(tStart(t)):fmtNice(tStart(t))+"–"+fmtNice(tEnd(t)))+'</span>'+esc(t.name);
      if(S._newIds && S._newIds[t.id]) lab.classList.add("an-new");
      board.appendChild(lab); labels[t.id]=lab;
      if(S.widths[t.id]){ lab.style.width=S.widths[t.id]+"px"; lab.style.maxWidth="none"; }
      lab._anchor=taskCells[t.id][Math.floor(taskCells[t.id].length/2)];
      if(S.placed[t.id]){ lab.style.left=S.placed[t.id].x+"px"; lab.style.top=S.placed[t.id].y+"px"; }
      makeLabelDrag(lab,t.id);
      var g=el("div","an-rsz"); lab.appendChild(g); makeResize(g,lab,t.id);
    });

    S.ctx={ tasks:shownTasks, labels:labels, svg:svg, color:idColor };
    board._anctx=S.ctx;
    layoutDefaults(false);        // place any not-yet-placed labels into the gutters
    applySelection();
    enableReschedule();

    // Paint-style canvas resize grips on the right edge, bottom edge and corner.
    ["r","b","br"].forEach(function(edge){ var h=el("div","an-edge an-edge-"+edge); board.appendChild(h); makeCanvasResize(h,edge); });
  }

  /* ---- month blocks: draggable by their header ---- */
  function positionMonths(monthEls, titleTop){
    var monthW = 7*CW+6*3;
    var cols = (S.perRow && S.perRow!=="auto") ? Math.max(1,+S.perRow)
             : Math.max(1, Math.floor((S.cw-24)/(monthW+44)));
    cols = Math.max(1, Math.min(cols, monthEls.length||1));
    var hs = monthEls.map(function(m){ return m.wrap.offsetHeight||260; });
    var rowMax=[]; monthEls.forEach(function(m,i){ var r=Math.floor(i/cols); rowMax[r]=Math.max(rowMax[r]||0, hs[i]); });
    var rowY=[], yy=titleTop; for(var r=0;r<rowMax.length;r++){ rowY[r]=yy; yy+=rowMax[r]+26; }
    var gridW = cols*monthW+(cols-1)*44, startX=Math.max(12, Math.round((S.cw-gridW)/2));
    monthEls.forEach(function(m,i){
      if(S.monthMoved[m.key] && S.monthPos[m.key]){ m.wrap.style.left=S.monthPos[m.key].x+"px"; m.wrap.style.top=S.monthPos[m.key].y+"px"; }
      else{ var col=i%cols, row=Math.floor(i/cols), x=startX+col*(monthW+44), y=rowY[row];
        m.wrap.style.left=x+"px"; m.wrap.style.top=y+"px"; S.monthPos[m.key]={x:x,y:y}; }
    });
  }
  function makeMonthDrag(wrap, key){
    var hdr=wrap.querySelector(".an-mhdr"); if(!hdr) return; hdr.style.cursor="move";
    hdr.addEventListener("mousedown", function(e){
      e.preventDefault(); e.stopPropagation(); var br=S.board.getBoundingClientRect();
      var dx=e.clientX-wrap.getBoundingClientRect().left, dy=e.clientY-wrap.getBoundingClientRect().top, pushed=false;
      function move(ev){ if(!pushed){pushUndo();pushed=true;} var x=Math.max(0, ev.clientX-br.left-dx), y=Math.max(0, ev.clientY-br.top-dy);
        wrap.style.left=x+"px"; wrap.style.top=y+"px"; S.monthPos[key]={x:x,y:y}; S.monthMoved[key]=1; drawArrows(); }
      function up(){ persist(); try{ layoutDefaults(false); }catch(e){} document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }
  function makeCanvasResize(h, edge){
    h.addEventListener("mousedown", function(e){
      e.preventDefault(); e.stopPropagation(); var br=S.board.getBoundingClientRect(), pushed=false;
      function move(ev){ if(!pushed){pushUndo();pushed=true;}
        if(edge==="r"||edge==="br") S.cw=Math.max(500, Math.round(ev.clientX-br.left));
        if(edge==="b"||edge==="br") S.ch=Math.max(360, Math.round(ev.clientY-br.top));
        S.board.style.width=S.cw+"px"; S.board.style.height=S.ch+"px";
        var svg=S.board.querySelector(".an-arrows"); if(svg){ svg.setAttribute("width",S.cw); svg.setAttribute("height",S.ch); }
        S.preset="custom";
      }
      function up(){ persist(); buildBar(monthsAvail(S.items)); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }

  function rectOf(e,b){ var r=e.getBoundingClientRect(); return {x:r.left-b.left,y:r.top-b.top,w:r.width,h:r.height}; }
  function centreOf(e,b){ var r=rectOf(e,b); return {x:r.x+r.w/2,y:r.y+r.h/2,w:r.w,h:r.h}; }

  /* Gutter-based auto layout: left-anchored tasks stack in the left margin,
     right-anchored ones on the right, so arrows fan outward instead of crossing. */
  function layoutDefaults(force){
    var c=S.ctx; if(!c) return; var b=S.board.getBoundingClientRect();
    // gutters are measured from the actual month blocks, not the full-width row
    var months=S.board.querySelectorAll(".an-month"); if(!months.length) return;
    var rowL=1e9, rowR=0, rowTop=1e9, rowBot=0;
    // The calendars themselves are obstacles — labels must never sit on top of them.
    var occ=[];
    months.forEach(function(m){ var r=m.getBoundingClientRect();
      var mr={x:r.left-b.left,y:r.top-b.top,w:r.width,h:r.height};
      rowL=Math.min(rowL,mr.x); rowR=Math.max(rowR,mr.x+mr.w); rowTop=Math.min(rowTop,mr.y); rowBot=Math.max(rowBot,mr.y+mr.h);
      occ.push(mr); });
    rowTop=Math.max(rowTop,40);
    // Labels the user dragged stay put and act as obstacles. Auto-placed labels
    // are re-flowed every layout change (deterministically) so they never go
    // stale when the canvas resizes, months toggle, or the sheet refreshes.
    if(!force) c.tasks.forEach(function(t){ if(!S.userPlaced[t.id]) return; var p=S.placed[t.id]; if(!p) return;
      var lab=c.labels[t.id]; lab.style.left=p.x+"px"; lab.style.top=p.y+"px";
      occ.push({x:p.x,y:p.y,w:lab.offsetWidth||220,h:lab.offsetHeight||40}); });
    function freeY(x,lw,lh,y0){                     // first y (top-down from y0) clear of every obstacle at this x
      var y=y0, guard=0;
      function hit(){ return occ.filter(function(o){ return x<o.x+o.w+10 && o.x<x+lw+10 && y<o.y+o.h+8 && o.y<y+lh+8; }); }
      var h=hit();
      while(h.length && guard++<300){ var nb=y+lh+12; h.forEach(function(o){ nb=Math.max(nb,o.y+o.h+12); }); y=nb; h=hit(); }
      return y;
    }
    c.tasks.slice().sort(function(a,bb){
      var ca=centreOf(c.labels[a.id]._anchor,b), cb=centreOf(c.labels[bb.id]._anchor,b);
      return ca.y-cb.y || ca.x-cb.x;
    }).forEach(function(t){
      var lab=c.labels[t.id]; if(!force && S.userPlaced[t.id]) return;
      var lw=lab.offsetWidth||220, lh=lab.offsetHeight||40, cc=centreOf(lab._anchor,b);
      var leftRoom = (rowL-8) >= (lw+32), rightRoom = (S.cw-8) >= (rowR+32+lw);
      var left = cc.x < S.cw/2, x, y0=rowTop;
      if(left && leftRoom) x=rowL-lw-32;             // clean left gutter
      else if(!left && rightRoom) x=rowR+32;         // clean right gutter
      else { x=Math.max(8, Math.min(S.cw-lw-8, cc.x-lw/2)); y0=rowBot+16; }  // no side room -> below its month
      var y = freeY(x,lw,lh,y0);
      lab.style.left=x+"px"; lab.style.top=y+"px"; S.placed[t.id]={x:x,y:y};
      occ.push({x:x,y:y,w:lw,h:lh});
    });
    drawArrows();
  }

  function fitToContent(){
    var b=S.board.getBoundingClientRect(), maxR=0, maxB=0;
    S.board.querySelectorAll(".an-cell:not(.empty), .an-lbl, .an-title").forEach(function(e){
      var r=e.getBoundingClientRect(); maxR=Math.max(maxR, r.right-b.left); maxB=Math.max(maxB, r.bottom-b.top);
    });
    S.cw=Math.round(maxR+24); S.ch=Math.round(maxB+24);
    S.board.style.width=S.cw+"px"; S.board.style.height=S.ch+"px";
    var svg=S.board.querySelector(".an-arrows"); if(svg){ svg.setAttribute("width",S.cw); svg.setAttribute("height",S.ch); }
    persist();
  }

  function drawArrows(){
    var c=S.ctx; if(!c) return; var b=S.board.getBoundingClientRect(); var out=""; var one=selOne();
    c.tasks.forEach(function(t){
      var lr=rectOf(c.labels[t.id],b), color=c.color[t.id], cc=centreOf(c.labels[t.id]._anchor,b);
      arrowsOf(t.id).forEach(function(s, ai){
        var start={x:lr.x+(s.startOff?s.startOff.x:lr.w/2), y:lr.y+(s.startOff?s.startOff.y:lr.h/2)};
        var tip;
        if(s.end){ tip={x:s.end.x,y:s.end.y}; }
        else if(ai===0){                                   // main arrow: meet the anchor cell's edge
          var wpd=s.wp||{x:(start.x+cc.x)/2,y:(start.y+cc.y)/2};
          var hw=cc.w/2,hh=cc.h/2,vx=wpd.x-cc.x,vy=wpd.y-cc.y; if(vx===0&&vy===0){vy=-1;}
          var k=Math.min(hw/(Math.abs(vx)||1e6),hh/(Math.abs(vy)||1e6)), ul=Math.hypot(vx,vy)||1;
          tip={x:cc.x+vx*k+(vx/ul), y:cc.y+vy*k+(vy/ul)};
        } else { tip={x:start.x+40,y:start.y-40}; }
        var wp=s.wp||{x:(start.x+tip.x)/2,y:(start.y+tip.y)/2};
        var sw=S.selSet[t.id]?3.5:2.5;
        var ang=Math.atan2(tip.y-wp.y,tip.x-wp.x), ah=8;
        var p1x=tip.x-ah*Math.cos(ang-0.42),p1y=tip.y-ah*Math.sin(ang-0.42);
        var p2x=tip.x-ah*Math.cos(ang+0.42),p2y=tip.y-ah*Math.sin(ang+0.42);
        out+='<path d="M '+start.x+' '+start.y+' Q '+wp.x+' '+wp.y+' '+tip.x+' '+tip.y+'" fill="none" stroke="'+color+'" stroke-width="'+sw+'"/>'
           +'<polygon points="'+tip.x+','+tip.y+' '+p1x+','+p1y+' '+p2x+','+p2y+'" fill="'+color+'"/>';
        if(t.id===one){ positionHandle(t.id,ai,"start",start); positionHandle(t.id,ai,"wp",wp); positionHandle(t.id,ai,"end",tip); }
      });
    });
    c.svg.innerHTML=out;
  }
  function positionHandle(id,ai,type,pt){ var h=S._handles && S._handles[id+"|"+ai+"|"+type]; if(h){ h.style.left=pt.x+"px"; h.style.top=pt.y+"px"; } }

  /* ---------------- selection (progressive disclosure) ---------------- */
  function selCount(){ return Object.keys(S.selSet).length; }
  function selOne(){ return selCount()===1 ? Object.keys(S.selSet)[0] : null; }
  function applySelection(){
    var c=S.ctx; if(!c) return;
    c.tasks.forEach(function(t){ c.labels[t.id].classList.toggle("an-sel", !!S.selSet[t.id]); });
    renderHandles();
    drawArrows();
  }
  /* Handles exist only for the single selected card, one set (start/bend/end) per arrow. */
  function renderHandles(){
    if(!S.board) return;
    S.board.querySelectorAll(".an-handle").forEach(function(h){ h.remove(); });
    S._handles={}; var one=selOne(); if(!one || !S.ctx) return;
    arrowsOf(one).forEach(function(s, ai){
      ["start","wp","end"].forEach(function(type){
        var h=mkHandle(type, one, ai, S.ctx.color[one]); S.board.appendChild(h);
        S._handles[one+"|"+ai+"|"+type]=h;
      });
    });
  }
  function selectLabel(id, additive){
    if(additive){ if(S.selSet[id]) delete S.selSet[id]; else S.selSet[id]=1; }
    else { S.selSet={}; S.selSet[id]=1; }
    S.sel=id; applySelection();
  }
  function clearSel(){ if(selCount()){ S.selSet={}; S.sel=null; applySelection(); } }

  /* ---------------- snapping ---------------- */
  function snapXY(x,y,selfId){
    if(!S.snap) return {x:x,y:y};
    x=Math.round(x/8)*8; y=Math.round(y/8)*8;               // soft grid
    var c=S.ctx; if(c) c.tasks.forEach(function(t){ if(t.id===selfId) return; var p=S.placed[t.id]; if(!p) return;
      if(Math.abs(p.x-x)<=6) x=p.x; if(Math.abs(p.y-y)<=6) y=p.y; });       // align to other labels
    return {x:x,y:y};
  }

  function mkHandle(type,id,ai,color){
    var h=el("div","an-handle"+(type==="wp"?" wp":type==="end"?" end":"")+(ai>0?" an-handle-extra":""));
    h.style.borderColor=color; if(type!=="wp") h.style.background=color;
    if(ai>0 && type==="end") h.title="Double-click to remove this arrow";
    h.addEventListener("mousedown",function(e){
      e.preventDefault(); e.stopPropagation(); var b=S.board.getBoundingClientRect(), pushed=false;
      function move(ev){ if(!pushed){pushUndo();pushed=true;} var px=ev.clientX-b.left, py=ev.clientY-b.top; var a=arrowRef(id,ai);
        if(type==="wp") a.wp={x:px,y:py};
        else if(type==="end"){
          if(S.snap){ var best=null,bd=14; S.board.querySelectorAll(".an-cell[data-date]").forEach(function(cel){ var r=cel.getBoundingClientRect(); var cx=r.left+r.width/2-b.left, cy=r.top+r.height/2-b.top; var d=Math.hypot(cx-px,cy-py); if(d<bd){bd=d;best={x:cx,y:cy};} }); if(best){px=best.x;py=best.y;} }
          a.end={x:px,y:py};
        } else { var lr=rectOf(S.ctx.labels[id],b); a.startOff={x:px-lr.x,y:py-lr.y}; }
        drawArrows(); }
      function up(){ persist(); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
    if(ai>0) h.addEventListener("dblclick", function(e){ e.preventDefault(); e.stopPropagation();
      pushUndo(); extras(id).splice(ai-1,1); persist(); applySelection(); });
    return h;
  }
  /* add a second (or third…) arrow from a card, for an action that spans months */
  function addExtraArrow(id){
    pushUndo(); var b=S.board.getBoundingClientRect();
    var a = S.ctx && S.ctx.labels[id] ? centreOf(S.ctx.labels[id]._anchor,b) : {x:120,y:120};
    extras(id).push({ startOff:null, wp:null, end:{x:a.x+70, y:a.y+70} });
    S.selSet={}; S.selSet[id]=1; S.sel=id; persist(); render();
    toast("Arrow added — drag its ◆ end onto the day it should point to");
  }

  function makeResize(g,lab,id){
    g.addEventListener("mousedown",function(e){
      e.preventDefault(); e.stopPropagation(); var startX=e.clientX, startW=lab.offsetWidth, pushed=false;
      function move(ev){ if(!pushed){pushUndo();pushed=true;} var w=Math.max(90,Math.min(600,startW+(ev.clientX-startX)));
        lab.style.width=w+"px"; lab.style.maxWidth="none"; S.widths[id]=w; drawArrows(); }
      function up(){ persist(); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }

  function makeLabelDrag(lab,id){
    lab.addEventListener("contextmenu", function(e){ e.preventDefault(); e.stopPropagation(); openEditor(id, e.clientX, e.clientY); });
    lab.addEventListener("mousedown",function(e){
      if(e.button!==0 || e.target.classList.contains("an-rsz")) return;
      e.preventDefault();
      // move all selected together if this label is part of a multi-selection
      var group = (S.selSet[id] && selCount()>1) ? Object.keys(S.selSet) : [id];
      var starts={}; group.forEach(function(g){ var l=S.ctx.labels[g]; var p=S.placed[g]||{x:parseFloat(l.style.left)||0,y:parseFloat(l.style.top)||0}; starts[g]={x:p.x,y:p.y}; });
      var ox=e.clientX, oy=e.clientY, moved=false;
      function move(ev){ if(!moved){ pushUndo(); moved=true; lab.classList.add("drag"); }
        var ddx=ev.clientX-ox, ddy=ev.clientY-oy;
        group.forEach(function(g){ var nx=Math.max(4,starts[g].x+ddx), ny=Math.max(4,starts[g].y+ddy);
          if(group.length===1){ var sp=snapXY(nx,ny,g); nx=sp.x; ny=sp.y; }
          var el=S.ctx.labels[g]; if(el){ el.style.left=nx+"px"; el.style.top=ny+"px"; el.classList.remove("an-new"); }
          S.placed[g]={x:nx,y:ny}; S.userPlaced[g]=1; });
        drawArrows(); }
      function up(){ lab.classList.remove("drag");
        if(moved) persist(); else selectLabel(id, e.shiftKey);   // click selects (shift = add); drag arranges
        document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }

  /* ---- right-click editor: rename + recolour ---- */
  function openEditor(id, cx, cy){
    closeEditor(); var t=byId(id); if(!t) return;
    var cur=(colorMap(S.items)[id]||"#2e86c1"); if(cur.length!==7) cur="#2e86c1";
    var pop=el("div","an-editor");
    pop.innerHTML='<label class="an-ed-l">Name</label><input class="an-ed-name" type="text">'+
      '<div class="an-ed-row"><label class="an-ed-l">Colour</label><input class="an-ed-color" type="color"></div>'+
      '<div class="an-ed-btns"><button class="an-btn" data-a="arrow">+ Arrow</button><button class="an-btn an-ed-del" data-a="del">Delete</button><span style="flex:1"></span>'+
      '<button class="an-btn" data-a="cancel">Cancel</button><button class="an-btn an-ed-save" data-a="save">Save</button></div>';
    document.body.appendChild(pop);
    var nameI=pop.querySelector(".an-ed-name"); nameI.value=t.name||"";
    var colI=pop.querySelector(".an-ed-color"); colI.value=cur;
    pop.style.left=Math.min(cx, window.innerWidth-pop.offsetWidth-12)+"px";
    pop.style.top=Math.min(cy, window.innerHeight-pop.offsetHeight-12)+"px";
    nameI.focus(); nameI.select();
    function commit(){
      var nn=nameI.value, nc=colI.value.toLowerCase();
      if(nn!==(t.name||"") && S.opts.onRename) S.opts.onRename(id, nn);
      if(nc!==cur){ pushUndo(); S.colorOverride[id]=nc; if(S.opts.onColor) S.opts.onColor(id, nc); persist(); render(); }
      closeEditor();
    }
    pop.addEventListener("mousedown", function(e){ e.stopPropagation(); });
    pop.addEventListener("click", function(e){ var a=e.target.getAttribute&&e.target.getAttribute("data-a"); if(!a) return;
      if(a==="save") commit(); else if(a==="cancel") closeEditor();
      else if(a==="arrow"){ closeEditor(); addExtraArrow(id); }
      else if(a==="del"){ closeEditor(); if(S.opts.onDelete) S.opts.onDelete(id); } });
    nameI.addEventListener("keydown", function(e){ if(e.key==="Enter"){e.preventDefault();commit();} else if(e.key==="Escape"){e.preventDefault();closeEditor();} });
    S._editor=pop; setTimeout(function(){ document.addEventListener("mousedown", outsideEditor, true); },0);
  }
  function outsideEditor(e){ if(S._editor && !S._editor.contains(e.target)) closeEditor(); }
  function closeEditor(){ document.removeEventListener("mousedown", outsideEditor, true); if(S._editor){ S._editor.remove(); S._editor=null; } }

  /* ---- save the design into the workbook (host persists it) ---- */
  function doSave(){ persist(); if(S.opts.onSaveDesign){ S.opts.onSaveDesign(localStorage.getItem("tt-anno:"+S.storeKey)||"{}"); } else toast("Saved locally"); }
  function applyDesign(jsonStr){ try{ var d=JSON.parse(jsonStr)||{};
      ["selMonths","cw","ch","preset","perRow","title","lock","snap","placed","widths","arrow","monthPos","monthMoved","colorOverride","extraArrows","userPlaced","known"]
        .forEach(function(k){ if(d[k]!==undefined && d[k]!==null) S[k]=d[k]; });
      persist(); buildBar(monthsAvail(S.items)); render();
    }catch(e){} }

  /* ---------------- reschedule (writes to Excel) ---------------- */
  var _tip=null;
  function showTip(x,y,txt){ if(!_tip){ _tip=el("div","an-tip"); document.body.appendChild(_tip); } _tip.textContent=txt; _tip.style.left=(x+14)+"px"; _tip.style.top=(y+16)+"px"; }
  function clearTip(){ if(_tip){ _tip.remove(); _tip=null; } }

  var EDGE=10, _dragging=false;
  function byId(id){ var t=null; S.items.some(function(x){ if(x.id===id){t=x;return true;} }); return t; }
  function zoneAt(cell,x){
    var t=byId(cell.dataset.pid); if(!t) return "move";
    var canR = S.opts.canResize ? S.opts.canResize() : true; if(!canR) return "move";
    var r=cell.getBoundingClientRect(), di=cell.dataset.date;
    if(di===t.start && (x-r.left)<=EDGE) return "start";
    if(di===iso(tEnd(t)) && (r.right-x)<=EDGE) return "end";
    return "move";
  }
  function enableReschedule(){
    var board=S.board;
    board.querySelectorAll(".an-cell.tinted[data-pid]").forEach(function(cell){
      if(S.lock){ cell.style.cursor="default"; return; }
      cell.addEventListener("mousemove",function(e){ if(_dragging) return;
        cell.style.cursor = zoneAt(cell,e.clientX)==="move" ? "grab" : "ew-resize"; });
      cell.addEventListener("mousedown",function(e){
        if(e.button!==0 || S.lock) return;
        var t=byId(cell.dataset.pid); if(!t) return; var id=t.id;
        e.preventDefault();
        var mode=zoneAt(cell,e.clientX);
        var grab=parse(cell.dataset.date), s0=parse(t.start), end0=t.end?parse(t.end):s0;
        var moved=false, pending=null; _dragging=true;
        board.classList.add("an-rescheduling");
        document.body.style.cursor = mode==="move" ? "grabbing" : "ew-resize";
        function cellUnder(x,y){ var el2=document.elementFromPoint(x,y); return el2&&el2.closest?el2.closest(".an-cell[data-date]"):null; }
        function apply(target,ev){
          var ns=s0, ne=end0;
          if(mode==="move"){ var d=daysBetween(grab,target); ns=addD(s0,d); ne=addD(end0,d); }
          else if(mode==="start"){ ns = target>end0 ? end0 : target; ne=end0; }
          else { ne = target<s0 ? s0 : target; ns=s0; }
          var lo=iso(ns), hi=iso(ne);
          board.querySelectorAll(".an-cell[data-date]").forEach(function(c){ c.classList.toggle("an-shift", c.dataset.date>=lo && c.dataset.date<=hi); });
          pending={ start:iso(ns), end: ne>ns ? iso(ne) : "" };
          var label = mode==="start"?"start ":mode==="end"?"end ":"";
          showTip(ev.clientX, ev.clientY, label + fmtNice(ns) + (ne>ns ? " → "+fmtNice(ne) : ""));
        }
        function move(ev){ var c=cellUnder(ev.clientX,ev.clientY); if(!c) return; moved=true; apply(parse(c.dataset.date),ev); }
        function up(){
          _dragging=false; document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up);
          document.body.style.cursor=""; board.classList.remove("an-rescheduling"); clearTip();
          board.querySelectorAll(".an-shift").forEach(function(c){ c.classList.remove("an-shift"); });
          if(moved && pending && (pending.start!==t.start || (pending.end||"")!==(t.end||""))){
            global.Cal && (global.Cal._suppressClick=true);
            if(S.opts.onEdit) S.opts.onEdit(id, pending.start, pending.end);
          }
        }
        document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
      });
    });
    // drag on empty canvas = rubber-band multi-select; a plain click clears
    if(!board._anRubber){ board._anRubber=true;
      board.addEventListener("mousedown", function(e){
        if(e.button!==0) return;
        if(e.target.closest(".an-lbl")||e.target.closest(".an-handle")||e.target.closest(".an-mhdr")||e.target.closest(".an-edge")||e.target.closest(".an-cell.tinted")) return;
        var br=board.getBoundingClientRect(), ox=e.clientX, oy=e.clientY, rect=null, dragged=false;
        function move(ev){ if(!dragged && Math.abs(ev.clientX-ox)+Math.abs(ev.clientY-oy)<4) return; dragged=true;
          if(!rect){ rect=el("div","an-rubber"); board.appendChild(rect); }
          rect.style.left=(Math.min(ox,ev.clientX)-br.left)+"px"; rect.style.top=(Math.min(oy,ev.clientY)-br.top)+"px";
          rect.style.width=Math.abs(ev.clientX-ox)+"px"; rect.style.height=Math.abs(ev.clientY-oy)+"px"; }
        function up(){ document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up);
          if(dragged && rect){ var rr=rect.getBoundingClientRect(); S.selSet={};
            S.ctx.tasks.forEach(function(t){ var lr=S.ctx.labels[t.id].getBoundingClientRect();
              if(lr.left<rr.right&&rr.left<lr.right&&lr.top<rr.bottom&&rr.top<lr.bottom) S.selSet[t.id]=1; });
            S.sel=selOne(); applySelection(); }
          else clearSel();
          if(rect) rect.remove(); }
        document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
      }); }
  }

  /* ---------------- export ---------------- */
  function capture(){ var h2c=global.html2canvas; var board=S.board; board.classList.add("an-exporting");
    return h2c(board,{backgroundColor:"#ffffff",scale:3,useCORS:true,logging:false,width:board.offsetWidth,height:board.offsetHeight,windowWidth:board.offsetWidth})
      .then(function(c){ board.classList.remove("an-exporting"); return c; })
      .catch(function(e){ board.classList.remove("an-exporting"); throw e; }); }
  function exportCanvas(){ return capture(); }
  function copyToClipboard(){
    if(!global.html2canvas){ toast("PNG library not loaded"); return; }
    capture().then(function(c){ c.toBlob(function(blob){
      if(navigator.clipboard && global.ClipboardItem){
        navigator.clipboard.write([new ClipboardItem({"image/png":blob})]).then(function(){ toast("Copied to clipboard"); },
          function(){ toast("Clipboard blocked — use Export PNG"); });
      } else toast("Clipboard not available — use Export PNG");
    },"image/png"); }).catch(function(e){ toast("Copy failed: "+(e&&e.message||e)); });
  }

  global.Annotated = { mount:mount, update:update, render:render, exportCanvas:exportCanvas,
                       applyDesign:applyDesign, undo:undo, isMounted:function(){ return S.mounted; } };
})(window);
