/* Annotated timeline view.
 *
 * A second way to look at the same items: compact month grids with solid
 * colour-per-task day cells, and the task descriptions living OUTSIDE the grid
 * as draggable/resizable label boxes tied to their days by colour arrows the
 * user can bend and re-anchor. Built for exporting a slide-ready PNG.
 *
 * It knows nothing about Excel: dragging a coloured run to reschedule calls
 * opts.onEdit(id, startISO, endISO) and the host persists it — the very same
 * contract the calendar view uses. Classes are namespaced `an-` so they never
 * collide with the calendar renderer's styles.
 */
(function (global) {
  "use strict";

  var MON = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  var DOW = ["Mo","Tu","We","Th","Fr","Sa","Su"];
  var CW = 46, CH = 34;
  // header tint per calendar-month (purely decorative)
  var HDR = {0:"#c47fb0",1:"#5aa0d6",2:"#7fb069",3:"#e0a13c",4:"#5aa0d6",5:"#7fb069",6:"#5aa0d6",7:"#d9736b",8:"#e0b93c",9:"#7fb069",10:"#b08bd6",11:"#5aa0d6"};
  var PALETTE = ["#2e86c1","#e67e22","#27ae60","#8e44ad","#c0392b","#16a085","#d4ac0d","#34495e","#d81b83","#2980b9","#af601a","#1abc9c"];

  function parse(s){ if(!s) return null; var p=String(s).split("-").map(Number); if(!p[0]) return null; return new Date(p[0],p[1]-1,p[2]); }
  function iso(d){ return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0"); }
  function midx(d){ return (d.getDay()+6)%7; }
  function addD(d,n){ var x=new Date(d); x.setDate(x.getDate()+n); return x; }
  function daysBetween(a,b){ return Math.round((b-a)/86400000); }
  function esc(s){ return String(s==null?"":s).replace(/[&<>]/g,function(c){return {"&":"&amp;","<":"&lt;",">":"&gt;"}[c];}); }
  function fmtNice(d){ return d.getDate()+" "+MON[d.getMonth()].slice(0,3); }
  function isMile(t){ return !t.end || t.end===t.start; }
  function tStart(t){ return parse(t.start); }
  function tEnd(t){ var e=t.end?parse(t.end):null, s=tStart(t); return (e&&e>s)?e:s; }
  function col(t,i){ return t.color || PALETTE[i%PALETTE.length]; }
  function ink(hex){ var h=String(hex||"#2e86c1").replace("#",""); var r=parseInt(h.slice(0,2),16)||0,g=parseInt(h.slice(2,4),16)||0,b=parseInt(h.slice(4,6),16)||0;
    return (0.299*r+0.587*g+0.114*b)>150?"#22303f":"#ffffff"; }

  var S = { board:null, bar:null, items:[], opts:{}, mounted:false, storeKey:"default",
            placed:{}, widths:{}, arrow:{}, selected:null, cw:1500, ch:950, ctx:null };
  function ast(id){ return S.arrow[id]||(S.arrow[id]={startOff:null,wp:null,end:null}); }

  /* Persist the whole arrangement (which months, canvas size, and every label
     position/width and arrow shape) so it survives re-renders, adding months,
     Excel refreshes, and reopening the dialog. Keyed per timetable. */
  function persist(){ try{ localStorage.setItem("tt-anno:"+S.storeKey, JSON.stringify(
      { selected:S.selected, cw:S.cw, ch:S.ch, placed:S.placed, widths:S.widths, arrow:S.arrow })); }catch(e){} }
  function load(){ try{
      var raw=localStorage.getItem("tt-anno:"+S.storeKey); if(!raw) return; var d=JSON.parse(raw)||{};
      if(d.selected) S.selected=d.selected; if(d.cw) S.cw=d.cw; if(d.ch) S.ch=d.ch;
      if(d.placed) S.placed=d.placed; if(d.widths) S.widths=d.widths; if(d.arrow) S.arrow=d.arrow;
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
    var avail = monthsAvail(S.items);
    if(S.selected===null){                       // first time: default to the earliest 2 months with tasks
      S.selected={}; avail.slice(0,2).forEach(function(m){ S.selected[m.k]=true; });
    }
    buildBar(avail); render();
  }

  function buildBar(avail){
    if(!S.bar) return; S.bar.innerHTML="";
    var lab=document.createElement("strong"); lab.textContent="Months:"; lab.style.fontSize="12.5px"; S.bar.appendChild(lab);
    avail.forEach(function(m){
      var l=document.createElement("label"); l.className="an-mo";
      var cb=document.createElement("input"); cb.type="checkbox"; cb.checked=!!S.selected[m.k];
      cb.onchange=function(){ S.selected[m.k]=cb.checked; persist(); render(); };
      l.appendChild(cb); l.appendChild(document.createTextNode(MON[m.m].slice(0,3)+" "+String(m.y).slice(2)));
      S.bar.appendChild(l);
    });
    var cg=document.createElement("span"); cg.className="an-grp";
    cg.innerHTML='Canvas ';
    var wi=document.createElement("input"); wi.type="number"; wi.step="20"; wi.value=S.cw;
    var hi=document.createElement("input"); hi.type="number"; hi.step="20"; hi.value=S.ch;
    wi.onchange=function(){ S.cw=+wi.value||1500; persist(); render(); };
    hi.onchange=function(){ S.ch=+hi.value||950; persist(); render(); };
    cg.appendChild(wi); cg.appendChild(document.createTextNode(" × ")); cg.appendChild(hi);
    S.bar.appendChild(cg);
    var hint=document.createElement("span"); hint.className="an-hint";
    hint.textContent="Drag a coloured run to reschedule · drag labels / arrow dots to arrange.";
    S.bar.appendChild(hint);
  }

  /* ---------------- render ---------------- */
  function render(){
    if(!S.board) return;
    var board=S.board; board.innerHTML="";
    board.style.setProperty("--an-cw",CW+"px"); board.style.setProperty("--an-ch",CH+"px");
    board.style.width=S.cw+"px"; board.style.height=S.ch+"px";

    var row=document.createElement("div"); row.className="an-months-row"; board.appendChild(row);
    var avail=monthsAvail(S.items);
    var shown=avail.filter(function(m){ return S.selected[m.k]; });
    var idColor={}, order={}; S.items.forEach(function(t,i){ idColor[t.id]=col(t,i); order[t.id]=i; });

    var taskCells={}, taskWrap={};
    shown.forEach(function(mm){
      var wrap=document.createElement("div"); wrap.className="an-month"; wrap.style.width=(7*CW+6*3)+"px";
      var hdr=document.createElement("div"); hdr.className="an-mhdr"; hdr.style.background=HDR[mm.m]||"#4a6b8a"; hdr.textContent=MON[mm.m]+" "+mm.y; wrap.appendChild(hdr);
      var dow=document.createElement("div"); dow.className="an-dow"; DOW.forEach(function(d){var s=document.createElement("div");s.textContent=d;dow.appendChild(s);}); wrap.appendChild(dow);
      var grid=document.createElement("div"); grid.className="an-grid"; wrap.appendChild(grid);
      var first=new Date(mm.y,mm.m,1), start=addD(first,-midx(first)), last=new Date(mm.y,mm.m+1,0);
      var weeks=Math.ceil((midx(first)+last.getDate())/7);
      for(var i=0;i<weeks*7;i++){
        var d=addD(start,i);
        var cell=document.createElement("div"); cell.className="an-cell";
        if(d.getMonth()!==mm.m){ cell.classList.add("empty"); grid.appendChild(cell); continue; }
        var di=iso(d); cell.dataset.date=di;
        cell.innerHTML='<span class="an-num">'+d.getDate()+'</span>';
        var covers=S.items.filter(function(t){ return !isMile(t)&&di>=t.start&&di<=iso(tEnd(t)); })
                   .concat(S.items.filter(function(t){ return isMile(t)&&t.start===di; }));
        if(covers.length){
          var primary=covers[0];
          cell.classList.add("tinted"); cell.style.background=idColor[primary.id]; cell.style.color=ink(idColor[primary.id]);
          cell.dataset.pid=primary.id; cell.dataset.jump=primary.id;
          if(covers.length>1){
            var bars=document.createElement("div"); bars.className="an-cbars";
            covers.slice(1).forEach(function(t){ var b=document.createElement("div"); b.className="an-cbar"; b.style.background=idColor[t.id]; bars.appendChild(b); });
            cell.appendChild(bars);
          }
          covers.forEach(function(t){ (taskCells[t.id]=taskCells[t.id]||[]).push(cell); taskWrap[t.id]=wrap; });
        } else cell.classList.add("plain");
        grid.appendChild(cell);
      }
      row.appendChild(wrap);
    });

    var svg=document.createElementNS("http://www.w3.org/2000/svg","svg"); svg.setAttribute("class","an-arrows"); board.appendChild(svg);
    svg.setAttribute("width",S.cw); svg.setAttribute("height",S.ch);

    var shownTasks=S.items.filter(function(t){ return taskCells[t.id]; }).sort(function(a,b){ return a.start<b.start?-1:1; });
    var labels={}, handles={}; var bRect=board.getBoundingClientRect(); var perWrap={};
    shownTasks.forEach(function(t){
      var color=idColor[t.id];
      var lab=document.createElement("div"); lab.className="an-lbl"; lab.style.borderLeftColor=color;
      lab.innerHTML='<span class="an-dt">'+(isMile(t)?fmtNice(tStart(t)):fmtNice(tStart(t))+"–"+fmtNice(tEnd(t)))+'</span>'+esc(t.name);
      S.board.appendChild(lab); labels[t.id]=lab;
      if(S.widths[t.id]){ lab.style.width=S.widths[t.id]+"px"; lab.style.maxWidth="none"; }
      lab._anchor=taskCells[t.id][Math.floor(taskCells[t.id].length/2)];
      if(S.placed[t.id]){ lab.style.left=S.placed[t.id].x+"px"; lab.style.top=S.placed[t.id].y+"px"; }
      else{ var wr=taskWrap[t.id].getBoundingClientRect(); var wkey=taskWrap[t.id].offsetLeft; perWrap[wkey]=(perWrap[wkey]||0);
        var x=wr.left-bRect.left, y=(wr.bottom-bRect.top)+18+perWrap[wkey]*56; lab.style.left=x+"px"; lab.style.top=y+"px"; S.placed[t.id]={x:x,y:y}; perWrap[wkey]++; }
      makeLabelDrag(lab,t.id);
      var g=document.createElement("div"); g.className="an-rsz"; lab.appendChild(g); makeResize(g,lab,t.id);
      handles[t.id]={ s:mkHandle("start",t.id,color), w:mkHandle("wp",t.id,color), e:mkHandle("end",t.id,color) };
      S.board.appendChild(handles[t.id].s); S.board.appendChild(handles[t.id].w); S.board.appendChild(handles[t.id].e);
    });

    S.ctx={ tasks:shownTasks, labels:labels, handles:handles, svg:svg, color:idColor };
    board._anctx=S.ctx;
    drawArrows();
    enableReschedule();
  }

  function rectOf(el,bRect){ var r=el.getBoundingClientRect(); return {x:r.left-bRect.left,y:r.top-bRect.top,w:r.width,h:r.height}; }
  function centreOf(el,bRect){ var r=rectOf(el,bRect); return {x:r.x+r.w/2,y:r.y+r.h/2,w:r.w,h:r.h}; }

  function drawArrows(){
    var c=S.ctx; if(!c) return; var bRect=S.board.getBoundingClientRect(); var out="";
    c.tasks.forEach(function(t){
      var s=ast(t.id); var lr=rectOf(c.labels[t.id],bRect);
      var start={x:lr.x+(s.startOff?s.startOff.x:lr.w/2), y:lr.y+(s.startOff?s.startOff.y:lr.h/2)};
      var cc=centreOf(c.labels[t.id]._anchor,bRect); var color=c.color[t.id]; var tip;
      if(s.end){ tip={x:s.end.x,y:s.end.y}; }
      else{ var wpd=s.wp||{x:(start.x+cc.x)/2,y:(start.y+cc.y)/2};
        var hw=cc.w/2,hh=cc.h/2,vx=wpd.x-cc.x,vy=wpd.y-cc.y; if(vx===0&&vy===0){vy=-1;}
        var k=Math.min(hw/(Math.abs(vx)||1e6),hh/(Math.abs(vy)||1e6)), ul=Math.hypot(vx,vy)||1;
        tip={x:cc.x+vx*k+(vx/ul), y:cc.y+vy*k+(vy/ul)}; }
      var wp=s.wp||{x:(start.x+tip.x)/2,y:(start.y+tip.y)/2};
      var ang=Math.atan2(tip.y-wp.y,tip.x-wp.x), ah=8;
      var p1x=tip.x-ah*Math.cos(ang-0.42),p1y=tip.y-ah*Math.sin(ang-0.42);
      var p2x=tip.x-ah*Math.cos(ang+0.42),p2y=tip.y-ah*Math.sin(ang+0.42);
      out+='<path d="M '+start.x+' '+start.y+' Q '+wp.x+' '+wp.y+' '+tip.x+' '+tip.y+'" fill="none" stroke="'+color+'" stroke-width="2.5"/>'
         +'<polygon points="'+tip.x+','+tip.y+' '+p1x+','+p1y+' '+p2x+','+p2y+'" fill="'+color+'"/>';
      var H=c.handles[t.id];
      H.s.style.left=start.x+"px"; H.s.style.top=start.y+"px";
      H.w.style.left=wp.x+"px"; H.w.style.top=wp.y+"px";
      H.e.style.left=tip.x+"px"; H.e.style.top=tip.y+"px";
    });
    c.svg.innerHTML=out;
  }

  function mkHandle(type,id,color){
    var h=document.createElement("div"); h.className="an-handle"+(type==="wp"?" wp":type==="end"?" end":"");
    h.style.borderColor=color; if(type!=="wp") h.style.background=color;
    h.addEventListener("mousedown",function(e){
      e.preventDefault(); e.stopPropagation(); var bRect=S.board.getBoundingClientRect();
      function move(ev){ var px=ev.clientX-bRect.left, py=ev.clientY-bRect.top;
        if(type==="wp") ast(id).wp={x:px,y:py};
        else if(type==="end") ast(id).end={x:px,y:py};
        else{ var lr=rectOf(S.ctx.labels[id],bRect); ast(id).startOff={x:px-lr.x,y:py-lr.y}; }
        drawArrows(); }
      function up(){ persist(); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
    return h;
  }

  function makeResize(g,lab,id){
    g.addEventListener("mousedown",function(e){
      e.preventDefault(); e.stopPropagation(); var startX=e.clientX, startW=lab.offsetWidth;
      function move(ev){ var w=Math.max(90,Math.min(600,startW+(ev.clientX-startX)));
        lab.style.width=w+"px"; lab.style.maxWidth="none"; S.widths[id]=w; drawArrows(); }
      function up(){ persist(); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }

  function makeLabelDrag(lab,id){
    lab.addEventListener("mousedown",function(e){
      if(e.target.classList.contains("an-rsz")) return;
      e.preventDefault(); lab.classList.add("drag"); var bRect=S.board.getBoundingClientRect();
      var dx=e.clientX-lab.getBoundingClientRect().left, dy=e.clientY-lab.getBoundingClientRect().top;
      function move(ev){ var x=Math.max(4,ev.clientX-bRect.left-dx), y=Math.max(4,ev.clientY-bRect.top-dy);
        lab.style.left=x+"px"; lab.style.top=y+"px"; S.placed[id]={x:x,y:y}; drawArrows(); }
      function up(){ lab.classList.remove("drag"); persist(); document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up); }
      document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
    });
  }

  /* Reschedule by dragging the coloured run itself:
       - grab the middle  → MOVE  (shift start & end together)
       - grab the left end → resize START
       - grab the right end→ resize END (or extend a milestone into a range)
     On drop we hand the new start/end to the host, which writes them to Excel —
     the same onEdit contract the calendar view uses. Edge-resize is only offered
     when the host says an End column exists (opts.canResize). */
  var EDGE=10, _dragging=false;
  function byId(id){ var t=null; S.items.some(function(x){ if(x.id===id){t=x;return true;} }); return t; }
  function zoneAt(cell,x){
    var t=byId(cell.dataset.pid); if(!t) return "move";
    var canR = S.opts.canResize ? S.opts.canResize() : true;
    if(!canR) return "move";
    var r=cell.getBoundingClientRect(), di=cell.dataset.date;
    if(di===t.start && (x-r.left)<=EDGE) return "start";
    if(di===iso(tEnd(t)) && (r.right-x)<=EDGE) return "end";
    return "move";
  }
  function enableReschedule(){
    var board=S.board;
    board.querySelectorAll(".an-cell.tinted[data-pid]").forEach(function(cell){
      cell.addEventListener("mousemove",function(e){ if(_dragging) return;
        cell.style.cursor = zoneAt(cell,e.clientX)==="move" ? "grab" : "ew-resize"; });
      cell.addEventListener("mousedown",function(e){
        if(e.button!==0) return;
        var t=byId(cell.dataset.pid); if(!t) return; var id=t.id;
        e.preventDefault();
        var mode=zoneAt(cell,e.clientX);
        var grab=parse(cell.dataset.date), s0=parse(t.start), end0=t.end?parse(t.end):s0;
        var moved=false, pending=null; _dragging=true;
        board.classList.add("an-rescheduling");    // let elementFromPoint see cells, not the handles/labels on top
        document.body.style.cursor = mode==="move" ? "grabbing" : "ew-resize";
        function cellUnder(x,y){ var el=document.elementFromPoint(x,y); return el&&el.closest?el.closest(".an-cell[data-date]"):null; }
        function apply(target){
          var ns=s0, ne=end0;
          if(mode==="move"){ var d=daysBetween(grab,target); ns=addD(s0,d); ne=addD(end0,d); }
          else if(mode==="start"){ ns = target>end0 ? end0 : target; ne=end0; }
          else { ne = target<s0 ? s0 : target; ns=s0; }
          var lo=iso(ns), hi=iso(ne);
          board.querySelectorAll(".an-cell[data-date]").forEach(function(c){
            c.classList.toggle("an-shift", c.dataset.date>=lo && c.dataset.date<=hi); });
          pending={ start:iso(ns), end: ne>ns ? iso(ne) : "" };
        }
        function move(ev){ var c=cellUnder(ev.clientX,ev.clientY); if(!c) return; moved=true; apply(parse(c.dataset.date)); }
        function up(){
          _dragging=false; document.removeEventListener("mousemove",move); document.removeEventListener("mouseup",up);
          document.body.style.cursor=""; board.classList.remove("an-rescheduling");
          board.querySelectorAll(".an-shift").forEach(function(c){ c.classList.remove("an-shift"); });
          if(moved && pending && (pending.start!==t.start || (pending.end||"")!==(t.end||""))){
            global.Cal && (global.Cal._suppressClick=true);
            if(S.opts.onEdit) S.opts.onEdit(id, pending.start, pending.end);
          }
        }
        document.addEventListener("mousemove",move); document.addEventListener("mouseup",up);
      });
    });
  }

  function exportCanvas(html2canvas){
    var board=S.board; board.classList.add("an-exporting");
    return html2canvas(board,{backgroundColor:"#ffffff",scale:3,useCORS:true,logging:false,
      width:board.offsetWidth,height:board.offsetHeight,windowWidth:board.offsetWidth})
      .then(function(c){ board.classList.remove("an-exporting"); return c; })
      .catch(function(e){ board.classList.remove("an-exporting"); throw e; });
  }

  global.Annotated = { mount:mount, update:update, render:render, exportCanvas:exportCanvas,
                       isMounted:function(){ return S.mounted; } };
})(window);
