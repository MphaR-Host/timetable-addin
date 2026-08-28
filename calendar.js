/* Timetable calendar renderer.
 *
 * Pure view code: give it items and it produces DOM. It knows nothing about
 * Excel, Office.js or transport, so the task pane and the pop-out dialog share
 * exactly one implementation and can never drift apart.
 */
(function (global) {
  "use strict";

  var DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  var MON = ["January", "February", "March", "April", "May", "June", "July",
             "August", "September", "October", "November", "December"];
  var STATUSES = ["Pending", "Ongoing", "Done"];
  var SG = { Pending: "○", Ongoing: "▶", Done: "✓" };

  var DENSITY = {
    compact:     {col:340,laneH:19,barH:16,barTop:26,minDay:30,dlH:15,barFs:10.5,dayFs:11.5,dlFs:10.5,numSz:19,dowFs:10.5,mheadFs:14,legFs:11.5},
    comfortable: {col:470,laneH:26,barH:22,barTop:28,minDay:34,dlH:18,barFs:12.5,dayFs:13,dlFs:12,numSz:24,dowFs:12,mheadFs:16,legFs:12.5},
    large:       {col:620,laneH:32,barH:27,barTop:34,minDay:44,dlH:23,barFs:14.5,dayFs:15,dlFs:14,numSz:28,dowFs:13.5,mheadFs:18,legFs:14},
    xlarge:      {col:820,laneH:40,barH:34,barTop:42,minDay:58,dlH:29,barFs:17,dayFs:17.5,dlFs:16.5,numSz:34,dowFs:15,mheadFs:21,legFs:16}
  };
  var D = DENSITY.comfortable;
  var CMAP = {};        // task id -> display colour, de-duplicated
  var EXPANDED = {};     // week key -> true while its hidden lanes are shown
  var LAST = null;       // last render arguments, so a "+N more" chip can redraw

  /* ---- dates ---- */
  function parseD(s){ if(!s) return null; var p=String(s).split("-").map(Number);
    if(!p[0]) return null; return new Date(p[0],p[1]-1,p[2]); }
  function fmtISO(d){ return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+
    "-"+String(d.getDate()).padStart(2,"0"); }
  function today0(){ var t=new Date(); return new Date(t.getFullYear(),t.getMonth(),t.getDate()); }
  function mondayIndex(d){ return (d.getDay()+6)%7; }
  function addDays(d,n){ var x=new Date(d); x.setDate(x.getDate()+n); return x; }
  function sameDay(a,b){ return a&&b&&a.getTime()===b.getTime(); }
  function daysBetween(a,b){ return Math.round((b-a)/86400000); }
  function isWeekend(d){ var g=d.getDay(); return g===0||g===6; }
  function iStart(it){ return parseD(it.start); }
  function iEnd(it){ var e=it.end?parseD(it.end):null, s=iStart(it); return (e&&e>s)?e:s; }
  function isSpan(it){ var e=it.end?parseD(it.end):null, s=iStart(it); return !!(e&&e>s); }
  function firstVis(it){ var d=iStart(it); if(it.async) return d;
    while(isWeekend(d)) d=addDays(d,1); return d; }
  function lastVis(it){ var d=iEnd(it); if(it.async) return d;
    while(isWeekend(d)) d=addDays(d,-1); return d; }
  function esc(s){ return String(s==null?"":s).replace(/[&<>"]/g,function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]; }); }
  function dm(d){ return d.getDate()+" "+MON[d.getMonth()].slice(0,3); }

  /* ---- colour ---- */
  function textOn(hex){ var h=String(hex).replace("#","");
    var r=parseInt(h.slice(0,2),16),g=parseInt(h.slice(2,4),16),b=parseInt(h.slice(4,6),16);
    return (0.299*r+0.587*g+0.114*b)>150?"#1a2530":"#ffffff"; }
  /* Task colour darkened until it genuinely clears WCAG AA on white. */
  function ink(hex){ var h=String(hex||"#2e86c1").replace("#","");
    var r=parseInt(h.slice(0,2),16)||0,g=parseInt(h.slice(2,4),16)||0,b=parseInt(h.slice(4,6),16)||0;
    function lin(v){ v/=255; return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4); }
    function contrast(){ return 1.05/(0.2126*lin(r)+0.7152*lin(g)+0.0722*lin(b)+0.05); }
    var guard=0;
    while(contrast()<4.5&&guard++<40){ r=Math.round(r*0.92); g=Math.round(g*0.92); b=Math.round(b*0.92); }
    return "rgb("+r+","+g+","+b+")"; }

  /* A mid-tone fill fails AA against BOTH white and near-black text, so for the
     solid Ongoing state we deepen the fill (hue preserved) until white text
     genuinely clears 4.5:1, rather than picking the least-bad label colour. */
  function solidFill(hex){
    var h=String(hex||"#2e86c1").replace("#","");
    var r=parseInt(h.slice(0,2),16)||0,g=parseInt(h.slice(2,4),16)||0,b=parseInt(h.slice(4,6),16)||0;
    function lin(v){ v/=255; return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4); }
    function whiteOn(){ return 1.05/(0.2126*lin(r)+0.7152*lin(g)+0.0722*lin(b)+0.05); }
    var guard=0;
    while(whiteOn()<4.5&&guard++<40){ r=Math.round(r*0.93); g=Math.round(g*0.93); b=Math.round(b*0.93); }
    return "rgb("+r+","+g+","+b+")";
  }

  var DONE_FILL="#e4e9ee", DONE_INK="#5a656f", DONE_EDGE="#9aa6b2";

  /* The host falls back to a small palette assigned by row, so a sheet with no
     Colour column hands us the same colour every ten tasks. Keep every colour
     that is unique and re-assign only the clashes, so no two visible tasks
     share a fill. An explicitly chosen colour is only moved if it collides. */
  var DISTINCT=["#2e86c1","#e67e22","#27ae60","#8e44ad","#c0392b","#16a085","#d4ac0d","#34495e",
                "#d81b83","#2980b9","#af601a","#1abc9c","#7f8c8d","#c0398b","#0e6655","#b7950b",
                "#6c3483","#1f618d","#a04000","#117864"];
  /* Past the hand-picked list, walk the hue circle by the golden angle so any
     number of tasks still gets a colour of its own. */
  function hslHex(h,s,l){
    s/=100; l/=100;
    var k=function(n){ return (n+h/30)%12; },
        a=s*Math.min(l,1-l),
        f=function(n){ var v=l-a*Math.max(-1,Math.min(k(n)-3,Math.min(9-k(n),1)));
          return Math.round(255*v).toString(16).padStart(2,"0"); };
    return "#"+f(0)+f(8)+f(4);
  }
  function distinctColors(items){
    var map={}, used={}, di=0, gi=0;
    items.forEach(function(it){
      var c=String(it.color||"").toLowerCase();
      if(!c || used[c]){
        c=null;
        while(di<DISTINCT.length){                       // curated colours first
          var cand=DISTINCT[di++].toLowerCase();
          if(!used[cand]){ c=cand; break; }
        }
        while(!c){                                       // then generated hues
          var gen=hslHex((gi*137.508)%360, 58, 38); gi++;
          if(!used[gen]) c=gen;
          if(gi>720) { c=gen; break; }
        }
      }
      used[c]=1; map[it.id]=c;
    });
    return map;
  }
  function colOf(it){ return CMAP[it.id] || it.color || "#2e86c1"; }

  /* ISO-8601 week number (weeks start Monday; week 1 holds the first Thursday). */
  function isoWeek(d){
    var x=new Date(d.getFullYear(),d.getMonth(),d.getDate());
    x.setDate(x.getDate()+3-mondayIndex(x));
    var first=new Date(x.getFullYear(),0,4);
    return 1+Math.round(((x-first)/86400000-3+mondayIndex(first))/7);
  }

  /* Status reads as three different fills, never three opacities.
     backgroundColor - not the `background` shorthand - or the inline value
     would reset background-image and wipe out the Ongoing stripes. */
  function paint(el,color,status,edge){
    if(status==="Done"){
      // Keep a stripe of the task's own colour: a finished item should still say
      // which workstream it belonged to, not just that it is finished.
      el.style.backgroundColor=DONE_FILL; el.style.color=DONE_INK;
      el.style.boxShadow="inset "+(edge||"3px")+" 0 0 "+color+", inset 0 0 0 1px "+DONE_EDGE;
    } else if(status==="Ongoing"){
      el.style.backgroundColor=solidFill(color); el.style.color="#ffffff";
      el.style.boxShadow=edge?("inset "+edge+" 0 0 rgba(0,0,0,.35)"):"0 0 0 1px rgba(255,255,255,.55)";
    } else {
      el.style.backgroundColor="#ffffff"; el.style.color=ink(color);
      el.style.boxShadow="inset 0 0 0 2px "+color;
    }
  }

  function applyDensity(name, perRow){
    D = DENSITY[name] || DENSITY.comfortable;
    var s=document.documentElement.style;
    s.setProperty("--colw",D.col+"px");   s.setProperty("--barh",D.barH+"px");
    s.setProperty("--barfs",D.barFs+"px");s.setProperty("--dayfs",D.dayFs+"px");
    s.setProperty("--dlfs",D.dlFs+"px");  s.setProperty("--dlh",D.dlH+"px");
    s.setProperty("--numsz",D.numSz+"px");s.setProperty("--mindaysz",D.minDay+"px");
    s.setProperty("--dowfs",D.dowFs+"px");s.setProperty("--mheadfs",D.mheadFs+"px");
    s.setProperty("--legfs",D.legFs+"px");
    var box=document.querySelector(".months");
    if(box) box.style.gridTemplateColumns = (!perRow||perRow==="auto")
      ? "repeat(auto-fill,minmax("+D.col+"px,1fr))" : "repeat("+perRow+",1fr)";
  }

  function assignLanes(events){
    var sorted=events.slice().sort(function(a,b){ return a.s-b.s||b.e-a.e; });
    var laneEnds=[], out=[];
    sorted.forEach(function(ev){
      var lane=laneEnds.findIndex(function(end){ return end<ev.s; });
      if(lane===-1){ lane=laneEnds.length; laneEnds.push(ev.e); } else laneEnds[lane]=ev.e;
      var o={}; for(var k in ev) o[k]=ev[k]; o.lane=lane; out.push(o);
    });
    return { placed:out, lanes:laneEnds.length };
  }

  function mkey(d){ return d.getFullYear()+"-"+d.getMonth(); }
  function parseMKey(s){ var p=String(s||"").split("-"); return p.length===2 && p[0]!==""
    ? new Date(+p[0], +p[1], 1) : null; }

  /* Every month a task touches, for the month pickers. */
  function monthsOf(items){
    var out=[], seen={};
    (items||[]).forEach(function(it){ var s=iStart(it); if(!s) return;
      var cur=new Date(s.getFullYear(),s.getMonth(),1), end=iEnd(it);
      while(cur<=end){ var k=mkey(cur);
        if(!seen[k]){ seen[k]=1; out.push({ k:k, y:cur.getFullYear(), m:cur.getMonth(),
          label:MON[cur.getMonth()].slice(0,3)+" "+cur.getFullYear() }); }
        cur=new Date(cur.getFullYear(),cur.getMonth()+1,1); }
    });
    out.sort(function(a,b){ return a.y-b.y||a.m-b.m; });
    return out;
  }

  /* The span of months to draw. Without an explicit from/to this is every month
     the data touches (plus today), which on a long project is a lot of empty
     grids - hence opts.from / opts.to / opts.hideEmpty. */
  function monthRange(items, from, to){
    var min=null,max=null;
    items.forEach(function(it){ var s=iStart(it),e=iEnd(it);
      if(!min||s<min)min=s; if(!max||e>max)max=e; });
    var t=today0();
    if(!min){ min=t; max=t; }
    if(t<min)min=t; if(t>max)max=t;
    var f=new Date(min.getFullYear(),min.getMonth(),1),
        l=new Date(max.getFullYear(),max.getMonth(),1);
    var pf=parseMKey(from), pl=parseMKey(to);
    if(pf&&pf>f) f=pf;
    if(pl&&pl<l) l=pl;
    if(f>l) l=new Date(f);
    return { from:f, to:l };
  }

  /* Milestone labels wrap, so their height is only known once in the DOM.
     Grow every day cell in a week to the tallest stack in that week. */
  function balanceWeeks(root){
    (root||document).querySelectorAll(".week").forEach(function(w){
      var msTop=+w.dataset.mstop||0, need=msTop+6;
      w.querySelectorAll(".msbox").forEach(function(b){
        need=Math.max(need,msTop+b.getBoundingClientRect().height+5); });
      w.querySelectorAll(".day").forEach(function(d){
        d.style.minHeight=Math.ceil(need)+"px"; });
    });
  }

  function renderMonth(year,month,vis,t,opts){
    opts=opts||{};
    var showDates=opts.showDates, weekNums=!!opts.weekNums;
    var maxLanes=opts.maxLanes||5;
    var wrap=document.createElement("div"); wrap.className="month"+(weekNums?" wknums":"");
    wrap.innerHTML='<div class="mhead">'+MON[month]+" "+year+'</div><div class="dow">'+
      DOW.map(function(d,i){ return '<div class="'+(i>=5?"we":"")+'">'+d+"</div>"; }).join("")+"</div>";
    var weeksBox=document.createElement("div"); weeksBox.className="weeks";
    var first=new Date(year,month,1);
    var startCell=addDays(first,-mondayIndex(first));
    var mStart=new Date(year,month,1), mEnd=new Date(year,month+1,0);
    var monthItems=vis.filter(function(it){ return iEnd(it)>=mStart&&iStart(it)<=mEnd; });
    var numWeeks=Math.ceil((mondayIndex(first)+mEnd.getDate())/7);

    for(var w=0;w<numWeeks;w++){
      var wkStart=addDays(startCell,w*7), wkEnd=addDays(wkStart,6);
      var week=document.createElement("div"); week.className="week";
      var days=document.createElement("div"); days.className="days";

      var evs=[];
      monthItems.filter(isSpan).forEach(function(it){
        var s=iStart(it), e=iEnd(it);
        if(e<wkStart||s>wkEnd) return;
        var segS=s<wkStart?wkStart:s, segE=e>wkEnd?wkEnd:e;
        if(!it.async){
          var fri=addDays(wkStart,4);
          if(segS>fri){
            /* Mon-Fri is all this week shows, and this segment is entirely on the
               weekend. Skip it - unless the task has no working day at all, which
               would drop it from the grid completely; pin that to the Friday. */
            if(firstVis(it)>lastVis(it)){ segS=fri; segE=fri; }
            else return;
          } else {
            if(segE>fri) segE=fri;
            if(segE<segS) return;
          }
        }
        evs.push({ it:it, s:segS.getTime(), e:segE.getTime(),
                   col:daysBetween(wkStart,segS), span:daysBetween(segS,segE)+1 });
      });
      var res=assignLanes(evs), placed=res.placed, allLanes=res.lanes;
      /* A busy week can stack a dozen lanes and make the row enormous, so cap it
         and offer the rest behind a "+N more" chip. */
      var wkKey=fmtISO(wkStart), openWk=!!EXPANDED[wkKey];
      var capped=(!openWk && allLanes>maxLanes);
      var lanes=capped?maxLanes:allLanes, overflow=0;
      if(capped){
        overflow=placed.filter(function(p){ return p.lane>=maxLanes; }).length;
        placed=placed.filter(function(p){ return p.lane<maxLanes; });
      }

      /* Single-day milestones live inside their own day cell only - a one-day
         event must never look like it spans several days. */
      var msByDay={};
      monthItems.filter(function(it){ return !isSpan(it); }).forEach(function(it){
        var s=iStart(it); if(s<wkStart||s>wkEnd) return;
        var c=daysBetween(wkStart,s); (msByDay[c]=msByDay[c]||[]).push(it);
      });
      var msTop=D.barTop+lanes*D.laneH;
      week.dataset.mstop=msTop;
      week.dataset.wkstart=fmtISO(wkStart);   // Monday of this row, for drag→date

      for(var d=0;d<7;d++){
        var cd=addDays(wkStart,d);
        var cell=document.createElement("div");
        cell.className="day"; cell.style.minHeight=(msTop+6)+"px";
        cell.dataset.date=fmtISO(cd);
        if(d>=5) cell.classList.add("weekend");
        if(cd.getMonth()!==month) cell.classList.add("other");
        if(sameDay(cd,t)) cell.classList.add("today");
        if(monthItems.some(function(it){ return it.key&&sameDay(iStart(it),cd); }))
          cell.classList.add("keybox");
        cell.innerHTML='<div class="num">'+cd.getDate()+"</div>";
        var dayMs=msByDay[d]||[];
        if(dayMs.length){
          var box=document.createElement("div"); box.className="msbox"; box.style.top=msTop+"px";
          dayMs.forEach(function(it){
            var el=document.createElement("div");
            el.className="msi st-"+esc(it.status)+(it.key?" key":"")+(it.async?" is-async":"");
            el.dataset.jump=it.id;
            el.dataset.start=fmtISO(iStart(it));
            el.dataset.async=it.async?"1":"";
            el.dataset.name=it.name||"";
            paint(el,colOf(it),it.status,"3px");
            var gcol=it.status==="Ongoing"?"#ffffff":(it.status==="Done"?DONE_INK:ink(colOf(it)));
            el.innerHTML='<span class="sg" style="color:'+gcol+'">'+(SG[it.status]||"○")+
              '</span><span class="lbl">'+(esc(it.name)||"(untitled)")+"</span>";
            el.setAttribute("aria-label",(it.name||"(untitled)")+", "+it.status+", "+fmtISO(iStart(it)));
            el.dataset.status=it.status;
            el.dataset.notes=it.notes||"";
            box.appendChild(el);
          });
          cell.appendChild(box);
        }
        days.appendChild(cell);
      }
      if(weekNums){
        var wn=document.createElement("div"); wn.className="wknum";
        wn.textContent=isoWeek(wkStart); wn.title="ISO week "+isoWeek(wkStart);
        week.appendChild(wn);
      }
      week.appendChild(days);

      placed.forEach(function(p){
        var bar=document.createElement("div");
        bar.className="bar st-"+esc(p.it.status)+(p.it.async?" is-async":"");
        paint(bar,colOf(p.it),p.it.status,"");
        bar.style.left="calc("+((p.col/7)*100)+"% + 3px)";
        bar.style.width="calc("+((p.span/7)*100)+"% - 6px)";
        bar.style.top=(D.barTop+p.lane*D.laneH)+"px";
        var fvis=firstVis(p.it).getTime(), lvis=lastVis(p.it).getTime();
        var contL=p.s>fvis, contR=p.e<lvis;
        var lr=contL?0:4, rr=contR?0:4;
        bar.style.borderRadius=lr+"px "+rr+"px "+rr+"px "+lr+"px";
        var label=(contL?"… ":((SG[p.it.status]||"")+" "))+(p.it.name||"(untitled)");
        if(showDates&&!contL) label+="  ("+dm(iStart(p.it))+"–"+dm(iEnd(p.it))+")";
        bar.textContent=label;
        bar.setAttribute("aria-label",(p.it.name||"(untitled)")+", "+p.it.status+", "+
          fmtISO(iStart(p.it))+" to "+fmtISO(iEnd(p.it)));
        bar.dataset.status=p.it.status;
        bar.dataset.notes=p.it.notes||"";
        bar.dataset.jump=p.it.id;
        // Whole-task dates + whether THIS segment is a continuation, so drag
        // can offer edge-resize only on the segment showing the real edge.
        bar.dataset.start=fmtISO(iStart(p.it));
        bar.dataset.end=fmtISO(iEnd(p.it));
        bar.dataset.contl=contL?"1":"";
        bar.dataset.contr=contR?"1":"";
        bar.dataset.async=p.it.async?"1":"";
        bar.dataset.name=p.it.name||"";
        week.appendChild(bar);
      });
      if(capped || (openWk && allLanes>maxLanes)){
        var chip=document.createElement("button");
        chip.type="button"; chip.className="lanemore";
        chip.textContent = capped ? ("+"+overflow+" more") : "− less";
        chip.style.top=(D.barTop+lanes*D.laneH)+"px";
        (function(k){ chip.addEventListener("click", function(e){
          e.preventDefault(); e.stopPropagation();
          EXPANDED[k]=!EXPANDED[k];
          if(LAST) render(LAST.target, LAST.items, LAST.opts);
        }); })(wkKey);
        week.appendChild(chip);
      }
      weeksBox.appendChild(week);
    }
    wrap.appendChild(weeksBox);

    var leg=document.createElement("div"); leg.className="legend";
    var list=monthItems.slice().sort(function(a,b){ return iStart(a)-iStart(b); });
    if(list.length){
      list.forEach(function(it){
        var s=iStart(it);
        var li=document.createElement("div"); li.className="li"; li.dataset.jump=it.id;
        var rng=isSpan(it)?(s.getDate()+"–"+iEnd(it).getDate()):String(s.getDate());
        li.innerHTML='<span class="d">'+rng+'</span><span class="swatch" style="background:'+
          esc(colOf(it))+'"></span><span class="sg" style="color:'+ink(colOf(it))+'">'+
          (SG[it.status]||"○")+'</span><span class="st-'+esc(it.status)+'">'+
          (esc(it.name)||"(untitled)")+"</span>";
        leg.appendChild(li);
      });
    } else leg.innerHTML='<span class="mini">No items this month</span>';
    wrap.appendChild(leg);
    return wrap;
  }

  function render(target, items, opts){
    opts = opts || {};
    LAST = { target:target, items:items, opts:opts };
    var hidden = opts.hidden || [];
    var dated = items.filter(function(it){ return iStart(it); });
    // Colours are derived from every dated task, not just the visible ones, so
    // hiding a status doesn't shuffle the colours of everything else.
    CMAP = distinctColors(dated);
    var vis = dated.filter(function(it){ return hidden.indexOf(it.status) === -1; });
    target.innerHTML = "";
    var r = monthRange(vis, opts.from, opts.to), t = today0();
    var tKey = mkey(t);
    var box = document.createElement("div"); box.className = "months";
    var cur = new Date(r.from), drawn = 0;
    while(cur <= r.to){
      var mStart=new Date(cur.getFullYear(),cur.getMonth(),1),
          mEnd=new Date(cur.getFullYear(),cur.getMonth()+1,0);
      var has = vis.some(function(it){ return iEnd(it)>=mStart && iStart(it)<=mEnd; });
      if(has || !opts.hideEmpty){
        var el = renderMonth(cur.getFullYear(), cur.getMonth(), vis, t, opts);
        el.dataset.mk = mkey(cur);
        if(mkey(cur)===tKey) el.classList.add("cur");
        box.appendChild(el); drawn++;
      }
      cur = new Date(cur.getFullYear(), cur.getMonth()+1, 1);
    }
    // Never end up with a blank pane: fall back to the month we are in.
    if(!drawn){
      var only = renderMonth(t.getFullYear(), t.getMonth(), vis, t, opts);
      only.dataset.mk = tKey; only.classList.add("cur"); box.appendChild(only);
    }
    target.appendChild(box);
    applyDensity(opts.density, opts.perRow);
    balanceWeeks(target);
    if(opts.scrollToday) scrollToToday(target);
  }

  /* Bring the month containing today into view - on a long project the calendar
     otherwise opens months away from where the work is. */
  function scrollToToday(target){
    var el=(target||document).querySelector(".month.cur");
    if(!el) return false;
    try { el.scrollIntoView({block:"nearest", inline:"nearest"}); } catch(e){ el.scrollIntoView(); }
    return true;
  }

  function statusKey(el, items){
    var counts={Pending:0,Ongoing:0,Done:0};
    items.forEach(function(it){ if(counts[it.status]!=null) counts[it.status]++; });
    el.innerHTML='<span class="ttl">Status</span>'+STATUSES.map(function(s){
      return '<span class="k" title="'+s+'"><span class="kb st-'+s+'" data-sw="'+s+'"></span>'+
        '<span class="g">'+SG[s]+'</span><span>'+s+'</span>'+
        '<span class="n">'+counts[s]+"</span></span>"; }).join("")+
      '<span class="grow"></span><span class="n">'+items.length+
      " item"+(items.length===1?"":"s")+"</span>";
    // Paint the key with the very same function the calendar uses.
    el.querySelectorAll("[data-sw]").forEach(function(sw){
      paint(sw,"#4a6b8a",sw.dataset.sw,""); });
  }

  /* ---------- direct manipulation: drag to move, drag an edge to reshape ----
   * Delegated on the persistent calendar root, so it survives every re-render.
   * On drop it calls opts.onEdit(id, startISO, endISO); the host decides how to
   * persist (write cells in the task pane, or message the pane from the dialog).
   * A drag sets Cal._suppressClick so the host's click→select doesn't also fire.
   */
  var EDGE = 10;  // px hit-zone at each end of a bar for resize
  function fmtNice(iso){ var d=parseD(iso); return d? d.getDate()+" "+MON[d.getMonth()].slice(0,3): iso; }

  /* Hover card. A native title tooltip is slow to appear, cannot be styled and
     buries the notes on a second line, so draw our own. */
  var hoverEl=null, hoverT=null;
  function hideHover(){ clearTimeout(hoverT); if(hoverEl){ hoverEl.remove(); hoverEl=null; } }
  function showHover(el,x,y){
    hideHover();
    var name=el.dataset.name||"(untitled)", st=el.dataset.status||"",
        s=el.dataset.start||"", e=el.dataset.end||"", notes=el.dataset.notes||"";
    var when = s ? ((e&&e!==s) ? fmtNice(s)+" → "+fmtNice(e) : fmtNice(s)) : "";
    var days = (s&&e&&e!==s) ? (daysBetween(parseD(s),parseD(e))+1)+" days" : "";
    var box=document.createElement("div"); box.className="hovercard";
    box.innerHTML='<div class="hc-t">'+esc(name)+'</div><div class="hc-m">'+esc(when)+
      (days?'<span class="hc-s"> · '+days+'</span>':'')+
      (st?'<span class="hc-s"> · '+esc(st)+'</span>':'')+
      (el.dataset.async?'<span class="hc-s"> · async</span>':'')+'</div>'+
      (notes?'<div class="hc-n">'+esc(notes)+'</div>':'');
    document.body.appendChild(box);
    var vw=window.innerWidth||1000, vh=window.innerHeight||800;
    var left=Math.min(x+14, vw-box.offsetWidth-10), top=y+18;
    if(top+box.offsetHeight>vh) top=Math.max(8, y-box.offsetHeight-12);
    box.style.left=Math.max(8,left)+"px"; box.style.top=top+"px";
    hoverEl=box;
  }

  function enableDrag(root, opts){
    opts = opts || {};
    var drag = null, tip = null, daySel = null;
    var canResize = opts.canResize || function(){ return true; };

    function paintRange(lo, hi){
      root.querySelectorAll(".day").forEach(function(c){
        c.classList.toggle("selrange", c.dataset.date >= lo && c.dataset.date <= hi);
      });
    }
    function clearRange(){
      root.querySelectorAll(".day.selrange").forEach(function(c){ c.classList.remove("selrange"); });
    }

    function weekUnder(x, y){
      var el = document.elementFromPoint(x, y);
      var w = el && el.closest ? el.closest(".week") : null;
      if (w) return w;
      var best = null, bestD = Infinity;
      root.querySelectorAll(".week").forEach(function(wk){
        var r = wk.getBoundingClientRect();
        var dy = y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0;
        var dx = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
        var d = dy * 4000 + dx;
        if (d < bestD) { bestD = d; best = wk; }
      });
      return best;
    }
    function dateAt(x, y){
      var wk = weekUnder(x, y); if (!wk || !wk.dataset.wkstart) return null;
      var r = wk.getBoundingClientRect();
      var idx = Math.floor((x - r.left) / (r.width / 7));
      idx = Math.max(0, Math.min(6, idx));
      return addDays(parseD(wk.dataset.wkstart), idx);
    }
    function zoneOf(bar, x){
      if (!canResize()) return "move";                        // no End column mapped
      var r = bar.getBoundingClientRect();
      // Milestones (single-day) have no continuation flags; allowing their edges
      // to resize is what lets you drag a one-day box out into a multi-day span.
      var msi = bar.classList.contains("msi");
      if ((msi || !bar.dataset.contl) && (x - r.left) <= EDGE) return "left";
      if ((msi || !bar.dataset.contr) && (r.right - x) <= EDGE) return "right";
      return "move";
    }
    function showTip(x, y, text){
      if (!tip) { tip = document.createElement("div"); tip.className = "dragtip"; document.body.appendChild(tip); }
      tip.textContent = text; tip.style.left = (x + 14) + "px"; tip.style.top = (y + 16) + "px";
    }
    function hideTip(){ if (tip) { tip.remove(); tip = null; } }

    // A small floating editor anchored to the clicked box. `cfg` = {title,value,save,extra}.
    function openPopup(anchor, cfg){
      closePopup();
      var r = anchor.getBoundingClientRect(), W = 250;
      var pop = document.createElement("div");
      pop.className = "cal-pop";
      pop.innerHTML =
        '<div class="cal-pop-t">' + esc(cfg.title) + '</div>' +
        '<input type="text" class="cal-pop-i">' +
        '<div class="cal-pop-r">' +
          (cfg.del ? '<button class="cal-pop-b danger" data-a="delete">Delete</button>' : '') +
          (cfg.extra ? '<button class="cal-pop-b" data-a="extra">' + esc(cfg.extra.label) + '</button>' : '') +
          '<span style="flex:1"></span>' +
          '<button class="cal-pop-b" data-a="cancel">Cancel</button>' +
          '<button class="cal-pop-b primary" data-a="save">Save</button>' +
        '</div>';
      document.body.appendChild(pop);
      var vw = window.innerWidth || 1000, vh = window.innerHeight || 800;
      var left = Math.min(Math.max(8, r.left), vw - W - 8);
      var top = r.bottom + 6;
      if (top + pop.offsetHeight > vh) top = Math.max(8, r.top - pop.offsetHeight - 6);
      pop.style.left = left + "px"; pop.style.top = top + "px"; pop.style.width = W + "px";
      var inp = pop.querySelector(".cal-pop-i");
      inp.value = cfg.value || ""; inp.focus(); inp.select();
      function commit(save){ closePopup(); if (save) cfg.save(inp.value); }
      inp.addEventListener("keydown", function(e){
        if (e.key === "Enter"){ e.preventDefault(); commit(true); }
        else if (e.key === "Escape"){ e.preventDefault(); commit(false); }
      });
      pop.addEventListener("mousedown", function(e){ e.stopPropagation(); });
      pop.addEventListener("click", function(e){
        var a = e.target.getAttribute && e.target.getAttribute("data-a");
        if (!a) return;
        if (a === "save") commit(true);
        else if (a === "cancel") commit(false);
        else if (a === "extra"){ closePopup(); if (cfg.extra) cfg.extra.run(); }
        else if (a === "delete"){
          // two-click confirm so a misclick can't delete a row
          if (e.target.dataset.armed){ closePopup(); if (cfg.del) cfg.del(); }
          else { e.target.dataset.armed = "1"; e.target.textContent = "Sure? Delete"; }
        }
      });
      global.Cal._pop = pop;
      setTimeout(function(){ document.addEventListener("mousedown", outsideClose, true); }, 0);
    }
    function outsideClose(e){
      var p = global.Cal._pop;
      if (p && !p.contains(e.target)) closePopup();
    }
    function closePopup(){
      document.removeEventListener("mousedown", outsideClose, true);
      var p = global.Cal._pop;
      if (p && p.parentNode) p.parentNode.removeChild(p);
      global.Cal._pop = null;
    }

    function computeDates(cur){
      var s = parseD(drag.origStart), e = drag.origEnd ? parseD(drag.origEnd) : null;
      if (drag.mode === "move") {
        var delta = daysBetween(drag.anchor, cur);
        return { start: fmtISO(addDays(s, delta)), end: e ? fmtISO(addDays(e, delta)) : "" };
      }
      if (drag.mode === "left") {
        var end = e || s;
        var ns = cur > end ? end : cur;
        return { start: fmtISO(ns), end: drag.origEnd };            // end unchanged
      }
      // right edge
      var ne = cur < s ? s : cur;
      return { start: drag.origStart, end: fmtISO(ne) === drag.origStart ? "" : fmtISO(ne) };
    }

    // hover: show the right cursor before the user commits
    root.addEventListener("mousemove", function(e){
      if (drag) return;
      var bar = e.target.closest ? e.target.closest(".bar,.msi") : null;
      if (!bar) return;
      var z = zoneOf(bar, e.clientX);
      bar.style.cursor = z === "move" ? "grab" : "ew-resize";
    });

    // Right-click a box (or an empty day) -> popup to add/edit its text.
    root.addEventListener("contextmenu", function(e){
      var bar = e.target.closest ? e.target.closest(".bar,.msi") : null;
      if (bar && bar.dataset.jump && opts.onRename){
        e.preventDefault();
        var bid = bar.dataset.jump;
        openPopup(bar, { title:"Edit text", value:bar.dataset.name||"",
          save:function(v){ if(v !== (bar.dataset.name||"")) opts.onRename(bid, v); },
          del: opts.onDelete ? function(){ opts.onDelete(bid); } : null,
          extra: opts.onSelect ? { label:"⤢ Excel", run:function(){ opts.onSelect(bid); } } : null });
        return;
      }
      var day = e.target.closest ? e.target.closest(".day") : null;
      if (day && day.dataset.date && opts.onAdd){
        e.preventDefault();
        openPopup(day, { title:"Add task on "+fmtNice(day.dataset.date), value:"",
          save:function(v){ if(v.trim()) opts.onAdd(day.dataset.date, v.trim()); } });
      }
    });

    root.addEventListener("mousedown", function(e){
      if (e.button !== 0) return;
      var bar = e.target.closest ? e.target.closest(".bar,.msi") : null;
      if (bar && bar.dataset.jump){
        e.preventDefault(); e.stopPropagation();
        drag = { id: bar.dataset.jump, bar: bar, mode: zoneOf(bar, e.clientX),
                 origStart: bar.dataset.start, origEnd: bar.dataset.end || "",
                 anchor: dateAt(e.clientX, e.clientY) || parseD(bar.dataset.start),
                 downX: e.clientX, moved: false, pending: null };
        bar.style.pointerEvents = "none";         // let dateAt see cells beneath
        document.body.classList.add("dragging");
        global.Cal._dragging = true;              // hosts pause refresh while this is set
        hideHover();
        return;
      }
      // Empty day: begin a create-selection. Drag across days for a span.
      var day = e.target.closest ? e.target.closest(".day") : null;
      if (day && day.dataset.date && opts.onCreate){
        e.preventDefault();
        daySel = { start: day.dataset.date, end: day.dataset.date, moved: false };
        global.Cal._dragging = true;
        hideHover();
        paintRange(day.dataset.date, day.dataset.date);
      }
    });

    document.addEventListener("mousemove", function(e){
      if (drag){
        var cur = dateAt(e.clientX, e.clientY); if (!cur) return;
        if (Math.abs(e.clientX - drag.downX) > 3) drag.moved = true;
        var d = computeDates(cur); drag.pending = d;
        paintRange(d.start, d.end || d.start);
        var verb = drag.mode === "move" ? "" : (drag.mode === "left" ? "start " : "end ");
        showTip(e.clientX, e.clientY, verb + fmtNice(d.start) + (d.end ? " → " + fmtNice(d.end) : ""));
        return;
      }
      if (daySel){
        var c2 = dateAt(e.clientX, e.clientY); if (!c2) return;
        daySel.end = fmtISO(c2); daySel.moved = true;
        var lo = daySel.start < daySel.end ? daySel.start : daySel.end;
        var hi = daySel.start < daySel.end ? daySel.end : daySel.start;
        paintRange(lo, hi);
        showTip(e.clientX, e.clientY, fmtNice(lo) + (lo !== hi ? " → " + fmtNice(hi) : ""));
      }
    });

    document.addEventListener("mouseup", function(){
      if (drag){
        var dr = drag; drag = null;
        document.body.classList.remove("dragging");
        global.Cal._dragging = false;
        dr.bar.style.pointerEvents = "";
        clearRange(); hideTip();
        if (dr.moved && dr.pending &&
            (dr.pending.start !== dr.origStart || dr.pending.end !== dr.origEnd)) {
          global.Cal._suppressClick = true;       // don't also fire click→select
          if (opts.onEdit) opts.onEdit(dr.id, dr.pending.start, dr.pending.end);
        }
        return;
      }
      if (daySel){
        var s = daySel.start, e2 = daySel.end; daySel = null;
        global.Cal._dragging = false;
        clearRange(); hideTip();
        var lo = s < e2 ? s : e2, hi = s < e2 ? e2 : s;
        if (opts.onCreate) opts.onCreate(lo, lo === hi ? "" : hi);
      }
    });

    /* hover card on bars and milestones */
    root.addEventListener("mouseover", function(e){
      if (drag || daySel) return;
      var b = e.target.closest ? e.target.closest(".bar,.msi") : null;
      if (!b || !b.dataset.jump) return;
      var x=e.clientX, y=e.clientY;
      clearTimeout(hoverT); hoverT=setTimeout(function(){ showHover(b,x,y); }, 160);
    });
    root.addEventListener("mouseout", function(e){
      if (e.target.closest && e.target.closest(".bar,.msi")) hideHover();
    });
    root.addEventListener("mousedown", hideHover, true);

    /* Escape abandons an in-flight move/resize or day selection */
    document.addEventListener("keydown", function(e){
      if (e.key !== "Escape") return;
      if (drag){
        var dr=drag; drag=null;
        document.body.classList.remove("dragging");
        global.Cal._dragging=false; global.Cal._suppressClick=true;
        dr.bar.style.pointerEvents=""; clearRange(); hideTip();
      } else if (daySel){
        daySel=null; global.Cal._dragging=false; clearRange(); hideTip();
      } else return;
      e.preventDefault();
    });
  }

  global.Cal = {
    render: render, statusKey: statusKey, applyDensity: applyDensity,
    balanceWeeks: balanceWeeks, paint: paint, ink: ink, textOn: textOn, esc: esc,
    parseD: parseD, fmtISO: fmtISO, today0: today0, addDays: addDays,
    daysBetween: daysBetween, iStart: iStart, iEnd: iEnd, isSpan: isSpan,
    enableDrag: enableDrag, _suppressClick: false, _dragging: false,
    monthsOf: monthsOf, scrollToToday: scrollToToday, colOf: colOf, isoWeek: isoWeek,
    STATUSES: STATUSES, SG: SG, MON: MON, DENSITY: DENSITY
  };
})(window);
