/* Excel data layer for the Timetable add-in.
 *
 * There is no sync and no second copy of the data: the worksheet IS the state.
 * Every read and write goes through Excel.run straight to cells, so edits made
 * in the calendar and edits typed into the sheet are the same operation.
 *
 * The configuration lives in Office.context.document.settings, which is stored
 * inside the workbook - so the column mapping travels with the file and a
 * colleague who opens it gets a working calendar with no setup.
 */
(function (global) {
  "use strict";

  var FIELDS = ["id","name","start","end","status","color","key","async","notes"];
  var REQUIRED = ["name","start"];
  var PALETTE = ["#2e86c1","#8e44ad","#27ae60","#e67e22","#c0392b",
                 "#16a085","#d4ac0d","#7b241c","#e74c3c","#2c3e50"];
  var NAMED = {blue:"#2e86c1",orange:"#e67e22",teal:"#16a085",purple:"#8e44ad",
               green:"#27ae60",gold:"#d4ac0d",red:"#c0392b",yellow:"#d4ac0d",
               grey:"#95a5a6",gray:"#95a5a6",black:"#2c3e50"};
  var SETTING = "timetableCalendar.v1";
  var EPOCH = Date.UTC(1899,11,30);

  /* ---- coercion (Excel is loosely typed) ---- */
  function serialToISO(n){
    var d=new Date(EPOCH+Math.round(n)*86400000);
    return d.getUTCFullYear()+"-"+String(d.getUTCMonth()+1).padStart(2,"0")+
           "-"+String(d.getUTCDate()).padStart(2,"0");
  }
  function isoToSerial(s){
    var p=String(s).slice(0,10).split("-").map(Number);
    if(!p[0]) return null;
    return Math.round((Date.UTC(p[0],p[1]-1,p[2])-EPOCH)/86400000);
  }
  var DATE_RE=[/^(\d{4})-(\d{1,2})-(\d{1,2})/,null];
  function toISO(v){
    if(v===null||v===undefined||v==="") return "";
    if(typeof v==="number") return isFinite(v)?serialToISO(v):"";
    var s=String(v).trim();
    var m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m) return m[1]+"-"+String(m[2]).padStart(2,"0")+"-"+String(m[3]).padStart(2,"0");
    m=s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
    if(m) return m[3]+"-"+String(m[2]).padStart(2,"0")+"-"+String(m[1]).padStart(2,"0");
    var d=new Date(s);
    return isNaN(d)?"":(d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+
                        "-"+String(d.getDate()).padStart(2,"0"));
  }
  function toBool(v){
    if(typeof v==="boolean") return v;
    return ["yes","true","y","1","x","ano","ja"].indexOf(String(v||"").trim().toLowerCase())>=0;
  }
  /* Order matters: "not started" must not be read as "started". */
  var PEND=["not started","nezaháj","nezahaj","not yet","planned","pending","to do",
            "todo","waiting","on hold","čeká","ceka","plán","plan","backlog"];
  var DONE=["done","complete","hotov","closed","finish","dokonč","dokonc","ukonč",
            "ukonc","delivered","100%","splněn","splnen"];
  var ONG =["ongoing","in progress","progress","wip","running","active","started",
            "probíh","probih","rozpracov","běží","bezi","current","underway"];
  function has(s,arr){ return arr.some(function(w){ return s.indexOf(w)>=0; }); }
  function toStatus(v){
    var s=String(v||"").trim().toLowerCase();
    if(!s) return "Pending";
    if(has(s,PEND)) return "Pending";
    if(has(s,DONE)) return "Done";
    if(has(s,ONG)) return "Ongoing";
    return "Pending";
  }
  function toColor(v,fallback){
    var s=String(v||"").trim();
    if(!s) return fallback;
    if(s.charAt(0)==="#"&&s.length===7) return s.toLowerCase();
    if(/^[0-9a-f]{6}$/i.test(s)) return "#"+s.toLowerCase();
    return NAMED[s.toLowerCase()]||fallback;
  }
  function newId(){ return "T"+(Date.now()%0xFFFFFFF).toString(16)+
    Math.floor(Math.random()*900+100); }
  function colName(n){ var s=""; while(n>0){ var r=(n-1)%26; s=String.fromCharCode(65+r)+s;
    n=Math.floor((n-1)/26); } return s; }

  /* ---- config, stored inside the workbook ---- */
  function defaults(){
    return { sheet:"", table:"", mapping:{}, density:"comfortable",
             perRow:"auto", showDates:false, hidden:[] };
  }
  function loadConfig(){
    var raw=null;
    try { raw=Office.context.document.settings.get(SETTING); } catch(e){}
    var cfg=defaults();
    if(raw&&typeof raw==="object"){
      Object.keys(cfg).forEach(function(k){ if(raw[k]!==undefined) cfg[k]=raw[k]; });
    }
    FIELDS.forEach(function(f){ if(cfg.mapping[f]===undefined) cfg.mapping[f]=""; });
    return cfg;
  }
  function saveConfig(cfg){
    return new Promise(function(res,rej){
      try{
        Office.context.document.settings.set(SETTING,cfg);
        Office.context.document.settings.saveAsync(function(r){
          if(r.status===Office.AsyncResultStatus.Succeeded) res(true);
          else rej(new Error(r.error?r.error.message:"could not save settings"));
        });
      }catch(e){ rej(e); }
    });
  }

  /* Annotated-view design, keyed by sheet, stored inside the workbook so it
     survives cache clears and travels with the file. */
  var ANNO_SETTING = "timetableAnno.v1";
  function loadDesign(key){
    try { var all=Office.context.document.settings.get(ANNO_SETTING); return (all&&all[key])||null; }
    catch(e){ return null; }
  }
  function saveDesign(key, json){
    return new Promise(function(res,rej){
      try{
        var all=Office.context.document.settings.get(ANNO_SETTING)||{};
        all[key]=json; Office.context.document.settings.set(ANNO_SETTING, all);
        Office.context.document.settings.saveAsync(function(r){
          if(r.status===Office.AsyncResultStatus.Succeeded) res(true);
          else rej(new Error(r.error?r.error.message:"could not save the design"));
        });
      }catch(e){ rej(e); }
    });
  }

  /* ---- locating the data ---- */
  function isCellRef(s){ return /^[A-Za-z]{1,3}[0-9]{1,7}$/.test(String(s||"").trim()); }

  /* Candidate header rows anywhere on a sheet: a real timetable rarely starts
     at A1 - there may be a title block, a logo, blank spacer rows. */
  function findAnchors(values, r0, c0, limit){
    var out=[];
    for(var ri=0; ri<Math.min(values.length,40); ri++){
      var row=values[ri], bestStart=null, bestLen=0, curStart=null, curLen=0;
      for(var ci=0; ci<Math.min(row.length,60); ci++){
        var v=row[ci];
        var isText = v!==null && v!==undefined && String(v).trim()!=="" &&
                     typeof v!=="number" && typeof v!=="boolean";
        if(isText){
          if(curStart===null) curStart=ci;
          curLen++;
          if(curLen>bestLen){ bestLen=curLen; bestStart=curStart; }
        } else { curStart=null; curLen=0; }
      }
      if(bestLen>=3){
        out.push({ cell: colName(c0+bestStart+1)+(r0+ri+1),
                   headers: row.slice(bestStart,bestStart+bestLen).map(function(x){
                     return String(x).trim(); }),
                   cols: bestLen });
        if(out.length>=(limit||4)) break;
      }
    }
    return out;
  }

  function listSources(){
    return Excel.run(function(ctx){
      var sheets=ctx.workbook.worksheets;
      sheets.load("items/name");
      var wb=ctx.workbook; wb.load("name");
      return ctx.sync().then(function(){
        var jobs=sheets.items.map(function(sh){
          var tables=sh.tables; tables.load("items/name");
          var used=sh.getUsedRangeOrNullObject();
          used.load("values,rowIndex,columnIndex,isNullObject");
          return { sh:sh, tables:tables, used:used };
        });
        return ctx.sync().then(function(){
          // getRange() covers header + body in one go. Note there is no
          // getDataBodyRangeOrNullObject() on Table - using it threw and took
          // the whole source listing down with it.
          var more=jobs.map(function(j){
            return j.tables.items.map(function(t){
              var r=t.getRange(); r.load("values,rowCount");
              return { t:t, r:r };
            });
          });
          return ctx.sync().then(function(){
            return { workbook: wb.name, sheets: jobs.map(function(j,i){
              return {
                name: j.sh.name,
                tables: more[i].map(function(x){
                  var vals=x.r.values||[];
                  return { name:x.t.name,
                           headers:(vals[0]||[]).map(function(v){
                             return String(v==null?"":v).trim(); }),
                           rows: Math.max(0,(x.r.rowCount||0)-1) };
                }),
                anchors: j.used.isNullObject?[]:
                  findAnchors(j.used.values, j.used.rowIndex, j.used.columnIndex)
              };
            })};
          });
        });
      });
    });
  }

  function headersAt(sheetName, cell){
    return Excel.run(function(ctx){
      var sh=ctx.workbook.worksheets.getItem(sheetName);
      var a=sh.getRange(cell); a.load("rowIndex,columnIndex");
      var used=sh.getUsedRangeOrNullObject();
      used.load("rowIndex,columnIndex,rowCount,columnCount,isNullObject");
      return ctx.sync().then(function(){
        if(used.isNullObject) return { cell:cell, headers:[] };
        var width=used.columnIndex+used.columnCount-a.columnIndex;
        if(width<=0) return { cell:cell, headers:[] };
        var r=sh.getRangeByIndexes(a.rowIndex,a.columnIndex,1,width);
        r.load("values");
        return ctx.sync().then(function(){
          var out=[];
          (r.values[0]||[]).forEach(function(v){
            if(out.stop) return;
            if(v===null||v===undefined||String(v).trim()===""){ out.stop=true; return; }
            out.push(String(v).trim());
          });
          return { cell:cell, headers:out };
        });
      });
    });
  }

  /* Resolve the mapping to column offsets and grab the body values. */
  function grid(ctx, cfg){
    var sh=ctx.workbook.worksheets.getItem(cfg.sheet);
    if(cfg.table && !isCellRef(cfg.table)){
      var t=sh.tables.getItem(cfg.table);
      // One range for header + body: Table has no getDataBodyRangeOrNullObject,
      // and getDataBodyRange() itself throws on an empty table.
      var r=t.getRange();
      r.load("values,rowIndex,columnIndex,columnCount,rowCount");
      return ctx.sync().then(function(){
        var vals=r.values||[];
        return { sheet:sh, table:t, isTable:true,
                 headers:(vals[0]||[]).map(function(v){ return String(v==null?"":v).trim(); }),
                 firstRow:r.rowIndex+1, firstCol:r.columnIndex, ncols:r.columnCount,
                 values:vals.slice(1) };
      });
    }
    var a=sh.getRange(cfg.table||"A1"); a.load("rowIndex,columnIndex");
    var used=sh.getUsedRangeOrNullObject();
    used.load("rowIndex,columnIndex,rowCount,columnCount,isNullObject");
    return ctx.sync().then(function(){
      if(used.isNullObject) throw new Error("That sheet is empty.");
      var width=used.columnIndex+used.columnCount-a.columnIndex;
      var hr=sh.getRangeByIndexes(a.rowIndex,a.columnIndex,1,Math.max(width,1));
      hr.load("values");
      return ctx.sync().then(function(){
        var hs=[];
        (hr.values[0]||[]).some(function(v){
          if(v===null||v===undefined||String(v).trim()==="") return true;
          hs.push(String(v).trim()); return false;
        });
        if(!hs.length) throw new Error("No header text found at "+(cfg.table||"A1")+".");
        var nrows=used.rowIndex+used.rowCount-(a.rowIndex+1);
        if(nrows<0) nrows=0;
        if(!nrows) return { sheet:sh, isTable:false, headers:hs,
                            firstRow:a.rowIndex+1, firstCol:a.columnIndex,
                            ncols:hs.length, values:[] };
        var body=sh.getRangeByIndexes(a.rowIndex+1,a.columnIndex,nrows,hs.length);
        body.load("values");
        return ctx.sync().then(function(){
          return { sheet:sh, isTable:false, headers:hs,
                   firstRow:a.rowIndex+1, firstCol:a.columnIndex,
                   ncols:hs.length, values:body.values };
        });
      });
    });
  }

  function resolveCols(g, cfg){
    var norm=g.headers.map(function(h){ return String(h||"").trim(); });
    var lower=norm.map(function(h){ return h.toLowerCase(); });
    var col={}, missing=[];
    FIELDS.forEach(function(f){
      var want=String(cfg.mapping[f]||"").trim();
      if(!want) return;
      var i=norm.indexOf(want);
      if(i<0) i=lower.indexOf(want.toLowerCase());
      if(i>=0) col[f]=i; else missing.push(f+" → '"+want+"'");
    });
    REQUIRED.forEach(function(f){
      if(col[f]===undefined)
        throw new Error("'"+f+"' must be mapped to an existing column. Columns found: "+norm.join(", "));
    });
    return { col:col, missing:missing };
  }

  function buildItems(g, col){
    var items=[], blanks=[], missingIds=[];
    g.values.forEach(function(row,r){
      function get(f){ return col[f]!==undefined?row[col[f]]:null; }
      var name=String(get("name")==null?"":get("name")).trim();
      var start=toISO(get("start"));
      var rid=String(get("id")==null?"":get("id")).trim();
      if(!name&&!start&&!rid){ blanks.push(r); return; }
      if(col.id!==undefined){ if(!rid){ rid=newId()+r; missingIds.push([r,rid]); } }
      else rid="row"+r;
      items.push({ id:rid, row:r, name:name, start:start, end:toISO(get("end")),
        status: col.status!==undefined?toStatus(get("status")):"Pending",
        color: toColor(get("color"), PALETTE[r%PALETTE.length]),
        key: toBool(get("key")), async: toBool(get("async")),
        notes: String(get("notes")==null?"":get("notes")).trim() });
    });
    return { items:items, blanks:blanks, missingIds:missingIds };
  }

  var lastCols={}, lastWarning="";

  function read(cfg){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg);
        var built=buildItems(g,rc.col);
        lastCols=rc.col;
        var warn=[];
        if(rc.missing.length) warn.push("Unmapped: "+rc.missing.join("; "));
        if(rc.col.id===undefined) warn.push("No ID column mapped — rows are addressed "+
          "by position, so re-sorting the sheet while the calendar is open can retarget an edit.");
        lastWarning=warn.join(" · ");
        // Backfill IDs for rows typed straight into the sheet.
        built.missingIds.forEach(function(pair){
          g.sheet.getCell(g.firstRow+pair[0], g.firstCol+rc.col.id).values=[[pair[1]]];
        });
        return ctx.sync().then(function(){
          return { items:built.items, blanks:built.blanks, headers:g.headers,
                   warning:lastWarning, available:Object.keys(rc.col) };
        });
      });
    });
  }

  function writeCell(g, col, row, field, value){
    var cell=g.sheet.getCell(g.firstRow+row, g.firstCol+col[field]);
    if(field==="start"||field==="end"){
      var s=isoToSerial(value);
      if(s===null||value===""){ cell.clear(Excel.ClearApplyTo.contents); }
      else { cell.values=[[s]]; cell.numberFormat=[["yyyy-mm-dd"]]; }
    } else if(field==="key"||field==="async"){
      cell.values=[[value?"Yes":""]];
    } else if(field==="status"){
      cell.values=[[toStatus(value)]];
    } else {
      cell.values=[[value]];
    }
  }

  function findRow(g, col, id){
    var built=buildItems(g,col);
    for(var i=0;i<built.items.length;i++)
      if(built.items[i].id===id) return built.items[i].row;
    throw new Error("That task no longer exists in the sheet.");
  }

  function update(cfg,id,field,value){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg);
        if(rc.col[field]===undefined)
          throw new Error("The '"+field+"' field is not mapped to a column in this table.");
        writeCell(g,rc.col,findRow(g,rc.col,id),field,value);
        return ctx.sync();
      });
    });
  }

  /* Make sure a Colour column exists and is mapped, adding one if needed, so a
     right-click recolour in the annotated view lands in the sheet (and colours
     the calendar view too). */
  function ensureColorColumn(cfg){
    if(cfg.mapping.color) return Promise.resolve(false);
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var lower=g.headers.map(function(h){ return String(h).toLowerCase().trim(); });
        var idx=lower.indexOf("colour"); if(idx<0) idx=lower.indexOf("color");
        if(idx>=0){ cfg.mapping.color=g.headers[idx]; return ctx.sync().then(function(){ return false; }); }
        if(g.isTable){
          g.table.columns.add(null, null, "Colour"); cfg.mapping.color="Colour";
          return ctx.sync().then(function(){ return true; });
        }
        // plain range: write a header cell just past the last column
        g.sheet.getCell(g.firstRow-1, g.firstCol+g.ncols).values=[["Colour"]];
        cfg.mapping.color="Colour";
        return ctx.sync().then(function(){ return true; });
      });
    }).then(function(added){ return saveConfig(cfg).then(function(){ return added; }); });
  }
  function setColor(cfg,id,value){
    return ensureColorColumn(cfg).then(function(){ return update(cfg,id,"color",value); });
  }

  // Move / resize writes both dates in one round trip so the bar never lands
  // half-updated. `end` of "" clears the End cell (task becomes single-day).
  function setDates(cfg,id,start,end){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg), row=findRow(g,rc.col,id);
        writeCell(g,rc.col,row,"start",start);
        if(rc.col.end!==undefined) writeCell(g,rc.col,row,"end",end||"");
        return ctx.sync();
      });
    });
  }

  function add(cfg,start,end,name){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg), col=rc.col;
        var built=buildItems(g,col);
        var target = built.blanks.length ? built.blanks[0] : g.values.length;
        if(g.isTable && !built.blanks.length){
          g.table.rows.add(null,[new Array(g.ncols).fill("")]);
        }
        var rid=newId();
        if(col.id!==undefined) writeCell(g,col,target,"id",rid);
        if(col.status!==undefined) writeCell(g,col,target,"status","Pending");
        if(col.color!==undefined)
          writeCell(g,col,target,"color",PALETTE[built.items.length%PALETTE.length]);
        writeCell(g,col,target,"name",name||"");
        writeCell(g,col,target,"start",start||Cal.fmtISO(Cal.today0()));
        if(col.end!==undefined) writeCell(g,col,target,"end",end||"");
        return ctx.sync().then(function(){ return { id:rid }; });
      });
    });
  }

  // Re-create a previously-deleted row with all of its mapped fields, so undo
  // brings the task back exactly as it was (same id when an ID column exists).
  function addFull(cfg,it){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg), col=rc.col;
        var built=buildItems(g,col);
        var target = built.blanks.length ? built.blanks[0] : g.values.length;
        if(g.isTable && !built.blanks.length){
          g.table.rows.add(null,[new Array(g.ncols).fill("")]);
        }
        if(col.id!==undefined) writeCell(g,col,target,"id", it.id||newId());
        ["name","status","color","notes","start","end","key","async"].forEach(function(f){
          if(col[f]!==undefined) writeCell(g,col,target,f, it[f]);
        });
        return ctx.sync().then(function(){ return { id: it.id }; });
      });
    });
  }

  function remove(cfg,id){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg);
        var row=findRow(g,rc.col,id);
        if(g.isTable){
          g.table.rows.getItemAt(row).delete();
        } else {
          // Only clear the mapped span - we cannot know what lives beside it.
          g.sheet.getRangeByIndexes(g.firstRow+row,g.firstCol,1,g.ncols)
                 .clear(Excel.ClearApplyTo.contents);
        }
        return ctx.sync();
      });
    });
  }

  function sortByDate(cfg){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg), col=rc.col;
        var built=buildItems(g,col);
        var order=built.items.slice().sort(function(a,b){
          var x=a.start||"9999-99-99", y=b.start||"9999-99-99";
          return x<y?-1:x>y?1:a.name.toLowerCase()<b.name.toLowerCase()?-1:1;
        });
        order.forEach(function(it,newRow){
          ["id","name","status","color","notes"].forEach(function(f){
            if(col[f]!==undefined) writeCell(g,col,newRow,f,it[f]); });
          ["key","async"].forEach(function(f){
            if(col[f]!==undefined) writeCell(g,col,newRow,f,it[f]); });
          ["start","end"].forEach(function(f){
            if(col[f]!==undefined) writeCell(g,col,newRow,f,it[f]); });
        });
        return ctx.sync();
      });
    });
  }

  function selectRow(cfg,id){
    return Excel.run(function(ctx){
      return grid(ctx,cfg).then(function(g){
        var rc=resolveCols(g,cfg);
        var row=findRow(g,rc.col,id);
        g.sheet.activate();
        g.sheet.getRangeByIndexes(g.firstRow+row,g.firstCol,1,g.ncols).select();
        return ctx.sync();
      });
    });
  }

  /* Sheet edits push straight back into the calendar. */
  function watch(cfg, cb){
    return Excel.run(function(ctx){
      var sh=ctx.workbook.worksheets.getItem(cfg.sheet);
      sh.onChanged.add(function(){ cb(); return Promise.resolve(); });
      ctx.workbook.worksheets.onActivated.add(function(){ cb(); return Promise.resolve(); });
      return ctx.sync();
    });
  }

  global.Xl = {
    FIELDS:FIELDS, REQUIRED:REQUIRED, PALETTE:PALETTE,
    loadConfig:loadConfig, saveConfig:saveConfig, defaults:defaults,
    loadDesign:loadDesign, saveDesign:saveDesign,
    listSources:listSources, headersAt:headersAt, isCellRef:isCellRef,
    read:read, update:update, setColor:setColor, setDates:setDates, add:add, addFull:addFull, remove:remove,
    sortByDate:sortByDate, selectRow:selectRow, watch:watch,
    toStatus:toStatus, toISO:toISO, isoToSerial:isoToSerial, serialToISO:serialToISO
  };
})(window);
