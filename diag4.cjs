const XLSX = require('xlsx');
const wb   = XLSX.readFile('C:/Users/LCK_KATHIA/Desktop/PRUEBA 1 (4).xlsx');

// ── DS ───────────────────────────────────────────────────────────────────────
const dsRaw = XLSX.utils.sheet_to_json(wb.Sheets['DS'], {header:1, defval:''});
const dHdr  = dsRaw[0].map(c=>String(c).trim());
const iDF   = dHdr.indexOf('Fraccion');
const iDS   = dHdr.indexOf('SecuenciaFraccion');
const iDC   = dHdr.indexOf('CantidadUMComercial');
const iDV   = dHdr.findIndex(h=>h==='Valor usd redondeado');
const iDP2  = dHdr.indexOf('Pedimento2');
const iDPais= dHdr.indexOf('PaisOrigenDestino');
const iDDesc= dHdr.indexOf('DescripcionMercancia');

const nFrac = s => String(s||'').replace(/\./g,'').trim();
const normS = s => String(s||'').trim().toUpperCase();
const nDesc = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');

const dsRows = [];
for(let i=1;i<dsRaw.length;i++){
  const r=dsRaw[i];
  const c=parseFloat(r[iDC])||0;
  if(c===0) continue;
  dsRows.push({ ped2:normS(r[iDP2]), frac:nFrac(r[iDF]), sec:r[iDS],
    cant:c, val:parseFloat(r[iDV])||0,
    pais:normS(r[iDPais]), desc:nDesc(r[iDDesc]) });
}

// ── Layout ────────────────────────────────────────────────────────────────────
const lyRaw = XLSX.utils.sheet_to_json(wb.Sheets['Layout'], {header:1, defval:''});
const lHdr  = lyRaw[0].map(c=>String(c).trim());
const iLF   = lHdr.indexOf('FraccionNico');
const iLC   = lHdr.indexOf('cantidad_comercial');
const iLV   = lHdr.indexOf('=REDONDEAR(FR1,0)');
const iLP   = lHdr.indexOf('pais_origen PARA LAYOUT');
const iLD   = lHdr.indexOf('descripcion');
const iLPed = lHdr.indexOf('pedimento');
const iLSec = lHdr.indexOf('SEC CALC');

const lyRows = [];
for(let i=1;i<lyRaw.length;i++){
  const r=lyRaw[i];
  const c=parseFloat(r[iLC])||0;
  if(c===0) continue;
  lyRows.push({ ped:normS(r[iLPed]), frac:nFrac(r[iLF]),
    cant:c, val:parseFloat(r[iLV])||0,
    pais:normS(r[iLP]), desc:nDesc(r[iLD]), sec:String(r[iLSec]||'') });
}

// ── Totales globales ─────────────────────────────────────────────────────────
const dsTotC = dsRows.reduce((a,r)=>a+r.cant,0);
const dsTotV = dsRows.reduce((a,r)=>a+r.val,0);
const lyTotC = lyRows.reduce((a,r)=>a+r.cant,0);
const lyTotV = lyRows.reduce((a,r)=>a+r.val,0);
console.log(`DS  Cant=${dsTotC.toLocaleString()} Val=$${dsTotV.toFixed(0)}`);
console.log(`LY  Cant=${lyTotC.toLocaleString()} Val=$${lyTotV.toFixed(0)}`);
console.log(`DIFF Cant=${lyTotC-dsTotC} Val=$${(lyTotV-dsTotV).toFixed(0)}`);
console.log(`DS rows=${dsRows.length}  LY rows=${lyRows.length}\n`);

// ── Agrupar Layout por Ped+Frac ───────────────────────────────────────────────
const lyByPF = new Map();
for(const r of lyRows){
  const k=r.ped+'|||'+r.frac;
  if(!lyByPF.has(k)) lyByPF.set(k,[]);
  lyByPF.get(k).push(r);
}

// ── Agrupar DS por Ped+Frac ───────────────────────────────────────────────────
const dsByPF = new Map();
for(const r of dsRows){
  const k=r.ped2+'|||'+r.frac;
  if(!dsByPF.has(k)) dsByPF.set(k,[]);
  dsByPF.get(k).push(r);
}

// ── Simular matching simple para detectar qué DS no matchea ──────────────────
// Para cada Ped+Frac: comparar suma DS vs suma Layout
let problemCount=0;
for(const [k, dsList] of dsByPF){
  const lyList = lyByPF.get(k)||[];
  const lyCantTotal = lyList.reduce((a,r)=>a+r.cant,0);
  const lyValTotal  = lyList.reduce((a,r)=>a+r.val,0);
  const dsCantTotal = dsList.reduce((a,r)=>a+r.cant,0);
  const dsValTotal  = dsList.reduce((a,r)=>a+r.val,0);
  const diffC = Math.abs(lyCantTotal - dsCantTotal);
  const diffV = Math.abs(lyValTotal  - dsValTotal);

  // Marcar como problema si hay más de 1 DS sec Y el layout no suma igual a cada sec individual
  if(dsList.length===1 && diffC>1){
    problemCount++;
    console.log(`[SIN LAYOUT?] ${k} — DS Sec${dsList[0].sec} Cant=${dsList[0].cant} | LY rows=${lyList.length} LY cant=${lyCantTotal} diff=${lyCantTotal-dsList[0].cant}`);
  }
  if(dsList.length>1){
    // Verificar si layout total coincide con DS total
    if(diffC>1){
      problemCount++;
      console.log(`[TOTAL NO COINCIDE] ${k} — DS total=${dsCantTotal} LY total=${lyCantTotal} diff=${lyCantTotal-dsCantTotal}`);
      for(const d of dsList) console.log(`  DS Sec${d.sec} C=${d.cant} V=${d.val} Pais=${d.pais} Desc=${d.desc.slice(0,30)}`);
      const lyByDesc = new Map();
      for(const r of lyList){ const dk=r.desc; if(!lyByDesc.has(dk)) lyByDesc.set(dk,[]); lyByDesc.get(dk).push(r); }
      for(const [dk,rows] of lyByDesc){
        const sc=rows.reduce((a,r)=>a+r.cant,0), sv=rows.reduce((a,r)=>a+r.val,0);
        console.log(`  LY desc="${dk.slice(0,30)}" rows=${rows.length} C=${sc} V=${sv}`);
      }
    } else {
      // Total coincide pero hay múltiples secs — mostrar si son problemáticas
      let anyProb=false;
      for(const d of dsList){
        const byDesc = lyList.filter(r=>r.desc===d.desc);
        const byDescP= byDesc.filter(r=>r.pais===d.pais);
        const matchDP = byDescP.reduce((a,r)=>a+r.cant,0);
        const matchD  = byDesc.reduce((a,r)=>a+r.cant,0);
        if(Math.abs(matchDP-d.cant)>1 && Math.abs(matchD-d.cant)>1){
          anyProb=true;
        }
      }
      if(anyProb){
        problemCount++;
        console.log(`[MULTI-SEC DIFICIL] ${k} — ${dsList.length} DS secs, LY total=${lyCantTotal} rows=${lyList.length}`);
        for(const d of dsList) console.log(`  DS Sec${d.sec} C=${d.cant} V=${d.val} Pais=${d.pais} Desc=${d.desc.slice(0,30)}`);
        const lyByDesc = new Map();
        for(const r of lyList){ const dk=r.desc.slice(0,30); if(!lyByDesc.has(dk)) lyByDesc.set(dk,[]); lyByDesc.get(dk).push(r); }
        for(const [dk,rows] of lyByDesc){
          const sc=rows.reduce((a,r)=>a+r.cant,0), sv=rows.reduce((a,r)=>a+r.val,0);
          const byPais = new Map();
          for(const r of rows){ if(!byPais.has(r.pais)) byPais.set(r.pais,{c:0,v:0,n:0}); const g=byPais.get(r.pais); g.c+=r.cant; g.v+=r.val; g.n++; }
          console.log(`  LY desc="${dk}" rows=${rows.length} C=${sc} V=${sv.toFixed(0)}`);
          for(const [p,g] of byPais) console.log(`    Pais=${p} n=${g.n} C=${g.c} V=${g.v.toFixed(0)}`);
        }
      }
    }
  }
}
console.log(`\nTotal grupos problema: ${problemCount}`);

// ── Contar DS que definitivamente no tienen Layout ────────────────────────────
let noLayout=0;
for(const [k,dsList] of dsByPF){
  if(!lyByPF.has(k)){
    noLayout+=dsList.length;
    console.log(`[NO LAYOUT] ${k} — ${dsList.length} secs`);
  }
}
console.log(`DS secs sin fracción en Layout: ${noLayout}`);
