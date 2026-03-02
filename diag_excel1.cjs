const XLSX=require('xlsx');
const wb=XLSX.readFile('C:/Users/LCK_KATHIA/Desktop/excel1.xlsx');
const dsR=XLSX.utils.sheet_to_json(wb.Sheets['DS'],{header:1,defval:''});
const lyR=XLSX.utils.sheet_to_json(wb.Sheets['Layout'],{header:1,defval:''});
const lH=lyR[1].map(c=>String(c).trim()); // header en fila 1, datos desde fila 2
const dH=dsR[0].map(c=>String(c).trim());

const iF=dH.indexOf('Fraccion'),iSec=dH.indexOf('SecuenciaFraccion'),iC=dH.indexOf('CantidadUMComercial');
const iV=dH.findIndex(h=>h==='Valor usd redondeado'),iP2=dH.indexOf('Pedimento2'),iPais=dH.indexOf('PaisOrigenDestino');
const iDesc=dH.indexOf('DescripcionMercancia'),iCand=dH.indexOf('Candado 551');
const iLSec=lH.indexOf('SEC CALC'),iLFrac=lH.indexOf('FraccionNico');
const iLC=lH.indexOf('cantidad_comercial');
const iLV=26; // =REDONDEAR(FR1,0)
const iLP=lH.indexOf('pais_origen PARA LAYOUT'),iLD=lH.indexOf('descripcion'),iLPed=lH.indexOf('pedimento');

const nF=s=>String(s||'').replace(/\./g,'').trim();
const nS=s=>String(s||'').trim().toUpperCase();

// Construir DS
const dsAll=[], dsByCand=new Map(), dsByPF=new Map();
for(let i=1;i<dsR.length;i++){
  const r=dsR[i];
  const c=parseFloat(r[iC])||0; if(c===0) continue;
  const ped2=nS(r[iP2]),frac=nF(r[iF]),sec=nS(r[iSec]);
  const cand=nS(r[iCand]);
  const obj={ped2,frac,sec,cant:c,val:parseFloat(r[iV])||0,pais:nS(r[iPais]),
             desc:String(r[iDesc]||'').toLowerCase().trim(),cand,i};
  dsAll.push(obj);
  if(cand) dsByCand.set(cand,obj);
  const kPF=ped2+'|||'+frac; if(!dsByPF.has(kPF)) dsByPF.set(kPF,[]); dsByPF.get(kPF).push(obj);
}

// Simular E0 para cada Layout row
const usedDS=new Set();
const results=[];
for(let i=2;i<lyR.length;i++){
  const r=lyR[i];
  const cant=parseFloat(r[iLC])||0; if(cant===0) continue;
  const sec=nS(r[iLSec]),frac=nF(r[iLFrac]),ped=nS(r[iLPed]);
  const val=parseFloat(r[iLV])||0;
  const isReal = sec && sec!=='.' && !isNaN(parseFloat(sec));

  let matched=null, how='';
  // E0 candado
  const cand=ped+'-'+frac+'-'+sec;
  if(isReal && dsByCand.has(cand) && !usedDS.has(dsByCand.get(cand).i)){
    matched=dsByCand.get(cand); how='E0-cand'; usedDS.add(matched.i);
  }
  // E0 Ped+Frac+Sec
  if(!matched && isReal){
    const kPF=ped+'|||'+frac;
    const cands=(dsByPF.get(kPF)||[]).filter(d=>d.sec===sec && !usedDS.has(d.i));
    if(cands.length){ matched=cands[0]; how='E0-sec'; usedDS.add(matched.i); }
  }
  const cantOk = matched ? Math.abs(matched.cant-cant)<=1 : false;
  const valOk  = matched ? Math.abs(matched.val-val)<=4 : false;
  const ok = matched ? (cantOk&&valOk?'OK':cantOk?'VAL?':'CANT?') : 'SIN_MATCH';
  results.push({i,sec,frac,cant,val,matched:matched?matched.sec:'—',matchedCant:matched?matched.cant:0,ok,how});
}

console.log('=== LAYOUT ROW POR ROW ===');
for(const r of results){
  const flag = r.ok==='OK'?'✓':'✗';
  console.log(flag,'Fila'+r.i,'SEC='+r.sec,'Frac='+r.frac,'Cant='+r.cant,'Val='+r.val,
    '→ DS_sec='+r.matched+'('+r.matchedCant+')','['+r.ok+']',r.how);
}

const unmatched = results.filter(r=>r.ok!=='OK');
console.log('\n--- Filas con problema:',unmatched.length,'de',results.length);

const unusedDS = dsAll.filter(d=>!usedDS.has(d.i));
console.log('DS no usadas:',unusedDS.length);
for(const d of unusedDS) console.log('  DS Sec='+d.sec,'Frac='+d.frac,'Cant='+d.cant,'Val='+d.val,'Pais='+d.pais);
