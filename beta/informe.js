/* =====================================================================================================
   GENERADOR DE INFORMES (Capa 15)
   Convierte las evidencias de UN evento en un informe PowerPoint (.pptx) o Word (.docx) de diagramación
   NEUTRA: sin logos, sin colores de marca y sin nombrar a los proveedores, porque quien lo recibe es el
   cliente final y el administrador le aplica después el patrón de lámina / plantilla que ese cliente pida.
   Corre 100% en el navegador del administrador:
   - los datos ya están en el panel (config del evento + filas de evidencias con la sesión admin);
   - cada foto se baja, se reduce a ~1.200 px y se suelta la original antes de pasar a la siguiente;
   - si el evento es grande, el informe sale POR PARTES (cada archivo con un tope de fotos) para que el
     navegador nunca tenga en memoria más de lo que cabe.
   Se carga perezosamente (sólo cuando alguien abre "Informe"), así no pesa en la app de captura.
   Librerías locales: pptxgen.bundle.js (PptxGenJS 4.0.1) y docx.umd.js (docx 9.7.1), también perezosas.
   ===================================================================================================== */
(function(){
"use strict";
const BASE=(document.currentScript&&document.currentScript.src)?document.currentScript.src.replace(/[^/]*$/,""):"";
/* Paleta deliberadamente neutra: el informe NO lleva marca (ni logo ni azul/oro de Estrella) para que el
   administrador le aplique el patrón del cliente final sin tener que borrar nada. Los únicos colores que
   quedan son los tres de estado (verde/ámbar/rojo): son semáforo de cobertura, no identidad. */
const TINTA="1F2937", LINEA="C9CFD6", GRIS="5A626B", VERDE_OK="1A9E56", ROJO="B3261E", AMBAR="B7791F", F="Arial";
const TOPE_POR_ARCHIVO=350;          // fotos por archivo; por encima se parte (medido en el ensayo: ver LEEME)
const LADO_PX=1200, CALIDAD=0.72;    // ~120-200 KB por foto: nítida a pantalla completa, liviana en el archivo
const CONC=4;                        // descargas simultáneas

/* ---------- utilidades ---------- */
const MESES=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const DIAS=["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
function fDate(iso){ const d=new Date(String(iso||"")+"T12:00:00"); return isNaN(d)?null:d; }
function fLarga(iso){ const d=fDate(iso); return d?`${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`:(iso||""); }
function fDiaLargo(iso){ const d=fDate(iso); return d?`${DIAS[d.getDay()]} ${d.getDate()} de ${MESES[d.getMonth()]}`:(iso||""); }
function fCorta(iso){ const d=fDate(iso); return d?`${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()}`:(iso||""); }
function rango(fe,ff){ if(!fe) return ""; if(!ff||ff===fe) return "el "+fLarga(fe);
  const a=fDate(fe), b=fDate(ff); if(a&&b&&a.getFullYear()===b.getFullYear()&&a.getMonth()===b.getMonth()) return `del ${a.getDate()} al ${b.getDate()} de ${MESES[b.getMonth()]} de ${b.getFullYear()}`;
  return `del ${fLarga(fe)} al ${fLarga(ff)}`; }
function recorta(t,n){ t=String(t==null?"":t).replace(/\s+/g," ").trim(); return t.length>n?t.slice(0,n-1).trimEnd()+"…":t; }
function esc(t){ return String(t==null?"":t).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
function cap(t){ t=String(t||"").trim(); return t?t.charAt(0).toUpperCase()+t.slice(1):t; }
function muestrear(arr,n){ if(!n||arr.length<=n) return arr.slice(); if(n===1) return [arr[0]];
  const out=[]; for(let k=0;k<n;k++) out.push(arr[Math.round(k*(arr.length-1)/(n-1))]); return [...new Set(out)]; }
function cargarScript(src){ return new Promise((ok,ko)=>{ if(document.querySelector(`script[data-inf="${src}"]`)) return ok();
  const s=document.createElement("script"); s.src=BASE+src; s.dataset.inf=src; s.onload=()=>ok(); s.onerror=()=>ko(new Error("No se pudo cargar "+src)); document.head.appendChild(s); }); }
function bytesADataURL(bytes,mime){ let bin=""; const CH=0x8000; for(let i=0;i<bytes.length;i+=CH) bin+=String.fromCharCode.apply(null,bytes.subarray(i,i+CH)); return `data:${mime};base64,`+btoa(bin); }

/* Reduce la foto en el navegador. La original se suelta apenas se dibuja (lección iOS: encoger el lienzo a 1x1). */
async function prepararFoto(blob){
  const bmp=await createImageBitmap(blob,{imageOrientation:"from-image"});
  const s=Math.min(1,LADO_PX/Math.max(bmp.width,bmp.height)); const w=Math.max(1,Math.round(bmp.width*s)), h=Math.max(1,Math.round(bmp.height*s));
  const c=document.createElement("canvas"); c.width=w; c.height=h; c.getContext("2d").drawImage(bmp,0,0,w,h); bmp.close();
  const out=await new Promise(r=>c.toBlob(r,"image/jpeg",CALIDAD)); c.width=1; c.height=1;
  if(!out) throw new Error("toBlob vacío");
  return {w,h,bytes:new Uint8Array(await out.arrayBuffer())};
}

/* ---------- MODELO: requerimientos del config + evidencias → qué va en el informe ---------- */
function construirModelo(ctx, op){
  const ev=ctx.ev, cfg=ev.config||{}, D=ctx.deps;
  const items=D.anotarIids((cfg.it||[]).slice());
  const porIid=new Map(items.map(it=>[it.iid,it]));
  const visibles=new Set(ctx.visibles||[]);
  const filtrado=ctx.visibles && ctx.visibles.length && ctx.visibles.length!==ctx.evidencias.length;
  const fotos=ctx.evidencias.filter((e,j)=>(!filtrado||visibles.has(j)) && (e.tipo||"foto")==="foto" && e.storage_path);
  const iidDeEvi=e=>{ const id=String(e.item_id||""); const pre=ev.uid+"|"; return id.startsWith(pre)?id.slice(pre.length):D.iidDe(e.item,e.categoria,1); };
  const grupos=new Map(); fotos.forEach(e=>{ const k=iidDeEvi(e); if(!grupos.has(k)) grupos.set(k,[]); grupos.get(k).push(e); });
  const orden=(a,b)=>String((a.fecha||"")+(a.hora||"")+(a.creado||"")).localeCompare(String((b.fecha||"")+(b.hora||"")+(b.creado||"")));
  const hoy=D.isoHoy();
  const reqs=[];
  /* v39.1: categoría por el prefijo del nombre ("Festival Clubes - …") para ítems/adicionales sin categoría */
  const nz=x=>String(x||"").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s+/g," ").trim();
  const catsEv=new Map(); items.forEach(it=>{ const c=String(it.cat||it.grupo||"").trim(); if(c && c!=="Adicionales") catsEv.set(nz(c),c); });
  const porPrefijo=req=>{ const m=String(req||"").match(/^\s*([^\-–—:]{3,60}?)\s*[-–—:]\s+\S/); return m?(catsEv.get(nz(m[1]))||""):""; };
  const armar=(it,lista,adicional)=>{
    lista=(lista||[]).slice().sort(orden);
    const porDia=!adicional && !!cfg.pd && D.candidatoPorDia(it) && it.pd!==false && D.fechasDeItem(it,cfg.fe||"").length>0;
    const fechas=porDia?D.fechasDeItem(it,cfg.fe||""):[];
    const bloques=[];
    if(porDia){
      fechas.forEach((f,k)=>{ const del=lista.filter(e=>e.fecha===f); bloques.push({titulo:`Día ${k+1} de ${fechas.length} · ${fDiaLargo(f)}`,fecha:f,total:del.length,fotos:muestrear(del,op.maxPorGrupo)}); });
      const otras=lista.filter(e=>!fechas.includes(e.fecha)); if(otras.length) bloques.push({titulo:"Otras fechas",total:otras.length,fotos:muestrear(otras,op.maxPorGrupo)});
    } else if(lista.length){ bloques.push({titulo:"",total:lista.length,fotos:muestrear(lista,op.maxPorGrupo)}); }
    const diasCub=fechas.filter(f=>lista.some(e=>e.fecha===f));
    const diasFalta=fechas.filter(f=>f<=hoy && !lista.some(e=>e.fecha===f));
    let estado, color;
    if(!lista.length){ estado="Sin evidencia"; color=ROJO; }
    else if(porDia && diasCub.length<fechas.length){ estado=`Parcial · ${diasCub.length}/${fechas.length} días`; color=AMBAR; }
    else { estado=porDia?`Completo · ${fechas.length}/${fechas.length} días`:"Con evidencia"; color=VERDE_OK; }
    /* `prov` se queda en el modelo (lo usa quien depure desde la consola) pero NO se imprime en ningún
       documento: el informe lo lee el cliente final y no tiene por qué conocer a los proveedores. */
    return {it, codigo:it.codigo||"", req:it.req||"", cat:it.cat||it.grupo||porPrefijo(it.req), grupo:(adicional?(it.grupo||""):(it.cat||it.grupo||""))||porPrefijo(it.req), cant:it.cant||"", um:it.um||"", prov:it.prov||"", car:it.car||"",
      adicional, solicitado:it.solicitado||"", justificacion:it.justificacion||"", porDia, fechas, diasCub, diasFalta, total:lista.length, bloques, estado, color};
  };
  items.forEach(it=>{ const lista=grupos.get(it.iid)||[]; if(filtrado && !lista.length) return; reqs.push(armar(it,lista,false)); });
  const extras=[];
  if(op.adicionales){ grupos.forEach((lista,k)=>{ if(porIid.has(k)) return; const e0=lista[0]||{};
    extras.push(armar({codigo:e0.item,req:e0.requerimiento,cat:e0.categoria,grupo:e0.grupo||"",cant:e0.cantidad,prov:e0.proveedor,solicitado:e0.solicitado_por,justificacion:e0.justificacion},lista,true)); }); }
  const conEv=reqs.filter(r=>r.total>0);
  return { ev, cfg, filtrado, reqs, extras, conEv,
    /* v39: cada adicional va al final de SU categoría (grupo), no todos juntos al final del informe */
    visibles: (op.version==="cliente" ? conEv.concat(extras.filter(r=>r.total)) : reqs.concat(extras)).map((r,i)=>({r,i})).sort((x,y)=>ordenInf(x.r,y.r)||x.i-y.i).map(x=>x.r),
    totalFotos: fotos.length };
}
function ordenInf(a,b){ const ca=String(a.grupo||"").trim(), cb=String(b.grupo||"").trim();
  if(!ca!==!cb) return ca?-1:1;
  const c=ca.localeCompare(cb,"es",{sensitivity:"base"}); if(c) return c;
  return (a.adicional?1:0)-(b.adicional?1:0); }   /* empate: se desempata por el orden de la plantilla */
/* Parte el informe por categorías, respetando el tope de fotos por archivo. Un requerimiento nunca se corta. */
function partir(lista){
  const partes=[]; let actual=[], n=0;
  lista.forEach(r=>{ const k=r.bloques.reduce((s,b)=>s+b.fotos.length,0);
    if(actual.length && n+k>TOPE_POR_ARCHIVO){ partes.push(actual); actual=[]; n=0; }
    actual.push(r); n+=k; });
  if(actual.length||!partes.length) partes.push(actual);
  return partes;
}
function fotosDe(lista){ const s=new Set(); lista.forEach(r=>r.bloques.forEach(b=>b.fotos.forEach(e=>s.add(e)))); return [...s]; }
function textoActividad(m, op, lista){
  const cl=m.ev.cl||"", nombre=op.nombreEvento||m.ev.ev||"";
  const servicios=lista.filter(r=>r.total>0).map(r=>`${String(r.req||"").toLowerCase()}${r.cant?` (${r.cant}${r.um?" "+String(r.um).toLowerCase():""})`:""}`);
  const P=[];
  P.push(`Para el evento ${nombre}${op.lugar?` realizado en ${op.lugar}`:""}${m.cfg.fe?` ${rango(m.cfg.fe,m.cfg.ff)}`:""}, se prestaron los servicios de: ${servicios.join("; ")||"(sin servicios con evidencia)"}.`);
  P.push(`Dicha actividad contó con la colaboración y supervisión por parte de los contratistas de ${cl} y se desarrolló de manera exitosa, además de contar con nuestro personal logístico para el desarrollo del evento.`);
  P.push("No se presentan novedades y se cumple con los requerimientos según las indicaciones dadas por el supervisor; se llevó a cabo sin contratiempos y fue recibido a satisfacción.");
  return P;
}
function novedades(lista){
  const N=[];
  lista.forEach(r=>{ if(!r.total) N.push(`${r.codigo} ${r.req}: sin evidencia registrada.`);
    else if(r.porDia && r.diasFalta.length) N.push(`${r.codigo} ${r.req}: sin foto ${r.diasFalta.map(fDiaLargo).join(", ")}.`); });
  return N;
}
function pie(e){ return `${fCorta(e.fecha)}${e.hora?" · "+String(e.hora).slice(0,5):""}${e.fecha_origen==="declarada"?" · fecha declarada":""}`; }

/* ---------- POWERPOINT ---------- */
async function generarPPTX(m, lista, fotos, op, parte){
  await cargarScript("pptxgen.bundle.js");
  const pptx=new window.PptxGenJS(); pptx.layout="LAYOUT_WIDE"; const W=13.333, H=7.5;
  pptx.author=op.operador; pptx.company=op.operador; pptx.title=`Informe · ${m.ev.ev||""}`;
  /* Diseño neutro: sin fondos, sin logo y sin colores de marca. Al pie, un filete gris separa el número de
     lámina; todo el contenido termina antes de y = 6,95 in, así nada se monta si luego se aplica otro patrón. */
  pptx.defineSlideMaster({title:"INF", background:{color:"FFFFFF"},
    objects:[{rect:{x:0.5,y:7.06,w:W-1.0,h:0.014,fill:{color:LINEA},line:{color:LINEA,width:0}}}],
    slideNumber:{x:11.9,y:7.1,w:0.95,h:0.3,fontFace:F,fontSize:10,color:GRIS,align:"right",valign:"middle"}});
  /* El cuerpo del título encoge según su largo: un requerimiento largo ocupaba dos renglones a 24 pt y
     se montaba sobre el subtítulo. Medido a 12,33 in de ancho: ~55 car. a 24 pt, ~72 a 20 pt, ~95 a 17 pt. */
  const cabecera=(s,t,sub)=>{ const tt=recorta(t,90), fs=tt.length>72?17:tt.length>54?20:24;
    s.addText(tt,{x:0.5,y:0.16,w:12.33,h:0.52,fontFace:F,fontSize:fs,bold:true,color:TINTA,valign:"middle",margin:0});
    if(sub) s.addText(recorta(sub,150),{x:0.5,y:0.70,w:12.33,h:0.3,fontFace:F,fontSize:13,color:GRIS,valign:"top",margin:0});
    s.addShape(pptx.ShapeType.line,{x:0.5,y:1.04,w:12.33,h:0,line:{color:LINEA,width:1}}); };
  const nombre=op.nombreEvento||m.ev.ev||"";
  const tituloParte=parte.total>1?` · Parte ${parte.n} de ${parte.total}`:"";

  // 1 · Portada: sin logo ni franjas de color; sólo el bloque de datos sobre un filete gris
  let s=pptx.addSlide({masterName:"INF"});
  s.addShape(pptx.ShapeType.line,{x:0.9,y:2.55,w:3.6,h:0,line:{color:LINEA,width:2.25}});
  s.addText([{text:"INFORME TÉCNICO DE EJECUCIÓN",options:{fontSize:30,bold:true,color:TINTA,breakLine:true,paraSpaceAfter:10}},
             {text:recorta(nombre,60),options:{fontSize:22,bold:true,color:"222222",breakLine:true,paraSpaceAfter:6}},
             {text:recorta((m.ev.cl||"")+(m.cfg.fe?" · "+cap(rango(m.cfg.fe,m.cfg.ff).replace(/^(el|del) /,"")):""),80),options:{fontSize:15,color:GRIS,breakLine:true,paraSpaceAfter:6}},
             {text:`OPERACIÓN LOGÍSTICA · ${op.operador.toUpperCase()}${tituloParte}`,options:{fontSize:11,color:GRIS}}],
    {x:0.9,y:2.85,w:11.5,h:3.2,fontFace:F,valign:"top",margin:0});
  if(op.version==="interno") s.addText("CONTROL INTERNO · no enviar al cliente",{x:0.9,y:6.35,w:8.3,h:0.4,fontFace:F,fontSize:12,bold:true,color:ROJO,margin:0});

  // 2 · Ficha del evento
  if(parte.n===1){
    s=pptx.addSlide({masterName:"INF"}); cabecera(s,"Ficha del evento",m.ev.cl||"");
    const filas=[["Nombre del evento",nombre],["Lugar",op.lugar||"—"],["Fecha",cap(rango(m.cfg.fe,m.cfg.ff).replace(/^el /,""))||"—"],["Operador",op.operador],["Número de contrato",op.contrato||"—"],["Cliente",m.ev.cl||"—"]];
    s.addTable(filas.map((f,k)=>[{text:`${k+1}. ${f[0]}`,options:{bold:true,color:TINTA}},{text:String(f[1])}]),
      {x:1.0,y:1.55,w:11.3,colW:[3.6,7.7],fontFace:F,fontSize:18,color:"222222",rowH:0.62,valign:"middle",border:{type:"solid",pt:0.5,color:LINEA}});
    // 3 · Resumen de cobertura (en tramos de 9 filas)
    const R=lista; const tramo=10;
    for(let i=0;i<Math.max(1,R.length);i+=tramo){
      s=pptx.addSlide({masterName:"INF"});
      const conFoto=R.filter(r=>r.total>0).length;
      cabecera(s,"Resumen de cobertura"+(R.length>tramo?` (${Math.floor(i/tramo)+1}/${Math.ceil(R.length/tramo)})`:""),`${conFoto} de ${R.length} requerimiento(s) con evidencia · ${R.reduce((a,r)=>a+r.total,0)} foto(s) registradas`);
      const head=["Cód.","Requerimiento","Categoría","Cant.","Fotos","Estado"].map(t=>({text:t,options:{bold:true,color:"FFFFFF",fill:{color:TINTA}}}));
      const rows=R.slice(i,i+tramo).map(r=>[r.codigo,recorta(r.req,62),recorta(r.cat,24),`${r.cant}${r.um?" "+String(r.um).toLowerCase():""}`,String(r.total),{text:r.estado,options:{bold:true,color:r.color}}]);
      s.addTable([head].concat(rows),{x:0.45,y:1.35,w:12.45,colW:[0.7,5.3,2.2,1.4,0.85,2.0],fontFace:F,fontSize:11,color:"222222",rowH:0.46,valign:"middle",border:{type:"solid",pt:0.5,color:LINEA},fill:{color:"FFFFFF"}});
    }
    if(op.hojasPegar){
      [["Requerimiento "+(m.ev.cl||"del cliente"),"Pegue aquí la imagen del requerimiento del cliente"],["Cotización operador","Pegue aquí la imagen de la cotización"]].forEach(([t,h])=>{
        s=pptx.addSlide({masterName:"INF"}); cabecera(s,t,"");
        s.addShape(pptx.ShapeType.rect,{x:1.2,y:1.4,w:10.9,h:5.4,fill:{color:"FAFAFB"},line:{color:LINEA,width:1.25,dashType:"dash"}});
        s.addText(h,{x:1.2,y:3.8,w:10.9,h:0.6,align:"center",fontFace:F,fontSize:16,color:GRIS}); });
    }
  }

  // 4 · Evidencias por requerimiento
  const COLS=4, GAP=0.22, MX=0.5, Y0=1.3, ALTO=5.55, PIE=0.32;   /* foto + pie terminan en y ≈ 6,85 in; la barra del pie empieza en 7,14 */
  const cw=(W-2*MX-(COLS-1)*GAP)/COLS;
  const hojaFotos=(r,bloque,grupo,k,nk)=>{
    const sl=pptx.addSlide({masterName:"INF"});
    const meta=[r.adicional?("ADICIONAL EN SITIO"+(r.grupo?" · "+r.grupo:"")):r.cat, r.cant?`Cantidad: ${r.cant}${r.um?" "+String(r.um).toLowerCase():""}`:"", bloque.titulo, nk>1?`(${k}/${nk})`:""].filter(Boolean).join("  ·  ");
    cabecera(sl,`Evidencias · ${r.codigo?r.codigo+" ":""}${r.req}`,meta);
    const nota=r.adicional?[r.solicitado?`Solicitado por: ${r.solicitado}`:"",r.justificacion].filter(Boolean).join(" · "):r.car;
    if(nota) sl.addText(recorta(nota,210),{x:0.5,y:1.10,w:12.33,h:0.32,fontFace:F,fontSize:11,italic:true,color:"6B6B6B",margin:0,valign:"top"});
    const y0=Y0+0.2, boxH=ALTO-0.2-PIE-0.05;
    const off=(W-2*MX-(grupo.length*cw+(grupo.length-1)*GAP))/2;
    grupo.forEach((e,c)=>{ const x=MX+off+c*(cw+GAP); const f=fotos.get(e);
      if(!f||f.falla){ sl.addShape(pptx.ShapeType.rect,{x,y:y0,w:cw,h:boxH,fill:{color:"F2F2F2"},line:{color:"CCCCCC",width:0.75}});
        sl.addText("Foto no disponible",{x,y:y0+boxH/2-0.2,w:cw,h:0.4,align:"center",fontFace:F,fontSize:11,color:"888888"}); }
      else { const esc2=Math.min(cw/f.w,boxH/f.h), w=f.w*esc2, h=f.h*esc2;
        sl.addImage({data:bytesADataURL(f.bytes,"image/jpeg"),x:x+(cw-w)/2,y:y0+boxH-h,w,h}); }
      sl.addText(pie(e),{x,y:y0+boxH+0.04,w:cw,h:PIE,align:"center",fontFace:F,fontSize:10,color:GRIS,fill:{color:"FFFFFF"},margin:0}); });
  };
  const hojaVacia=(r,bloque)=>{ const sl=pptx.addSlide({masterName:"INF"});
    cabecera(sl,`Evidencias · ${r.codigo?r.codigo+" ":""}${r.req}`,[r.cat,bloque?bloque.titulo:""].filter(Boolean).join("  ·  "));
    sl.addShape(pptx.ShapeType.rect,{x:2.5,y:2.6,w:8.3,h:1.6,fill:{color:"FDECEA"},line:{color:ROJO,width:1}});
    sl.addText(bloque?"SIN EVIDENCIA ESTE DÍA":"SIN EVIDENCIA REGISTRADA",{x:2.5,y:2.6,w:8.3,h:1.6,align:"center",valign:"middle",fontFace:F,fontSize:22,bold:true,color:ROJO}); };
  lista.forEach(r=>{
    if(!r.total){ if(op.version==="interno") hojaVacia(r,null); return; }
    r.bloques.forEach(b=>{
      if(!b.fotos.length){ if(op.version==="interno") hojaVacia(r,b); return; }
      const nk=Math.ceil(b.fotos.length/COLS);
      for(let i=0;i<b.fotos.length;i+=COLS) hojaFotos(r,b,b.fotos.slice(i,i+COLS),i/COLS+1,nk);
    });
  });

  // 5 · Informe de actividad + novedades · 6 · Contraportada (sólo en la última parte)
  if(parte.n===parte.total){
    s=pptx.addSlide({masterName:"INF"}); cabecera(s,"Informe de actividad",nombre);
    const P=textoActividad(m,op,m.visibles);
    s.addText(P.map((t,k)=>({text:t,options:{breakLine:k<P.length-1,paraSpaceAfter:10}})),{x:1.0,y:1.45,w:11.3,h:5.3,fontFace:F,fontSize:16,color:"222222",valign:"top",margin:0,fit:"shrink"});
    const N=op.version==="interno"?novedades(m.visibles):[];
    if(N.length){ s=pptx.addSlide({masterName:"INF"}); cabecera(s,"Novedades de cobertura (control interno)",`${N.length} novedad(es)`);
      s.addText(N.slice(0,16).map((t,k)=>({text:t,options:{bullet:true,breakLine:k<Math.min(N.length,16)-1}})),{x:0.9,y:1.4,w:11.6,h:4.95,fontFace:F,fontSize:13,color:"222222",valign:"top",margin:0});
      if(N.length>16) s.addText(`… y ${N.length-16} más (ver el resumen de cobertura).`,{x:0.9,y:6.45,w:11,h:0.35,fontFace:F,fontSize:11,italic:true,color:GRIS,margin:0}); }
    s=pptx.addSlide({masterName:"INF"});
    s.addShape(pptx.ShapeType.line,{x:4.92,y:3.05,w:3.5,h:0,line:{color:LINEA,width:2.25}});
    s.addText([{text:"Gracias",options:{fontSize:34,bold:true,color:TINTA,breakLine:true,paraSpaceAfter:12}},
               {text:op.contacto||op.operador,options:{fontSize:18,bold:true,color:TINTA,breakLine:true,paraSpaceAfter:4}},
               {text:op.cargo||"",options:{fontSize:15,color:GRIS}}],{x:0.8,y:3.35,w:11.7,h:2.0,fontFace:F,align:"center",valign:"top",margin:0});
  }
  return await pptx.write({outputType:"blob"});
}

/* ---------- WORD ---------- */
async function generarDOCX(m, lista, fotos, op, parte){
  await cargarScript("docx.umd.js");
  const X=window.docx;
  const {Document,Packer,Paragraph,TextRun,ImageRun,Table,TableRow,TableCell,WidthType,AlignmentType,HeadingLevel,BorderStyle,Header,Footer,PageNumber,ShadingType,VerticalAlign,PageBreak,TableLayoutType}=X;
  const nombre=op.nombreEvento||m.ev.ev||"";
  const T=(t,o)=>new TextRun(Object.assign({text:String(t==null?"":t),font:F},o||{}));
  const P=(t,o,po)=>new Paragraph(Object.assign({children:[T(t,o)]},po||{}));
  const H1=t=>new Paragraph({heading:HeadingLevel.HEADING_1,spacing:{before:240,after:120},children:[T(t,{bold:true,color:TINTA,size:30})]});
  const H2=t=>new Paragraph({heading:HeadingLevel.HEADING_2,spacing:{before:200,after:60},keepNext:true,children:[T(t,{bold:true,color:TINTA,size:24})]});
  const H3=t=>new Paragraph({spacing:{before:120,after:60},keepNext:true,children:[T(t,{bold:true,color:"333333",size:20})]});
  const sinBorde={top:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}};
  const borde={top:{style:BorderStyle.SINGLE,size:4,color:LINEA},bottom:{style:BorderStyle.SINGLE,size:4,color:LINEA},left:{style:BorderStyle.SINGLE,size:4,color:LINEA},right:{style:BorderStyle.SINGLE,size:4,color:LINEA}};
  const celda=(hijos,o)=>new TableCell(Object.assign({children:Array.isArray(hijos)?hijos:[hijos],margins:{top:60,bottom:60,left:90,right:90},verticalAlign:VerticalAlign.CENTER,borders:borde},o||{}));
  const ANCHO_UTIL=9360; // twips: carta con márgenes de 1,9 cm ≈ 16,5 cm
  const tabla=(filas,anchos)=>new Table({width:{size:ANCHO_UTIL,type:WidthType.DXA},columnWidths:anchos,layout:TableLayoutType.FIXED,rows:filas});
  const img=(d,wpx,hpx)=>new ImageRun({type:"jpg",data:d.bytes,transformation:{width:Math.round(wpx),height:Math.round(hpx)}});
  const hijos=[];
  // Portada (sin logo ni colores de marca: el administrador aplica después la plantilla del cliente final)
  hijos.push(P("INFORME TÉCNICO DE EJECUCIÓN",{bold:true,size:44,color:TINTA},{alignment:AlignmentType.CENTER,spacing:{before:2600,after:200}}));
  hijos.push(P(nombre,{bold:true,size:32,color:"222222"},{alignment:AlignmentType.CENTER,spacing:{after:120}}));
  hijos.push(P(m.ev.cl||"",{size:26,color:GRIS},{alignment:AlignmentType.CENTER,spacing:{after:80}}));
  if(m.cfg.fe) hijos.push(P(cap(rango(m.cfg.fe,m.cfg.ff).replace(/^(el|del) /,"")),{size:24,color:GRIS},{alignment:AlignmentType.CENTER,spacing:{after:600}}));
  hijos.push(P(`Operación logística · ${op.operador}${parte.total>1?` · Parte ${parte.n} de ${parte.total}`:""}`,{size:20,color:GRIS},{alignment:AlignmentType.CENTER}));
  if(op.version==="interno") hijos.push(P("CONTROL INTERNO · no enviar al cliente",{bold:true,size:20,color:ROJO},{alignment:AlignmentType.CENTER,spacing:{before:200}}));
  hijos.push(new Paragraph({children:[new PageBreak()]}));
  let sec=1;
  if(parte.n===1){
    hijos.push(H1(`${sec++}. Ficha del evento`));
    const ficha=[["Nombre del evento",nombre],["Lugar",op.lugar||"—"],["Fecha",cap(rango(m.cfg.fe,m.cfg.ff).replace(/^el /,""))||"—"],["Operador",op.operador],["Número de contrato",op.contrato||"—"],["Cliente",m.ev.cl||"—"]];
    hijos.push(tabla(ficha.map(f=>new TableRow({children:[celda(P(f[0],{bold:true,color:TINTA,size:20}),{shading:{type:ShadingType.CLEAR,fill:"F3F4F6",color:"auto"},width:{size:3000,type:WidthType.DXA}}),celda(P(f[1],{size:20}),{width:{size:6360,type:WidthType.DXA}})]})),[3000,6360]));
    hijos.push(H1(`${sec++}. Resumen de cobertura`));
    hijos.push(P(`${lista.filter(r=>r.total>0).length} de ${lista.length} requerimiento(s) con evidencia · ${lista.reduce((a,r)=>a+r.total,0)} foto(s) registradas.`,{size:20,color:GRIS},{spacing:{after:120}}));
    const anch=[600,3400,1700,1100,660,1900];
    const cab=new TableRow({tableHeader:true,children:["Cód.","Requerimiento","Categoría","Cant.","Fotos","Estado"].map((t,k)=>celda(P(t,{bold:true,color:"FFFFFF",size:17}),{shading:{type:ShadingType.CLEAR,fill:TINTA,color:"auto"},width:{size:anch[k],type:WidthType.DXA}}))});
    hijos.push(tabla([cab].concat(lista.map(r=>new TableRow({cantSplit:true,children:[r.codigo,r.req,r.cat,`${r.cant}${r.um?" "+String(r.um).toLowerCase():""}`,String(r.total),r.estado].map((t,k)=>celda(P(t,k===5?{bold:true,color:r.color,size:16}:{size:17}),{width:{size:anch[k],type:WidthType.DXA}}))}))),anch));
  }
  hijos.push(H1(`${sec++}. Evidencias por requerimiento`));
  const COLS=3, ANCHO_FOTO=196, ALTO_MAX=250; // px a 96 dpi ≈ 5,2 × 6,6 cm
  const rejilla=grupo=>{ const filas=[];
    for(let i=0;i<grupo.length;i+=COLS){ const fila=grupo.slice(i,i+COLS);
      const celdas=fila.map(e=>{ const f=fotos.get(e); const h=[];
        if(!f||f.falla) h.push(P("Foto no disponible",{size:16,color:"888888"},{alignment:AlignmentType.CENTER}));
        else { const k=Math.min(ANCHO_FOTO/f.w,ALTO_MAX/f.h); h.push(new Paragraph({alignment:AlignmentType.CENTER,children:[img(f,f.w*k,f.h*k)]})); }
        h.push(P(pie(e),{size:15,color:GRIS},{alignment:AlignmentType.CENTER}));
        return celda(h,{borders:sinBorde,width:{size:3120,type:WidthType.DXA},verticalAlign:VerticalAlign.BOTTOM}); });
      while(celdas.length<COLS) celdas.push(celda(P(""),{borders:sinBorde,width:{size:3120,type:WidthType.DXA}}));
      filas.push(new TableRow({cantSplit:true,children:celdas})); }
    return tabla(filas,[3120,3120,3120]); };
  lista.forEach(r=>{
    if(!r.total && op.version!=="interno") return;
    hijos.push(H2(`${r.codigo?r.codigo+" · ":""}${r.req}`));
    hijos.push(P([r.adicional?("Adicional en sitio"+(r.grupo?" · "+r.grupo:"")):r.cat, r.cant?`Cantidad: ${r.cant}${r.um?" "+String(r.um).toLowerCase():""}`:"", `Estado: ${r.estado}`].filter(Boolean).join("  ·  "),{size:18,color:GRIS},{keepNext:true}));
    const nota=r.adicional?[r.solicitado?`Solicitado por: ${r.solicitado}`:"",r.justificacion].filter(Boolean).join(" · "):r.car;
    if(nota) hijos.push(P(recorta(nota,400),{size:17,italics:true,color:"6B6B6B"},{keepNext:true,spacing:{after:80}}));
    if(!r.total){ hijos.push(P("SIN EVIDENCIA REGISTRADA",{bold:true,color:ROJO,size:20})); return; }
    r.bloques.forEach(b=>{
      if(b.titulo) hijos.push(H3(b.titulo+(b.total>b.fotos.length?` · ${b.fotos.length} de ${b.total} fotos`:"")));
      else if(b.total>b.fotos.length) hijos.push(P(`${b.fotos.length} de ${b.total} fotos`,{size:16,color:GRIS}));
      if(!b.fotos.length){ if(op.version==="interno") hijos.push(P("SIN EVIDENCIA ESTE DÍA",{bold:true,color:ROJO,size:18})); return; }
      hijos.push(rejilla(b.fotos));
    });
  });
  if(parte.n===parte.total){
    hijos.push(H1(`${sec++}. Informe de actividad`));
    textoActividad(m,op,m.visibles).forEach(t=>hijos.push(P(t,{size:21},{spacing:{after:140},alignment:AlignmentType.JUSTIFIED})));
    const N=op.version==="interno"?novedades(m.visibles):[];
    if(N.length){ hijos.push(H2("Novedades de cobertura (control interno)")); N.forEach(t=>hijos.push(new Paragraph({bullet:{level:0},children:[T(t,{size:19})]}))); }
    hijos.push(H1(`${sec++}. Firmas`));
    const fir=["Elaboró","Revisó","Aprobó"];
    hijos.push(tabla([new TableRow({children:fir.map(t=>celda([P(" ",{size:20},{spacing:{before:700}}),P("______________________________",{size:18},{alignment:AlignmentType.CENTER}),P(t,{bold:true,color:TINTA,size:19},{alignment:AlignmentType.CENTER})],{borders:sinBorde,width:{size:3120,type:WidthType.DXA}}))})],[3120,3120,3120]));
  }
  const doc=new Document({creator:op.operador,title:`Informe · ${nombre}`,styles:{default:{document:{run:{font:F,size:20}}}},
    sections:[{properties:{page:{size:{width:12240,height:15840},margin:{top:1080,bottom:1000,left:1440,right:1440}}},
      headers:{default:new Header({children:[new Paragraph({children:[T("Informe técnico de ejecución · "+recorta(nombre,60),{size:16,color:GRIS})]})]})},
      footers:{default:new Footer({children:[new Paragraph({alignment:AlignmentType.RIGHT,children:[T("Página ",{size:16,color:GRIS}),new TextRun({children:[PageNumber.CURRENT],font:F,size:16,color:GRIS})]})]})},
      children:hijos}]});
  return await Packer.toBlob(doc);
}

/* ---------- ORQUESTADOR ---------- */
async function generar(ctx, op, ui){
  const t0=performance.now();
  const m=construirModelo(ctx,op);
  const lista=m.visibles;
  if(!lista.length) throw new Error("No hay requerimientos para el informe con estas opciones.");
  const partes=partir(lista);
  const resumen={partes:partes.length, fotos:0, fallas:0, bytes:0, archivos:[]};
  for(let p=0;p<partes.length;p++){
    if(ui.cancelado()) break;
    const L=partes[p], F2=fotosDe(L), fotos=new Map(); let hechas=0, fallas=0;
    const cola=F2.slice();
    ui.progreso(`Parte ${p+1} de ${partes.length} · preparando fotos`,0,F2.length,0);
    const trabajador=async()=>{ while(cola.length && !ui.cancelado()){ const e=cola.shift();
      try{ const b=await ctx.bajarBlob(e); if(!b) throw new Error("sin archivo"); fotos.set(e,await prepararFoto(b)); }
      catch(_){ fotos.set(e,{falla:true}); fallas++; }
      hechas++; ui.progreso(`Parte ${p+1} de ${partes.length} · preparando fotos`,hechas,F2.length,fallas); } };
    await Promise.all(Array.from({length:CONC},trabajador));
    if(ui.cancelado()) break;
    ui.progreso(`Parte ${p+1} de ${partes.length} · armando ${op.formato==="docx"?"Word":"PowerPoint"}`,F2.length,F2.length,fallas);
    const parte={n:p+1,total:partes.length};
    const blob=op.formato==="docx"?await generarDOCX(m,L,fotos,op,parte):await generarPPTX(m,L,fotos,op,parte);
    fotos.clear();
    const nom=`INFORME_${ctx.deps.slug(m.ev.cl||"")}_${ctx.deps.slug(op.nombreEvento||m.ev.ev||"")}_${op.version}${partes.length>1?`_parte${p+1}de${partes.length}`:""}.${op.formato}`;
    ctx.deps.bajar(blob,nom);
    resumen.fotos+=F2.length; resumen.fallas+=fallas; resumen.bytes+=blob.size; resumen.archivos.push({nom,bytes:blob.size,fotos:F2.length});
    if(p<partes.length-1) await new Promise(r=>setTimeout(r,800));   // Chrome agrupa descargas seguidas; un respiro evita el bloqueo
  }
  resumen.segundos=Math.round((performance.now()-t0)/1000);
  resumen.cancelado=ui.cancelado();
  return resumen;
}

/* ---------- PANEL (dentro de la vista de evidencias del evento) ---------- */
function lugarFrecuente(evidencias){ const c={}; evidencias.forEach(e=>{ const l=String(e.lugar||"").split(",").slice(0,3).join(",").trim(); if(l) c[l]=(c[l]||0)+1; });
  return Object.entries(c).sort((a,b)=>b[1]-a[1]).map(x=>x[0])[0]||""; }
function abrir(ctx){
  const box=ctx.contenedor; if(!box) return;
  if(!box.classList.contains("hide") && box.dataset.uid===ctx.ev.uid){ box.classList.add("hide"); return; }
  const D=ctx.deps, key=D.NS+"informe_"+ctx.ev.uid; const g=D.loadJSON(key,{})||{};
  const filtrado=ctx.visibles && ctx.visibles.length && ctx.visibles.length!==ctx.evidencias.length;
  const v=(k,d)=>esc(g[k]!=null?g[k]:d);
  box.dataset.uid=ctx.ev.uid;
  box.innerHTML=`<div style="font-weight:800;margin-bottom:6px">🧾 Informe del evento</div>
    <div class="row"><div><label>Formato</label><select id="infFmt"><option value="pptx">PowerPoint (.pptx)</option><option value="docx">Word (.docx)</option></select></div>
      <div><label>Versión</label><select id="infVer"><option value="cliente">Cliente · sólo con evidencia</option><option value="interno">Control interno · todo + faltantes</option></select></div></div>
    <div class="row" style="margin-top:8px"><div><label>Fotos por requerimiento${ctx.ev.config&&ctx.ev.config.pd?" (por día)":""}</label><select id="infMax"><option value="4">Hasta 4</option><option value="8">Hasta 8</option><option value="12">Hasta 12</option><option value="0">Todas</option></select></div>
      <div><label>Número de contrato</label><input id="infContrato" type="text" value="${v("contrato","")}" placeholder="Ej.: OL-001-2026"></div></div>
    <div class="fld" style="margin-top:8px"><label>Nombre del evento</label><input id="infNombre" type="text" value="${v("nombreEvento",ctx.ev.ev||"")}"></div>
    <div class="fld" style="margin-top:8px"><label>Lugar</label><input id="infLugar" type="text" value="${v("lugar",lugarFrecuente(ctx.evidencias)||(ctx.ev.config&&ctx.ev.config.ci)||"")}"></div>
    <div class="row" style="margin-top:8px"><div><label>Operador</label><input id="infOper" type="text" value="${v("operador","Estrella Grupo Empresarial S.A")}"></div>
      <div><label>Contacto (contraportada)</label><input id="infContacto" type="text" value="${v("contacto","")}" placeholder="Nombre · cargo"></div></div>
    <label style="display:flex;gap:8px;align-items:center;margin-top:10px;font-size:13px;font-weight:400;text-transform:none;letter-spacing:0;color:inherit"><input id="infPegar" type="checkbox" style="width:auto;flex:none;margin:0" ${g.hojasPegar===false?"":"checked"}> Láminas para pegar requerimiento del cliente y cotización</label>
    <label style="display:flex;gap:8px;align-items:center;margin-top:6px;font-size:13px;font-weight:400;text-transform:none;letter-spacing:0;color:inherit"><input id="infAdic" type="checkbox" style="width:auto;flex:none;margin:0" ${g.adicionales?"checked":""}> Incluir requerimientos adicionales reportados en sitio</label>
    ${filtrado?`<div class="mut" style="color:var(--oro-d)">Se usará el filtro actual del visor: ${ctx.visibles.length} de ${ctx.evidencias.length} evidencias.</div>`:""}
    <div class="mut" id="infEst" style="margin-top:6px"></div>
    <button class="btn" id="infGo" style="margin-top:8px">Generar informe</button>
    <div id="infProg" class="hide" style="margin-top:8px"></div>`;
  box.classList.remove("hide");
  if(g.formato) box.querySelector("#infFmt").value=g.formato; if(g.version) box.querySelector("#infVer").value=g.version; if(g.maxPorGrupo!=null) box.querySelector("#infMax").value=String(g.maxPorGrupo);
  const leer=()=>{ const q=s=>box.querySelector(s);
    const ct=q("#infContacto").value.split("·");
    return {formato:q("#infFmt").value, version:q("#infVer").value, maxPorGrupo:+q("#infMax").value, contrato:q("#infContrato").value.trim(),
      nombreEvento:q("#infNombre").value.trim(), lugar:q("#infLugar").value.trim(), operador:q("#infOper").value.trim()||"Estrella Grupo Empresarial S.A",
      contacto:(ct[0]||"").trim(), cargo:(ct[1]||"").trim(), contactoTxt:q("#infContacto").value, hojasPegar:q("#infPegar").checked, adicionales:q("#infAdic").checked}; };
  const estimar=()=>{ try{ const op=leer(); const m=construirModelo(ctx,op); const partes=partir(m.visibles); const n=fotosDe(m.visibles).length;
      box.querySelector("#infEst").textContent=`${m.visibles.length} requerimiento(s) · ${n} foto(s) al informe${partes.length>1?` · saldrá en ${partes.length} archivos (tope ${TOPE_POR_ARCHIVO} fotos c/u)`:""} · ~${Math.max(1,Math.round(n*0.25))} MB`;
    }catch(err){ box.querySelector("#infEst").textContent="⚠ "+err.message; } };
  ["#infFmt","#infVer","#infMax","#infAdic"].forEach(s=>box.querySelector(s).addEventListener("change",estimar)); estimar();
  let cancel=false;
  box.querySelector("#infGo").addEventListener("click",async()=>{
    const op=leer(); const btn=box.querySelector("#infGo"), pr=box.querySelector("#infProg");
    D.saveJSON(key,{formato:op.formato,version:op.version,maxPorGrupo:op.maxPorGrupo,contrato:op.contrato,nombreEvento:op.nombreEvento,lugar:op.lugar,operador:op.operador,contacto:op.contactoTxt,hojasPegar:op.hojasPegar,adicionales:op.adicionales});
    btn.disabled=true; cancel=false; pr.classList.remove("hide");
    const ui={cancelado:()=>cancel, progreso:(t,h,n,f)=>{ pr.innerHTML=`<div style="font-weight:700">${esc(t)}</div><div style="height:7px;border-radius:6px;background:#e9e9ee;overflow:hidden;margin:5px 0"><div style="height:100%;width:${n?Math.round(h/n*100):0}%;background:var(--azul)"></div></div><div class="mut" style="margin:0">${h}/${n} foto(s)${f?` · <span style="color:#b3261e">no disponibles ${f}</span>`:""} · <a class="lnk" id="infCancel">cancelar</a></div>`;
      const c=pr.querySelector("#infCancel"); if(c) c.onclick=()=>{ cancel=true; }; }};
    try{
      const r=await generar(ctx,op,ui);
      window.__informeUltimo=r;
      pr.innerHTML=r.cancelado?`<div class="mut">Cancelado.</div>`:`<div style="font-weight:700;color:#1a9e56">✓ Informe listo (${r.segundos}s)</div>`
        +r.archivos.map(a=>`<div class="mut" style="margin:2px 0 0">${esc(a.nom)} · ${a.fotos} foto(s) · ${(a.bytes/1048576).toFixed(1)} MB</div>`).join("")
        +(r.fallas?`<div class="mut" style="color:#b3261e">${r.fallas} foto(s) no se pudieron bajar: quedaron como "Foto no disponible". Vuelve a generar si la red falló.</div>`:"")
        +`<div class="mut">El texto del informe de actividad es un borrador: revísalo antes de enviar.</div>`;
    }catch(err){ window.__informeUltimo={error:String(err&&err.message||err)}; pr.innerHTML=`<div style="color:#b3261e">⚠ ${esc(err&&err.message||err)}</div>`; }
    finally{ btn.disabled=false; }
  });
}
window.InformeEstrella={abrir, generar, construirModelo, partir, TOPE_POR_ARCHIVO};
})();
