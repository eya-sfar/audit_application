const ExcelJS=require('exceljs'),fs=require('fs'),path=require('path');
const dim=(b,ext)=>{try{if(ext==='png')return[b.readUInt32BE(16),b.readUInt32BE(20)];let i=2;while(i<b.length){if(b[i]!==0xFF){i++;continue}const m=b[i+1];if(m>=0xC0&&m<=0xCF&&![0xC4,0xC8,0xCC].includes(m))return[b.readUInt16BE(i+7),b.readUInt16BE(i+5)];i+=2+b.readUInt16BE(i+2)}}catch{}return[600,400]};
const MAXW=720,MAXPAGES=12;
async function pdfPages(buf){const mupdf=await import('mupdf');const doc=mupdf.Document.openDocument(buf,'application/pdf'),n=doc.countPages(),out=[];
 for(let i=0;i<Math.min(n,MAXPAGES);i++){const pix=doc.loadPage(i).toPixmap(mupdf.Matrix.scale(1.25,1.25),mupdf.ColorSpace.DeviceRGB,false,true);out.push({buf:Buffer.from(pix.asPNG()),w:pix.getWidth(),h:pix.getHeight()})}return{pages:out,total:n}}
const ST={pending:'À déposer',submitted:'À valider',validated:'Validée',rejected:'Rejetée'};
const AN={arrival:'Arrivée',approval:'Approbation',end:'Fin'};
const D=s=>s?new Date((s.length===10?s+'T00:00:00':s).replace(/Z?$/,'Z')):null;
const fx=(formula,result)=>({formula,result});
const COL={'À déposer':'FFE9ECEF','À valider':'FFFFE8B3','Validée':'FFC8EBD7','Rejetée':'FFF6C6CB'};
module.exports=async function build({program,wave,generated,eps,types,names,uploads}){
 const wb=new ExcelJS.Workbook();wb.creator='Preuves EP';
 const sum=wb.addWorksheet('Résumé'),epS=wb.addWorksheet('EP'),prS=wb.addWorksheet('Preuves'),mx=wb.addWorksheet('Matrice'),lt=wb.addWorksheet('Retards'),rl=wb.addWorksheet('Règles');
 const head=(ws,cols)=>{ws.columns=cols.map(c=>({header:c[0],key:c[1],width:c[2]}));const r=ws.getRow(1);r.height=30;r.eachCell(c=>{c.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF14213D'}};c.alignment={vertical:'middle',wrapText:true}});ws.views=[{state:'frozen',ySplit:1}]};
 const DF='dd/mm/yyyy';
 // --- Preuves (une ligne par preuve et par EP)
 head(prS,[['N° dossier','id',11],['EP','ep',24],['Phase','phase',13],['N°','n',6],['Preuve','proof',44],['Deadline','dl',12],['Statut','st',12],['Déposée le','sub',12],['Déposée par','by',18],['Décision le','rev',12],['Motif de rejet','note',30],['Jours de retard','late',10],['Fichiers','files',50]]);
 const rows=[];let r=2;
 eps.forEach(e=>e.proofs.forEach((p,i)=>{const st=ST[p.status],dl=p.deadline,late=p.late?Math.max(0,Math.round((Date.parse(generated)-Date.parse(dl))/864e5)):0;
  rows.push({ep:e,p,st,late});
  prS.addRow([e.id,e.name,p.phase,i+1,p.name,D(dl),st,D(p.submitted_at),p.submitted_by?names[p.submitted_by]:null,D(p.reviewed_at),p.note,fx(`IF(AND(F${r}<>"",OR(G${r}="À déposer",G${r}="Rejetée")),MAX(0,TODAY()-F${r}),0)`,late),p.files.join(', ')]);r++}));
 ['F','H','J'].forEach(c=>prS.getColumn(c).numFmt=DF);prS.autoFilter={from:'A1',to:'M1'};
 // --- EP
 head(epS,[['N° dossier','id',11],['EP','name',26],['ID EXPA','x',12],['Phase actuelle','ph',14],['Approbation','ap',13],['Arrivée','ar',12],['Durée (sem.)','du',9],['Fin','end',12],['Buddy','b',18],['Preuves attendues','t',11],['Déposées','dep',10],['Non déposées','mis',11],['Validées','val',10],['Rejetées','rej',10],['En retard','lt',10],['Taux validé','tx',10]]);
 const P='Preuves!';
 eps.forEach((e,i)=>{const n=i+2,rs=rows.filter(x=>x.ep.id===e.id),c=s=>rs.filter(x=>x.st===s).length,dep=c('À valider')+c('Validée');
  epS.addRow([e.id,e.name,e.expa_id,e.phase,D(e.approval_date),D(e.arrival_date),e.duration_weeks,D(e.end_date),e.buddy||'Non affecté',
   fx(`COUNTIF(${P}$A:$A,A${n})`,rs.length),fx(`COUNTIFS(${P}$A:$A,A${n},${P}$G:$G,"À valider")+COUNTIFS(${P}$A:$A,A${n},${P}$G:$G,"Validée")`,dep),fx(`J${n}-K${n}`,rs.length-dep),
   fx(`COUNTIFS(${P}$A:$A,A${n},${P}$G:$G,"Validée")`,c('Validée')),fx(`COUNTIFS(${P}$A:$A,A${n},${P}$G:$G,"Rejetée")`,c('Rejetée')),
   fx(`COUNTIFS(${P}$A:$A,A${n},${P}$L:$L,">0")`,rs.filter(x=>x.late>0).length),fx(`IF(J${n}=0,0,M${n}/J${n})`,rs.length?c('Validée')/rs.length:0)])});
 ['E','F','H'].forEach(c=>epS.getColumn(c).numFmt=DF);epS.getColumn('P').numFmt='0%';epS.autoFilter={from:'A1',to:'P1'};
 // --- Résumé
 const tot=rows.length,cn=s=>rows.filter(x=>x.st===s).length,dep=cn('À valider')+cn('Validée'),lateN=rows.filter(x=>x.late>0).length;
 sum.getColumn(1).width=34;[2,3,4,5].forEach(i=>sum.getColumn(i).width=15);
 sum.getCell('A1').value=`Audit ${program}${wave?' · '+wave:''}`;sum.getCell('A1').font={name:'Arial',size:16,bold:true};
 sum.getCell('A2').value=`Export du ${generated.split('-').reverse().join('/')}. Les chiffres se recalculent à partir des onglets EP et Preuves.`;
 const k=[['EP suivis',fx('COUNT(EP!$A:$A)',eps.length)],['Preuves attendues',fx(`COUNT(${P}$A:$A)`,tot)],['Preuves déposées',fx(`COUNTIF(${P}$G:$G,"À valider")+COUNTIF(${P}$G:$G,"Validée")`,dep)],['Non déposées',fx('B5-B6',tot-dep)],['Validées',fx(`COUNTIF(${P}$G:$G,"Validée")`,cn('Validée'))],['À valider',fx(`COUNTIF(${P}$G:$G,"À valider")`,cn('À valider'))],['Rejetées',fx(`COUNTIF(${P}$G:$G,"Rejetée")`,cn('Rejetée'))],['En retard',fx(`COUNTIF(${P}$L:$L,">0")`,lateN)],['Taux de preuves validées',fx('IF(B5=0,0,B8/B5)',tot?cn('Validée')/tot:0)]];
 k.forEach((x,i)=>{sum.getCell(4+i,1).value=x[0];sum.getCell(4+i,2).value=x[1]});sum.getCell('B12').numFmt='0%';
 sum.getCell('A3').value='Indicateurs globaux';sum.getCell('A3').font={name:'Arial',bold:true};
 const th=(row,labels)=>labels.forEach((l,i)=>{const c=sum.getCell(row,i+1);c.value=l;c.font={name:'Arial',size:10,bold:true,color:{argb:'FFFFFFFF'}};c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF14213D'}}});
 th(15,['Par phase de la guideline','Attendues','Déposées','Validées','En retard']);
 ['Approbation','Réalisation','Fin'].forEach((f,i)=>{const n=16+i,rs=rows.filter(x=>x.p.phase===f);sum.getRow(n).values=[f,
  fx(`COUNTIF(${P}$C:$C,A${n})`,rs.length),fx(`COUNTIFS(${P}$C:$C,A${n},${P}$G:$G,"À valider")+COUNTIFS(${P}$C:$C,A${n},${P}$G:$G,"Validée")`,rs.filter(x=>x.st==='À valider'||x.st==='Validée').length),
  fx(`COUNTIFS(${P}$C:$C,A${n},${P}$G:$G,"Validée")`,rs.filter(x=>x.st==='Validée').length),fx(`COUNTIFS(${P}$C:$C,A${n},${P}$L:$L,">0")`,rs.filter(x=>x.late>0).length)]});
 th(20,['EP par phase actuelle','Nombre']);
 ['Approbation','Réalisation','Fin'].forEach((f,i)=>{const n=21+i;sum.getRow(n).values=[f,fx(`COUNTIF(EP!$D:$D,A${n})`,eps.filter(e=>e.phase===f).length)]});
 // --- Matrice EP x preuves
 head(mx,[['EP','ep',24],['Phase actuelle','ph',14],['Buddy','b',16],...types.map(t=>[t.name,'t'+t.id,16])]);
 mx.getRow(1).height=75;
 eps.forEach(e=>{const row=mx.addRow([e.name,e.phase,e.buddy||'Non affecté',...types.map(t=>{const p=e.proofs.find(x=>x.type_id===t.id);return p?ST[p.status]:'—'})]);
  row.eachCell((c,i)=>{if(i>3&&COL[c.value]){c.fill={type:'pattern',pattern:'solid',fgColor:{argb:COL[c.value]}};c.alignment={horizontal:'center',vertical:'middle',wrapText:true}}})});
 // --- Retards
 head(lt,[['EP','ep',24],['Preuve','p',44],['Phase','ph',13],['Deadline','dl',12],['Jours de retard','l',10],['Statut','st',12],['Buddy','b',18]]);
 rows.filter(x=>x.late>0).sort((a,b)=>b.late-a.late).forEach(x=>lt.addRow([x.ep.name,x.p.name,x.p.phase,D(x.p.deadline),x.late,x.st,x.ep.buddy||'Non affecté']));
 lt.getColumn('D').numFmt=DF;
 // --- Règles
 head(rl,[['N°','n',5],['Preuve','name',44],['Phase','ph',13],['Programme','pr',11],['Mode de collecte','m',14],['Formats','f',16],['Critères de conformité','c',50],['Point de départ','a',13],['Jours (− avant, + après)','o',12],['Pages attendues','pg',9]]);
 types.forEach((t,i)=>rl.addRow([i+1,t.name,t.phase,t.program,t.mode,t.formats,t.criteria.join(' ; '),AN[t.anchor],t.offset_days,t.expected_pages]));

 // --- Annexes : les fichiers eux-mêmes, un onglet par EP
 const link={};const used=new Set(['Résumé','EP','Preuves','Matrice','Retards','Règles']);
 for(const e of eps){const withFiles=e.proofs.filter(p=>p.files.length);if(!withFiles.length)continue;
  let nm=`EP${e.id} ${e.name}`.replace(/[\[\]:*?\/\\]/g,'').slice(0,31);while(used.has(nm))nm=nm.slice(0,28)+'_'+used.size;used.add(nm);
  const ws=wb.addWorksheet(nm);ws.getColumn(1).width=110;ws.views=[{showGridLines:false}];
  ws.getCell('A1').value={text:'← Retour à la liste des preuves',hyperlink:"#'Preuves'!A1"};ws.getCell('A1').font={name:'Arial',size:10,color:{argb:'FF0B5FFF'},underline:true};
  ws.getCell('A2').value=`${e.name} · ${e.phase} · buddy : ${e.buddy||'non affecté'}`;ws.getCell('A2').font={name:'Arial',size:14,bold:true};
  let row=4;
  for(const p of withFiles){link[p.id]={nm,row};const hc=ws.getCell(row,1),st=ST[p.status];
   hc.value=`${p.phase} · ${p.name} · ${st}${p.deadline?' · deadline '+p.deadline.split('-').reverse().join('/'):''}${p.submitted_at?' · déposée le '+p.submitted_at.slice(0,10).split('-').reverse().join('/')+(p.submitted_by?' par '+names[p.submitted_by]:''):''}`;
   hc.font={name:'Arial',size:11,bold:true};hc.fill={type:'pattern',pattern:'solid',fgColor:{argb:COL[st]}};ws.getRow(row).height=22;row++;
   if(p.note){ws.getCell(row,1).value='Motif de rejet : '+p.note;ws.getCell(row,1).font={name:'Arial',size:10,color:{argb:'FFC62839'}};row++}
   const put=(buf,ext,w,h)=>{const k=Math.min(1,MAXW/w),dw=Math.round(w*k),dh=Math.round(h*k);ws.addImage(wb.addImage({buffer:buf,extension:ext}),{tl:{col:0,row:row-1},ext:{width:dw,height:dh}});row+=Math.ceil(dh/20)+1};
   const note=t=>{ws.getCell(row,1).value=t;ws.getCell(row,1).font={name:'Arial',size:10,italic:true,color:{argb:'FF5B6478'}};row++};
   for(const f of p.files){const ext=path.extname(f).slice(1).toLowerCase();let buf;try{buf=fs.readFileSync(path.join(uploads,f))}catch{note(`Fichier introuvable : ${f}`);continue}
    note(`Fichier : ${f}`);
    if(ext==='png'||ext==='jpg'||ext==='jpeg'){const[w,h]=dim(buf,ext);put(buf,ext==='png'?'png':'jpeg',w,h)}
    else if(ext==='pdf'){try{const r=await pdfPages(buf);r.pages.forEach(pg=>put(pg.buf,'png',pg.w,pg.h));if(r.total>MAXPAGES)note(`${r.total-MAXPAGES} page(s) non affichée(s), le PDF compte ${r.total} pages.`)}catch{note('PDF illisible : impossible de l’afficher.')}}
    else note(`Le format .${ext} ne peut pas être affiché dans Excel (nom du fichier conservé ci-dessus).`)}
   row++}}
 rows.forEach((x,i)=>{const l=link[x.p.id];if(l){const c=prS.getCell(i+2,13);c.value={text:`Voir le fichier (${x.p.files.length})`,hyperlink:`#'${l.nm}'!A${l.row}`};c.font={name:'Arial',size:10,color:{argb:'FF0B5FFF'},underline:true}}});
 [sum,epS,prS,mx,lt,rl].forEach(ws=>ws.eachRow(row=>row.eachCell(c=>{c.font={...(c.font||{}),name:'Arial',size:(c.font&&c.font.size&&c.font.size!==11)?c.font.size:10}})));
 return wb};
