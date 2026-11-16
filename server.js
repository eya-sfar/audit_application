const express=require('express'),Database=require('better-sqlite3'),multer=require('multer'),crypto=require('crypto'),path=require('path'),fs=require('fs');
const db=new Database('data.db');db.pragma('foreign_keys=ON');fs.mkdirSync('uploads',{recursive:true});
db.exec(`
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,name TEXT,email TEXT UNIQUE,role TEXT,program TEXT,pw TEXT);
CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,user_id INTEGER);
CREATE TABLE IF NOT EXISTS proof_types(id INTEGER PRIMARY KEY,pos INTEGER,name TEXT,phase TEXT,program TEXT,mode TEXT,formats TEXT,criteria TEXT,anchor TEXT,offset_days INTEGER,expected_pages INTEGER);
CREATE TABLE IF NOT EXISTS eps(id INTEGER PRIMARY KEY,expa_id TEXT,name TEXT,program TEXT,arrival_date TEXT,approval_date TEXT,duration_weeks INTEGER,end_date TEXT,buddy_id INTEGER,created_by INTEGER);
CREATE TABLE IF NOT EXISTS proofs(id INTEGER PRIMARY KEY,ep_id INTEGER REFERENCES eps(id) ON DELETE CASCADE,type_id INTEGER REFERENCES proof_types(id),status TEXT DEFAULT 'pending',files TEXT DEFAULT '[]',auto TEXT DEFAULT '[]',submitted_by INTEGER,submitted_at TEXT,reviewed_at TEXT,note TEXT,UNIQUE(ep_id,type_id));
CREATE TABLE IF NOT EXISTS notifs(id INTEGER PRIMARY KEY,user_id INTEGER,msg TEXT,ep_id INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP,is_read INTEGER DEFAULT 0);`);

const hash=(p,s=crypto.randomBytes(8).toString('hex'))=>s+':'+crypto.scryptSync(p,s,32).toString('hex');
const check=(p,h)=>hash(p,h.split(':')[0])===h;
const addDays=(d,n)=>{if(!d)return null;const x=new Date(d+'T00:00:00Z');x.setUTCDate(x.getUTCDate()+n);return x.toISOString().slice(0,10)};
const today=()=>new Date().toISOString().slice(0,10);
const A=['email_pdf','screenshot','scan','multi','form','photo_doc','tracker','closing'];
// [nom, phase, programme, mode, formats, critères, ancre, décalage(j), pages]
const SEED=[
['Contact EP (10 jours avant)','Approbation','ALL','email_pdf','pdf',['Email téléchargé en PDF','Date visible','Nom de l’EP visible'],'arrival',-10],
['Infos pick-up','Approbation','ALL','email_pdf','pdf',['Email téléchargé en PDF','Date visible'],'arrival',-3],
['Contact du manager','Approbation','ALL','email_pdf','pdf',['Email téléchargé en PDF','Date visible'],'arrival',-7],
['Visa + lettre d’invitation','Approbation','ALL','multi','pdf,png,jpg',['Mail du visa joint','Lettre d’invitation jointe','Date visible'],'approval',14],
['Infos fournies à l’EP (coût de la vie, assurance, expectation setting)','Approbation','ALL','multi','pdf,png,jpg',['Tous les éléments fournis sont joints','Date visible'],'approval',14],
['Invitation IPS','Approbation','ALL','email_pdf','pdf',['Email téléchargé en PDF','Date visible'],'arrival',-5],
['JD form signé par l’EP','Approbation','ALL','form','pdf,png,jpg',['Signature de l’EP présente','Date visible'],'approval',14],
['Contrat OP','Réalisation','ALL','scan','pdf',['8 pages','Tampon LC','Signature VP','Tampon OP'],'arrival',7,8],
['EP goals (Google Form)','Réalisation','ALL','form','pdf,png,jpg',['Nom de l’EP visible','Date visible'],'arrival',7],
['Accommodation visit','Réalisation','ALL','screenshot','png,jpg,pdf',['Date visible','Capture lisible'],'arrival',3],
['Preuve de livraison IPS','Réalisation','ALL','photo_doc','png,jpg,pdf',['Photo de présence','Présentation jointe','Date visible'],'arrival',14],
['LC Standards tracker','Réalisation','ALL','tracker','pdf,png,jpg,xlsx',['Nom de l’EP présent'],'arrival',14],
['Preuve de départ','Fin','ALL','screenshot','png,jpg,pdf',['Date visible','Nom de l’EP visible'],'end',0],
['Standards survey EXPA','Fin','ALL','screenshot','png,jpg,pdf',['Date visible','Nom de l’EP visible'],'end',7],
['Closing space avec les parties prenantes','Fin','IGV','closing','pdf,png,jpg',['Parties prenantes présentes','Date visible'],'end',3]];
if(!db.prepare('SELECT 1 FROM users').get()){
 const u=db.prepare('INSERT INTO users(name,email,role,program,pw) VALUES(?,?,?,?,?)');
 u.run('TL Audit IGV','tl.igv@demo.tn','tl','IGV',hash('demo'));u.run('TL Audit IGT','tl.igt@demo.tn','tl','IGT',hash('demo'));u.run('Buddy Demo','buddy@demo.tn','buddy',null,hash('demo'));
 const t=db.prepare('INSERT INTO proof_types(pos,name,phase,program,mode,formats,criteria,anchor,offset_days,expected_pages) VALUES(?,?,?,?,?,?,?,?,?,?)');
 SEED.forEach((s,i)=>t.run(i+1,s[0],s[1],s[2],s[3],s[4],JSON.stringify(s[5]),s[6],s[7],s[8]||null));
}
const C=new Map(),st=s=>{let x=C.get(s);if(!x){x=db.prepare(s);C.set(s,x)}return x}; // requêtes préparées mises en cache
const q=(s,...p)=>st(s).all(...p),one=(s,...p)=>st(s).get(...p),run=(s,...p)=>st(s).run(...p);
const notify=(uid,msg,ep)=>uid&&run('INSERT INTO notifs(user_id,msg,ep_id) VALUES(?,?,?)',uid,msg,ep);
const notifyTL=(prog,msg,ep)=>q("SELECT id FROM users WHERE role='tl' AND program=?",prog).forEach(u=>notify(u.id,msg,ep));
const endOf=b=>b.program==='IGV'?addDays(b.arrival_date,(+b.duration_weeks||4)*7):(b.end_date||addDays(b.arrival_date,(+b.duration_weeks||8)*7));
const deadline=(ep,t)=>addDays({arrival:ep.arrival_date,approval:ep.approval_date,end:ep.end_date}[t.anchor],t.offset_days);
function proofsOf(ep){return q('SELECT p.*,t.name,t.phase,t.mode,t.formats,t.criteria,t.anchor,t.offset_days,t.expected_pages FROM proofs p JOIN proof_types t ON t.id=p.type_id WHERE p.ep_id=? ORDER BY t.pos',ep.id).map(p=>{
 const dl=deadline(ep,p);return{...p,criteria:JSON.parse(p.criteria),files:JSON.parse(p.files),auto:JSON.parse(p.auto),deadline:dl,late:!!dl&&dl<today()&&['pending','rejected'].includes(p.status)}})}
const PH=['Approbation','Réalisation','Fin'];
const phaseOf=ep=>{const t=today();return t<ep.arrival_date?'Approbation':t<=ep.end_date?'Réalisation':'Fin'};
const summary=ep=>{const ps=proofsOf(ep),open=ps.filter(p=>['pending','rejected'].includes(p.status)&&p.deadline).sort((a,b)=>a.deadline<b.deadline?-1:1);
 const dep=p=>['submitted','validated'].includes(p.status);
 const byPhase=Object.fromEntries(PH.map(f=>{const l=ps.filter(p=>p.phase===f);return[f,{total:l.length,deposited:l.filter(dep).length,missing:l.filter(p=>!dep(p)).length,validated:l.filter(p=>p.status==='validated').length,late:l.filter(p=>p.late).length}]}));
 return{...ep,phase:phaseOf(ep),byPhase,deposited:ps.filter(dep).length,missing:ps.filter(p=>!dep(p)).length,buddy:(one('SELECT name FROM users WHERE id=?',ep.buddy_id)||{}).name,total:ps.length,validated:ps.filter(p=>p.status==='validated').length,submitted:ps.filter(p=>p.status==='submitted').length,rejected:ps.filter(p=>p.status==='rejected').length,late:ps.filter(p=>p.late).length,next:open[0]?{name:open[0].name,deadline:open[0].deadline}:null}};
const mkProofs=ep=>q("SELECT id FROM proof_types WHERE program IN ('ALL',?)",ep.program).forEach(t=>run('INSERT OR IGNORE INTO proofs(ep_id,type_id) VALUES(?,?)',ep.id,t.id));

const app=express();app.use(express.json());app.use(express.static('public'));
const up=multer({storage:multer.memoryStorage(),limits:{fileSize:15e6}});
const wrap=f=>(req,res)=>{try{f(req,res)}catch(e){console.error(e);res.status(500).json({error:e.message})}};
app.post('/api/login',wrap((req,res)=>{const u=one('SELECT * FROM users WHERE email=?',req.body.email);
 if(!u||!check(req.body.password||'',u.pw))return res.status(401).json({error:'Email ou mot de passe incorrect'});
 const token=crypto.randomBytes(24).toString('hex');run('INSERT INTO sessions VALUES(?,?)',token,u.id);res.json({token,user:{id:u.id,name:u.name,role:u.role,program:u.program}})}));
app.use('/api',(req,res,next)=>{const tk=(req.headers.authorization||'').slice(7)||req.query.t;const s=one('SELECT u.* FROM sessions s JOIN users u ON u.id=s.user_id WHERE token=?',tk);
 if(!s)return res.status(401).json({error:'Non connecté'});req.user=s;next()});
const tlOnly=(req,res,next)=>req.user.role==='tl'?next():res.status(403).json({error:'Réservé au Team Leader'});
const can=(u,ep)=>ep&&(u.role==='tl'?ep.program===u.program:ep.buddy_id===u.id);
const loadEP=(req,res)=>{const ep=one('SELECT * FROM eps WHERE id=?',req.params.id);if(!can(req.user,ep)){res.status(404).json({error:'EP introuvable'});return null}return ep};

app.get('/api/eps',wrap((req,res)=>{const l=req.user.role==='tl'?q('SELECT * FROM eps WHERE program=? ORDER BY arrival_date',req.user.program):q('SELECT * FROM eps WHERE buddy_id=? ORDER BY arrival_date',req.user.id);res.json(l.map(summary))}));
app.post('/api/eps',tlOnly,wrap((req,res)=>{const b={...req.body,program:req.user.program};
 if(!b.name||!b.arrival_date||!b.approval_date)return res.status(400).json({error:'Nom, date d’approbation et date d’arrivée obligatoires'});
 b.end_date=endOf(b);const r=run('INSERT INTO eps(expa_id,name,program,arrival_date,approval_date,duration_weeks,end_date,buddy_id,created_by) VALUES(?,?,?,?,?,?,?,?,?)',b.expa_id||null,b.name,b.program,b.arrival_date,b.approval_date,b.duration_weeks||null,b.end_date,b.buddy_id||null,req.user.id);
 const ep=one('SELECT * FROM eps WHERE id=?',r.lastInsertRowid);mkProofs(ep);notifyTL(ep.program,`Nouvel EP : ${ep.name}, dossier créé, ${proofsOf(ep).length} preuves à suivre`,ep.id);notify(ep.buddy_id,`Tu es buddy de ${ep.name}`,ep.id);res.json(summary(ep))}));
app.get('/api/eps/:id',wrap((req,res)=>{const ep=loadEP(req,res);if(ep)res.json({ep:summary(ep),proofs:proofsOf(ep)})}));
app.put('/api/eps/:id',tlOnly,wrap((req,res)=>{const ep=loadEP(req,res);if(!ep)return;const b={...ep,...req.body,program:ep.program};b.end_date=endOf(b);
 run('UPDATE eps SET name=?,expa_id=?,arrival_date=?,approval_date=?,duration_weeks=?,end_date=?,buddy_id=? WHERE id=?',b.name,b.expa_id,b.arrival_date,b.approval_date,b.duration_weeks,b.end_date,b.buddy_id||null,ep.id);
 if(b.buddy_id&&b.buddy_id!==ep.buddy_id)notify(b.buddy_id,`Tu es maintenant buddy de ${b.name}`,ep.id); // l'historique des preuves reste
 res.json(summary(one('SELECT * FROM eps WHERE id=?',ep.id)))})); // les deadlines sont recalculées à chaque lecture
app.get('/api/buddies',tlOnly,wrap((req,res)=>res.json(q("SELECT id,name,email FROM users WHERE role='buddy'"))));
app.post('/api/buddies',tlOnly,wrap((req,res)=>{const{name,email,password}=req.body;if(!name||!email)return res.status(400).json({error:'Nom et email obligatoires'});
 try{const r=run("INSERT INTO users(name,email,role,pw) VALUES(?,?,'buddy',?)",name,email,hash(password||'demo'));res.json({id:r.lastInsertRowid,name,email})}catch{res.status(400).json({error:'Email déjà utilisé'})}}));
app.get('/api/proof-types',tlOnly,wrap((req,res)=>res.json(q('SELECT * FROM proof_types ORDER BY pos').map(t=>({...t,criteria:JSON.parse(t.criteria)})))));
app.put('/api/proof-types/:id',tlOnly,wrap((req,res)=>{const b=req.body;run('UPDATE proof_types SET anchor=?,offset_days=?,formats=?,criteria=?,expected_pages=? WHERE id=?',b.anchor,+b.offset_days,b.formats,JSON.stringify(b.criteria),b.expected_pages||null,req.params.id);res.json({ok:1})}));

app.post('/api/proofs/:id/upload',up.array('files',8),wrap((req,res)=>{
 const p=one('SELECT p.*,t.name tname,t.phase,t.formats,t.criteria,t.expected_pages FROM proofs p JOIN proof_types t ON t.id=p.type_id WHERE p.id=?',req.params.id);
 const ep=p&&one('SELECT * FROM eps WHERE id=?',p.ep_id);if(!can(req.user,ep))return res.status(404).json({error:'Preuve introuvable'});
 if(p.status==='validated')return res.status(400).json({error:'Preuve déjà validée'});
 const crit=JSON.parse(p.criteria),ticked=JSON.parse(req.body.checks||'[]');
 if(!req.files.length)return res.status(400).json({error:'Aucun fichier'});
 if(ticked.length!==crit.length||!ticked.every(Boolean))return res.status(400).json({error:'Coche tous les critères avant l’envoi'});
 const ok=p.formats.split(','),auto=[];const saved=[];
 for(const [i,f] of req.files.entries()){const ext=path.extname(f.originalname).slice(1).toLowerCase();
  if(!ok.includes(ext))return res.status(400).json({error:`Format .${ext} refusé. Formats acceptés : ${ok.join(', ')}`});
  if(!f.size)return res.status(400).json({error:'Fichier vide'});
  if(ext==='pdf'&&p.expected_pages){const n=(f.buffer.toString('latin1').match(/\/Type\s*\/Page[^s]/g)||[]).length;
   if(n&&n!==p.expected_pages)return res.status(400).json({error:`Le PDF doit faire ${p.expected_pages} pages (${n} détectées)`});if(!n)auto.push('Nombre de pages non vérifiable')}
  const name=`${ep.id}_${p.phase}_${p.tname.replace(/[^\w]+/g,'-').slice(0,30)}_${today()}_${p.id}_${i}.${ext}`;
  fs.writeFileSync(path.join('uploads',name),f.buffer);saved.push(name)}
 run("UPDATE proofs SET status='submitted',files=?,auto=?,submitted_by=?,submitted_at=?,note=NULL WHERE id=?",JSON.stringify(saved),JSON.stringify(auto),req.user.id,new Date().toISOString(),p.id);
 const left=deadline(ep,p);notifyTL(ep.program,`${ep.name} : « ${p.tname} » déposée${left?' (deadline '+left+')':''}`,ep.id);res.json({ok:1,auto})}));
app.post('/api/proofs/:id/review',tlOnly,wrap((req,res)=>{const p=one('SELECT p.*,t.name tname FROM proofs p JOIN proof_types t ON t.id=p.type_id WHERE p.id=?',req.params.id);
 const ep=p&&one('SELECT * FROM eps WHERE id=?',p.ep_id);if(!can(req.user,ep))return res.status(404).json({error:'Preuve introuvable'});
 if(p.status!=='submitted')return res.status(400).json({error:'Aucun dépôt à valider'});
 const{decision,note}=req.body;if(!['validated','rejected'].includes(decision))return res.status(400).json({error:'Décision invalide'});
 if(decision==='rejected'&&!note)return res.status(400).json({error:'Un motif de rejet est obligatoire'});
 run('UPDATE proofs SET status=?,note=?,reviewed_at=? WHERE id=?',decision,note||null,new Date().toISOString(),p.id);
 notify(p.submitted_by||ep.buddy_id,`${ep.name} : « ${p.tname} » ${decision==='validated'?'validée':'rejetée — '+note}`,ep.id);res.json({ok:1})}));
app.get('/api/proofs/:id/file/:i',wrap((req,res)=>{const p=one('SELECT * FROM proofs WHERE id=?',req.params.id);const ep=p&&one('SELECT * FROM eps WHERE id=?',p.ep_id);
 const f=can(req.user,ep)&&JSON.parse(p.files)[req.params.i];if(!f)return res.sendStatus(404);res.sendFile(path.resolve('uploads',f))}));
app.get('/api/notifications',wrap((req,res)=>res.json(q('SELECT * FROM notifs WHERE user_id=? ORDER BY id DESC LIMIT 40',req.user.id))));
app.post('/api/notifications/read',wrap((req,res)=>{run('UPDATE notifs SET is_read=1 WHERE user_id=?',req.user.id);res.json({ok:1})}));
app.get('/api/digest',wrap((req,res)=>{const eps=(req.user.role==='tl'?q('SELECT * FROM eps WHERE program=?',req.user.program):q('SELECT * FROM eps WHERE buddy_id=?',req.user.id));
 const tm=addDays(today(),1);let toValidate=0,tomorrow=0,late=0;eps.forEach(ep=>proofsOf(ep).forEach(p=>{if(p.status==='submitted')toValidate++;if(p.deadline===tm&&p.status!=='validated')tomorrow++;if(p.late)late++}));res.json({toValidate,tomorrow,late})}));
app.get('/api/dashboard',tlOnly,wrap((req,res)=>{const eps=q('SELECT * FROM eps WHERE program=?',req.user.program).map(summary);
 const sum=k=>eps.reduce((a,e)=>a+e[k],0),t=today(),w=addDays(t,7);
 const phases=Object.fromEntries(PH.map(f=>[f,eps.filter(e=>e.phase===f).length]));
 const upcoming=[],late=[];q('SELECT * FROM eps WHERE program=?',req.user.program).forEach(ep=>proofsOf(ep).forEach(p=>{const r={ep_id:ep.id,ep:ep.name,proof:p.name,deadline:p.deadline,buddy:(one('SELECT name FROM users WHERE id=?',ep.buddy_id)||{}).name};
  if(p.late)late.push(r);else if(p.deadline&&p.deadline<=w&&p.deadline>=t&&['pending','rejected'].includes(p.status))upcoming.push(r)}));
 const by=a=>a.sort((x,y)=>x.deadline<y.deadline?-1:1);
 const buddies={};eps.forEach(e=>{const k=e.buddy||'Non affecté';const b=buddies[k]=buddies[k]||{eps:0,missing:0,late:0};b.eps++;b.missing+=e.missing;b.late+=e.late});
 res.json({eps:eps.length,phases,total:sum('total'),deposited:sum('deposited'),missing:sum('missing'),validated:sum('validated'),toValidate:sum('submitted'),rejected:sum('rejected'),late:sum('late'),lateList:by(late),upcoming:by(upcoming),buddies,list:eps})}));
app.get('/api/export.xlsx',tlOnly,async(req,res)=>{try{
 const eps=q('SELECT * FROM eps WHERE program=? ORDER BY arrival_date',req.user.program).map(ep=>({...summary(ep),proofs:proofsOf(ep)}));
 const names=Object.fromEntries(q('SELECT id,name FROM users').map(u=>[u.id,u.name]));
 const types=q("SELECT * FROM proof_types WHERE program IN ('ALL',?) ORDER BY pos",req.user.program).map(t=>({...t,criteria:JSON.parse(t.criteria)}));
 const wb=await require('./export')({program:req.user.program,wave:String(req.query.wave||'').slice(0,60),generated:today(),eps,types,names,uploads:path.resolve('uploads')});
 res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
 res.setHeader('Content-Disposition',`attachment; filename="audit_${req.user.program}_${today()}.xlsx"`);
 await wb.xlsx.write(res);res.end()}catch(e){console.error(e);res.status(500).json({error:e.message})}});
app.listen(process.env.PORT||3000,()=>console.log('http://localhost:'+(process.env.PORT||3000)));
