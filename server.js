const express=require('express');
const path=require('path');
const fs=require('fs');
const os=require('os');
const crypto=require('crypto');
const multer=require('multer');
const Database=require('better-sqlite3');
const PDFDocument=require('pdfkit');
const ExcelJS=require('exceljs');
const {ClientSecretCredential}=require('@azure/identity');

const app=express();
const ROOT=process.cwd();
// Vercel Functions can write only to temporary storage. Never write under /var/task.
const RUNTIME_ROOT=path.join(os.tmpdir(),'mahy-khoory-pms');
const DATA=path.join(RUNTIME_ROOT,'data');
const UPLOADS=path.join(RUNTIME_ROOT,'uploads');
fs.mkdirSync(DATA,{recursive:true});
fs.mkdirSync(UPLOADS,{recursive:true});

const PORT=process.env.PORT||3000;
const BASE=(process.env.BASE_URL||'').replace(/\/$/,'')||`http://localhost:${PORT}`;
const db=new Database(path.join(DATA,'pms.db'));
db.pragma('journal_mode=WAL');
db.exec('CREATE TABLE IF NOT EXISTS evaluations(id TEXT PRIMARY KEY,createdAt TEXT,status TEXT,currentStep INTEGER,data TEXT,workflow TEXT,attachments TEXT,history TEXT)');
const upload=multer({dest:UPLOADS,limits:{fileSize:15*1024*1024,files:10}});
app.use(express.json({limit:'2mb'}));
app.use(express.static(path.join(ROOT,'public')));

function makeToken(){return crypto.randomBytes(24).toString('hex')}
function getRecord(id){const r=db.prepare('SELECT * FROM evaluations WHERE id=?').get(id);if(!r)return null;for(const k of ['data','workflow','attachments','history'])r[k]=JSON.parse(r[k]);return r}
function saveRecord(r){db.prepare('INSERT OR REPLACE INTO evaluations VALUES(?,?,?,?,?,?,?,?)').run(r.id,r.createdAt,r.status,r.currentStep,JSON.stringify(r.data),JSON.stringify(r.workflow),JSON.stringify(r.attachments),JSON.stringify(r.history))}
async function sendMail(to,subject,html,attachments=[]){
  const required=['AZURE_TENANT_ID','AZURE_CLIENT_ID','AZURE_CLIENT_SECRET','MAIL_SENDER'];
  for(const k of required)if(!process.env[k])throw new Error(`Missing Vercel environment variable: ${k}`);
  const cred=new ClientSecretCredential(process.env.AZURE_TENANT_ID,process.env.AZURE_CLIENT_ID,process.env.AZURE_CLIENT_SECRET);
  const access=await cred.getToken('https://graph.microsoft.com/.default');
  const message={subject,body:{contentType:'HTML',content:html},toRecipients:[{emailAddress:{address:to}}],attachments:attachments.map(a=>({'@odata.type':'#microsoft.graph.fileAttachment',name:a.name,contentType:a.type||'application/octet-stream',contentBytes:fs.readFileSync(a.path).toString('base64')}))};
  const response=await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.MAIL_SENDER)}/sendMail`,{method:'POST',headers:{Authorization:`Bearer ${access.token}`,'Content-Type':'application/json'},body:JSON.stringify({message,saveToSentItems:true})});
  if(!response.ok)throw new Error(`Microsoft Graph Mail.Send failed: ${response.status} ${await response.text()}`);
}
function approvalMail(r,step){const d=r.data;const link=`${BASE}/approve.html?id=${encodeURIComponent(r.id)}&token=${encodeURIComponent(step.token)}`;return `<h2>PMS approval required</h2><p><b>Employee:</b> ${d.employeeName}</p><p><b>Period:</b> ${d.from} to ${d.to}</p><p><b>Rating:</b> ${d.overallRating} (${d.percentage}%)</p><p><b>Approval step:</b> ${step.role}</p><p><a href="${link}" style="padding:12px 18px;background:#123b66;color:white;text-decoration:none;border-radius:6px">Open approval</a></p>`}
async function notifyCurrent(r){const s=r.workflow[r.currentStep];await sendMail(s.email,`PMS approval: ${r.data.employeeName}`,approvalMail(r,s),r.attachments)}
function pdfPath(r){return path.join(DATA,`${r.id}.pdf`)}
function buildPdf(r){return new Promise((resolve,reject)=>{const out=pdfPath(r);const doc=new PDFDocument({margin:35,size:'A4'});const stream=fs.createWriteStream(out);doc.pipe(stream);doc.fontSize(16).fillColor('#123b66').text('M.A.H.Y KHOORY GROUP OF COMPANIES',{align:'center'}).fontSize(13).text('PERFORMANCE EVALUATION SHEET - SKILLED / UNSKILLED WORKERS',{align:'center'}).moveDown();doc.fillColor('black').fontSize(9);for(const [k,v] of Object.entries(r.data)){if(k!=='ratings')doc.text(`${k.replace(/([A-Z])/g,' $1')}: ${v??''}`)}doc.moveDown().fontSize(11).fillColor('#123b66').text('Performance Attributes').fillColor('black').fontSize(8);(r.data.ratings||[]).forEach((x,i)=>doc.text(`${i+1}. ${x.attribute} | Rating: ${x.rating} | Comments: ${x.comment||''}`));doc.moveDown().fontSize(11).fillColor('#123b66').text('Approval History').fillColor('black').fontSize(8);r.history.forEach(h=>doc.text(`${h.at} | ${h.action} | ${h.by}${h.comment?' | '+h.comment:''}`));doc.end();stream.on('finish',()=>resolve(out));stream.on('error',reject)})}

app.post('/api/evaluations',upload.array('attachments',10),async(req,res)=>{try{const data=JSON.parse(req.body.data||'{}');const workflow=JSON.parse(req.body.workflow||'[]');if(!data.employeeName||!workflow.length)return res.status(400).json({error:'Employee and workflow are required'});for(const step of workflow){if(!/^\S+@\S+\.\S+$/.test(step.email||''))return res.status(400).json({error:'Every approver needs a valid email'});step.token=makeToken();step.status='Pending'}const r={id:crypto.randomUUID(),createdAt:new Date().toISOString(),status:'Pending',currentStep:0,data,workflow,attachments:(req.files||[]).map(f=>({name:f.originalname,type:f.mimetype,path:f.path,size:f.size})),history:[{at:new Date().toISOString(),action:'Initiated',by:data.initiatorEmail||'Initiator'}]};saveRecord(r);await notifyCurrent(r);res.json({ok:true,id:r.id,message:`Sent to ${workflow[0].email}`})}catch(e){console.error(e);res.status(500).json({error:e.message})}});
app.get('/api/evaluations',(_req,res)=>{const rows=db.prepare('SELECT id FROM evaluations ORDER BY createdAt DESC').all().map(x=>getRecord(x.id)).map(r=>({id:r.id,createdAt:r.createdAt,status:r.status,currentStep:r.currentStep,data:r.data,workflow:r.workflow.map(({token,...s})=>s),attachments:r.attachments.map(({path,...a})=>a),history:r.history}));res.json(rows)});
app.get('/api/approval/:id',(req,res)=>{const r=getRecord(req.params.id);const s=r&&r.workflow[r.currentStep];if(!r||!s||s.token!==req.query.token)return res.status(403).json({error:'Invalid or expired approval link'});res.json({id:r.id,status:r.status,data:r.data,step:{role:s.role,email:s.email},history:r.history})});
app.post('/api/approval/:id',async(req,res)=>{try{const r=getRecord(req.params.id);const s=r&&r.workflow[r.currentStep];if(!r||!s||s.token!==req.body.token)return res.status(403).json({error:'Invalid or expired approval link'});if(!['Approve','Reject'].includes(req.body.action))return res.status(400).json({error:'Invalid action'});s.status=req.body.action==='Approve'?'Approved':'Rejected';s.comment=req.body.comment||'';s.actedAt=new Date().toISOString();r.history.push({at:s.actedAt,action:s.status,by:s.email,comment:s.comment});if(req.body.action==='Reject')r.status='Rejected';else if(r.currentStep<r.workflow.length-1){r.currentStep++;await notifyCurrent(r)}else{r.status='Approved';const pdf=await buildPdf(r);await sendMail(r.data.initiatorEmail,`PMS workflow completed: ${r.data.employeeName}`,'<p>All approval stages are complete. The filled PMS PDF is attached.</p>',[{name:`PMS-${r.data.employeeCode||r.id}.pdf`,type:'application/pdf',path:pdf}])}saveRecord(r);res.json({ok:true,status:r.status})}catch(e){console.error(e);res.status(500).json({error:e.message})}});
app.get('/api/evaluations/:id/pdf',async(req,res)=>{const r=getRecord(req.params.id);if(!r||r.status!=='Approved')return res.status(400).send('PDF is available after all approvals');const p=fs.existsSync(pdfPath(r))?pdfPath(r):await buildPdf(r);res.download(p,`PMS-${r.data.employeeCode||r.id}.pdf`)});
app.get('/api/export.xlsx',async(_req,res)=>{const wb=new ExcelJS.Workbook();const ws=wb.addWorksheet('PMS Records');const rows=db.prepare('SELECT id FROM evaluations ORDER BY createdAt DESC').all().map(x=>getRecord(x.id));const keys=[...new Set(rows.flatMap(r=>Object.keys(r.data).filter(k=>k!=='ratings')))];ws.columns=[{header:'Employee Name',key:'employeeName',width:28},...keys.filter(k=>k!=='employeeName').map(k=>({header:k.replace(/([A-Z])/g,' $1'),key:k,width:22})),{header:'Ratings and Comments',key:'ratingsFlat',width:70},{header:'Workflow',key:'workflowFlat',width:65},{header:'Status',key:'status',width:14},{header:'Approval History',key:'historyFlat',width:70}];for(const r of rows)ws.addRow({...r.data,ratingsFlat:(r.data.ratings||[]).map(x=>`${x.attribute}: ${x.rating} (${x.comment||''})`).join(' | '),workflowFlat:r.workflow.map((x,i)=>`${i+1}. ${x.role} <${x.email}> [${x.status}]`).join(' | '),status:r.status,historyFlat:r.history.map(h=>`${h.at} ${h.action} by ${h.by}${h.comment?' ('+h.comment+')':''}`).join(' | ')});ws.getRow(1).font={bold:true};res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');res.setHeader('Content-Disposition','attachment; filename=PMS_All_Fields.xlsx');await wb.xlsx.write(res);res.end()});
app.get('/api/health',(_req,res)=>res.json({ok:true,runtimeDataPath:DATA,environment:process.env.VERCEL?'vercel':'local'}));

// Local Node server; Vercel imports the exported Express app instead.
if(require.main===module)app.listen(PORT,()=>console.log(`PMS running at ${BASE}`));
module.exports=app;
