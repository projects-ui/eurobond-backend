const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const jwt = require('jsonwebtoken');
const XLSX = require('xlsx');
const cloudinary = require('cloudinary').v2;
const { CloudinaryStorage } = require('multer-storage-cloudinary');
const { MongoClient } = require('mongodb');

dotenv.config();

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

const app = express();
app.use(cors());
app.use(express.json());

const UPLOADS_DIR = path.join(__dirname, 'uploads');
app.use('/uploads', express.static(UPLOADS_DIR));
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });
if (!fs.existsSync(path.join(__dirname, 'data'))) fs.mkdirSync(path.join(__dirname, 'data'), { recursive: true });

// Cache
const stockCache = {};
let db = null;

async function getDB() {
  if (db) return db;
  if (!process.env.MONGODB_URI) return null;
  try {
    const client = new MongoClient(process.env.MONGODB_URI);
    await client.connect();
    db = client.db('eurobond');
    console.log('MongoDB connected!');
    return db;
  } catch(e) { console.log('MongoDB error:', e.message); return null; }
}

function readJSON(f, fallback=[]) { try { return JSON.parse(fs.readFileSync(f,'utf8')); } catch { return fallback; } }
function writeJSON(f, d) { fs.writeFileSync(f, JSON.stringify(d,null,2)); }

const USERS_FILE = path.join(__dirname,'data','users.json');
const LOGS_FILE = path.join(__dirname,'data','logs.json');
const RESOURCES_FILE = path.join(__dirname,'data','resources.json');

async function getUsers() { const d=await getDB(); if(d) return d.collection('users').find().toArray(); return readJSON(USERS_FILE); }
async function saveUser(u) { const d=await getDB(); if(d) return d.collection('users').insertOne(u); const us=readJSON(USERS_FILE); us.push(u); writeJSON(USERS_FILE,us); }
async function deleteUser(id) { const d=await getDB(); if(d){try{const {ObjectId}=require('mongodb');return d.collection('users').deleteOne({_id:new ObjectId(id)});}catch(e){return d.collection('users').deleteOne({id});}} writeJSON(USERS_FILE,readJSON(USERS_FILE).filter(u=>u.id!==id)); }
async function getResources() { const d=await getDB(); if(d){const files=await d.collection('files').find().toArray();const links=await d.collection('links').find().toArray();return{files,links};} return readJSON(RESOURCES_FILE,{files:[],links:[]}); }
async function addResource(type,data) { const d=await getDB(); if(d) return d.collection(type==='file'?'files':'links').insertOne(data); const r=readJSON(RESOURCES_FILE,{files:[],links:[]}); if(type==='file') r.files.push(data); else r.links.push(data); writeJSON(RESOURCES_FILE,r); }
async function deleteResource(type,id) { const d=await getDB(); if(d){const col=d.collection(type==='file'?'files':'links');try{const {ObjectId}=require('mongodb');return col.deleteOne({_id:new ObjectId(id)});}catch(e){return col.deleteOne({id});}} const r=readJSON(RESOURCES_FILE,{files:[],links:[]}); if(type==='file') r.files=r.files.filter(f=>f.id!==id); else r.links=r.links.filter(l=>l.id!==id); writeJSON(RESOURCES_FILE,r); }
async function addLog(log) { const d=await getDB(); if(d) return d.collection('logs').insertOne(log); const logs=readJSON(LOGS_FILE); logs.push(log); writeJSON(LOGS_FILE,logs); }
async function getLogs() { const d=await getDB(); if(d) return d.collection('logs').find().sort({loginTime:-1}).limit(100).toArray(); return readJSON(LOGS_FILE).reverse(); }

setTimeout(async()=>{ const adminEmail=process.env.ADMIN_EMAIL; if(!adminEmail) return; const users=await getUsers(); if(!users.find(u=>u.email===adminEmail)){ await saveUser({id:'1',name:process.env.ADMIN_NAME||'Admin',phone:'9999999999',email:adminEmail,password:process.env.ADMIN_PASSWORD||'eurobond123',role:'Sales Manager',region:'Hyderabad',active:true,createdAt:new Date().toISOString(),lastLogin:null}); console.log('Default user created: '+adminEmail); } },2000);

const cloudinaryStorage = new CloudinaryStorage({ cloudinary, params: async(req,file)=>({ folder:'eurobond', resource_type:'raw', access_mode:'public', type:'upload', public_id:Date.now()+'_'+file.originalname.replace(/[^a-zA-Z0-9._-]/g,'_'), use_filename:false, unique_filename:false, upload_preset:'eurobond_public' }) });
const upload = multer({ storage: cloudinaryStorage });

function authMiddleware(req,res,next){ const token=req.headers.authorization?.split(' ')[1]; if(!token) return res.status(401).json({success:false,message:'No token'}); try{req.user=jwt.verify(token,process.env.JWT_SECRET);next();}catch{res.status(401).json({success:false,message:'Invalid token'});} }
function adminMiddleware(req,res,next){ const token=req.headers.authorization?.split(' ')[1]; if(!token) return res.status(401).json({success:false,message:'No token'}); try{const d=jwt.verify(token,process.env.JWT_SECRET);if(d.role!=='admin') return res.status(403).json({success:false,message:'Admin only'});req.user=d;next();}catch{res.status(401).json({success:false,message:'Invalid token'});} }

// Normalize item code: "er 147 4ti" -> "ER-147-4TI"
function normalizeCode(s) { return s.toUpperCase().trim().replace(/\s+/g,'-'); }
function extractCodes(msg) { const p=/[A-Za-z]{2,3}[\s-]\d{3,4}[\s-][A-Za-z0-9]+/g; return (msg.match(p)||[]).map(normalizeCode); }

const SITES = ['AHM','BBS','BHW','DDN','DEL','GWH','HBL','IND','JAI','KLK','LKO','NGP','PAT','PNB','PUNE','RAI','RNC','UMG'];

function processStock(rows, userMsg) {
  const codes = extractCodes(userMsg);
  const sites = SITES.filter(s => userMsg.toUpperCase().includes(s));
  const wantsBatch = /batch|lot|batchnum|batch.wise/i.test(userMsg);

  let filtered = [];
  if (codes.length > 0) {
    filtered = rows.filter(row => {
      const item = normalizeCode(String(row['Item_No']||''));
      const site = String(row['Site']||'').toUpperCase().trim();
      const itemMatch = codes.some(c => item === c);
      const siteMatch = sites.length===0 || sites.includes(site);
      return itemMatch && siteMatch;
    });
  }
  if (filtered.length===0 && codes.length===0) {
    const words = userMsg.toUpperCase().split(/[\s,]+/).filter(w=>w.length>3);
    filtered = rows.filter(row => {
      const rowStr = Object.values(row).join(' ').toUpperCase();
      return words.some(w=>rowStr.includes(w));
    });
  }
  if (filtered.length===0) return null;

  // Batch mode - group same batch+size
  if (wantsBatch) {
    const batchSummary = {};
    filtered.forEach(row => {
      const key = `${row['Item_No']}|${row['Site']}|${row['Size']}|${row['BatchNum']}`;
      if (!batchSummary[key]) batchSummary[key] = { item:row['Item_No'], site:row['Site'], size:row['Size'], batch:row['BatchNum'], nos:0 };
      batchSummary[key].nos += Number(row['NOS']||0);
    });
    let result = 'BATCHES:\n';
    Object.values(batchSummary).forEach(b => {
      result += `${b.item}|${b.site}|${b.size}|${b.nos}NOS|Batch:${b.batch}\n`;
    });
    result += `Total: ${Object.values(batchSummary).reduce((sum,b)=>sum+b.nos,0)} NOS\n`;
    return result;
  }

  // Summary mode
  const summary = {};
  filtered.forEach(row => {
    const item = normalizeCode(String(row['Item_No']||''));
    const site = String(row['Site']||'').trim();
    const size = String(row['Size']||'').trim().replace(/\s+/g,' ');
    const nos = Number(row['NOS']||0);
    const sqm = Number(row['SQM']||0);
    if (!summary[item]) summary[item]={total_nos:0,total_sqm:0,sites:{},sizes:{},site_sizes:{}};
    summary[item].total_nos+=nos;
    summary[item].total_sqm+=sqm;
    summary[item].sites[site]=(summary[item].sites[site]||0)+nos;
    summary[item].sizes[size]=(summary[item].sizes[size]||0)+nos;
    if (!summary[item].site_sizes[site]) summary[item].site_sizes[site]={};
    summary[item].site_sizes[site][size]=(summary[item].site_sizes[site][size]||0)+nos;
  });

  let result = 'STOCK:\n';
  Object.entries(summary).forEach(([item,data])=>{
    result += `${item}|${data.total_nos}NOS|${data.total_sqm.toFixed(0)}SQM\n`;
    result += `Sizes:${Object.entries(data.sizes).map(([s,n])=>`${s}=${n}`).join(',')}\n`;
    result += `Loc:${Object.entries(data.sites).map(([s,n])=>`${s}=${n}`).join(',')}\n`;
    result += `LocSizes:${Object.entries(data.site_sizes).map(([s,sz])=>`${s}(${Object.entries(sz).map(([k,v])=>`${k}=${v}`).join(',')})`).join('|')}\n`;
  });
  return result;
}

async function getStockRows(fileUrl) {
  if (stockCache[fileUrl]) return stockCache[fileUrl];
  const https=require('https'),http=require('http');
  const tmp=path.join(UPLOADS_DIR,'tmp_'+Date.now()+'.xlsx');
  await new Promise((resolve,reject)=>{
    const p=fileUrl.startsWith('https')?https:http;
    const f=fs.createWriteStream(tmp);
    p.get(fileUrl,res=>{res.pipe(f);f.on('finish',()=>{f.close();resolve();});}).on('error',reject);
  });
  const wb=XLSX.readFile(tmp);
  try{fs.unlinkSync(tmp);}catch(e){}
  const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
  stockCache[fileUrl]=rows;
  return rows;
}

const STOCK_KW = /\b(stock|nos|sqm|inventory|quantity|available|item|size|location|site|batch|ahm|del|bbs|bhw|ddn|gwh|hbl|ind|jai|klk|lko|ngp|pat|pnb|pune|rai|rnc|umg)\b/i;

app.post('/api/admin/login',(req,res)=>{ const{username,password}=req.body; if(username==='admin'&&password==='eurobond@2024'){return res.json({success:true,token:jwt.sign({role:'admin'},process.env.JWT_SECRET,{expiresIn:'24h'})});} res.status(401).json({success:false,message:'Invalid credentials'}); });

app.post('/api/auth/login-email',async(req,res)=>{ const{email,password}=req.body; const users=await getUsers(); const user=users.find(u=>u.email===email&&u.active); if(!user) return res.status(404).json({success:false,message:'User not found. Contact admin.'}); if(user.password!==password) return res.status(400).json({success:false,message:'Invalid password'}); await addLog({name:user.name,phone:user.phone||'',email:user.email,loginTime:new Date()}); const token=jwt.sign({id:user._id?.toString()||user.id,name:user.name,role:user.role,region:user.region},process.env.JWT_SECRET,{expiresIn:'7d'}); res.json({success:true,token,user:{name:user.name,role:user.role,region:user.region,email:user.email}}); });

app.post('/api/chat',authMiddleware,async(req,res)=>{
  const{messages}=req.body;
  const apiKey=process.env.ANTHROPIC_API_KEY;
  if(!apiKey) return res.json({success:true,reply:'AI not configured.'});

  const resources=await getResources();
  const userMsg=messages[messages.length-1]?.content||'';
  // Extract item codes from conversation history for follow-up messages
  const historyText=messages.map(m=>m.content).join(' ');
  const enrichedMsg=extractCodes(userMsg).length>0 ? userMsg : userMsg+' '+extractCodes(historyText).join(' ');
  const isStock=STOCK_KW.test(userMsg)||extractCodes(enrichedMsg).length>0;

  let resourceContext='';
  let excelDownloadUrl='';

  for(const f of resources.files){
    const ext=(f.filename||'').toLowerCase();
    if(ext.includes('.xlsx')||ext.includes('.xls')){
      const fileUrl=f.cloudinaryUrl||f.url;
      excelDownloadUrl=fileUrl;
      if(fileUrl){
        try{
          const rows=await getStockRows(fileUrl);
          const isStockFile=(f.name||'').toLowerCase().includes('stock');
          if(isStockFile&&isStock){
            const stockData=processStock(rows,enrichedMsg);
            if(stockData) resourceContext+=`Stock: ${f.name}\n${stockData}\n`;
          } else {
            const words=userMsg.toUpperCase().split(/[\s,]+/).filter(w=>w.length>2);
            const filtered=rows.filter(row=>{
              const rowStr=Object.values(row).join(' ').toUpperCase();
              return words.some(w=>rowStr.includes(w));
            });
            if(filtered.length>0){
              resourceContext+=`File: ${f.name}\n`;
              filtered.slice(0,5).forEach(row=>{
                resourceContext+=Object.entries(row).map(([k,v])=>`${k}: ${v}`).join(' | ')+'\n';
              });
              if(filtered.length>5) resourceContext+=`+${filtered.length-5} more results\n`;
            }
          }
        }catch(e){ resourceContext+=`File: ${f.name} available\n`; }
      }
    } else {
      resourceContext+=`File: ${f.name} | Link: ${f.cloudinaryUrl||f.url}\n`;
    }
  }

  const linksCtx=resources.links.map(l=>`${l.title}: ${l.url}`).join('\n');

  const systemPrompt=`You are EUROBOND AI Sales Assistant.

RULES:
1. LANGUAGE: Always reply in exact same language user writes. Telugu→Telugu, Hindi→Hindi, English→English, any language→same language.
2. GENERAL QUESTIONS: Answer immediately and helpfully like ChatGPT. Hi/hello → friendly 1-2 line reply. General questions → answer directly.
3. STOCK DATA: Show ONLY exact data from RESOURCES. Valid locations with full names: AHM=Ahmedabad, BBS=Bhubaneswar, BHW=Bhiwandi, DDN=Dehradun, DEL=Delhi, GWH=Guwahati, HBL=Hubballi, IND=Indore, JAI=Jaipur, KLK=Kolkata, LKO=Lucknow, NGP=Nagpur, PAT=Patna, PNB=Punjab, PUNE=Pune, RAI=Raipur, RNC=Ranchi, UMG=Factory only. NEVER invent other locations.
4. BATCH: Show exact batch numbers from data. Never guess.
5. CONTACTS: Always share contact details when asked from data.
6. LINKS: Use EXACT URLs only. Never modify.
7. FORMAT: Short replies. Plain text. No tables.
8. ERRORS: Never say "something went wrong" or "unable to process". Always give a helpful reply.
9. GENERAL KNOWLEDGE: Answer any question - distances, calculations, geography, history, science, business, etc. like a knowledgeable assistant.
10. SALES HELP: Help with product recommendations, pricing queries, customer objections, presentation tips.
11. Be friendly, professional and helpful in every response.

User: ${req.user.name} | Role: ${req.user.role} | Region: ${req.user.region}
${resourceContext?'RESOURCES:\n'+resourceContext:''}
${linksCtx?'LINKS:\n'+linksCtx:''}`;

  try{
    const response=await fetch('https://api.anthropic.com/v1/messages',{
      method:'POST',
      headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      body:JSON.stringify({model:'claude-haiku-4-5-20251001',max_tokens:isStock?1500:500,system:systemPrompt,messages:messages.slice(-10)})
    });
    const data=await response.json();
    if(data.error){
      // Rate limit or other error - give friendly message
      const errMsg=data.error.message||'';
      if(errMsg.includes('rate limit')||errMsg.includes('tokens')){
        return res.json({success:true,reply:'Please wait a moment and try again.'});
      }
      return res.json({success:true,reply:'Unable to process request. Please try again.'});
    }
    res.json({success:true,reply:data.content[0].text});
  }catch(err){
    res.json({success:true,reply:'Connection issue. Please try again in a moment.'});
  }
});

app.get('/api/admin/users',adminMiddleware,async(req,res)=>res.json({success:true,users:await getUsers()}));
app.post('/api/admin/users',adminMiddleware,async(req,res)=>{ const{name,phone,email,password,role,region}=req.body; const users=await getUsers(); if(users.find(u=>u.email===email)) return res.status(400).json({success:false,message:'Email already exists'}); await saveUser({id:Date.now().toString(),name,phone,email,password:password||'eurobond123',role,region,active:true,createdAt:new Date().toISOString(),lastLogin:null}); res.json({success:true}); });
app.delete('/api/admin/users/:id',adminMiddleware,async(req,res)=>{await deleteUser(req.params.id);res.json({success:true});});
app.post('/api/admin/upload',adminMiddleware,upload.single('file'),async(req,res)=>{ if(!req.file) return res.status(400).json({success:false,message:'No file'}); Object.keys(stockCache).forEach(k=>delete stockCache[k]); await addResource('file',{id:Date.now().toString(),name:req.body.title||req.file.originalname,filename:req.file.originalname,cloudinaryUrl:req.file.path,url:req.file.path,type:req.file.mimetype,size:req.file.size,uploadedAt:new Date().toISOString()}); res.json({success:true}); });
app.delete('/api/admin/files/:id',adminMiddleware,async(req,res)=>{await deleteResource('file',req.params.id);res.json({success:true});});
app.post('/api/admin/links',adminMiddleware,async(req,res)=>{ const{title,url}=req.body; await addResource('link',{id:Date.now().toString(),title,url,addedAt:new Date().toISOString()}); res.json({success:true}); });
app.delete('/api/admin/links/:id',adminMiddleware,async(req,res)=>{await deleteResource('link',req.params.id);res.json({success:true});});
app.get('/api/admin/resources',adminMiddleware,async(req,res)=>res.json({success:true,resources:await getResources()}));
app.get('/api/admin/logs',adminMiddleware,async(req,res)=>res.json({success:true,logs:await getLogs()}));
app.get('/api/health',(req,res)=>res.json({status:'ok',apiKey:process.env.ANTHROPIC_API_KEY?.startsWith('sk-')?'configured':'not configured',mongodb:db?'connected':'disconnected'}));

const PORT=process.env.PORT||5000;
app.listen(PORT,()=>{console.log('EUROBOND Backend running on port '+PORT);getDB();});
