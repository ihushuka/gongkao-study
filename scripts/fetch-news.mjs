import { mkdir, writeFile } from "node:fs/promises";

const now = new Date();
const fmt = new Intl.DateTimeFormat("en-CA", {timeZone:"Asia/Shanghai",year:"numeric",month:"2-digit",day:"2-digit"}).formatToParts(now);
const get=t=>fmt.find(x=>x.type===t)?.value;
const today=`${get('year')}-${get('month')}-${get('day')}`;

const clean=s=>String(s||"").replace(/<[^>]*>/g,"").replace(/&amp;/g,"&").replace(/&quot;/g,'"').replace(/\s+/g," ").trim();
const abs=(u,b)=>{try{return new URL(u,b).href}catch{return ""}};
async function fetchText(url){
 const r=await fetch(url,{headers:{"user-agent":"Mozilla/5.0"}});
 if(!r.ok) throw Error(url);
 return await r.text();
}
const unique=a=>Array.from(new Map(a.filter(x=>x.title&&x.link).map(x=>[x.title,x])).values()).slice(0,10);

async function people(){
 let rows=[];
 const urls=[
  "https://www.people.com.cn/",
  "http://www.people.com.cn/GB/59476/index.html"
 ];
 for(const url of urls){
  try{
   const html=await fetchText(url);
   const reg=/<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
   let m;
   while((m=reg.exec(html))){
    let title=clean(m[2]);
    if(title.length>8) rows.push({title,link:abs(m[1],url),source:"人民日报·人民网",date:today});
   }
  }catch(e){console.log(e.message)}
 }
 return unique(rows);
}

async function gov(){
 let rows=[];
 const urls=["https://www.gov.cn/yaowen/index.htm","https://www.gov.cn/"];
 for(const url of urls){
  try{
   const html=await fetchText(url);
   const reg=/<a[^>]+href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
   let m;
   while((m=reg.exec(html))){
    let title=clean(m[2]);
    if(title.length>8) rows.push({title,link:abs(m[1],url),source:"中国政府网",date:today});
   }
  }catch(e){console.log(e.message)}
 }
 return unique(rows);
}

const xuexi=["https://www.xuexi.cn/lgdata/1jscb6pu1n2.json","https://www.xuexi.cn/lgdata/1crqb964p71.json"];
function walk(v,r){
 if(!v||typeof v!=="object")return;
 if(Array.isArray(v)){v.forEach(x=>walk(x,r));return;}
 const t=clean(v.title||v.name||v.text),u=v.url||v.link;
 if(t.length>8&&u)r.push({title:t,link:abs(u,"https://www.xuexi.cn"),source:"学习强国",date:today});
 Object.values(v).forEach(x=>walk(x,r));
}
async function xuexiNews(){let r=[];for(const u of xuexi){try{walk(JSON.parse(await fetchText(u)),r)}catch(e){}}return unique(r)}

const result=await Promise.allSettled([xuexiNews(),people(),gov()]);
const getRows=i=>result[i].status==="fulfilled"?result[i].value:[];
const out={date:today,updatedAt:now.toISOString(),sources:{xuexi:getRows(0),people:getRows(1),gov:getRows(2)}};
await mkdir("public",{recursive:true});
await writeFile("public/daily-news.json",JSON.stringify(out,null,2),"utf8");
console.log(`news ${today}:`,out.sources);
