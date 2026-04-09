import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

const WD={Commit:100,Upside:50,Pipeline:20,"N/A":0};
const QQ={
  NA:{total:21606356,cloud:6610587,op:14995769,fy26Cloud:532253,fy26Op:11370709,
    Q:{Q1:{t:12540356,c:1115587,o:11424769,s:"2026-04-01",e:"2026-06-30"},
       Q2:{t:15101356,c:2110587,o:12990769,s:"2026-07-01",e:"2026-09-30"},
       Q3:{t:18186356,c:3550587,o:14635769,s:"2026-10-01",e:"2026-12-31"},
       Q4:{t:21606356,c:6610587,o:14995789,s:"2027-01-01",e:"2027-03-31"}}},
  EMEA:{total:4529189,cloud:2930308,op:1598881,fy26Cloud:628881,fy26Op:343641,
    Q:{Q1:{t:1362522,c:733641,o:628881,s:"2026-04-01",e:"2026-06-30"},
       Q2:{t:2197522,c:1113641,o:1083881,s:"2026-07-01",e:"2026-09-30"},
       Q3:{t:3330855,c:1851974,o:1478881,s:"2026-10-01",e:"2026-12-31"},
       Q4:{t:4529189,c:2930308,o:1598881,s:"2027-01-01",e:"2027-03-31"}}},
};

const RAW=[
  ["NA","TiDB Cloud","Voxo","Voxo - MySQL Consolidation","Claiborne Adams","Non-KA","5/1/2026",42000,42000,"Commit","Commit","Negotiation",60],
  ["NA","TiDB Cloud","PayNearMe","PayNearMe Aurora Global Replace","Estyn Cannan","KA","5/1/2026",216000,216000,"Upside","Upside","Technical Validation",30],
  ["NA","TiDB Cloud","Empower Project","Empower Project - TiDB X Migration","Claiborne Adams","Non-KA","4/30/2026",26400,26400,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Cloud","Meza AI","Meza Ai - TiDB Migration","Claiborne Adams","Non-KA","6/30/2026",26400,26400,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Cloud","Solvea","Solvea Aurora Replacement","Fan Wang","Non-KA","6/1/2026",60000,60000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Cloud","Connecty AI","ConnectyAI - Partner integration","Christopher Hofmann","Non-KA","6/30/2026",26400,26400,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Cloud","futureagi.com","Future AGI - Partner integration","Christopher Hofmann","Non-KA","6/30/2026",26400,26400,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Cloud","Zeta Global","Zeta Global","Yashasvi Chandrabhatta","KA","8/31/2026",120000,120000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Cloud","OneTrust","OneTrust CosmosDB Replace","Yashasvi Chandrabhatta","KA","12/31/2026",120000,120000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Cloud","SafetyCulture","Citus Postgres Replace","Yashasvi Chandrabhatta","Non-KA","11/30/2026",840000,840000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Cloud","Audible","Audible HTAP Data Platform Build","Yashasvi Chandrabhatta","Non-KA","12/20/2026",120000,120000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Cloud","Prove","Snowflake Replacement Project","Yashasvi Chandrabhatta","Non-KA","2/28/2027",60000,60000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Cloud","Monitor Latino","TiDB Cloud on Alibaba - MonitorLatino","Zeno Wang","Non-KA","9/30/2026",96000,96000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Affirm Inc","Affirm landing opp","Louis Fahrberger","SKA","6/30/2026",250000,250000,"Upside","Upside","Business Validation",50],
  ["NA","TiDB Enterprise Subscription","Atlassian Pty Ltd","Atlassian - True Up Event","Aaron Alvarez","SKA","6/30/2026",588000,588000,"Upside","Upside","Technical Validation",30],
  ["NA","TiDB Enterprise Subscription","Amazon.com Inc.","Amazon - Fraud Detection Modernization","Sawyer Hulme","SKA","6/30/2026",250000,250000,"Upside","Upside","Technical Validation",30],
  ["NA","TiDB Enterprise Subscription","Bling","Bling - AWS Virginia","Aaron Alvarez","Non-KA","6/15/2026",300000,300000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","StubHub","Stubhub - Move from SQL server","Scott Einaugler","Non-KA","6/15/2026",150000,150000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Heartland Dental","Heartland - Move and Consolidate DBs","Scott Einaugler","Non-KA","6/30/2026",150000,150000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Coupang Inc.","Coupang - MySQL Modernization","Estyn Cannan","SKA","3/26/2027",120000,120000,"Upside","Upside","Business Validation",50],
  ["NA","TiDB Enterprise Subscription","Visa Inc.","Visa TiDB On Premise - OP","Louis Fahrberger","SKA","9/30/2026",249600,249600,"Upside","Upside","Business Validation",50],
  ["NA","TiDB Enterprise Subscription","Life360","Life360 - Aurora MySQL Replacement","Sawyer Hulme","Non-KA","7/10/2026",150000,150000,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Enterprise Subscription","Kraken","Kraken - MariaDB Modernization","Sawyer Hulme","Non-KA","7/24/2026",250000,250000,"Pipeline","N/A","Technical Validation",30],
  ["NA","TiDB Enterprise Subscription","Intuit Inc.","Intuit - AntiFraud ODS","Aaron Alvarez","SKA","8/5/2026",500000,500000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Pluralsight","PluralSight - Aurora Postgres to TiDB","Scott Einaugler","Non-KA","9/20/2026",150000,150000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Intuit Inc.","Intuit E-Filing Application Oracle Replace","Aaron Alvarez","SKA","9/30/2026",300000,300000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Walmart","Walmart - No More Sharding","Scott Einaugler","SKA","9/30/2026",150000,150000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Salesforce Inc.","Razlink - Salesforce HTAP and AI","Louis Fahrberger","SKA","10/30/2026",250000,250000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Reddit","Reddit - ThingDB replace","Estyn Cannan","KA","11/11/2026",240000,240000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Block Inc.","Block - Aurora MySQL/Vitess Replacement","Sawyer Hulme","SKA","11/24/2026",150000,150000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Visa Inc.","Visa ERS - CS Reconciliation","Louis Fahrberger","SKA","11/30/2026",480000,480000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Visa Inc.","Visa - AI Next Gen Payment Platform","Louis Fahrberger","SKA","11/30/2026",480000,480000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Robinhood","Robinhood - PostgresAurora Replace","Estyn Cannan","SKA","11/30/2026",360000,360000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Planview","Planview - Switch from Postgres to TiDB","Scott Einaugler","KA","12/1/2026",150000,150000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Kochava","Kochava - TiDB adoption","Aaron Alvarez","Non-KA","12/18/2026",500000,500000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","HubSpot Inc.","HubSpot - HBase Replacement","Sawyer Hulme","SKA","12/19/2026",250000,250000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","Etsy Inc","Etsy - MySQL Monolith Replacement","Sawyer Hulme","SKA","12/19/2026",150000,150000,"Pipeline","N/A","Discovery & Qualification",20],
  ["NA","TiDB Enterprise Subscription","UnitedHealth Group","United health - Database migration","Scott Einaugler","Non-KA","3/1/2027",250000,250000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Drata","Drata - Aurora MySQL Replacement","Sawyer Hulme","KA","3/6/2027",150000,150000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Abbott","Abbott - Move to TiDB","Scott Einaugler","SKA","3/25/2027",48000,48000,"Pipeline","N/A","Prospecting",10],
  ["NA","TiDB Enterprise Subscription","Axon","Axon MySQL Replacement","Louis Fahrberger","SKA","3/30/2027",250000,250000,"Pipeline","N/A","Prospecting",10],
  ["EMEA","TiDB Cloud","Bolt Technology","Education Bolt","Nikita Khristenko","SKA","5/1/2026",0,0,"Commit","Commit","Negotiation",60],
  ["EMEA","TiDB Cloud","CAPITAL COM IP LTD","Capital.com Fixed term contract","Jamie Mckee","Non-KA","5/1/2026",320000,320000,"Commit","Commit","Negotiation",60],
  ["EMEA","TiDB Cloud","Finaloop","Finaloop - Snowflake replace","Jamie Mckee","Non-KA","5/1/2026",50000,50000,"Commit","Commit","Negotiation",60],
  ["EMEA","TiDB Cloud","Bolt Technology","Bolt TiDBX","Nikita Khristenko","SKA","4/1/2026",50000,50000,"Upside","Upside","Negotiation",60],
  ["EMEA","TiDB Cloud","OFZIO","OFZIO MySQL Migration","Nikita Khristenko","Non-KA","5/13/2026",50000,50000,"Pipeline","N/A","Business Validation",50],
  ["EMEA","TiDB Cloud","Wafi Group","TiDB Cloud on Alibaba - Wafi Group","Zeno Wang","Non-KA","9/30/2026",20000,20000,"Pipeline","N/A","Prospecting",10],
  ["EMEA","TiDB Cloud","Mr. Mandoob","TiDB Cloud on Alibaba - Mr Mandoob","Zeno Wang","Non-KA","9/30/2026",40000,40000,"Pipeline","N/A","Prospecting",10],
  ["EMEA","TiDB Cloud","Trustly","Trustly - MySQL replace - Dedicated AWS","Jamie Mckee","Non-KA","10/1/2026",100000,100000,"Pipeline","N/A","Discovery & Qualification",20],
  ["EMEA","TiDB Cloud","Etraveli Group","Etraveli Group - New Logo Training","Jamie Mckee","Non-KA","12/1/2026",0,0,"Upside","Upside","Negotiation",60],
  ["EMEA","TiDB Enterprise Subscription","Bolt Technology","Bolt renewalFY27","Nikita Khristenko","SKA","5/1/2026",154640,154640,"Commit","Commit","Negotiation",60],
  ["EMEA","TiDB Enterprise Subscription","LHV Bank","LHV PilotProject","Nikita Khristenko","SKA","6/1/2026",50000,50000,"Upside","Upside","Business Validation",50],
  ["EMEA","TiDB Enterprise Subscription","Akamai Technologies EMEA","Akamai TiKV migration","Nikita Khristenko","Non-KA","5/29/2026",300000,300000,"Pipeline","N/A","Business Validation",50],
  ["EMEA","TiDB Enterprise Subscription","Merck KGaA","Merck Group - IAM expansion","Udo Straesser","Non-KA","6/25/2026",100000,100000,"Pipeline","N/A","Technical Validation",30],
  ["EMEA","TiDB Enterprise Subscription","Mayflower","Mayflower TiDBMigration","Nikita Khristenko","Non-KA","9/17/2026",350000,350000,"Pipeline","N/A","Technical Validation",30],
  ["EMEA","TiDB Enterprise Subscription","Forter","Forter - Aurora MySQL replace","Jamie Mckee","Non-KA","10/1/2026",120000,120000,"Pipeline","N/A","Discovery & Qualification",20],
  ["EMEA","TiDB Enterprise Subscription","Batch","Batch Cassandra migration - AWS cloud","Udo Straesser","Non-KA","10/29/2026",150000,150000,"Pipeline","N/A","Technical Validation",30],
  ["EMEA","TiDB Enterprise Subscription","Siemens","Siemens TiDB migration","Nikita Khristenko","Non-KA","10/15/2026",200000,200000,"Pipeline","N/A","Prospecting",10],
  ["EMEA","TiDB Enterprise Subscription","Etraveli Group","Etraveli Group - MySQL Replacement Year 1","Jamie Mckee","Non-KA","12/1/2026",105000,105000,"Upside","Upside","Negotiation",60],
];

const INIT=RAW.map(r=>({
  region:r[0],product:r[1],account:r[2],opp:r[3],owner:r[4],seg:r[5],
  launchDate:r[6],netNewARR:parseInt(r[7])||0,fcstARR:parseInt(r[8])||0,
  fcstType:r[9],fcstTypeOverride:r[10],stage:r[11],prob:parseInt(r[12])||0,weightedARR:null
}));

// Date helpers - all local time to avoid UTC timezone issues
const pd=str=>{
  if(!str)return null;
  const p=str.split("/");
  if(p.length===3)return new Date(parseInt(p[2]),parseInt(p[0])-1,parseInt(p[1])).getTime();
  return new Date(str).getTime();
};
const qS=str=>{const p=str.split("-");return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2])).getTime();};
const qE=str=>{const p=str.split("-");return new Date(parseInt(p[0]),parseInt(p[1])-1,parseInt(p[2]),23,59,59).getTime();};

// Styles
const CY="#00FFFF",YW="#FFFF00",TL="#00B0A0",LB="#ADD8E6";
const S={
  wrap:{fontFamily:"Calibri,Arial,sans-serif",fontSize:13,background:"#fff",minHeight:"100vh"},
  tabs:{display:"flex",borderBottom:"2px solid #aaa",background:"#f0f0f0",padding:"0 8px",overflowX:"auto"},
  tab:(a,c)=>({padding:"7px 14px",cursor:"pointer",border:"1px solid #bbb",borderBottom:a?"2px solid #fff":"none",background:a?(c||"#fff"):"#e0e0e0",fontWeight:a?700:400,borderRadius:"4px 4px 0 0",marginRight:2,marginTop:4,fontSize:12,whiteSpace:"nowrap"}),
  th:(bg,fg="#000")=>({background:bg,color:fg,padding:"5px 8px",border:"1px solid #bbb",fontWeight:700,whiteSpace:"nowrap",fontSize:11}),
  td:(bg="#fff",bold=false,align="left")=>({background:bg,padding:"4px 8px",border:"1px solid #ddd",fontWeight:bold?"700":"400",textAlign:align,whiteSpace:"nowrap",fontSize:11}),
};
const fmt=n=>n===0||n==null?"0":Number(n).toLocaleString();
const FC={Commit:YW,Upside:YW,Pipeline:"#fff","N/A":"#fff"};
const SC={"Negotiation":"#C6EFCE","Business Validation":"#DDEBF7","Technical Validation":"#FFF2CC","Discovery & Qualification":"#FCE4D6","Prospecting":"#F2F2F2"};
const FO=["Commit","Upside","Pipeline","N/A"];
const SO=["Negotiation","Business Validation","Technical Validation","Discovery & Qualification","Prospecting"];
const GO=["KA","SKA","Non-KA"];
const THDR=["Region","Product","Account Name","Pipeline Name","Account Owner","Segmentation","Launch Date","Q Net New ARR $","Forecast Net New ARR$","Forecast Type","Forecast Type (Override)","Forecast Weighted Net New ARR","Opp Stage","Prob %"];
const toTSV=(h,rows)=>[h,...rows].map(r=>r.join("\t")).join("\n");
const toArr=r=>[r.region,r.product,r.account,r.opp,r.owner,r.seg,r.launchDate,r.netNewARR,r.fcstARR,r.fcstType,r.fcstTypeOverride,r.weightedARR??0,r.stage,r.prob];
const cp=(t,l)=>navigator.clipboard.writeText(t).then(()=>alert("Copied: "+l+"\nGo to Google Sheets -> A1 -> Ctrl+V")).catch(()=>alert("Copy failed"));

// Core calculation engine
function useCalc(data,weights,baselines){
  const W=k=>(weights[k]||0)/100;
  const wtOf=r=>r.weightedARR!=null?r.weightedARR:Math.round((r.fcstARR||0)*W(r.fcstTypeOverride||"N/A"));
  const pwOf=r=>Math.round((r.fcstARR||0)*(r.prob||0)/100);
  const QKS=["Q1","Q2","Q3","Q4"];

  const filt=(region,prod,s,e)=>data.filter(r=>{
    if(r.region!==region)return false;
    if(prod&&r.product!==prod)return false;
    if(!prod&&r.product!=="TiDB Cloud"&&r.product!=="TiDB Enterprise Subscription")return false;
    const d=pd(r.launchDate);
    return d!=null&&(s==null||d>=s)&&d<=e;
  });

  const calc=(region,cat)=>{
    const q=QQ[region],bl=baselines[region];
    const prod=cat==="Cloud"?"TiDB Cloud":cat==="OP"?"TiDB Enterprise Subscription":null;
    const fy26=cat==="Cloud"?bl.fy26Cloud:cat==="OP"?bl.fy26Op:bl.fy26Cloud+bl.fy26Op;
    const fyQ=cat==="Cloud"?q.cloud:cat==="OP"?q.op:q.total;

    const qRows=QKS.map(qk=>{
      const s=qS(q.Q[qk].s),e=qE(q.Q[qk].e);
      const inc=filt(region,prod,s,e);
      const cum=filt(region,prod,null,e);
      const quota=cat==="Cloud"?q.Q[qk].c:cat==="OP"?q.Q[qk].o:q.Q[qk].t;
      const net=inc.reduce((a,r)=>a+(r.fcstARR||0),0);
      const wt =inc.reduce((a,r)=>a+wtOf(r),0);
      const pwt=inc.reduce((a,r)=>a+pwOf(r),0);
      const cNet=cum.reduce((a,r)=>a+(r.fcstARR||0),0);
      const cWt =cum.reduce((a,r)=>a+wtOf(r),0);
      const cPwt=cum.reduce((a,r)=>a+pwOf(r),0);
      const fcst=fy26+cNet,wfcst=fy26+cWt,pfcst=fy26+cPwt;
      const gap=fcst-quota,att=quota?(fcst/quota*100).toFixed(0)+"%":"0%";
      return {quota,net,wt,pwt,fcst,wfcst,pfcst,gap,att};
    });

    const fyNet=qRows.reduce((a,r)=>a+r.net,0);
    const fyWt =qRows.reduce((a,r)=>a+r.wt,0);
    const fyPwt=qRows.reduce((a,r)=>a+r.pwt,0);
    const fyF=fy26+fyNet,fyWF=fy26+fyWt,fyPF=fy26+fyPwt;
    const fyG=fyF-fyQ,fyWG=fyWF-fyQ,fyPG=fyPF-fyQ;
    const fyA=fyQ?(fyF/fyQ*100).toFixed(0)+"%":"0%";
    return {qRows,fy:{quota:fyQ,net:fyNet,wt:fyWt,pwt:fyPwt,fcst:fyF,wfcst:fyWF,pfcst:fyPF,gap:fyG,wgap:fyWG,pgap:fyPG,att:fyA},fy26};
  };
  return {calc,wtOf,pwOf};
}

// Summary Tab
function SummaryTab({data,baselines,setBL,weights,setWeights}){
  const {calc}=useCalc(data,weights,baselines);
  const HDRS=["Region","FY27 Quota","FY26 Ending ARR","FY27 Net New ARR","FY27 Ending ARR (FCST)","Attain%","GAP"];
  const HBG=i=>i===0?"#4472C4":i===1?LB:i===2?YW:i===3?"#fff":i===4?CY:TL;
  const HFG=i=>(i===0||i===6)?"#fff":"#000";
  const CBG=i=>i===1?LB:i===2?YW:i===3?"#fff":i===4?CY:TL;

  return (
    <div style={{padding:14}}>
      <div style={{color:"#888",fontSize:11,marginBottom:10}}>FY27 Full Year - 4/1/2026 to 3/31/2027 - {new Date().toLocaleDateString()}</div>
      <div style={{background:"#faf5ff",border:"1px solid #d8b4fe",borderRadius:8,padding:12,marginBottom:16}}>
        <div style={{fontWeight:700,fontSize:12,color:"#7e22ce",marginBottom:8}}>Forecast Type Weights</div>
        <div style={{display:"flex",gap:16,flexWrap:"wrap",alignItems:"center"}}>
          {Object.keys(weights).map(k=>(
            <label key={k} style={{fontSize:12,display:"flex",alignItems:"center",gap:5}}>
              <span style={{fontWeight:700,color:k==="Commit"?"#16a34a":k==="Upside"?"#ca8a04":k==="Pipeline"?"#2563eb":"#9ca3af"}}>{k}</span>
              <input type="number" min={0} max={100} value={weights[k]} onChange={e=>setWeights(w=>({...w,[k]:parseInt(e.target.value)||0}))}
                style={{width:52,border:"1px solid #d8b4fe",borderRadius:4,padding:"2px 5px",fontSize:12,textAlign:"center"}}/>
              <span style={{color:"#888"}}>%</span>
            </label>
          ))}
          <button onClick={()=>setWeights({...WD})} style={{background:"#7e22ce",color:"#fff",border:"none",borderRadius:4,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>Reset</button>
        </div>
      </div>
      {["NA","EMEA"].map(region=>{
        const bl=baselines[region];
        return (
          <div key={region} style={{marginBottom:24}}>
            <div style={{display:"flex",gap:10,marginBottom:8,flexWrap:"wrap",background:"#f8f8f8",padding:8,borderRadius:6,border:"1px solid #ddd",alignItems:"center"}}>
              <span style={{fontWeight:700,color:region==="NA"?"#0070C0":"#217346",fontSize:13}}>{region}</span>
              <label style={{fontSize:11}}>Cloud FY26 ARR:
                <input type="number" value={bl.fy26Cloud} onChange={e=>setBL(b=>({...b,[region]:{...b[region],fy26Cloud:parseInt(e.target.value)||0}}))}
                  style={{border:"1px solid #bbb",borderRadius:3,padding:"2px 4px",fontSize:11,width:95,marginLeft:3}}/>
              </label>
              <label style={{fontSize:11}}>OP FY26 ARR:
                <input type="number" value={bl.fy26Op} onChange={e=>setBL(b=>({...b,[region]:{...b[region],fy26Op:parseInt(e.target.value)||0}}))}
                  style={{border:"1px solid #bbb",borderRadius:3,padding:"2px 4px",fontSize:11,width:95,marginLeft:3}}/>
              </label>
            </div>
            {["Total","Cloud","OP"].map((cat,ci)=>{
              const {fy,fy26}=calc(region,cat);
              const wvals=[region,fy.quota,fy26,fy.wt,fy.wfcst,fy.quota?(fy.wfcst/fy.quota*100).toFixed(0)+"%":"0%",fy.wgap];
              const pvals=[region,fy.quota,fy26,fy.pwt,fy.pfcst,fy.quota?(fy.pfcst/fy.quota*100).toFixed(0)+"%":"0%",fy.pgap];
              return (
                <div key={cat} style={{marginBottom:10}}>
                  {ci>0&&<div style={{fontWeight:700,color:ci===1?"#0070C0":"#FF6600",marginBottom:2,fontSize:12}}>{cat} Forecast</div>}
                  <table style={{borderCollapse:"collapse"}}>
                    <thead><tr>{HDRS.map((h,i)=><th key={i} style={S.th(HBG(i),HFG(i))}>{h}</th>)}</tr></thead>
                    <tbody>
                      <tr>
                        {[region,fy.quota,fy26,fy.net,fy.fcst,fy.att,fy.gap].map((v,i)=>(
                          <td key={i} style={S.td(CBG(i),false,"right")}>
                            {i===0?<span style={{fontSize:10,color:"#555"}}>Forecast</span>:i===6?<span style={{color:fy.gap<0?"red":"#000"}}>{fmt(v)}</span>:i!==5?fmt(v):v}
                          </td>
                        ))}
                      </tr>
                      <tr style={{borderTop:"1px dashed #aaa"}}>
                        {wvals.map((v,i)=>(
                          <td key={i} style={S.td(i===4?"#d4edda":i===6?"#7B5EA7":CBG(i),false,"right")}>
                            {i===0?<span style={{fontSize:10,color:"#7e22ce",fontWeight:700}}>Forecast Weighted</span>:i===6?<span style={{color:fy.wgap<0?"red":"#fff"}}>{fmt(v)}</span>:i!==5?fmt(v):v}
                          </td>
                        ))}
                      </tr>
                      <tr style={{borderTop:"1px dashed #aaa"}}>
                        {pvals.map((v,i)=>(
                          <td key={i} style={S.td(i===4?"#fde68a":i===6?"#b45309":CBG(i),false,"right")}>
                            {i===0?<span style={{fontSize:10,color:"#b45309",fontWeight:700}}>Probability Weighted</span>:i===6?<span style={{color:fy.pgap<0?"red":"#fff"}}>{fmt(v)}</span>:i!==5?fmt(v):v}
                          </td>
                        ))}
                      </tr>
                    </tbody>
                  </table>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

// Quarterly Tab
function QuarterlyTab({data,baselines,weights}){
  const {calc}=useCalc(data,weights,baselines);
  const QKS=["Q1","Q2","Q3","Q4"];
  const QLBL={Q1:"Q1 Apr-Jun 26",Q2:"Q2 Jul-Sep 26",Q3:"Q3 Oct-Dec 26",Q4:"Q4 Jan-Mar 27"};
  const CC={Total:"#1a3a5c",Cloud:"#0070C0",OP:"#FF6600"};

  return (
    <div style={{padding:14}}>
      <div style={{color:"#888",fontSize:11,marginBottom:12}}>FY27 Quarterly Breakdown - Incremental Net New per quarter, Cumulative Ending ARR</div>
      {["NA","EMEA"].map(region=>(
        <div key={region} style={{marginBottom:24}}>
          <div style={{fontWeight:700,fontSize:13,color:region==="NA"?"#0070C0":"#217346",marginBottom:8,padding:"5px 10px",background:region==="NA"?"#e8f4ff":"#e8f9f0",borderRadius:6}}>
            {region} - Quarterly Breakdown
          </div>
          <div style={{overflowX:"auto"}}>
            <table style={{borderCollapse:"collapse"}}>
              <thead>
                <tr>
                  <th style={S.th("#4472C4","#fff")}>Category</th>
                  <th style={S.th("#4472C4","#fff")}>Metric</th>
                  {QKS.map(qk=><th key={qk} style={S.th("#4472C4","#fff")}>{QLBL[qk]}</th>)}
                  <th style={S.th("#1a3a5c","#fff")}>FY27 Full Year</th>
                </tr>
              </thead>
              <tbody>
                {["Total","Cloud","OP"].map((cat,ci)=>{
                  const {qRows,fy}=calc(region,cat);
                  const bg=ci%2===0?"#fff":"#fafafa";
                  const ms=[
                    {lbl:"Quota",                   vals:qRows.map(r=>r.quota), fyV:fy.quota, bg:LB},
                    {lbl:"Net New ARR",              vals:qRows.map(r=>r.net),   fyV:fy.net,   bg:"#fff"},
                    {lbl:"Forecast Weighted Net New ARR",     vals:qRows.map(r=>r.wt),    fyV:fy.wt,    bg:"#e6ccff"},
                    {lbl:"Probability Weighted Net New ARR",vals:qRows.map(r=>r.pwt),   fyV:fy.pwt,   bg:"#fef3c7"},
                    {lbl:"Ending ARR (FCST)",        vals:qRows.map(r=>r.fcst),  fyV:fy.fcst,  bg:CY},
                    {lbl:"Forecast Weighted Ending ARR",      vals:qRows.map(r=>r.wfcst), fyV:fy.wfcst, bg:"#d4edda"},
                    {lbl:"Probability Weighted Ending ARR", vals:qRows.map(r=>r.pfcst), fyV:fy.pfcst, bg:"#fde68a"},
                  ];
                  return [
                    ...ms.map((m,mi)=>(
                      <tr key={cat+mi} style={{background:bg}}>
                        {mi===0&&<td style={{...S.td(),fontWeight:700,color:CC[cat],borderRight:"2px solid #ddd",verticalAlign:"middle"}} rowSpan={10}>{cat}</td>}
                        <td style={S.td(m.bg,true)}>{m.lbl}</td>
                        {m.vals.map((v,i)=><td key={i} style={S.td(m.bg,false,"right")}>{fmt(v)}</td>)}
                        <td style={S.td(m.bg,true,"right")}>{fmt(m.fyV)}</td>
                      </tr>
                    )),
                    <tr key={cat+"gap1"} style={{background:bg,borderBottom:"1px dashed #aaa"}}>
                      <td style={S.td(TL,true)}><span style={{color:"#fff"}}>GAP / Attain%</span></td>
                      {qRows.map((r,i)=>(
                        <td key={i} style={S.td(TL,false,"right")}>
                          <div style={{color:r.gap<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(r.gap)}</div>
                          <div style={{color:"#cffafe",fontSize:10}}>{r.att}</div>
                        </td>
                      ))}
                      <td style={S.td(TL,true,"right")}>
                        <div style={{color:fy.gap<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(fy.gap)}</div>
                        <div style={{color:"#cffafe",fontSize:10}}>{fy.att}</div>
                      </td>
                    </tr>,
                    <tr key={cat+"gap2"} style={{background:bg,borderBottom:"1px dashed #aaa"}}>
                      <td style={S.td("#7B5EA7",true)}><span style={{color:"#fff"}}>Forecast Weighted GAP / Attain%</span></td>
                      {qRows.map((r,i)=>{
                        const wg=r.wfcst-r.quota,wa=r.quota?(r.wfcst/r.quota*100).toFixed(0)+"%":"0%";
                        return <td key={i} style={S.td("#7B5EA7",false,"right")}><div style={{color:wg<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(wg)}</div><div style={{color:"#e9d5ff",fontSize:10}}>{wa}</div></td>;
                      })}
                      <td style={S.td("#7B5EA7",true,"right")}>
                        <div style={{color:fy.wgap<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(fy.wgap)}</div>
                        <div style={{color:"#e9d5ff",fontSize:10}}>{fy.quota?(fy.wfcst/fy.quota*100).toFixed(0)+"%":"0%"}</div>
                      </td>
                    </tr>,
                    <tr key={cat+"gap3"} style={{background:bg,borderBottom:"2px solid #bbb"}}>
                      <td style={S.td("#b45309",true)}><span style={{color:"#fff"}}>Probability Weighted GAP / Attain%</span></td>
                      {qRows.map((r,i)=>{
                        const pg=r.pfcst-r.quota,pa=r.quota?(r.pfcst/r.quota*100).toFixed(0)+"%":"0%";
                        return <td key={i} style={S.td("#b45309",false,"right")}><div style={{color:pg<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(pg)}</div><div style={{color:"#fef3c7",fontSize:10}}>{pa}</div></td>;
                      })}
                      <td style={S.td("#b45309",true,"right")}>
                        <div style={{color:fy.pgap<0?"#ffcccc":"#fff",fontWeight:700}}>{fmt(fy.pgap)}</div>
                        <div style={{color:"#fef3c7",fontSize:10}}>{fy.quota?(fy.pfcst/fy.quota*100).toFixed(0)+"%":"0%"}</div>
                      </td>
                    </tr>,
                  ];
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}

// Detail Tab
function DetailTab({data,setData,region,product,weights}){
  const [filters,setFilters]=useState({owner:"All",seg:"All",stage:"All",fcstType:"All"});
  const [editing,setEditing]=useState(null);
  const [newRow,setNewRow]=useState(null);
  const W=k=>(weights[k]||0)/100;
  const uniq=(arr,key)=>["All",...new Set(arr.map(r=>r[key]).filter(Boolean))];
  const filtered=useMemo(()=>data.filter(r=>
    (filters.owner==="All"||r.owner===filters.owner)&&
    (filters.seg==="All"||r.seg===filters.seg)&&
    (filters.stage==="All"||r.stage===filters.stage)&&
    (filters.fcstType==="All"||r.fcstType===filters.fcstType)
  ),[data,filters]);
  const tARR=filtered.reduce((s,r)=>s+(r.netNewARR||0),0);
  const tFcst=filtered.reduce((s,r)=>s+(r.fcstARR||0),0);
  const tWt=filtered.reduce((s,r)=>{const a=Math.round((r.fcstARR||0)*W(r.fcstTypeOverride||"N/A"));return s+(r.weightedARR!=null?r.weightedARR:a);},0);
  const COLS=[
    {key:"account",         lbl:"Account Name"},
    {key:"opp",             lbl:"Pipeline Name"},
    {key:"owner",           lbl:"Account Owner"},
    {key:"seg",             lbl:"Segmentation",dd:GO},
    {key:"launchDate",      lbl:"Launch Date"},
    {key:"netNewARR",       lbl:"Q Net New ARR $",num:true,hb:"#FF6600",hc:"#fff"},
    {key:"fcstARR",         lbl:"Forecast Net New ARR$",num:true,hb:"#7030A0",hc:"#fff"},
    {key:"fcstType",        lbl:"Forecast Type",dd:FO},
    {key:"fcstTypeOverride",lbl:"Forecast Type (Override)",dd:FO},
    {key:"weightedARR",     lbl:"Forecast Weighted Net New ARR",num:true,hb:"#7B2D8B",hc:"#fff"},
    {key:"stage",           lbl:"Opp Stage",dd:SO},
    {key:"prob",            lbl:"Prob %",num:true},
  ];
  const cbg=(key,row)=>{
    if(key==="fcstARR")return row.fcstARR!==row.netNewARR?YW:"#DDEEFF";
    if(key==="weightedARR"){const a=Math.round((row.fcstARR||0)*W(row.fcstTypeOverride||"N/A"));return row.weightedARR!=null&&row.weightedARR!==a?"#f3e8ff":"#ede0ff";}
    if(key==="fcstTypeOverride"||key==="fcstType")return FC[row[key]]||"#fff";
    if(key==="stage")return SC[row.stage]||"#fff";
    return "#fff";
  };
  const edit=(idx,key,val)=>{
    const upd=[...data],r={...upd[idx]};
    if(key==="fcstARR"){const v=parseInt(val.toString().replace(/,/g,""))||0;r.fcstARR=v;if(r.weightedARR==null)r.weightedARR=Math.round(v*W(r.fcstTypeOverride||"N/A"));}
    else if(key==="weightedARR"){r.weightedARR=parseInt(val.toString().replace(/,/g,""))||0;}
    else if(key==="fcstTypeOverride"){r.fcstTypeOverride=val;if(r.weightedARR==null)r.weightedARR=Math.round((r.fcstARR||0)*W(val));}
    else if(key==="netNewARR"||key==="prob"){r[key]=parseInt(val.toString().replace(/,/g,""))||0;}
    else r[key]=val;
    upd[idx]=r;setData(upd);
  };
  const blank={region,product,account:"",opp:"",owner:"",seg:"KA",launchDate:"",netNewARR:0,fcstARR:0,fcstType:"Upside",fcstTypeOverride:"Upside",weightedARR:null,stage:"Technical Validation",prob:30};
  const sel=(key,opts)=><select value={filters[key]} onChange={e=>setFilters(f=>({...f,[key]:e.target.value}))} style={{border:"1px solid #bbb",borderRadius:3,padding:"2px 4px",fontSize:11}}>{opts.map(o=><option key={o}>{o}</option>)}</select>;
  return (
    <div style={{padding:10}}>
      <div style={{display:"flex",gap:10,marginBottom:8,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontWeight:700,fontSize:11,color:"#555"}}>FILTER:</span>
        <label style={{fontSize:11}}>Owner {sel("owner",uniq(data,"owner"))}</label>
        <label style={{fontSize:11}}>Seg {sel("seg",uniq(data,"seg"))}</label>
        <label style={{fontSize:11}}>Stage {sel("stage",uniq(data,"stage"))}</label>
        <label style={{fontSize:11}}>Forecast {sel("fcstType",uniq(data,"fcstType"))}</label>
        <button onClick={()=>setNewRow({...blank})} style={{marginLeft:"auto",background:"#4472C4",color:"#fff",border:"none",borderRadius:4,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600}}>+ Add Row</button>
        <button onClick={()=>cp(toTSV(THDR,filtered.map(toArr)),region+" "+(product==="TiDB Cloud"?"Cloud":"OP")+" FY27")} style={{background:"#217346",color:"#fff",border:"none",borderRadius:4,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:600}}>Copy to Sheets</button>
      </div>
      <div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse",minWidth:900}}>
          <thead>
            <tr>{COLS.map(c=><th key={c.key} style={S.th(c.hb||"#d8d8d8",c.hc||"#000")}>{c.key==="netNewARR"?fmt(tARR):c.key==="fcstARR"?fmt(tFcst):c.key==="weightedARR"?fmt(tWt):""}</th>)}<th style={S.th("#d8d8d8")}></th></tr>
            <tr>{COLS.map(c=><th key={c.key} style={S.th(c.hb||"#4472C4","#fff")}>{c.lbl}</th>)}<th style={S.th("#4472C4","#fff")}>Act</th></tr>
          </thead>
          <tbody>
            {filtered.map((row,i)=>{
              const ri=data.indexOf(row);
              return (
                <tr key={i} style={{background:i%2===0?"#fff":"#f9f9f9"}}>
                  {COLS.map(col=>{
                    const bg=cbg(col.key,row);
                    const isE=editing?.ri===ri&&editing?.key===col.key;
                    if(col.key==="fcstARR")return(
                      <td key={col.key} style={{...S.td(bg,false,"right"),minWidth:110,padding:"2px 4px"}}>
                        <input type="number" value={row.fcstARR} onChange={e=>edit(ri,"fcstARR",e.target.value)}
                          style={{width:"100%",border:row.fcstARR!==row.netNewARR?"1px solid #cca800":"1px solid #90b8e0",borderRadius:3,padding:"2px 4px",fontSize:11,textAlign:"right",background:bg,fontWeight:row.fcstARR!==row.netNewARR?700:400}}/>
                      </td>
                    );
                    if(col.key==="weightedARR"){
                      const a=Math.round((row.fcstARR||0)*W(row.fcstTypeOverride||"N/A"));
                      const val=row.weightedARR!=null?row.weightedARR:a;
                      const ovr=row.weightedARR!=null&&row.weightedARR!==a;
                      return(
                        <td key={col.key} style={{...S.td(bg,false,"right"),minWidth:100,padding:"2px 4px"}}>
                          <input type="number" value={val} onChange={e=>edit(ri,"weightedARR",e.target.value)}
                            title={ovr?"Manual override (auto="+fmt(a)+")":"Auto: "+(weights[row.fcstTypeOverride||"N/A"]||0)+"% of Forecast"}
                            style={{width:"100%",border:ovr?"1px solid #9333ea":"1px solid #c4b5fd",borderRadius:3,padding:"2px 4px",fontSize:11,textAlign:"right",background:bg,fontWeight:ovr?700:400}}/>
                        </td>
                      );
                    }
                    if(col.dd)return(
                      <td key={col.key} style={S.td(bg)}>
                        <select value={row[col.key]} onChange={e=>edit(ri,col.key,e.target.value)}
                          style={{width:"100%",border:"1px solid #cbd5e1",borderRadius:3,padding:"2px 3px",fontSize:11,background:bg,cursor:"pointer"}}>
                          {col.dd.map(o=><option key={o} value={o}>{o}</option>)}
                        </select>
                      </td>
                    );
                    const disp=col.key==="prob"?row[col.key]+"%":col.num?fmt(row[col.key]):row[col.key];
                    return(
                      <td key={col.key} style={S.td(bg,false,col.num?"right":"left")} onDoubleClick={()=>setEditing({ri,key:col.key})}>
                        {isE?<input autoFocus defaultValue={row[col.key]} onBlur={e=>{edit(ri,col.key,e.target.value);setEditing(null);}} style={{width:"100%",border:"1px solid #4472C4",fontSize:11,padding:2}}/>:disp}
                      </td>
                    );
                  })}
                  <td style={S.td()}><button onClick={()=>setData(data.filter((_,j)=>j!==ri))} style={{background:"#e74c3c",color:"#fff",border:"none",borderRadius:3,padding:"1px 6px",cursor:"pointer",fontSize:10}}>x</button></td>
                </tr>
              );
            })}
            {newRow&&(
              <tr style={{background:"#fffde7"}}>
                {COLS.map(col=>(
                  <td key={col.key} style={S.td("#fffde7",false,col.num?"right":"left")}>
                    {col.dd
                      ?<select value={newRow[col.key]||""} onChange={e=>setNewRow(r=>({...r,[col.key]:e.target.value}))} style={{width:"100%",border:"1px solid #f59e0b",fontSize:11,padding:2}}>{col.dd.map(o=><option key={o}>{o}</option>)}</select>
                      :<input value={newRow[col.key]||""} onChange={e=>setNewRow(r=>({...r,[col.key]:col.num?(parseInt(e.target.value)||0):e.target.value}))} style={{width:"100%",border:"1px solid #f59e0b",fontSize:11,padding:2}}/>
                    }
                  </td>
                ))}
                <td style={S.td()}>
                  <button onClick={()=>{setData([...data,newRow]);setNewRow(null);}} style={{background:"#22c55e",color:"#fff",border:"none",borderRadius:3,padding:"1px 5px",cursor:"pointer",fontSize:10,marginRight:2}}>ok</button>
                  <button onClick={()=>setNewRow(null)} style={{background:"#94a3b8",color:"#fff",border:"none",borderRadius:3,padding:"1px 5px",cursor:"pointer",fontSize:10}}>x</button>
                </td>
              </tr>
            )}
            {!filtered.length&&!newRow&&<tr><td colSpan={COLS.length+1} style={{padding:16,textAlign:"center",color:"#888"}}>No records match filters.</td></tr>}
          </tbody>
          <tfoot>
            <tr style={{background:"#f0f0f0"}}>
              <td colSpan={5} style={S.td("#f0f0f0",true)}>Total ({filtered.length} opps)</td>
              <td style={S.td("#FF6600",true,"right")}>{fmt(tARR)}</td>
              <td style={S.td("#7030A0",true,"right")}><span style={{color:"#fff"}}>{fmt(tFcst)}</span></td>
              <td style={S.td("#f0f0f0")}></td><td style={S.td("#f0f0f0")}></td>
              <td style={S.td("#7B2D8B",true,"right")}><span style={{color:"#fff"}}>{fmt(tWt)}</span></td>
              <td style={S.td("#f0f0f0")}></td><td style={S.td("#f0f0f0")}></td>
              <td style={S.td("#f0f0f0")}></td>
            </tr>
          </tfoot>
        </table>
      </div>
    </div>
  );
}

// Refresh Panel
function RefreshPanel({onRefresh}){
  const [msg,setMsg]=useState(""), [lastRefresh,setLastRefresh]=useState(null), [loading,setLoading]=useState(false);
  const FY27_S=qS("2026-04-01"), FY27_E=qE("2027-03-31");

  const parseRows=useCallback((rows)=>{
    if(!rows||rows.length<2){setMsg("No data found in file.");return;}
    const h=rows[0].map(s=>String(s||"").trim());
    const idx=k=>h.findIndex(x=>x.toLowerCase().includes(k.toLowerCase()));
    const iO=idx("Opportunity Name"),iOw=idx("Opportunity Owner"),iP=idx("Product Family");
    const iSt=idx("Stage"),iPb=idx("Probability"),iA=idx("Account Name");
    const iSg=idx("Segmentation"),iN=idx("Net New ARR Amount"),iL=idx("Launch Date");
    const iF=idx("Forecast Type"),iR=idx("Region");
    const parsed=rows.slice(1).reduce((acc,c)=>{
      if(c.length<5)return acc;
      const rawDate=c[iL];
      let d=null,launchDate="";
      if(rawDate instanceof Date){d=rawDate.getTime();launchDate=`${rawDate.getMonth()+1}/${rawDate.getDate()}/${rawDate.getFullYear()}`;}
      else if(typeof rawDate==="number"){const dt=new Date(Math.round((rawDate-25569)*86400000));d=dt.getTime();launchDate=`${dt.getMonth()+1}/${dt.getDate()}/${dt.getFullYear()}`;}
      else{const s=String(rawDate||"");d=pd(s);launchDate=s;}
      if(!d||d<FY27_S||d>FY27_E)return acc;
      const region=String(c[iR]||"").trim();
      if(region!=="NA"&&region!=="EMEA")return acc;
      const netNew=parseFloat(String(c[iN]||"0").replace(/[^0-9.-]/g,""))||0;
      const fcst=String(c[iF]||"Pipeline").trim();
      acc.push({region,product:String(c[iP]||"").trim(),account:String(c[iA]||""),opp:String(c[iO]||""),owner:String(c[iOw]||""),seg:String(c[iSg]||""),launchDate,netNewARR:netNew,fcstARR:netNew,fcstType:fcst,fcstTypeOverride:fcst,stage:String(c[iSt]||""),prob:parseInt(c[iPb])||0,weightedARR:null});
      return acc;
    },[]);
    if(!parsed.length){setMsg("No matching rows found — check the file has Region & Launch Date columns.");return;}
    onRefresh(parsed);setLastRefresh(new Date());setMsg(parsed.length+" opps loaded!");
  },[onRefresh,FY27_S,FY27_E]);

  const handleFile=useCallback(e=>{
    const file=e.target.files[0];
    if(!file)return;
    setLoading(true);setMsg("Reading file...");
    const reader=new FileReader();
    const isCSV=file.name.toLowerCase().endsWith(".csv");
    reader.onload=ev=>{
      try{
        if(isCSV){
          const rows=ev.target.result.trim().split("\n").map(l=>l.split(",").map(s=>s.trim().replace(/^"|"$/g,"")));
          parseRows(rows);
        } else {
          
          
          const wb=XLSX.read(new Uint8Array(ev.target.result),{type:"array",cellDates:true});
          const ws=wb.Sheets[wb.SheetNames[0]];
          parseRows(XLSX.utils.sheet_to_json(ws,{header:1,raw:false,dateNF:"m/d/yyyy"}));
        }
      } catch(err){setMsg("Error reading file: "+err.message);}
      setLoading(false);
      e.target.value="";
    };
    if(isCSV)reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
  },[parseRows]);

  return (
    <>
      
      <div style={{padding:"6px 14px",background:"#f0f7ff",borderBottom:"1px solid #bfdbfe",display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:11,fontWeight:700,color:"#1d4ed8"}}>Refresh from SFDC:</span>
        <a href="https://pingcap.lightning.force.com/lightning/r/Report/00ORC00000EKIC12AP/view?queryScope=userFolders" target="_blank" rel="noreferrer"
          style={{background:"#0070C0",color:"#fff",borderRadius:4,padding:"3px 10px",fontSize:11,fontWeight:600,textDecoration:"none",whiteSpace:"nowrap"}}>
          1. Open SFDC Report ↗
        </a>
        <span style={{fontSize:11,color:"#555"}}>→ Export → Details Only → CSV or Excel →</span>
        <label style={{background:"#217346",color:"#fff",borderRadius:4,padding:"3px 10px",fontSize:11,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
          2. Upload File (.csv / .xlsx)
          <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
        </label>
        {loading&&<span style={{fontSize:11,color:"#2563eb",fontWeight:600}}>⏳ Loading...</span>}
        {msg&&<span style={{fontSize:11,color:msg.includes("opps")?"#15803d":"#dc2626",fontWeight:600}}>{msg}</span>}
        {lastRefresh&&<span style={{fontSize:11,color:"#555"}}>Last refreshed: {lastRefresh.toLocaleString()}</span>}
      </div>
    </>
  );
}

// App
export default function App(){
  const [data,setData]=useState(INIT);
  const [tab,setTab]=useState("summary");
  const [baselines,setBL]=useState({NA:{fy26Cloud:QQ.NA.fy26Cloud,fy26Op:QQ.NA.fy26Op},EMEA:{fy26Cloud:QQ.EMEA.fy26Cloud,fy26Op:QQ.EMEA.fy26Op}});
  const [weights,setWeights]=useState({...WD});
  const get=(region,prod)=>data.filter(r=>r.region===region&&r.product===prod);
  const set=(region,prod)=>upd=>setData(prev=>[...prev.filter(r=>!(r.region===region&&r.product===prod)),...upd]);
  const TABS=[
    {id:"summary",   lbl:"ARR FCST Summary", color:"#fff"},
    {id:"quarterly", lbl:"Quarterly View",    color:"#f0e6ff"},
    {id:"na_cloud",  lbl:"NA Cloud FCST",     color:"#cce5ff"},
    {id:"na_op",     lbl:"NA OP FCST",        color:"#ffe5cc"},
    {id:"emea_cloud",lbl:"EMEA Cloud FCST",   color:"#ccf5e0"},
    {id:"emea_op",   lbl:"EMEA OP FCST",      color:"#fff0cc"},
  ];
  return (
    <div style={S.wrap}>
      <div style={{background:"#1a3a5c",color:"#fff",padding:"6px 14px",fontSize:13,fontWeight:700}}>
        PingCAP ARR Forecast FY27 Full Year - NA and EMEA - 4/1/2026 to 3/31/2027
      </div>
      <RefreshPanel onRefresh={setData}/>
      <div style={S.tabs}>{TABS.map(t=><div key={t.id} style={S.tab(tab===t.id,t.color)} onClick={()=>setTab(t.id)}>{t.lbl}</div>)}</div>
      {tab==="summary"   &&<SummaryTab   data={data} baselines={baselines} setBL={setBL} weights={weights} setWeights={setWeights}/>}
      {tab==="quarterly" &&<QuarterlyTab data={data} baselines={baselines} weights={weights}/>}
      {tab==="na_cloud"  &&<DetailTab    data={get("NA","TiDB Cloud")}                   setData={set("NA","TiDB Cloud")}                   region="NA"   product="TiDB Cloud"                   weights={weights}/>}
      {tab==="na_op"     &&<DetailTab    data={get("NA","TiDB Enterprise Subscription")} setData={set("NA","TiDB Enterprise Subscription")} region="NA"   product="TiDB Enterprise Subscription" weights={weights}/>}
      {tab==="emea_cloud"&&<DetailTab    data={get("EMEA","TiDB Cloud")}                 setData={set("EMEA","TiDB Cloud")}                 region="EMEA" product="TiDB Cloud"                   weights={weights}/>}
      {tab==="emea_op"   &&<DetailTab    data={get("EMEA","TiDB Enterprise Subscription")}setData={set("EMEA","TiDB Enterprise Subscription")}region="EMEA" product="TiDB Enterprise Subscription" weights={weights}/>}
    </div>
  );
}
