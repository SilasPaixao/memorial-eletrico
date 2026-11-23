import {useState,useMemo} from "react";
import crtImage from "./assets/crt.png";
import {ImageRun,Table,TableRow,TableCell,WidthType,AlignmentType,BorderStyle} from "docx";
import {Document,Packer,Paragraph,HeadingLevel,TextRun} from "docx";
import {calculate,type Load,type Unit} from "./calculations";
import { jsPDF } from "jspdf";
import { save } from "@tauri-apps/plugin-dialog";
import { writeFile } from "@tauri-apps/plugin-fs";
import { isTauri } from "@tauri-apps/api/core";
type EditableLoad = Load & {efficiency?:number;powerFactor?:number;demandFactor?:number};
function defaultParameters(l:Load):EditableLoad {
 const motor=l.category.toLowerCase().includes("motor")||l.unit==="CV"||l.unit==="HP";
 const lighting=l.category.toLowerCase().includes("ilumina");
 return {...l,efficiency:lighting?undefined:90,powerFactor:lighting?0.92:0.85,demandFactor:lighting?100:undefined};
}


const defaultSections:Record<string,string>={
 "4.1 Dados técnicos":"• Tensão de Isolação de Baixa Tensão .....................................750 V\n• Tensão de Operação Trifásica de Baixa Tensão ....................380 V",
 "4.2 Ramal de Entrada":"• Cabo de cobre isolado 750 V bitola 4x16(16) mm²;",
 "4.3 Proteção geral":"• Disjuntor tripolar de 63A 10 kA;",
 "4.4 Eletroduto":"• Eletroduto PVC 1 1/4\";",
 "4.5 Aterramento":"- 01 haste de aço cobreado;\n- Cabo de cobre 16 mm²;\n- Eletroduto de PVC rígido de 25 mm;\n- Conector ATR;",
 "4.6 Medição":"• Será instalado um conjunto de medição com caixa de medição indireta padrão Coelba conforme projeto em anexo.",
 "5. DEMANDA E DIMENSIONAMENTO":"A potência elétrica absorvida e a demanda final deverão ser confirmadas pelas placas dos motores, considerando rendimento, fator de potência, regime de funcionamento, simultaneidade e corrente de partida. A potência em CV representa potência mecânica e não deve ser adotada diretamente como potência elétrica de entrada.\nPara dimensionar ramal, disjuntor, condutores, medição e eventual sistema de partida, é necessário verificar correntes nominais e de partida, tensão de alimentação e operação simultânea. O dimensionamento deverá ser validado pelo responsável técnico e conforme padrão vigente da distribuidora.",
 "6. CARACTERÍSTICAS DO FORNECIMENTO":"Tipo de fornecimento pretendido: Trifásico.\nTensão de atendimento: a confirmar junto à distribuidora, conforme disponibilidade da rede local e tensão nominal dos motores.",
 "7. ATERRAMENTO E PROTEÇÕES":"O aterramento, as proteções contra sobrecorrente, curto-circuito e choque elétrico, e os dispositivos de seccionamento e proteção individual dos motores deverão ser especificados no projeto executivo conforme ABNT NBR 5410 e normas da distribuidora.",
 "8. RESPONSABILIDADE TÉCNICA":"A execução e o dimensionamento final deverão ser acompanhados de responsabilidade técnica emitida por profissional legalmente habilitado, conforme atribuições aplicáveis."
};
const sectionKeys=Object.keys(defaultSections);
function readSections():Record<string,string>{try{return {...defaultSections,...JSON.parse(localStorage.getItem("memorial-modelo-secoes-v1")||"{}")}}catch{return {...defaultSections}}}

type Project={consumer:string;cpf:string;address:string;supply:"Monofásica"|"Trifásica";voltage:string;professional:string;registration:string;loads:EditableLoad[]};
const defaults:Project={consumer:"",cpf:"",address:"",supply:"Trifásica",voltage:"380",professional:"",registration:"",loads:[
{id:"1",description:"Motor elétrico",category:"Motores",quantity:1,power:15,unit:"CV"},
{id:"2",description:"Motor elétrico",category:"Motores",quantity:4,power:1,unit:"CV"},
{id:"3",description:"Motor elétrico",category:"Motores",quantity:2,power:5,unit:"CV"},
{id:"4",description:"Motor elétrico",category:"Motores",quantity:4,power:2,unit:"CV"},
{id:"5",description:"Iluminação",category:"Iluminação",quantity:10,power:50,unit:"W"}].map(l=>defaultParameters(l as Load))};
function initial():Project {
 try {
  const raw=localStorage.getItem("memorial-eletrico-v1");
  if(!raw)return defaults;
  const previous=JSON.parse(raw);
  const loads:Array<EditableLoad>=Array.isArray(previous.loads)?previous.loads.map((l:EditableLoad)=>({...defaultParameters(l),...l})):defaults.loads;
  // Migrates old global estimates to individual rows only when there are no row-level values.
  if(previous.eff||previous.fp||previous.lightingFp){
   for(const l of loads){
    if(l.category.toLowerCase().includes("motor")){
     if(previous.eff && !previous.loads.find((x:EditableLoad)=>x.id===l.id)?.efficiency)l.efficiency=Number(previous.eff);
     if(previous.fp && !previous.loads.find((x:EditableLoad)=>x.id===l.id)?.powerFactor)l.powerFactor=Number(previous.fp);
    }else if(l.category.toLowerCase().includes("ilumina")&&previous.lightingFp&&!previous.loads.find((x:EditableLoad)=>x.id===l.id)?.powerFactor)l.powerFactor=Number(previous.lightingFp);
   }
  }
  return {...defaults,...previous,loads};
 }catch{return defaults}
}
function fmt(v:number|undefined){return v===undefined?"indeterminado":v.toLocaleString("pt-BR",{maximumFractionDigits:3})}
async function saveBytes(bytes:Uint8Array,filename:string,filterName:string,extension:string){
 if(isTauri()){
  const path=await save({defaultPath:filename,filters:[{name:filterName,extensions:[extension]}]});
  if(path)await writeFile(path,bytes);
 }else{
  const blob=new Blob([new Uint8Array(bytes)],{type:"application/octet-stream"});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");a.href=url;a.download=filename;document.body.appendChild(a);a.click();a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),30000);
 }
}
async function saveBlob(blob:Blob,name:string,filterName:string,extension:string){
 await saveBytes(new Uint8Array(await blob.arrayBuffer()),name,filterName,extension);
}
async function imageBytes():Promise<Uint8Array>{return new Uint8Array(await (await fetch(crtImage)).arrayBuffer())}
function groupLoads(loads:EditableLoad[],motor:boolean){return loads.filter(l=>motor?isMotor(l):isLighting(l))}
function otherLoads(loads:EditableLoad[]){return loads.filter(l=>!isMotor(l)&&!isLighting(l))}
function lightingDescription(l:EditableLoad){
 const description=l.description.trim();
 if(!description || description.toLowerCase()==="iluminação")return l.quantity===1?"ponto de iluminação":"pontos de iluminação";
 return description.toLowerCase();
}
function money(n:number){return n.toLocaleString("pt-BR",{maximumFractionDigits:2})}
function powerKw(l:EditableLoad){const factor:Record<Unit,number>={W:0.001,kW:1,CV:0.7355,HP:0.7457,VA:0.001,kVA:1};return l.quantity*l.power*factor[l.unit]}
function num(n:number,digits=2){return n.toLocaleString("pt-BR",{minimumFractionDigits:digits,maximumFractionDigits:digits})}
function param(n:number){return n.toLocaleString("pt-BR",{maximumFractionDigits:4})}
function isMotor(l:EditableLoad){return l.unit==="CV"||l.unit==="HP"||l.category.toLowerCase().includes("motor")}
function isLighting(l:EditableLoad){return l.category.toLowerCase().includes("ilumina")}
function calculationDetails(loads:EditableLoad[]):string[]{
 const lines:string[]=[];
 const motors=loads.filter(isMotor),lights=loads.filter(isLighting),others=loads.filter(l=>!isMotor(l)&&!isLighting(l));
 if(motors.length){
  const allCV=motors.every(l=>l.unit==="CV");const totalCV=allCV?motors.reduce((a,l)=>a+l.quantity*l.power,0):undefined;
  const mechanical=motors.reduce((a,l)=>a+powerKw(l),0);
  if(totalCV!==undefined)lines.push(`Conversão: ${param(totalCV)} × 0,7355 = ${num(mechanical)} kW`);
  else lines.push(`Potência mecânica nominal total dos motores: ${num(mechanical)} kW (conversões: 1 CV = 0,7355 kW; 1 HP = 0,7457 kW).`);
  const sameEta=motors.every(l=>l.efficiency===motors[0].efficiency),sameFp=motors.every(l=>l.powerFactor===motors[0].powerFactor);
  if(sameEta&&sameFp&&motors[0].efficiency&&motors[0].powerFactor){
   const eta=motors[0].efficiency/100,fp=motors[0].powerFactor;
   const active=mechanical/eta,apparent=active/fp;
   lines.push(`Rendimento estimado: ${param(motors[0].efficiency)}%`);
   lines.push(`Fator de potência: ${param(fp)}`);
   lines.push(`A potência total nominal dos motores corresponde a ${totalCV!==undefined?param(totalCV)+" CV, equivalente a aproximadamente ":"aproximadamente "}${num(mechanical)} kW de potência mecânica. Considerando rendimento estimado de ${param(motors[0].efficiency)}% e fator de potência de ${param(fp)}, obtém-se uma potência aparente estimada de ${num(apparent)} kVA.`);
   lines.push(`Potência ativa estimada: P = ${num(mechanical)} / ${param(eta)} = ${num(active)} kW.`);
   lines.push(`Potência aparente estimada: S = P / FP = ${num(active)} / ${param(fp)} = ${num(apparent)} kVA.`);
   lines.push(`Potência reativa estimada: Q = √(S² − P²) = ${num(Math.sqrt(Math.max(0,apparent**2-active**2)))} kvar.`);
  }else{
   lines.push('Parâmetros de rendimento e fator de potência considerados individualmente por equipamento:');
   for(const l of motors){
    const eta=(l.efficiency??0)/100,fp=l.powerFactor??0,pmech=powerKw(l);
    if(!(eta>0&&eta<=1&&fp>0&&fp<=1)){lines.push(`${l.description} (${l.quantity} × ${param(l.power)} ${l.unit}): rendimento ou fator de potência não informado.`);continue}
    const active=pmech/eta,apparent=active/fp;
    lines.push(`${l.description} (${l.quantity} × ${param(l.power)} ${l.unit}): rendimento estimado ${param(l.efficiency!)}%; fator de potência ${param(fp)}; P = ${num(pmech)} / ${param(eta)} = ${num(active)} kW; S = ${num(active)} / ${param(fp)} = ${num(apparent)} kVA; Q = ${num(Math.sqrt(Math.max(0,apparent**2-active**2)))} kvar.`);
   }
   const entries=motors.filter(l=>l.efficiency&&l.powerFactor).map(l=>{const active=powerKw(l)/(l.efficiency!/100);return {active,reactive:active*Math.sqrt(1/(l.powerFactor!**2)-1)}});
   if(entries.length===motors.length){const active=entries.reduce((a,x)=>a+x.active,0),reactive=entries.reduce((a,x)=>a+x.reactive,0);lines.push(`Totais dos motores: P = ${num(active)} kW; Q = ${num(reactive)} kvar; S = √(P² + Q²) = ${num(Math.hypot(active,reactive))} kVA.`)}
  }
 }
 for(const l of lights){
  const c=powerKw(l),fd=(l.demandFactor??100)/100,fp=l.powerFactor??0;
  lines.push(`Para ${l.quantity} ${lightingDescription(l)} com carga total de ${num(c)} kW.`);
  lines.push(`Fd = Fator de Demanda para Iluminação (parâmetro estimado/editável) = ${param(l.demandFactor??100)}%.`);
  lines.push(`FpI (fator de potência para iluminação) = ${param(fp)}.`);
  if(fd>=0&&fd<=1&&fp>0&&fp<=1){const demand=c*fd/fp;lines.push('Dm = (C × Fd) / FpI');lines.push(`Dm = (${num(c)} × ${param(fd)}) / ${param(fp)} = ${num(demand)} kVA.`);lines.push(`Potência reativa estimada na demanda: Q = √(Dm² − (C × Fd)²) = ${num(Math.sqrt(Math.max(0,demand**2-(c*fd)**2)))} kvar.`)}
  else lines.push('Dm: parâmetros de demanda ou fator de potência inválidos.');
 }
 for(const l of others){
  const base=powerKw(l),fp=l.powerFactor??0,eta=(l.efficiency??100)/100;
  if(l.unit==="VA"||l.unit==="kVA"){
   lines.push(`${l.description} (${l.quantity} × ${param(l.power)} ${l.unit}): potência aparente S = ${num(base)} kVA.`);
   if(fp>0&&fp<=1)lines.push(`Fator de potência: ${param(fp)}; P = S × FP = ${num(base*fp)} kW; Q = S × √(1 − FP²) = ${num(base*Math.sqrt(1-fp**2))} kvar.`)
  }else if(fp>0&&fp<=1&&eta>0&&eta<=1){const active=base/eta,apparent=active/fp;lines.push(`${l.description} (${l.quantity} × ${param(l.power)} ${l.unit}): potência ativa estimada P = ${num(base)} / ${param(eta)} = ${num(active)} kW; fator de potência ${param(fp)}; potência aparente S = P / FP = ${num(apparent)} kVA; potência reativa Q = √(S² − P²) = ${num(Math.sqrt(Math.max(0,apparent**2-active**2)))} kvar.`)}
  else lines.push(`${l.description}: rendimento ou fator de potência não informado para cálculo de potência aparente.`)
 }
 return lines;
}

export function App(){
 const [p,setP]=useState<Project>(initial);
 const [editing,setEditing]=useState<string[]>([]);
 const [notice,setNotice]=useState("");
 const [sections,setSections]=useState<Record<string,string>>(readSections);
 const [editingSections,setEditingSections]=useState(false);
 const changeSection=(key:string,value:string)=>setSections(old=>{const next={...old,[key]:value};localStorage.setItem("memorial-modelo-secoes-v1",JSON.stringify(next));return next});
 const restoreSections=()=>{const next={...defaultSections};setSections(next);localStorage.setItem("memorial-modelo-secoes-v1",JSON.stringify(next))};
 const set=(patch:Partial<Project>)=>setP(old=>{const next={...old,...patch};localStorage.setItem("memorial-eletrico-v1",JSON.stringify(next));return next});
 const patch=(id:string,change:Partial<EditableLoad>)=>set({loads:p.loads.map(x=>x.id===id?{...x,...change}:x)});
 const toggleParameters=(load:EditableLoad)=>{
  if(editing.includes(load.id)){setEditing(old=>old.filter(id=>id!==load.id));return}
  // Existing materials saved by older versions may not have FP or efficiency.
  const suggested=defaultParameters(load);
  const patchValues:Partial<EditableLoad>={};
  if(load.powerFactor===undefined)patchValues.powerFactor=suggested.powerFactor;
  if(load.efficiency===undefined&&!isLighting(load))patchValues.efficiency=90;
  if(isLighting(load)&&load.demandFactor===undefined)patchValues.demandFactor=100;
  if(Object.keys(patchValues).length)patch(load.id,patchValues);
  setEditing(old=>[...old,load.id]);
 };
 const calculations=useMemo(()=>p.loads.map(load=>{
  try{return {load,result:calculate(load,load.efficiency===undefined?undefined:load.efficiency/100,load.powerFactor)}}
  catch(e){return {load,error:String(e)}}
 }),[p]);
 const valid=calculations.every(x=>!x.error);
 async function exportDocx(){
  const motor=groupLoads(p.loads,true),lights=groupLoads(p.loads,false);
  const heading=(t:string)=>new Paragraph({children:[new TextRun({text:t,bold:true})],spacing:{before:220,after:100}});
  const para=(t:string)=>new Paragraph({text:t,alignment:AlignmentType.JUSTIFIED,spacing:{after:100},style:"Normal"});
  const tableFor=(items:EditableLoad[])=>new Table({width:{size:100,type:WidthType.PERCENTAGE},rows:[
   new TableRow({children:["Descrição","Quantidade","Potência unitária","Potência total"].map(t=>new TableCell({children:[new Paragraph({children:[new TextRun({text:t,bold:true})],alignment:AlignmentType.CENTER})]}))}),
   ...items.map((l,i)=>new TableRow({children:[l.description,String(l.quantity),`${money(l.power)} ${l.unit}`,`${money(l.power*l.quantity)} ${l.unit}`].map(t=>new TableCell({shading:i%2===0?{fill:"D9D9D9"}:undefined,children:[new Paragraph({text:t,alignment:AlignmentType.CENTER})]}))}))
  ],borders:{top:{style:BorderStyle.SINGLE,size:4,color:"777777"},bottom:{style:BorderStyle.SINGLE,size:4,color:"777777"},left:{style:BorderStyle.SINGLE,size:4,color:"777777"},right:{style:BorderStyle.SINGLE,size:4,color:"777777"},insideHorizontal:{style:BorderStyle.SINGLE,size:4,color:"777777"},insideVertical:{style:BorderStyle.SINGLE,size:4,color:"777777"}}});
  const children:any[]=[
   new Paragraph({text:"MEMORIAL DESCRITIVO",alignment:AlignmentType.CENTER,heading:HeadingLevel.HEADING_1}),
   new Paragraph({children:[new TextRun({text:`INSTALAÇÃO ELÉTRICA – ALIMENTAÇÃO ${p.supply.toUpperCase()}`,bold:true})],alignment:AlignmentType.CENTER,spacing:{after:420}}),
   heading("1. IDENTIFICAÇÃO DO CONSUMIDOR"),para(`Titular: ${p.consumer||"(não informado)"}`),para(`CPF: ${p.cpf||"(não informado)"}`),para(`Endereço: ${p.address||"(não informado)"}`),
   heading("2. OBJETIVO"),para(`Este memorial apresenta a relação de cargas prevista para atendimento do consumidor identificado, por meio de fornecimento ${p.supply.toLowerCase()}.`),
   heading("3. RELAÇÃO DE CARGA INSTALADA"),
   ...(motor.length?[heading("3.1 MOTORES"),tableFor(motor),heading("Memória de cálculo — Motores"),...calculationDetails(motor).map(para)]:[]),
   ...(lights.length?[heading("3.2 ILUMINAÇÃO"),tableFor(lights),heading("Memória de cálculo — Iluminação"),...calculationDetails(lights).map(para)]:[]),
   ...(otherLoads(p.loads).length?[heading("3.3 OUTROS EQUIPAMENTOS"),tableFor(otherLoads(p.loads)),...otherLoads(p.loads).flatMap(l=>[heading(`Memória de cálculo — ${l.description}`),...calculationDetails([l]).map(para)])]:[]),
   heading("4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA"),
   ...sectionKeys.flatMap(key=>key.startsWith("4.")?[heading(key),...sections[key].split("\n").map(para)]:[]),
   ...sectionKeys.flatMap(key=>key.startsWith("4.")?[]:[heading(key),...sections[key].split("\n").map(para)]),
   new Paragraph({text:"",spacing:{before:750,after:500}})

  ];
  const img=await imageBytes();
  // Keep the identification image at the end of the document, not on every page.
  children.push(new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:img,transformation:{width:214,height:75},type:"png"})]}));
  const doc=new Document({styles:{default:{document:{run:{font:"Times New Roman",size:24},paragraph:{spacing:{line:300}}}}},sections:[{properties:{page:{margin:{top:1150,bottom:1300,left:900,right:900}}},children}]});
  await saveBlob(await Packer.toBlob(doc),"memorial-eletrico.docx","Documento Word","docx");
 }
 async function makePdf():Promise<Uint8Array>{
  const pdf=new jsPDF({unit:"mm",format:"a4"});pdf.setFont("times","normal");pdf.setFontSize(12);
  const left=18,right=192,usable=right-left;let y=20;
  const next=(height:number)=>{if(y+height>265){pdf.addPage();y=22}};
  const line=(t:string,opts?:{bold?:boolean;center?:boolean;gap?:number})=>{pdf.setFont("times",opts?.bold?"bold":"normal");const lines=pdf.splitTextToSize(t,usable) as string[];next(lines.length*5+5);pdf.text(lines,opts?.center?105:left,y,{align:opts?.center?"center":"left"});y+=lines.length*5+(opts?.gap??3)};
  const h=(t:string)=>{y+=4;line(t,{bold:true,gap:3})};
  const table=(items:EditableLoad[])=>{const cols=[left,65,100,148,right];const widths=[47,35,48,44];const rows=[ ["Descrição","Quantidade","Potência unitária","Potência total"],...items.map(l=>[l.description,String(l.quantity),`${money(l.power)} ${l.unit}`,`${money(l.power*l.quantity)} ${l.unit}`])];next(rows.length*9+5);for(let ri=0;ri<rows.length;ri++){if(ri%2===1){pdf.setFillColor(217,217,217);pdf.rect(left,y-6,usable,9,"F")}pdf.setDrawColor(100);pdf.rect(left,y-6,usable,9);for(let j=1;j<cols.length-1;j++)pdf.line(cols[j],y-6,cols[j],y+3);pdf.setFont("times",ri===0?"bold":"normal");pdf.setFontSize(9);rows[ri].forEach((v,j)=>{const max=widths[j]-2;const clipped=pdf.getTextWidth(v)>max?v.slice(0,Math.max(2,Math.floor(v.length*max/pdf.getTextWidth(v))-1))+"…":v;pdf.text(clipped,cols[j]+widths[j]/2,y,{align:"center"})});pdf.setFontSize(12);y+=9}y+=5};
  line("MEMORIAL DESCRITIVO",{bold:true,center:true,gap:5});line(`INSTALAÇÃO ELÉTRICA – ALIMENTAÇÃO ${p.supply.toUpperCase()}`,{bold:true,center:true,gap:12});
  h("1. IDENTIFICAÇÃO DO CONSUMIDOR");line(`Titular: ${p.consumer||"(não informado)"}`);line(`CPF: ${p.cpf||"(não informado)"}`);line(`Endereço: ${p.address||"(não informado)"}`);
  h("2. OBJETIVO");line(`Este memorial apresenta a relação de cargas prevista para atendimento do consumidor identificado, por meio de fornecimento ${p.supply.toLowerCase()}.`);
  h("3. RELAÇÃO DE CARGA INSTALADA");
  const motors=groupLoads(p.loads,true),lighting=groupLoads(p.loads,false),others=otherLoads(p.loads);
  if(motors.length){h("3.1 MOTORES");table(motors);h("Memória de cálculo — Motores");for(const detail of calculationDetails(motors))line(detail)}
  if(lighting.length){h("3.2 ILUMINAÇÃO");table(lighting);h("Memória de cálculo — Iluminação");for(const detail of calculationDetails(lighting))line(detail)}
  if(others.length){h("3.3 OUTROS EQUIPAMENTOS");table(others);for(const material of others){h(`Memória de cálculo — ${material.description}`);for(const detail of calculationDetails([material]))line(detail)}}
  h("4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA");
  for(const key of sectionKeys){if(key.startsWith("4.")){h(key);for(const part of sections[key].split("\n"))line(part)}}
  for(const key of sectionKeys){if(!key.startsWith("4.")){h(key);for(const part of sections[key].split("\n"))line(part)}}
  // Leave room for signing above the professional identification on the final page.
  next(41);y+=15;
  const img=await imageBytes();const b64=btoa(String.fromCharCode(...img));
  next(25);pdf.addImage(`data:image/png;base64,${b64}`,"PNG",77,y,56,19.7);
  return new Uint8Array(pdf.output("arraybuffer"));
 }
 async function exportPdf(){try{const bytes=await makePdf();await saveBytes(bytes,"memorial-eletrico.pdf","Documento PDF","pdf");setNotice("PDF salvo.")}catch(e){setNotice("Erro ao salvar PDF: "+String(e))}}
 async function exportWord(){try{await exportDocx();setNotice("DOCX salvo.")}catch(e){setNotice("Erro ao salvar DOCX: "+String(e))}}
 async function exportBackup(){try{await saveBlob(new Blob([JSON.stringify({...p,templateSections:sections},null,2)],{type:"application/json"}),"memorial-backup.json","Backup JSON","json");setNotice("Backup salvo.")}catch(e){setNotice("Erro ao salvar backup: "+String(e))}}
 return <main><header><h1>⚡ Memorial Elétrico</h1><p>Aplicativo offline • exportação em PDF e DOCX</p></header>
 <section><h2>1. Consumidor e fornecimento</h2><div className="grid">
 <label>Consumidor<input value={p.consumer} onChange={e=>set({consumer:e.target.value})}/></label>
 <label>CPF<input value={p.cpf} onChange={e=>set({cpf:e.target.value})}/></label>
 <label>Endereço<input value={p.address} onChange={e=>set({address:e.target.value})}/></label>
 <label>Tipo<select value={p.supply} onChange={e=>set({supply:e.target.value as Project["supply"]})}><option>Monofásica</option><option>Trifásica</option></select></label>
 <label>Tensão informada (V)<input type="number" value={p.voltage} onChange={e=>set({voltage:e.target.value})}/></label>
 <label>Responsável técnico<input value={p.professional} onChange={e=>set({professional:e.target.value})}/></label>
 <label>Registro profissional<input value={p.registration} onChange={e=>set({registration:e.target.value})}/></label></div></section>
 <section><h2>2. Cargas</h2><p>Rendimento, fator de potência e fator de demanda são parâmetros estimados/editáveis. Clique no lápis da linha para editá-los.</p><div className="scroll"><table><thead><tr><th>Descrição</th><th>Categoria</th><th>Quantidade</th><th>Potência</th><th>Unidade</th><th>Rend. (%)</th><th>FP</th><th>Fd (%)</th><th>Ações</th></tr></thead><tbody>{p.loads.map(l=>{
 const unlocked=editing.includes(l.id);
 return <tr key={l.id}><td><input value={l.description} onChange={e=>patch(l.id,{description:e.target.value})}/></td><td><input value={l.category} onChange={e=>patch(l.id,{category:e.target.value})}/></td><td><input type="number" min="1" step="1" value={l.quantity} onChange={e=>patch(l.id,{quantity:Number(e.target.value)})}/></td><td><input type="number" min="0.01" value={l.power} onChange={e=>patch(l.id,{power:Number(e.target.value)})}/></td><td><select value={l.unit} onChange={e=>patch(l.id,{unit:e.target.value as Unit})}>{(["W","kW","CV","HP","VA","kVA"] as Unit[]).map(u=><option key={u}>{u}</option>)}</select></td>
 <td><input aria-label={"Rendimento de "+l.description} className={"parameter "+(!unlocked?"muted":"")} type="number" min="0.01" max="100" step="0.1" placeholder="—" disabled={!unlocked||isLighting(l)} value={l.efficiency??""} onChange={e=>patch(l.id,{efficiency:Number(e.target.value)})}/></td>
 <td><input aria-label={"FP de "+l.description} className={"parameter "+(!unlocked?"muted":"")} type="number" min="0.01" max="1" step="0.01" placeholder="—" disabled={!unlocked} value={l.powerFactor??""} onChange={e=>patch(l.id,{powerFactor:Number(e.target.value)})}/></td>
 <td><input aria-label={"Fator de demanda de "+l.description} className={"parameter "+(!unlocked?"muted":"")} type="number" min="0" max="100" step="1" placeholder="—" disabled={!unlocked||!isLighting(l)} value={l.demandFactor??""} onChange={e=>patch(l.id,{demandFactor:Number(e.target.value)})}/></td>
 <td><div className="actions"><button title={unlocked?"Concluir edição":"Editar parâmetros"} onClick={()=>toggleParameters(l)}>{unlocked?"✓":"✎"}</button><button onClick={()=>set({loads:p.loads.filter(x=>x.id!==l.id)})}>Remover</button></div></td></tr>})}</tbody></table></div>
 <button onClick={()=>set({loads:[...p.loads,defaultParameters({id:crypto.randomUUID(),description:"Novo equipamento",category:"Outros",quantity:1,power:1,unit:"W"})]})}>+ Adicionar equipamento</button></section>
 <section><h2>3. Textos do modelo (seções 4 a 8)</h2>
 <p>Os textos abaixo são os valores iniciais do modelo.pdf. Edite e salve para reutilizá-los em novos documentos. Confira os dados técnicos de cada instalação antes de exportar.</p>
 <div className="actions"><button onClick={()=>setEditingSections(!editingSections)}>{editingSections?"Concluir edição":"✎ Editar textos do modelo"}</button><button onClick={restoreSections}>Restaurar textos originais</button></div>
 {sectionKeys.map(key=><label key={key} className="section-editor"><strong>{key}</strong><textarea rows={Math.max(2,sections[key].split("\n").length+2)} disabled={!editingSections} value={sections[key]} onChange={e=>changeSection(key,e.target.value)}/></label>)}
 <h3>Identificação profissional</h3><div className="crt-preview"><img src={crtImage} alt="Identificação CRT/CFT do profissional" /></div>
 <p>A imagem será inserida no fim do documento, abaixo do espaço reservado para assinatura.</p>
 </section>
 <section><h2>4. Resultados indicativos</h2>{calculations.map(({load,result,error})=><div className="result" key={load.id}><b>{load.description}</b>: {error||`Mecânica ${fmt(result?.mechanicalKw)} kW | Ativa ${fmt(result?.activeKw)} kW | Aparente ${fmt(result?.apparentKva)} kVA`}<small>{result?.warnings.join(" ")}</small></div>)}
 {!valid&&<p className="warning">Corrija os dados inválidos antes de utilizar os resultados.</p>}
 <p className="hint">Os dados técnicos são editáveis e não são dimensionados automaticamente pelo programa.</p>
 <div className="actions"><button onClick={exportWord}>Exportar DOCX</button><button onClick={exportPdf}>Salvar PDF</button><button onClick={exportBackup}>Backup JSON</button><label className="import">Importar backup JSON<input type="file" accept=".json,application/json" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.loads))throw Error("Backup inválido");set(data);if(data.templateSections){const merged={...defaultSections,...data.templateSections};setSections(merged);localStorage.setItem("memorial-modelo-secoes-v1",JSON.stringify(merged))}}catch{alert("Arquivo inválido")}}}/></label></div>
 {notice&&<p role="status">{notice}</p>}
 </section></main>
}
