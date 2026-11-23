import {useState,useMemo} from "react";
import {Document,Packer,Paragraph,HeadingLevel,TextRun} from "docx";
import {calculate,type Load,type Unit} from "./calculations";
import { jsPDF } from "jspdf";
import { save } from "@tauri-apps/plugin-dialog";
import { writeFile } from "@tauri-apps/plugin-fs";
import { isTauri } from "@tauri-apps/api/core";
type EditableLoad = Load & {efficiency?:number;powerFactor?:number};
function defaultParameters(l:Load):EditableLoad {
 const motor=l.category.toLowerCase().includes("motor")||l.unit==="CV"||l.unit==="HP";
 const lighting=l.category.toLowerCase().includes("ilumina");
 return {...l,efficiency:motor?90:undefined,powerFactor:motor?0.85:lighting?0.92:undefined};
}

type Project={consumer:string;cpf:string;address:string;supply:"Monofásica"|"Trifásica";voltage:string;professional:string;registration:string;loads:EditableLoad[]};
const defaults:Project={consumer:"",cpf:"",address:"",supply:"Trifásica",voltage:"380",professional:"",registration:"",loads:[
{id:"1",description:"Motor elétrico",category:"Motores",quantity:1,power:15,unit:"CV"},
{id:"2",description:"Motor elétrico",category:"Motores",quantity:4,power:1,unit:"CV"},
{id:"3",description:"Motor elétrico",category:"Motores",quantity:2,power:5,unit:"CV"},
{id:"4",description:"Motor elétrico",category:"Motores",quantity:4,power:2,unit:"CV"},
{id:"5",description:"Iluminação",category:"Iluminação",quantity:10,power:50,unit:"W"}].map(l => defaultParameters(l as Load))};
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
export function App(){
 const [p,setP]=useState<Project>(initial);
 const [editing,setEditing]=useState<string[]>([]);
 const [notice,setNotice]=useState("");
 const set=(patch:Partial<Project>)=>setP(old=>{const next={...old,...patch};localStorage.setItem("memorial-eletrico-v1",JSON.stringify(next));return next});
 const patch=(id:string,change:Partial<EditableLoad>)=>set({loads:p.loads.map(x=>x.id===id?{...x,...change}:x)});
 const calculations=useMemo(()=>p.loads.map(load=>{
  try{return {load,result:calculate(load,load.efficiency===undefined?undefined:load.efficiency/100,load.powerFactor)}}
  catch(e){return {load,error:String(e)}}
 }),[p]);
 const valid=calculations.every(x=>!x.error);
 const report=[
 "RASCUNHO — NÃO VALIDADO PELO RESPONSÁVEL TÉCNICO",
 "MEMORIAL DESCRITIVO — INSTALAÇÃO ELÉTRICA",
 "1. IDENTIFICAÇÃO DO CONSUMIDOR",`Titular: ${p.consumer||"(não informado)"}`,`CPF: ${p.cpf||"(não informado)"}`,`Endereço: ${p.address||"(não informado)"}`,
 "2. OBJETIVO",`Relação preliminar de cargas para fornecimento ${p.supply.toLowerCase()}. Os parâmetros de cálculo estimados estão indicados junto às cargas.`,
 "3. RELAÇÃO DE CARGA INSTALADA",...calculations.map(({load,result,error})=>`${load.description}: ${load.quantity} × ${load.power} ${load.unit}; rendimento estimado ${load.efficiency===undefined?"não aplicável":load.efficiency+"%"}; FP estimado ${load.powerFactor===undefined?"não informado":load.powerFactor}; potência mecânica ${fmt(result?.mechanicalKw)} kW; ativa ${fmt(result?.activeKw)} kW; aparente ${fmt(result?.apparentKva)} kVA. ${error||result?.warnings.join(" ")||""}`),
 "4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA",`Fornecimento pretendido: ${p.supply}; tensão informada: ${p.voltage||"não informada"} V. Condutores, disjuntores, aterramento, medição e padrão de entrada: não informados nesta versão.`,
 "5. DEMANDA E DIMENSIONAMENTO","Demanda final, simultaneidade, correntes de partida e dimensionamento normativo ainda não determinados. Potência em CV é potência mecânica.",
 "6. CARACTERÍSTICAS DO FORNECIMENTO",`Tipo: ${p.supply}. Tensão e disponibilidade devem ser confirmadas junto à distribuidora.`,
 "7. ATERRAMENTO E PROTEÇÕES","Informações de aterramento e proteções não cadastradas nesta versão.",
 "8. RESPONSABILIDADE TÉCNICA",`Profissional: ${p.professional||"não informado"}. Registro: ${p.registration||"não informado"}. Documento em rascunho.`,
 "MEMORIAL DE CÁLCULO — PRELIMINAR","Os parâmetros estimados são informados individualmente na relação de cargas e podem ser editados.",
 "CV → kW mecânicos: CV × 0,7355. kW elétricos = kW mecânicos / rendimento. kVA = kW elétricos / FP.",
 "Observação: valores calculados com parâmetros estimados editáveis. Não foram calculados fatores de demanda nem correntes de partida."
 ];
 async function exportDocx(){
  const paragraphs=report.map(line=>new Paragraph({children:[new TextRun(line)],heading:/^(MEMORIAL |[1-8]\. )/.test(line)?HeadingLevel.HEADING_2:undefined,spacing:{after:140}}));
  const doc=new Document({sections:[{children:paragraphs}]});
  await saveBlob(await Packer.toBlob(doc),"memorial-eletrico-RASCUNHO.docx","Documento Word","docx");
 }
 async function exportPdf(){
  try{
   const pdf=new jsPDF({unit:"mm",format:"a4"});
   const margin=18,width=174;let y=24;
   pdf.setTextColor(190,190,190);pdf.setFontSize(43);pdf.text("RASCUNHO",105,145,{align:"center",angle:35});
   pdf.setTextColor(30,42,58);pdf.setFontSize(10);
   for(const line of report){
    const wrapped=pdf.splitTextToSize(line,width) as string[];
    if(y+wrapped.length*5>277){pdf.addPage();y=24;pdf.setTextColor(190,190,190);pdf.setFontSize(43);pdf.text("RASCUNHO",105,145,{align:"center",angle:35});pdf.setTextColor(30,42,58);pdf.setFontSize(10)}
    pdf.text(wrapped,margin,y);y+=wrapped.length*5+3;
   }
   await saveBytes(new Uint8Array(pdf.output("arraybuffer")),"memorial-eletrico-RASCUNHO.pdf","Documento PDF","pdf");
   setNotice("PDF salvo.");
  }catch(e){setNotice("Erro ao salvar PDF: "+String(e))}
 }
 async function exportWord(){try{await exportDocx();setNotice("DOCX salvo.")}catch(e){setNotice("Erro ao salvar DOCX: "+String(e))}}
 async function exportBackup(){try{await saveBlob(new Blob([JSON.stringify(p,null,2)],{type:"application/json"}),"memorial-backup.json","Backup JSON","json");setNotice("Backup salvo.")}catch(e){setNotice("Erro ao salvar backup: "+String(e))}}
 return <main><header><h1>⚡ Memorial Elétrico</h1><p>Protótipo offline • documentos sempre em RASCUNHO • sem dimensionamento normativo automático</p></header>
 <section><h2>1. Consumidor e fornecimento</h2><div className="grid">
 <label>Consumidor<input value={p.consumer} onChange={e=>set({consumer:e.target.value})}/></label>
 <label>CPF<input value={p.cpf} onChange={e=>set({cpf:e.target.value})}/></label>
 <label>Endereço<input value={p.address} onChange={e=>set({address:e.target.value})}/></label>
 <label>Tipo<select value={p.supply} onChange={e=>set({supply:e.target.value as Project["supply"]})}><option>Monofásica</option><option>Trifásica</option></select></label>
 <label>Tensão informada (V)<input type="number" value={p.voltage} onChange={e=>set({voltage:e.target.value})}/></label>
 <label>Responsável técnico<input value={p.professional} onChange={e=>set({professional:e.target.value})}/></label>
 <label>Registro profissional<input value={p.registration} onChange={e=>set({registration:e.target.value})}/></label></div></section>
 <section><h2>2. Cargas</h2><p>Rendimento e FP são estimativas iniciais. Clique no lápis da linha para editá-los.</p><div className="scroll"><table><thead><tr><th>Descrição</th><th>Categoria</th><th>Quantidade</th><th>Potência</th><th>Unidade</th><th>Rend. (%)</th><th>FP</th><th>Ações</th></tr></thead><tbody>{p.loads.map(l=>{
 const unlocked=editing.includes(l.id);
 return <tr key={l.id}><td><input value={l.description} onChange={e=>patch(l.id,{description:e.target.value})}/></td><td><input value={l.category} onChange={e=>patch(l.id,{category:e.target.value})}/></td><td><input type="number" min="1" step="1" value={l.quantity} onChange={e=>patch(l.id,{quantity:Number(e.target.value)})}/></td><td><input type="number" min="0.01" value={l.power} onChange={e=>patch(l.id,{power:Number(e.target.value)})}/></td><td><select value={l.unit} onChange={e=>patch(l.id,{unit:e.target.value as Unit})}>{(["W","kW","CV","HP","VA","kVA"] as Unit[]).map(u=><option key={u}>{u}</option>)}</select></td>
 <td><input aria-label={"Rendimento de "+l.description} className={"parameter "+(!unlocked?"muted":"")} type="number" min="0.01" max="100" step="0.1" placeholder="—" disabled={!unlocked||l.efficiency===undefined} value={l.efficiency??""} onChange={e=>patch(l.id,{efficiency:Number(e.target.value)})}/></td>
 <td><input aria-label={"FP de "+l.description} className={"parameter "+(!unlocked?"muted":"")} type="number" min="0.01" max="1" step="0.01" placeholder="—" disabled={!unlocked||l.powerFactor===undefined} value={l.powerFactor??""} onChange={e=>patch(l.id,{powerFactor:Number(e.target.value)})}/></td>
 <td><div className="actions"><button title={unlocked?"Concluir edição":"Editar parâmetros"} onClick={()=>setEditing(old=>unlocked?old.filter(id=>id!==l.id):[...old,l.id])}>{unlocked?"✓":"✎"}</button><button onClick={()=>set({loads:p.loads.filter(x=>x.id!==l.id)})}>Remover</button></div></td></tr>})}</tbody></table></div>
 <button onClick={()=>set({loads:[...p.loads,defaultParameters({id:crypto.randomUUID(),description:"Novo equipamento",category:"Outros",quantity:1,power:1,unit:"W"})]})}>+ Adicionar equipamento</button></section>
 <section><h2>4. Resultados indicativos</h2>{calculations.map(({load,result,error})=><div className="result" key={load.id}><b>{load.description}</b>: {error||`Mecânica ${fmt(result?.mechanicalKw)} kW | Ativa ${fmt(result?.activeKw)} kW | Aparente ${fmt(result?.apparentKva)} kVA`}<small>{result?.warnings.join(" ")}</small></div>)}
 {!valid&&<p className="warning">Corrija os dados inválidos antes de utilizar os resultados.</p>}
 <p className="warning">RASCUNHO. O armazenamento atual é local no navegador/WebView (localStorage), não SQLite. Não há aprovação técnica, assinatura, cálculo de demanda normativa ou dimensionamento de entrada.</p>
 <div className="actions"><button onClick={exportWord}>Exportar DOCX (rascunho)</button><button onClick={exportPdf}>Salvar PDF (rascunho)</button><button onClick={exportBackup}>Backup JSON</button><label className="import">Importar backup JSON<input type="file" accept=".json,application/json" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.loads))throw Error("Backup inválido");set(data)}catch{alert("Arquivo inválido")}}}/></label></div>
 {notice&&<p role="status">{notice}</p>}
 </section></main>
}
