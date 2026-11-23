import {useState,useMemo} from "react";
import {Document,Packer,Paragraph,HeadingLevel,TextRun} from "docx";
import {calculate,type Load,type Unit} from "./calculations";
type Project={consumer:string;cpf:string;address:string;supply:"Monofásica"|"Trifásica";voltage:string;professional:string;registration:string;eff:string;fp:string;lightingFp:string;loads:Load[]};
const defaults:Project={consumer:"",cpf:"",address:"",supply:"Trifásica",voltage:"380",professional:"",registration:"",eff:"",fp:"",lightingFp:"",loads:[
{id:"1",description:"Motor elétrico",category:"Motores",quantity:1,power:15,unit:"CV"},
{id:"2",description:"Motor elétrico",category:"Motores",quantity:4,power:1,unit:"CV"},
{id:"3",description:"Motor elétrico",category:"Motores",quantity:2,power:5,unit:"CV"},
{id:"4",description:"Motor elétrico",category:"Motores",quantity:4,power:2,unit:"CV"},
{id:"5",description:"Iluminação",category:"Iluminação",quantity:10,power:50,unit:"W"}]};
function initial():Project{try{const raw=localStorage.getItem("memorial-eletrico-v1");return raw?{...defaults,...JSON.parse(raw)}:defaults}catch{return defaults}}
function fmt(v:number|undefined){return v===undefined?"indeterminado":v.toLocaleString("pt-BR",{maximumFractionDigits:3})}
function saveBlob(blob:Blob,name:string){const a=document.createElement("a");const url=URL.createObjectURL(blob);a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),2000)}
export function App(){
 const [p,setP]=useState<Project>(initial);
 const set=(patch:Partial<Project>)=>setP(old=>{const next={...old,...patch};localStorage.setItem("memorial-eletrico-v1",JSON.stringify(next));return next});
 const patch=(id:string,change:Partial<Load>)=>set({loads:p.loads.map(x=>x.id===id?{...x,...change}:x)});
 const calculations=useMemo(()=>p.loads.map(load=>{try{const isMotor=load.category.toLowerCase().includes("motor");const isLight=load.category.toLowerCase().includes("ilumina");return {load,result:calculate(load,isMotor&&p.eff?Number(p.eff)/100:undefined,isMotor&&p.fp?Number(p.fp):isLight&&p.lightingFp?Number(p.lightingFp):undefined)}}catch(e){return {load,error:String(e)}}}),[p]);
 const valid=calculations.every(x=>!x.error);
 const report=[
 "RASCUNHO — NÃO VALIDADO PELO RESPONSÁVEL TÉCNICO",
 "MEMORIAL DESCRITIVO — INSTALAÇÃO ELÉTRICA",
 "1. IDENTIFICAÇÃO DO CONSUMIDOR",`Titular: ${p.consumer||"(não informado)"}`,`CPF: ${p.cpf||"(não informado)"}`,`Endereço: ${p.address||"(não informado)"}`,
 "2. OBJETIVO",`Relação preliminar de cargas para fornecimento ${p.supply.toLowerCase()}. A ligação e as proteções dependem de verificação normativa e aprovação do responsável técnico.`,
 "3. RELAÇÃO DE CARGA INSTALADA",...calculations.map(({load,result,error})=>`${load.description}: ${load.quantity} × ${load.power} ${load.unit}; potência mecânica ${fmt(result?.mechanicalKw)} kW; ativa ${fmt(result?.activeKw)} kW; aparente ${fmt(result?.apparentKva)} kVA. ${error||result?.warnings.join(" ")||""}`),
 "4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA",`Fornecimento pretendido: ${p.supply}; tensão informada: ${p.voltage||"não informada"} V. Condutores, disjuntores, aterramento, medição e padrão de entrada: PENDENTES DE DIMENSIONAMENTO.`,
 "5. DEMANDA E DIMENSIONAMENTO","Demanda final, simultaneidade, correntes de partida e dimensionamento normativo ainda não determinados. Potência em CV é potência mecânica.",
 "6. CARACTERÍSTICAS DO FORNECIMENTO",`Tipo: ${p.supply}. Tensão e disponibilidade devem ser confirmadas junto à distribuidora.`,
 "7. ATERRAMENTO E PROTEÇÕES","Definição sujeita a projeto técnico, ABNT NBR 5410 e regras vigentes da distribuidora.",
 "8. RESPONSABILIDADE TÉCNICA",`Profissional: ${p.professional||"não informado"}. Registro: ${p.registration||"não informado"}. Sem assinatura ou aprovação nesta versão.`,
 "MEMORIAL DE CÁLCULO — PRELIMINAR",`Premissas de motores: rendimento ${p.eff||"não informado"}%; FP ${p.fp||"não informado"}. Iluminação: FP ${p.lightingFp||"não informado"}.`,
 "CV → kW mecânicos: CV × 0,7355. kW elétricos = kW mecânicos / rendimento. kVA = kW elétricos / FP.",
 "ATENÇÃO: os valores são indicativos; não incluem fatores de demanda normativos, partida de motores ou validação de proteção. Nenhuma aprovação técnica foi registrada."
 ];
 async function exportDocx(){
  const paragraphs=report.map(line=>new Paragraph({children:[new TextRun(line)],heading:/^(MEMORIAL |[1-8]\. )/.test(line)?HeadingLevel.HEADING_2:undefined,spacing:{after:140}}));
  const doc=new Document({sections:[{children:paragraphs}]});
  saveBlob(await Packer.toBlob(doc),"memorial-eletrico-RASCUNHO.docx");
 }
 function printPdf(){
  const w=window.open("","_blank");if(!w){alert("Permita janelas pop-up para imprimir/salvar PDF.");return}
  const escape=(s:string)=>s.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]!));
  w.document.write(`<html><head><meta charset="utf-8"><title>Memorial RASCUNHO</title><style>body{font:12pt Arial;margin:18mm;line-height:1.5}p{margin:0 0 10px}body:before{content:"RASCUNHO";position:fixed;top:42%;left:12%;font-size:80pt;color:#eee;transform:rotate(-30deg);z-index:-1}@page{size:A4;margin:18mm}</style></head><body>${report.map(t=>`<p>${escape(t)}</p>`).join("")}</body></html>`);
  w.document.close();w.focus();setTimeout(()=>w.print(),300);
 }
 return <main><header><h1>⚡ Memorial Elétrico</h1><p>Protótipo offline • documentos sempre em RASCUNHO • sem dimensionamento normativo automático</p></header>
 <section><h2>1. Consumidor e fornecimento</h2><div className="grid">
 <label>Consumidor<input value={p.consumer} onChange={e=>set({consumer:e.target.value})}/></label>
 <label>CPF<input value={p.cpf} onChange={e=>set({cpf:e.target.value})}/></label>
 <label>Endereço<input value={p.address} onChange={e=>set({address:e.target.value})}/></label>
 <label>Tipo<select value={p.supply} onChange={e=>set({supply:e.target.value as Project["supply"]})}><option>Monofásica</option><option>Trifásica</option></select></label>
 <label>Tensão informada (V)<input type="number" value={p.voltage} onChange={e=>set({voltage:e.target.value})}/></label>
 <label>Responsável técnico<input value={p.professional} onChange={e=>set({professional:e.target.value})}/></label>
 <label>Registro profissional<input value={p.registration} onChange={e=>set({registration:e.target.value})}/></label></div></section>
 <section><h2>2. Cargas</h2><div className="scroll"><table><thead><tr><th>Descrição</th><th>Categoria</th><th>Quantidade</th><th>Potência</th><th>Unidade</th><th></th></tr></thead><tbody>{p.loads.map(l=><tr key={l.id}><td><input value={l.description} onChange={e=>patch(l.id,{description:e.target.value})}/></td><td><input value={l.category} onChange={e=>patch(l.id,{category:e.target.value})}/></td><td><input type="number" min="1" step="1" value={l.quantity} onChange={e=>patch(l.id,{quantity:Number(e.target.value)})}/></td><td><input type="number" min="0.01" value={l.power} onChange={e=>patch(l.id,{power:Number(e.target.value)})}/></td><td><select value={l.unit} onChange={e=>patch(l.id,{unit:e.target.value as Unit})}>{(["W","kW","CV","HP","VA","kVA"] as Unit[]).map(u=><option key={u}>{u}</option>)}</select></td><td><button onClick={()=>set({loads:p.loads.filter(x=>x.id!==l.id)})}>Remover</button></td></tr>)}</tbody></table></div><button onClick={()=>set({loads:[...p.loads,{id:crypto.randomUUID(),description:"Novo equipamento",category:"Outros",quantity:1,power:1,unit:"W"}]})}>+ Adicionar equipamento</button></section>
 <section><h2>3. Premissas de cálculo</h2><p>Não são dados de placa. Informe apenas premissas justificadas pelo profissional.</p><div className="grid"><label>Rendimento motores (%)<input type="number" min="0" max="100" value={p.eff} onChange={e=>set({eff:e.target.value})}/></label><label>FP motores<input type="number" min="0" max="1" step=".01" value={p.fp} onChange={e=>set({fp:e.target.value})}/></label><label>FP iluminação<input type="number" min="0" max="1" step=".01" value={p.lightingFp} onChange={e=>set({lightingFp:e.target.value})}/></label></div></section>
 <section><h2>4. Resultados indicativos</h2>{calculations.map(({load,result,error})=><div className="result" key={load.id}><b>{load.description}</b>: {error||`Mecânica ${fmt(result?.mechanicalKw)} kW | Ativa ${fmt(result?.activeKw)} kW | Aparente ${fmt(result?.apparentKva)} kVA`}<small>{result?.warnings.join(" ")}</small></div>)}
 {!valid&&<p className="warning">Corrija os dados inválidos antes de utilizar os resultados.</p>}
 <p className="warning">RASCUNHO. O armazenamento atual é local no navegador/WebView (localStorage), não SQLite. Não há aprovação técnica, assinatura, cálculo de demanda normativa ou dimensionamento de entrada.</p>
 <div className="actions"><button onClick={exportDocx}>Exportar DOCX (rascunho)</button><button onClick={printPdf}>Imprimir / salvar PDF (rascunho)</button><button onClick={()=>saveBlob(new Blob([JSON.stringify(p,null,2)],{type:"application/json"}),"memorial-backup.json")}>Backup JSON</button><label className="import">Importar backup JSON<input type="file" accept=".json,application/json" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.loads))throw Error("Backup inválido");set(data)}catch{alert("Arquivo inválido")}}}/></label></div>
 </section></main>
}
