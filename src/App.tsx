import {useState,useMemo} from "react";
import crtImage from "./assets/crt.png";
import {ImageRun,Table,TableRow,TableCell,WidthType,AlignmentType,BorderStyle,Footer} from "docx";
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
type SignatureRecord={name:string;date:string;size:number;documentHash:string;status:"importado-sem-verificacao"};
function bytesFromDataUrl(dataUrl:string):Uint8Array{
 const raw=atob(dataUrl.split(",")[1]);return Uint8Array.from(raw,c=>c.charCodeAt(0));
}
async function imageBytes():Promise<Uint8Array>{return new Uint8Array(await (await fetch(crtImage)).arrayBuffer())}
async function digest(bytes:Uint8Array):Promise<string>{const h=await crypto.subtle.digest("SHA-256",new Uint8Array(bytes));return [...new Uint8Array(h)].map(x=>x.toString(16).padStart(2,"0")).join("")}
function groupLoads(loads:EditableLoad[],motor:boolean){return loads.filter(l=>motor?(l.category.toLowerCase().includes("motor")||l.unit==="CV"||l.unit==="HP"):l.category.toLowerCase().includes("ilumina"))}
function money(n:number){return n.toLocaleString("pt-BR",{maximumFractionDigits:2})}
export function App(){
 const [p,setP]=useState<Project>(initial);
 const [editing,setEditing]=useState<string[]>([]);
 const [notice,setNotice]=useState("");
 const [signatureRecords,setSignatureRecords]=useState<SignatureRecord[]>(()=>{try{return JSON.parse(localStorage.getItem("memorial-assinaturas-v1")||"[]")}catch{return []}});
 const [lastPdfHash,setLastPdfHash]=useState("");
 const [showCrt,setShowCrt]=useState(true);
 const [includeSignatureBlock,setIncludeSignatureBlock]=useState(true);
 const [importedPdf,setImportedPdf]=useState("");
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
   heading("3. RELAÇÃO DE CARGA INSTALADA"),heading("3.1 MOTORES"),tableFor(motor),
   para(`Potência mecânica total: ${money(motor.reduce((a,l)=>a+l.quantity*l.power*(l.unit==="CV"?0.7355:l.unit==="HP"?0.7457:l.unit==="kW"?1:l.unit==="W"?0.001:0),0))} kW.`),
   ...motor.map(l=>para(`${l.description} (${l.quantity} × ${money(l.power)} ${l.unit}): rendimento estimado ${l.efficiency??"—"}%; FP estimado ${l.powerFactor??"—"}.`)),
   heading("3.2 ILUMINAÇÃO"),tableFor(lights),
   ...lights.map(l=>para(`${l.description}: ${l.quantity} × ${money(l.power)} ${l.unit}; FP estimado ${l.powerFactor??"—"}.`)),
   heading("4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA"),para(`Tensão de operação informada: ${p.voltage||"não informada"} V. Os demais dados técnicos devem ser preenchidos conforme o projeto específico.`),
   heading("5. DEMANDA E DIMENSIONAMENTO"),para("Os resultados apresentados utilizam parâmetros estimados editáveis por equipamento."),
   heading("6. CARACTERÍSTICAS DO FORNECIMENTO"),para(`Tipo de fornecimento pretendido: ${p.supply}.`),
   heading("7. ATERRAMENTO E PROTEÇÕES"),para("Informações de aterramento e proteções não cadastradas nesta versão."),
   heading("8. RESPONSABILIDADE TÉCNICA"),para("Área reservada à assinatura digital do documento pelo profissional responsável."),
   para("RASCUNHO — documento ainda não assinado digitalmente.")
  ];
  const img=await imageBytes();
  const foot=showCrt?new Footer({children:[new Paragraph({alignment:AlignmentType.CENTER,children:[new ImageRun({data:img,transformation:{width:214,height:75},type:"png"})]})]}):undefined;
  const doc=new Document({styles:{default:{document:{run:{font:"Times New Roman",size:24},paragraph:{spacing:{line:300}}}}},sections:[{properties:{page:{margin:{top:1150,bottom:1300,left:900,right:900}}},footers:foot?{default:foot}:undefined,children}]});
  await saveBlob(await Packer.toBlob(doc),"memorial-eletrico-RASCUNHO.docx","Documento Word","docx");
 }
 async function makePdf():Promise<Uint8Array>{
  const pdf=new jsPDF({unit:"mm",format:"a4"});pdf.setFont("times","normal");pdf.setFontSize(12);
  const left=18,right=192,usable=right-left;let y=20;
  const watermark=()=>{pdf.setTextColor(210,210,210);pdf.setFontSize(37);pdf.text("RASCUNHO",105,145,{align:"center",angle:35});pdf.setTextColor(0,0,0);pdf.setFontSize(12);pdf.setFont("times","normal")};
  const next=(height:number)=>{if(y+height>265){pdf.addPage();watermark();y=22}};
  const line=(t:string,opts?:{bold?:boolean;center?:boolean;gap?:number})=>{pdf.setFont("times",opts?.bold?"bold":"normal");const lines=pdf.splitTextToSize(t,usable) as string[];next(lines.length*5+5);pdf.text(lines,opts?.center?105:left,y,{align:opts?.center?"center":"left"});y+=lines.length*5+(opts?.gap??3)};
  const h=(t:string)=>{y+=4;line(t,{bold:true,gap:3})};
  const table=(items:EditableLoad[])=>{const cols=[left,65,100,148,right];const widths=[47,35,48,44];const rows=[ ["Descrição","Quantidade","Potência unitária","Potência total"],...items.map(l=>[l.description,String(l.quantity),`${money(l.power)} ${l.unit}`,`${money(l.power*l.quantity)} ${l.unit}`])];next(rows.length*9+5);for(let ri=0;ri<rows.length;ri++){if(ri%2===1){pdf.setFillColor(217,217,217);pdf.rect(left,y-6,usable,9,"F")}pdf.setDrawColor(100);pdf.rect(left,y-6,usable,9);for(let j=1;j<cols.length-1;j++)pdf.line(cols[j],y-6,cols[j],y+3);pdf.setFont("times",ri===0?"bold":"normal");pdf.setFontSize(9);rows[ri].forEach((v,j)=>{const max=widths[j]-2;const clipped=pdf.getTextWidth(v)>max?v.slice(0,Math.max(2,Math.floor(v.length*max/pdf.getTextWidth(v))-1))+"…":v;pdf.text(clipped,cols[j]+widths[j]/2,y,{align:"center"})});pdf.setFontSize(12);y+=9}y+=5};
  watermark();line("MEMORIAL DESCRITIVO",{bold:true,center:true,gap:5});line(`INSTALAÇÃO ELÉTRICA – ALIMENTAÇÃO ${p.supply.toUpperCase()}`,{bold:true,center:true,gap:12});
  h("1. IDENTIFICAÇÃO DO CONSUMIDOR");line(`Titular: ${p.consumer||"(não informado)"}`);line(`CPF: ${p.cpf||"(não informado)"}`);line(`Endereço: ${p.address||"(não informado)"}`);
  h("2. OBJETIVO");line(`Este memorial apresenta a relação de cargas prevista para atendimento do consumidor identificado, por meio de fornecimento ${p.supply.toLowerCase()}.`);
  h("3. RELAÇÃO DE CARGA INSTALADA");h("3.1 MOTORES");table(groupLoads(p.loads,true));
  for(const l of groupLoads(p.loads,true))line(`${l.description}: ${l.quantity} × ${money(l.power)} ${l.unit}; rendimento estimado ${l.efficiency??"—"}%; FP estimado ${l.powerFactor??"—"}.`);
  h("3.2 ILUMINAÇÃO");table(groupLoads(p.loads,false));for(const l of groupLoads(p.loads,false))line(`${l.description}: ${l.quantity} × ${money(l.power)} ${l.unit}; FP estimado ${l.powerFactor??"—"}.`);
  h("4. DADOS TÉCNICOS DA ENTRADA DE ENERGIA");line(`Tensão de operação informada: ${p.voltage||"não informada"} V. Demais dados técnicos não cadastrados.`);
  h("5. DEMANDA E DIMENSIONAMENTO");line("Resultados baseados em parâmetros estimados editáveis por equipamento.");
  h("6. CARACTERÍSTICAS DO FORNECIMENTO");line(`Tipo de fornecimento pretendido: ${p.supply}.`);
  h("7. ATERRAMENTO E PROTEÇÕES");line("Informações de aterramento e proteções não cadastradas nesta versão.");
  h("8. RESPONSABILIDADE TÉCNICA");if(includeSignatureBlock)line("Espaço reservado à assinatura digital GOV.BR no PDF gerado.");line("Documento em rascunho, ainda não assinado digitalmente.");
  if(showCrt){const img=await imageBytes();const b64=btoa(String.fromCharCode(...img));for(let page=1;page<=pdf.getNumberOfPages();page++){pdf.setPage(page);pdf.addImage(`data:image/png;base64,${b64}`,"PNG",77,271,56,19.7)}}
  return new Uint8Array(pdf.output("arraybuffer"));
 }
 async function exportPdf(){try{const bytes=await makePdf();await saveBytes(bytes,"memorial-eletrico-RASCUNHO.pdf","Documento PDF","pdf");setLastPdfHash(await digest(bytes));setNotice("PDF gerado. Assine este PDF individualmente pelo GOV.BR e depois importe o arquivo assinado.")}catch(e){setNotice("Erro ao salvar PDF: "+String(e))}}
 async function importSignedPdf(file:File){
  if(!file.name.toLowerCase().endsWith(".pdf"))throw Error("Selecione um arquivo PDF");
  const bytes=new Uint8Array(await file.arrayBuffer());
  if(bytes.length<5||new TextDecoder().decode(bytes.slice(0,5))!=="%PDF-")throw Error("Arquivo não parece ser PDF");
  const record:SignatureRecord={name:file.name,date:new Date().toISOString(),size:file.size,documentHash:await digest(bytes),status:"importado-sem-verificacao"};
  const records=[record,...signatureRecords];setSignatureRecords(records);localStorage.setItem("memorial-assinaturas-v1",JSON.stringify(records));setImportedPdf(file.name);setNotice("PDF assinado importado para registro. A autenticidade da assinatura não foi verificada pelo aplicativo. Guarde o PDF original assinado.");
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
 <section><h2>3. Responsabilidade técnica e assinatura GOV.BR</h2>
 <p>Identificação profissional do modelo, exibida no rodapé de cada página. A imagem não é assinatura digital.</p>
 <div className="crt-preview"><img src={crtImage} alt="Identificação CRT/CFT do profissional" /></div>
 <label className="check"><input type="checkbox" checked={showCrt} onChange={e=>setShowCrt(e.target.checked)}/> Inserir CRT.PNG no rodapé</label>
 <label className="check"><input type="checkbox" checked={includeSignatureBlock} onChange={e=>setIncludeSignatureBlock(e.target.checked)}/> Reservar área de assinatura GOV.BR no documento</label>
 <p><strong>Fluxo de assinatura:</strong> gere o PDF, assine o arquivo pelo GOV.BR e importe o PDF assinado. Cada PDF precisa receber sua própria assinatura digital; a assinatura de um documento não pode ser reaproveitada criptograficamente em outro.</p>
 <div className="actions"><button onClick={exportPdf}>1. Gerar PDF para assinatura</button><a href="https://assinador.iti.br/" target="_blank" rel="noreferrer">2. Abrir assinador GOV.BR</a><label className="import">3. Registrar PDF assinado<input type="file" accept=".pdf,application/pdf" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;try{await importSignedPdf(f)}catch(err){setNotice(String(err))}e.target.value=""}}/></label></div>
 {lastPdfHash&&<p className="hint">SHA-256 do último PDF gerado: <code>{lastPdfHash}</code></p>}
 {importedPdf&&<p>Último arquivo registrado: <strong>{importedPdf}</strong> (assinatura ainda não verificada).</p>}
 <h3>Arquivos assinados registrados neste computador</h3>
 {signatureRecords.length===0?<p>Nenhum arquivo registrado.</p>:<ul>{signatureRecords.slice(0,10).map(x=><li key={x.documentHash}>{x.name} — {new Date(x.date).toLocaleString("pt-BR")} — importado, assinatura não verificada</li>)}</ul>}
 <p className="hint">O aplicativo guarda apenas o registro do arquivo importado, não uma cópia do PDF. Guarde o PDF assinado no local escolhido.</p>
 </section>
 <section><h2>4. Resultados indicativos</h2>{calculations.map(({load,result,error})=><div className="result" key={load.id}><b>{load.description}</b>: {error||`Mecânica ${fmt(result?.mechanicalKw)} kW | Ativa ${fmt(result?.activeKw)} kW | Aparente ${fmt(result?.apparentKva)} kVA`}<small>{result?.warnings.join(" ")}</small></div>)}
 {!valid&&<p className="warning">Corrija os dados inválidos antes de utilizar os resultados.</p>}
 <p className="warning">RASCUNHO. O armazenamento atual é local no navegador/WebView (localStorage), não SQLite. Não há verificação criptográfica de assinatura, cálculo de demanda normativa ou dimensionamento de entrada.</p>
 <div className="actions"><button onClick={exportWord}>Exportar DOCX (rascunho)</button><button onClick={exportPdf}>Salvar PDF (rascunho)</button><button onClick={exportBackup}>Backup JSON</button><label className="import">Importar backup JSON<input type="file" accept=".json,application/json" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;try{const data=JSON.parse(await f.text());if(!Array.isArray(data.loads))throw Error("Backup inválido");set(data)}catch{alert("Arquivo inválido")}}}/></label></div>
 {notice&&<p role="status">{notice}</p>}
 </section></main>
}
