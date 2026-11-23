export type Unit = "W" | "kW" | "CV" | "HP" | "VA" | "kVA";
export type Load = {id:string;description:string;category:string;quantity:number;power:number;unit:Unit};
export type Result = {mechanicalKw?:number;activeKw?:number;apparentKva?:number;warnings:string[]};
export function calculate(load:Load, efficiency?:number, powerFactor?:number):Result {
 if(!Number.isInteger(load.quantity)||load.quantity<=0||!Number.isFinite(load.power)||load.power<=0)throw Error("Quantidade inteira e potência positiva obrigatórias.");
 if(efficiency!==undefined && !(efficiency>0&&efficiency<=1))throw Error("Rendimento inválido.");
 if(powerFactor!==undefined && !(powerFactor>0&&powerFactor<=1))throw Error("Fator de potência inválido.");
 const factors:Record<Unit,number>={W:.001,kW:1,CV:.7355,HP:.7457,VA:.001,kVA:1};
 const p=load.quantity*load.power*factors[load.unit];
 if(load.unit==="CV"||load.unit==="HP"){
  const activeKw=efficiency?p/efficiency:undefined;
  return {mechanicalKw:p,activeKw,apparentKva:activeKw&&powerFactor?activeKw/powerFactor:undefined,warnings:[...(!efficiency?["Rendimento não informado."]:[]),...(!powerFactor?["FP não informado."]:[])]};
 }
 if(load.unit==="VA"||load.unit==="kVA")return {apparentKva:p,activeKw:powerFactor?p*powerFactor:undefined,warnings:powerFactor?[]:["FP não informado."]};
 return {activeKw:p,apparentKva:powerFactor?p/powerFactor:undefined,warnings:powerFactor?[]:["FP não informado."]};
}
export function current(kva:number,volts:number,phases:1|3){
 if(!(kva>0&&volts>0))throw Error("Potência e tensão devem ser positivas.");
 return kva*1000/(phases===3?Math.sqrt(3)*volts:volts);
}
