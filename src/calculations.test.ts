import {describe,it,expect} from "vitest";
import {calculate,current} from "./calculations";
const motor={id:"m",description:"Motor elétrico",category:"Motores",quantity:1,power:37,unit:"CV" as const};
describe("potências e correntes",()=>{
 it("não inventa FP nem rendimento",()=>{const r=calculate(motor);expect(r.mechanicalKw).toBeCloseTo(27.2135);expect(r.apparentKva).toBeUndefined()});
 it("usa premissas explícitas",()=>expect(calculate(motor,.9,.85).apparentKva).toBeCloseTo(35.5732,3));
 it("corrente trifásica",()=>expect(current(35.5732,380,3)).toBeCloseTo(54.05,1));
 it("rejeita quantidade fracionária",()=>expect(()=>calculate({...motor,quantity:1.5})).toThrow());
});
