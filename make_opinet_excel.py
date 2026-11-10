"""Create a reader-friendly Excel workbook from existing Opinet CSV files.

This script never calls the Opinet API. It reads only ./opinet_data/*.csv.
"""
from __future__ import annotations

import csv, json, os, subprocess, tempfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "opinet_data"
OUT = ROOT / "data" / "opinet" / "2026-11-10" / "유가_기간비교.xlsx"
NODE = Path(r"C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe")
NODE_MODULES = Path(r"C:\Users\USER\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules")

def read_csv(path):
    with path.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))

def num(v):
    try: return float(v)
    except (TypeError, ValueError): return None

def main():
    meta = json.loads((DATA / "collection_metadata.json").read_text(encoding="utf-8"))
    national, brands, areas = [], [], []
    for p in sorted(DATA.glob("*.csv")):
        if p.name == "sido_codes.csv": continue
        rows = read_csv(p)
        period = "최근 7일" if p.name.startswith("recent_7d") else "직전 7일"
        if "national_product" in p.name:
            for r in rows: national.append({"기간": period, "날짜": r.get("DATE",""), "제품코드": r.get("PRODCD",""), "제품명": r.get("PRODNM",r.get("PRODCD","")), "평균가격(원/리터)": num(r.get("PRICE"))})
        elif "national_brand" in p.name:
            for r in rows: brands.append({"기간": period, "날짜": r.get("DATE",""), "제품코드": r.get("PRODCD",""), "제품명": r.get("PRODNM",r.get("PRODCD","")), "상표코드": r.get("POLL_DIV_CD",""), "상표명": r.get("POLL_DIV_NM",r.get("POLL_DIV_CD","")), "평균가격(원/리터)": num(r.get("PRICE"))})
        elif "area_product" in p.name:
            for r in rows: areas.append({"기간": period, "날짜": r.get("DATE",""), "지역코드": r.get("AREA_CD",""), "지역명": r.get("AREA_NM",r.get("AREA_CD","")), "제품코드": r.get("PRODCD",""), "제품명": r.get("PRODNM",r.get("PRODCD","")), "평균가격(원/리터)": num(r.get("PRICE"))})
    products = sorted({r["제품코드"] for r in national})
    names = {r["제품코드"]: r["제품명"] for r in national if r["제품명"]}
    def avg(rows, code, period, key=None):
        vals = [r["평균가격(원/리터)"] for r in rows if r["제품코드"] == code and r["기간"] == period and (key is None or key(r)) and r["평균가격(원/리터)"] is not None]
        return round(sum(vals)/len(vals), 2) if vals and len(vals) == 7 else None
    summary=[]
    for code in products:
        a=avg(national,code,"최근 7일"); b=avg(national,code,"직전 7일")
        summary.append([names.get(code,code),code,a,b,(a-b if a is not None and b is not None else None),((a-b)/b if a is not None and b not in (None,0) else None)])
    payload={"summary":summary,"national":national,"brands":brands,"areas":areas,"products":[names.get(x,x) for x in products],"meta":meta,"out":str(OUT)}
    js = r'''import fs from "node:fs/promises"; import {Workbook,SpreadsheetFile} from "file:///C:/Users/USER/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";
const d=JSON.parse(await fs.readFile(process.argv[2],"utf8")); const wb=Workbook.create();
const sh=Object.fromEntries(["기간별 비교","전국 일별","상표별 일별","지역별 일별"].map(n=>[n,wb.worksheets.add(n)]));
const navy="#1F4E78", blue="#D9EAF7", pale="#F5F8FB", green="#E2F0D9", red="#FCE4D6", font={name:"Arial",size:10};
function setup(s,title,subtitle,last){s.showGridLines=false; s.getRange("A1").values=[[title]]; s.getRange("A1").format={font:{name:"Arial",size:15,bold:true,color:navy}}; s.getRange("A2").values=[[subtitle]]; s.getRange("A2").format={font:{name:"Arial",size:10,italic:true,color:"#666666"}}; s.getRange(`A1:${last}1`).format.borders={bottom:{style:"medium",color:navy}};}
function table(s,start,headers,rows,name){const m=start.match(/^([A-Z]+)(\d+)$/); const c=m[1], row=Number(m[2]); const end=String.fromCharCode(64+headers.length); const n=rows.length+row; const range=`${c}${row}:${end}${n}`; s.getRange(range).values=[headers,...rows]; const t=s.tables.add(range,true,name); t.style="TableStyleMedium2"; s.getRange(range).format.font=font; s.getRange(range).format.verticalAlignment="center"; return t;}
setup(sh["기간별 비교"],"유가 기간 비교","기준일: 2026-11-10 | 단위: 원/리터 | 일별 데이터 7일 완비 시에만 평균·증감 계산", "F");
sh["기간별 비교"].getRange("A4:F4").values=[["제품명","제품코드","최근 7일 평균","직전 7일 평균","증감액","증감률"]]; sh["기간별 비교"].getRange("A5:F"+(4+d.summary.length)).values=d.summary; sh["기간별 비교"].tables.add("A4:F"+(4+d.summary.length),true,"PeriodCompare");
sh["기간별 비교"].getRange("A4:F4").format={fill:navy,font:{name:"Arial",size:10,bold:true,color:"#FFFFFF"}}; sh["기간별 비교"].getRange("C5:E"+(4+d.summary.length)).format.numberFormat="#,##0.00"; sh["기간별 비교"].getRange("F5:F"+(4+d.summary.length)).format.numberFormat="0.0%";
sh["기간별 비교"].getRange("A"+(7+d.summary.length)).values=[["상표별·지역별 비교는 아래 상세 시트에서 제품명 필터를 사용해 선택하세요."]]; sh["기간별 비교"].getRange("A"+(7+d.summary.length)).format={font:{name:"Arial",size:10,italic:true,color:"#666666"}}; sh["기간별 비교"].freezePanes.freezeRows(4);
setup(sh["전국 일별"],"전국 일별 평균가격","제품명 필터로 제품을 선택할 수 있습니다. 단위: 원/리터", "E"); table(sh["전국 일별"],"A4",["기간","날짜","제품코드","제품명","평균가격(원/리터)"],d.national.map(x=>[x["기간"],x["날짜"],x["제품코드"],x["제품명"],x["평균가격(원/리터)"]]),"NationalDaily"); sh["전국 일별"].freezePanes.freezeRows(4);
setup(sh["상표별 일별"],"상표별 일별 평균가격","제품명 필터로 제품을 선택한 뒤 상표별 가격을 비교하세요. 단위: 원/리터", "G"); table(sh["상표별 일별"],"A4",["기간","날짜","제품코드","제품명","상표코드","상표명","평균가격(원/리터)"],d.brands.map(x=>[x["기간"],x["날짜"],x["제품코드"],x["제품명"],x["상표코드"],x["상표명"],x["평균가격(원/리터)"]]),"BrandDaily"); sh["상표별 일별"].freezePanes.freezeRows(4);
setup(sh["지역별 일별"],"지역별 일별 평균가격","시도(2자리)만 포함했습니다. 제품명·지역명 필터로 비교하세요. 단위: 원/리터", "G"); table(sh["지역별 일별"],"A4",["기간","날짜","지역코드","지역명","제품코드","제품명","평균가격(원/리터)"],d.areas.map(x=>[x["기간"],x["날짜"],x["지역코드"],x["지역명"],x["제품코드"],x["제품명"],x["평균가격(원/리터)"]]),"AreaDaily"); sh["지역별 일별"].freezePanes.freezeRows(4);
for(const s of Object.values(sh)){const u=s.getUsedRange();u.format.font=font;u.format.autofitColumns();u.format.rowHeight=18;} for(const n of ["전국 일별","상표별 일별","지역별 일별"]){sh[n].getRange("E5:Z10000").format.numberFormat="#,##0.00";}
wb.recalculate(); const out=await SpreadsheetFile.exportXlsx(wb); await out.save(d.out);'''
    with tempfile.TemporaryDirectory() as td:
        payload_path=Path(td)/"payload.json"; js_path=Path(td)/"build.mjs"; payload_path.write_text(json.dumps(payload,ensure_ascii=False),encoding="utf-8"); js_path.write_text(js,encoding="utf-8")
        env=os.environ.copy(); env["NODE_PATH"]=str(NODE_MODULES); OUT.parent.mkdir(parents=True,exist_ok=True)
        subprocess.run([str(NODE),str(js_path),str(payload_path)],check=True,env=env)
    print(f"created {OUT}")

if __name__=="__main__": main()
