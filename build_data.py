#!/usr/bin/env python3
import sys, re, json
from pathlib import Path
import pandas as pd
import numpy as np

def strip_html(s):
    if pd.isna(s):
        return ""
    s=str(s)
    s=re.sub(r"<br\s*/?>", " ", s, flags=re.I)
    s=re.sub(r"</?sub>", "", s, flags=re.I)
    s=re.sub(r"</?sup>", "", s, flags=re.I)
    s=re.sub(r"<[^>]+>", "", s)
    s=s.replace("&nbsp;"," ").replace("\u00a0"," ")
    return re.sub(r"\s+"," ",s).strip()

element_re=re.compile(r"([A-Z][a-z]?)")
def extract_elements(formula_text):
    if not formula_text:
        return []
    els=element_re.findall(formula_text)
    seen=set(); out=[]
    for e in els:
        if e not in seen:
            seen.add(e); out.append(e)
    return out

def normalize_country(code: str, country: str) -> str:
    c=(country or "").strip()
    code=(code or "").strip()
    if code == "Хабаровский" or c == "Хабаровский":
        return "Россия"
    if c == "Англия":
        return "Великобритания"
    return c

def safe_sum(s):
    s=pd.to_numeric(s, errors="coerce")
    return float(np.nansum(s.values)) if len(s) else 0.0

def main(xlsx_path: str):
    xlsx_path=Path(xlsx_path)
    df=pd.read_excel(xlsx_path, sheet_name="data")
    df=df[df["col-ID"].notna()].copy()

    nations=pd.read_excel(xlsx_path, sheet_name="nations")
    alpha3_to_name=dict(zip(nations["Alpha3"], nations["Наименование"]))

    records=[]
    for _,r in df.iterrows():
        formula_html=r.get("Формула")
        formula_txt=strip_html(formula_html)

        y=r.get("Год открытия")
        discovery_year=None
        if not pd.isna(y):
            try:
                yf=float(y)
                discovery_year=int(yf) if yf.is_integer() else yf
            except Exception:
                discovery_year=str(y)

        code="" if pd.isna(r.get("Страна")) else str(r.get("Страна")).strip()
        mapped=alpha3_to_name.get(code, code)
        country=normalize_country(code, mapped)

        records.append({
            "id": str(r.get("col-ID","")).strip(),
            "name_ru": str(r.get("Название","")).strip(),
            "name_other": ("" if pd.isna(r.get("Прочие")) else str(r.get("Прочие")).strip()),
            "ima_name": ("" if pd.isna(r.get("IMA Name")) else str(r.get("IMA Name")).strip()),
            "abbr": ("" if pd.isna(r.get("Сокращение")) else str(r.get("Сокращение")).strip()),
            "class": ("" if pd.isna(r.get("Класс")) else str(r.get("Класс")).strip()),
            "formula_html": ("" if pd.isna(formula_html) else str(formula_html)),
            "formula_text": formula_txt,
            "elements": extract_elements(formula_txt),
            "syngony": ("" if pd.isna(r.get("Сингония")) else str(r.get("Сингония")).strip()),
            "locality": ("" if pd.isna(r.get("Месторождение")) else str(r.get("Месторождение")).strip()),
            "country_code": code,
            "country": country,
            "discovery_year": discovery_year,
            "cost": (None if pd.isna(r.get("Стоимость")) else float(r.get("Стоимость"))),
            "ue_band": ("" if pd.isna(r.get("УЕ")) else str(r.get("УЕ")).strip()),
        })

    df_sum=pd.DataFrame(records)
    summary={"overall":{
        "count": int(len(df_sum)),
        "total_cost": safe_sum(df_sum["cost"]),
        "avg_cost": float(np.nanmean(pd.to_numeric(df_sum["cost"], errors="coerce"))) if df_sum["cost"].notna().any() else None,
        "countries": int(df_sum["country"].replace("",np.nan).dropna().nunique()),
        "localities": int(df_sum["locality"].replace("",np.nan).dropna().nunique()),
        "classes": int(df_sum["class"].replace("",np.nan).dropna().nunique()),
    }}
    by_class=[]
    for cls,g in df_sum.groupby("class", dropna=False):
        cls="" if (cls is None or (isinstance(cls,float) and np.isnan(cls))) else str(cls)
        if not cls:
            continue
        costs=pd.to_numeric(g["cost"], errors="coerce")
        by_class.append({
            "class": cls,
            "count": int(len(g)),
            "total_cost": float(np.nansum(costs.values)),
            "avg_cost": float(np.nanmean(costs.values)) if costs.notna().any() else None,
            "countries": int(g["country"].replace("",np.nan).dropna().nunique()),
            "localities": int(g["locality"].replace("",np.nan).dropna().nunique()),
        })
    summary["by_class"]=sorted(by_class, key=lambda x: x["count"], reverse=True)

    out_dir=Path(__file__).resolve().parent
    (out_dir/"data.json").write_text(json.dumps(records, ensure_ascii=False, indent=2), encoding="utf-8")
    (out_dir/"summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print("OK: data.json + summary.json обновлены")

if __name__=="__main__":
    if len(sys.argv)<2:
        print("Usage: python build_data.py <path-to-xlsx>")
        raise SystemExit(2)
    main(sys.argv[1])
